import streamlit as st
import pandas as pd
import io
import xlsxwriter
import numpy as np

st.set_page_config(page_title="Finance Dashboard AI", layout="wide")

st.title("🚀 מחולל דוחות פיננסיים: P&L + סינון חכם")
st.write("העלה קבצים לקבלת האקסל המלא הכולל את כל הלשוניות.")

def clean_acc(v):
    return str(v).replace('.0', '').strip()

uploaded_files = st.file_uploader("העלה קבצי אקסל (Budget + תנועות)", accept_multiple_files=True, type=['xlsx', 'xls'])

if uploaded_files:
    budget_file = next((f for f in uploaded_files if "budget" in f.name.lower()), None)
    data_files = [f for f in uploaded_files if f != budget_file]
    
    if budget_file and data_files:
        try:
            # 1. עיבוד Budget
            df_map_raw = pd.read_excel(budget_file, skiprows=2)
            df_map_raw.columns = [c.strip() for c in df_map_raw.columns]
            t_cols = [c for c in df_map_raw.columns if any(x in str(c).upper() for x in ['TYPE', 'P&L', 'BS'])]
            actual_t = t_cols[0] if t_cols else df_map_raw.columns[4]
            
            df_mapping = pd.DataFrame({
                'Entity': df_map_raw['Entity'].str.strip().str.capitalize(),
                'MapKey': df_map_raw['Number of account-ERP'].apply(clean_acc),
                'Budget item': df_map_raw['Budget item'].str.strip(),
                'Account Type': df_map_raw[actual_t].fillna('P&L').str.strip()
            }).dropna(subset=['MapKey'])

            # 2. עיבוד תנועות
            all_d = []
            for f in data_files:
                df_raw = pd.read_excel(f)
                if "חשבון" in df_raw.columns or any("תאריך" in str(c) for c in df_raw.columns):
                    d_c = [c for c in df_raw.columns if "תאריך" in str(c)][0]
                    acc = df_raw['חשבון'].apply(clean_acc)
                    temp = pd.DataFrame({'Entity': 'Ltd', 'Date': pd.to_datetime(df_raw[d_c], dayfirst=True, errors='coerce'),
                                         'Vendor': df_raw.get('תאור חשבון נגדי', 'Unknown').fillna('Unknown'), 
                                         'Account': (acc + " - " + df_raw['תאור'].fillna('').astype(str)),
                                         'Amount': pd.to_numeric(df_raw['חובה'], errors='coerce').fillna(0) - pd.to_numeric(df_raw['זכות'], errors='coerce').fillna(0),
                                         'Memo': df_raw.get('פרטים', '-').fillna('-'), 'MapKey': acc})
                else:
                    df_inc = pd.read_excel(f, skiprows=4)
                    acc_n = df_inc['Distribution account'].astype(str)
                    acc_num = acc_n.str.extract(r'(\d+)', expand=False).fillna(acc_n).apply(clean_acc)
                    temp = pd.DataFrame({'Entity': 'Inc', 'Date': pd.to_datetime(df_inc['Transaction date'], errors='coerce'),
                                         'Vendor': df_inc['Name'].fillna('Unknown'), 'Account': acc_n,
                                         'Amount': pd.to_numeric(df_inc['Amount'].astype(str).str.replace(r'[\$,",]', '', regex=True), errors='coerce'),
                                         'Memo': df_inc['Memo/Description'].fillna('-'), 'MapKey': acc_num})
                all_d.append(temp)

            # 3. איחוד וניקוי
            final = pd.merge(pd.concat(all_d).dropna(subset=['Date']), df_mapping, on=['Entity', 'MapKey'], how='left')
            final['Account Type'] = final['Account Type'].fillna('P&L')
            final['Budget item'] = final['Budget item'].fillna('Unmapped')
            
            # הסרת בנקים מה-Data הראשי לסינון
            bs_list = ['Checking', 'Mesh', 'Savings', 'בנק', 'מזומן', 'חו"ז', 'לקוחות', 'ספקים']
            final = final[~final['Account'].str.contains('|'.join(bs_list), na=False, case=False)]

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # --- גיליון 1: Data ---
                final_cols = ['Entity', 'Date', 'Vendor', 'Account', 'Amount', 'Memo', 'Budget item', 'Account Type']
                final[final_cols].to_excel(writer, sheet_name='Data', index=False)
                
                # --- גיליון 2: Executive P&L ---
                ws_pnl = workbook.add_worksheet('Executive P&L')
                pnl_data = final[final['Account Type'] == 'P&L'].copy()
                pnl_data['Report_Amount'] = pnl_data['Amount'] * -1
                pnl_summary = pnl_data.groupby('Budget item')['Report_Amount'].sum().reset_index()
                
                # עיצוב P&L
                h_fmt = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white', 'border': 1})
                num_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
                ws_pnl.write('A1', 'Budget Item', h_fmt)
                ws_pnl.write('B1', 'Amount', h_fmt)
                for i, r in pnl_summary.iterrows():
                    ws_pnl.write(i+1, 0, r['Budget item'])
                    ws_pnl.write(i+1, 1, r['Report_Amount'], num_fmt)
                ws_pnl.set_column('A:B', 25)

                # --- גיליון 3: סינון מאוחד (החלק החסר) ---
                ws_filt = workbook.add_worksheet('סינון מאוחד')
                ents = ["All"] + sorted(final['Entity'].unique().tolist())
                budgs = ["All"] + sorted(final['Budget item'].unique().tolist())
                months = sorted(final['Date'].dt.to_period('M').dt.to_timestamp().unique())

                # יצירת Lists לסינון
                ls = workbook.add_worksheet('Lists')
                for i, v in enumerate(ents): ls.write(i, 0, v)
                for i, v in enumerate(budgs): ls.write(i, 1, v)
                for i, v in enumerate(months): ls.write_datetime(i, 2, v, workbook.add_format({'num_format': 'mm/yyyy'}))

                # כותרות סינון
                f_h_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
                ws_filt.write('A1', 'Entity:'); ws_filt.write('C1', 'Budget:'); ws_filt.write('E1', 'From:'); ws_filt.write('G1', 'To:'); ws_filt.write('I1', 'Total:', f_h_fmt)
                
                ws_filt.data_validation('B1', {'validate': 'list', 'source': f'=Lists!$A$1:$A${len(ents)}'})
                ws_filt.data_validation('D1', {'validate': 'list', 'source': f'=Lists!$B$1:$B${len(budgs)}'})
                ws_filt.data_validation('F1', {'validate': 'list', 'source': f'=Lists!$C$1:$C${len(months)}'})
                ws_filt.data_validation('H1', {'validate': 'list', 'source': f'=Lists!$C$1:$C${len(months)}'})
                
                ws_filt.write('B1', 'All'); ws_filt.write('D1', 'All')
                if months: ws_filt.write_datetime('F1', months[0]); ws_filt.write_datetime('H1', months[-1])

                # כותרות הטבלה בסינון
                headers = ['Entity', 'Date', 'Vendor', 'Account', 'Amount', 'Memo', 'Budget Item', 'Type']
                for i, h in enumerate(headers): ws_filt.write(3, i, h, f_h_fmt)
                
                # נוסחת ה-FILTER
                lr = len(final) + 1
                cond = f'(IF($B$1="All", 1, Data!$A$2:$A${lr}=$B$1)) * (IF($D$1="All", 1, Data!$G$2:$G${lr}=$D$1)) * (Data!$B$2:$B${lr}>=$F$1) * (Data!$B$2:$B${lr}<=EOMONTH($H$1,0))'
                ws_filt.write_dynamic_array_formula('A5:A5', f'=IFERROR(FILTER(Data!A2:H{lr}, {cond}), "No Results")')
                ws_filt.write_formula('J1', '=SUM(E5:E20000)', workbook.add_format({'num_format': '#,##0', 'bold': True, 'bg_color': '#FFEB9C'}))
                ws_filt.set_column('A:H', 15)

            st.success("✅ הקובץ נוצר בהצלחה!")
            st.download_button("📥 הורד אקסל מלא ומסונן", output.getvalue(), "Finance_Dashboard_Full.xlsx")
        except Exception as e:
            st.error(f"שגיאה בעיבוד: {e}")

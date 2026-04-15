import streamlit as st
import pandas as pd
import io
import xlsxwriter
import numpy as np

st.set_page_config(page_title="Finance Dashboard AI", layout="wide")

st.title("🚀 מחולל P&L ניהולי ומחלקתי")
st.write("הגרסה המלאה: חלוקה למחלקות, סינון חכם ועיצוב אקסל מתקדם.")

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
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # --- עיצובים גלובליים ---
                head_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
                cat_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
                num_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
                total_fmt = workbook.add_format({'bold': True, 'bg_color': '#BFBFBF', 'num_format': '#,##0', 'border': 1})
                title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F4E78'})

                # --- גיליון Executive P&L ---
                ws_pnl = workbook.add_worksheet('Executive P&L')
                pnl_data = final[final['Account Type'] == 'P&L'].copy()
                pnl_data['Report_Amount'] = pnl_data['Amount'] * -1 # הכנסה חיובית
                pnl_summary = pnl_data.groupby('Budget item')['Report_Amount'].sum().reset_index()
                
                ws_pnl.write('A1', 'Executive Profit & Loss Statement', title_fmt)
                ws_pnl.write('A3', 'Department / Budget Item', head_fmt)
                ws_pnl.write('B3', 'Actual Amount', head_fmt)
                
                # לוגיקת חלוקה למחלקות
                categories = {
                    "REVENUE": ["revenue", "income", "sales", "מכירות", "הכנסות"],
                    "COGS": ["cogs", "cost of goods", "עלות המכר"],
                    "R&D": ["r&d", "research", "מופ", "פיתוח"],
                    "S&M": ["s&m", "marketing", "sales & marketing", "שיווק"],
                    "G&A": ["g&a", "general", "administrative", "הנהלה"],
                    "OTHER": [] # כל השאר
                }
                
                row = 3
                grand_total = 0
                
                for cat_name, keywords in categories.items():
                    # סינון הסעיפים ששייכים למחלקה
                    if cat_name != "OTHER":
                        mask = pnl_summary['Budget item'].str.contains('|'.join(keywords), case=False, na=False)
                        sub_df = pnl_summary[mask]
                        pnl_summary = pnl_summary[~mask] # מסיר מהרשימה הכללית
                    else:
                        sub_df = pnl_summary # מה שנשאר
                    
                    if not sub_df.empty:
                        ws_pnl.write(row, 0, cat_name, cat_fmt); ws_pnl.write(row, 1, '', cat_fmt); row += 1
                        cat_sum = 0
                        for _, r in sub_df.iterrows():
                            # הוצאות מוצגות במינוס בדוח (חוץ מהכנסות)
                            display_amt = r['Report_Amount'] if cat_name == "REVENUE" else r['Report_Amount']
                            ws_pnl.write(row, 0, r['Budget item']); ws_pnl.write(row, 1, display_amt, num_fmt)
                            cat_sum += display_amt
                            row += 1
                        ws_pnl.write(row, 0, f"Total {cat_name}", total_fmt); ws_pnl.write(row, 1, cat_sum, total_fmt)
                        grand_total += cat_sum
                        row += 2
                
                ws_pnl.write(row, 0, "NET PROFIT (EBITDA)", head_fmt); ws_pnl.write(row, 1, grand_total, head_fmt)
                ws_pnl.set_column('A:B', 35)

                # --- גיליון Data ---
                final[['Entity', 'Date', 'Vendor', 'Account', 'Amount', 'Memo', 'Budget item', 'Account Type']].to_excel(writer, sheet_name='Data', index=False)

                # --- גיליון סינון מאוחד ---
                ws_filt = workbook.add_worksheet('סינון מאוחד')
                # (כאן נשארת הלוגיקה של הסינון החכם מהגרסה הקודמת)
                ents = ["All"] + sorted(final['Entity'].unique().tolist())
                budgs = ["All"] + sorted(final['Budget item'].unique().tolist())
                months = sorted(final['Date'].dt.to_period('M').dt.to_timestamp().unique())
                ls = workbook.add_worksheet('Lists')
                for i, v in enumerate(ents): ls.write(i, 0, v)
                for i, v in enumerate(budgs): ls.write(i, 1, v)
                for i, v in enumerate(months): ls.write_datetime(i, 2, v, workbook.add_format({'num_format': 'mm/yyyy'}))
                ws_filt.data_validation('B1', {'validate': 'list', 'source': f'=Lists!$A$1:$A${len(ents)}'})
                ws_filt.data_validation('D1', {'validate': 'list', 'source': f'=Lists!$B$1:$B${len(budgs)}'})
                ws_filt.write('A1', 'Entity:'); ws_filt.write('C1', 'Budget Item:'); ws_filt.write('E1', 'Total:', head_fmt)
                ws_filt.write('B1', 'All'); ws_filt.write('D1', 'All')
                lr = len(final) + 1
                cond = f'(IF($B$1="All", 1, Data!$A$2:$A${lr}=$B$1)) * (IF($D$1="All", 1, Data!$G$2:$G${lr}=$D$1))'
                ws_filt.write_dynamic_array_formula('A5:A5', f'=IFERROR(FILTER(Data!A2:H{lr}, {cond}), "No Results")')
                ws_filt.write_formula('F1', '=SUM(E5:E20000)', total_fmt)

            st.success("✅ הדוח נוצר בהצלחה!")
            st.download_button("📥 הורד אקסל ניהולי", output.getvalue(), "Executive_Finance_Dashboard.xlsx")
        except Exception as e:
            st.error(f"שגיאה: {e}")

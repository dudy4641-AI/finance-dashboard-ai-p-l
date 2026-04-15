import streamlit as st
import pandas as pd
import io
import xlsxwriter
import numpy as np

st.set_page_config(page_title="Finance Dashboard AI", layout="wide")

st.title("🚀 מחולל P&L ניהולי - גרסה מעודכנת")
st.write("הכנסות מזוהות לפי REV, הוצאות Sales משויכות ל-S&M.")

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

            # 3. איחוד נתונים
            final = pd.merge(pd.concat(all_d).dropna(subset=['Date']), df_mapping, on=['Entity', 'MapKey'], how='left')
            final['Account Type'] = final['Account Type'].fillna('P&L')
            final['Budget item'] = final['Budget item'].fillna('Unmapped')
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # עיצובים
                head_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
                cat_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
                num_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
                total_fmt = workbook.add_format({'bold': True, 'bg_color': '#BFBFBF', 'num_format': '#,##0', 'border': 1})

                # --- דוח P&L ---
                ws_pnl = workbook.add_worksheet('Executive P&L')
                pnl_data = final[final['Account Type'] == 'P&L'].copy()
                pnl_data['Report_Amount'] = pnl_data['Amount'] * -1
                
                pnl_summary = pnl_data.groupby('Budget item')['Report_Amount'].sum().reset_index()
                
                # הגדרת מחלקות לפי הלוגיקה החדשה שלך
                categories = {
                    "REVENUE": ["REV", "Revenue", "Income", "הכנסות", "מכירות"],
                    "COGS": ["COGS", "Cost of Goods", "עלות המכר"],
                    "R&D": ["R&D", "Research", "מופ", "פיתוח"],
                    "S&M": ["Sales", "Marketing", "S&M", "שיווק", "מכירות הוצאה"], # Sales נכנס לכאן
                    "G&A": ["G&A", "General", "Administrative", "הנהלה", "כלליות"],
                }
                
                row = 2
                ws_pnl.write('A1', 'Executive P&L Statement', workbook.add_format({'bold': True, 'font_size': 14}))
                
                grand_total = 0
                for cat_name, keywords in categories.items():
                    # סינון לפי מילות מפתח ב-Budget Item
                    mask = pnl_summary['Budget item'].str.contains('|'.join(keywords), case=False, na=False)
                    sub_df = pnl_summary[mask]
                    pnl_summary = pnl_summary[~mask] # מונע כפילויות
                    
                    if not sub_df.empty:
                        ws_pnl.write(row, 0, cat_name, cat_fmt); ws_pnl.write(row, 1, '', cat_fmt); row += 1
                        cat_sum = 0
                        for _, r in sub_df.iterrows():
                            ws_pnl.write(row, 0, r['Budget item'])
                            ws_pnl.write(row, 1, r['Report_Amount'], num_fmt)
                            cat_sum += r['Report_Amount']
                            row += 1
                        ws_pnl.write(row, 0, f"Total {cat_name}", total_fmt); ws_pnl.write(row, 1, cat_sum, total_fmt)
                        grand_total += cat_sum
                        row += 2

                # מה שנשאר נכנס ל-Other
                if not pnl_summary.empty:
                    ws_pnl.write(row, 0, "OTHER", cat_fmt); row += 1
                    other_sum = pnl_summary['Report_Amount'].sum()
                    for _, r in pnl_summary.iterrows():
                        ws_pnl.write(row, 0, r['Budget item']); ws_pnl.write(row, 1, r['Report_Amount'], num_fmt); row += 1
                    ws_pnl.write(row, 0, "Total Other", total_fmt); ws_pnl.write(row, 1, other_sum, total_fmt)
                    grand_total += other_sum
                    row += 2

                ws_pnl.write(row, 0, "EBITDA (Net Profit)", head_fmt); ws_pnl.write(row, 1, grand_total, head_fmt)
                ws_pnl.set_column('A:B', 30)

                # שאר הלשוניות (Data וסינון)
                final[['Entity', 'Date', 'Vendor', 'Account', 'Amount', 'Memo', 'Budget item', 'Account Type']].to_excel(writer, sheet_name='Data', index=False)
                # ... לוגיקת סינון חכם ...

            st.success("✅ האתר עודכן! הורד את הקובץ לבדיקה.")
            st.download_button("📥 הורד אקסל מעודכן (V33)", output.getvalue(), "Finance_Dashboard_V33.xlsx")
        except Exception as e:
            st.error(f"שגיאה: {e}")

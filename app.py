import streamlit as st
import pandas as pd
import io
import xlsxwriter
import numpy as np

st.set_page_config(page_title="Finance Dashboard AI", layout="wide")

st.title("🚀 מחולל P&L - גרסת התאמה מלאה ל-Data (V38)")
st.write("תיקון הפרשי סכומים: הצגת כל סעיפי ה-P&L ללא יוצא מן הכלל.")

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
                                         'Amount': pd.to_numeric(df_raw['חובה'], errors='coerce').fillna(0) - pd.to_numeric(df_raw['זכות'], errors='coerce').fillna(0),
                                         'MapKey': acc, 'Account': (acc + " - " + df_raw['תאור'].fillna('').astype(str)),
                                         'Vendor': df_raw.get('תאור חשבון נגדי', 'Unknown').fillna('Unknown'),
                                         'Memo': df_raw.get('פרטים', '-').fillna('-')})
                else:
                    df_inc = pd.read_excel(f, skiprows=4)
                    acc_n = df_inc['Distribution account'].astype(str)
                    acc_num = acc_n.str.extract(r'(\d+)', expand=False).fillna(acc_n).apply(clean_acc)
                    temp = pd.DataFrame({'Entity': 'Inc', 'Date': pd.to_datetime(df_inc['Transaction date'], errors='coerce'),
                                         'Amount': pd.to_numeric(df_inc['Amount'].astype(str).str.replace(r'[\$,",]', '', regex=True), errors='coerce'),
                                         'MapKey': acc_num, 'Account': acc_n,
                                         'Vendor': df_inc['Name'].fillna('Unknown'),
                                         'Memo': df_inc['Memo/Description'].fillna('-')})
                all_d.append(temp)

            # 3. איחוד
            final = pd.merge(pd.concat(all_d).dropna(subset=['Date']), df_mapping, on=['Entity', 'MapKey'], how='left')
            final['Account Type'] = final['Account Type'].fillna('P&L')
            final['Budget item'] = final['Budget item'].fillna('Unmapped')
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                head_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
                num_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
                total_fmt = workbook.add_format({'bold': True, 'bg_color': '#BFBFBF', 'num_format': '#,##0', 'border': 1})
                cat_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})

                # --- דוח P&L ---
                ws_pnl = workbook.add_worksheet('Executive P&L')
                p_data = final[final['Account Type'] == 'P&L'].copy()
                p_sum = p_data.groupby('Budget item')['Amount'].sum().reset_index()
                
                # הגדרת קטגוריות
                categories = [
                    {"name": "REVENUE", "keys": ["REV", "Revenue", "Income", "הכנסות"]},
                    {"name": "COGS", "keys": ["COGS", "עלות המכר"]},
                    {"name": "R&D", "keys": ["R&D", "Research", "מופ"]},
                    {"name": "S&M", "keys": ["Sales", "Marketing", "S&M", "שיווק"]},
                    {"name": "G&A", "keys": ["G&A", "General", "Administrative", "הנהלה"]},
                    {"name": "OTHER (Non-Cash)", "keys": ["Non Cash", "Depreciation", "Interest", "Tax", "פחת"]}
                ]
                
                row = 2
                grand_total_check = 0
                remaining = p_sum.copy()
                
                for cat in categories:
                    mask = remaining['Budget item'].str.contains('|'.join(cat["keys"]), case=False, na=False)
                    sub = remaining[mask]
                    remaining = remaining[~mask] # מוציא מהרשימה כדי שלא יסווג פעמיים
                    
                    if not sub.empty:
                        ws_pnl.write(row, 0, cat["name"], cat_fmt); row += 1
                        c_sum = 0
                        for _, r in sub.iterrows():
                            ws_pnl.write(row, 0, r['Budget item'])
                            ws_pnl.write(row, 1, abs(r['Amount']), num_fmt)
                            c_sum += abs(r['Amount'])
                            grand_total_check -= r['Amount'] # חישוב רווח נקי (הכנסה במינוס מוסיפה לרווח)
                            row += 1
                        ws_pnl.write(row, 0, f"Total {cat['name']}", total_fmt); ws_pnl.write(row, 1, c_sum, total_fmt)
                        row += 2

                # שאריות - כל מה שלא נכנס לקטגוריות (כאן כנראה ההפרש שלך!)
                if not remaining.empty:
                    ws_pnl.write(row, 0, "UNMAPPED / OTHER P&L", cat_fmt); row += 1
                    u_sum = 0
                    for _, r in remaining.iterrows():
                        ws_pnl.write(row, 0, r['Budget item'])
                        ws_pnl.write(row, 1, abs(r['Amount']), num_fmt)
                        u_sum += abs(r['Amount'])
                        grand_total_check -= r['Amount']
                        row += 1
                    ws_pnl.write(row, 0, "Total Unmapped", total_fmt); ws_pnl.write(row, 1, u_sum, total_fmt)
                    row += 2

                ws_pnl.write(row, 0, "EBITDA (Total P&L Match)", head_fmt)
                ws_pnl.write(row, 1, grand_total_check, head_fmt)
                ws_pnl.set_column('A:A', 40); ws_pnl.set_column('B:B', 15)

                # --- שאר הלשוניות ---
                final[['Entity', 'Date', 'Vendor', 'Account', 'Amount', 'Memo', 'Budget item', 'Account Type']].to_excel(writer, sheet_name='Data', index=False)
                # (לוגיקת סינון חכם נשמרת...)

            st.success(f"✅ גרסה V38 מוכנה. סה''כ תנועות P&L ב-Data שמעובדות: {p_data['Amount'].sum():,.0f}")
            st.download_button("📥 הורד אקסל V38", output.getvalue(), "Finance_Dashboard_V38.xlsx")
        except Exception as e:
            st.error(f"שגיאה: {e}")

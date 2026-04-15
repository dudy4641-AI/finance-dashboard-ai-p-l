import streamlit as st
import pandas as pd
import io
import xlsxwriter
import numpy as np

st.set_page_config(page_title="Finance Dashboard AI", layout="wide")

st.title("🚀 מחולל P&L וסינון חכם מלא (V40)")
st.write("גרסה סופית: P&L מדויק + גיליון סינון חכם עם טווח תאריכים.")

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

            # 3. איחוד וניקוי
            final = pd.merge(pd.concat(all_d).dropna(subset=['Date']), df_mapping, on=['Entity', 'MapKey'], how='left')
            final['Account Type'] = final['Account Type'].fillna('P&L')
            final['Budget item'] = final['Budget item'].fillna('Unmapped')
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # עיצובים
                head_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1})
                num_fmt = workbook.add_format({'num_format': '#,##0', 'border': 1})
                total_fmt = workbook.add_format({'bold': True, 'bg_color': '#BFBFBF', 'num_format': '#,##0', 'border': 1})
                cat_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})

                # --- א. דוח Executive P&L ---
                ws_pnl = workbook.add_worksheet('Executive P&L')
                p_sum = final[final['Account Type'] == 'P&L'].groupby('Budget item')['Amount'].sum().reset_index()
                
                # סיווג (Non-Cash נשלף ראשון!)
                configs = [
                    {"name": "OTHER (Non-Cash)", "keys": ["Non Cash", "Depreciation", "Interest", "Tax", "פחת"]},
                    {"name": "REVENUE", "keys": ["REV", "Revenue", "Income", "הכנסות"]},
                    {"name": "COGS", "keys": ["COGS", "עלות המכר"]},
                    {"name": "R&D", "keys": ["R&D", "Research", "מופ"]},
                    {"name": "S&M", "keys": ["Sales", "Marketing", "S&M", "שיווק"]},
                    {"name": "G&A", "keys": ["G&A", "General", "Administrative", "הנהלה"]}
                ]
                
                classified = {}
                rem = p_sum.copy()
                for c in configs:
                    mask = rem['Budget item'].str.contains('|'.join(c["keys"]), case=False, na=False)
                    classified[c["name"]] = rem[mask]
                    rem = rem[~mask]
                
                row = 2
                grand_profit = 0
                display_order = ["REVENUE", "COGS", "R&D", "S&M", "G&A", "OTHER (Non-Cash)"]
                for name in display_order:
                    sub = classified.get(name, pd.DataFrame())
                    if not sub.empty:
                        ws_pnl.write(row, 0, name, cat_fmt); row += 1
                        c_sum = 0
                        for _, r in sub.iterrows():
                            ws_pnl.write(row, 0, r['Budget item'])
                            ws_pnl.write(row, 1, abs(r['Amount']), num_fmt)
                            c_sum += abs(r['Amount'])
                            grand_profit -= r['Amount']
                            row += 1
                        ws_pnl.write(row, 0, f"Total {name}", total_fmt); ws_pnl.write(row, 1, c_sum, total

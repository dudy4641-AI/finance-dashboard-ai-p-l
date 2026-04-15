import streamlit as st
import pandas as pd
import io
import xlsxwriter
import numpy as np

st.set_page_config(page_title="Finance Dashboard AI", layout="wide")

st.title("🚀 מחולל דוחות פיננסיים ו-P&L")
st.write("העלה את קבצי התקציב והתנועות כדי לקבל את האקסל המלוטש.")

def clean_acc(v):
    return str(v).replace('.0', '').strip()

uploaded_files = st.file_uploader("העלה קבצי אקסל (Budget + תנועות)", accept_multiple_files=True, type=['xlsx', 'xls'])

if uploaded_files:
    budget_file = next((f for f in uploaded_files if "budget" in f.name.lower()), None)
    data_files = [f for f in uploaded_files if f != budget_file]
    
    if budget_file and data_files:
        try:
            # עיבוד Budget
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

            all_d = []
            for f in data_files:
                df_raw = pd.read_excel(f)
                if "חשבון" in df_raw.columns or any("תאריך" in str(c) for c in df_raw.columns):
                    d_c = [c for c in df_raw.columns if "תאריך" in str(c)][0]
                    acc = df_raw['חשבון'].apply(clean_acc)
                    temp = pd.DataFrame({'Entity': 'Ltd', 'Date': pd.to_datetime(df_raw[d_c], dayfirst=True, errors='coerce'),
                                         'Vendor': df_raw.get('תאור חשבון נגדי', 'Unknown'), 'Account': (acc + " - " + df_raw['תאור'].fillna('').astype(str)),
                                         'Amount': pd.to_numeric(df_raw['חובה'], errors='coerce').fillna(0) - pd.to_numeric(df_raw['זכות'], errors='coerce').fillna(0),
                                         'Memo': df_raw.get('פרטים', '-'), 'MapKey': acc})
                else:
                    df_inc = pd.read_excel(f, skiprows=4)
                    acc_n = df_inc['Distribution account'].astype(str)
                    acc_num = acc_n.str.extract(r'(\d+)', expand=False).fillna(acc_n).apply(clean_acc)
                    temp = pd.DataFrame({'Entity': 'Inc', 'Date': pd.to_datetime(df_inc['Transaction date'], errors='coerce'),
                                         'Vendor': df_inc['Name'].fillna('Unknown'), 'Account': acc_n,
                                         'Amount': pd.to_numeric(df_inc['Amount'].astype(str).str.replace(r'[\$,",]', '', regex=True), errors='coerce'),
                                         'Memo': df_inc['Memo/Description'].fillna('-'), 'MapKey': acc_num})
                all_d.append(temp)

            final = pd.merge(pd.concat(all_d).dropna(subset=['Date']), df_mapping, on=['Entity', 'MapKey'], how='left')
            final['Account Type'] = final['Account Type'].fillna('P&L')
            final['Budget item'] = final['Budget item'].fillna('Unmapped')
            
            # יצירת הקובץ להורדה
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # לשונית Data
                final[['Entity', 'Account Type', 'Date', 'Vendor', 'Account', 'Amount', 'Budget item', 'Memo']].to_excel(writer, sheet_name='Data', index=False)
                
                # לשונית Executive P&L
                ws_pnl = writer.book.add_worksheet('Executive P&L')
                pnl_summary = final[final['Account Type'] == 'P&L'].groupby('Budget item')['Amount'].sum().reset_index()
                pnl_summary['Amount'] *= -1 # הכנסה בפלוס
                
                header_fmt = writer.book.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white'})
                ws_pnl.write('A1', 'Budget Item', header_fmt)
                ws_pnl.write('B1', 'Total Amount', header_fmt)
                for i, row in pnl_summary.iterrows():
                    ws_pnl.write(i+1, 0, row['Budget item'])
                    ws_pnl.write(i+1, 1, row['Amount'])

            st.success("✅ הקובץ מוכן!")
            st.download_button("📥 הורד אקסל מעוצב", output.getvalue(), "Finance_Dashboard.xlsx")
        except Exception as e:
            st.error(f"שגיאה: {e}")

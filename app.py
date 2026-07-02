import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font

st.title("📊 Trade Analyzer (BUY→SELL / SELL→BUY)")

# ⏱ Time Box (user input in minutes)
time_limit = st.number_input("Time difference limit (minutes)", min_value=1, max_value=120, value=30)

uploaded_file = st.file_uploader("Excel file upload karein", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read first sheet
        df_dict = pd.read_excel(uploaded_file, sheet_name=None)
        sheet_name = list(df_dict.keys())[0]
        df = df_dict[sheet_name]

        df.dropna(subset=['Order Time', 'Execute'], inplace=True)
        df['Order Time'] = pd.to_datetime(df['Order Time'], errors='coerce', dayfirst=True)
        df['Execute'] = pd.to_datetime(df['Execute'], errors='coerce').dt.time

        # ⭐ CLIENT SORT CANCEL — only Symbol sort
        df = df.sort_values(by=['Symbol', 'Order Time'])

        # ---------------------------------------------------------
        # BUY → SELL
        # ---------------------------------------------------------
        buy_sell_results = []
        df_buy_sell = df.copy()

        for (client, symbol), group in df_buy_sell.groupby(['Client', 'Symbol']):
            buys = group[group['B/S'].str.upper() == 'BUY'].copy()
            sells = group[group['B/S'].str.upper() == 'SELL'].copy()

            for buy_idx, buy in buys.iterrows():
                for sell_idx, sell in sells.iterrows():

                    buy_time = datetime.combine(buy['Order Time'].date(), buy['Execute'])
                    sell_time = datetime.combine(sell['Order Time'].date(), sell['Execute'])

                    time_diff = (sell_time - buy_time).total_seconds() / 60

                    if 0 <= time_diff <= time_limit and buy['Qty'] > 0 and sell['Qty'] > 0:

                        qty_traded = min(buy['Qty'], sell['Qty'])
                        rate_diff = round(sell['Order Price'] - buy['Order Price'], 2)
                        profit_loss = round((sell['Order Price'] - buy['Order Price']) * qty_traded, 2)

                        buy_sell_results.append({
                            'Type': 'BUY→SELL',
                            'Client': client,
                            'Symbol': symbol,
                            'Buy Qty': qty_traded,
                            'Sell Qty': qty_traded,
                            'Buy Rate': buy['Order Price'],
                            'Sell Rate': sell['Order Price'],
                            'Rate Diff': rate_diff,
                            'Buy Time': buy_time.strftime("%H:%M:%S"),
                            'Sell Time': sell_time.strftime("%H:%M:%S"),
                            'Time Diff': f"{int(time_diff // 60):02}:{int(time_diff % 60):02}:00",
                            'Profit/Loss': profit_loss
                        })

                        buys.loc[buy_idx, 'Qty'] -= qty_traded
                        sells.loc[sell_idx, 'Qty'] -= qty_traded

                        if buys.loc[buy_idx, 'Qty'] == 0:
                            break

        buy_sell_df = pd.DataFrame(buy_sell_results)

        # ---------------------------------------------------------
        # SELL → BUY
        # ---------------------------------------------------------
        sell_buy_results = []
        df_sell_buy = df.copy()

        for (client, symbol), group in df_sell_buy.groupby(['Client', 'Symbol']):
            sells = group[group['B/S'].str.upper() == 'SELL'].copy()
            buys = group[group['B/S'].str.upper() == 'BUY'].copy()

            for sell_idx, sell in sells.iterrows():
                for buy_idx, buy in buys.iterrows():

                    sell_time = datetime.combine(sell['Order Time'].date(), sell['Execute'])
                    buy_time = datetime.combine(buy['Order Time'].date(), buy['Execute'])

                    time_diff = (buy_time - sell_time).total_seconds() / 60

                    if 0 <= time_diff <= time_limit and sell['Qty'] > 0 and buy['Qty'] > 0:

                        qty_traded = min(sell['Qty'], buy['Qty'])
                        rate_diff = round(sell['Order Price'] - buy['Order Price'], 2)
                        profit_loss = round((sell['Order Price'] - buy['Order Price']) * qty_traded, 2)

                        sell_buy_results.append({
                            'Type': 'SELL→BUY',
                            'Client': client,
                            'Symbol': symbol,
                            'Buy Qty': qty_traded,
                            'Sell Qty': qty_traded,
                            'Sell Rate': sell['Order Price'],
                            'Buy Rate': buy['Order Price'],
                            'Rate Diff': rate_diff,
                            'Sell Time': sell_time.strftime("%H:%M:%S"),
                            'Buy Time': buy_time.strftime("%H:%M:%S"),
                            'Time Diff': f"{int(time_diff // 60):02}:{int(time_diff % 60):02}:00",
                            'Profit/Loss': profit_loss
                        })

                        sells.loc[sell_idx, 'Qty'] -= qty_traded
                        buys.loc[buy_idx, 'Qty'] -= qty_traded

                        if sells.loc[sell_idx, 'Qty'] == 0:
                            break

        sell_buy_df = pd.DataFrame(sell_buy_results)

        # ---------------------------------------------------------
        # MERGE BOTH INTO ONE SHEET
        # ---------------------------------------------------------
        final_df = pd.concat([buy_sell_df, sell_buy_df], ignore_index=True)

        # ⭐ FINAL SORT — ONLY SYMBOL A→Z
        final_df = final_df.sort_values(by=['Symbol'])

        # REMOVE NEGATIVE PROFIT/LOSS
        final_df = final_df[final_df['Profit/Loss'] > 0]

        # Symbol-wise blank row
        final_output = []
        prev_symbol = None

        for _, row in final_df.iterrows():
            if prev_symbol is not None and prev_symbol != row['Symbol']:
                final_output.append({col: '' for col in final_df.columns})
            final_output.append(row.to_dict())
            prev_symbol = row['Symbol']

        final_df = pd.DataFrame(final_output)

        # ---------------------------------------------------------
        # Excel in memory (for download)
        # ---------------------------------------------------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='FINAL_TRADES')

        output.seek(0)

        st.success("✅ Processing complete! Niche se file download karein.")
        st.download_button(
            label="⬇ Download FINAL Excel",
            data=output,
            file_name="FINAL_TRADES.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("Preview:")
        st.dataframe(final_df)

    except Exception as e:
        st.error(f"❌ Error: {e}")

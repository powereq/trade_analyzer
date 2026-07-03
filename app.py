import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# -----------------------------
# 🔥 STYLISH PAGE CONFIG
# -----------------------------
st.set_page_config(
    page_title="Trade Analyzer",
    page_icon="📊",
    layout="wide"
)

# -----------------------------
# 🔥 CUSTOM CSS (STYLISH UI)
# -----------------------------
st.markdown("""
<style>

body {
    background-color: #f5f7fa;
}

.big-title {
    font-size: 42px;
    font-weight: 800;
    color: #0047AB;
    text-align: center;
    margin-bottom: 10px;
}

.sub-title {
    font-size: 20px;
    font-weight: 600;
    color: #333;
    text-align: center;
    margin-bottom: 30px;
}

.box {
    padding: 20px;
    background: white;
    border-radius: 12px;
    border: 1px solid #e0e0e0;
    box-shadow: 0px 3px 10px rgba(0,0,0,0.05);
    margin-bottom: 20px;
}

.card {
    padding: 18px;
    background: #ffffff;
    border-radius: 12px;
    border: 1px solid #d0d0d0;
    text-align: center;
    box-shadow: 0px 3px 10px rgba(0,0,0,0.08);
}

.card-title {
    font-size: 18px;
    font-weight: 600;
    color: #0047AB;
}

.card-value {
    font-size: 26px;
    font-weight: 700;
    color: #111;
}

</style>
""", unsafe_allow_html=True)

# -----------------------------
# 🔥 HEADER
# -----------------------------
st.markdown("<div class='big-title'>📊 Trade Analyzer</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-title'>BUY→SELL & SELL→BUY Auto Matching + PNL Summary Dashboard</div>", unsafe_allow_html=True)

# -----------------------------
# 🔥 INPUT BOXES (TIME + PNL)
# -----------------------------
with st.container():
    st.markdown("<div class='box'>", unsafe_allow_html=True)

    time_limit = st.number_input(
        "⏱ Time difference limit (minutes)",
        min_value=1, max_value=500, value=30
    )

    pnl_limit = st.number_input(
        "💰 Minimum PNL filter (₹)",
        min_value=0, max_value=1000000, value=1000
    )

    uploaded_file = st.file_uploader("📁 Excel file upload karein", type=["xlsx"])

    st.markdown("</div>", unsafe_allow_html=True)

# -----------------------------
# 🔥 PROCESSING LOGIC
# -----------------------------
if uploaded_file is not None:
    try:
        df_dict = pd.read_excel(uploaded_file, sheet_name=None)
        sheet_name = list(df_dict.keys())[0]
        df = df_dict[sheet_name]

        df.dropna(subset=['Order Time', 'Execute'], inplace=True)

        df['Order Time'] = pd.to_datetime(df['Order Time'], errors='coerce', dayfirst=True)
        df['Execute'] = pd.to_datetime(df['Execute'], errors='coerce').dt.time

        df['Date'] = df['Order Time'].dt.date

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
                            'Date': buy['Date'],
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
                            'Date': sell['Date'],
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
        # MERGE BOTH
        # ---------------------------------------------------------
        final_df = pd.concat([buy_sell_df, sell_buy_df], ignore_index=True)

        final_df = final_df.sort_values(by=['Symbol'])

        final_df = final_df[final_df['Profit/Loss'] >= pnl_limit]

        # ---------------------------------------------------------
        # ⭐ SYMBOL‑WISE BLANK ROW LOGIC
        # ---------------------------------------------------------
        final_output = []
        prev_symbol = None

        for _, row in final_df.iterrows():
            if prev_symbol is not None and prev_symbol != row['Symbol']:
                final_output.append({col: '' for col in final_df.columns})
            final_output.append(row.to_dict())
            prev_symbol = row['Symbol']

        final_df = pd.DataFrame(final_output)

        # ---------------------------------------------------------
        # ⭐ SUMMARY CALCULATIONS
        # ---------------------------------------------------------
        symbol_summary = final_df.groupby('Symbol')['Profit/Loss'].sum().reset_index()
        client_summary = final_df.groupby('Client')['Profit/Loss'].sum().reset_index()
        grand_total = final_df['Profit/Loss'].sum()

        # ---------------------------------------------------------
        # ⭐ SUMMARY CARDS
        # ---------------------------------------------------------
        col1, col2, col3 = st.columns(3)

        col1.markdown(f"<div class='card'><div class='card-title'>Total Symbols</div><div class='card-value'>{symbol_summary.shape[0]}</div></div>", unsafe_allow_html=True)
        col2.markdown(f"<div class='card'><div class='card-title'>Total Clients</div><div class='card-value'>{client_summary.shape[0]}</div></div>", unsafe_allow_html=True)
        col3.markdown(f"<div class='card'><div class='card-title'>Grand Total PNL</div><div class='card-value'>₹ {grand_total:,.2f}</div></div>", unsafe_allow_html=True)

        # ---------------------------------------------------------
        # DOWNLOAD + PREVIEW
        # ---------------------------------------------------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='FINAL_TRADES')
            symbol_summary.to_excel(writer, index=False, sheet_name='SYMBOL_SUMMARY')
            client_summary.to_excel(writer, index=False, sheet_name='CLIENT_SUMMARY')

        output.seek(0)

        st.success("✅ Processing complete! Niche se file download karein.")

        st.download_button(
            label="⬇ Download FINAL Excel (with summaries)",
            data=output,
            file_name="FINAL_TRADES.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("📄 Preview:")
        st.dataframe(final_df)

        st.subheader("📘 Symbol-wise Summary")
        st.dataframe(symbol_summary)

        st.subheader("📙 Client-wise Summary")
        st.dataframe(client_summary)

    except Exception as e:
        st.error(f"❌ Error: {e}")

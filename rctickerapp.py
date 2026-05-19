import requests
import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="RC Risk Monitor", layout="wide")

st.title("RC Risk Monitor (ticker version)")

isin_input = st.text_area(
    "Enter ISIN codes separated by commas",
    placeholder="CH1234567890, CH0987654321"
)

def create_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Products")
        worksheet = writer.sheets["Products"]

        percent_columns = ["Strike", "Strike Level Distance", "Performance"]

        for col_name in percent_columns:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name) + 1
                for row in range(2, len(df) + 2):
                    worksheet.cell(row=row, column=col_idx).number_format = "0.00%"

        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))

            worksheet.column_dimensions[column_letter].width = max_length + 2

    output.seek(0)
    return output


if st.button("Run Risk Monitor"):
    if not isin_input.strip():
        st.warning("Please enter at least one ISIN.")
    else:
        isins = [isin.strip().upper() for isin in isin_input.split(",")]

        all_rows = []

        with st.spinner("Fetching product data..."):
            for isin in isins:
                url = f"https://structuredproducts-ch.leonteq.com/isin/{isin}/json"

                try:
                    response = requests.get(url, timeout=20)
                    response.raise_for_status()
                    data = response.json()

                    product = data.get("product", {})

                    product_name = product.get("identification", {}).get("name")
                    initial_fixing_date = product.get("calendar", {}).get("initialFixingDate")
                    last_trading_date = product.get("calendar", {}).get("lastTradingDate")
                    investment_currency = product.get("investment", {}).get("investmentCurrency")

                    baskets = product.get("baskets", [])

                    for basket in baskets:
                        levels = basket.get("levels", {})
                        strike = pd.to_numeric(levels.get("strike"), errors="coerce")

                        underlyings = basket.get("underlyings", [])

                        for underlying in underlyings:
                            underlying_feed = underlying.get("underlyingFeed", {})
                            dynamic = underlying.get("dynamic", {})
                            identifiers = underlying.get("identifiers", {})
                            ticker = identifiers.get("bloombergTicker", {})

                            initial = pd.to_numeric(
                                underlying.get("initialFixingLevelAbs"),
                                errors="coerce"
                            )

                            spot = pd.to_numeric(
                                underlying_feed.get("spot"),
                                errors="coerce"
                            )

                            strike_level_abs = pd.to_numeric(
                                dynamic.get("strikeLevelAbs"),
                                errors="coerce"
                            )

                            strike_level_distance = pd.to_numeric(
                                dynamic.get("strikeLevelDistance"),
                                errors="coerce"
                            )

                            spot_below_strike = (
                                "Yes"
                                if pd.notna(spot)
                                and pd.notna(strike_level_abs)
                                and spot < strike_level_abs
                                else "No"
                            )

                            performance = (
                                (spot - initial) / initial
                                if pd.notna(spot) and pd.notna(initial) and initial != 0
                                else np.nan
                            )

                            all_rows.append({
                                "ISIN": isin,
                                "Product Name": product_name,
                                "Initial Fixing Date": initial_fixing_date,
                                "Last Trading Date": last_trading_date,
                                "Ticker": ticker,
                                "Initial Fixing Level": initial,
                                "Spot": spot,
                                "Strike Level Abs": strike_level_abs,
                                "Strike Level Distance": strike_level_distance,
                                "Strike": strike,
                                "Investment Currency": investment_currency,
                                "Spot Below Strike": spot_below_strike,
                                "Performance": performance
                            })

                except Exception as e:
                    all_rows.append({
                        "ISIN": isin,
                        "Product Name": f"ERROR: {e}",
                        "Initial Fixing Date": None,
                        "Last Trading Date": None,
                        "Ticker": None,
                        "Initial Fixing Level": None,
                        "Spot": None,
                        "Strike Level Abs": None,
                        "Strike Level Distance": None,
                        "Strike": None,
                        "Investment Currency": None,
                        "Spot Below Strike": None,
                        "Performance": np.nan
                    })

        df = pd.DataFrame(all_rows)

        df = (
            df.sort_values(by=["ISIN", "Performance"], ascending=[True, True])
              .groupby("ISIN", as_index=False)
              .first()
        )

        st.success("Done.")
        st.dataframe(df, use_container_width=True)

        excel_file = create_excel(df)

        st.download_button(
            label="Download Excel file",
            data=excel_file,
            file_name="rc_risk_monitor.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

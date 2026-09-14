# ============================================================
# 20-DAY SUPPORT / RESISTANCE ZONE SCANNER
# GOOGLE COLAB
# ============================================================
!pip -q install openpyxl pandas numpy

import openpyxl
import pandas as pd
import numpy as np
from google.colab import files

# ============================================================
# SETTINGS
# ============================================================

ZONE_TOLERANCE = 0.0075       # 0.75% price clustering
MIN_TOUCHES = 2               # minimum touches for a zone
TOP_ZONES = 3                 # top 3 support + resistance
RECENT_WEIGHT_DAYS = 20

# Scores
TOUCH_SCORE_MAX = 30
RECENCY_SCORE_MAX = 20
REJECTION_SCORE_MAX = 20
O2_SCORE_MAX = 15
C2O_SCORE_MAX = 10
EXTRA_SCORE_MAX = 5

# ============================================================
# UPLOAD
# ============================================================

print("Upload your 20-day Excel file")
uploaded = files.upload()
input_file = next(iter(uploaded))

print("\nLoading:", input_file)

wb = openpyxl.load_workbook(
    input_file,
    data_only=False
)

# ============================================================
# DELETE OLD SR_ZONES
# ============================================================

if "SR_Zones" in wb.sheetnames:
    del wb["SR_Zones"]

# ============================================================
# FIRST 20 SHEETS
# ============================================================

sheets = wb.worksheets[:]

if len(sheets) < 20:
    raise ValueError(
        f"Only {len(sheets)} sheets found. "
        "At least 20 sheets are required."
    )

sheets20 = sheets[:20]

print("\n20 sheets being used:")

for i, ws in enumerate(sheets20, 1):
    print(f"{i}: {ws.title}")

# ============================================================
# HEADER NORMALIZATION
# ============================================================

def normalize_header(x):
    if x is None:
        return ""

    x = str(x).strip().upper()
    x = x.replace("%", "")
    x = x.replace(" ", "")
    x = x.replace("_", "")
    x = x.replace("-", "")

    return x


# ============================================================
# FIND COLUMN
# ============================================================

def find_column(ws, possible_names):

    headers = {}

    for cell in ws[1]:
        headers[normalize_header(cell.value)] = cell.column

    for name in possible_names:

        key = normalize_header(name)

        if key in headers:
            return headers[key]

    return None


# ============================================================
# COLUMN IDENTIFICATION
# ============================================================

def get_columns(ws):

    symbol_col = find_column(
        ws,
        [
            "SYMBOL",
            "STOCK",
            "SCRIP",
            "TICKER",
            "NAME"
        ]
    )

    open_col = find_column(
        ws,
        [
            "OPEN",
            "O"
        ]
    )

    high_col = find_column(
        ws,
        [
            "HIGH",
            "H"
        ]
    )

    low_col = find_column(
        ws,
        [
            "LOW",
            "L"
        ]
    )

    prev_close_col = find_column(
        ws,
        [
            "PREV CLOSE",
            "PREVCLOSE",
            "PREVIOUS CLOSE",
            "PREVIOUSCLOSE"
        ]
    )

    o2h_col = find_column(
        ws,
        [
            "O2H",
            "O2H%",
            "OPEN2HIGH",
            "OPENHIGH"
        ]
    )

    o2l_col = find_column(
        ws,
        [
            "O2L",
            "O2L%",
            "OPEN2LOW",
            "OPENLOW"
        ]
    )

    c2o_col = find_column(
        ws,
        [
            "C2O",
            "C2O%",
            "LTP-OPEN",
            "LTP_OPEN",
            "CLOSE2OPEN"
        ]
    )

    return {
        "symbol": symbol_col,
        "open": open_col,
        "high": high_col,
        "low": low_col,
        "prev_close": prev_close_col,
        "o2h": o2h_col,
        "o2l": o2l_col,
        "c2o": c2o_col
    }


# ============================================================
# CHECK COLUMNS
# ============================================================

print("\nColumn detection:")

for ws in sheets20:

    cols = get_columns(ws)

    print(
        ws.title,
        "=>",
        cols
    )

    if cols["symbol"] is None:
        raise ValueError(
            f"Symbol column not found in sheet: {ws.title}"
        )

    if cols["high"] is None:
        raise ValueError(
            f"High column not found in sheet: {ws.title}"
        )

    if cols["low"] is None:
        raise ValueError(
            f"Low column not found in sheet: {ws.title}"
        )

    if cols["open"] is None:
        raise ValueError(
            f"Open column not found in sheet: {ws.title}"
        )


# ============================================================
# READ ALL 20 DAYS
# ============================================================

records = []

for day_index, ws in enumerate(sheets20, 1):

    cols = get_columns(ws)

    for row in range(2, ws.max_row + 1):

        symbol = ws.cell(
            row,
            cols["symbol"]
        ).value

        if symbol is None:
            continue

        symbol = str(symbol).strip()

        if symbol == "":
            continue

        def get_value(col):

            if col is None:
                return np.nan

            value = ws.cell(row, col).value

            if value is None:
                return np.nan

            try:
                return float(value)

            except:

                try:
                    return float(
                        str(value)
                        .replace(",", "")
                        .replace("%", "")
                        .strip()
                    )

                except:
                    return np.nan

        o = get_value(cols["open"])
        h = get_value(cols["high"])
        l = get_value(cols["low"])
        pc = get_value(cols["prev_close"])
        o2h = get_value(cols["o2h"])
        o2l = get_value(cols["o2l"])
        c2o = get_value(cols["c2o"])

        # ----------------------------------------------------
        # Calculate missing O2H
        # ----------------------------------------------------

        if pd.isna(o2h):

            if (
                not pd.isna(o)
                and not pd.isna(h)
                and o != 0
            ):

                o2h = (
                    (h - o)
                    / o
                    * 100
                )

        # ----------------------------------------------------
        # Calculate missing O2L
        # ----------------------------------------------------

        if pd.isna(o2l):

            if (
                not pd.isna(o)
                and not pd.isna(l)
                and o != 0
            ):

                o2l = (
                    (l - o)
                    / o
                    * 100
                )

        # ----------------------------------------------------
        # C2O
        #
        # If column unavailable, use a CLOSE/LTP column.
        # Otherwise remains blank.
        # ----------------------------------------------------

        if pd.isna(c2o):

            close_col = find_column(
                ws,
                [
                    "LTP",
                    "PRICE",
                    "CLOSE",
                    "LAST PRICE",
                    "LASTPRICE"
                ]
            )

            if close_col is not None:

                close_value = get_value(close_col)

                if (
                    not pd.isna(close_value)
                    and not pd.isna(o)
                    and o != 0
                ):

                    c2o = (
                        (close_value - o)
                        / o
                        * 100
                    )

        records.append(
            {
                "symbol": symbol,
                "day": day_index,
                "sheet": ws.title,
                "open": o,
                "high": h,
                "low": l,
                "prev_close": pc,
                "o2h": o2h,
                "o2l": o2l,
                "c2o": c2o
            }
        )

df = pd.DataFrame(records)

print("\nTotal records:", len(df))


# ============================================================
# NUMERIC CONVERSION
# ============================================================

numeric_cols = [
    "open",
    "high",
    "low",
    "prev_close",
    "o2h",
    "o2l",
    "c2o"
]

for col in numeric_cols:

    df[col] = pd.to_numeric(
        df[col],
        errors="coerce"
    )


# ============================================================
# ZONE CREATION FUNCTION
# ============================================================

def create_zones(prices, records_for_stock, zone_type):

    prices = [
        float(x)
        for x in prices
        if not pd.isna(x)
        and x > 0
    ]

    if len(prices) == 0:
        return []

    # --------------------------------------------------------
    # SORT PRICES
    # --------------------------------------------------------

    prices_sorted = sorted(prices)

    clusters = []

    current_cluster = [
        prices_sorted[0]
    ]

    # --------------------------------------------------------
    # CLUSTER NEARBY PRICES
    # --------------------------------------------------------

    for price in prices_sorted[1:]:

        cluster_mean = np.mean(
            current_cluster
        )

        if (
            abs(price - cluster_mean)
            / cluster_mean
            <= ZONE_TOLERANCE
        ):

            current_cluster.append(price)

        else:

            clusters.append(
                current_cluster
            )

            current_cluster = [
                price
            ]

    clusters.append(
        current_cluster
    )

    zones = []

    # --------------------------------------------------------
    # EACH CLUSTER
    # --------------------------------------------------------

    for cluster in clusters:

        touches = len(cluster)

        if touches < MIN_TOUCHES:
            continue

        zone_low = min(cluster)
        zone_high = max(cluster)

        midpoint = (
            zone_low + zone_high
        ) / 2

        # ----------------------------------------------------
        # TOUCH INFORMATION
        # ----------------------------------------------------

        touch_rows = []

        for _, row in records_for_stock.iterrows():

            if zone_type == "RESISTANCE":
                price = row["high"]

            else:
                price = row["low"]

            if pd.isna(price):
                continue

            if (
                abs(price - midpoint)
                / midpoint
                <= ZONE_TOLERANCE
            ):

                touch_rows.append(row)

        touch_df = pd.DataFrame(
            touch_rows
        )

        # ----------------------------------------------------
        # RECENCY
        #
        # day 1 = newest
        # day 20 = oldest
        # ----------------------------------------------------

        if len(touch_df) > 0:

            recent_days = (
                touch_df["day"]
                .tolist()
            )

            recency_raw = np.mean(
                [
                    (
                        RECENT_WEIGHT_DAYS
                        - d
                        + 1
                    )
                    / RECENT_WEIGHT_DAYS
                    for d in recent_days
                ]
            )

        else:

            recency_raw = 0

        recency_score = (
            recency_raw
            * RECENCY_SCORE_MAX
        )

        # ----------------------------------------------------
        # TOUCH SCORE
        # ----------------------------------------------------

        touch_score = min(
            touches / 6,
            1
        ) * TOUCH_SCORE_MAX

        # ----------------------------------------------------
        # REJECTION SCORE
        # ----------------------------------------------------

        rejection_score = 0

        if len(touch_df) > 0:

            if zone_type == "RESISTANCE":

                rejection_values = []

                for _, rr in touch_df.iterrows():

                    o2h = rr["o2h"]
                    c2o = rr["c2o"]

                    if (
                        not pd.isna(o2h)
                        and not pd.isna(c2o)
                    ):

                        if c2o < o2h:
                            rejection_values.append(1)

                if rejection_values:

                    rejection_score = (
                        np.mean(
                            rejection_values
                        )
                        * REJECTION_SCORE_MAX
                    )

            else:

                rejection_values = []

                for _, rr in touch_df.iterrows():

                    o2l = rr["o2l"]
                    c2o = rr["c2o"]

                    if (
                        not pd.isna(o2l)
                        and not pd.isna(c2o)
                    ):

                        if c2o > o2l:
                            rejection_values.append(1)

                if rejection_values:

                    rejection_score = (
                        np.mean(
                            rejection_values
                        )
                        * REJECTION_SCORE_MAX
                    )

        # ----------------------------------------------------
        # O2 SCORE
        # ----------------------------------------------------

        o2_score = 0

        if len(touch_df) > 0:

            if zone_type == "RESISTANCE":

                vals = touch_df[
                    "o2h"
                ].dropna()

                if len(vals) > 0:

                    consistency = (
                        1
                        /
                        (
                            1
                            + vals.std()
                        )
                    )

                    o2_score = min(
                        consistency,
                        1
                    ) * O2_SCORE_MAX

            else:

                vals = touch_df[
                    "o2l"
                ].dropna()

                if len(vals) > 0:

                    consistency = (
                        1
                        /
                        (
                            1
                            + abs(vals.std())
                        )
                    )

                    o2_score = min(
                        consistency,
                        1
                    ) * O2_SCORE_MAX

        # ----------------------------------------------------
        # C2O SCORE
        # ----------------------------------------------------

        c2o_score = 0

        if len(touch_df) > 0:

            c2o_values = touch_df[
                "c2o"
            ].dropna()

            if len(c2o_values) > 0:

                if zone_type == "RESISTANCE":

                    rejected = (
                        c2o_values < 0
                    )

                else:

                    rejected = (
                        c2o_values > 0
                    )

                c2o_score = (
                    rejected.mean()
                    * C2O_SCORE_MAX
                )

        # ----------------------------------------------------
        # EXTRA SCORE
        # ----------------------------------------------------

        extra_score = min(
            touches / 10,
            1
        ) * EXTRA_SCORE_MAX

        # ----------------------------------------------------
        # TOTAL SCORE
        # ----------------------------------------------------

        total_score = (
            touch_score
            + recency_score
            + rejection_score
            + o2_score
            + c2o_score
            + extra_score
        )

        total_score = min(
            round(total_score, 1),
            100
        )

        # ----------------------------------------------------
        # STRENGTH
        # ----------------------------------------------------

        if total_score >= 80:
            strength = "VERY STRONG"

        elif total_score >= 65:
            strength = "STRONG"

        elif total_score >= 50:
            strength = "MODERATE"

        else:
            strength = "WEAK"

        zones.append(
            {
                "zone_type": zone_type,
                "zone_low": zone_low,
                "zone_high": zone_high,
                "zone_mid": midpoint,
                "touches": touches,
                "score": total_score,
                "strength": strength,
                "touch_score": round(
                    touch_score,
                    1
                ),
                "recency_score": round(
                    recency_score,
                    1
                ),
                "rejection_score": round(
                    rejection_score,
                    1
                ),
                "o2_score": round(
                    o2_score,
                    1
                ),
                "c2o_score": round(
                    c2o_score,
                    1
                )
            }
        )

    # --------------------------------------------------------
    # SORT BY SCORE
    # --------------------------------------------------------

    zones = sorted(
        zones,
        key=lambda x: (
            x["score"],
            x["touches"]
        ),
        reverse=True
    )

    return zones[:TOP_ZONES]


# ============================================================
# BUILD SR RESULTS
# ============================================================

results = []

symbols = sorted(
    df["symbol"]
    .dropna()
    .unique()
)

print(
    "\nProcessing",
    len(symbols),
    "stocks..."
)

for symbol in symbols:

    stock = df[
        df["symbol"] == symbol
    ].copy()

    # --------------------------------------------------------
    # CURRENT PRICE
    # --------------------------------------------------------

    latest = stock[
        stock["day"] == 1
    ]

    if len(latest) == 0:
        continue

    latest_row = latest.iloc[0]

    current_price = np.nan

    ws_latest = sheets20[0]

    close_col = find_column(
        ws_latest,
        [
            "LTP",
            "PRICE",
            "CLOSE",
            "LAST PRICE",
            "LASTPRICE"
        ]
    )

    if close_col is not None:

        symbol_col = find_column(
            ws_latest,
            [
                "SYMBOL",
                "STOCK",
                "SCRIP",
                "TICKER",
                "NAME"
            ]
        )

        if symbol_col is not None:

            for rr in range(
                2,
                ws_latest.max_row + 1
            ):

                s = ws_latest.cell(
                    rr,
                    symbol_col
                ).value

                if (
                    s is not None
                    and str(s).strip() == symbol
                ):

                    try:
                        current_price = float(
                            ws_latest.cell(
                                row=rr,
                                column=close_col
                            ).value
                        )

                    except:

                        try:
                            current_price = float(
                                str(
                                    ws_latest.cell(
                                        row=rr,
                                        column=close_col
                                    ).value
                                )
                                .replace(",", "")
                                .replace("%", "")
                                .strip()
                            )

                        except:
                            current_price = np.nan

                    break

    # Fallback to the most recent open if no current/LTP
    # value is available.
    if pd.isna(current_price):

        if not pd.isna(latest_row["open"]):
            current_price = float(
                latest_row["open"]
            )

    # --------------------------------------------------------
    # SUPPORT / RESISTANCE ZONES
    # --------------------------------------------------------

    support_zones = create_zones(
        stock["low"].tolist(),
        stock,
        "SUPPORT"
    )

    resistance_zones = create_zones(
        stock["high"].tolist(),
        stock,
        "RESISTANCE"
    )

    # --------------------------------------------------------
    # ORIGINAL SUMMARY STATISTICS
    # --------------------------------------------------------

    o2h_values = stock["o2h"].dropna()
    o2l_values = stock["o2l"].dropna()
    c2o_values = stock["c2o"].dropna()

    o2h_median = (
        float(o2h_values.median())
        if len(o2h_values) > 0
        else np.nan
    )

    o2h_80 = (
        float(o2h_values.quantile(0.80))
        if len(o2h_values) > 0
        else np.nan
    )

    o2l_median = (
        float(o2l_values.median())
        if len(o2l_values) > 0
        else np.nan
    )

    o2l_20 = (
        float(o2l_values.quantile(0.20))
        if len(o2l_values) > 0
        else np.nan
    )

    c2o_median = (
        float(c2o_values.median())
        if len(c2o_values) > 0
        else np.nan
    )

    # --------------------------------------------------------
    # ORIGINAL NEAREST RESISTANCE / SUPPORT
    #
    # Nearest resistance = lowest zone midpoint >= price
    # Nearest support    = highest zone midpoint <= price
    # --------------------------------------------------------

    nearest_resistance = np.nan
    nearest_support = np.nan

    if not pd.isna(current_price):

        resistance_above = [
            z["zone_mid"]
            for z in resistance_zones
            if z["zone_mid"] >= current_price
        ]

        support_below = [
            z["zone_mid"]
            for z in support_zones
            if z["zone_mid"] <= current_price
        ]

        if resistance_above:
            nearest_resistance = min(
                resistance_above
            )

        if support_below:
            nearest_support = max(
                support_below
            )

    # --------------------------------------------------------
    # DISTANCE TO NEAREST RESISTANCE / SUPPORT
    # --------------------------------------------------------

    distance_to_resistance = np.nan
    distance_to_support = np.nan

    if (
        not pd.isna(current_price)
        and current_price != 0
    ):

        if not pd.isna(nearest_resistance):

            distance_to_resistance = (
                (
                    nearest_resistance
                    - current_price
                )
                / current_price
                * 100
            )

        if not pd.isna(nearest_support):

            distance_to_support = (
                (
                    current_price
                    - nearest_support
                )
                / current_price
                * 100
            )

    # --------------------------------------------------------
    # SR BIAS
    # --------------------------------------------------------

    if (
        not pd.isna(nearest_resistance)
        and not pd.isna(nearest_support)
    ):

        if (
            distance_to_support
            < distance_to_resistance
        ):
            sr_bias = "RESISTANCE NEAR"

        elif (
            distance_to_resistance
            < distance_to_support
        ):
            sr_bias = "SUPPORT NEAR"

        else:
            sr_bias = "EQUILIBRIUM"

    elif not pd.isna(nearest_resistance):

        sr_bias = "RESISTANCE ONLY"

    elif not pd.isna(nearest_support):

        sr_bias = "SUPPORT ONLY"

    else:

        sr_bias = "NO NEAR SR"

    # --------------------------------------------------------
    # BUILD ORIGINAL OUTPUT ROW
    # --------------------------------------------------------

    result = {
        "Symbol": symbol,
        "Current Price": current_price
    }

    # --------------------------------------------------------
    # RESISTANCE 1-3
    # --------------------------------------------------------

    for i in range(TOP_ZONES):

        n = i + 1

        if i < len(resistance_zones):

            z = resistance_zones[i]

            result[f"Resistance {n} Low"] = z["zone_low"]
            result[f"Resistance {n} High"] = z["zone_high"]
            result[f"Resistance {n} Mid"] = z["zone_mid"]
            result[f"Resistance {n} Touches"] = z["touches"]
            result[f"Resistance {n} Score"] = z["score"]
            result[f"Resistance {n} Strength"] = z["strength"]

        else:

            result[f"Resistance {n} Low"] = np.nan
            result[f"Resistance {n} High"] = np.nan
            result[f"Resistance {n} Mid"] = np.nan
            result[f"Resistance {n} Touches"] = np.nan
            result[f"Resistance {n} Score"] = np.nan
            result[f"Resistance {n} Strength"] = ""

    # --------------------------------------------------------
    # SUPPORT 1-3
    # --------------------------------------------------------

    for i in range(TOP_ZONES):

        n = i + 1

        if i < len(support_zones):

            z = support_zones[i]

            result[f"Support {n} Low"] = z["zone_low"]
            result[f"Support {n} High"] = z["zone_high"]
            result[f"Support {n} Mid"] = z["zone_mid"]
            result[f"Support {n} Touches"] = z["touches"]
            result[f"Support {n} Score"] = z["score"]
            result[f"Support {n} Strength"] = z["strength"]

        else:

            result[f"Support {n} Low"] = np.nan
            result[f"Support {n} High"] = np.nan
            result[f"Support {n} Mid"] = np.nan
            result[f"Support {n} Touches"] = np.nan
            result[f"Support {n} Score"] = np.nan
            result[f"Support {n} Strength"] = ""

    # --------------------------------------------------------
    # ORIGINAL ADDITIONAL OUTPUT
    # --------------------------------------------------------

    result["O2H Median"] = o2h_median
    result["O2H 80%"] = o2h_80

    result["O2L Median"] = o2l_median
    result["O2L 20%"] = o2l_20

    result["C2O Median"] = c2o_median

    result["Nearest Resistance"] = nearest_resistance
    result["Distance to Resistance %"] = distance_to_resistance

    result["Nearest Support"] = nearest_support
    result["Distance to Support %"] = distance_to_support

    result["SR Bias"] = sr_bias

    results.append(result)


# ============================================================
# CREATE SR_ZONES SHEET
# ============================================================

ws_sr = wb.create_sheet("SR_Zones")

if results:

    result_df = pd.DataFrame(results)

    # Write headers
    for col_num, column_name in enumerate(
        result_df.columns,
        1
    ):

        ws_sr.cell(
            row=1,
            column=col_num,
            value=column_name
        )

    # Write values
    for row_num, row_data in enumerate(
        result_df.itertuples(index=False),
        2
    ):

        for col_num, value in enumerate(
            row_data,
            1
        ):

            if pd.isna(value):
                value = None

            ws_sr.cell(
                row=row_num,
                column=col_num,
                value=value
            )

else:

    ws_sr.cell(
        row=1,
        column=1,
        value="No SR zones found"
    )


# ============================================================
# FORMAT SR_ZONES
# ============================================================

from openpyxl.styles import Font, PatternFill, Alignment

header_fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

header_font = Font(
    color="FFFFFF",
    bold=True
)

for cell in ws_sr[1]:

    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

ws_sr.freeze_panes = "A2"
ws_sr.auto_filter.ref = ws_sr.dimensions

for row in ws_sr.iter_rows(
    min_row=2,
    max_row=ws_sr.max_row
):

    for cell in row:

        if isinstance(
            cell.value,
            (int, float)
        ):

            cell.number_format = "0.00"


for column_cells in ws_sr.columns:

    max_length = 0
    column_letter = column_cells[0].column_letter

    for cell in column_cells:

        try:
            length = len(str(cell.value))

            if length > max_length:
                max_length = length

        except:
            pass

    ws_sr.column_dimensions[
        column_letter
    ].width = min(
        max(max_length + 2, 10),
        25
    )

# ============================================================
# SAVE
# ============================================================

output_file = (
    "SR_Zones_20_Day.xlsx"
)

wb.save(output_file)

print(
    "\n============================================================"
)

print(
    "SR ZONES COMPLETED"
)

print(
    "Stocks processed:",
    len(results)
)

print(
    "Output:",
    output_file
)

print(
    "============================================================"
)

# ============================================================
# DOWNLOAD
# ============================================================

files.download(output_file)

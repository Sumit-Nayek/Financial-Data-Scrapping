import requests
import pandas as pd

from bs4 import BeautifulSoup
from datetime import datetime
from pathlib import Path


# ============================================================
# 1. CONFIGURATION
# ============================================================

BASE_URL = "https://www.indiabondinfo.nsdl.com/"

OUTPUT_FOLDER = Path("data")
OUTPUT_FILE = OUTPUT_FOLDER / "bond_data.xlsx"

OUTPUT_FOLDER.mkdir(exist_ok=True)


# ------------------------------------------------------------
# Search filters
# ------------------------------------------------------------

INSTRUMENT_TYPE = ""
ALLOTMENT_DATE_FROM = ""
ALLOTMENT_DATE_TO = ""
MATURITY_DATE_FROM = ""
MATURITY_DATE_TO = ""


# ------------------------------------------------------------
# Website headers
# ------------------------------------------------------------

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 "
        "(Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/120.0 Safari/537.36"
    )
}


# ============================================================
# 2. CREATE SESSION
# ============================================================

session = requests.Session()
session.headers.update(HEADERS)


# ============================================================
# 3. CHECK WHETHER TODAY IS A WORKING DAY
# ============================================================

def is_market_day():

    today = datetime.today()

    # Saturday = 5
    # Sunday = 6
    if today.weekday() >= 5:
        return False, "Weekend"

    return True, "Working Day"


# ============================================================
# 4. DOWNLOAD WEB PAGE
# ============================================================

def get_page(url):

    try:

        response = session.get(
            url,
            timeout=30
        )

        response.raise_for_status()

        print("Page downloaded successfully.")

        return response.text

    except requests.RequestException as error:

        print("Error while downloading page:")
        print(error)

        return None


# ============================================================
# 5. EXTRACT TABLES
# ============================================================

def extract_tables(html):

    if html is None:
        return []

    try:

        tables = pd.read_html(html)

        print(f"Number of tables found: {len(tables)}")

        return tables

    except ValueError:

        print("No HTML tables found.")

        return []


# ============================================================
# 6. CLEAN DATA
# ============================================================

def clean_data(df):

    # Remove completely empty rows
    df = df.dropna(how="all")

    # Remove duplicate rows
    df = df.drop_duplicates()

    # Clean column names
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # Remove extra spaces from text
    for column in df.columns:

        if df[column].dtype == "object":

            df[column] = (
                df[column]
                .astype(str)
                .str.strip()
            )

    return df


# ============================================================
# 7. STANDARDIZE COLUMN NAMES
# ============================================================

def standardize_columns(df):

    column_mapping = {

        "ISIN No.": "ISIN",
        "ISIN Number": "ISIN",

        "Issuer Name": "Issuer",
        "Company Name": "Issuer",

        "Instrument Type": "Type of Instrument",

        "Allotment Date": "Allotment Date",

        "Maturity Date": "Maturity Date",

        "Face Value": "Face Value",

        "Issue Size": "Issue Size",

        "Coupon Rate": "Coupon Rate",

        "Credit Rating": "Credit Rating"
    }

    df = df.rename(
        columns=column_mapping
    )

    return df


# ============================================================
# 8. PROCESS DATES
# ============================================================

def process_dates(df):

    date_columns = [
        "Allotment Date",
        "Maturity Date"
    ]

    for column in date_columns:

        if column in df.columns:

            df[column] = pd.to_datetime(
                df[column],
                errors="coerce",
                dayfirst=True
            )

    return df


# ============================================================
# 9. APPLY USER FILTERS
# ============================================================

def apply_filters(df):

    # --------------------------------------------------------
    # Instrument type
    # --------------------------------------------------------

    if (
        INSTRUMENT_TYPE
        and "Type of Instrument" in df.columns
    ):

        df = df[
            df["Type of Instrument"]
            .astype(str)
            .str.contains(
                INSTRUMENT_TYPE,
                case=False,
                na=False
            )
        ]


    # --------------------------------------------------------
    # Allotment date
    # --------------------------------------------------------

    if "Allotment Date" in df.columns:

        if ALLOTMENT_DATE_FROM:

            start_date = pd.to_datetime(
                ALLOTMENT_DATE_FROM
            )

            df = df[
                df["Allotment Date"] >= start_date
            ]


        if ALLOTMENT_DATE_TO:

            end_date = pd.to_datetime(
                ALLOTMENT_DATE_TO
            )

            df = df[
                df["Allotment Date"] <= end_date
            ]


    # --------------------------------------------------------
    # Maturity date
    # --------------------------------------------------------

    if "Maturity Date" in df.columns:

        if MATURITY_DATE_FROM:

            start_date = pd.to_datetime(
                MATURITY_DATE_FROM
            )

            df = df[
                df["Maturity Date"] >= start_date
            ]


        if MATURITY_DATE_TO:

            end_date = pd.to_datetime(
                MATURITY_DATE_TO
            )

            df = df[
                df["Maturity Date"] <= end_date
            ]


    return df


# ============================================================
# 10. TRANSLATION
# ============================================================

def translate_text(df):

    """
    Translation stage.

    For financial datasets, numeric fields such as
    coupon, face value and issue size do not require
    translation.

    Text fields can be translated here if required.
    """

    if "Issuer" in df.columns:

        df["Issuer (Translated)"] = df["Issuer"]

    if "Type of Instrument" in df.columns:

        df["Type of Instrument (Translated)"] = (
            df["Type of Instrument"]
        )

    return df


# ============================================================
# 11. SAVE DATA
# ============================================================

def save_to_excel(df):

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Bond Data",
            index=False
        )

    print()
    print("Data saved successfully.")
    print(f"Output file: {OUTPUT_FILE}")


# ============================================================
# 12. MAIN EXTRACTION PIPELINE
# ============================================================

def run_extraction():

    print("=" * 60)
    print("INDIA BOND DATA EXTRACTION")
    print("=" * 60)

    # --------------------------------------------------------
    # Check working day
    # --------------------------------------------------------

    run_today, reason = is_market_day()

    if not run_today:

        print(f"Extraction skipped: {reason}")

        return


    print("Today is a working day.")
    print()


    # --------------------------------------------------------
    # Download website
    # --------------------------------------------------------

    html = get_page(BASE_URL)

    if html is None:

        return


    # --------------------------------------------------------
    # Extract tables
    # --------------------------------------------------------

    tables = extract_tables(html)

    if not tables:

        print("No data table found.")

        return


    # --------------------------------------------------------
    # Select largest table
    # --------------------------------------------------------

    df = max(
        tables,
        key=lambda table: table.shape[0]
    )


    print(
        f"Initial records: {len(df)}"
    )


    # --------------------------------------------------------
    # Clean
    # --------------------------------------------------------

    df = clean_data(df)


    # --------------------------------------------------------
    # Standardize
    # --------------------------------------------------------

    df = standardize_columns(df)


    # --------------------------------------------------------
    # Process dates
    # --------------------------------------------------------

    df = process_dates(df)


    # --------------------------------------------------------
    # Apply filters
    # --------------------------------------------------------

    df = apply_filters(df)


    # --------------------------------------------------------
    # Translation
    # --------------------------------------------------------

    df = translate_text(df)


    # --------------------------------------------------------
    # Final information
    # --------------------------------------------------------

    print(
        f"Final records: {len(df)}"
    )


    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    save_to_excel(df)


# ============================================================
# 13. RUN PROGRAM
# ============================================================

if __name__ == "__main__":

    run_extraction()
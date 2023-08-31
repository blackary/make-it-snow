from pathlib import Path
from typing import Dict

import pandas as pd
import streamlit as st

from ppt import populate_slide, BUCKETS

st.set_page_config(page_title="Fresh Snow", page_icon="❄️")

st.title(":snowflake: Fresh Snow!")

st.write("Upload a spreadsheet of recent hires, and select the week")


def get_bucket(user: pd.Series) -> str:
    if user.Country == "India":
        return "INDIA"

    elif user.Country == "United Arab Emirates":
        return "UAE"

    elif user.Country == "Israel":
        return "ISRAEL"

    elif user.Country in ["United States of America", "Canada"]:
        if user.Country == "United States of America":
            country = "USA"
        else:
            country = "CAN"

        state = user.Location.split("-")[1].replace(" Metro", "")
        # Find matching row in dataframe
        if country == "CAN":
            if state in ["British Columbia"]:
                return "WESTERN_NA"
            else:
                return "EASTERN_NA"

        elif country == "USA":
            if state in ["WA", "OR", "CA", "ID", "NV", "UT", "AZ", "MT", "WY"]:
                return "WESTERN_NA"
            elif state in [
                "CO",
                "NM",
                "TX",
                "OK",
                "KS",
                "SD",
                "ND",
                "MN",
                "IA",
                "MO",
                "AR",
                "LA",
            ]:
                return "CENTRAL_NA"
            elif state in [
                "WI",
                "IL",
                "MS",
                "AL",
                "GA",
                "KY",
                "IN",
                "MI",
                "OH",
                "TN",
                "WV",
                "FL",
                "SC",
                "NC",
                "VA",
                "DE",
                "MD",
                "PA",
                "NY",
                "NJ",
                "CT",
                "MA",
                "VT",
                "NH",
                "ME",
                "NB",
                "NS",
                "DC",
            ]:
                return "EASTERN_NA"

    elif user.Country == "Mexico":
        return "CENTRAL_NA"

    elif user.Country in ["Japan", "Korea, Republic of", "China"]:
        return "JAPAN_SOUTH_KOREA"

    elif user.Country in [
        "United Kingdom",
        "France",
        "Germany",
        "Italy",
        "Netherlands",
        "Poland",
        "Switzerland",
        "Spain",
        "Sweden",
        "Slovakia",
        "Denmark",
        "Ireland",
        "Finland",
        "Norway",
    ]:
        return "EUROPE"

    elif user.Country == "Indonesia":
        return "INDONESIA"

    elif user.Country in ["New Zealand"]:
        return "NEW_ZEALAND"

    elif user.Country in ["Singapore", "Malaysia"]:
        return "SINGAPORE_MALAYSIA"

    elif user.Country in ["Japan", "Korea, Republic of"]:
        return "JAPAN_SOUTH_KOREA"

    elif user.Country == "Philippines":
        return "PHILIPPINES"

    elif user.Country in ["Australia"]:
        city = user.Location.split("-")[1]
        if city in [
            "Sydney",
            "Melbourne",
            "Brisbane",
            "Adelaide",
            "Victoria",
            "New South Wales",
            "Queensland",
            "Canberra",
        ]:
            return "AUSTRALIA_EAST"
        elif city in ["Perth"]:
            return "AUSTRALIA_WEST"

    elif user.Country == "Brazil":
        return "BRAZIL"

    elif user.Country in ["Colombia", "Costa Rica"]:
        return "COSTA_RICA"

    raise LookupError(user)


def map_users_to_buckets(users: pd.DataFrame) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for bucket in BUCKETS:
        counts[bucket] = 0

    for row in users.itertuples():
        try:
            bucket = get_bucket(row)
            counts[bucket] += 1
        except LookupError as e:
            st.error(f"No bucket found for {e}")

    return counts


uploaded_file = st.file_uploader("Upload an excel spreadsheet", type="xlsx")

use_sample = st.checkbox("Use sample data", help="Use example set of fake data")

df = None

if uploaded_file is not None:
    # read csv
    df = pd.read_excel(
        uploaded_file, usecols=["Location", "Country", "Hire Date"], header=3
    )
elif use_sample:
    df = pd.read_csv("example.csv")
    df["Hire Date"] = pd.to_datetime(df["Hire Date"])


if df is not None:
    st.expander("Show all data").write(df)

    # Add column for start of week
    df["hire_week"] = df["Hire Date"].dt.to_period("W")

    hire_week = st.select_slider("Select hire week", df["hire_week"].unique())

    users = df[df["hire_week"] == hire_week]

    buckets = map_users_to_buckets(users)

    st.json(dict(buckets))

    populate_slide(buckets, hire_week.start_time)

    # Add a streamlit download button to download output.pptx
    st.download_button(
        "Download output.pptx",
        data=Path("output.pptx").read_bytes(),
        file_name=f"fresh-snow-{hire_week.start_time:%Y-%m-%d}.pptx",
    )

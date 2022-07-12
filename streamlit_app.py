from cgitb import lookup
from collections import defaultdict
from typing import DefaultDict, Dict

import streamlit as st

st.title(":snowflake: Fresh Snow!")

st.write("Upload a spreadsheet of recent hires, and select the week")

import pandas as pd
import streamlit as st

state_timezeones = pd.read_csv("state_tz.csv")

BUCKETS = [
    "WESTERN_NA",
    "CENTRAL_NA",
    "EASTERN_NA",
    "EUROPE",
    "INDIA",
    "SOUTHEAST_ASIA",
    "EAST_ASIA",
    "AUSTRALIA",
]


def get_bucket(user: pd.Series) -> str:
    if user.Country in ["India", "United Arab Emirates"]:
        return "INDIA"

    elif user.Country in ["United States of America", "Canada"]:
        if user.Country == "United States of America":
            country = "USA"
        else:
            country = "CAN"

        state = user.Location.split("-")[1]
        # Find matching row in dataframe
        try:
            timezone = state_timezeones.query(
                f"STATE == '{state}' & COUNTRY == '{country}'"
            ).TZ.values[0]
        except LookupError:
            return "EASTERN_NA"

        if timezone == "PST":
            return "WESTERN_NA"

        if state in ["UT", "MT"]:
            return "WESTERN_NA"

        if state in ["TX", "CO"]:
            return "CENTRAL_NA"

        if timezone in ["EST", "CST"]:
            return "EASTERN_NA"

        if timezone in "MST":
            return "CENTRAL_NA"

    elif user.Country in ["Mexico"]:
        return "CENTRAL_NA"

    elif user.Country in ["Japan", "Korea, Republic of", "China"]:
        return "EAST_ASIA"

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
    ]:
        return "EUROPE"

    elif user.Country in ["Indonesia", "Singapore", "Philippines"]:
        return "SOUTHEAST_ASIA"

    elif user.Country in ["Australia", "New Zealand"]:
        return "AUSTRALIA"

    raise LookupError(user)


def map_users_to_buckets(users: pd.DataFrame) -> Dict[str, int]:
    counts: DefaultDict[str, int] = defaultdict(int)
    for row in users.itertuples():
        try:
            bucket = get_bucket(row)
            counts[bucket] += 1
        except LookupError as e:
            st.error(f"No bucket found for {e}")

    return counts


uploaded_file = st.file_uploader("Upload an excel spreadsheet", type="xlsx")
if uploaded_file is not None:
    # read csv
    df = pd.read_excel(
        uploaded_file, usecols=["Location", "Country", "Hire Date"], header=3
    )

    st.expander("Show all data").write(df)

    # Add column for start of week
    df["hire_week"] = df["Hire Date"].dt.to_period("W")

    hire_week = st.select_slider("Select hire week", df["hire_week"].unique())

    users = df[df["hire_week"] == hire_week]

    st.json(dict(map_users_to_buckets(users)))

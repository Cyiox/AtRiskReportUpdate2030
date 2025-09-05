import pandas as pd


# Load your data
df = pd.read_excel("2030 Forecast E-U-R Draft -Master.xlsx", sheet_name="Sheet1").fillna("")

# 🔧 Convert numeric columns to numbers
df["TotalUnits"] = pd.to_numeric(df["TotalUnits"], errors="coerce")
df["S8_PBAUnits"] = pd.to_numeric(df["S8_PBAUnits"], errors="coerce")

# 🔧 Convert the date column to datetime
df["S8_ExpDate"] = pd.to_datetime(df["S8_ExpDate"], errors="coerce")

# ✅ Define the condition
condition = (
    (df["S8_PBAUnits"] == (df["TotalUnits"] - 1)) &
    (df["S8_ExpDate"] > pd.Timestamp("2030-12-31"))
)

# ✅ Update multiple columns
df.loc[condition, ["Jameer notes", "1st Cut Status"]] = [
    "100% S.8 -1 unit, Expiring after forecast date",
    "PRESERVED"
]

df.to_excel("2030 Forecast E-U-R Draft -Master-Updated.xlsx", index=False)

'''
import pandas as pd

# Load data
df = pd.read_excel("2030 Forecast E-U-R Draft -Master.xlsx", sheet_name="Main Forecast").fillna("")
incomingInfo = pd.read_excel("MergedForecast.xlsx").fillna("")

# Only keep columns in incomingInfo that are NOT already in df (except the key column)
key = "PropertyID"
new_columns = [col for col in incomingInfo.columns if col != key and col not in df.columns]

# Reduce incomingInfo to just the key + new columns
incomingInfo_reduced = incomingInfo[[key] + new_columns]

# Merge with only the new columns added
df = df.merge(incomingInfo_reduced, on=key, how="left")
print("Frames mereged")
df.to_excel("test.xlsx", index=False)
'''
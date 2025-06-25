from openpyxl import load_workbook
import pandas as pd

source = "jp_eqhist_data_filtered.xlsx"
wb = load_workbook(filename=source)
ws = wb["Data"]

df = pd.read_excel(source, sheet_name="Data")

# Ensure 'Time' is datetime
df["Time"] = pd.to_datetime(df["Time"])

# Extract year
df["Year"] = df["Time"].dt.year
# Extract month
df["Month"] = df["Time"].dt.month
# Extract day
df["Day"] = df["Time"].dt.day

# Group by year and seismic intensity, count occurrences
result_year = df.groupby(["Year", "Seismic Intensity"]).size().unstack(fill_value=0)

# Write to a new sheet in the same Excel file
with pd.ExcelWriter(
    source, engine="openpyxl", mode="a", if_sheet_exists="replace"
) as writer:
    result_year.to_excel(writer, sheet_name="Count by year")

# Group by year and seismic intensity, count occurrences
result_month = (
    df.groupby(["Year", "Month", "Seismic Intensity"]).size().unstack(fill_value=0)
)

# Write to a new sheet in the same Excel file
with pd.ExcelWriter(
    source, engine="openpyxl", mode="a", if_sheet_exists="replace"
) as writer:
    result_month.to_excel(writer, sheet_name="Count by month")

# Group by year and seismic intensity, count occurrences
result_day = (
    df.groupby(["Year", "Month", "Day", "Seismic Intensity"])
    .size()
    .unstack(fill_value=0)
)

# Write to a new sheet in the same Excel file
with pd.ExcelWriter(
    source, engine="openpyxl", mode="a", if_sheet_exists="replace"
) as writer:
    result_day.to_excel(writer, sheet_name="Count by day")

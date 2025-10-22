# create_wind_template.py
import pandas as pd

# Define template data
data = {
    "Wind Speed Bin (m/s)": ["0–2", "2–4", "4–6", "6–8", "8–10", "10–12", "12–14"],
    "Jan": [120, 200, 250, 180, 100, 60, 20],
    "Feb": [110, 180, 230, 160, 90, 50, 20],
    "Mar": [130, 210, 260, 190, 110, 70, 30],
    "Apr": [100, 190, 240, 170, 100, 60, 25],
    "May": [90, 170, 220, 150, 90, 50, 20],
    "Jun": [80, 150, 200, 140, 80, 40, 15],
    "Jul": [70, 140, 190, 130, 70, 30, 10],
    "Aug": [80, 150, 200, 140, 80, 40, 15],
    "Sep": [100, 160, 210, 150, 90, 50, 20],
    "Oct": [120, 180, 230, 160, 100, 60, 25],
    "Nov": [130, 190, 240, 170, 110, 70, 30],
    "Dec": [140, 200, 250, 180, 120, 80, 35],
}

# Create DataFrame
df = pd.DataFrame(data)

# Define file name and path (same folder as app)
template_filename = "wind_frequency_template.xlsx"

# Save Excel file
df.to_excel(template_filename, index=False)

print(f"✅ Excel template '{template_filename}' successfully created in the current folder.")

import os
import json
import gspread
import pandas as pd

# 1. Pull the secret string from GitHub Actions environment
creds_string = os.environ.get('GOOGLE_CREDENTIALS')

# 2. Convert that string back into a JSON dictionary format
creds_dict = json.loads(creds_string)

# 3. Log in using the dictionary
gc = gspread.service_account_from_dict(creds_dict)

# Now extract your data exactly like before:
sheet = gc.open("Cypress Sales Transaction Log")
raw_tab = sheet.worksheet("Original Responses")
df_responses = pd.DataFrame(raw_tab.get_all_records())

# Pull Weather Data (using get_all_values to avoid blank column crash)
weather_tab = sheet.worksheet("Weather Data")
weather_data = weather_tab.get_all_values()
df_weather = pd.DataFrame(weather_data[1:], columns=weather_data[0])
df_weather = df_weather.loc[:, df_weather.columns != '']

# Pull Gross Sales Data (using get_all_values to avoid blank column crash)
gross_tab = sheet.worksheet("Gross Sales Data")
gross_data = gross_tab.get_all_values()
df_gross = pd.DataFrame(gross_data[1:], columns=gross_data[0])
df_gross = df_gross.loc[:, df_gross.columns != '']

# Normalizing Date Columns to create a derived key to join
df_responses["Date"] = pd.to_datetime(df_responses["Timestamp"]).dt.normalize()
df_weather["Date"] = pd.to_datetime(df_weather["Date"])
df_gross["Date"] = pd.to_datetime(df_gross["Date"])

# Organizing Sterling Silver items
df_responses["Item Category"] = df_responses["Item Category"].replace({
    'Earring, Sterling Silver' : 'Sterling Silver Earring',
    'Ring, Sterling Silver' : 'Sterling Silver Ring',
    'Necklace, Sterling Silver' : 'Sterling Silver Necklace',
    'Necklace, Earring, Sterling Silver' : 'Sterling Silver Earring, Necklace'})

# Merging gross sales data with weather data
df_gross_weather = df_gross.merge(df_weather, on=["Date"], how = "left")

# Adding transaction IDs to each transaction in df_responses table
df_responses.insert(0, 'Transaction ID', "TXN-" + (df_responses.index + 1).astype(str))

# Copying responses table and splitting with delimiter comma (",")
df_responses_copy = df_responses.copy()
df_responses_copy['Item Category'] = df_responses_copy['Item Category'].str.split(pat=', ')

# Exploding "Item Category" to create separate item rows
df_responses_copy = df_responses_copy.explode('Item Category')

# 1. Connect to your newly created destination tabs
ws_items = sheet.worksheet("Clean Response Data")
ws_daily = sheet.worksheet("Clean Sales_Weather Data")

# 2. Clear the tabs to ensure a clean slate
ws_items.clear()
ws_daily.clear()

# 3. Create safe copies for the final export
df_responses_clean = df_responses_copy.copy()
df_gross_weather_clean = df_gross_weather.copy()

# 4. Convert Pandas Timestamps to plain strings so JSON can read them
df_responses_clean['Timestamp'] = df_responses_clean['Timestamp'].astype(str)
df_responses_clean['Date'] = df_responses_clean['Date'].astype(str)
df_gross_weather_clean['Date'] = df_gross_weather_clean['Date'].astype(str)

# 5. Replace any NaNs or NaTs (blanks) with empty strings
df_responses_clean = df_responses_clean.fillna('').replace('NaT', '')
df_gross_weather_clean = df_gross_weather_clean.fillna('').replace('NaT', '')

# 6. Export the DataFrames using USER_ENTERED to auto-parse dates
ws_items.update(
    values=[df_responses_clean.columns.values.tolist()] + df_responses_clean.values.tolist(),
    range_name='A1',
    value_input_option='USER_ENTERED'
)

ws_daily.update(
    values=[df_gross_weather_clean.columns.values.tolist()] + df_gross_weather_clean.values.tolist(),
    range_name='A1',
    value_input_option='USER_ENTERED'
)

# 7. Format the headers (Freeze, Bold, Filter)
ws_items.freeze(rows=1)
ws_daily.freeze(rows=1)
ws_items.format('1:1', {'textFormat': {'bold': True}})
ws_daily.format('1:1', {'textFormat': {'bold': True}})
ws_items.set_basic_filter()
ws_daily.set_basic_filter()

print("ETL Pipeline Complete! Data successfully pushed to Google Sheets.")
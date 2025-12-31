import pandas as pd

# Just add this "encoding='latin1'" part to your code
df = pd.read_csv(r"F:\sample projects\Sales Performance Analysis\superstore.csv", encoding='latin1')
print("\nData shape (rows, columns):", df.shape)
print("\nColumn names:",df.columns.tolist())
print(df.head())

#Convert date
df['Order Date'] = pd.to_datetime(df['Order Date'])
df['Ship Date'] = pd.to_datetime(df['Ship Date'])

df['Order Year'] = df['Order Date'].dt.year
df['Order Month'] = df['Order Date'].dt.month
df['Order Month Name'] = df['Order Date'].dt.month_name()

#Profit margin
df['Profit Margin'] = (df['Profit'] / df['Sales']) * 100
df['Profit Margin'] = df['Profit Margin'].fillna(0)

#export to excel
# Save to your project folder
df.to_excel(r"F:\sample projects\Sales Performance Analysis\superstore_cleaned.xlsx", index=False)

print("Cleaned dataset saved successfully.")

#append new data
# Simulate new incoming sales data
new_data = {
    'Row ID': [99999],
    'Order ID': ['CA-2025-NEW01'],
    'Order Date': [pd.to_datetime('2025-01-15')],
    'Ship Date': [pd.to_datetime('2025-01-18')],
    'Ship Mode': ['Second Class'],
    'Customer ID': ['CG-12520'],
    'Customer Name': ['Test Customer'],
    'Segment': ['Consumer'],
    'Country': ['United States'],
    'City': ['Los Angeles'],
    'State': ['California'],
    'Postal Code': [90001],
    'Region': ['West'],
    'Product ID': ['OFF-NEW-100'],
    'Category': ['Office Supplies'],
    'Sub-Category': ['Binders'],
    'Product Name': ['New Binder Pack'],
    'Sales': [250],
    'Quantity': [5],
    'Discount': [0.1],
    'Profit': [45]
}

new_df = pd.DataFrame(new_data)

# Recalculate features for new data
new_df['Order Year'] = new_df['Order Date'].dt.year
new_df['Order Month'] = new_df['Order Date'].dt.month
new_df['Order Month Name'] = new_df['Order Date'].dt.month_name()
new_df['Profit Margin'] = (new_df['Profit'] / new_df['Sales']) * 100

# Append to existing data
df = pd.concat([df, new_df], ignore_index=True)

# Save updated dataset
df.to_excel(r"F:\sample projects\Sales Performance Analysis\superstore_cleaned.xlsx", index=False)
print("New data appended and file updated.")

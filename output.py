import pandas as pd
import matplotlib.pyplot as plt

# Load the CSV file
file_path = "C:\\Users\\alone\\Amazon Sale Report.csv"  # Adjust the path as necessary
sales_data = pd.read_csv(file_path)

# Convert 'Date' to datetime format
sales_data['Date'] = pd.to_datetime(sales_data['Date'])

# Replace 'Amount' and 'Qty' with the correct column names from your dataset
sales_column = 'Amount'  # Column name for sales
quantity_column = 'Qty'  # Column name for quantity sold

# Check if the necessary columns exist
if sales_column not in sales_data.columns or quantity_column not in sales_data.columns:
    raise ValueError("The specified sales or quantity columns do not exist in the data.")

# Set 'Date' as the index for resampling
sales_data.set_index('Date', inplace=True)

# 1. Sales Overview

# Summary Statistics
total_sales = sales_data[sales_column].sum()
average_sales_per_transaction = sales_data[sales_column].mean()
total_quantity_sold = sales_data[quantity_column].sum()

# Trend Analysis: Plot sales over time (daily)
sales_trend_daily = sales_data[sales_column].resample('D').sum()

# Monthly and Quarterly Trends
sales_trend_monthly = sales_data[sales_column].resample('ME').sum()
sales_trend_quarterly = sales_data[sales_column].resample('QE').sum()

# Sales Growth
sales_trend_monthly_pct_change = sales_trend_monthly.pct_change().fillna(0)

# Top Performers: Best-performing products and categories based on sales volume and revenue
top_products = sales_data.groupby('Category')[sales_column].sum().sort_values(ascending=False).head(10)
top_categories = sales_data.groupby('Category')[sales_column].sum().sort_values(ascending=False).head(10)

# Save results to Excel file
output_file_path = "C:\\Users\\alone\\Amazon_Sales_Report_Analysis.xlsx"  # Adjust the path as necessary
with pd.ExcelWriter(output_file_path) as writer:
    sales_trend_daily.to_frame(name='Daily Sales').to_excel(writer, sheet_name='Daily Sales')
    sales_trend_monthly.to_frame(name='Monthly Sales').to_excel(writer, sheet_name='Monthly Sales')
    sales_trend_quarterly.to_frame(name='Quarterly Sales').to_excel(writer, sheet_name='Quarterly Sales')
    sales_trend_monthly_pct_change.to_frame(name='Month-over-Month Sales Growth').to_excel(writer, sheet_name='Sales Growth')
    top_products.to_frame(name='Top 10 Products').to_excel(writer, sheet_name='Top 10 Products')
    top_categories.to_frame(name='Top 10 Categories').to_excel(writer, sheet_name='Top 10 Categories')

# Display visualizations
plt.figure(figsize=(14, 10))

# Daily Sales Trend
plt.subplot(3, 1, 1)
sales_trend_daily.plot()
plt.title('Daily Sales Trend')
plt.xlabel('Date')
plt.ylabel(f'Sales ({sales_column})')
plt.grid(True)

# Monthly Sales Trend
plt.subplot(3, 1, 2)
sales_trend_monthly.plot(kind='bar')
plt.title('Monthly Sales Trend')
plt.xlabel('Month')
plt.ylabel(f'Sales ({sales_column})')
plt.grid(True)

# Quarterly Sales Trend
plt.subplot(3, 1, 3)
sales_trend_quarterly.plot(kind='bar', color='orange')
plt.title('Quarterly Sales Trend')
plt.xlabel('Quarter')
plt.ylabel(f'Sales ({sales_column})')
plt.grid(True)

plt.tight_layout()
plt.show()

# Print the summary statistics and growth rate
print(f"Total Sales: ${total_sales}")
print(f"Average Sales per Transaction: ${average_sales_per_transaction}")
print(f"Total Quantity Sold: {total_quantity_sold}")
print(f"Month-over-Month Sales Growth:\n{sales_trend_monthly_pct_change}")
print(f"Top 10 Categories by Sales:\n{top_categories}")

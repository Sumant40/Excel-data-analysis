#!/usr/bin/env python3
"""
Retail Sales Data Generator
Creates sample dataset for Excel Analytics Project

Usage: python generate_sales_data.py
Output: retail_sales_data.csv (1000 transactions)
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

# Set random seed for reproducibility
np.random.seed(42)
random.seed(42)

print("Generating Retail Sales Dataset...")
print("=" * 60)

# Configuration
NUM_RECORDS = 1000
START_DATE = datetime(2023, 1, 1)
END_DATE = datetime(2023, 12, 31)

# Master Data
REGIONS = ["North", "South", "East", "West", "Central"]

CITIES = {
    "North": ["Delhi", "Jaipur", "Lucknow", "Chandigarh"],
    "South": ["Bangalore", "Chennai", "Hyderabad", "Kochi"],
    "East": ["Kolkata", "Bhubaneswar", "Patna", "Ranchi"],
    "West": ["Mumbai", "Pune", "Ahmedabad", "Surat"],
    "Central": ["Bhopal", "Indore", "Nagpur", "Raipur"]
}

PRODUCT_CATEGORIES = {
    "Laptops": {
        "products": ["Dell Inspiron", "HP Pavilion", "Lenovo ThinkPad", "ASUS VivoBook", "Acer Aspire"],
        "price_range": (35000, 80000),
        "avg_quantity": 1.5
    },
    "Smartphones": {
        "products": ["Samsung Galaxy S23", "iPhone 14", "OnePlus Nord", "Xiaomi 13", "Realme GT"],
        "price_range": (15000, 60000),
        "avg_quantity": 1.3
    },
    "Tablets": {
        "products": ["iPad Air", "Samsung Tab S8", "Lenovo Tab P11", "Amazon Fire HD"],
        "price_range": (20000, 50000),
        "avg_quantity": 1.2
    },
    "Accessories": {
        "products": ["Wireless Mouse", "Mechanical Keyboard", "USB-C Cable", "Laptop Bag", "Webcam"],
        "price_range": (500, 3000),
        "avg_quantity": 2.5
    },
    "Wearables": {
        "products": ["Apple Watch Series 8", "Fitbit Versa 4", "Samsung Galaxy Watch", "Mi Band 7"],
        "price_range": (10000, 35000),
        "avg_quantity": 1.1
    }
}

SALES_REPS = ["Rahul Sharma", "Priya Patel", "Amit Kumar", "Sneha Reddy", "Vikram Singh", 
              "Anjali Gupta", "Rohan Mehta", "Kavya Iyer"]

CUSTOMER_NAMES = [
    "Rajesh Kumar", "Priya Singh", "Amit Patel", "Sneha Sharma", "Vikram Reddy",
    "Anjali Gupta", "Rohan Mehta", "Kavya Iyer", "Arjun Nair", "Divya Shah",
    "Karan Verma", "Neha Agarwal", "Sanjay Joshi", "Pooja Desai", "Rahul Kapoor",
    "Simran Kaur", "Aditya Rao", "Riya Malhotra", "Varun Chopra", "Ishita Bansal",
    "Mohit Saxena", "Tanvi Kulkarni", "Nikhil Jain", "Shreya Pandey", "Harsh Sinha",
    "Ananya Bose", "Kunal Dixit", "Nisha Ghosh", "Siddharth Tiwari", "Meera Chawla",
    "Gaurav Bhatt", "Ritika Menon", "Akash Pillai", "Sakshi Rathore", "Yash Dhawan",
    "Aarti Shetty", "Manish Arora", "Pallavi Krishnan", "Sameer Dutta", "Swati Nambiar"
]

# Discount tiers (more realistic distribution)
DISCOUNT_OPTIONS = [0, 0, 0, 5, 5, 10, 10, 15, 20, 25]  # Weighted towards lower discounts

print(f"Generating {NUM_RECORDS} transaction records...")
print(f"Date Range: {START_DATE.date()} to {END_DATE.date()}")
print()

# Generate data
data = []

for i in range(NUM_RECORDS):
    # Generate random date
    days_diff = (END_DATE - START_DATE).days
    random_days = random.randint(0, days_diff)
    order_date = START_DATE + timedelta(days=random_days)
    
    # Select random region and city
    region = random.choice(REGIONS)
    city = random.choice(CITIES[region])
    
    # Select product category and product
    category = random.choice(list(PRODUCT_CATEGORIES.keys()))
    category_data = PRODUCT_CATEGORIES[category]
    product = random.choice(category_data["products"])
    
    # Generate quantity (slightly random around average)
    avg_qty = category_data["avg_quantity"]
    quantity = max(1, int(np.random.normal(avg_qty, 0.5)))
    
    # Generate price within range (normal distribution)
    min_price, max_price = category_data["price_range"]
    mid_price = (min_price + max_price) / 2
    std_dev = (max_price - min_price) / 4
    unit_price = int(np.random.normal(mid_price, std_dev))
    unit_price = max(min_price, min(max_price, unit_price))  # Clamp to range
    
    # Round to nearest 100 for realistic pricing
    unit_price = round(unit_price / 100) * 100
    
    # Discount (weighted towards lower values)
    discount_percent = random.choice(DISCOUNT_OPTIONS)
    
    # Customer
    customer_name = random.choice(CUSTOMER_NAMES)
    customer_id = f"CUST-{hash(customer_name) % 10000:04d}"
    
    # Sales rep
    sales_rep = random.choice(SALES_REPS)
    
    # Order ID
    order_id = f"ORD-{i+1:04d}"
    
    # Calculate amounts
    total_sales = quantity * unit_price
    discount_amount = total_sales * (discount_percent / 100)
    net_sales = total_sales - discount_amount
    
    # Append record
    data.append({
        'Order_ID': order_id,
        'Order_Date': order_date.strftime('%d-%m-%Y'),
        'Customer_ID': customer_id,
        'Customer_Name': customer_name,
        'Region': region,
        'City': city,
        'Product_Category': category,
        'Product_Name': product,
        'Quantity': quantity,
        'Unit_Price': unit_price,
        'Discount_Percent': discount_percent,
        'Sales_Rep': sales_rep
    })

# Create DataFrame
df = pd.DataFrame(data)

# Sort by date
df = df.sort_values('Order_Date').reset_index(drop=True)

# Regenerate Order_IDs after sorting
df['Order_ID'] = [f"ORD-{i+1:04d}" for i in range(len(df))]

print("Data generation complete!")
print()
print("Dataset Statistics:")
print("=" * 60)
print(f"Total Records: {len(df)}")
print(f"Date Range: {df['Order_Date'].min()} to {df['Order_Date'].max()}")
print(f"Unique Customers: {df['Customer_ID'].nunique()}")
print(f"Unique Products: {df['Product_Name'].nunique()}")
print()

print("Regional Distribution:")
print(df['Region'].value_counts().to_string())
print()

print("Product Category Distribution:")
print(df['Product_Category'].value_counts().to_string())
print()

# Calculate summary metrics
df['Total_Sales'] = df['Quantity'] * df['Unit_Price']
df['Discount_Amount'] = df['Total_Sales'] * (df['Discount_Percent'] / 100)
df['Net_Sales'] = df['Total_Sales'] - df['Discount_Amount']

total_revenue = df['Net_Sales'].sum()
total_orders = len(df)
avg_order_value = total_revenue / total_orders

print("Revenue Metrics:")
print("=" * 60)
print(f"Total Revenue: ₹{total_revenue:,.0f}")
print(f"Average Order Value: ₹{avg_order_value:,.0f}")
print(f"Total Discount Given: ₹{df['Discount_Amount'].sum():,.0f}")
print(f"Average Discount: {df['Discount_Percent'].mean():.1f}%")
print()

# Save to CSV (without calculated columns - user will add them in Excel)
output_cols = ['Order_ID', 'Order_Date', 'Customer_ID', 'Customer_Name', 'Region', 
               'City', 'Product_Category', 'Product_Name', 'Quantity', 'Unit_Price', 
               'Discount_Percent', 'Sales_Rep']

output_df = df[output_cols]

# Save to CSV
output_file = 'retail_sales_data.csv'
output_df.to_csv(output_file, index=False)

print(f"Dataset saved to: {output_file}")
print()

# Preview first 10 records
print("Sample Data (First 10 Records):")
print("=" * 60)
print(output_df.head(10).to_string(index=False))
print()

print("Dataset generation complete!")
print()
print("Next Steps:")
print("  1. Open the CSV file in Excel")
print("  2. Follow the README instructions to build your dashboard")
print("  3. Add calculated columns (Total_Sales, Discount_Amount, Net_Sales, etc.)")
print("  4. Create Pivot Tables and Charts")
print("  5. Build the interactive dashboard")
print()
print("Good luck with your Excel Analytics Project!")
print("=" * 60)

# Generate additional insights for documentation
print()
print("Quick Insights for Your Analysis:")
print("=" * 60)

# Top selling category
top_category = df.groupby('Product_Category')['Net_Sales'].sum().sort_values(ascending=False)
print(f"Top Category: {top_category.index[0]} (₹{top_category.values[0]:,.0f})")

# Best region
top_region = df.groupby('Region')['Net_Sales'].sum().sort_values(ascending=False)
print(f"Top Region: {top_region.index[0]} (₹{top_region.values[0]:,.0f})")

# Best month
df['Month'] = pd.to_datetime(df['Order_Date'], format='%d-%m-%Y').dt.to_period('M')
top_month = df.groupby('Month')['Net_Sales'].sum().sort_values(ascending=False)
print(f"Best Month: {top_month.index[0]} (₹{top_month.values[0]:,.0f})")

# Top sales rep
top_rep = df.groupby('Sales_Rep')['Net_Sales'].sum().sort_values(ascending=False)
print(f"Top Sales Rep: {top_rep.index[0]} (₹{top_rep.values[0]:,.0f})")

print()
print("Use these insights to validate your Excel formulas!")
print("=" * 60)

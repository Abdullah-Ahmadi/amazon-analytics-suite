#!/usr/bin/env python3
"""
Generate sample Amazon data for testing the analytics tool.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

def generate_sample_data():
    """Generate sample CSV files for testing."""
    print("Generating sample Amazon data...")

    # Generate sales data
    dates = pd.date_range(start='2023-01-01', end='2023-12-31', freq='D')
    products = [
        ('Wireless Earbuds Pro', 'B08XYZ1234', 'Electronics', 79.99),
        ('Yoga Mat Premium', 'B09ABC5678', 'Sports', 34.99),
        ('Organic Coffee 2lb', 'B07DEF9012', 'Grocery', 24.99),
        ('LED Desk Lamp', 'B10GHI3456', 'Home', 45.50),
        ('Phone Case iPhone', 'B11JKL7890', 'Accessories', 19.99),
    ]

    sales_data = []
    for i in range(1000):
        product_name, sku, category, price = random.choice(products)
        sales_data.append({
            'order_id': f'171-{i:07d}',
            'order_date': random.choice(dates).strftime('%Y-%m-%d'),
            'sku': sku,
            'product_name': product_name,
            'category': category,
            'quantity': random.randint(1, 3),
            'unit_price': price,
            'total_amount': price * random.randint(1, 3),
            'marketplace': random.choice(['Amazon.com', 'Amazon.ca']),
            'fulfillment': random.choice(['FBA', 'FBM'])
        })

    sales_df = pd.DataFrame(sales_data)
    sales_df.to_csv('sample_sales_data.csv', index=False)
    print(f"✅ Generated sample_sales_data.csv ({len(sales_df)} rows)")

    # Generate inventory data
    inventory_data = []
    for product_name, sku, category, price in products:
        inventory_data.append({
            'asin': sku,
            'product_name': product_name,
            'current_stock': random.randint(0, 500),
            'inbound_to_amazon': random.randint(0, 100),
            'days_of_supply': random.randint(1, 60),
            'stranded_status': 'Yes' if random.random() < 0.1 else 'No'
        })

    inventory_df = pd.DataFrame(inventory_data)
    inventory_df.to_csv('sample_inventory_data.csv', index=False)
    print(f"✅ Generated sample_inventory_data.csv ({len(inventory_df)} rows)")

    # Generate advertising data
    advertising_data = []
    for i in range(100):
        advertising_data.append({
            'date': random.choice(dates).strftime('%Y-%m-%d'),
            'campaign_name': random.choice(['Brand_Campaign', 'Auto_Targeting']),
            'spend': random.uniform(10, 200),
            'clicks': random.randint(10, 500),
            'impressions': random.randint(1000, 10000),
            'sales_attributed': random.uniform(20, 500)
        })

    advertising_df = pd.DataFrame(advertising_data)
    advertising_df.to_csv('sample_advertising_data.csv', index=False)
    print(f"✅ Generated sample_advertising_data.csv ({len(advertising_df)} rows)")

    # Generate reviews data
    reviews_data = []
    for i in range(200):
        product_name, sku, category, price = random.choice(products)
        reviews_data.append({
            'review_date': random.choice(dates).strftime('%Y-%m-%d'),
            'rating': random.choices([1, 2, 3, 4, 5], weights=[0.05, 0.1, 0.15, 0.3, 0.4])[0],
            'review_title': f'Review {i+1}',
            'product_name': product_name,
            'verified_purchase': random.choice(['Yes', 'No'])
        })

    reviews_df = pd.DataFrame(reviews_data)
    reviews_df.to_csv('sample_reviews_data.csv', index=False)
    print(f"✅ Generated sample_reviews_data.csv ({len(reviews_df)} rows)")

    print("\n" + "="*60)
    print("Sample data generation complete!")
    print("Place these files in the same directory as main.py and run:")
    print("  python main.py")
    print("="*60)

if __name__ == "__main__":
    generate_sample_data()

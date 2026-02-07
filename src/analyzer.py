"""
Data analysis and business logic module - FIXED VERSION
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import logging

class DataAnalyzer:
    """Performs business analysis on the loaded data."""
    
    def __init__(self, config, logger=None):
        self.config = config
        self.logger = logger or logging.getLogger(__name__)
        self.results = {}
        self.alerts = []
        
    def analyze_all(self, data_dict):
        """Run all analyses."""
        self.logger.info("Starting comprehensive analysis...")
        
        try:
            # Sales analysis
            if 'sales' in data_dict and data_dict['sales'] is not None:
                self._analyze_sales(data_dict['sales'])
            
            # Inventory analysis
            if 'inventory' in data_dict and data_dict['inventory'] is not None:
                self._analyze_inventory(data_dict['inventory'])
            
            # Advertising analysis
            if 'advertising' in data_dict and data_dict['advertising'] is not None:
                self._analyze_advertising(data_dict['advertising'])
            
            # Reviews analysis
            if 'reviews' in data_dict and data_dict['reviews'] is not None:
                self._analyze_reviews(data_dict['reviews'])
            
            self.logger.info(f"Generated {len(self.alerts)} alerts")
            
        except Exception as e:
            self.logger.error(f"Error during analysis: {e}")
            raise
    
    def _analyze_sales(self, sales_df):
        """Analyze sales data."""
        self.logger.info("Analyzing sales data...")
        
        # Find revenue/total amount column
        revenue_col = None
        for col in sales_df.columns:
            col_lower = col.lower()
            if any(key in col_lower for key in ['total', 'revenue', 'amount', 'price']):
                if 'unit' not in col_lower:  # Skip unit_price
                    revenue_col = col
                    break
        
        # Find quantity column
        quantity_col = None
        for col in sales_df.columns:
            col_lower = col.lower()
            if any(key in col_lower for key in ['quantity', 'qty']):
                quantity_col = col
                break
        
        # Basic sales metrics
        sales_metrics = {
            'total_transactions': len(sales_df),
            'total_quantity': sales_df[quantity_col].sum() if quantity_col and quantity_col in sales_df.columns else 0,
            'total_revenue': sales_df[revenue_col].sum() if revenue_col and revenue_col in sales_df.columns else 0,
            'avg_order_value': sales_df[revenue_col].mean() if revenue_col and revenue_col in sales_df.columns else 0
        }
        
        # Try to identify date column
        date_col = None
        for col in sales_df.columns:
            if 'date' in col.lower():
                date_col = col
                break
        
        if date_col and date_col in sales_df.columns:
            try:
                sales_df['_date'] = pd.to_datetime(sales_df[date_col], errors='coerce')
                sales_df['_month_str'] = sales_df['_date'].dt.strftime('%Y-%m')  # Convert to string for Excel
                
                if revenue_col and revenue_col in sales_df.columns:
                    # Monthly trends
                    monthly_sales = sales_df.groupby('_month_str').agg({
                        revenue_col: 'sum'
                    }).reset_index()
                    
                    monthly_sales.columns = ['month', 'revenue']
                    sales_metrics['monthly_trends'] = monthly_sales
                    
                    # Recent performance (last 30 days)
                    recent_cutoff = datetime.now() - timedelta(days=30)
                    recent_sales = sales_df[sales_df['_date'] >= recent_cutoff]
                    sales_metrics['recent_transactions'] = len(recent_sales)
                    sales_metrics['recent_revenue'] = recent_sales[revenue_col].sum() if revenue_col in recent_sales.columns else 0
                    
                    # Growth rate (if enough data)
                    if len(monthly_sales) > 1:
                        last_month = monthly_sales.iloc[-1]['revenue']
                        prev_month = monthly_sales.iloc[-2]['revenue'] if len(monthly_sales) > 1 else 0
                        if prev_month > 0:
                            growth = ((last_month - prev_month) / prev_month) * 100
                            sales_metrics['monthly_growth'] = growth
            
            except Exception as e:
                self.logger.warning(f"Could not analyze date-based sales metrics: {e}")
        
        # Product performance
        product_col = None
        for col in sales_df.columns:
            if 'product' in col.lower() and 'name' in col.lower():
                product_col = col
                break
        
        if product_col and product_col in sales_df.columns and revenue_col and revenue_col in sales_df.columns:
            try:
                product_performance = sales_df.groupby(product_col).agg({
                    revenue_col: 'sum'
                }).round(2)
                
                product_performance.columns = ['total_revenue']
                product_performance = product_performance.reset_index()
                
                if quantity_col and quantity_col in sales_df.columns:
                    product_qty = sales_df.groupby(product_col).agg({
                        quantity_col: 'sum'
                    }).round(2)
                    product_performance['total_quantity'] = product_qty[quantity_col].values
                
                sales_metrics['top_products'] = product_performance.sort_values(
                    'total_revenue', ascending=False
                ).head(10)
            except:
                self.logger.warning("Could not calculate product performance")
        
        self.results['sales'] = sales_metrics
        
        # Generate sales alerts
        if sales_metrics.get('monthly_growth', 0) < -10:
            self.alerts.append({
                'type': 'sales',
                'level': 'warning',
                'title': 'Sales Decline',
                'message': f"Monthly sales declined by {abs(sales_metrics['monthly_growth']):.1f}%"
            })
    
    def _analyze_inventory(self, inventory_df):
        """Analyze inventory data."""
        self.logger.info("Analyzing inventory data...")
        
        inventory_metrics = {
            'total_products': len(inventory_df)
        }
        
        # Find stock column
        stock_col = None
        for col in inventory_df.columns:
            col_lower = col.lower()
            if 'stock' in col_lower and ('current' in col_lower or 'available' in col_lower):
                stock_col = col
                break
        
        # Find days supply column
        days_col = None
        for col in inventory_df.columns:
            col_lower = col.lower()
            if 'days' in col_lower:
                days_col = col
                break
        
        if stock_col and stock_col in inventory_df.columns:
            try:
                total_stock = pd.to_numeric(inventory_df[stock_col], errors='coerce').sum()
                inventory_metrics['total_stock'] = total_stock
                
                # Low stock analysis
                if days_col and days_col in inventory_df.columns:
                    days_series = pd.to_numeric(inventory_df[days_col], errors='coerce')
                    low_stock = inventory_df[days_series < self.config.THRESHOLDS['low_stock_days']]
                    overstock = inventory_df[days_series > self.config.THRESHOLDS['overstock_days']]
                    
                    inventory_metrics['low_stock_count'] = len(low_stock)
                    inventory_metrics['overstock_count'] = len(overstock)
                    inventory_metrics['low_stock_products'] = low_stock.head(10)
                    inventory_metrics['overstock_products'] = overstock.head(10)
                    
                    # Generate inventory alerts
                    if len(low_stock) > 0:
                        critical_low = low_stock[days_series[low_stock.index] < 3]
                        if len(critical_low) > 0:
                            self.alerts.append({
                                'type': 'inventory',
                                'level': 'critical',
                                'title': 'Critical Stock Shortage',
                                'message': f"{len(critical_low)} products have less than 3 days of stock"
                            })
                        
                        self.alerts.append({
                            'type': 'inventory',
                            'level': 'warning',
                            'title': 'Low Stock Alert',
                            'message': f"{len(low_stock)} products have less than {self.config.THRESHOLDS['low_stock_days']} days of stock"
                        })
                    
                    if len(overstock) > 0:
                        self.alerts.append({
                            'type': 'inventory',
                            'level': 'info',
                            'title': 'Overstock Detected',
                            'message': f"{len(overstock)} products have more than {self.config.THRESHOLDS['overstock_days']} days of stock"
                        })
            except:
                self.logger.warning("Could not analyze inventory stock data")
        
        self.results['inventory'] = inventory_metrics
    
    def _analyze_advertising(self, advertising_df):
        """Analyze advertising data."""
        self.logger.info("Analyzing advertising data...")
        
        advertising_metrics = {}
        
        # Find spend column
        spend_col = None
        for col in advertising_df.columns:
            if 'spend' in col.lower():
                spend_col = col
                break
        
        # Find clicks column
        clicks_col = None
        for col in advertising_df.columns:
            if 'clicks' in col.lower():
                clicks_col = col
                break
        
        # Find impressions column
        impressions_col = None
        for col in advertising_df.columns:
            if 'impressions' in col.lower():
                impressions_col = col
                break
        
        if spend_col and spend_col in advertising_df.columns:
            advertising_metrics['total_spend'] = pd.to_numeric(advertising_df[spend_col], errors='coerce').sum()
        
        if clicks_col and clicks_col in advertising_df.columns:
            advertising_metrics['total_clicks'] = pd.to_numeric(advertising_df[clicks_col], errors='coerce').sum()
        
        if impressions_col and impressions_col in advertising_df.columns:
            advertising_metrics['total_impressions'] = pd.to_numeric(advertising_df[impressions_col], errors='coerce').sum()
        
        # Calculate metrics if data is available
        if advertising_metrics.get('total_spend', 0) > 0 and advertising_metrics.get('total_clicks', 0) > 0:
            advertising_metrics['avg_cpc'] = (
                advertising_metrics['total_spend'] / advertising_metrics['total_clicks']
            )
            
            if advertising_metrics.get('total_impressions', 0) > 0:
                advertising_metrics['ctr'] = (
                    advertising_metrics['total_clicks'] / advertising_metrics['total_impressions']
                ) * 100
        
        # Campaign performance
        if 'campaign_name' in advertising_df.columns:
            try:
                campaign_performance = advertising_df.groupby('campaign_name').agg({
                    spend_col: 'sum' if spend_col and spend_col in advertising_df.columns else pd.Series([0]).sum,
                    clicks_col: 'sum' if clicks_col and clicks_col in advertising_df.columns else pd.Series([0]).sum
                }).reset_index()
                
                campaign_performance.columns = ['campaign_name', 'spend', 'clicks']
                
                if 'sales' in advertising_df.columns:
                    campaign_performance['sales'] = advertising_df.groupby('campaign_name')['sales'].sum().values
                    campaign_performance['roas'] = (
                        campaign_performance['sales'] / campaign_performance['spend']
                    ).replace([np.inf, -np.inf], 0).fillna(0).round(2)
                    
                    # Identify underperforming campaigns
                    high_acos_campaigns = campaign_performance[
                        campaign_performance['roas'] < (100 / self.config.THRESHOLDS['high_acos'])
                    ]
                    
                    if len(high_acos_campaigns) > 0:
                        advertising_metrics['high_acos_campaigns'] = high_acos_campaigns.head(5)
                        
                        self.alerts.append({
                            'type': 'advertising',
                            'level': 'warning',
                            'title': 'High ACOS Campaigns',
                            'message': f"{len(high_acos_campaigns)} campaigns have ACOS > {self.config.THRESHOLDS['high_acos']}%"
                        })
                
                advertising_metrics['campaign_performance'] = campaign_performance.sort_values(
                    'spend', ascending=False
                ).head(10)
            except:
                self.logger.warning("Could not analyze campaign performance")
        
        self.results['advertising'] = advertising_metrics
    
    def _analyze_reviews(self, reviews_df):
        """Analyze customer reviews."""
        self.logger.info("Analyzing customer reviews...")
        
        reviews_metrics = {
            'total_reviews': len(reviews_df)
        }
        
        # Rating analysis
        if 'rating' in reviews_df.columns:
            try:
                reviews_metrics['avg_rating'] = pd.to_numeric(reviews_df['rating'], errors='coerce').mean()
                
                # Low-rated products
                ratings = pd.to_numeric(reviews_df['rating'], errors='coerce')
                low_rated = reviews_df[ratings < self.config.THRESHOLDS['low_rating']]
                if len(low_rated) > 0:
                    reviews_metrics['low_rated_count'] = len(low_rated)
                    
                    self.alerts.append({
                        'type': 'reviews',
                        'level': 'warning',
                        'title': 'Low Ratings Alert',
                        'message': f"{len(low_rated)} reviews have rating < {self.config.THRESHOLDS['low_rating']}"
                    })
            except:
                self.logger.warning("Could not analyze review ratings")
        
        self.results['reviews'] = reviews_metrics
    
    def get_kpi_summary(self):
        """Generate a summary of key performance indicators."""
        kpis = {}
        
        # Sales KPIs
        if 'sales' in self.results:
            sales = self.results['sales']
            kpis['total_revenue'] = sales.get('total_revenue', 0)
            kpis['total_transactions'] = sales.get('total_transactions', 0)
            kpis['avg_order_value'] = sales.get('avg_order_value', 0)
            kpis['monthly_growth'] = sales.get('monthly_growth', 0)
        
        # Inventory KPIs
        if 'inventory' in self.results:
            inventory = self.results['inventory']
            kpis['total_products'] = inventory.get('total_products', 0)
            kpis['low_stock_products'] = inventory.get('low_stock_count', 0)
            kpis['overstock_products'] = inventory.get('overstock_count', 0)
        
        # Advertising KPIs
        if 'advertising' in self.results:
            ads = self.results['advertising']
            kpis['total_ad_spend'] = ads.get('total_spend', 0)
            kpis['avg_cpc'] = ads.get('avg_cpc', 0)
        
        # Reviews KPIs
        if 'reviews' in self.results:
            reviews = self.results['reviews']
            kpis['avg_rating'] = reviews.get('avg_rating', 0)
            kpis['total_reviews'] = reviews.get('total_reviews', 0)
        
        return kpis

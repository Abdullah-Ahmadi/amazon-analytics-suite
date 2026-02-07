"""
Data loading and validation module - FIXED VERSION
"""

import pandas as pd
import numpy as np
from pathlib import Path
import re
from datetime import datetime
import logging

class DataLoader:
    """Handles loading and validation of CSV files."""
    
    def __init__(self, logger=None):
        self.logger = logger or logging.getLogger(__name__)
        self.data = {}
        self.stats = {}
        
    def discover_files(self, directory="."):
        """
        Automatically discover CSV files in the directory.
        
        Returns:
            dict: Dictionary mapping file types to file paths
        """
        self.logger.info("Discovering CSV files...")
        directory = Path(directory)
        
        csv_files = list(directory.glob("*.csv"))
        self.logger.info(f"Found {len(csv_files)} CSV files")
        
        # Categorize files based on filename patterns
        categorized = {
            'sales': [],
            'inventory': [],
            'advertising': [],
            'reviews': [],
            'other': []
        }
        
        for file_path in csv_files:
            filename = file_path.name.lower()
            
            if any(pattern in filename for pattern in ["sales", "transaction", "order"]):
                categorized['sales'].append(file_path)
            elif any(pattern in filename for pattern in ["inventory", "stock"]):
                categorized['inventory'].append(file_path)
            elif any(pattern in filename for pattern in ["ad", "spend", "campaign"]):
                categorized['advertising'].append(file_path)
            elif any(pattern in filename for pattern in ["review", "rating"]):
                categorized['reviews'].append(file_path)
            else:
                categorized['other'].append(file_path)
        
        # Log discovery results
        for file_type, files in categorized.items():
            if files:
                self.logger.info(f"  {file_type.upper()}: {len(files)} files")
                for f in files[:3]:  # Show first 3 files
                    self.logger.info(f"    - {f.name}")
                if len(files) > 3:
                    self.logger.info(f"    ... and {len(files)-3} more")
        
        return categorized
    
    def load_sales_data(self, file_paths):
        """Load and validate sales data."""
        self.logger.info("Loading sales data...")
        
        if not file_paths:
            self.logger.warning("No sales files found")
            return None
        
        try:
            # Load first sales file found
            df = pd.read_csv(file_paths[0])
            self.logger.info(f"Loaded {len(df)} rows from {file_paths[0].name}")
            
            # Standardize column names
            df.columns = [self._standardize_column_name(col) for col in df.columns]
            
            # Clean and preprocess
            df = self._clean_sales_data(df)
            
            self.data['sales'] = df
            self.stats['sales'] = {
                'rows': len(df),
                'columns': len(df.columns),
                'date_range': self._get_date_range(df, self._find_date_column(df.columns))
            }
            
            return df
            
        except Exception as e:
            self.logger.error(f"Error loading sales data: {e}")
            raise
    
    def load_inventory_data(self, file_paths):
        """Load and validate inventory data."""
        self.logger.info("Loading inventory data...")
        
        if not file_paths:
            self.logger.warning("No inventory files found")
            return None
        
        try:
            df = pd.read_csv(file_paths[0])
            self.logger.info(f"Loaded {len(df)} rows from {file_paths[0].name}")
            
            # Standardize column names
            df.columns = [self._standardize_column_name(col) for col in df.columns]
            
            # Clean and preprocess
            df = self._clean_inventory_data(df)
            
            self.data['inventory'] = df
            self.stats['inventory'] = {
                'rows': len(df),
                'columns': len(df.columns),
                'unique_products': df['product_name'].nunique() if 'product_name' in df.columns else 0
            }
            
            return df
            
        except Exception as e:
            self.logger.error(f"Error loading inventory data: {e}")
            raise
    
    def load_advertising_data(self, file_paths):
        """Load and validate advertising data."""
        self.logger.info("Loading advertising data...")
        
        if not file_paths:
            self.logger.warning("No advertising files found")
            return None
        
        try:
            df = pd.read_csv(file_paths[0])
            self.logger.info(f"Loaded {len(df)} rows from {file_paths[0].name}")
            
            # Standardize column names
            df.columns = [self._standardize_column_name(col) for col in df.columns]
            
            self.data['advertising'] = df
            self.stats['advertising'] = {
                'rows': len(df),
                'columns': len(df.columns)
            }
            
            return df
            
        except Exception as e:
            self.logger.error(f"Error loading advertising data: {e}")
            raise
    
    def load_reviews_data(self, file_paths):
        """Load and validate reviews data."""
        self.logger.info("Loading reviews data...")
        
        if not file_paths:
            self.logger.warning("No reviews files found")
            return None
        
        try:
            df = pd.read_csv(file_paths[0])
            self.logger.info(f"Loaded {len(df)} rows from {file_paths[0].name}")
            
            # Standardize column names
            df.columns = [self._standardize_column_name(col) for col in df.columns]
            
            self.data['reviews'] = df
            self.stats['reviews'] = {
                'rows': len(df),
                'columns': len(df.columns),
                'average_rating': df['rating'].mean() if 'rating' in df.columns else 0
            }
            
            return df
            
        except Exception as e:
            self.logger.error(f"Error loading reviews data: {e}")
            raise
    
    def _standardize_column_name(self, name):
        """Standardize column names for consistency."""
        if not isinstance(name, str):
            name = str(name)
        
        # Convert to lowercase, replace spaces/special chars with underscore
        name = name.lower().strip()
        name = re.sub(r'[^\w\s]', '_', name)  # Replace special chars
        name = re.sub(r'\s+', '_', name)  # Replace spaces
        name = re.sub(r'_+', '_', name)  # Remove multiple underscores
        
        return name
    
    def _find_date_column(self, columns):
        """Find date column in list of columns."""
        for col in columns:
            col_lower = col.lower()
            if any(key in col_lower for key in ['date', 'time', 'order_date', 'transaction_date']):
                return col
        return None
    
    def _clean_sales_data(self, df):
        """Clean and preprocess sales data."""
        df_clean = df.copy()
        
        # Find date column and convert to datetime
        date_col = self._find_date_column(df_clean.columns)
        if date_col:
            try:
                df_clean[date_col] = pd.to_datetime(df_clean[date_col], errors='coerce')
            except:
                self.logger.warning(f"Could not parse date column: {date_col}")
        
        # Convert numeric columns
        numeric_patterns = ['quantity', 'price', 'total', 'amount', 'qty', 'cost', 'revenue']
        for col in df_clean.columns:
            col_lower = col.lower()
            if any(pattern in col_lower for pattern in numeric_patterns):
                try:
                    df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce')
                except:
                    pass
        
        return df_clean
    
    def _clean_inventory_data(self, df):
        """Clean and preprocess inventory data."""
        df_clean = df.copy()
        
        # Convert numeric columns
        numeric_patterns = ['stock', 'quantity', 'qty', 'days', 'supply', 'inbound']
        for col in df_clean.columns:
            col_lower = col.lower()
            if any(pattern in col_lower for pattern in numeric_patterns):
                try:
                    df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)
                except:
                    pass
        
        return df_clean
    
    def _get_date_range(self, df, date_column):
        """Get date range from dataframe."""
        if date_column and date_column in df.columns:
            try:
                dates = pd.to_datetime(df[date_column], errors='coerce')
                valid_dates = dates.dropna()
                if len(valid_dates) > 0:
                    return f"{valid_dates.min().date()} to {valid_dates.max().date()}"
            except:
                pass
        return "Not available"

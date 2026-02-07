"""
Configuration settings for the Amazon Analytics Suite.
"""

from pathlib import Path
from dataclasses import dataclass
from datetime import datetime

@dataclass
class Config:
    """Application configuration."""

    # File patterns to detect
    SALES_PATTERNS = ["*sales*", "*transaction*", "*order*"]
    INVENTORY_PATTERNS = ["*inventory*", "*stock*"]
    ADVERTISING_PATTERNS = ["*ad*", "*spend*", "*campaign*"]
    REVIEWS_PATTERNS = ["*review*", "*rating*"]

    # Directory settings
    OUTPUT_DIR = "output"
    LOG_DIR = "logs"

    # Excel settings
    DEFAULT_SHEET_NAMES = {
        "dashboard": "Executive Dashboard",
        "sales": "Sales Analysis",
        "inventory": "Inventory Health",
        "advertising": "Advertising ROI",
        "reviews": "Customer Reviews",
        "alerts": "Actionable Alerts",
        "raw_data": "Raw Data"
    }

    # Business thresholds
    THRESHOLDS = {
        "low_stock_days": 7,
        "overstock_days": 30,
        "excessive_stock_days": 90,
        "low_profit_margin": 15,
        "high_acos": 40,  # Advertising Cost of Sale
        "low_rating": 3.0
    }

    # Colors for Excel formatting (RGB hex)
    COLORS = {
        "urgent_red": "FFC7CE",
        "warning_yellow": "FFEB9C",
        "good_green": "C6EFCE",
        "info_blue": "BDD7EE",
        "header_blue": "4472C4",
        "dark_blue": "2F5496"
    }

    @property
    def timestamp(self):
        """Get current timestamp for file naming."""
        return datetime.now().strftime("%Y%m%d_%H%M%S")

    @property
    def output_filename(self):
        """Generate output filename."""
        return f"Amazon_Dashboard_{self.timestamp}.xlsx"

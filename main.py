#!/usr/bin/env python3
"""
Amazon Analytics Suite - Main Entry Point
A comprehensive tool for processing Amazon seller reports and generating insights.
"""

import sys
import os
from datetime import datetime
from pathlib import Path

# Add src directory to Python path
src_path = Path(__file__).parent / 'src'
sys.path.insert(0, str(src_path))

from ui_handler import UIManager
from logger import setup_logger
from config import Config

def main():
    """Main execution function."""
    try:
        # Setup logging
        logger = setup_logger()
        logger.info("=" * 60)
        logger.info("Amazon Analytics Suite - Starting Execution")
        logger.info("=" * 60)

        # Initialize configuration
        config = Config()

        # Create necessary directories
        output_dir = Path(config.OUTPUT_DIR)
        output_dir.mkdir(exist_ok=True)

        # Initialize and run UI
        ui = UIManager(config, logger)
        ui.run()

    except Exception as e:
        print(f"Critical error: {e}")
        print("Please check the log file for details.")
        sys.exit(1)

if __name__ == "__main__":
    main()

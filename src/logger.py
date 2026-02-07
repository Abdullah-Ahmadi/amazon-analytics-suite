"""
Logging configuration for the Amazon Analytics Suite.
"""

import logging
from pathlib import Path
from datetime import datetime

def setup_logger():
    """Configure and return logger instance."""

    # Create logs directory
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    # Create log filename with timestamp
    log_file = log_dir / f"amazon_analytics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    # Configure logger
    logger = logging.getLogger("AmazonAnalytics")
    logger.setLevel(logging.DEBUG)

    # Remove existing handlers
    logger.handlers.clear()

    # File handler (overwrite mode)
    file_handler = logging.FileHandler(log_file, mode='w')
    file_handler.setLevel(logging.DEBUG)

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # Formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger

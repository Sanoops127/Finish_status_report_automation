import logging
import os
import sys
from pathlib import Path


def setup_logger():
    # Ensure logs directory exists
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    log_file = log_dir / "automation.log"

    # Create root logger
    logger = logging.getLogger("depotnet_bot")
    logger.setLevel(logging.INFO)
    logger.propagate = False  # Strictly prevent double logging

    # Avoid adding multiple handlers if setup is called multiple times
    if not logger.handlers:
        # File handler (always needed for persistence)
        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(logging.INFO)

        # Formatting
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        # Console handler: Skip if console logging is already managed (e.g. by Celery)
        # We detect Celery by checking if the task is running or if the worker is starting
        is_celery = any('celery' in arg.lower() for arg in sys.argv)
        if not is_celery:
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.INFO)
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)

    return logger

# Create a singleton instance for easy import
logger = setup_logger()

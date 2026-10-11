"""
Logging utility with Rich console output and file handler.
"""
import logging
import sys
from pathlib import Path
from typing import Optional
from rich.console import Console
from rich.logging import RichHandler
from config import DEFAULT_LOG_PATH

console = Console()


def setup_logger(log_file: Optional[Path] = DEFAULT_LOG_PATH, level: int = logging.INFO) -> logging.Logger:
    """Setup and configure application logger with Rich console output and file logging."""
    logger = logging.getLogger("leads_extractor")
    logger.setLevel(level)

    # Clear existing handlers to prevent duplication
    if logger.hasHandlers():
        logger.handlers.clear()

    # Rich Console Handler
    rich_handler = RichHandler(
        console=console,
        show_time=True,
        show_path=False,
        markup=True,
        rich_tracebacks=True,
    )
    rich_handler.setLevel(level)
    logger.addHandler(rich_handler)

    # File Handler
    if log_file:
        Path(log_file).parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_formatter = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        file_handler.setFormatter(file_formatter)
        file_handler.setLevel(level)
        logger.addHandler(file_handler)

    return logger

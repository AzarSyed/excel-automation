"""Centralized logging setup for all modules."""
import logging
import sys
from pathlib import Path
from datetime import datetime

LOG_DIR = Path(__file__).parent.parent / "exports" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)


def setup_logger(name: str, level: int = logging.INFO) -> logging.Logger:
    """Return a named logger with console + daily file handlers."""
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(level)
    fmt = logging.Formatter(
        "%(asctime)s | %(name)-15s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    log_file = LOG_DIR / f"{datetime.now().strftime('%Y-%m-%d')}.log"
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger

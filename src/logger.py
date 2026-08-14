import logging
import sys

# Standard logger setup
logger = logging.getLogger("GLASS_PI")
logger.setLevel(logging.INFO)

if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)

def info(msg):
    logger.info(msg)

def warning(msg):
    logger.warning(msg)

def error(msg):
    logger.error(msg)
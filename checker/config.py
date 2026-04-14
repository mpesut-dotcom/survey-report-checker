"""Configuration for pipeline v2."""
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# API
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# Models
FLASH_MODEL = os.getenv("FLASH_MODEL", "gemini-2.5-flash-lite")
PRO_MODEL = os.getenv("PRO_MODEL", FLASH_MODEL)

# Paths
BASE_DIR = Path(__file__).parent.parent  # workspace root

# LibreOffice
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"

# Image rendering
SLIDE_DPI = 150  # 150 DPI → zoom ~2.08 for PyMuPDF
SLIDE_ZOOM = SLIDE_DPI / 72.0

# LLM settings
FLASH_TEMPERATURE = 0.1
PRO_TEMPERATURE = 0.1
MAX_OUTPUT_TOKENS = 65536
MAX_RETRIES = 3
RETRY_BASE_DELAY = 2  # seconds, exponential backoff

# Batch sizes
EXTRACTION_BATCH_SIZE = 1   # 1 slide per call (with image)
VERIFICATION_BATCH_SIZE = 1 # 1 slide per call (thorough)

# Concurrency
MAX_CONCURRENT_LLM = 4  # max parallel LLM calls per phase

# Cost estimates (per 1M tokens, approximate)
COST_ESTIMATES = {
    "gemini-2.5-flash-lite": {"input": 0.10, "output": 0.40},
    "gemini-3-flash-preview": {"input": 0.15, "output": 0.60},
    "gemini-3.1-pro-preview": {"input": 1.25, "output": 10.00},
}

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))
from dotenv import load_dotenv

from src.excel_mcp.server import run_streamable_http

load_dotenv(override=True)

if __name__ == "__main__":
    run_streamable_http()

"""Source entry point for the isolated desktop Office PDF setup check."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from core.pdf_setup import main

if __name__ == "__main__":
    sys.exit(main())

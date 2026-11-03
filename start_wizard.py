"""Packaged entry point: capture errors before importing the application or Tk."""
from core.startup import launch

if __name__ == '__main__':
    raise SystemExit(launch())

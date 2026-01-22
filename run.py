#!/usr/bin/env python
"""Quick start script for TIQ Assistant."""

import subprocess
import sys
from pathlib import Path

# Add src to path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))


def main():
    """Run the Streamlit app."""
    app_path = src_path / "tiq_assistant" / "web" / "streamlit_app.py"

    print("Starting TIQ Assistant...")
    print(f"Open http://localhost:8501 in your browser")
    print("Press Ctrl+C to stop")
    print()

    subprocess.run([
        sys.executable, "-m", "streamlit", "run",
        str(app_path),
        "--server.headless", "true",
    ])


if __name__ == "__main__":
    main()

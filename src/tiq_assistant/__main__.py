"""Entry point for TIQ Assistant."""

import subprocess
import sys


def main():
    """Run TIQ Assistant based on command line arguments."""
    # Check for mode argument
    if len(sys.argv) > 1:
        mode = sys.argv[1].lower()
    else:
        mode = "desktop"  # Default to desktop mode

    if mode == "web":
        run_web()
    elif mode == "desktop":
        run_desktop()
    else:
        print(f"Unknown mode: {mode}")
        print("Usage: python -m tiq_assistant [desktop|web]")
        print("  desktop - Run as system tray application (default)")
        print("  web     - Run as Streamlit web application")
        sys.exit(1)


def run_web():
    """Run the Streamlit web app."""
    subprocess.run([
        sys.executable, "-m", "streamlit", "run",
        "src/tiq_assistant/web/streamlit_app.py",
        "--server.headless", "true",
    ])


def run_desktop():
    """Run the desktop application."""
    from tiq_assistant.desktop.app import main as desktop_main
    sys.exit(desktop_main())


if __name__ == "__main__":
    main()

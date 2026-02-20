# run_standalone.py
import os
import sys
import socket
import webbrowser
from threading import Timer
from app import create_app
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def find_free_port(start_port=8000, max_attempts=100):
    """Find a free port starting from start_port."""
    for port in range(start_port, start_port + max_attempts):
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            result = sock.connect_ex(('127.0.0.1', port))
            sock.close()
            if result != 0:  # Port is free
                return port
        except Exception:
            continue
    return start_port  # Fallback

def open_browser(port):
    """Open browser after a short delay."""
    def open_url():
        url = f"http://127.0.0.1:{port}"
        try:
            webbrowser.open(url)
            logger.info(f"Application started! Opening browser at {url}")
        except Exception as e:
            logger.warning(f"Could not open browser: {e}")
            print(f"\n======================================")
            print(f" DQA Dashboard is running!")
            print(f" Open your browser and go to:")
            print(f" {url}")
            print(f"======================================\n")
    
    Timer(1.5, open_url).start()

def main():
    """Main entry point for standalone application."""
    logger.info("Starting DQA Dashboard in standalone mode...")
    
    # Force standalone mode
    import app.routes
    app.routes.STANDALONE_MODE = True
    
    # Log system info for debugging
    logger.info(f"Python executable: {sys.executable}")
    logger.info(f"Current directory: {os.getcwd()}")
    if getattr(sys, 'frozen', False):
        logger.info("Running as PyInstaller executable")
        if hasattr(sys, '_MEIPASS'):
            logger.info(f"Bundled data location: {sys._MEIPASS}") # type: ignore
    
    # Create app
    app = create_app()
    
    # Find free port
    port = find_free_port()
    logger.info(f"Using port: {port}")
    
    # Open browser after startup
    open_browser(port)
    
    # Run the app
    try:
        print(f"\n{'='*50}")
        print(f"  DQA Dashboard - Standalone Version")
        print(f"  Running on: http://127.0.0.1:{port}")
        print(f"  Data directory: {os.path.join(os.path.dirname(sys.executable), 'DQA_Data')}")
        print(f"  Press Ctrl+C to stop")
        print(f"{'='*50}\n")
        
        app.run(
            host="127.0.0.1",
            port=port,
            debug=False,
            threaded=True,
            use_reloader=False
        )
    except KeyboardInterrupt:
        logger.info("Shutting down...")
        print("\nApplication stopped.")
    except Exception as e:
        logger.error(f"Error starting application: {e}")
        print(f"\nError: {e}")
        print("Please check the log file for details.")
        sys.exit(1)

if __name__ == "__main__":
    main()
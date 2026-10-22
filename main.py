import tkinter as tk
import logging
import sys
import os
import atexit

# Add the current directory to sys.path to ensure modules can be imported
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# First-run setup for executable distribution
if getattr(sys, 'frozen', False):
    from modules.first_run_setup import run_first_time_setup
    run_first_time_setup()

from modules.logging_config import setup_logging

# Setup logging
setup_logging(log_file="aplikasiAlatUji.log")
logger = logging.getLogger(__name__)

from modules.ui.main_window import ExcelProcessorGUI
from modules.db_controller import initialize_db_pool, close_pool
from modules.shutdown_manager import get_shutdown_manager

def main():
    # Initialize shutdown manager
    shutdown_mgr = get_shutdown_manager()
    
    # Initialize configuration from environment variables/.env
    try:
        from modules.config import get_config
        config = get_config()
        config.print_config_summary()
        logger.info("Configuration initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize configuration: {e}")
        print(f"Warning: Configuration initialization failed: {e}")
    
    # Initialize database connection pool at startup
    try:
        initialize_db_pool()
        logger.info("Database connection pool initialized")
    except Exception as e:
        logger.error(f"Failed to initialize database pool: {e}")
        print(f"Warning: Database pool initialization failed: {e}")
    
    # Register graceful shutdown on exit
    atexit.register(shutdown_mgr.shutdown)
    
    root = tk.Tk()
    # Set icon if available
    try:
        # root.iconbitmap("icon.ico") 
        pass
    except:
        pass
        
    app = ExcelProcessorGUI(root)
    
    # Register the root window for shutdown cleanup
    def on_closing():
        logger.info("User initiated application close")
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
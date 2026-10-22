"""First-run setup for executable distribution"""
import subprocess
import sys
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


def setup_playwright():
    """Download Playwright browser on first run"""
    marker = Path.home() / '.odoo_app_playwright_setup'
    
    if marker.exists():
        return True
    
    try:
        logger.info("First-time setup: Installing Playwright browser...")
        print("First-time setup: Downloading browser components (170MB)...")
        print("This only happens once and may take 2-3 minutes.")
        
        result = subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            capture_output=True,
            text=True,
            timeout=300
        )
        
        if result.returncode == 0:
            marker.touch()
            logger.info("Playwright browser installed successfully")
            print("Setup complete!")
            return True
        else:
            logger.error(f"Playwright install failed: {result.stderr}")
            print(f"Warning: Browser install failed. App may not work correctly.")
            return False
            
    except subprocess.TimeoutExpired:
        logger.error("Playwright install timeout")
        print("Warning: Browser download timed out.")
        return False
    except Exception as e:
        logger.error(f"Playwright install error: {e}")
        print(f"Warning: Browser install error: {e}")
        return False


def get_config_dir():
    """Get user config directory for executable"""
    if getattr(sys, 'frozen', False):
        # Running as executable - use AppData
        appdata = Path.home() / 'AppData' / 'Roaming' / 'OdooCompressionTest'
        appdata.mkdir(parents=True, exist_ok=True)
        return appdata
    else:
        # Running as script - use project root
        return Path(__file__).parent.parent


def ensure_env_file():
    """Ensure .env exists in config directory"""
    config_dir = get_config_dir()
    env_path = config_dir / '.env'
    
    if env_path.exists():
        return env_path
    
    # Copy .env.example if available
    if getattr(sys, 'frozen', False):
        # In executable, .env.example is bundled
        import os
        bundle_dir = Path(sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.dirname(sys.executable))
        example = bundle_dir / '.env.example'
    else:
        # In development
        example = Path(__file__).parent.parent / '.env.example'
    
    if example.exists():
        import shutil
        shutil.copy(example, env_path)
        logger.info(f"Created .env from template at {env_path}")
    else:
        # Create minimal .env
        env_path.write_text("""# Database Configuration
DB_HOST=localhost
DB_PORT=5432
DB_NAME=alatujidb
DB_USER=postgres
DB_PASSWORD=

# Web Automation
SESSION_TOKEN=

# Application
LOG_LEVEL=INFO
ENVIRONMENT=production
""")
        logger.info(f"Created minimal .env at {env_path}")
    
    return env_path


def get_data_directories():
    """Get data directories (AppData for exe, project root for dev)"""
    config_dir = get_config_dir()
    
    dirs = {
        'output': config_dir / 'Output',
        'rencana': config_dir / 'Rencana',
        'pengujian': config_dir / 'Pengujian',
        'data_input': config_dir / 'Data Input Pengujian',
    }
    
    # Create directories
    for dir_path in dirs.values():
        dir_path.mkdir(parents=True, exist_ok=True)
    
    return dirs


def run_first_time_setup():
    """Run all first-time setup tasks"""
    logger.info("Running first-time setup checks...")
    
    # Ensure .env exists
    env_path = ensure_env_file()
    logger.info(f"Config file: {env_path}")
    
    # Ensure data directories exist
    dirs = get_data_directories()
    logger.info(f"Data directories ready: {len(dirs)} folders")
    
    # Setup Playwright (only if running as executable)
    if getattr(sys, 'frozen', False):
        setup_playwright()
    
    logger.info("First-time setup complete")
    return True

"""
Configuration Management Module

Centralized credential and configuration management using environment variables.
Supports .env files via python-dotenv for local development.

Security Best Practices:
- All sensitive credentials sourced from environment variables
- .env file template provided (never commit actual .env)
- Fallback to config.cnf only for backward compatibility
- Validate required credentials on startup
"""

import os
import logging
from pathlib import Path
from typing import Optional

try:
    # pyrefly: ignore [missing-import]
    from dotenv import load_dotenv, find_dotenv
except ImportError:
    # Fallback if python-dotenv not installed
    def load_dotenv(*args, **kwargs):
        pass
    def find_dotenv(*args, **kwargs):
        return None

logger = logging.getLogger(__name__)


class EnvironmentConfig:
    """
    Centralized configuration management.
    
    Priority order for credentials:
    1. Environment variables (highest priority)
    2. .env file (local development)
    3. config.cnf (legacy fallback)
    4. Defaults (if specified)
    """
    
    # Class-level loaded flag
    _loaded = False
    
    def __init__(self):
        """Initialize configuration on first use"""
        if not EnvironmentConfig._loaded:
            self._load_configuration()
            EnvironmentConfig._loaded = True
    
    @staticmethod
    def _load_configuration():
        """Load environment configuration"""
        # Try to load from .env file
        env_file = find_dotenv()
        if env_file:
            logger.info(f"Loading .env from: {env_file}")
            load_dotenv(env_file)
        else:
            logger.debug(".env file not found, using system environment variables")
    
    # ========================================================================
    # DATABASE CONFIGURATION
    # ========================================================================
    
    @staticmethod
    def get_db_host() -> str:
        """Get database host"""
        return os.getenv('DB_HOST', 'localhost')
    
    @staticmethod
    def get_db_port() -> str:
        """Get database port"""
        return os.getenv('DB_PORT', '5432')
    
    @staticmethod
    def get_db_name() -> str:
        """Get database name"""
        return os.getenv('DB_NAME', 'alatujidb')
    
    @staticmethod
    def get_db_user() -> str:
        """Get database username"""
        return os.getenv('DB_USER', 'postgres')
    
    @staticmethod
    def get_db_password() -> Optional[str]:
        """Get database password (no default - requires explicit setting)"""
        password = os.getenv('DB_PASSWORD')
        if not password:
            logger.warning("DB_PASSWORD not set in environment variables")
        return password
    
    @staticmethod
    def get_db_connect_timeout() -> int:
        """Get database connection timeout"""
        return int(os.getenv('DB_CONNECT_TIMEOUT', '10'))
    
    @staticmethod
    def get_db_statement_timeout() -> int:
        """Get database statement timeout"""
        return int(os.getenv('DB_STATEMENT_TIMEOUT', '30'))
    
    # ========================================================================
    # CONNECTION POOL CONFIGURATION
    # ========================================================================
    
    @staticmethod
    def get_db_pool_size() -> int:
        """Get connection pool size"""
        return int(os.getenv('DB_POOL_SIZE', '5'))
    
    @staticmethod
    def get_db_pool_overflow() -> int:
        """Get connection pool overflow"""
        return int(os.getenv('DB_POOL_OVERFLOW', '10'))
    
    # ========================================================================
    # SELENIUM/WEB AUTOMATION CONFIGURATION
    # ========================================================================
    
    @staticmethod
    def get_selenium_url() -> str:
        """Get Selenium test URL"""
        return os.getenv('SELENIUM_URL', 'https://rmc.adhimix.web.id/web/login')
    
    @staticmethod
    def get_selenium_user() -> Optional[str]:
        """Get Selenium username (no default - requires explicit setting)"""
        user = os.getenv('SELENIUM_USER')
        if not user:
            logger.warning("SELENIUM_USER not set in environment variables")
        return user
    
    @staticmethod
    def get_selenium_password() -> Optional[str]:
        """Get Selenium password (no default - requires explicit setting)"""
        password = os.getenv('SELENIUM_PASSWORD')
        if not password:
            logger.warning("SELENIUM_PASSWORD not set in environment variables")
        return password
    
    @staticmethod
    def get_session_token() -> Optional[str]:
        """Get session token for token-based auth (alternative to username/password)"""
        return os.getenv('SESSION_TOKEN')
    
    # ========================================================================
    # APPLICATION CONFIGURATION
    # ========================================================================
    
    @staticmethod
    def get_log_level() -> str:
        """Get application log level"""
        return os.getenv('LOG_LEVEL', 'INFO')
    
    @staticmethod
    def get_environment() -> str:
        """Get application environment (development, staging, production)"""
        return os.getenv('ENVIRONMENT', 'development')
    
    @staticmethod
    def is_production() -> bool:
        """Check if running in production"""
        return EnvironmentConfig.get_environment().lower() == 'production'
    
    @staticmethod
    def is_development() -> bool:
        """Check if running in development"""
        return EnvironmentConfig.get_environment().lower() == 'development'
    
    # ========================================================================
    # VALIDATION & UTILITIES
    # ========================================================================
    
    @staticmethod
    def validate_required_credentials() -> bool:
        """
        Validate that all required credentials are set.
        
        Returns:
            bool: True if all required credentials are present
            
        Raises:
            ValueError: If any required credential is missing
        """
        config = EnvironmentConfig()
        missing = []
        
        # Check required credentials
        if not config.get_db_password():
            missing.append('DB_PASSWORD')
        
        # For web automation: either token OR username+password required
        has_token = bool(config.get_session_token())
        has_credentials = bool(config.get_selenium_user() and config.get_selenium_password())
        
        if not has_token and not has_credentials:
            missing.append('SESSION_TOKEN or (SELENIUM_USER + SELENIUM_PASSWORD)')
        
        if missing:
            error_msg = f"Missing required environment variables: {', '.join(missing)}"
            logger.error(error_msg)
            raise ValueError(error_msg)
        
        logger.info("All required credentials validated successfully")
        return True
    
    @staticmethod
    def get_database_url() -> str:
        """
        Build database connection URL for SQLAlchemy.
        
        Returns:
            str: PostgreSQL connection string
        """
        config = EnvironmentConfig()
        password = config.get_db_password()
        
        return (
            f"postgresql://{config.get_db_user()}:{password}"
            f"@{config.get_db_host()}:{config.get_db_port()}"
            f"/{config.get_db_name()}"
        )
    
    @staticmethod
    def print_config_summary():
        """Print configuration summary (without sensitive values)"""
        config = EnvironmentConfig()
        
        print("\n" + "="*70)
        print("APPLICATION CONFIGURATION SUMMARY")
        print("="*70)
        print(f"\nDatabase:")
        print(f"  Host: {config.get_db_host()}")
        print(f"  Port: {config.get_db_port()}")
        print(f"  Name: {config.get_db_name()}")
        print(f"  User: {config.get_db_user()}")
        print(f"  Password: {'***' if config.get_db_password() else 'NOT SET'}")
        
        print(f"\nConnection Pool:")
        print(f"  Pool Size: {config.get_db_pool_size()}")
        print(f"  Max Overflow: {config.get_db_pool_overflow()}")
        
        print(f"\nApplication:")
        print(f"  Environment: {config.get_environment()}")
        print(f"  Log Level: {config.get_log_level()}")
        
        print(f"\nSelenium:")
        print(f"  URL: {config.get_selenium_url()}")
        print(f"  User: {config.get_selenium_user() or 'NOT SET'}")
        print(f"  Password: {'***' if config.get_selenium_password() else 'NOT SET'}")
        
        print("="*70 + "\n")


# Singleton instance
_config_instance = None

def get_config() -> EnvironmentConfig:
    """Get global configuration instance"""
    global _config_instance
    if _config_instance is None:
        _config_instance = EnvironmentConfig()
    return _config_instance

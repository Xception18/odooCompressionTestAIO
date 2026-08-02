"""
Database Connection Pooling Module

Provides efficient database connection management using SQLAlchemy connection pooling.
Replaces the inefficient pattern of creating new connections for each query.

Uses centralized environment-based configuration for credentials.
"""

import logging
from sqlalchemy import create_engine, text
from sqlalchemy.pool import QueuePool
from sqlalchemy.exc import SQLAlchemyError
from modules.config import get_config

logger = logging.getLogger(__name__)


class DatabasePool:
    """
    Centralized database connection pool manager.
    
    Provides a single point of configuration for database connections across the application.
    Uses SQLAlchemy's QueuePool for efficient connection reuse.
    """
    
    _instance = None  # Singleton instance
    _engine = None
    
    def __new__(cls):
        """Ensure singleton pattern - only one pool instance"""
        if cls._instance is None:
            cls._instance = super(DatabasePool, cls).__new__(cls)
        return cls._instance
    
    def __init__(self):
        """Initialize database pool configuration"""
        if self._engine is None:
            self._config = self._load_config()
            self._engine = self._create_engine()
    
    @staticmethod
    def _load_config():
        """Load database configuration from environment variables or config.cnf"""
        config = get_config()
        
        return {
            'database': config.get_db_name(),
            'host': config.get_db_host(),
            'user': config.get_db_user(),
            'password': config.get_db_password(),
            'port': config.get_db_port(),
        }
    
    def _create_engine(self):
        """
        Create SQLAlchemy engine with connection pooling.
        
        Configuration:
        - pool_size: Number of connections to keep in pool (5)
        - max_overflow: Additional connections above pool_size (10)
        - pool_recycle: Recycle connections every hour (3600 seconds)
        - pool_pre_ping: Test connections before using
        """
        try:
            connection_string = (
                f"postgresql://{self._config['user']}:{self._config['password']}"
                f"@{self._config['host']}:{self._config['port']}"
                f"/{self._config['database']}"
            )
            
            engine = create_engine(
                connection_string,
                poolclass=QueuePool,
                pool_size=5,                    # Keep 5 connections ready
                max_overflow=10,                # Allow up to 10 additional connections
                pool_recycle=3600,              # Recycle connections after 1 hour
                pool_pre_ping=True,             # Test connection before using
                echo=False,                     # Set True for SQL logging
                connect_args={
                    'connect_timeout': 10,
                    'options': '-c statement_timeout=30000'  # 30 second statement timeout
                }
            )
            
            # Test connection
            with engine.connect() as conn:
                result = conn.execute(text("SELECT 1"))
                logger.info("✓ Database connection pool initialized successfully")
            
            return engine
            
        except Exception as e:
            logger.error(f"✗ Failed to create database engine: {e}")
            raise
    
    def get_engine(self):
        """Get the SQLAlchemy engine"""
        if self._engine is None:
            self._engine = self._create_engine()
        return self._engine
    
    def get_connection(self):
        """
        Get a database connection from the pool.
        
        Returns:
            sqlalchemy.engine.Connection: A connection from the pool
            
        Raises:
            SQLAlchemyError: If unable to get connection
        """
        try:
            return self._engine.connect()
        except SQLAlchemyError as e:
            logger.error(f"Failed to get database connection: {e}")
            raise
    
    def execute_query(self, query, params=None):
        """
        Execute a SELECT query and return results.
        
        Args:
            query (str): SQL query string
            params (dict): Query parameters
            
        Returns:
            list: Query results as list of dictionaries
        """
        try:
            conn = self.get_connection()
            try:
                result = conn.execute(text(query), params or {})
                return [dict(row) for row in result.fetchall()]
            finally:
                conn.close()
        except Exception as e:
            logger.error(f"Query execution failed: {e}")
            raise
    
    def execute_insert(self, query, params=None):
        """
        Execute an INSERT query.
        
        Args:
            query (str): SQL INSERT query
            params (dict): Query parameters
            
        Returns:
            int: Number of rows inserted
        """
        try:
            conn = self.get_connection()
            try:
                result = conn.execute(text(query), params or {})
                conn.commit()
                rows_affected = result.rowcount
                logger.info(f"Inserted {rows_affected} row(s)")
                return rows_affected
            except Exception as e:
                conn.rollback()
                raise
            finally:
                conn.close()
        except Exception as e:
            logger.error(f"Insert execution failed: {e}")
            raise
    
    def execute_update(self, query, params=None):
        """
        Execute an UPDATE query.
        
        Args:
            query (str): SQL UPDATE query
            params (dict): Query parameters
            
        Returns:
            int: Number of rows updated
        """
        try:
            conn = self.get_connection()
            try:
                result = conn.execute(text(query), params or {})
                conn.commit()
                rows_affected = result.rowcount
                logger.info(f"Updated {rows_affected} row(s)")
                return rows_affected
            except Exception as e:
                conn.rollback()
                raise
            finally:
                conn.close()
        except Exception as e:
            logger.error(f"Update execution failed: {e}")
            raise
    
    def close(self):
        """Close all connections in the pool"""
        if self._engine:
            self._engine.dispose()
            logger.info("Database pool closed")
    
    def get_pool_status(self):
        """Get current pool status for monitoring"""
        if self._engine and hasattr(self._engine.pool, 'size'):
            return {
                'pool_size': self._engine.pool.size(),
                'checked_out': self._engine.pool.checkedout(),
                'overflow': self._engine.pool.overflow(),
                'config': self._config
            }
        return None


# Singleton accessor function
def get_db_pool():
    """Get or create the global database pool"""
    return DatabasePool()


def get_db_connection():
    """Convenience function to get a connection from the pool"""
    return get_db_pool().get_connection()


def test_connection():
    """Test database connectivity"""
    try:
        pool = get_db_pool()
        conn = pool.get_connection()
        result = conn.execute(text("SELECT 1 as test"))
        conn.close()
        logger.info("✓ Database connection test passed")
        return True
    except Exception as e:
        logger.error(f"✗ Database connection test failed: {e}")
        return False

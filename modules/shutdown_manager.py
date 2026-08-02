"""
Graceful Shutdown Manager for Background Threads

Manages application shutdown, ensuring clean resource cleanup:
- Gracefully stops background threads
- Closes database connections
- Closes Selenium WebDriver instances
- Prevents orphaned processes
"""

import threading
import logging
import time
from typing import List, Optional

logger = logging.getLogger(__name__)


class ShutdownManager:
    """
    Centralized shutdown management for background threads and resources.
    
    Features:
    - Thread lifecycle tracking
    - Graceful thread termination with timeout
    - Resource cleanup coordination
    - Single point for application shutdown
    """
    
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        """Singleton pattern implementation"""
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super(ShutdownManager, cls).__new__(cls)
                    cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        """Initialize shutdown manager"""
        if self._initialized:
            return
        
        self.threads: List[threading.Thread] = []
        self.stop_event = threading.Event()
        self.drivers: List = []  # Selenium WebDriver instances
        self.cleanup_callbacks = []
        self._initialized = True
        
        logger.info("ShutdownManager initialized")
    
    # ========================================================================
    # Thread Management
    # ========================================================================
    
    def register_thread(self, thread: threading.Thread):
        """
        Register a background thread for graceful shutdown.
        
        Args:
            thread: The thread to track
        """
        if thread and not thread in self.threads:
            self.threads.append(thread)
            logger.debug(f"Registered thread: {thread.name}")
    
    def unregister_thread(self, thread: threading.Thread):
        """
        Unregister a thread (call when thread completes).
        
        Args:
            thread: The thread to remove from tracking
        """
        if thread in self.threads:
            self.threads.remove(thread)
            logger.debug(f"Unregistered thread: {thread.name}")
    
    def get_active_threads(self) -> List[threading.Thread]:
        """Get list of active threads"""
        return [t for t in self.threads if t.is_alive()]
    
    # ========================================================================
    # Resource Management
    # ========================================================================
    
    def register_driver(self, driver):
        """
        Register a Selenium WebDriver for cleanup.
        
        Args:
            driver: Selenium WebDriver instance
        """
        if driver:
            self.drivers.append(driver)
            logger.debug("Registered Selenium WebDriver")
    
    def register_cleanup_callback(self, callback):
        """
        Register a cleanup callback function.
        
        Args:
            callback: Callable to execute during shutdown
        """
        if callable(callback):
            self.cleanup_callbacks.append(callback)
            logger.debug(f"Registered cleanup callback: {callback.__name__}")
    
    # ========================================================================
    # Shutdown Sequence
    # ========================================================================
    
    def shutdown(self, timeout: int = 10):
        """
        Gracefully shutdown all resources and threads.
        
        Args:
            timeout: Maximum seconds to wait for thread termination
        """
        logger.info("=" * 70)
        logger.info("APPLICATION SHUTDOWN INITIATED")
        logger.info("=" * 70)
        
        # Signal all threads to stop
        logger.info("Setting stop event for background threads...")
        self.stop_event.set()
        
        # Wait for threads to terminate
        logger.info(f"Waiting for {len(self.get_active_threads())} active thread(s) to finish...")
        self._wait_for_threads(timeout)
        
        # Execute cleanup callbacks
        logger.info("Executing cleanup callbacks...")
        self._execute_cleanup_callbacks()
        
        # Close Selenium drivers
        logger.info("Closing Selenium WebDriver instances...")
        self._close_drivers()
        
        # Close database connections
        logger.info("Closing database connections...")
        self._close_database_connections()
        
        logger.info("=" * 70)
        logger.info("APPLICATION SHUTDOWN COMPLETE")
        logger.info("=" * 70)
    
    def _wait_for_threads(self, timeout: int):
        """
        Wait for all registered threads to complete.
        
        Args:
            timeout: Maximum seconds to wait per thread
        """
        active_threads = self.get_active_threads()
        
        if not active_threads:
            logger.info("No active threads to wait for")
            return
        
        for thread in active_threads:
            try:
                logger.debug(f"Waiting for thread: {thread.name}")
                thread.join(timeout=timeout)
                
                if thread.is_alive():
                    logger.warning(
                        f"Thread {thread.name} did not terminate within {timeout}s"
                    )
                else:
                    logger.info(f"Thread {thread.name} terminated successfully")
                    
            except Exception as e:
                logger.error(f"Error waiting for thread {thread.name}: {e}")
    
    def _execute_cleanup_callbacks(self):
        """Execute all registered cleanup callbacks"""
        for callback in self.cleanup_callbacks:
            try:
                logger.debug(f"Executing cleanup callback: {callback.__name__}")
                callback()
            except Exception as e:
                logger.error(f"Error in cleanup callback {callback.__name__}: {e}")
    
    def _close_drivers(self):
        """Close all registered Selenium WebDriver instances"""
        for driver in self.drivers:
            try:
                if driver:
                    logger.debug("Closing WebDriver...")
                    driver.quit()
                    logger.info("WebDriver closed successfully")
            except Exception as e:
                logger.error(f"Error closing WebDriver: {e}")
        
        self.drivers.clear()
    
    def _close_database_connections(self):
        """Close database connections"""
        try:
            from modules.db_controller import close_pool
            logger.debug("Closing database connection pool...")
            close_pool()
            logger.info("Database connection pool closed")
        except Exception as e:
            logger.error(f"Error closing database pool: {e}")
    
    # ========================================================================
    # Status & Monitoring
    # ========================================================================
    
    def get_status(self) -> dict:
        """
        Get shutdown manager status.
        
        Returns:
            dict: Status information
        """
        return {
            'stop_event_set': self.stop_event.is_set(),
            'active_threads': len(self.get_active_threads()),
            'total_threads': len(self.threads),
            'registered_drivers': len(self.drivers),
            'cleanup_callbacks': len(self.cleanup_callbacks),
        }
    
    def print_status(self):
        """Print shutdown manager status"""
        status = self.get_status()
        print("\nShutdown Manager Status:")
        print(f"  Stop event set: {status['stop_event_set']}")
        print(f"  Active threads: {status['active_threads']}/{status['total_threads']}")
        print(f"  Registered drivers: {status['registered_drivers']}")
        print(f"  Cleanup callbacks: {status['cleanup_callbacks']}")
        print()


# Singleton accessor
def get_shutdown_manager() -> ShutdownManager:
    """Get or create the global shutdown manager"""
    return ShutdownManager()

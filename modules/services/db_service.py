"""Database service - extracted from main_window.py"""
import logging
import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.exc import SQLAlchemyError
from dataclasses import dataclass
from datetime import datetime
import os

logger = logging.getLogger(__name__)

@dataclass
class DbResult:
    """Result of database operation"""
    success: bool
    message: str
    rows_uploaded: int = 0
    duplicate: bool = False
    error: str = None


def test_connection(host: str, port: str, user: str, password: str, database: str) -> tuple[bool, str, int]:
    """
    Test database connection.
    
    Args:
        host: Database host
        port: Database port
        user: Database user
        password: Database password
        database: Database name
    
    Returns:
        Tuple of (success, message, elapsed_ms)
    """
    import time
    
    try:
        start = time.time()
        
        url = f"postgresql://{user}:{password}@{host}:{port}/{database}"
        engine = create_engine(url, pool_pre_ping=True, connect_args={"connect_timeout": 5})
        
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        
        elapsed_ms = int((time.time() - start) * 1000)
        
        return (True, f"Connected in {elapsed_ms} ms", elapsed_ms)
        
    except Exception as e:
        error_msg = str(e)
        
        # Shorten common errors
        if "could not connect" in error_msg or "refused" in error_msg:
            short_msg = "Connection refused"
        elif "password authentication failed" in error_msg:
            short_msg = "Authentication failed"
        elif "does not exist" in error_msg:
            short_msg = "Database not found"
        elif "timeout" in error_msg.lower():
            short_msg = "Connection timeout"
        else:
            short_msg = error_msg[:100]
        
        return (False, short_msg, 0)


def upload_csv_to_db(
    csv_path: str,
    host: str,
    port: str,
    user: str,
    password: str,
    database: str,
    table: str
) -> DbResult:
    """
    Upload CSV file to PostgreSQL database.
    
    Args:
        csv_path: Path to CSV file
        host: Database host
        port: Database port
        user: Database user
        password: Database password
        database: Database name
        table: Table name
    
    Returns:
        DbResult with success status
    """
    try:
        if not os.path.isfile(csv_path):
            logger.error(f"File {csv_path} does not exist.")
            return DbResult(
                success=False,
                message=f"File not found: {csv_path}",
                error="FileNotFoundError"
            )
        
        # Create database connection
        url = f"postgresql+psycopg2://{user}:{password}@{host}:{port}/{database}"
        engine = create_engine(url)
        
        # Read CSV
        df = pd.read_csv(csv_path, delimiter=';')
        logger.info(f"Number of rows to upload: {len(df)}")
        
        # Upload to PostgreSQL
        try:
            df.to_sql(table, engine, if_exists='append', index=False)
            logger.info("Upload completed successfully")
            
            return DbResult(
                success=True,
                message=f"Uploaded {len(df)} rows to {table}",
                rows_uploaded=len(df)
            )
            
        except SQLAlchemyError as e:
            # Check for duplicate key error
            orig = getattr(e, 'orig', None)
            pgcode = getattr(orig, 'pgcode', None)
            
            if pgcode == '23505':
                logger.error("Upload failed: DUPLICATE DATA DETECTED")
                return DbResult(
                    success=False,
                    message="Duplicate data detected (unique constraint violation)",
                    duplicate=True,
                    error="23505"
                )
            else:
                logger.error(f"Upload error: {e}")
                return DbResult(
                    success=False,
                    message=f"Database error: {str(e)}",
                    error=str(e)
                )
                
    except Exception as e:
        logger.error(f"Error during upload: {str(e)}")
        return DbResult(
            success=False,
            message=f"Failed: {str(e)}",
            error=str(e)
        )


def upload_today_csv(
    output_dir: str,
    host: str,
    port: str,
    user: str,
    password: str,
    database: str,
    table: str
) -> DbResult:
    """
    Upload today's CSV file to database.
    
    Args:
        output_dir: Directory containing CSV files
        host: Database host
        port: Database port
        user: Database user
        password: Database password
        database: Database name
        table: Table name
    
    Returns:
        DbResult with success status
    """
    today = datetime.now().strftime('%Y-%m-%d')
    csv_file = os.path.join(output_dir, f'{today}.csv')
    
    return upload_csv_to_db(csv_file, host, port, user, password, database, table)

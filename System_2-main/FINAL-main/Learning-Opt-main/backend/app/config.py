# config.py
import os
import mysql.connector
from mysql.connector import pooling
from dotenv import load_dotenv
import logging

# Load environment variables
load_dotenv()

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Database configuration
db_config = {
    "host": os.getenv("DB_HOST", "localhost"),
    "port": os.getenv("DB_PORT", 3306),
    "user": os.getenv("DB_USER", "root"),
    "password": os.getenv("DB_PASSWORD", ""),
    "database": os.getenv("DB_NAME", "creo_certificate"),
    "autocommit": True
}

# Create connection pool
try:
    connection_pool = pooling.MySQLConnectionPool(
        pool_name="creo_pool",
        pool_size=10,
        pool_reset_session=True,
        **db_config
    )
    logger.info("✅ MySQL connection pool created successfully")
    
    # Test connection
    connection = connection_pool.get_connection()
    cursor = connection.cursor()
    cursor.execute("SELECT 1 + 1 AS solution")
    result = cursor.fetchone()
    logger.info(f"✅ Database test query result: {result[0]}")
    logger.info(f"✅ Connected to {db_config['host']}:{db_config['port']} ({db_config['database']})")
    cursor.close()
    connection.close()
    
except mysql.connector.Error as err:
    logger.error(f"❌ Database connection error: {err}")
    logger.info("Please verify:")
    logger.info("1. MySQL is running")
    logger.info('2. Database "creo_certificate" exists')
    logger.info('3. User "root" has no password set')
    exit(1)

# Database helper function
def execute_query(query, params=None):
    connection = connection_pool.get_connection()
    cursor = connection.cursor(dictionary=True)
    try:
        cursor.execute(query, params)
        if query.strip().lower().startswith("select"):
            return cursor.fetchall()
        return cursor.rowcount
    finally:
        cursor.close()
        connection.close()
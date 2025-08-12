# routes/auth.py
import sys
import os
from flask import Blueprint, request, jsonify
from flask_cors import CORS
import logging

current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir)

import config  # Relative import

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

auth_bp = Blueprint('auth', __name__)

# Enable CORS for the auth blueprint
CORS(auth_bp)

@auth_bp.route('/login', methods=['POST'])
def login():
    data = request.get_json()
    logger.info(f"🔒 Received login request: {data}")
    
    # Trim inputs
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()
    
    if not username or not password:
        logger.error("❌ Missing credentials")
        return jsonify({"message": "Username and password are required"}), 400
    
    try:
        # Query database
        query = "SELECT * FROM credential_tbl WHERE credential_username = %s"
        logger.info(f"🔍 Executing SQL: {query} [{username}]")
        
        users = config.execute_query(query, (username,))
        logger.info(f"📊 Found {len(users)} matching users")
        
        if not users:
            logger.error(f"❌ User not found: {username}")
            return jsonify({"message": "Invalid username or password"}), 401
        
        user = users[0]
        logger.info(f"👤 User record: ID={user['credential_id']}, Username={user['credential_username']}")
        
        # Compare passwords (plaintext)
        if password != user['credential_password']:
            logger.error(f"❌ Password mismatch for user: {username}")
            return jsonify({"message": "Invalid username or password"}), 401
        
        logger.info(f"✅ User logged in successfully: {username}")
        return jsonify({
            "user": {
                "id": user['credential_id'],
                "username": user['credential_username']
            }
        })
        
    except Exception as e:
        logger.error(f"❌ Login error: {str(e)}")
        return jsonify({"message": "Server error"}), 500
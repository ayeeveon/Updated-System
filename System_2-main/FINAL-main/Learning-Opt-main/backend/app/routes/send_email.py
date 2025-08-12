import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from flask import Flask, request, jsonify
import logging
from datetime import datetime

app = Flask(__name__)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@app.route('/send_email', methods=['POST'])
def send_email():
    try:
        data = request.get_json()
        
        # Required parameters
        to_email = 'kcangelbarbosa@gmail.com'
        subject = data.get('subject', 'No Subject')
        body = data.get('body')
        
        if not body:
            raise ValueError("Email body is required")

        # SMTP Configuration
        smtp_config = {
            'host': 'gator4119.hostgator.com',
            'port': 465,
            'username': 'creoappsadmin@creoapps.creoteconlinelearning.com',
            'password': 'xFkK_}t5pkA~'
        }

        # Create message
        msg = MIMEMultipart()
        msg['From'] = smtp_config['username']
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))

        # Send email
        with smtplib.SMTP_SSL(smtp_config['host'], smtp_config['port']) as server:
            server.login(smtp_config['username'], smtp_config['password'])
            server.send_message(msg)
        
        return jsonify({'success': True, 'message': 'Email sent successfully'})

    except Exception as e:
        logger.error(f"Error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)  # debug=False prevents auto-reloader issues
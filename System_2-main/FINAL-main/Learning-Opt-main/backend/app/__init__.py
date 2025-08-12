# backend/app/__init__.py
from flask import Flask
from flask_cors import CORS

def create_app():
    app = Flask(__name__)

    # Enable CORS
    CORS(app)

    @app.route('/')
    def index():
        return 'Backend server is running'

    # Import blueprints
    from .routes.upload import bp as upload_bp
    from .routes.generate import bp as generate_bp
    from .routes.immersion import immersion_bp

    # Register blueprints with URL prefixes
    app.register_blueprint(upload_bp, url_prefix='/upload')
    app.register_blueprint(generate_bp, url_prefix='/generate')
    app.register_blueprint(immersion_bp, url_prefix='/immersion')

    return app


# If you run this module directly, start Flask app with debug enabled
if __name__ == "__main__":
    app = create_app()
    app.run(debug=True, host='0.0.0.0', port=5000)

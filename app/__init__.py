# DQA/app/__init__.py

import os
from flask import Flask

def create_app():
    app = Flask(__name__)

    # Secret key needed for flash() / sessions
    app.config["SECRET_KEY"] = os.environ.get(
        "FLASK_SECRET_KEY",
        "dev-secret-key-change-me"  # change this in production
    )

    app.config["UPLOAD_FOLDER"] = "data/uploaded_files"

    from . import routes
    app.register_blueprint(routes.bp)

    return app
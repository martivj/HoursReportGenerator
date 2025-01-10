from flask import Flask
from pathlib import Path
from dotenv import load_dotenv
import os


def create_app():
    load_dotenv()

    app = Flask(__name__)

    app.config.update(
        SECRET_KEY=os.getenv("SECRET_KEY", "dev"),
        MAX_CONTENT_LENGTH=16 * 1024 * 1024,  # 16MB max file size
        UPLOAD_FOLDER=Path(os.getenv("UPLOAD_FOLDER", "temp/uploads")),
        MAX_FILES=int(os.getenv("MAX_FILES", 10)),
        ALLOWED_EXTENSIONS={"csv"},
    )

    # Ensure upload directory exists
    app.config["UPLOAD_FOLDER"].mkdir(parents=True, exist_ok=True)

    # Register blueprints
    from app.routes import main

    app.register_blueprint(main)

    return app

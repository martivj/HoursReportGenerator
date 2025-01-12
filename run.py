import logging
from logging.handlers import RotatingFileHandler
import os
from app import create_app
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_wtf.csrf import CSRFProtect
from pathlib import Path

app = create_app()


# Set up logging configuration
def configure_logging():
    Path("logs").mkdir(exist_ok=True)

    # Create formatter
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    # File handler with rotation
    file_handler = RotatingFileHandler(
        "logs/app.log",
        maxBytes=10 * 1024 * 1024,  # 10MB
        backupCount=10,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.INFO)

    # Stream handler for console output
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    console_handler.setLevel(logging.INFO)

    # Configure Flask app logger
    app.logger.setLevel(logging.INFO)
    app.logger.addHandler(file_handler)
    app.logger.addHandler(console_handler)

    # Remove default Flask handler
    app.logger.removeHandler(
        default_handler
        for default_handler in app.logger.handlers
        if isinstance(default_handler, logging.StreamHandler)
    )

    app.logger.info("Logging system initialized")


# Initialize logging if not in debug mode
if not app.debug:
    configure_logging()
    Path("temp/uploads").mkdir(parents=True, exist_ok=True)

csrf = CSRFProtect(app)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

limiter = Limiter(
    app=app, key_func=get_remote_address, default_limits=["200 per day", "50 per hour"]
)

if __name__ == "__main__":
    from hypercorn.config import Config
    from hypercorn.asyncio import serve
    import asyncio

    config = Config()
    config.bind = ["localhost:5000"]
    asyncio.run(serve(app, config))

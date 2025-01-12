import logging
from logging.handlers import RotatingFileHandler
import os
import sys
from app import create_app
from flask_wtf.csrf import CSRFProtect
from pathlib import Path
from flask_limiter.util import get_remote_address
from flask_limiter import Limiter
import redis

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
        mode="a",  # Open in text mode, not binary
    )
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.INFO)

    # Stream handler for console output
    console_handler = logging.StreamHandler(sys.stdout)  # Explicitly use stdout
    console_handler.setFormatter(formatter)
    console_handler.setLevel(logging.INFO)

    # Configure Flask app logger
    app.logger.setLevel(logging.INFO)
    app.logger.addHandler(file_handler)
    app.logger.addHandler(console_handler)

    # Remove default Flask handler
    for handler in app.logger.handlers[:]:
        if isinstance(handler, logging.StreamHandler):
            app.logger.removeHandler(handler)

    app.logger.info("Logging system initialized")


# Initialize logging if not in debug mode
if not app.debug:
    configure_logging()
    Path("temp/uploads").mkdir(parents=True, exist_ok=True)

csrf = CSRFProtect(app)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

redis_url = os.environ.get("REDIS_URL")

if redis_url:
    redis_client = redis.from_url(redis_url)
else:
    redis_client = redis.Redis(host="localhost", port=6379, db=0)

limiter = Limiter(
    app=app,
    key_func=get_remote_address,
    storage_uri=redis_url if redis_url else "redis://localhost:6379/0",
)

if __name__ == "__main__":
    from hypercorn.config import Config
    from hypercorn.asyncio import serve
    import asyncio
    import os

    # port = int(os.environ.get("PORT", 10000))
    # app.run(host="0.0.0.0", port=port)

    config = Config()
    port = int(os.environ.get("PORT", 10000))
    config.bind = [f"0.0.0.0:{port}"]
    asyncio.run(serve(app, config))

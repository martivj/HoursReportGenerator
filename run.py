import logging
from logging.handlers import RotatingFileHandler
import os
from app import create_app
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_wtf.csrf import CSRFProtect
from pathlib import Path

app = create_app()

# Set up logging and temp directories
if not app.debug:
    Path("logs").mkdir(exist_ok=True)
    Path("temp/uploads").mkdir(parents=True, exist_ok=True)
    file_handler = RotatingFileHandler(
        "logs/app.log",
        maxBytes=10240,
        backupCount=10,
    )
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s: %(message)s")
    )
    app.logger.addHandler(file_handler)

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

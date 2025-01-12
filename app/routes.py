import asyncio
from importlib import import_module
import inspect
import io
import os
from pathlib import Path
import shutil
import uuid
from flask import Blueprint, current_app, request, render_template, send_file, jsonify
from flask_wtf.csrf import generate_csrf
from werkzeug.utils import secure_filename

from src.project_configs.web_dev import WebDevConfig
from src.project_configs.itp2 import ITP2Config
from src.report_generator import ReportGenerator
from src.types.project_config import ProjectConfig

main = Blueprint("main", __name__)


def load_configs():
    """Dynamically load all ProjectConfig classes from project_configs folder"""
    configs = {}
    project_configs_path = Path(__file__).parent.parent / "src" / "project_configs"

    # Skip __init__.py and base classes
    for file_path in project_configs_path.glob("*.py"):
        if file_path.stem in ["__init__", "project_config"]:
            continue

        # Import the module dynamically
        module_name = f"src.project_configs.{file_path.stem}"
        module = import_module(module_name)

        # Find all classes that inherit from ProjectConfig
        for name, obj in inspect.getmembers(module):
            if (
                inspect.isclass(obj)
                and issubclass(obj, ProjectConfig)
                and obj != ProjectConfig
            ):

                # Create config instance to get display name
                config = obj()
                key = file_path.stem.lower()
                configs[key] = {"name": config.display_name, "class": obj}

    return configs


# Load configs once at startup
CONFIGS = load_configs()


def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower()
        in current_app.config["ALLOWED_EXTENSIONS"]
    )


@main.route("/")
def index():
    config_options = [{"key": k, "name": v["name"]} for k, v in CONFIGS.items()]
    return render_template(
        "index.html",
        configs=config_options,
        csrf_token=generate_csrf(),  # Use Flask-WTF's token generator
    )


@main.route("/health")
def health_check():
    """
    Basic health check endpoint that verifies:
    1. API is responsive
    2. Upload directory exists and is writable
    3. Can access configuration
    """
    try:
        # Log health check attempt
        current_app.logger.info("Health check initiated")

        upload_dir = current_app.config["UPLOAD_FOLDER"]
        if not upload_dir.exists() or not os.access(upload_dir, os.W_OK):
            current_app.logger.error(f"Upload directory not accessible: {upload_dir}")
            return (
                jsonify(
                    {"status": "error", "message": "Upload directory not accessible"}
                ),
                500,
            )

        if not CONFIGS:
            current_app.logger.error("Project configurations not loaded")
            return (
                jsonify(
                    {"status": "error", "message": "Project configurations not loaded"}
                ),
                500,
            )

        current_app.logger.info(
            f"Health check successful - Configs loaded: {len(CONFIGS)}, "
            f"Upload dir: {upload_dir}, Env: {current_app.config['FLASK_ENV']}"
        )
        return jsonify(
            {
                "status": "healthy",
                "configs_loaded": len(CONFIGS),
                "upload_dir": str(upload_dir),
                "environment": current_app.config["FLASK_ENV"],
            }
        )

    except Exception as e:
        current_app.logger.error(f"Health check failed: {str(e)}", exc_info=True)
        return jsonify({"status": "error", "message": str(e)}), 500


@main.route("/generate", methods=["POST"])
async def generate_report():
    # Get filename from form data, default to HoursReport
    output_filename = request.form.get("filename", "HoursReport")

    # Add .xlsx extension if not present
    if not output_filename.endswith(".xlsx"):
        output_filename += ".xlsx"

    # Validate config
    config_name = request.form.get("config")
    if config_name not in CONFIGS:
        return jsonify({"error": "Invalid config"}), 400

    # Validate files
    if "files" not in request.files:
        return jsonify({"error": "No files uploaded"}), 400

    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No selected files"}), 400

    if len(files) > current_app.config["MAX_FILES"]:
        return (
            jsonify(
                {"error": f"Maximum {current_app.config['MAX_FILES']} files allowed"}
            ),
            400,
        )

    # Create unique upload directory
    upload_dir = (
        Path(current_app.root_path).parent
        / current_app.config["UPLOAD_FOLDER"]
        / str(uuid.uuid4())
    )
    upload_dir.mkdir(parents=True, exist_ok=True)

    try:
        # Save uploaded files and generate report
        for file in files:
            if file and file.filename and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file.save(upload_dir / filename)
            else:
                return jsonify({"error": "Invalid file type"}), 400

        # Generate report
        config = CONFIGS[config_name]["class"]()
        generator = ReportGenerator(
            config,
            total_sheet_first=True,
            close_open_excel=False,
            data_dir=upload_dir,
        )

        output_path = upload_dir / output_filename
        generator.generate(str(output_path))

        if output_path.exists():
            memory_file = io.BytesIO()

            max_retries = 5
            retry_count = 0

            while retry_count < max_retries:
                try:
                    # Try to read and delete file
                    with open(output_path, "rb") as excel_file:
                        memory_file.write(excel_file.read())
                    memory_file.seek(0)

                    # Wait a moment before cleanup
                    await asyncio.sleep(0.5)

                    # Try cleanup
                    if upload_dir.exists():
                        shutil.rmtree(upload_dir)
                    break

                except PermissionError:
                    retry_count += 1
                    if retry_count == max_retries:
                        raise
                    await asyncio.sleep(1)  # Wait before retry

            return send_file(
                memory_file,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name=output_filename,
            )
        else:
            return jsonify({"error": "Failed to generate report"}), 500

    except Exception as e:
        current_app.logger.error(f"Error generating report: {str(e)}")
        return jsonify({"error": str(e)}), 500

    finally:
        # Schedule cleanup for later if directory still exists
        if upload_dir.exists():

            async def delayed_cleanup():
                await asyncio.sleep(2)  # Wait 2 seconds
                try:
                    if upload_dir.exists():
                        shutil.rmtree(upload_dir)
                except Exception as e:
                    current_app.logger.error(f"Error in delayed cleanup: {str(e)}")

            asyncio.create_task(delayed_cleanup())

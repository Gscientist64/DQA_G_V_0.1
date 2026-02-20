# DQA/run.py

from app import create_app
import os

# Create the app instance
app = create_app()

# Start the app
if __name__ == "__main__":
    os.environ["FLASK_ENV"] = "development"
    app.run(host="0.0.0.0", port=8000, debug=True)

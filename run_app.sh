#!/bin/bash
# Activating the virtual environment
source /Users/tayasarbhat/excel_compare/venv/bin/activate
# Setting environment variables
export FLASK_APP=app.py
export FLASK_ENV=development
# Starting the Flask application
flask run

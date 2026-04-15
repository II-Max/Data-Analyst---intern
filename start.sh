#!/bin/bash
python3 -m venv venv
source venv/bin/activate
pip install flask flask-cors pandas openpyxl python-dotenv Werkzeug
python3 app.py

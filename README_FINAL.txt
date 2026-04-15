Excel Data Entry Application

QUICK START
===========

Windows:
- Double-click: start.bat

Mac/Linux:
- chmod +x start.sh && ./start.sh

Then open browser: http://localhost:5000


FEATURES
========
+ Input/edit data on web
+ Upload Excel files
+ Export Excel files
+ Real-time validation
+ Data statistics


HOW TO USE
==========
1. Add row - Click + button
2. Upload file - Click upload button
3. Edit data - Click cell, type, auto-save
4. Delete - Click trash icon
5. Export - Click export button


FILES
=====
app.py             - Main application
config.py          - Configuration
start.bat          - Windows start script
start.sh           - Mac/Linux start script
templates/index.html - Web interface


REQUIREMENTS
============
Python 3.8+
512MB RAM
100MB disk


TROUBLESHOOTING
===============
Python not found?
- Install from python.org

Module error?
- Run: pip install flask flask-cors pandas openpyxl python-dotenv Werkzeug

Port error?
- Edit app.py, change port number

Upload error?
- Use .xlsx format (not .xls)


DATA LOCATION
=============
project/uploads/
Format: data_YYYYMMDD_HHMMSS.xlsx


SHARING
=======
Option 1: Export file, email/send to team
Option 2: Run on server, team access http://[IP]:5000


VERSION
=======
v1.0 - MIT License

Ready to go!

@echo off
echo Installing dependencies...
pip install -r requirements.txt
echo.
echo Starting GloBE XML Export App...
echo Open http://localhost:8501 in your browser.
echo.
streamlit run globe_xml_app.py
pause

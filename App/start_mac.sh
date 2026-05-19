#!/bin/bash
echo "Installing dependencies..."
pip3 install -r requirements.txt
echo ""
echo "Starting GloBE XML Export App..."
echo "Open http://localhost:8501 in your browser."
echo ""
streamlit run globe_xml_app.py

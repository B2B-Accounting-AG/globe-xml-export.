import runpy, os
# Deployment entry point (Streamlit Cloud). Runs the v2 app (new multi-sheet GIR
# template: multi-jurisdiction + safe harbours). The v1 app remains at
# App/globe_xml_app.py as a fallback for the old single-sheet template.
runpy.run_path(os.path.join(os.path.dirname(__file__), "App", "globe_xml_app_v2.py"), run_name="__main__")

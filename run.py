import subprocess
import threading
import time
import webbrowser

def open_browser():
    time.sleep(2)
    webbrowser.open("http://localhost:8501")

threading.Thread(
    target=open_browser,
    daemon=True
).start()

subprocess.run([
    "streamlit",
    "run",
    "app.py"
])
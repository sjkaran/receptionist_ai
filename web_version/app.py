"""Entry point:  python app.py   ->  opens the control panel in your browser."""
import threading
import webbrowser

import uvicorn

from web_version.core import database as db

HOST, PORT = "127.0.0.1", 8000

if __name__ == "__main__":
    db.init_db()
    threading.Timer(1.2, lambda: webbrowser.open(f"http://{HOST}:{PORT}")).start()
    uvicorn.run("web.server:app", host=HOST, port=PORT, log_level="warning")

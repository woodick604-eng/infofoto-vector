import threading
import time
import webbrowser
from app import app

def open_browser():
    # Donem un petit marge perquè el servidor arrenqui
    time.sleep(2)
    webbrowser.open("http://127.0.0.1:5051/")

# Executem l'obertura del navegador en un fil separat
thread = threading.Thread(target=open_browser, daemon=True)
thread.start()

print("Servidor iniciat. Obring el navegador...")
app.run(host="127.0.0.1", port=5051, debug=False)

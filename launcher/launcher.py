import pystray
from PIL import Image, ImageDraw
import subprocess
import os
import threading
import sys

def create_image():
    # Create a simple tray icon
    image = Image.new('RGB', (64, 64), color=(0, 120, 212))
    d = ImageDraw.Draw(image)
    d.text((10, 20), "OAI", fill=(255, 255, 255))
    return image

server_process = None

def start_server(icon, item):
    global server_process
    if server_process is None:
        backend_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'backend'))
        
        cert_dir = os.path.expanduser("~/.office-addin-dev-certs")
        ssl_keyfile = os.path.join(cert_dir, "localhost.key")
        ssl_certfile = os.path.join(cert_dir, "localhost.crt")
        
        cmd = [sys.executable, "-m", "uvicorn", "main:app", "--host", "127.0.0.1", "--port", "8000"]
        if os.path.exists(ssl_keyfile) and os.path.exists(ssl_certfile):
            cmd.extend(["--ssl-keyfile", ssl_keyfile, "--ssl-certfile", ssl_certfile])
            
        server_process = subprocess.Popen(cmd, cwd=backend_dir)
        print("Backend Server Started")

def stop_server(icon, item):
    global server_process
    if server_process is not None:
        server_process.terminate()
        server_process = None
        print("Backend Server Stopped")

def exit_app(icon, item):
    stop_server(icon, item)
    icon.stop()

def setup_tray():
    image = create_image()
    menu = pystray.Menu(
        pystray.MenuItem('Start Server', start_server),
        pystray.MenuItem('Stop Server', stop_server),
        pystray.MenuItem('Exit', exit_app)
    )
    icon = pystray.Icon("OfficeAI", image, "OfficeAI Pro", menu)
    icon.run()

if __name__ == "__main__":
    setup_tray()

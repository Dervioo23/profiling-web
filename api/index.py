import os
import sys

# Tambahkan root directory ke path agar processor dan modul lain terbaca
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app

# Export app untuk WSGI handler Vercel
app = app

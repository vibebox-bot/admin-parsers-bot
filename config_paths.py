import os

BASE_DIR = os.getenv("BASE_DIR", "/app")

def path(*parts):
    return os.path.join(BASE_DIR, *parts)

import os

BASE_DIR = "/app"

def path(*parts):
    return os.path.join(BASE_DIR, *parts)

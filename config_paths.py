import os

BASE_DIR = "/data"

def path(p):
    return os.path.join(BASE_DIR, p)

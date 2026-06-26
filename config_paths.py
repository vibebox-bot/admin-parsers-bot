import os

BASE_DIR = os.getcwd()

def path(*parts):
    return os.path.join(BASE_DIR, *parts)

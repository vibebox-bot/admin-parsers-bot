FROM python:3.12-bookworm

WORKDIR /app

COPY requirements.txt .

RUN echo "=== 1. INSTALL PYTHON ===" && \
    pip install --no-cache-dir -r requirements.txt

RUN echo "=== 2. INSTALL FIREFOX + XVFB ===" && \
    apt-get update && \
    apt-get install -y xvfb xauth && \
    rm -rf /var/lib/apt/lists/* && \
    python -m playwright install --with-deps firefox

COPY . .

RUN echo "=== 3. FILES ===" && \
    ls -la && \
    echo "=== BOT.PY ===" && \
    ls -l bot.py

RUN echo "=== 4. PYTHON ===" && \
    python --version && \
    python -c "import aiogram; print('aiogram OK')" && \
    python -c "import playwright; print('playwright OK')" && \
    python -c "import psutil; print('psutil OK')" && \
    python -c "import pytz; print('pytz OK')"

CMD ["bash", "-c", "Xvfb :99 -screen 0 1440x900x24 >/tmp/xvfb.log 2>&1 & export DISPLAY=:99; python -u bot.py"]

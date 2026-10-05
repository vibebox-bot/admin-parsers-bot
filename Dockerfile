FROM python:3.12-bookworm

WORKDIR /app

COPY requirements.txt .

RUN echo "=== 1. INSTALL PYTHON ===" && \
    pip install --no-cache-dir -r requirements.txt

RUN echo "=== 2. INSTALL FIREFOX + XVFB + VNC ===" && \
    apt-get update && \
    apt-get install -y xvfb xauth x11vnc novnc && \
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

EXPOSE 8080

CMD ["bash", "-c", "echo '=== START XVFB ==='; Xvfb :99 -screen 0 1440x900x24 >/tmp/xvfb.log 2>&1 & sleep 2; echo '=== START X11VNC ==='; x11vnc -display :99 -forever -shared -localhost -rfbport 5900 -nopw >/tmp/x11vnc.log 2>&1 & sleep 2; echo '=== START NOVNC ==='; /usr/share/novnc/utils/novnc_proxy --vnc localhost:5900 --listen 8080 >/tmp/novnc.log 2>&1 & sleep 3; echo '=== NOVNC LOG ==='; cat /tmp/novnc.log; echo '=== PORT 8080 ==='; (curl -s http://127.0.0.1:8080/ | head -c 200 || true); echo; echo '=== START BOT ==='; export DISPLAY=:99; exec python -u bot.py"]



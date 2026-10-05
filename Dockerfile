FROM python:3.12-bookworm

WORKDIR /app

COPY requirements.txt .

RUN echo "=== 1. INSTALL PYTHON ===" && \
    pip install --no-cache-dir -r requirements.txt

RUN echo "=== 2. INSTALL FIREFOX ===" && \
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

CMD ["python", "-u", "bot.py"]

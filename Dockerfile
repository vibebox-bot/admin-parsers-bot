FROM python:3.12-bookworm

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

RUN python -m playwright install --with-deps firefox

RUN apt-get update \
    && apt-get install -y xvfb \
    && rm -rf /var/lib/apt/lists/*

COPY . .

CMD ["xvfb-run", "-a", "python", "bot.py"]

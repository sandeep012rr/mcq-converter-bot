FROM python:3.10-slim
RUN apt-get update && apt-get install -y \
    build-essential libpango-1.0-0 libharfbuzz0b libpangoft2-1.0-0 \
    libffi-dev libjpeg-dev libopenjp2-7 fonts-dejavu fonts-freefont-ttf \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY . /app
RUN pip install --no-cache-dir -r requirements.txt
CMD ["python", "bot.py"]

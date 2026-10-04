FROM python:3.10-slim

# wkhtmltopdf aur Devanagari fonts install karein
RUN apt-get update && apt-get install -y \
    wkhtmltopdf \
    fonts-noto-core \
    fonts-noto-extra \
    fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . /app

RUN pip install --no-cache-dir -r requirements.txt

EXPOSE 10000

CMD ["python", "bot.py"]

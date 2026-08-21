FROM python:3.11-slim

# Install Google Chrome Stable from Google's official repo
# This is far more reliable than Debian's chromium package
RUN apt-get update && apt-get install -y \
    wget curl gnupg unzip poppler-utils \
    --no-install-recommends && \
    wget -q -O /tmp/chrome.deb https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb && \
    apt-get install -y /tmp/chrome.deb --fix-broken && \
    rm /tmp/chrome.deb && \
    rm -rf /var/lib/apt/lists/* && \
    echo "Chrome: $(google-chrome --version)" && \
    echo "ChromeDriver: $(chromedriver --version 2>/dev/null || echo 'not found')" && \
    which google-chrome && \
    which chromedriver || find / -name chromedriver 2>/dev/null | head -5

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

EXPOSE 8501

CMD ["streamlit", "run", "app.py", \
     "--server.port=8501", \
     "--server.address=0.0.0.0", \
     "--server.headless=true"]

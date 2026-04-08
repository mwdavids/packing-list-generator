FROM python:3.11-slim

WORKDIR /app

# Install LiteStream
ADD https://github.com/benbjohnson/litestream/releases/download/v0.3.13/litestream-v0.3.13-linux-amd64.tar.gz /tmp/litestream.tar.gz
RUN tar -xzf /tmp/litestream.tar.gz -C /usr/local/bin/ && rm /tmp/litestream.tar.gz

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

ARG CACHEBUST=1
COPY . .
RUN chmod +x start.sh

EXPOSE 8000

CMD ["./start.sh"]

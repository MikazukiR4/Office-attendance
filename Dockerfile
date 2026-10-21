FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    WORDCRAFT_VERSION=0.8.0 \
    WORDCRAFT_CLI=/usr/local/bin/wordcraft-cli

RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates curl fontconfig fonts-noto-core tar \
    && rm -rf /var/lib/apt/lists/* \
    && curl -fsSL -o /tmp/wordcraft.tar.gz \
       "https://github.com/storytold/wordcraft/releases/download/v${WORDCRAFT_VERSION}/wordcraft-${WORDCRAFT_VERSION}-linux-x86_64.tar.gz" \
    && mkdir /tmp/wordcraft \
    && tar -xzf /tmp/wordcraft.tar.gz -C /tmp/wordcraft \
    && find /tmp/wordcraft -type f -name wordcraft-cli -exec cp {} /usr/local/bin/wordcraft-cli \; \
    && chmod +x /usr/local/bin/wordcraft-cli \
    && test -x /usr/local/bin/wordcraft-cli \
    && rm -rf /tmp/wordcraft /tmp/wordcraft.tar.gz

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN mkdir -p /usr/local/share/fonts/th-sarabun \
    && cp fonts/*.ttf /usr/local/share/fonts/th-sarabun/ \
    && fc-cache -f

EXPOSE 10000
CMD ["sh", "-c", "streamlit run app.py --server.address=0.0.0.0 --server.port=${PORT:-10000}"]

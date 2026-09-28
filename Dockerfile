FROM python:3.12-slim

# Small image, no build tools needed - the script only uses requests and openpyxl.
RUN pip install --no-cache-dir requests openpyxl

WORKDIR /app
COPY streamer_finder_v3.py /app/streamer_finder_v3.py

# Data lives in a mounted volume so the database survives between runs.
VOLUME ["/data"]
WORKDIR /data

# Credentials come from the environment, never baked into the image:
#   TWITCH_CLIENT_ID, TWITCH_CLIENT_SECRET   (required)
#   ALERT_WEBHOOK_URL                        (optional Slack/Discord webhook)
#   KEEP_DAYS                                (optional, default 30)
ENTRYPOINT ["python", "/app/streamer_finder_v3.py"]
CMD ["ci"]

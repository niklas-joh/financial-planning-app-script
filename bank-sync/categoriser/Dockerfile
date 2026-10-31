FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app
COPY pyproject.toml ./
COPY src ./src
RUN pip install . && useradd --create-home --uid 1000 app

# credentials/, data/ and models/ are mounted at runtime (see docker-compose.yml).
USER app
ENV GOOGLE_APPLICATION_CREDENTIALS=/app/credentials/service-account.json
CMD ["python", "-m", "categoriser.run"]

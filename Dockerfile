FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

LABEL project="tg_dashi"

WORKDIR /app

COPY requirements.txt requirements-dev.txt ./
ARG DEV=0
RUN pip install --upgrade pip \
    && if [ "$DEV" = "1" ]; then pip install -r requirements-dev.txt; else pip install -r requirements.txt; fi

COPY . .

CMD ["python", "-m", "app.bot"]

FROM python:3.11

RUN apt update && apt -y install nodejs

WORKDIR /app

COPY . /app

RUN pip install --no-cache-dir --user --no-warn-script-location -r requirements.txt

WORKDIR /data

ENTRYPOINT ["python3", "/app/md2docx.py"]

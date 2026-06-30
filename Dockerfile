FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV HOST=0.0.0.0
ENV XIAJIU_DATA_DIR=/tmp/xiajiu-weekly-panel

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY fill_report_dashboard.py ./
COPY work/build_xiajiu_report.py ./work/build_xiajiu_report.py

RUN mkdir -p /tmp/xiajiu-weekly-panel

CMD ["python", "fill_report_dashboard.py"]

FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

RUN useradd --create-home --uid 10001 appuser
COPY --chown=appuser:appuser pv_blink_dashboard.py PV_Blink_Login_Form.py ./
COPY --chown=appuser:appuser assets/pvblink_logo.png ./assets/pvblink_logo.png
USER appuser

EXPOSE 8501
HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
    CMD ["python", "-c", "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8501/_stcore/health', timeout=3).read()"]

ENTRYPOINT ["python", "-m", "streamlit", "run", "pv_blink_dashboard.py", "--server.address=0.0.0.0", "--server.port=8501", "--server.headless=true"]

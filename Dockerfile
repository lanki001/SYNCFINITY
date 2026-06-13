FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY syncfinity_2_0.py .

CMD ["python", "syncfinity_2_0.py"]
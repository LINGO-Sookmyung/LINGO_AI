FROM python:3.11

WORKDIR /lingo2025-ai

COPY ./requirements.txt /lingo2025-ai/requirements.txt
RUN pip install --no-cache-dir --upgrade -r /lingo2025-ai/requirements.txt

COPY . /lingo2025-ai

WORKDIR /lingo2025-ai

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
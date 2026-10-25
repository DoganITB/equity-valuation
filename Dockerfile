# Web service (Render.com, Hugging Face, any Docker host): serves the Equity Valuation Platform and its data service on one public URL.
FROM python:3.11-slim
RUN apt-get update && apt-get install -y --no-install-recommends git ca-certificates && rm -rf /var/lib/apt/lists/*
RUN useradd -m -u 1000 user
USER user
ENV HOME=/home/user PATH=/home/user/.local/bin:$PATH PYTHONUNBUFFERED=1
WORKDIR /home/user/app
COPY --chown=user requirements.txt .
RUN pip install --no-cache-dir --user -r requirements.txt
# Optional IDX filings source (MIT licence, by celeebi). The site still runs on Yahoo if the clone fails.
RUN git clone --depth 1 https://github.com/celeebi/IDX_GetFinancialStatement || echo "IDX package not cloned"
COPY --chown=user . .
EXPOSE 7860
# Hosts like Render pass the port in $PORT; 7860 otherwise.
CMD ["sh", "-c", "python data_connector.py --host 0.0.0.0 --port ${PORT:-7860}"]

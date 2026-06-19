# Image du serveur MCP Outlook en transport HTTP (Azure Container Apps).
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    OUTLOOK_MCP_TRANSPORT=http \
    PORT=8000

WORKDIR /app

# Dépendances d'abord (cache de build)
COPY pyproject.toml README.md ./
COPY src ./src
RUN pip install --no-cache-dir .

EXPOSE 8000

# Sonde simple + démarrage
CMD ["python", "-m", "outlook_graph_mcp", "serve-http"]

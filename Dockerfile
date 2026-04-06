# Use a slim Python base image
FROM python:3.11-slim

# Prevent Python from writing pyc files & enable stdout flushing
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Install system dependencies (adjust if needed)
RUN apt-get update && apt-get install -y \
    curl \
    build-essential \
    git \
    && rm -rf /var/lib/apt/lists/*

# Install uv (fast Python package manager)
RUN curl -Ls https://astral.sh/uv/install.sh | sh

# Add uv to PATH
ENV PATH="/root/.local/bin:$PATH"

# Set working directory
WORKDIR /app

# Copy project files
COPY . /app

# (Optional but recommended) create virtual environment
RUN uv venv /app/.venv

# Install dependencies (auto-detects pyproject.toml / requirements.txt)
RUN uv pip install --python /app/.venv/bin/python .

# Expose port if MCP runs a server (adjust as needed)
EXPOSE 8000

# Default command using uvx (ephemeral runner)
CMD ["uvx", "mcpo", "--port", "8000", "--", "uvx", "--python", "/app/.venv/bin/python", "-m", "mcp_excel.main"]

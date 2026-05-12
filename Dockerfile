FROM python:3.13-slim

WORKDIR /app

# Install uv
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

# Copy project files
COPY pyproject.toml uv.lock ./
COPY main.py scraper.py converter.py ./

# Install dependencies
RUN uv sync --frozen

# Create directories for temp files
RUN mkdir -p /tmp/makalahmcp/images

# Expose port
EXPOSE 8000

# Run MCP server in HTTP mode
CMD ["uv", "run", "main.py", "--http"]

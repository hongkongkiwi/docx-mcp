# Multi-stage Docker build for docx-mcp
FROM rust:1.75-slim as builder

# Install system dependencies for building
RUN apt-get update && apt-get install -y \
    pkg-config \
    libssl-dev \
    libfontconfig1-dev \
    libfreetype6-dev \
    libjpeg-dev \
    libpng-dev \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy manifests
COPY Cargo.toml Cargo.lock ./
COPY build.rs ./

# Copy source code
COPY src/ ./src/
COPY benches/ ./benches/
COPY tests/ ./tests/

# Build the application
RUN cargo build --release --all-features

# Runtime stage
FROM debian:bookworm-slim

# Install runtime dependencies
RUN apt-get update && apt-get install -y \
    libssl3 \
    libfontconfig1 \
    libfreetype6 \
    libjpeg62-turbo \
    libpng16-16 \
    ca-certificates \
    libreoffice \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# Create non-root user
RUN groupadd -r docxmcp && useradd -r -g docxmcp -s /bin/bash -d /app docxmcp

# Create app directory and set ownership
WORKDIR /app
RUN chown -R docxmcp:docxmcp /app

# Copy the built binary from builder stage
COPY --from=builder /app/target/release/docx-mcp /usr/local/bin/docx-mcp
RUN chmod +x /usr/local/bin/docx-mcp

# Copy additional files if needed
COPY README.md LICENSE ./

# Switch to non-root user
USER docxmcp

# Create temp directory for document processing
RUN mkdir -p /tmp/docx-mcp && chmod 755 /tmp/docx-mcp

# Expose default MCP port (though MCP typically uses stdin/stdout)
EXPOSE 8080

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD /usr/local/bin/docx-mcp --version || exit 1

# Set environment variables
ENV RUST_LOG=info
ENV DOCX_MCP_TEMP_DIR=/tmp/docx-mcp

# Default command
CMD ["/usr/local/bin/docx-mcp"]
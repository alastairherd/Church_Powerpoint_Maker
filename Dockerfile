FROM rust:1-slim AS builder
WORKDIR /app
COPY . .
RUN cargo build --release -p server

FROM debian:stable-slim AS runtime
RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates \
    && rm -rf /var/lib/apt/lists/*
COPY --from=builder /app/target/release/server /usr/local/bin/church-deck-server
ENV PORT=8080
EXPOSE 8080
CMD ["/usr/local/bin/church-deck-server"]

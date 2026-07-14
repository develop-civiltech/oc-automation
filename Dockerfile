# syntax=docker/dockerfile:1

# ---- deps: instala node_modules (con toolchain de compilación como fallback para better-sqlite3) ----
FROM node:20-slim AS deps
RUN apt-get update \
    && apt-get install -y --no-install-recommends python3 make g++ \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

# ---- runtime: imagen final compartida por los servicios "app" y "mailer" ----
FROM node:20-slim AS runtime

# supercronic (usado solo por el servicio "mailer"; inofensivo en el servicio "app")
ENV SUPERCRONIC_URL=https://github.com/aptible/supercronic/releases/download/v0.2.47/supercronic-linux-amd64 \
    SUPERCRONIC_SHA1SUM=712d2ece75da6f6e530192a151488578153e4e96 \
    SUPERCRONIC=supercronic-linux-amd64

RUN apt-get update \
    && apt-get install -y --no-install-recommends curl ca-certificates \
    && rm -rf /var/lib/apt/lists/* \
    && curl -fsSLO "$SUPERCRONIC_URL" \
    && echo "${SUPERCRONIC_SHA1SUM}  ${SUPERCRONIC}" | sha1sum -c - \
    && chmod +x "$SUPERCRONIC" \
    && mv "$SUPERCRONIC" "/usr/local/bin/${SUPERCRONIC}" \
    && ln -s "/usr/local/bin/${SUPERCRONIC}" /usr/local/bin/supercronic

RUN groupadd --system --gid 10001 app \
    && useradd --system --uid 10001 --gid 10001 --home-dir /app --no-create-home app

WORKDIR /app
COPY --from=deps /app/node_modules ./node_modules
COPY . .

RUN chown -R app:app /app

USER app
ENV NODE_ENV=production
EXPOSE 3001

CMD ["node", "src/servidor-cotizaciones.js"]

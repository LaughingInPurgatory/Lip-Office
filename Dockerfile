FROM node:24-alpine

WORKDIR /app

COPY --chown=node:node package.json server.js ./
COPY --chown=node:node lib ./lib
COPY --chown=node:node public ./public
COPY --chown=node:node shared ./shared

RUN mkdir -p /app/documents/Docs /app/documents/Spreadsheets \
    && chown -R node:node /app/documents

ENV HOST=0.0.0.0 \
    PORT=3000 \
    LIP_OFFICE_CONTAINER=1

USER node
EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:3000/').then(r => process.exit(r.ok ? 0 : 1)).catch(() => process.exit(1))"

CMD ["node", "server.js"]

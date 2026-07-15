FROM node:20-alpine AS build

WORKDIR /app

COPY mcp-server/package*.json ./mcp-server/
RUN cd mcp-server && npm ci

COPY mcp-server/tsconfig.json ./mcp-server/
COPY mcp-server/src ./mcp-server/src
COPY mcp-server/scripts ./mcp-server/scripts
COPY src/data/serviceIconMapping.ts src/data/fabricIconCatalog.ts ./src/data/
COPY src/data/pricing/regions ./src/data/pricing/regions
COPY Azure_Public_Service_Icons ./Azure_Public_Service_Icons

RUN cd mcp-server && npm run build

FROM node:20-alpine

ENV NODE_ENV=production \
    MCP_TRANSPORT=http \
    MCP_HTTP_HOST=127.0.0.1 \
    MCP_HTTP_PORT=3030 \
    MCP_HTTP_PATH=/mcp

WORKDIR /app

COPY mcp-server/package*.json ./
RUN npm ci --omit=dev && npm cache clean --force
COPY --from=build /app/mcp-server/dist ./dist

USER node
EXPOSE 3030

CMD ["node", "dist/index.js", "--http"]

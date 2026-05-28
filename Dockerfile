# syntax=docker/dockerfile:1

# ── Stage 1: Build + SSR prerender ──────────────────────────────────────────
FROM node:22-alpine AS build
WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci

COPY . .

# Three-step build (defined in package.json):
#   1. vite build              → dist/         (client bundle)
#   2. vite build --ssr ...    → dist/server/  (Node SSR bundle)
#   3. node scripts/prerender  → patches dist/index.html with rendered HTML
RUN npm run build

# Remove the SSR server bundle — it's only needed during build, not at runtime
RUN rm -rf /app/dist/server

# ── Stage 2: Serve with nginx ────────────────────────────────────────────────
FROM nginx:1.27-alpine
COPY --from=build /app/dist /usr/share/nginx/html
COPY nginx.conf /etc/nginx/conf.d/default.conf

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]

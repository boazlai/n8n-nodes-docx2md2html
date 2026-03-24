# ── Stage 1: Build ────────────────────────────────────────────────────────────
FROM node:20-alpine AS builder

WORKDIR /build

# Install all deps (including devDeps for the build toolchain)
# --ignore-scripts skips native addon compilation (isolated-vm etc.) — not needed at compile time
COPY package.json ./
RUN npm install --ignore-scripts

# Compile TypeScript → dist/
COPY tsconfig.json ./
COPY nodes/ ./nodes/
RUN npm run build

# ── Stage 2: n8n with the custom node installed ───────────────────────────────
FROM n8nio/n8n:latest

USER root

# Create the custom-extensions directory and copy the built package
WORKDIR /custom-nodes/n8n-nodes-docx2md

COPY --from=builder /build/dist ./dist
COPY --from=builder /build/package.json ./package.json

# Install only production dependencies (mammoth, marked, html-to-docx)
# --legacy-peer-deps avoids trying to install n8n-workflow again (provided by n8n)
RUN npm install --omit=dev --legacy-peer-deps --ignore-scripts

# Switch back to the unprivileged n8n user
USER node

# Tell n8n where to find custom node packages
ENV N8N_CUSTOM_EXTENSIONS=/custom-nodes

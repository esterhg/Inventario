# ── Imagen para Control de Activos OCC ──────────────────────────────────────
FROM node:20-alpine

# Directorio de trabajo dentro del contenedor
WORKDIR /app

# Instalar dependencias primero (aprovecha la cache de Docker si el código
# cambia pero no las dependencias)
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

# Copiar el resto del código de la aplicación
COPY . .

# Carpeta donde se guardan las imágenes subidas.
# En Coolify, monta un volumen persistente en esta ruta para que las
# imágenes no se pierdan al re-desplegar el contenedor.
RUN mkdir -p /app/uploads

# Puerto en el que escucha la app (coincide con PORT / valor por defecto 4000)
EXPOSE 4000

# Usuario sin privilegios (más seguro que correr como root)
RUN chown -R node:node /app
USER node

ENV NODE_ENV=production
ENV PORT=4000

CMD ["node", "server.js"]

# Usar una imagen base que ya contenga Node.js y npm
FROM node:latest

# Instalar CLASP globalmente
RUN npm install -g @google/clasp

# Establecer el directorio de trabajo predeterminado dentro del contenedor
WORKDIR /app

# Comando para iniciar el contenedor
CMD ["/bin/bash"]
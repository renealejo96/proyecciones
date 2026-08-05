#!/bin/bash
echo "=== Desplegando actualización de Proyecciones en VPS ==="

# Crear .env si no existe a partir de .env.example
if [ ! -f .env ]; then
    echo "Creando archivo .env desde .env.example..."
    cp .env.example .env
fi

# Descargar últimos cambios de GitHub
git pull origin main

# Reconstruir y reiniciar contenedores Docker en producción
docker-compose up -d --build

echo "=== Despliegue completado exitosamente ==="

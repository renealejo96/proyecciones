#!/bin/bash
echo "=== Desplegando actualización de Proyecciones en VPS ==="
git pull origin main
docker-compose up -d --build
echo "=== Despliegue completado exitosamente ==="

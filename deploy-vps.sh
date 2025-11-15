#!/bin/bash

# =============================================================================
# SCRIPT DE DESPLIEGUE PARA VPS
# =============================================================================
# Este script automatiza el despliegue de Report Tuner en tu VPS
# =============================================================================

set -e  # Salir si hay algún error

# Colores para output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # Sin color

# Configuración
IMAGE_NAME="gimzalo/empower-reports:latest"
CONTAINER_NAME="empower-reports-app"
COMPOSE_FILE="docker-compose.prod.yml"

echo -e "${BLUE}================================================${NC}"
echo -e "${BLUE}   DESPLIEGUE DE REPORT TUNER EN VPS${NC}"
echo -e "${BLUE}================================================${NC}"
echo ""

# Función para mostrar mensajes
log_info() {
    echo -e "${BLUE}ℹ${NC} $1"
}

log_success() {
    echo -e "${GREEN}✓${NC} $1"
}

log_warning() {
    echo -e "${YELLOW}⚠${NC} $1"
}

log_error() {
    echo -e "${RED}✗${NC} $1"
}

# Verificar que Docker está instalado
log_info "Verificando instalación de Docker..."
if ! command -v docker &> /dev/null; then
    log_error "Docker no está instalado. Por favor instala Docker primero."
    echo "Visita: https://docs.docker.com/engine/install/"
    exit 1
fi
log_success "Docker está instalado"

# Verificar que Docker Compose está instalado
log_info "Verificando instalación de Docker Compose..."
if ! command -v docker-compose &> /dev/null && ! docker compose version &> /dev/null; then
    log_error "Docker Compose no está instalado. Por favor instala Docker Compose primero."
    echo "Visita: https://docs.docker.com/compose/install/"
    exit 1
fi
log_success "Docker Compose está instalado"

# Verificar que el archivo docker-compose.prod.yml existe
if [ ! -f "$COMPOSE_FILE" ]; then
    log_error "No se encuentra el archivo $COMPOSE_FILE"
    exit 1
fi

# Verificar si hay un contenedor corriendo
log_info "Verificando si hay una versión anterior corriendo..."
if docker ps -a --format '{{.Names}}' | grep -q "^${CONTAINER_NAME}$"; then
    log_warning "Encontrado contenedor existente. Deteniendo y eliminando..."
    docker-compose -f "$COMPOSE_FILE" down
    log_success "Contenedor anterior eliminado"
else
    log_info "No hay contenedor anterior"
fi

# Descargar la última imagen
log_info "Descargando la última versión de la imagen..."
docker pull "$IMAGE_NAME"
log_success "Imagen descargada exitosamente"

# Crear directorios necesarios
log_info "Creando directorios necesarios..."
mkdir -p logs
mkdir -p data
log_success "Directorios creados"

# Iniciar el contenedor
log_info "Iniciando el contenedor..."
docker-compose -f "$COMPOSE_FILE" up -d
log_success "Contenedor iniciado"

# Esperar a que el contenedor esté saludable
log_info "Esperando que el contenedor esté listo..."
sleep 5

# Verificar que el contenedor está corriendo
if docker ps --format '{{.Names}}' | grep -q "^${CONTAINER_NAME}$"; then
    log_success "Contenedor está corriendo correctamente"
    
    # Mostrar información del contenedor
    echo ""
    echo -e "${BLUE}================================================${NC}"
    echo -e "${BLUE}   INFORMACIÓN DEL DESPLIEGUE${NC}"
    echo -e "${BLUE}================================================${NC}"
    echo -e "${GREEN}Estado:${NC} Activo"
    echo -e "${GREEN}Contenedor:${NC} $CONTAINER_NAME"
    echo -e "${GREEN}Imagen:${NC} $IMAGE_NAME"
    echo -e "${GREEN}Puerto:${NC} 80 (HTTP)"
    echo ""
    
    # Obtener la IP del VPS
    SERVER_IP=$(curl -s ifconfig.me 2>/dev/null || echo "No disponible")
    echo -e "${GREEN}Acceso:${NC}"
    echo -e "  • Local: http://localhost"
    if [ "$SERVER_IP" != "No disponible" ]; then
        echo -e "  • Remoto: http://$SERVER_IP"
    fi
    echo ""
    
    # Comandos útiles
    echo -e "${BLUE}Comandos útiles:${NC}"
    echo -e "  • Ver logs: docker logs -f $CONTAINER_NAME"
    echo -e "  • Detener: docker-compose -f $COMPOSE_FILE down"
    echo -e "  • Reiniciar: docker-compose -f $COMPOSE_FILE restart"
    echo -e "  • Estado: docker ps"
    echo ""
    
    log_success "Despliegue completado exitosamente!"
    
else
    log_error "El contenedor no se inició correctamente"
    echo ""
    echo "Mostrando logs del contenedor:"
    docker logs "$CONTAINER_NAME"
    exit 1
fi

echo -e "${BLUE}================================================${NC}"
echo ""


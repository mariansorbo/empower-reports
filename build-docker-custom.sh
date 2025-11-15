#!/bin/bash

# =============================================================================
# SCRIPT PARA CONSTRUIR IMAGEN DOCKER PERSONALIZADA
# =============================================================================
# Este script construye una imagen Docker con tus propias variables de Azure
# =============================================================================

set -e  # Salir si hay algún error

# Colores para output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

echo -e "${BLUE}================================================${NC}"
echo -e "${BLUE}   BUILD PERSONALIZADO - REPORT TUNER${NC}"
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

# Solicitar variables de entorno
echo -e "${YELLOW}Este script te ayudará a construir una imagen personalizada${NC}"
echo -e "${YELLOW}con tus propias credenciales de Azure.${NC}"
echo ""

# Valores por defecto
DEFAULT_ACCOUNT_NAME=""
DEFAULT_CONTAINER_NAME="pbits-in"
DEFAULT_APP_NAME="Report Tuner"
DEFAULT_MAX_FILE_SIZE="31457280"
DEFAULT_IMAGE_NAME="empower-reports"
DEFAULT_TAG="custom"

# Solicitar Azure Account Name
read -p "$(echo -e ${BLUE}Azure Account Name ${NC}[${DEFAULT_ACCOUNT_NAME}]: )" AZURE_ACCOUNT_NAME
AZURE_ACCOUNT_NAME=${AZURE_ACCOUNT_NAME:-$DEFAULT_ACCOUNT_NAME}

# Solicitar SAS Token (obligatorio)
read -sp "$(echo -e ${BLUE}Azure SAS Token ${NC}(obligatorio): )" AZURE_SAS_TOKEN
echo ""
if [ -z "$AZURE_SAS_TOKEN" ]; then
    log_error "El SAS Token es obligatorio"
    exit 1
fi

# Solicitar Container Name
read -p "$(echo -e ${BLUE}Container Name ${NC}[${DEFAULT_CONTAINER_NAME}]: )" CONTAINER_NAME
CONTAINER_NAME=${CONTAINER_NAME:-$DEFAULT_CONTAINER_NAME}

# Solicitar App Name
read -p "$(echo -e ${BLUE}App Name ${NC}[${DEFAULT_APP_NAME}]: )" APP_NAME
APP_NAME=${APP_NAME:-$DEFAULT_APP_NAME}

# Solicitar Max File Size
read -p "$(echo -e ${BLUE}Max File Size ${NC}[${DEFAULT_MAX_FILE_SIZE}]: )" MAX_FILE_SIZE
MAX_FILE_SIZE=${MAX_FILE_SIZE:-$DEFAULT_MAX_FILE_SIZE}

# Solicitar nombre de imagen
read -p "$(echo -e ${BLUE}Nombre de la imagen ${NC}[${DEFAULT_IMAGE_NAME}]: )" IMAGE_NAME
IMAGE_NAME=${IMAGE_NAME:-$DEFAULT_IMAGE_NAME}

# Solicitar tag
read -p "$(echo -e ${BLUE}Tag de la imagen ${NC}[${DEFAULT_TAG}]: )" TAG
TAG=${TAG:-$DEFAULT_TAG}

FULL_IMAGE_NAME="${IMAGE_NAME}:${TAG}"

echo ""
echo -e "${BLUE}================================================${NC}"
echo -e "${BLUE}   RESUMEN DE CONFIGURACIÓN${NC}"
echo -e "${BLUE}================================================${NC}"
echo -e "${GREEN}Azure Account:${NC} $AZURE_ACCOUNT_NAME"
echo -e "${GREEN}SAS Token:${NC} ${AZURE_SAS_TOKEN:0:20}... (oculto)"
echo -e "${GREEN}Container:${NC} $CONTAINER_NAME"
echo -e "${GREEN}App Name:${NC} $APP_NAME"
echo -e "${GREEN}Max File Size:${NC} $MAX_FILE_SIZE bytes"
echo -e "${GREEN}Imagen:${NC} $FULL_IMAGE_NAME"
echo ""

# Confirmar
read -p "$(echo -e ${YELLOW}¿Continuar con el build? ${NC}[S/n]: )" CONFIRM
CONFIRM=${CONFIRM:-S}

if [[ ! $CONFIRM =~ ^[Ss]$ ]]; then
    log_warning "Build cancelado por el usuario"
    exit 0
fi

echo ""
log_info "Iniciando construcción de la imagen..."
echo ""

# Construir la imagen
docker build \
  --build-arg VITE_AZURE_ACCOUNT_NAME="$AZURE_ACCOUNT_NAME" \
  --build-arg VITE_AZURE_SAS_TOKEN="$AZURE_SAS_TOKEN" \
  --build-arg VITE_CONTAINER_NAME="$CONTAINER_NAME" \
  --build-arg VITE_APP_NAME="$APP_NAME" \
  --build-arg VITE_MAX_FILE_SIZE="$MAX_FILE_SIZE" \
  -t "$FULL_IMAGE_NAME" \
  .

if [ $? -eq 0 ]; then
    echo ""
    log_success "Imagen construida exitosamente: $FULL_IMAGE_NAME"
    echo ""
    
    # Mostrar información de la imagen
    log_info "Información de la imagen:"
    docker images "$IMAGE_NAME"
    echo ""
    
    # Instrucciones siguientes
    echo -e "${BLUE}================================================${NC}"
    echo -e "${BLUE}   PRÓXIMOS PASOS${NC}"
    echo -e "${BLUE}================================================${NC}"
    echo ""
    echo -e "${GREEN}1. Ejecutar localmente:${NC}"
    echo "   docker run -d -p 3000:80 --name empower-reports $FULL_IMAGE_NAME"
    echo ""
    echo -e "${GREEN}2. Probar en navegador:${NC}"
    echo "   http://localhost:3000"
    echo ""
    echo -e "${GREEN}3. Ver logs:${NC}"
    echo "   docker logs -f empower-reports"
    echo ""
    echo -e "${GREEN}4. Subir a Docker Hub (opcional):${NC}"
    echo "   docker login"
    echo "   docker tag $FULL_IMAGE_NAME tu-usuario/$IMAGE_NAME:$TAG"
    echo "   docker push tu-usuario/$IMAGE_NAME:$TAG"
    echo ""
    echo -e "${GREEN}5. Desplegar en VPS:${NC}"
    echo "   # En tu VPS:"
    echo "   docker pull tu-usuario/$IMAGE_NAME:$TAG"
    echo "   docker run -d -p 80:80 --name empower-reports tu-usuario/$IMAGE_NAME:$TAG"
    echo ""
    log_success "¡Listo! Tu imagen personalizada está construida."
else
    log_error "Error al construir la imagen"
    exit 1
fi


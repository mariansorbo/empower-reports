# 🐳 Despliegue con Docker - Report Tuner

Esta guía te explica cómo desplegar la aplicación Report Tuner usando Docker en tu VPS de Hostinger.

## 📋 Prerrequisitos

- Docker instalado en tu VPS
- Cuenta de Docker Hub
- Repositorio en GitHub

## 🚀 Opción 1: Despliegue Automático con GitHub Actions

### Paso 1: Configurar GitHub Secrets

1. Ve a tu repositorio en GitHub
2. Navega a **Settings** → **Secrets and variables** → **Actions**
3. Agrega estos secrets:
   - `DOCKER_USERNAME`: Tu usuario de Docker Hub
   - `DOCKER_PASSWORD`: Tu token de acceso de Docker Hub

### Paso 2: Subir código a GitHub

```bash
git add .
git commit -m "Add Docker configuration"
git push origin main
```

### Paso 3: Verificar el build automático

- Ve a la pestaña **Actions** en tu repositorio
- Verifica que el workflow se ejecute correctamente
- La imagen se subirá automáticamente a Docker Hub

### Paso 4: Desplegar en tu VPS

```bash
# En tu VPS de Hostinger
docker pull tu-usuario/empower-reports:latest
docker run -d -p 3000:80 --name empower-reports tu-usuario/empower-reports:latest
```

## 🔧 Opción 2: Despliegue Manual

### Paso 1: Construir la imagen localmente

```bash
# En tu máquina local
docker build -t empower-reports .
```

### Paso 2: Subir a Docker Hub

```bash
# Etiquetar la imagen
docker tag empower-reports tu-usuario/empower-reports:latest

# Subir a Docker Hub
docker push tu-usuario/empower-reports:latest
```

### Paso 3: Desplegar en VPS

```bash
# En tu VPS
docker pull tu-usuario/empower-reports:latest
docker run -d -p 3000:80 --name empower-reports tu-usuario/empower-reports:latest
```

## 🐙 Usando Docker Compose (Recomendado)

### Crear archivo de configuración para producción

```yaml
# docker-compose.prod.yml
version: '3.8'

services:
  empower-reports:
    image: tu-usuario/empower-reports:latest
    container_name: empower-reports-app
    ports:
      - "80:80"  # Cambiar puerto según necesites
    restart: unless-stopped
    environment:
      - NODE_ENV=production
    volumes:
      - ./data:/usr/share/nginx/html/data:ro
    networks:
      - empower-network

networks:
  empower-network:
    driver: bridge
```

### Desplegar con Docker Compose

```bash
# Descargar la imagen
docker-compose -f docker-compose.prod.yml pull

# Ejecutar en segundo plano
docker-compose -f docker-compose.prod.yml up -d
```

## 🔄 Actualizaciones Automáticas

### Usando Watchtower (Opcional)

```bash
# Instalar Watchtower para actualizaciones automáticas
docker run -d \
  --name watchtower \
  -v /var/run/docker.sock:/var/run/docker.sock \
  containrrr/watchtower \
  --interval 300 \
  empower-reports-app
```

## 🛠️ Comandos Útiles

```bash
# Ver logs de la aplicación
docker logs empower-reports-app

# Entrar al contenedor
docker exec -it empower-reports-app sh

# Detener la aplicación
docker stop empower-reports-app

# Eliminar la aplicación
docker rm empower-reports-app

# Ver imágenes disponibles
docker images

# Limpiar imágenes no utilizadas
docker system prune -a
```

## 🌐 Configuración de Dominio (Opcional)

Si quieres usar un dominio personalizado:

1. Configura el DNS de tu dominio para apuntar a la IP de tu VPS
2. Usa un proxy reverso como Nginx o Traefik
3. Configura SSL con Let's Encrypt

### Ejemplo con Nginx

```nginx
server {
    listen 80;
    server_name tu-dominio.com;
    
    location / {
        proxy_pass http://localhost:3000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

## 🔍 Verificación del Despliegue

1. **Verificar que el contenedor esté corriendo:**
   ```bash
   docker ps
   ```

2. **Verificar logs:**
   ```bash
   docker logs empower-reports-app
   ```

3. **Acceder a la aplicación:**
   - Abre tu navegador
   - Ve a `http://tu-ip:3000` o `http://tu-dominio.com`

## 🆘 Solución de Problemas

### Error: Puerto ya en uso
```bash
# Ver qué proceso usa el puerto
sudo netstat -tulpn | grep :3000

# Cambiar puerto en docker-compose.yml
ports:
  - "3001:80"  # Usar puerto 3001 en lugar de 3000
```

### Error: No se puede conectar a Docker Hub
```bash
# Verificar conexión
docker login

# Verificar que la imagen existe
docker search tu-usuario/empower-reports
```

### Error: Permisos de archivos
```bash
# Dar permisos correctos
sudo chown -R $USER:$USER ./data
chmod -R 755 ./data
```

## 📞 Soporte

Si tienes problemas con el despliegue:

1. Revisa los logs: `docker logs empower-reports-app`
2. Verifica la configuración de red: `docker network ls`
3. Asegúrate de que el puerto esté abierto en tu VPS
4. Verifica que Docker esté funcionando: `docker version`

---

¡Tu aplicación Report Tuner estará lista para usar! 🎉

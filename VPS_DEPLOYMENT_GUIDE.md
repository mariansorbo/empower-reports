# 🚀 Guía Completa de Despliegue en VPS

Esta guía te llevará paso a paso para desplegar **Report Tuner** en tu VPS usando Docker.

## 📋 Tabla de Contenidos

1. [Requisitos Previos](#requisitos-previos)
2. [Preparación del VPS](#preparación-del-vps)
3. [Configuración de Variables de Entorno](#configuración-de-variables-de-entorno)
4. [Opciones de Despliegue](#opciones-de-despliegue)
5. [Gestión del Contenedor](#gestión-del-contenedor)
6. [Configuración de Dominio y SSL](#configuración-de-dominio-y-ssl)
7. [Monitoreo y Logs](#monitoreo-y-logs)
8. [Solución de Problemas](#solución-de-problemas)

---

## 📋 Requisitos Previos

Antes de comenzar, asegúrate de tener:

- ✅ Un VPS activo (Ubuntu 20.04+ o Debian 11+ recomendado)
- ✅ Acceso SSH al VPS
- ✅ Cuenta en Azure con un Storage Account configurado
- ✅ Token SAS de Azure con permisos necesarios
- ✅ (Opcional) Dominio apuntando a la IP de tu VPS

---

## 🔧 Preparación del VPS

### 1. Conectarse al VPS

```bash
ssh usuario@tu-ip-vps
```

### 2. Actualizar el Sistema

```bash
sudo apt update && sudo apt upgrade -y
```

### 3. Instalar Docker

```bash
# Instalar Docker
curl -fsSL https://get.docker.com -o get-docker.sh
sudo sh get-docker.sh

# Agregar tu usuario al grupo docker
sudo usermod -aG docker $USER

# Reiniciar sesión para aplicar cambios
exit
# Vuelve a conectarte
ssh usuario@tu-ip-vps

# Verificar instalación
docker --version
```

### 4. Instalar Docker Compose

```bash
# Instalar Docker Compose
sudo curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-$(uname -s)-$(uname -m)" -o /usr/local/bin/docker-compose
sudo chmod +x /usr/local/bin/docker-compose

# Verificar instalación
docker-compose --version
```

### 5. Configurar Firewall

```bash
# Permitir puerto 80 (HTTP)
sudo ufw allow 80/tcp

# Permitir puerto 443 (HTTPS) si usarás SSL
sudo ufw allow 443/tcp

# Permitir SSH
sudo ufw allow 22/tcp

# Habilitar firewall
sudo ufw enable

# Ver estado
sudo ufw status
```

---

## ⚙️ Configuración de Variables de Entorno

### Opción A: Variables de Entorno para Docker Compose (Recomendado para producción simple)

Las variables de entorno de Vite deben estar presentes **en tiempo de build**, no en tiempo de ejecución. Para esta aplicación, tienes dos opciones:

#### 1. Usar imagen pre-construida de Docker Hub

Si usas la imagen `gimzalo/empower-reports:latest` de Docker Hub, esta ya viene con las variables compiladas. Solo necesitas asegurarte de que la imagen se construyó con las variables correctas.

#### 2. Construir la imagen en el VPS con tus propias variables

Crea un archivo `.env.production` en tu VPS:

```bash
nano .env.production
```

Contenido del archivo:

```env
# Azure Storage Configuration
VITE_AZURE_ACCOUNT_NAME=
VITE_AZURE_SAS_TOKEN=
VITE_CONTAINER_NAME=pbits-in
VITE_APP_NAME=Report Tuner
VITE_MAX_FILE_SIZE=31457280
```

**IMPORTANTE:** No subas este archivo a GitHub. Mantenlo solo en tu VPS.

### Opción B: Construir Imagen Personalizada en el VPS

Si necesitas construir con tus propias variables:

```bash
# Clonar el repositorio
git clone https://github.com/mariansorbo/empower-reports.git
cd empower-reports

# Construir con variables de entorno
docker build \
  --build-arg VITE_AZURE_ACCOUNT_NAME= \
  --build-arg VITE_AZURE_SAS_TOKEN= \
  --build-arg VITE_CONTAINER_NAME=pbits-in \
  --build-arg VITE_APP_NAME="Report Tuner" \
  --build-arg VITE_MAX_FILE_SIZE=31457280 \
  -t empower-reports:custom .
```

---

## 🚀 Opciones de Despliegue

### Opción 1: Despliegue Automático con Script (⭐ Recomendado)

```bash
# Descargar el script de despliegue
wget https://raw.githubusercontent.com/mariansorbo/empower-reports/main/deploy-vps.sh

# Dar permisos de ejecución
chmod +x deploy-vps.sh

# Ejecutar el script
./deploy-vps.sh
```

Este script:
- ✅ Verifica que Docker esté instalado
- ✅ Descarga la última imagen
- ✅ Detiene versiones anteriores
- ✅ Inicia el nuevo contenedor
- ✅ Verifica que todo funcione correctamente

### Opción 2: Despliegue Manual con Docker Compose

```bash
# Crear directorio para el proyecto
mkdir -p ~/empower-reports
cd ~/empower-reports

# Descargar docker-compose.prod.yml
wget https://raw.githubusercontent.com/mariansorbo/empower-reports/main/docker-compose.prod.yml

# Iniciar el contenedor
docker-compose -f docker-compose.prod.yml up -d

# Ver logs
docker-compose -f docker-compose.prod.yml logs -f
```

### Opción 3: Despliegue con Docker Run

```bash
docker pull gimzalo/empower-reports:latest

docker run -d \
  --name empower-reports-app \
  -p 80:80 \
  --restart unless-stopped \
  gimzalo/empower-reports:latest
```

---

## 🔄 Gestión del Contenedor

### Ver Estado del Contenedor

```bash
docker ps
```

### Ver Logs en Tiempo Real

```bash
docker logs -f empower-reports-app
```

### Reiniciar el Contenedor

```bash
docker restart empower-reports-app
```

### Detener el Contenedor

```bash
docker stop empower-reports-app
```

### Eliminar el Contenedor

```bash
docker rm -f empower-reports-app
```

### Actualizar a la Última Versión

```bash
# Opción 1: Con script
./deploy-vps.sh

# Opción 2: Manual
docker pull gimzalo/empower-reports:latest
docker stop empower-reports-app
docker rm empower-reports-app
docker run -d \
  --name empower-reports-app \
  -p 80:80 \
  --restart unless-stopped \
  gimzalo/empower-reports:latest
```

---

## 🌐 Configuración de Dominio y SSL

### 1. Configurar DNS

Apunta tu dominio a la IP de tu VPS:

```
A Record: @ -> tu-ip-vps
A Record: www -> tu-ip-vps
```

### 2. Instalar Certbot para SSL (Let's Encrypt)

```bash
# Instalar Certbot
sudo apt install certbot python3-certbot-nginx -y
```

### 3. Configurar Nginx como Proxy Reverso

```bash
# Instalar Nginx
sudo apt install nginx -y

# Crear configuración
sudo nano /etc/nginx/sites-available/empower-reports
```

Contenido del archivo:

```nginx
server {
    listen 80;
    server_name tu-dominio.com www.tu-dominio.com;

    location / {
        proxy_pass http://localhost:80;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_cache_bypass $http_upgrade;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

```bash
# Habilitar sitio
sudo ln -s /etc/nginx/sites-available/empower-reports /etc/nginx/sites-enabled/

# Verificar configuración
sudo nginx -t

# Reiniciar Nginx
sudo systemctl restart nginx
```

### 4. Obtener Certificado SSL

```bash
sudo certbot --nginx -d tu-dominio.com -d www.tu-dominio.com
```

Sigue las instrucciones y Certbot configurará automáticamente SSL.

### 5. Configurar Renovación Automática

```bash
# Probar renovación
sudo certbot renew --dry-run

# La renovación automática ya está configurada por defecto
```

---

## 📊 Monitoreo y Logs

### Ver Logs de Nginx (dentro del contenedor)

```bash
# Logs de acceso
docker exec empower-reports-app cat /var/log/nginx/access.log

# Logs de errores
docker exec empower-reports-app cat /var/log/nginx/error.log
```

### Logs Persistentes (si configuraste volúmenes)

```bash
# Si configuraste el volumen de logs en docker-compose.prod.yml
tail -f ~/empower-reports/logs/access.log
tail -f ~/empower-reports/logs/error.log
```

### Monitorear Recursos del Contenedor

```bash
# Ver uso de CPU y memoria
docker stats empower-reports-app

# Ver todos los contenedores
docker stats
```

### Health Check

```bash
# Verificar estado de salud
docker inspect empower-reports-app | grep -A 10 "Health"
```

---

## 🆘 Solución de Problemas

### El contenedor no inicia

```bash
# Ver logs completos
docker logs empower-reports-app

# Ver los últimos 100 logs
docker logs --tail 100 empower-reports-app

# Verificar que la imagen se descargó correctamente
docker images | grep empower-reports
```

### Error de puerto en uso

```bash
# Ver qué proceso usa el puerto 80
sudo netstat -tulpn | grep :80

# Opción 1: Detener el servicio que usa el puerto
sudo systemctl stop apache2  # Si es Apache
sudo systemctl stop nginx    # Si es Nginx

# Opción 2: Cambiar puerto en docker-compose.prod.yml
ports:
  - "8080:80"  # Usar puerto 8080 en lugar de 80
```

### La aplicación no carga archivos

Verifica que las variables de entorno de Azure están correctas:

```bash
# Si construiste la imagen en el VPS
# Reconstruye con las variables correctas
docker build \
  --build-arg VITE_AZURE_ACCOUNT_NAME= \
  --build-arg VITE_AZURE_SAS_TOKEN= \
  --build-arg VITE_CONTAINER_NAME=pbits-in \
  -t empower-reports:custom .
```

### Limpiar espacio en disco

```bash
# Ver uso de disco
df -h

# Limpiar imágenes sin usar
docker system prune -a

# Limpiar todo (imágenes, contenedores, volúmenes)
docker system prune -a --volumes
```

### Reiniciar todo desde cero

```bash
# Detener y eliminar todo
docker stop empower-reports-app
docker rm empower-reports-app
docker rmi gimzalo/empower-reports:latest

# Volver a desplegar
./deploy-vps.sh
```

---

## 🔐 Seguridad y Mejores Prácticas

### 1. Mantener el Sistema Actualizado

```bash
# Crear script de actualización automática
sudo apt install unattended-upgrades -y
sudo dpkg-reconfigure --priority=low unattended-upgrades
```

### 2. Configurar Fail2Ban (Protección contra ataques)

```bash
sudo apt install fail2ban -y
sudo systemctl enable fail2ban
sudo systemctl start fail2ban
```

### 3. Backup Regular

```bash
# Crear script de backup
nano ~/backup.sh
```

Contenido:

```bash
#!/bin/bash
DATE=$(date +%Y%m%d_%H%M%S)
BACKUP_DIR="$HOME/backups"
mkdir -p $BACKUP_DIR

# Backup de configuración
docker inspect empower-reports-app > "$BACKUP_DIR/config_$DATE.json"

echo "Backup completado: $BACKUP_DIR/config_$DATE.json"
```

```bash
chmod +x ~/backup.sh

# Agregar a crontab para backup diario
crontab -e
# Agregar: 0 2 * * * ~/backup.sh
```

---

## 📈 Actualizaciones Automáticas con Watchtower (Opcional)

Si quieres que tu aplicación se actualice automáticamente cuando haya nuevas versiones:

```bash
docker run -d \
  --name watchtower \
  -v /var/run/docker.sock:/var/run/docker.sock \
  containrrr/watchtower \
  --interval 3600 \
  --cleanup \
  empower-reports-app
```

Esto verificará cada hora si hay una nueva versión y la instalará automáticamente.

---

## 📞 Verificación Final

Después del despliegue, verifica que todo funcione:

1. **Acceso HTTP:**
   ```bash
   curl http://tu-ip-vps
   ```

2. **Acceso desde navegador:**
   - Ve a `http://tu-ip-vps` o `http://tu-dominio.com`

3. **Verificar subida de archivos:**
   - Intenta subir un archivo .pbit
   - Verifica los logs: `docker logs -f empower-reports-app`

4. **Verificar Azure Storage:**
   - Comprueba que los archivos aparecen en tu contenedor de Azure

---

## 🎉 ¡Listo!

Tu aplicación Report Tuner está ahora desplegada en tu VPS y lista para usar.

### Enlaces Útiles

- [Documentación de Docker](https://docs.docker.com/)
- [Docker Compose](https://docs.docker.com/compose/)
- [Let's Encrypt](https://letsencrypt.org/)
- [Nginx](https://nginx.org/en/docs/)

### Soporte

Si tienes problemas:
1. Revisa los logs: `docker logs empower-reports-app`
2. Verifica el estado: `docker ps`
3. Comprueba los recursos: `docker stats`

---

**Última actualización:** Noviembre 2025
**Versión de la guía:** 1.0


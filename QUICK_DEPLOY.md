# ⚡ Guía Rápida de Despliegue - Report Tuner

Esta es una guía condensada para desplegar rápidamente en tu VPS. Para más detalles, consulta [VPS_DEPLOYMENT_GUIDE.md](VPS_DEPLOYMENT_GUIDE.md).

## 🚀 Despliegue en 5 Minutos

### 1️⃣ Preparar el VPS

```bash
# Conectarse al VPS
ssh usuario@tu-ip-vps

# Instalar Docker
curl -fsSL https://get.docker.com -o get-docker.sh
sudo sh get-docker.sh
sudo usermod -aG docker $USER

# Cerrar sesión y volver a conectar
exit
ssh usuario@tu-ip-vps

# Instalar Docker Compose
sudo curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-$(uname -s)-$(uname -m)" -o /usr/local/bin/docker-compose
sudo chmod +x /usr/local/bin/docker-compose

# Configurar firewall
sudo ufw allow 80/tcp
sudo ufw allow 443/tcp
sudo ufw allow 22/tcp
sudo ufw enable
```

### 2️⃣ Desplegar la Aplicación

**Opción A: Con script automático (⭐ Recomendado)**

```bash
# Descargar y ejecutar script
wget https://raw.githubusercontent.com/mariansorbo/empower-reports/main/deploy-vps.sh
chmod +x deploy-vps.sh
./deploy-vps.sh
```

**Opción B: Manual**

```bash
# Crear directorio
mkdir -p ~/empower-reports && cd ~/empower-reports

# Descargar configuración
wget https://raw.githubusercontent.com/mariansorbo/empower-reports/main/docker-compose.prod.yml

# Iniciar
docker-compose -f docker-compose.prod.yml up -d
```

**Opción C: Docker directo**

```bash
docker pull gimzalo/empower-reports:latest

docker run -d \
  --name empower-reports-app \
  -p 80:80 \
  --restart unless-stopped \
  gimzalo/empower-reports:latest
```

### 3️⃣ Verificar

```bash
# Ver estado
docker ps

# Ver logs
docker logs -f empower-reports-app

# Probar en navegador
# http://tu-ip-vps
```

---

## 📝 Comandos Útiles

### Gestión Básica

```bash
# Ver logs
docker logs -f empower-reports-app

# Reiniciar
docker restart empower-reports-app

# Detener
docker stop empower-reports-app

# Ver recursos
docker stats empower-reports-app
```

### Actualizar a Nueva Versión

```bash
# Con script
./deploy-vps.sh

# Manual
docker pull gimzalo/empower-reports:latest
docker stop empower-reports-app
docker rm empower-reports-app
docker run -d --name empower-reports-app -p 80:80 --restart unless-stopped gimzalo/empower-reports:latest
```

### Solución Rápida de Problemas

```bash
# Ver logs detallados
docker logs --tail 100 empower-reports-app

# Reiniciar todo
docker restart empower-reports-app

# Eliminar y volver a crear
docker rm -f empower-reports-app
./deploy-vps.sh

# Limpiar espacio
docker system prune -a
```

---

## 🌐 Configurar Dominio (Opcional)

### 1. Configurar DNS
- Crea un registro A apuntando a tu IP del VPS

### 2. Instalar SSL

```bash
# Instalar Certbot y Nginx
sudo apt install certbot python3-certbot-nginx nginx -y

# Crear configuración de Nginx
sudo nano /etc/nginx/sites-available/empower-reports
```

Pega esto:

```nginx
server {
    listen 80;
    server_name tu-dominio.com www.tu-dominio.com;

    location / {
        proxy_pass http://localhost:80;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

```bash
# Activar sitio
sudo ln -s /etc/nginx/sites-available/empower-reports /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl restart nginx

# Obtener certificado SSL
sudo certbot --nginx -d tu-dominio.com -d www.tu-dominio.com
```

---

## 🔥 Construcción Personalizada (con tus propias variables)

Si necesitas usar tus propias credenciales de Azure:

```bash
# Clonar repositorio
git clone https://github.com/mariansorbo/empower-reports.git
cd empower-reports

# Construir con tus variables
docker build \
  --build-arg VITE_AZURE_ACCOUNT_NAME= \
  --build-arg VITE_AZURE_SAS_TOKEN= \
  --build-arg VITE_CONTAINER_NAME=pbits-in \
  --build-arg VITE_APP_NAME="Report Tuner" \
  --build-arg VITE_MAX_FILE_SIZE=31457280 \
  -t empower-reports:custom .

# Ejecutar tu versión personalizada
docker run -d \
  --name empower-reports-app \
  -p 80:80 \
  --restart unless-stopped \
  empower-reports:custom
```

---

## ✅ Checklist de Verificación

- [ ] Docker instalado (`docker --version`)
- [ ] Docker Compose instalado (`docker-compose --version`)
- [ ] Firewall configurado (puertos 80, 443, 22)
- [ ] Contenedor corriendo (`docker ps`)
- [ ] Aplicación accesible desde navegador
- [ ] Logs sin errores (`docker logs empower-reports-app`)
- [ ] (Opcional) Dominio configurado
- [ ] (Opcional) SSL instalado

---

## 📞 Ayuda Rápida

**Problema: Puerto en uso**
```bash
sudo netstat -tulpn | grep :80
# Cambiar puerto en docker run: -p 8080:80
```

**Problema: Contenedor no inicia**
```bash
docker logs empower-reports-app
docker system prune -a
./deploy-vps.sh
```

**Problema: Sin espacio en disco**
```bash
df -h
docker system prune -a --volumes
```

---

## 📚 Documentación Completa

Para instrucciones detalladas, configuración avanzada y solución de problemas:
- [VPS_DEPLOYMENT_GUIDE.md](VPS_DEPLOYMENT_GUIDE.md) - Guía completa
- [DOCKER_DEPLOYMENT.md](DOCKER_DEPLOYMENT.md) - Opciones de Docker

---

**¡Tu aplicación estará corriendo en minutos!** 🎉


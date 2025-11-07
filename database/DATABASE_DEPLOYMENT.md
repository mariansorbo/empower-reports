# 🗄️ Guía de Despliegue - Empower Reports Database

Esta guía te ayudará a desplegar la base de datos de Empower Reports en diferentes entornos.

## 📋 Tabla de Contenidos

1. [SQL Server en Docker (Local/VPS)](#opción-1-sql-server-en-docker)
2. [Azure SQL Database (Cloud)](#opción-2-azure-sql-database)
3. [Conexión y Verificación](#conexión-y-verificación)
4. [Mantenimiento](#mantenimiento)

---

## 🐳 Opción 1: SQL Server en Docker (Local/VPS)

### Requisitos Previos

- Docker instalado
- Docker Compose instalado
- Al menos 2GB de RAM disponible
- Puerto 1433 libre

### Paso 1: Configurar Variables de Entorno

```bash
cd database/

# Copiar archivo de ejemplo
cp .env.example .env

# Editar y cambiar la contraseña
nano .env
```

**Importante:** Cambia `YourStrong!Passw0rd` por una contraseña segura.

### Paso 2: Construir y Levantar el Contenedor

```bash
# Construir la imagen
docker-compose build

# Levantar el contenedor
docker-compose up -d

# Ver logs
docker-compose logs -f
```

### Paso 3: Verificar que Esté Corriendo

```bash
# Ver estado
docker-compose ps

# Verificar logs de inicialización
docker logs empower-reports-db

# Debería mostrar: "¡Base de datos inicializada correctamente!"
```

### Credenciales de Acceso

```
Servidor: localhost,1433
Usuario: sa
Contraseña: YourStrong!Passw0rd (la que configuraste)
Base de datos: empower_reports
```

---

## ☁️ Opción 2: Azure SQL Database (Cloud)

### Paso 1: Crear Azure SQL Database

#### Desde el Portal de Azure:

1. **Ir a "SQL databases"** en el Portal de Azure
2. Click en **"+ Create"**
3. Configurar:
   - **Subscription:** Tu suscripción
   - **Resource Group:** Crear nuevo o usar existente
   - **Database name:** `empower-reports-db`
   - **Server:** Crear nuevo servidor
     - **Server name:** `empower-reports-server` (debe ser único)
     - **Location:** Elige tu región
     - **Authentication:** SQL authentication
     - **Admin login:** `sqladmin`
     - **Password:** Tu contraseña segura
   - **Compute + storage:** 
     - Basic (5 DTUs) para desarrollo
     - Standard o Premium para producción
   - **Backup storage redundancy:** Locally-redundant

4. Click **"Review + create"** → **"Create"**
5. Esperar ~5 minutos a que se cree

#### Desde Azure CLI:

```bash
# Login
az login

# Variables
RESOURCE_GROUP="empower-reports-rg"
LOCATION="eastus"
SERVER_NAME="empower-reports-server"
DB_NAME="empower-reports-db"
ADMIN_USER="sqladmin"
ADMIN_PASSWORD="TuContraseñaSegura123!"

# Crear resource group
az group create --name $RESOURCE_GROUP --location $LOCATION

# Crear SQL Server
az sql server create \
  --name $SERVER_NAME \
  --resource-group $RESOURCE_GROUP \
  --location $LOCATION \
  --admin-user $ADMIN_USER \
  --admin-password $ADMIN_PASSWORD

# Crear base de datos
az sql db create \
  --resource-group $RESOURCE_GROUP \
  --server $SERVER_NAME \
  --name $DB_NAME \
  --edition Basic \
  --capacity 5

# Permitir acceso desde Azure services
az sql server firewall-rule create \
  --resource-group $RESOURCE_GROUP \
  --server $SERVER_NAME \
  --name AllowAzureServices \
  --start-ip-address 0.0.0.0 \
  --end-ip-address 0.0.0.0

# Permitir tu IP actual
MY_IP=$(curl -s ifconfig.me)
az sql server firewall-rule create \
  --resource-group $RESOURCE_GROUP \
  --server $SERVER_NAME \
  --name AllowMyIP \
  --start-ip-address $MY_IP \
  --end-ip-address $MY_IP
```

### Paso 2: Configurar Firewall

En el Portal de Azure:

1. Ve a tu SQL Server
2. Click en **"Networking"** o **"Firewalls and virtual networks"**
3. Agregar reglas:
   - **Allow Azure services:** ON
   - **Add client IP:** Click para agregar tu IP actual
   - **Add your VPS IP:** Si vas a conectarte desde un VPS

4. Click **"Save"**

### Paso 3: Ejecutar Scripts de Inicialización

#### Opción A: Desde Azure Data Studio (Recomendado)

1. Descargar [Azure Data Studio](https://aka.ms/azuredatastudio)
2. Conectarse:
   - **Server:** `empower-reports-server.database.windows.net`
   - **Authentication:** SQL Login
   - **User:** `sqladmin`
   - **Password:** Tu contraseña
   - **Database:** `empower-reports-db`

3. Ejecutar scripts en orden:
   ```
   1. schema.sql
   2. organization_workflows.sql
   3. state_machine_and_workflows.sql
   4. constraints_and_validations.sql
   5. migrations/001_fix_billing_cycle_and_organization_null.sql
   ```

#### Opción B: Desde sqlcmd (Línea de comandos)

```bash
# Instalar sqlcmd
# Windows: https://learn.microsoft.com/en-us/sql/tools/sqlcmd-utility
# Linux: sudo apt-get install mssql-tools

# Variables
SERVER="empower-reports-server.database.windows.net"
USER="sqladmin"
PASSWORD="TuContraseña"
DATABASE="empower-reports-db"

# Ejecutar scripts
sqlcmd -S $SERVER -U $USER -P $PASSWORD -d $DATABASE -i schema.sql
sqlcmd -S $SERVER -U $USER -P $PASSWORD -d $DATABASE -i organization_workflows.sql
sqlcmd -S $SERVER -U $USER -P $PASSWORD -d $DATABASE -i state_machine_and_workflows.sql
sqlcmd -S $SERVER -U $USER -P $PASSWORD -d $DATABASE -i constraints_and_validations.sql
sqlcmd -S $SERVER -U $USER -P $PASSWORD -d $DATABASE -i migrations/001_fix_billing_cycle_and_organization_null.sql
```

### Credenciales de Acceso (Azure SQL)

```
Servidor: empower-reports-server.database.windows.net
Puerto: 1433
Usuario: sqladmin
Contraseña: [tu contraseña]
Base de datos: empower-reports-db
```

**Connection String:**
```
Server=tcp:empower-reports-server.database.windows.net,1433;
Database=empower-reports-db;
User ID=sqladmin;
Password={tu_contraseña};
Encrypt=yes;
TrustServerCertificate=no;
Connection Timeout=30;
```

---

## 🔌 Conexión y Verificación

### Probar Conexión

**Desde sqlcmd:**
```bash
# Local (Docker)
sqlcmd -S localhost,1433 -U sa -P YourStrong!Passw0rd -d empower_reports -Q "SELECT COUNT(*) FROM plans"

# Azure SQL
sqlcmd -S empower-reports-server.database.windows.net -U sqladmin -P TuContraseña -d empower-reports-db -Q "SELECT COUNT(*) FROM plans"
```

**Desde Python:**
```python
import pyodbc

# Local
conn_str = (
    'Driver={ODBC Driver 18 for SQL Server};'
    'Server=localhost,1433;'
    'Database=empower_reports;'
    'UID=sa;'
    'PWD=YourStrong!Passw0rd;'
    'TrustServerCertificate=yes;'
)

# Azure SQL
conn_str = (
    'Driver={ODBC Driver 18 for SQL Server};'
    'Server=tcp:empower-reports-server.database.windows.net,1433;'
    'Database=empower-reports-db;'
    'UID=sqladmin;'
    'PWD=TuContraseña;'
    'Encrypt=yes;'
    'TrustServerCertificate=no;'
)

conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
cursor.execute("SELECT COUNT(*) FROM plans")
print(f"Planes en la base de datos: {cursor.fetchone()[0]}")
```

**Desde Node.js:**
```javascript
const sql = require('mssql');

// Local
const config = {
    user: 'sa',
    password: 'YourStrong!Passw0rd',
    server: 'localhost',
    database: 'empower_reports',
    options: {
        encrypt: false,
        trustServerCertificate: true
    }
};

// Azure SQL
const config = {
    user: 'sqladmin',
    password: 'TuContraseña',
    server: 'empower-reports-server.database.windows.net',
    database: 'empower-reports-db',
    options: {
        encrypt: true,
        trustServerCertificate: false
    }
};

async function testConnection() {
    try {
        await sql.connect(config);
        const result = await sql.query`SELECT COUNT(*) as count FROM plans`;
        console.log(`Planes: ${result.recordset[0].count}`);
    } catch (err) {
        console.error('Error:', err);
    }
}
```

### Verificar Tablas Creadas

```sql
-- Ver todas las tablas
SELECT TABLE_NAME 
FROM INFORMATION_SCHEMA.TABLES 
WHERE TABLE_TYPE = 'BASE TABLE'
ORDER BY TABLE_NAME;

-- Ver datos de planes
SELECT * FROM plans;

-- Ver usuarios (debería estar vacío)
SELECT COUNT(*) as total_users FROM users;
```

---

## 🛠️ Mantenimiento

### Backup (Docker)

```bash
# Crear backup
docker exec empower-reports-db /opt/mssql-tools/bin/sqlcmd \
  -S localhost -U sa -P YourStrong!Passw0rd \
  -Q "BACKUP DATABASE empower_reports TO DISK='/var/opt/mssql/backup/empower_reports.bak'"

# Copiar backup a host
docker cp empower-reports-db:/var/opt/mssql/backup/empower_reports.bak ./backup/
```

### Backup (Azure SQL)

Azure SQL tiene backups automáticos. Para backups manuales:

```bash
# Desde Azure CLI
az sql db export \
  --resource-group empower-reports-rg \
  --server empower-reports-server \
  --name empower-reports-db \
  --admin-user sqladmin \
  --admin-password TuContraseña \
  --storage-key-type StorageAccessKey \
  --storage-key [storage_key] \
  --storage-uri https://[storage_account].blob.core.windows.net/backups/empower-reports.bacpac
```

### Restaurar Backup

**Docker:**
```bash
# Copiar backup al contenedor
docker cp backup/empower_reports.bak empower-reports-db:/var/opt/mssql/backup/

# Restaurar
docker exec empower-reports-db /opt/mssql-tools/bin/sqlcmd \
  -S localhost -U sa -P YourStrong!Passw0rd \
  -Q "RESTORE DATABASE empower_reports FROM DISK='/var/opt/mssql/backup/empower_reports.bak' WITH REPLACE"
```

### Monitoreo

**Docker:**
```bash
# Ver uso de recursos
docker stats empower-reports-db

# Ver logs
docker logs -f empower-reports-db

# Ver conexiones activas
docker exec empower-reports-db /opt/mssql-tools/bin/sqlcmd \
  -S localhost -U sa -P YourStrong!Passw0rd \
  -Q "SELECT COUNT(*) as connections FROM sys.dm_exec_sessions WHERE is_user_process = 1"
```

**Azure SQL:**
- Ve al Portal de Azure → Tu base de datos
- Métricas disponibles: DTU, Storage, Connections, etc.

---

## 🚨 Solución de Problemas

### Puerto 1433 en uso

```bash
# Windows
netstat -ano | findstr :1433

# Linux/Mac
lsof -i :1433

# Cambiar puerto en docker-compose.yml
ports:
  - "1434:1433"  # Usa puerto 1434 en el host
```

### Contenedor no inicia

```bash
# Ver logs detallados
docker logs empower-reports-db

# Verificar permisos del script
chmod +x init-db.sh

# Reconstruir
docker-compose down -v
docker-compose build --no-cache
docker-compose up -d
```

### No puedo conectarme a Azure SQL

1. Verifica firewall en Azure Portal
2. Verifica que tu IP esté permitida
3. Verifica credenciales
4. Prueba conexión desde Azure Data Studio

---

## 📊 Costos Estimados

### Docker (Local/VPS)
- **Gratis** (solo costos del VPS)
- VPS recomendado: 2GB RAM (~$10-20/mes)

### Azure SQL Database
- **Basic:** ~$5/mes (desarrollo)
- **Standard S0:** ~$15/mes (producción pequeña)
- **Standard S1:** ~$30/mes (producción media)
- **Premium:** $465+/mes (alta disponibilidad)

[Calculadora de precios Azure SQL](https://azure.microsoft.com/en-us/pricing/details/azure-sql-database/)

---

## 🎯 Próximos Pasos

1. ✅ Base de datos desplegada
2. ⏭️ Crear backend API para conectarse a la BD
3. ⏭️ Conectar frontend con backend
4. ⏭️ Implementar autenticación
5. ⏭️ Desplegar aplicación completa

---

¿Necesitas ayuda con alguno de estos pasos?


# 🗄️ Guía de Despliegue - Report Tuner Database en Azure SQL Database

Esta guía te ayudará a desplegar la base de datos de Report Tuner en Azure SQL Database.

## 📋 Tabla de Contenidos

1. [Requisitos Previos](#requisitos-previos)
2. [Crear Azure SQL Database](#crear-azure-sql-database)
3. [Configurar Firewall](#configurar-firewall)
4. [Ejecutar Scripts de Inicialización](#ejecutar-scripts-de-inicialización)
5. [Conexión y Verificación](#conexión-y-verificación)
6. [Backup y Restauración](#backup-y-restauración)
7. [Monitoreo y Mantenimiento](#monitoreo-y-mantenimiento)
8. [Solución de Problemas](#solución-de-problemas)
9. [Costos Estimados](#costos-estimados)

---

## ✅ Requisitos Previos

- Suscripción activa de Azure
- Azure CLI instalado (opcional, para automatización)
- Herramienta para ejecutar scripts SQL:
  - [Azure Data Studio](https://aka.ms/azuredatastudio) (Recomendado)
  - [SQL Server Management Studio (SSMS)](https://docs.microsoft.com/en-us/sql/ssms/download-sql-server-management-studio-ssms)
  - `sqlcmd` (línea de comandos)

---

## ☁️ Crear Azure SQL Database

### Opción 1: Desde el Portal de Azure (Recomendado para principiantes)

1. **Ir a "SQL databases"** en el Portal de Azure
2. Click en **"+ Create"** o **"Create SQL database"**
3. **Configurar la base de datos:**
   - **Subscription:** Tu suscripción de Azure
   - **Resource Group:** Crear nuevo o usar existente
     - **Name:** `report-tuner-rg` (o el nombre que prefieras)
     - **Region:** Elige tu región más cercana
   - **Database name:** `report-tuner-db` (o el nombre que prefieras)
   - **Server:** Crear nuevo servidor
     - **Server name:** `report-tuner-server` (debe ser único globalmente)
     - **Location:** Misma región que el Resource Group
     - **Authentication method:** SQL authentication
     - **Server admin login:** `sqladmin` (o el nombre que prefieras)
     - **Password:** Genera una contraseña segura (guárdala bien)
     - **Allow Azure services to access this server:** ✅ Yes (marca esta opción)
   - **Compute + storage:**
     - **Service tier:** 
       - **Basic** (5 DTUs) - Para desarrollo/pruebas (~$5/mes)
       - **Standard S0** (10 DTUs) - Para producción pequeña (~$15/mes)
       - **Standard S1** (20 DTUs) - Para producción media (~$30/mes)
       - **Premium** - Para alta disponibilidad (desde $465/mes)
     - **Backup storage redundancy:** Locally-redundant (LRS) para desarrollo, Geo-redundant (GRS) para producción

4. Click **"Review + create"** → Revisar configuración → **"Create"**
5. Esperar ~3-5 minutos a que se cree la base de datos

### Opción 2: Desde Azure CLI (Para automatización)

```bash
# Login en Azure
az login

# Variables de configuración
RESOURCE_GROUP="report-tuner-rg"
LOCATION="eastus"  # Cambia por tu región preferida
SERVER_NAME="report-tuner-server"  # Debe ser único globalmente
DB_NAME="report-tuner-db"
ADMIN_USER="sqladmin"
ADMIN_PASSWORD="TuContraseñaSegura123!"  # Cambia por una contraseña segura

# Crear Resource Group
az group create \
  --name $RESOURCE_GROUP \
  --location $LOCATION

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
  --capacity 5 \
  --backup-storage-redundancy Local

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

---

## 🔥 Configurar Firewall

### Desde el Portal de Azure:

1. Ve a tu **SQL Server** (no la base de datos)
2. Click en **"Networking"** o **"Firewalls and virtual networks"**
3. **Configurar reglas:**
   - **Allow Azure services and resources to access this server:** ✅ **ON** (muy importante)
   - **Add your client IPv4 address:** Click para agregar tu IP actual
   - **Add your VPS/server IP:** Si vas a conectarte desde un servidor, agrega su IP
   - **Add firewall rule:** Para agregar rangos de IP específicos

4. Click **"Save"**

### Desde Azure CLI:

```bash
# Variables (ajusta según tu configuración)
RESOURCE_GROUP="report-tuner-rg"
SERVER_NAME="report-tuner-server"

# Permitir acceso desde Azure services (ya debería estar hecho)
az sql server firewall-rule create \
  --resource-group $RESOURCE_GROUP \
  --server $SERVER_NAME \
  --name AllowAzureServices \
  --start-ip-address 0.0.0.0 \
  --end-ip-address 0.0.0.0

# Agregar tu IP actual
MY_IP=$(curl -s ifconfig.me)
az sql server firewall-rule create \
  --resource-group $RESOURCE_GROUP \
  --server $SERVER_NAME \
  --name AllowMyIP \
  --start-ip-address $MY_IP \
  --end-ip-address $MY_IP

# Agregar IP de tu VPS/servidor (si aplica)
VPS_IP="1.2.3.4"  # Cambia por la IP de tu servidor
az sql server firewall-rule create \
  --resource-group $RESOURCE_GROUP \
  --server $SERVER_NAME \
  --name AllowVPS \
  --start-ip-address $VPS_IP \
  --end-ip-address $VPS_IP
```

### Importante sobre Firewall:

- **Azure services:** Debe estar habilitado si tu aplicación está en Azure (App Service, Functions, etc.)
- **Tu IP:** Necesitas agregar tu IP para conectarte desde tu máquina local
- **IP dinámica:** Si tu IP cambia, necesitarás agregar la nueva IP
- **Sin IP pública:** Si tu aplicación está en Azure, puede usar "Allow Azure services" sin necesidad de IP específica

---

## 📜 Ejecutar Scripts de Inicialización

**⚠️ Importante:** La base de datos ya debe estar creada en Azure SQL Database. Los scripts NO crean la base de datos, solo crean las tablas, vistas, funciones y procedimientos dentro de la base de datos existente.

### Orden de Ejecución de Scripts:

Ejecuta los scripts en el siguiente orden:

1. **`schema.sql`** ⭐ - Schema principal (tablas, vistas, funciones básicas, triggers)
2. **`organization_workflows.sql`** - Procedures y funciones para creación/unión a organizaciones
3. **`state_machine_and_workflows.sql`** - Máquina de estados de suscripciones
4. **`constraints_and_validations.sql`** - Validaciones adicionales
5. **`documentation_procedures.sql`** - Procedures para gestionar documentación
6. **`migrations/001_fix_billing_cycle_and_organization_null.sql`** - Migraciones (si aplica)
7. **`enterprise_pro_plan_v2.sql`** - Enterprise Pro multi-organización (solo si necesitas esta funcionalidad)

### Opción A: Desde Azure Data Studio (Recomendado)

1. **Descargar e instalar [Azure Data Studio](https://aka.ms/azuredatastudio)**
2. **Conectarse a la base de datos:**
   - Click en **"New Connection"**
   - **Server:** `report-tuner-server.database.windows.net`
   - **Authentication type:** SQL Login
   - **User name:** `sqladmin` (o el que configuraste)
   - **Password:** Tu contraseña
   - **Database:** `report-tuner-db` (selecciona de la lista)
   - Click **"Connect"**

3. **Ejecutar scripts en orden:**
   - Abre cada archivo `.sql` en Azure Data Studio
   - **IMPORTANTE:** Antes de ejecutar `schema.sql`, modifica la línea que dice `USE master;` y `CREATE DATABASE` - **NO las ejecutes**, simplemente comenta o elimina esas líneas, ya que la base de datos ya existe
   - Click en **"Run"** (F5) para ejecutar el script completo
   - Verifica que no haya errores en la pestaña "Messages"

### Opción B: Desde sqlcmd (Línea de comandos)

#### Instalar sqlcmd:

**Windows:**
- Descargar desde: https://learn.microsoft.com/en-us/sql/tools/sqlcmd-utility
- O instalar con Chocolatey: `choco install sqlcmd`

**Linux:**
```bash
# Ubuntu/Debian
curl https://packages.microsoft.com/keys/microsoft.asc | sudo apt-key add -
curl https://packages.microsoft.com/config/ubuntu/20.04/prod.list | sudo tee /etc/apt/sources.list.d/mssql-release.list
sudo apt-get update
sudo apt-get install -y mssql-tools unixodbc-dev
echo 'export PATH="$PATH:/opt/mssql-tools/bin"' >> ~/.bashrc
source ~/.bashrc
```

**macOS:**
```bash
brew install mssql-tools
```

#### Ejecutar scripts:

```bash
# Variables (ajusta según tu configuración)
SERVER="report-tuner-server.database.windows.net"
USER="sqladmin"
PASSWORD="TuContraseñaSegura123!"
DATABASE="report-tuner-db"

# IMPORTANTE: Antes de ejecutar schema.sql, edítalo para comentar/eliminar las líneas:
# - USE master;
# - CREATE DATABASE empower_reports;
# La base de datos ya existe en Azure SQL Database

# Ejecutar scripts en orden
cd database/

# 1. Schema principal
sqlcmd -S $SERVER -U $USER -P $PASSWORD -d $DATABASE -i schema.sql

# 2. Organization workflows
sqlcmd -S $SERVER -U $USER -P $PASSWORD -d $DATABASE -i organization_workflows.sql

# 3. State machine and workflows
sqlcmd -S $SERVER -U $USER -P $PASSWORD -d $DATABASE -i state_machine_and_workflows.sql

# 4. Constraints and validations
sqlcmd -S $SERVER -P $PASSWORD -d $DATABASE -i constraints_and_validations.sql

# 5. Documentation procedures
sqlcmd -S $SERVER -U $USER -P $PASSWORD -d $DATABASE -i documentation_procedures.sql

# 6. Migrations (si aplica)
sqlcmd -S $SERVER -U $USER -P $PASSWORD -d $DATABASE -i migrations/001_fix_billing_cycle_and_organization_null.sql

# 7. Enterprise Pro (solo si necesitas esta funcionalidad)
# sqlcmd -S $SERVER -U $USER -P $PASSWORD -d $DATABASE -i enterprise_pro_plan_v2.sql
```

### Opción C: Desde SQL Server Management Studio (SSMS)

1. **Descargar e instalar [SSMS](https://docs.microsoft.com/en-us/sql/ssms/download-sql-server-management-studio-ssms)**
2. **Conectarse:**
   - Server name: `report-tuner-server.database.windows.net`
   - Authentication: SQL Server Authentication
   - Login: `sqladmin`
   - Password: Tu contraseña
   - Click **"Connect"**

3. **Ejecutar scripts:**
   - Abre cada archivo `.sql` en SSMS
   - **IMPORTANTE:** Modifica `schema.sql` para comentar/eliminar las líneas de creación de base de datos
   - Selecciona la base de datos `report-tuner-db` en el dropdown
   - Click **"Execute"** (F5)

### Modificar schema.sql para Azure SQL Database

**Antes de ejecutar `schema.sql`, edítalo y comenta o elimina estas líneas:**

```sql
-- COMENTAR O ELIMINAR ESTAS LÍNEAS:
-- USE master;
-- GO
-- 
-- IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'empower_reports')
-- BEGIN
--     CREATE DATABASE empower_reports;
-- END
-- GO
-- 
-- USE empower_reports;
-- GO

-- EN SU LUGAR, SOLO DEJA:
USE report-tuner-db;  -- O el nombre de tu base de datos
GO
```

**O simplemente elimina esas líneas y asegúrate de estar conectado a la base de datos correcta antes de ejecutar el script.**

---

## 🔌 Conexión y Verificación

### Credenciales de Acceso

```
Servidor: report-tuner-server.database.windows.net
Puerto: 1433
Usuario: sqladmin
Contraseña: [tu contraseña]
Base de datos: report-tuner-db
```

### Connection String

**Para aplicaciones .NET:**
```
Server=tcp:report-tuner-server.database.windows.net,1433;
Database=report-tuner-db;
User ID=sqladmin;
Password={tu_contraseña};
Encrypt=yes;
TrustServerCertificate=no;
Connection Timeout=30;
```

**Para Node.js (mssql):**
```javascript
const config = {
    user: 'sqladmin',
    password: 'TuContraseñaSegura123!',
    server: 'report-tuner-server.database.windows.net',
    database: 'report-tuner-db',
    options: {
        encrypt: true,
        trustServerCertificate: false,
        enableArithAbort: true
    }
};
```

**Para Python (pyodbc):**
```python
conn_str = (
    'Driver={ODBC Driver 18 for SQL Server};'
    'Server=tcp:report-tuner-server.database.windows.net,1433;'
    'Database=report-tuner-db;'
    'UID=sqladmin;'
    'PWD=TuContraseñaSegura123!;'
    'Encrypt=yes;'
    'TrustServerCertificate=no;'
    'Connection Timeout=30;'
)
```

### Probar Conexión

**Desde sqlcmd:**
```bash
sqlcmd -S report-tuner-server.database.windows.net \
  -U sqladmin \
  -P TuContraseñaSegura123! \
  -d report-tuner-db \
  -Q "SELECT COUNT(*) as plan_count FROM plans"
```

**Desde Node.js:**
```javascript
const sql = require('mssql');

const config = {
    user: 'sqladmin',
    password: 'TuContraseñaSegura123!',
    server: 'report-tuner-server.database.windows.net',
    database: 'report-tuner-db',
    options: {
        encrypt: true,
        trustServerCertificate: false
    }
};

async function testConnection() {
    try {
        await sql.connect(config);
        const result = await sql.query`SELECT COUNT(*) as count FROM plans`;
        console.log(`✅ Conexión exitosa. Planes en la BD: ${result.recordset[0].count}`);
        await sql.close();
    } catch (err) {
        console.error('❌ Error de conexión:', err);
    }
}

testConnection();
```

### Verificar Tablas Creadas

```sql
-- Ver todas las tablas
SELECT TABLE_NAME 
FROM INFORMATION_SCHEMA.TABLES 
WHERE TABLE_TYPE = 'BASE TABLE'
ORDER BY TABLE_NAME;

-- Verificar que los planes se insertaron correctamente
SELECT id, name, max_users, max_reports, price_monthly 
FROM plans 
ORDER BY id;

-- Verificar triggers
SELECT name, type_desc 
FROM sys.objects 
WHERE type = 'TR'
ORDER BY name;

-- Verificar vistas
SELECT name 
FROM sys.views 
ORDER BY name;

-- Verificar funciones
SELECT name, type_desc 
FROM sys.objects 
WHERE type IN ('FN', 'IF', 'TF')
ORDER BY name;
```

### Verificar Datos Iniciales

```sql
-- Ver planes predefinidos
SELECT * FROM plans;

-- Deberías ver 5 planes: free_trial, basic, teams, enterprise, enterprise_pro

-- Verificar que no hay usuarios aún
SELECT COUNT(*) as total_users FROM users;

-- Verificar que no hay organizaciones aún
SELECT COUNT(*) as total_organizations FROM organizations;
```

---

## 💾 Backup y Restauración

### Backup Automático

Azure SQL Database realiza backups automáticos:
- **Backups completos:** Diarios
- **Backups diferenciales:** Cada 12 horas
- **Backups de transacciones:** Cada 5-10 minutos
- **Retención:** 
  - Basic/Standard: 7 días
  - Premium: 35 días (configurable hasta 10 años)

### Backup Manual (Exportar a BACPAC)

**Desde Azure Portal:**
1. Ve a tu base de datos en Azure Portal
2. Click en **"Export"**
3. Configurar:
   - **Storage account:** Selecciona o crea una cuenta de almacenamiento
   - **Container:** Selecciona o crea un contenedor
   - **Blob name:** `report-tuner-backup-YYYY-MM-DD.bacpac`
   - **Authentication type:** SQL Server authentication
   - **Server admin login:** `sqladmin`
   - **Password:** Tu contraseña
4. Click **"OK"** y esperar a que se complete la exportación

**Desde Azure CLI:**
```bash
# Variables
RESOURCE_GROUP="report-tuner-rg"
SERVER_NAME="report-tuner-server"
DB_NAME="report-tuner-db"
STORAGE_ACCOUNT="reporttunerstorage"  # Nombre de tu cuenta de almacenamiento
CONTAINER="backups"
BLOB_NAME="report-tuner-backup-$(date +%Y%m%d).bacpac"
ADMIN_USER="sqladmin"
ADMIN_PASSWORD="TuContraseñaSegura123!"

# Obtener clave de acceso de la cuenta de almacenamiento
STORAGE_KEY=$(az storage account keys list \
  --resource-group $RESOURCE_GROUP \
  --account-name $STORAGE_ACCOUNT \
  --query '[0].value' \
  --output tsv)

# Exportar base de datos
az sql db export \
  --resource-group $RESOURCE_GROUP \
  --server $SERVER_NAME \
  --name $DB_NAME \
  --admin-user $ADMIN_USER \
  --admin-password $ADMIN_PASSWORD \
  --storage-key-type StorageAccessKey \
  --storage-key $STORAGE_KEY \
  --storage-uri "https://$STORAGE_ACCOUNT.blob.core.windows.net/$CONTAINER/$BLOB_NAME"
```

### Restaurar desde Backup

**Desde Azure Portal:**
1. Ve a tu SQL Server en Azure Portal
2. Click en **"Restore database"**
3. Selecciona:
   - **Source:** Point in time restore
   - **Restore point:** Selecciona la fecha y hora
   - **Target database:** Nuevo nombre para la base de datos restaurada
4. Click **"Review + create"** → **"Create"**

**Desde Azure CLI:**
```bash
# Restaurar a un punto en el tiempo
az sql db restore \
  --resource-group $RESOURCE_GROUP \
  --server $SERVER_NAME \
  --name $DB_NAME \
  --dest-name $DB_NAME-restored \
  --time "2024-01-15T10:00:00Z"
```

---

## 📊 Monitoreo y Mantenimiento

### Métricas en Azure Portal

1. Ve a tu base de datos en Azure Portal
2. Sección **"Monitoring"** muestra:
   - **DTU/CPU:** Uso de recursos
   - **Storage:** Espacio utilizado
   - **Connections:** Conexiones activas
   - **Deadlocks:** Deadlocks detectados
   - **Query performance:** Rendimiento de consultas

### Alertas

Configurar alertas para:
- **DTU > 80%:** Escalar o optimizar consultas
- **Storage > 80%:** Considerar aumentar el tamaño
- **Connections > límite:** Revisar conexiones abiertas
- **Deadlocks:** Optimizar transacciones

### Optimización

**Query Performance Insight:**
1. Ve a tu base de datos → **"Query Performance Insight"**
2. Revisa las consultas más costosas
3. Optimiza índices y consultas según sea necesario

**Index Advisor:**
1. Ve a tu base de datos → **"Advisor"**
2. Revisa recomendaciones de índices
3. Aplica índices sugeridos si son relevantes

### Escalado

**Escalar verticalmente (aumentar DTUs):**
```bash
# Escalar a Standard S1 (20 DTUs)
az sql db update \
  --resource-group $RESOURCE_GROUP \
  --server $SERVER_NAME \
  --name $DB_NAME \
  --service-objective S1
```

**Escalar horizontalmente (leer réplicas):**
- Azure SQL Database soporta hasta 4 réplicas de lectura
- Útil para distribuir carga de lectura
- Configurable desde Azure Portal

---

## 🚨 Solución de Problemas

### Error: "Cannot open server"

**Causa:** Firewall bloqueando la conexión

**Solución:**
1. Verifica que tu IP esté en las reglas de firewall
2. Verifica que "Allow Azure services" esté habilitado
3. Agrega tu IP actual desde Azure Portal

### Error: "Login failed for user"

**Causa:** Credenciales incorrectas

**Solución:**
1. Verifica usuario y contraseña
2. Verifica que el usuario tenga permisos en la base de datos
3. Intenta restablecer la contraseña desde Azure Portal

### Error: "Server is not available"

**Causa:** Servidor pausado o no disponible

**Solución:**
1. Verifica el estado del servidor en Azure Portal
2. Si está pausado, reactívalo
3. Si es serverless, espera ~30 segundos (se inicia automáticamente)

### Error al ejecutar scripts: "Database 'empower_reports' does not exist"

**Causa:** El script `schema.sql` intenta crear/usar una base de datos que no existe

**Solución:**
1. Edita `schema.sql` y comenta/elimina las líneas que crean la base de datos
2. Asegúrate de estar conectado a la base de datos correcta antes de ejecutar
3. Cambia `USE empower_reports;` por `USE report-tuner-db;` (o el nombre de tu BD)

### Error: "The server principal is not able to access the database"

**Causa:** El usuario no tiene permisos en la base de datos

**Solución:**
1. Conéctate como administrador del servidor
2. Ejecuta: `ALTER ROLE db_owner ADD MEMBER [sqladmin];`
3. O crea un usuario específico para la aplicación con permisos limitados

### Conexión lenta

**Causa:** DTUs insuficientes o consultas no optimizadas

**Solución:**
1. Revisa métricas de DTU en Azure Portal
2. Si DTU > 80%, considera escalar
3. Optimiza consultas usando Query Performance Insight
4. Agrega índices según recomendaciones

---

## 💰 Costos Estimados

### Niveles de Servicio

| Nivel | DTUs | Precio/mes | Uso recomendado |
|-------|------|------------|-----------------|
| **Basic** | 5 | ~$5 | Desarrollo, pruebas |
| **Standard S0** | 10 | ~$15 | Producción pequeña |
| **Standard S1** | 20 | ~$30 | Producción media |
| **Standard S2** | 50 | ~$75 | Producción grande |
| **Standard S3** | 100 | ~$150 | Producción muy grande |
| **Premium P1** | 125 | ~$465 | Alta disponibilidad |
| **Premium P2** | 250 | ~$930 | Alta disponibilidad + rendimiento |

### Storage

- **Incluido:** 2GB - 1TB (dependiendo del nivel)
- **Extra:** ~$0.12/GB/mes

### Backup Storage

- **Incluido:** 100% del tamaño de la base de datos
- **Extra:** ~$0.12/GB/mes

### Calculadora de Precios

Usa la [Calculadora de precios de Azure SQL Database](https://azure.microsoft.com/en-us/pricing/details/azure-sql-database/) para estimar costos según tu uso.

### Recomendaciones de Costo

- **Desarrollo:** Basic (5 DTUs) - ~$5/mes
- **Producción pequeña:** Standard S0 (10 DTUs) - ~$15/mes
- **Producción media:** Standard S1 (20 DTUs) - ~$30/mes
- **Alta disponibilidad:** Premium P1 (125 DTUs) - ~$465/mes

### Optimización de Costos

1. **Usar serverless:** Para cargas de trabajo intermitentes
2. **Pausar en desarrollo:** Pausar la base de datos cuando no se use
3. **Reserved capacity:** Ahorra hasta 33% con capacidad reservada (1-3 años)
4. **Monitor DTU usage:** Escala solo cuando sea necesario

---

## 🎯 Próximos Pasos

1. ✅ Base de datos desplegada en Azure SQL Database
2. ⏭️ Configurar variables de entorno en la aplicación
3. ⏭️ Conectar backend con la base de datos
4. ⏭️ Implementar autenticación
5. ⏭️ Desplegar aplicación completa
6. ⏭️ Configurar alertas y monitoreo
7. ⏭️ Configurar backups automáticos (ya está habilitado por defecto)

---

## 📚 Recursos Adicionales

- [Documentación de Azure SQL Database](https://docs.microsoft.com/en-us/azure/azure-sql/database/)
- [Guía de migración a Azure SQL Database](https://docs.microsoft.com/en-us/azure/azure-sql/database/migrate-to-database-from-sql-server)
- [Mejores prácticas de Azure SQL Database](https://docs.microsoft.com/en-us/azure/azure-sql/database/performance-guidance)
- [Azure SQL Database Pricing](https://azure.microsoft.com/en-us/pricing/details/azure-sql-database/)

---

¿Necesitas ayuda con alguno de estos pasos? Consulta la documentación adicional en el directorio `database/` o crea un issue en el repositorio.

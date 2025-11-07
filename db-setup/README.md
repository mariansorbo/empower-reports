# 🗄️ Azure SQL Database Setup

Scripts para configurar y conectar la base de datos Azure SQL de Empower Reports.

## 📋 Pre-requisitos

1. **Base de datos Azure SQL creada** ✅
   - Servidor: `empowerbi-server.database.windows.net`
   - Base de datos: `EmpowerBI-DB`

2. **Firewall configurado**
   - Ve a Azure Portal → SQL Server (`empowerbi-server`)
   - Settings → Networking
   - Agrega tu IP pública o activa "Allow Azure services and resources to access this server"

3. **Credenciales de acceso**
   - Usuario: `CloudSAe222b635`
   - Contraseña: (la que configuraste)

## 🚀 Instalación

1. **Instalar dependencias:**
   ```bash
   cd db-setup
   npm install
   ```

2. **Configurar variables de entorno:**
   ```bash
   cp .env.example .env
   ```
   
   Edita `.env` y agrega tu contraseña:
   ```env
   DB_PASSWORD=tu_contraseña_aquí
   ```

## 🧪 Paso 1: Probar Conexión

Antes de crear el schema, verifica que la conexión funcione:

```bash
npm run test-connection
```

**Salida esperada:**
```
✅ ¡Conexión exitosa!
📊 Información de la base de datos:
   Base de datos: EmpowerBI-DB
   Usuario actual: CloudSAe222b635
   Tablas existentes: 0
```

**Si falla:**
- ❌ Verifica la contraseña
- ❌ Verifica que tu IP esté en el firewall de Azure
- ❌ Verifica que el servidor esté activo (serverless puede pausarse)

## 🏗️ Paso 2: Crear Schema

Una vez que la conexión funcione, crea todas las tablas:

```bash
npm run create-schema
```

Este script ejecutará en orden:
1. `schema.sql` - Tablas principales
2. `organization_workflows.sql` - Procedures y funciones
3. `state_machine_and_workflows.sql` - Máquina de estados
4. `constraints_and_validations.sql` - Validaciones

**Salida esperada:**
```
✅ schema.sql ejecutado: 45 batches
✅ organization_workflows.sql ejecutado: 12 batches
✅ state_machine_and_workflows.sql ejecutado: 18 batches
✅ constraints_and_validations.sql ejecutado: 8 batches

📋 Tablas creadas:
   1. plans
   2. users
   3. organizations
   4. organization_members
   5. subscriptions
   6. subscription_history
   7. reports
   8. organization_documentation

🎉 ¡Schema creado exitosamente!
```

## 🔍 Verificar en Azure Portal

1. Ve a Azure Portal
2. Abre tu base de datos `EmpowerBI-DB`
3. Click en "Query editor (preview)"
4. Ingresa tus credenciales
5. Ejecuta:
   ```sql
   SELECT TABLE_NAME 
   FROM INFORMATION_SCHEMA.TABLES 
   WHERE TABLE_TYPE = 'BASE TABLE'
   ORDER BY TABLE_NAME;
   ```

## 📊 Estructura del Schema

### Tablas principales (8):
- `plans` - Planes de suscripción
- `users` - Usuarios del sistema
- `organizations` - Organizaciones
- `organization_members` - Miembros de organizaciones
- `organization_documentation` - URLs de documentación
- `subscriptions` - Suscripciones activas
- `subscription_history` - Historial de cambios
- `reports` - Reportes subidos

### Triggers automáticos:
- Actualización de timestamps (`updated_at`)
- Validación de límites de usuarios
- Validación de límites de reportes
- Validación de organización primaria única

### Procedures y funciones:
- Creación/unión a organizaciones
- Gestión de suscripciones
- Workflows de Enterprise Pro
- Validaciones de negocio

## 🛠️ Troubleshooting

### Error: "Login failed for user"
```bash
# Verifica las credenciales en .env
cat .env | grep PASSWORD
```

### Error: "Cannot open server"
```bash
# Verifica el firewall en Azure Portal
# Agrega tu IP pública actual
```

### Error: "Server is not available"
```bash
# Tu base de datos serverless puede estar pausada
# Azure la iniciará automáticamente (toma ~30 seg)
# Intenta de nuevo en un momento
```

### Ver tu IP pública
```bash
# Windows PowerShell
(Invoke-WebRequest -Uri "https://api.ipify.org").Content

# O visita: https://whatismyipaddress.com/
```

## 📝 Notas

- La base de datos es **serverless** (GP_S_Gen5_1)
- Se pausa después de 60 minutos de inactividad
- El primer request después de pausa toma ~30 segundos
- Capacidad mínima: 0.5 vCores
- Capacidad máxima: 1 vCore
- Almacenamiento máximo: 32 GB

## 🔄 Próximos pasos

Una vez creado el schema:
1. ✅ Crear backend API con Node.js/Express
2. ✅ Implementar autenticación con JWT
3. ✅ Conectar frontend a las APIs
4. ✅ Implementar gestión de organizaciones
5. ✅ Implementar sistema de suscripciones


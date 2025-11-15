# 🔧 Configuración de Azure Storage

Esta guía explica cómo configurar las credenciales de Azure Storage para que la aplicación pueda leer y gestionar los archivos .pbit.

## 📋 Requisitos

- Cuenta de Azure Storage: Configura tu cuenta de Azure Storage
- Container: `pbits-in`
- Credenciales de acceso (SAS Token o Connection String)

## ⚙️ Configuración Recomendada: SAS Token

**🔐 IMPORTANTE: Para producción, usa SAS Token en lugar de Connection String**

### Paso 1: Generar SAS Token

1. Ve al Azure Portal
2. Navega a tu Storage Account en Azure Portal
3. Ve a **Security + networking** → **Shared access signature**
4. Configura los permisos:
   - ✅ **Read** (r)
   - ✅ **Write** (w)
   - ✅ **Delete** (d)
   - ✅ **List** (l)
5. Allowed resource types:
   - ✅ **Container**
   - ✅ **Object**
6. Set expiration date (ej: 1 año)
7. Click **Generate SAS and connection string**
8. Copia el **SAS token** (empieza con `sv=...`)

### Paso 2: Configurar Variables de Entorno

Crea un archivo `.env` en la raíz del proyecto:

```bash
# Azure Storage Configuration (Frontend)
VITE_AZURE_ACCOUNT_NAME=
VITE_AZURE_SAS_TOKEN=
VITE_CONTAINER_NAME=pbits-in
VITE_APP_NAME=Report Tuner
VITE_MAX_FILE_SIZE=31457280
```

## ⚠️ Alternativa: Connection String (Solo para desarrollo)

**ADVERTENCIA:** El Connection String contiene tu Account Key completa, lo cual es un riesgo de seguridad si se expone en el frontend.

Si necesitas usar Connection String temporalmente, necesitas modificar el código:

### Archivo: `src/services/azureStorageService.js`

Reemplaza la función `getBlobServiceClient()`:

```javascript
const getBlobServiceClient = () => {
  // Opción 1: SAS Token (RECOMENDADO)
  if (accountName && sasToken) {
    const serviceUrl = `https://${accountName}.blob.core.windows.net?${sasToken}`
    return new BlobServiceClient(serviceUrl)
  }
  
  // Opción 2: Connection String (SOLO DESARROLLO)
  const connectionString = import.meta.env.VITE_AZURE_CONNECTION_STRING
  if (connectionString) {
    return BlobServiceClient.fromConnectionString(connectionString)
  }
  
  throw new Error('Missing Azure Storage configuration')
}
```

Luego en tu `.env`:

```bash
# ⚠️ SOLO PARA DESARROLLO LOCAL - NO SUBIR A GIT
VITE_AZURE_CONNECTION_STRING=DefaultEndpointsProtocol=https;AccountName=;AccountKey=;EndpointSuffix=core.windows.net
VITE_CONTAINER_NAME=pbits-in
VITE_APP_NAME=Report Tuner
VITE_MAX_FILE_SIZE=31457280
```

## 🚀 Probar la Configuración

1. Asegúrate de tener el archivo `.env` configurado
2. Reinicia el servidor de desarrollo:
   ```bash
   npm run dev
   ```
3. Abre la aplicación y ve a **Reports**
4. Deberías ver la lista de archivos .pbit del container

## 🔒 Seguridad

### ✅ Buenas Prácticas

- ✅ Usa SAS Token con permisos mínimos necesarios
- ✅ Configura fecha de expiración para el SAS Token
- ✅ NUNCA subas el archivo `.env` a Git (ya está en `.gitignore`)
- ✅ Usa variables de entorno del servidor para producción
- ✅ Rota las credenciales regularmente

### ❌ Evita

- ❌ Exponer Connection String en el frontend
- ❌ Subir credenciales a GitHub
- ❌ Dar permisos excesivos al SAS Token
- ❌ SAS Tokens sin fecha de expiración

## 📝 Funcionalidades Implementadas

### 1. **Listar Reportes**
- Lee todos los archivos `.pbit` del container
- Muestra nombre, fecha, tamaño
- Ordenados por fecha de modificación (más reciente primero)

### 2. **Eliminar Reportes**
- Selección múltiple con checkboxes
- Confirmación antes de eliminar
- Feedback visual del resultado

### 3. **Carga Automática**
- Los reportes se cargan automáticamente al abrir el modal
- Se actualizan después de eliminar archivos

## 🐛 Troubleshooting

### Error: "Missing Azure Storage configuration"
- Verifica que el archivo `.env` existe
- Verifica que las variables empiezan con `VITE_`
- Reinicia el servidor de desarrollo

### Error: "Failed to load reports"
- Verifica que las credenciales son correctas
- Verifica que el container `pbits-in` existe
- Verifica que el SAS Token tiene permiso de **List** y **Read**

### No se ven los archivos
- Verifica que hay archivos `.pbit` en el container
- Abre la consola del navegador para ver errores detallados
- Verifica la conexión a Internet

## 📚 Recursos

- [Azure Blob Storage SDK for JavaScript](https://docs.microsoft.com/en-us/javascript/api/@azure/storage-blob/)
- [SAS Token Documentation](https://docs.microsoft.com/en-us/azure/storage/common/storage-sas-overview)
- [Best Practices for SAS](https://docs.microsoft.com/en-us/azure/storage/common/storage-sas-best-practices)


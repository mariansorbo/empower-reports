# Report Tuner

Una aplicación web MVP para subir archivos .pbit a Azure Blob Storage.

## Características

- ✅ Subida de archivos .pbit a Azure Blob Storage
- ✅ Validación de tipo de archivo (.pbit únicamente)
- ✅ Validación de tamaño (máximo 30MB)
- ✅ Barra de progreso en tiempo real
- ✅ Interfaz drag & drop
- ✅ Mensajes de error y éxito claros
- ✅ Diseño responsive y moderno

## Configuración

1. **Instalar dependencias:**
   ```bash
   npm install
   ```

2. **Configurar variables de entorno:**
   - Copia `env.example` a `.env`
   - Actualiza las credenciales de Azure Storage

3. **Ejecutar en desarrollo:**
   ```bash
   npm run dev
   ```

4. **Construir para producción:**
   ```bash
   npm run build
   ```

## Variables de Entorno

```env
VITE_AZURE_STORAGE_CONNECTION_STRING=tu_connection_string_aqui
VITE_CONTAINER_NAME=pbits-in
VITE_APP_NAME=Report Tuner
VITE_MAX_FILE_SIZE=31457280
```

## Uso

1. Abre la aplicación en tu navegador
2. Arrastra un archivo .pbit o haz clic para seleccionarlo
3. Haz clic en "Subir Archivo"
4. Observa el progreso de carga
5. Recibe confirmación de éxito

## Tecnologías

- React 18
- Vite
- Azure Blob Storage SDK
- CSS3 con diseño moderno

## Estructura del Proyecto

```
src/
├── components/
│   ├── FileUpload.jsx
│   └── FileUpload.css
├── App.jsx
├── App.css
└── main.jsx
```

# Configuration Guide - Azure Storage & EmailJS

## 🔧 Configuration Required

The application needs two configurations:
1. **Azure Storage** - Para leer los reportes .pbit
2. **EmailJS** - Para enviar correos del formulario de contacto

## 📝 Setup Instructions

### 1. Create Environment File

Create a file named `.env.local` in the root directory with the following content:

```env
# ===== Azure Storage Configuration =====
# NOTA: Este proyecto tiene configurada la conexión a procesadorastorage
VITE_AZURE_CONNECTION_STRING=DefaultEndpointsProtocol=https;AccountName=YOUR_ACCOUNT_NAME;AccountKey=YOUR_ACCOUNT_KEY;EndpointSuffix=core.windows.net
VITE_CONTAINER_NAME=pbits

# ===== EmailJS Configuration =====
# Obtén estas credenciales de https://www.emailjs.com/
VITE_EMAILJS_SERVICE_ID=your_service_id_here
VITE_EMAILJS_TEMPLATE_ID=your_template_id_here
VITE_EMAILJS_PUBLIC_KEY=your_public_key_here
```

### 2. Alternative Azure Configuration (More Secure for Production)

Instead of using the connection string, you can use Account Name + SAS Token:

```env
VITE_AZURE_ACCOUNT_NAME=YOUR_ACCOUNT_NAME
VITE_AZURE_SAS_TOKEN=your-sas-token-here
VITE_CONTAINER_NAME=pbits
```

---

## 📧 EmailJS Setup (Para el Formulario de Contacto)

### Paso 1: Obtener las Credenciales

#### 🔑 Service ID
1. En EmailJS, ve al menú izquierdo → **"Email Services"**
2. Haz click en tu servicio de Gmail/correo
3. Verás el **Service ID** (ej: `service_abc123`)
4. Cópialo a tu `.env.local`

#### 📋 Template ID
1. En EmailJS, asegúrate de estar en la página de tu template "Contact Us"
2. Mira la **URL del navegador** - tendrá: `emailjs.com/.../templates/template_XXXXXXX`
3. O ve a la pestaña **"Settings"** y verás el **Template ID**
4. Cópialo a tu `.env.local`

#### 🔐 Public Key
1. En EmailJS, ve al menú izquierdo → **"Account"**
2. Busca la sección **"API Keys"** o **"General"**
3. Verás tu **Public Key** (una cadena alfanumérica)
4. Cópialo a tu `.env.local`

### Paso 2: Verifica el Template en EmailJS

Tu template debe tener estas variables configuradas:
- **Subject**: `Contact Us: {{title}}`
- **To Email**: `mariansorbo@gmail.com` ✅ (ya lo tienes)
- **From Name**: `{{from_name}}` ✅
- **Reply To**: `{{email}}` ✅
- **Content**: Debe incluir `{{from_name}}`, `{{email}}`, `{{message}}`

### Paso 3: Ejemplo de .env.local Completo

```env
# Azure Storage
VITE_AZURE_CONNECTION_STRING=DefaultEndpointsProtocol=https;AccountName=YOUR_ACCOUNT_NAME;AccountKey=YOUR_ACCOUNT_KEY;EndpointSuffix=core.windows.net
VITE_CONTAINER_NAME=pbits

# EmailJS (reemplaza con tus valores reales)
VITE_EMAILJS_SERVICE_ID=service_abc123
VITE_EMAILJS_TEMPLATE_ID=template_xyz789
VITE_EMAILJS_PUBLIC_KEY=aBcDeFgHiJkLmNoPqRsTuV
```

## Changes Made

### Updated `src/services/azureStorageService.js`

- ✅ Removed the hardcoded "Ejemplo 1/" folder prefix
- ✅ Now reads all .pbit files from the root of the "pbits" container
- ✅ Displays only the filename (not the full path with folders)
- ✅ Sorts reports by last modified date (most recent first)

### How It Works

1. The `ReportsModal` component calls `listReports()` from the Azure Storage service
2. The service connects to your Azure Storage account using the credentials from `.env.local`
3. It lists all files with `.pbit` extension in the "pbits" container
4. Files are displayed with their name, date, and size

## Container Structure

The application will read .pbit files from:
- **Storage Account**: `YOUR_ACCOUNT_NAME`
- **Container**: `pbits`
- **Files**: All `.pbit` files in the container (including subfolders)

## Testing

After creating the `.env.local` file:

1. Restart your development server:
   ```bash
   npm run dev
   ```

2. Open the application in your browser
3. Click on "📋 View Reports" button
4. You should see all .pbit files from your Azure Storage container

## Troubleshooting

### "Failed to load reports" Error

- Verify your connection string is correct
- Check that the container name is "pbits"
- Ensure the storage account has public access or proper CORS settings
- Verify the account key hasn't expired

### No Reports Showing

- Check if there are actually .pbit files in the "pbits" container
- Verify the container name is spelled correctly
- Check browser console for detailed error messages

### CORS Issues

If you get CORS errors, you need to configure CORS in your Azure Storage account:

1. Go to Azure Portal → Your Storage Account
2. Navigate to "Resource sharing (CORS)"
3. Add these settings:
   - Allowed origins: `http://localhost:5173` (or your domain)
   - Allowed methods: `GET, PUT, POST, DELETE, HEAD, OPTIONS`
   - Allowed headers: `*`
   - Exposed headers: `*`
   - Max age: `3600`

## Security Notes

- ⚠️ Never commit `.env.local` to git (it's already in `.gitignore`)
- ⚠️ The connection string contains sensitive credentials
- 🔒 For production, use SAS Token with limited permissions instead of Account Key
- 🔒 Rotate your keys periodically

## Production Deployment

For Docker deployment, you can pass environment variables using:

```bash
docker run -d \
  --name empower-reports-app \
  -p 80:80 \
  -e VITE_AZURE_CONNECTION_STRING="your-connection-string" \
  -e VITE_CONTAINER_NAME="pbits" \
  --restart unless-stopped \
  gimzalo/empower-reports:latest
```

Or use build arguments:

```bash
docker build \
  --build-arg VITE_AZURE_CONNECTION_STRING="your-connection-string" \
  --build-arg VITE_CONTAINER_NAME="pbits" \
  -t empower-reports:custom .
```


# 🚀 Guía de Configuración Rápida - Report Tuner

Esta guía te ayudará a configurar las variables de entorno necesarias para que la aplicación funcione correctamente.

## ✅ Checklist de Configuración

- [ ] 1. Configurar Azure Storage (leer reportes)
- [ ] 2. Configurar EmailJS (enviar correos de contacto)
- [ ] 3. Crear archivo `.env.local`
- [ ] 4. Reiniciar servidor de desarrollo

---

## 📋 Paso 1: Obtener Credenciales de EmailJS

### A. Service ID

1. Ve a [EmailJS Dashboard](https://dashboard.emailjs.com/)
2. En el menú izquierdo, haz click en **"Email Services"**
3. Haz click en tu servicio de Gmail (o el que hayas configurado)
4. Copia el **Service ID** (algo como `service_abc123`)

### B. Template ID

**Opción 1 - Desde la URL:**
1. Ve a tu template "Contact Us"
2. Mira la URL del navegador: `emailjs.com/.../templates/template_XXXXXXX`
3. El Template ID es: `template_XXXXXXX`

**Opción 2 - Desde Settings:**
1. Estando en tu template, haz click en la pestaña **"Settings"**
2. Verás el **Template ID**

### C. Public Key

1. En el menú izquierdo, haz click en **"Account"**
2. Busca la sección **"API Keys"** o en el tab **"General"**
3. Verás tu **Public Key** (una cadena alfanumérica)
4. Cópiala

---

## 📝 Paso 2: Crear Archivo de Configuración

### Opción A: Usar el Script Automático (PowerShell)

```powershell
.\setup-env.ps1
```

El script te pedirá las 3 credenciales y creará el archivo `.env.local` automáticamente.

### Opción B: Crear Manualmente

Crea un archivo llamado `.env.local` en la raíz del proyecto con este contenido:

```env
# ===== Azure Storage Configuration =====
VITE_AZURE_CONNECTION_STRING=DefaultEndpointsProtocol=https;AccountName=;AccountKey=;EndpointSuffix=core.windows.net
VITE_CONTAINER_NAME=pbits

# ===== EmailJS Configuration =====
VITE_EMAILJS_SERVICE_ID=TU_SERVICE_ID_AQUI
VITE_EMAILJS_TEMPLATE_ID=TU_TEMPLATE_ID_AQUI
VITE_EMAILJS_PUBLIC_KEY=TU_PUBLIC_KEY_AQUI
```

**⚠️ Importante**: Reemplaza `TU_SERVICE_ID_AQUI`, `TU_TEMPLATE_ID_AQUI`, y `TU_PUBLIC_KEY_AQUI` con tus valores reales.

---

## 🔧 Paso 3: Verificar Template en EmailJS

Asegúrate de que tu template "Contact Us" tenga estas configuraciones:

### Variables del Template

El template debe usar estas variables (ya configurado en tu captura):
- `{{from_name}}` - Nombre del usuario
- `{{email}}` - Email del usuario (para Reply-To)
- `{{message}}` - Mensaje/feedback del usuario
- `{{title}}` - Título (automático: "New Feedback")
- `{{time}}` - Timestamp (automático)

### Configuración de Email

- ✅ **To Email**: `mariansorbo@gmail.com` (ya configurado)
- ✅ **From Name**: `{{from_name}}`
- ✅ **Reply To**: `{{email}}`

---

## 🎯 Paso 4: Probar la Configuración

### 1. Reiniciar el servidor

```bash
# Detén el servidor si está corriendo (Ctrl+C)
# Luego inicia nuevamente:
npm run dev
```

### 2. Probar el formulario de contacto

1. Abre la aplicación en el navegador (http://localhost:5173)
2. Scroll hasta la sección de "Contacto"
3. Llena el formulario con datos de prueba
4. Haz click en "Send Feedback"
5. Deberías ver un mensaje de éxito: ✅ "Thank you! Your message has been sent successfully."
6. Revisa tu correo `mariansorbo@gmail.com` - debería llegar el mensaje

### 3. Verificar Azure Storage (Reportes)

1. En la aplicación, haz click en "📋 View Reports"
2. Deberías ver la lista de archivos .pbit del container "pbits"
3. Si ves un error, verifica la connection string en `.env.local`

---

## 🐛 Troubleshooting

### ❌ Error: "EmailJS credentials not configured"

**Solución**: 
- Verifica que el archivo `.env.local` existe en la raíz del proyecto
- Verifica que las 3 variables de EmailJS estén presentes y sin comillas
- Reinicia el servidor (`npm run dev`)

### ❌ Error: "Failed to load reports"

**Solución**:
- Verifica que `VITE_AZURE_CONNECTION_STRING` esté correctamente configurado
- Verifica que el container se llame "pbits"
- Verifica que el storage account tenga archivos .pbit

### ❌ El correo no llega

**Solución**:
1. Verifica en [EmailJS Dashboard](https://dashboard.emailjs.com/) → "Email History"
2. Busca si el email fue enviado
3. Si dice "failed", revisa:
   - Que el Service ID sea correcto
   - Que el servicio de Gmail esté activo
   - Que no hayas excedido el límite (200/mes en plan gratuito)
4. Revisa la carpeta de spam en `mariansorbo@gmail.com`

### 🔍 Ver logs de errores

Abre la consola del navegador (F12) y busca mensajes de error en rojo.

---

## 📚 Documentación Adicional

- [Documentación de EmailJS](https://www.emailjs.com/docs/)
- [Configuración de Azure Storage](./AZURE_STORAGE_CONFIG.md)
- [Deployment Guide](./VPS_DEPLOYMENT_GUIDE.md)

---

## ✅ Todo Listo!

Una vez completados estos pasos, tu aplicación estará completamente configurada:

- ✅ Leer reportes .pbit desde Azure Storage
- ✅ Enviar correos de contacto a mariansorbo@gmail.com
- ✅ Formulario funcional con feedback al usuario

---

## 🔒 Seguridad

**⚠️ Importante**:
- El archivo `.env.local` está en `.gitignore` - nunca se subirá a Git
- No compartas tus credenciales públicamente
- Para producción, considera usar variables de entorno del servidor en lugar de `.env.local`

**Respuesta a tu pregunta: "¿Desde qué mail se envía?"**

El correo se envía desde la cuenta de Gmail (o el servicio) que conectaste en EmailJS → Email Services. Por ejemplo:
- Si conectaste `tu-cuenta@gmail.com` en EmailJS
- Los correos se enviarán desde: `tu-cuenta@gmail.com`
- Llegarán a: `mariansorbo@gmail.com`
- El usuario puede responder directamente (Reply-To está configurado con su email)

---

¡Listo! Si tienes problemas, revisa la sección de Troubleshooting arriba. 🚀





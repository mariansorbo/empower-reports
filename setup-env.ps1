# Script para configurar variables de entorno
# Ejecuta este script para crear tu archivo .env.local

Write-Host "==================================" -ForegroundColor Cyan
Write-Host "  Report Tuner - ENV Setup" -ForegroundColor Cyan
Write-Host "==================================" -ForegroundColor Cyan
Write-Host ""

$envContent = @"
# ===== Azure Storage Configuration =====
VITE_AZURE_CONNECTION_STRING=DefaultEndpointsProtocol=https;AccountName=YOUR_ACCOUNT_NAME;AccountKey=YOUR_ACCOUNT_KEY;EndpointSuffix=core.windows.net
VITE_CONTAINER_NAME=pbits

# ===== EmailJS Configuration =====
# Obtén estas credenciales de https://www.emailjs.com/
"@

Write-Host "Por favor, ingresa tus credenciales de EmailJS:" -ForegroundColor Yellow
Write-Host ""

# Pedir Service ID
Write-Host "1. Service ID (Ejemplo: service_abc123)" -ForegroundColor Green
$serviceId = Read-Host "   Service ID"
$envContent += "`nVITE_EMAILJS_SERVICE_ID=$serviceId"

# Pedir Template ID
Write-Host ""
Write-Host "2. Template ID (Ejemplo: template_xyz789)" -ForegroundColor Green
$templateId = Read-Host "   Template ID"
$envContent += "`nVITE_EMAILJS_TEMPLATE_ID=$templateId"

# Pedir Public Key
Write-Host ""
Write-Host "3. Public Key (cadena alfanumérica)" -ForegroundColor Green
$publicKey = Read-Host "   Public Key"
$envContent += "`nVITE_EMAILJS_PUBLIC_KEY=$publicKey"

# Guardar archivo
try {
    $envContent | Out-File -FilePath ".env.local" -Encoding UTF8 -NoNewline
    Write-Host ""
    Write-Host "✅ Archivo .env.local creado exitosamente!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Ahora puedes ejecutar:" -ForegroundColor Cyan
    Write-Host "  npm run dev" -ForegroundColor White
    Write-Host ""
} catch {
    Write-Host ""
    Write-Host "❌ Error al crear el archivo: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Crea manualmente el archivo .env.local con este contenido:" -ForegroundColor Yellow
    Write-Host $envContent -ForegroundColor White
}





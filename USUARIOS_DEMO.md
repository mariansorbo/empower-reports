# Usuarios de Demostración

## Archivo CSV de Usuarios

El sistema de autenticación utiliza un archivo CSV ubicado en `public/data/usuarios.csv` para almacenar las credenciales de los usuarios.

### Estructura del CSV

```csv
email,password
admin@reporttuner.com,admin123
demo@reporttuner.com,demo123
```

### Usuarios Predefinidos

| Email | Contraseña | Descripción |
|-------|------------|-------------|
| admin@reporttuner.com | admin123 | Usuario administrador |
| demo@reporttuner.com | demo123 | Usuario de demostración |

## Cómo Funciona

1. **Login**: Las credenciales ingresadas se validan contra los usuarios almacenados
2. **Registro**: Los nuevos usuarios se agregan y persisten en localStorage
3. **Validación**: Se verifica que el email no esté duplicado
4. **Persistencia**: 
   - El estado de login se mantiene en localStorage
   - Los usuarios se almacenan en localStorage (simulando el CSV)
   - Al recargar la página, los usuarios persisten

## Notas Técnicas

- Los usuarios se almacenan en localStorage del navegador (simulando persistencia del CSV)
- El archivo CSV inicial se carga solo la primera vez
- En un entorno de producción, esto sería reemplazado por una API REST con base de datos
- Las contraseñas se almacenan en texto plano (solo para demo - en producción usar hash)
- El sistema maneja errores de red y fallbacks apropiados

## Funciones de Debugging

Para debugging, puedes usar estas funciones en la consola del navegador:

```javascript
// Ver todos los usuarios
import { getAllUsers } from './src/services/userService.js'
getAllUsers().then(users => console.log(users))

// Limpiar todos los usuarios
import { clearUsers } from './src/services/userService.js'
clearUsers()

// Reiniciar a usuarios por defecto
import { resetUsers } from './src/services/userService.js'
resetUsers()
```

## Agregar Nuevos Usuarios

Los nuevos usuarios se agregan automáticamente cuando se registran. Para agregar usuarios manualmente, edita el archivo `public/data/usuarios.csv` y reinicia la aplicación.

## Probar el Sistema

1. Inicia la aplicación
2. Haz clic en "Iniciar sesión"
3. Usa cualquiera de las credenciales de la tabla anterior
4. O crea una nueva cuenta usando el formulario de registro

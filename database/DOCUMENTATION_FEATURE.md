# Feature: Documentación por Organización

## 🎯 Objetivo

Permitir que cada organización tenga un link personalizado a su documentación. El botón "Ver documentación" en el frontend se habilita solo cuando la organización tiene una URL configurada.

## 📊 Tabla: `organization_documentation`

### Estructura

```sql
CREATE TABLE organization_documentation (
    id UNIQUEIDENTIFIER PRIMARY KEY,
    organization_id UNIQUEIDENTIFIER NOT NULL UNIQUE,  -- Una org = una URL
    documentation_url VARCHAR(500) NOT NULL,
    description TEXT NULL,
    is_active BIT NOT NULL DEFAULT 1,
    created_by UNIQUEIDENTIFIER NULL,
    created_at DATETIME2 NOT NULL,
    updated_at DATETIME2 NOT NULL
);
```

### Características

- **Una organización = una URL**: Constraint UNIQUE en `organization_id`
- **Cascada**: Si se elimina la organización, se elimina su documentación
- **Auditoría**: Guarda quién configuró la URL
- **Activación/Desactivación**: Campo `is_active` para deshabilitar sin eliminar

---

## 🔧 Stored Procedures

### 1. Establecer/Actualizar URL

```sql
EXEC sp_set_organization_documentation
    @organization_id = '<GUID>',
    @documentation_url = 'https://docs.miempresa.com/power-bi',
    @description = 'Documentación de reportes Power BI',
    @created_by = '<user_id>';
```

**¿Qué hace?**
- Si no existe: Crea nuevo registro
- Si existe: Actualiza URL y descripción
- Valida que el usuario sea admin

**Validaciones:**
- Organización existe y no está archivada
- Usuario es admin o admin_global de la organización

### 2. Desactivar URL

```sql
EXEC sp_disable_organization_documentation
    @organization_id = '<GUID>';
```

**¿Qué hace?**
- Marca `is_active = 0`
- No elimina el registro, solo lo desactiva

### 3. Eliminar URL

```sql
EXEC sp_remove_organization_documentation
    @organization_id = '<GUID>';
```

**¿Qué hace?**
- Elimina el registro completamente

---

## 🔍 Función

### `fn_get_organization_documentation_url`

```sql
SELECT dbo.fn_get_organization_documentation_url('<organization_id>');
-- Retorna: 'https://docs.miempresa.com' o NULL si no tiene
```

**Uso en frontend:**
```javascript
const docUrl = await db.query(
  'SELECT dbo.fn_get_organization_documentation_url(@org_id)',
  { org_id: currentOrganization.id }
);

if (docUrl) {
  // Habilitar botón amarillo
  setDocumentationButtonEnabled(true);
  setDocumentationUrl(docUrl);
} else {
  // Mostrar botón gris deshabilitado
  setDocumentationButtonEnabled(false);
}
```

---

## 🎨 Implementación Frontend

### 1. Componente del Botón

```jsx
import { useState, useEffect } from 'react';
import { useOrganization } from '../contexts/OrganizationContext';

export function DocumentationButton() {
  const { currentOrganization } = useOrganization();
  const [docUrl, setDocUrl] = useState(null);
  const [loading, setLoading] = useState(true);
  
  useEffect(() => {
    if (currentOrganization?.id) {
      fetchDocumentationUrl();
    }
  }, [currentOrganization]);
  
  const fetchDocumentationUrl = async () => {
    try {
      const response = await fetch(`/api/organizations/${currentOrganization.id}/documentation`);
      const data = await response.json();
      setDocUrl(data.documentation_url);
    } catch (error) {
      console.error('Error fetching documentation URL:', error);
    } finally {
      setLoading(false);
    }
  };
  
  const handleClick = () => {
    if (docUrl) {
      window.open(docUrl, '_blank');
    }
  };
  
  return (
    <button
      className={`btn ${docUrl ? 'btn-primary' : 'btn-disabled'}`}
      onClick={handleClick}
      disabled={!docUrl}
      title={docUrl ? 'Ver documentación' : 'Sin documentación configurada'}
    >
      📚 Ver Documentación
    </button>
  );
}
```

### 2. CSS para el botón

```css
.btn-primary {
  background-color: #F3C911; /* Amarillo */
  color: #000;
  cursor: pointer;
}

.btn-disabled {
  background-color: #ccc; /* Gris */
  color: #666;
  cursor: not-allowed;
}
```

### 3. Actualizar vista de usuario

La vista `vw_users_with_primary_org` ya incluye:
- `organization_documentation_url` - La URL
- `has_documentation` - 1 si tiene, 0 si no

```sql
SELECT * FROM vw_users_with_primary_org WHERE id = @user_id;
-- Retorna: ..., organization_documentation_url, has_documentation, ...
```

---

## 📡 API Backend

### GET /api/organizations/:id/documentation

```javascript
router.get('/api/organizations/:id/documentation', authenticate, async (req, res) => {
  const orgId = req.params.id;
  
  // Verificar que el usuario pertenece a la organización
  const isMember = await db.query(`
    SELECT 1 FROM organization_members
    WHERE organization_id = @org_id AND user_id = @user_id AND left_at IS NULL
  `, { org_id: orgId, user_id: req.user.id });
  
  if (!isMember) {
    return res.status(403).json({ error: 'No perteneces a esta organización' });
  }
  
  // Obtener URL
  const url = await db.query(
    'SELECT dbo.fn_get_organization_documentation_url(@org_id) AS url',
    { org_id: orgId }
  );
  
  res.json({
    documentation_url: url?.url || null,
    has_documentation: !!url?.url
  });
});
```

### POST /api/organizations/:id/documentation (Admin only)

```javascript
router.post('/api/organizations/:id/documentation', authenticate, async (req, res) => {
  const orgId = req.params.id;
  const { documentation_url, description } = req.body;
  
  const result = await db.execute('sp_set_organization_documentation', {
    organization_id: orgId,
    documentation_url,
    description,
    created_by: req.user.id
  });
  
  if (result.success) {
    res.json({ success: true, message: result.message });
  } else {
    res.status(400).json({ success: false, error: result.message });
  }
});
```

---

## 🎯 Casos de Uso

### Caso 1: Organización sin documentación

```
Usuario logueado → Botón "Ver documentación" gris y deshabilitado
```

### Caso 2: Organización con documentación

```
Usuario logueado → Botón "Ver documentación" amarillo y habilitado
Click en botón → Abre URL en nueva pestaña
```

### Caso 3: Admin configura documentación

```
Admin va a Settings → Organización → Configurar documentación
Ingresa URL: https://docs.miempresa.com
Click en "Guardar" → Backend llama sp_set_organization_documentation
Success → Botón se habilita para todos los miembros
```

### Caso 4: Admin actualiza URL

```
Admin cambia URL: https://docs.miempresa.com/v2
Click en "Guardar" → sp_set_organization_documentation actualiza el registro
Success → Todos los miembros ven la nueva URL
```

---

## 📋 Datos Dummy

```sql
-- Ejemplo de datos
INSERT INTO organization_documentation (organization_id, documentation_url, description, created_by)
VALUES
    ('<citenza_org_id>', 'https://docs.citenza.com/power-bi', 'Documentación de Power BI para Citenza', '<gonzalo_user_id>'),
    ('<data_latam_org_id>', 'https://datalatam.notion.site/docs', 'Wiki de documentación Data LATAM', '<camila_user_id>');
```

---

## 🎯 Flujo Completo

```
1. Admin va a Settings → Organización
   ↓
2. Sección "Documentación"
   ├─ Input: URL de documentación
   ├─ Textarea: Descripción (opcional)
   └─ Botón: "Guardar"
   ↓
3. Click en "Guardar"
   ↓
4. Frontend llama POST /api/organizations/:id/documentation
   ↓
5. Backend llama sp_set_organization_documentation
   ├─ Valida: Usuario es admin
   ├─ Valida: Organización activa
   └─ INSERT o UPDATE
   ↓
6. Success → Frontend muestra mensaje "URL guardada"
   ↓
7. Header se actualiza → Botón "Ver documentación" se habilita (amarillo)
   ↓
8. Miembros de la org ven el botón habilitado
```

---

## 🔄 Integración con Vista de Usuario

```sql
-- En el login, el frontend obtiene
SELECT * FROM vw_users_with_primary_org WHERE id = @user_id;

-- Retorna:
-- ..., organization_documentation_url, has_documentation, ...

-- Frontend usa:
if (user.has_documentation) {
  enableDocumentationButton(user.organization_documentation_url);
}
```

---

## 📊 Resumen

- **Tabla nueva**: `organization_documentation` (1:1 con organizations)
- **Trigger nuevo**: `trg_org_documentation_updated_at`
- **Procedures nuevos**: 3 (set, disable, remove)
- **Función nueva**: `fn_get_organization_documentation_url`
- **Vista actualizada**: `vw_users_with_primary_org` incluye URL y flag
- **UX**: Botón gris (sin URL) o amarillo (con URL)

**Total agregado:**
- 1 tabla
- 1 trigger
- 3 procedures
- 1 función
- 1 vista actualizada







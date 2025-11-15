# 🏢 Enterprise Pro Plan V2 - Documentación

## 📋 Descripción

El plan **Enterprise Pro** está diseñado para empresas de consultoría de analytics que trabajan con múltiples clientes y necesitan:

1. **Separación completa de metadata** entre clientes
2. **Gestión centralizada** por un admin global
3. **Múltiples organizaciones independientes** (hasta 5) para diferentes clientes
4. **Usuarios que pueden trabajar en múltiples organizaciones** sin ver metadata de otras

## 🎯 Caso de Uso

**Empresa de consultoría de analytics:**
- Trabaja con 3-5 clientes diferentes
- Cada cliente tiene su propia organización **independiente** (confidencialidad)
- Algunos consultores trabajan en múltiples clientes
- Los consultores NO deben ver metadata de reportes de clientes en los que no trabajan
- El admin global de la consultoría gestiona todas las organizaciones desde un panel central

## 🏗️ Modelo: Organizaciones Independientes

### **Concepto Clave**

**NO hay jerarquía padre/hijo.** Las organizaciones son completamente independientes. Lo único que comparten es que un **admin global** las gestiona.

```
┌─────────────────────────────────────┐
│  Organización Enterprise Pro         │
│  (Consultoría de Analytics)        │
│                                     │
│  Admin Global: Juan Pérez          │
│  Plan: Enterprise Pro               │
│  Puede gestionar: hasta 5 orgs     │
└─────────────────────────────────────┘
         │ gestiona (no jerarquía)
         │
    ┌────┼────┬─────┬─────┬─────┐
    │    │    │     │     │     │
┌───▼─┐ ┌▼──┐ ┌▼──┐ ┌▼──┐ ┌▼──┐
│Org A│ │Org B│ │Org C│ │Org D│ │Org E│
│     │ │    │ │    │ │    │ │    │
│Users:│ │Users:│ │Users:│ │Users:│ │Users:│
│- Ana │ │- Ana │ │- Bob │ │- Bob │ │- Ana │
│- Bob │ │- Carlos│ │- Carlos│ │- Carlos│ │- Bob │
└─────┘ └─────┘ └─────┘ └─────┘ └─────┘

Todas son organizaciones independientes
Solo comparten que el mismo admin_global las gestiona
```

## 🔑 Componentes Clave

### **1. Rol `admin_global`**

Nuevo rol en `organization_members`:
- **`admin_global`**: Puede gestionar múltiples organizaciones (solo en Enterprise Pro)
- **`admin`**: Admin normal de una organización específica
- **`member`**: Miembro regular
- **`viewer`**: Solo lectura

### **2. Tabla `enterprise_pro_managed_organizations`**

Relaciona organizaciones Enterprise Pro con las organizaciones que gestionan:

```sql
enterprise_pro_org_id     → Organización con plan Enterprise Pro
managed_organization_id   → Organización gestionada (independiente)
admin_user_id            → Usuario admin_global que gestiona
```

**Relación**: Una organización puede ser gestionada por UNA Enterprise Pro, pero las organizaciones son independientes.

### **3. Organizaciones Independientes**

- Cada organización tiene su propio `id`, `name`, `slug`
- Cada organización tiene su propia suscripción (inicialmente `free_trial`)
- Cada organización tiene sus propios usuarios y reportes
- **NO hay campo `parent_organization_id`** - son completamente independientes

## 📊 Características del Plan

### **Límites**
- **Máximo 5 organizaciones gestionadas** por organización Enterprise Pro
- **50 usuarios** en la organización Enterprise Pro
- **1000 reportes** totales (acumulados entre todas las organizaciones gestionadas)
- **200GB de almacenamiento** total

### **Features**
- ✅ API access
- ✅ Branding personalizado
- ✅ Audit log completo
- ✅ Priority support
- ✅ **Multi-organization management** (nuevo)
- ✅ **Organization isolation** (nuevo)
- ✅ **Advanced user management** (nuevo)
- ✅ **Global admin role** (nuevo)

## 🔐 Separación de Metadata

### **Reportes**
Cada reporte está vinculado a una `organization_id` específica:
```sql
-- Reporte del Cliente A
INSERT INTO reports (organization_id, user_id, name, ...)
VALUES ('org-cliente-a', 'user-ana', 'Reporte Ventas Cliente A', ...);

-- Reporte del Cliente B (ANA NO puede verlo aunque pertenezca a Cliente A)
INSERT INTO reports (organization_id, user_id, name, ...)
VALUES ('org-cliente-b', 'user-ana', 'Reporte Ventas Cliente B', ...);
```

### **Acceso**
- Ana pertenece a Cliente A → Solo ve reportes de Cliente A
- Ana también pertenece a Cliente B → Ve reportes de Cliente A Y Cliente B
- Carlos NO pertenece a Cliente A → NO ve reportes de Cliente A

## 📝 Uso de la API

### **1. Asignar rol admin_global a un usuario**

```sql
-- Juan Pérez es admin_global de la organización Enterprise Pro
INSERT INTO organization_members (
    organization_id,
    user_id,
    role,
    is_primary
)
VALUES (
    'org-enterprise-pro-123',
    'user-juan-id',
    'admin_global',  -- NUEVO ROL
    1  -- Es su organización primaria
);
```

### **2. Crear Organización Gestionada**

```sql
DECLARE @new_org_id UNIQUEIDENTIFIER;
DECLARE @message VARCHAR(500);

EXEC sp_create_managed_organization
    @enterprise_pro_org_id = 'org-enterprise-pro-123',
    @organization_name = 'Cliente ABC Corp',
    @organization_slug = 'cliente-abc',
    @created_by_user_id = 'user-juan-id',  -- Debe ser admin_global
    @organization_id = @new_org_id OUTPUT,
    @message = @message OUTPUT;

SELECT @new_org_id AS new_organization_id, @message AS message;
```

### **3. Verificar si puede gestionar más organizaciones**

```sql
SELECT dbo.fn_can_manage_more_organizations('org-enterprise-pro-123');
-- Retorna 1 si puede, 0 si no puede (límite alcanzado)
```

### **4. Contar organizaciones gestionadas**

```sql
SELECT dbo.fn_get_managed_organizations_count('org-enterprise-pro-123');
-- Retorna número de organizaciones gestionadas activas
```

### **5. Ver todas las organizaciones Enterprise Pro**

```sql
SELECT * FROM vw_enterprise_pro_organizations;
-- Muestra: enterprise_pro_org, current_managed_orgs, remaining_slots, etc.
```

### **6. Ver organizaciones gestionadas**

```sql
SELECT * FROM vw_managed_organizations;
-- Muestra: managed_org, enterprise_pro_org, admin_user, member_count, reports_count
```

### **7. Obtener organizaciones gestionadas por un usuario**

```sql
SELECT * FROM dbo.fn_get_user_managed_organizations('user-juan-id');
-- Retorna todas las organizaciones que Juan gestiona como admin_global
```

### **8. Verificar si usuario puede gestionar una organización**

```sql
SELECT dbo.fn_can_user_manage_organization('user-juan-id', 'org-cliente-a');
-- Retorna 1 si puede gestionar (es admin_global o admin), 0 si no
```

## 🔄 Flujo de Trabajo

### **Setup Inicial**

1. **Cliente se suscribe a Enterprise Pro**
   ```sql
   -- Crear organización Enterprise Pro
   INSERT INTO organizations (name, ...) VALUES ('Consultoría Analytics Pro', ...);
   
   -- Crear suscripción Enterprise Pro
   INSERT INTO subscriptions (organization_id, plan_id, status, ...)
   VALUES ('org-enterprise-pro', 'enterprise_pro', 'active', ...);
   ```

2. **Asignar admin_global**
   ```sql
   INSERT INTO organization_members (organization_id, user_id, role, ...)
   VALUES ('org-enterprise-pro', 'user-juan-id', 'admin_global', ...);
   ```

3. **Admin global crea organización para Cliente A**
   ```sql
   EXEC sp_create_managed_organization
       @enterprise_pro_org_id = 'org-enterprise-pro',
       @organization_name = 'Cliente A',
       @created_by_user_id = 'user-juan-id',
       ...
   ```

4. **Admin global invita usuarios a Cliente A**
   ```sql
   -- Ana se une a Cliente A como member
   INSERT INTO organization_members (organization_id, user_id, role, ...)
   VALUES ('org-cliente-a', 'user-ana', 'member', ...);
   ```

5. **Ana sube reportes a Cliente A**
   ```sql
   -- Reporte vinculado SOLO a Cliente A
   INSERT INTO reports (organization_id, user_id, name, ...)
   VALUES ('org-cliente-a', 'user-ana', 'Reporte Ventas', ...);
   ```

### **Usuario Multi-Organización**

1. **Ana también trabaja en Cliente B**
   ```sql
   -- Ana se une a Cliente B
   INSERT INTO organization_members (organization_id, user_id, role, ...)
   VALUES ('org-cliente-b', 'user-ana', 'member', ...);
   ```

2. **Ana ahora ve reportes de Cliente A Y Cliente B**
   ```sql
   -- Query para obtener reportes de Ana
   SELECT * FROM reports 
   WHERE user_id = 'user-ana'
   AND organization_id IN (
       SELECT organization_id 
       FROM organization_members 
       WHERE user_id = 'user-ana' 
       AND left_at IS NULL
   )
   AND is_deleted = 0;
   ```

3. **Ana NO ve reportes de Cliente C** (donde no pertenece)

## 🔒 Seguridad y Validaciones

### **Validaciones Automáticas**

1. **Límite de organizaciones gestionadas**
   - Trigger `trg_ep_managed_check_limit` valida antes de insertar
   - Solo permite gestionar hasta 5 organizaciones

2. **Solo Enterprise Pro puede gestionar múltiples orgs**
   - Validación en `sp_create_managed_organization`
   - Otros planes no pueden gestionar múltiples organizaciones

3. **Solo admin_global puede crear orgs gestionadas**
   - Validación de rol 'admin_global' en organización Enterprise Pro
   - Usuarios 'admin', 'member' o 'viewer' no pueden crear

4. **Separación de metadata**
   - Todos los reportes tienen `organization_id`
   - Queries filtran por `organization_id` del usuario
   - No hay "cross-contamination" entre organizaciones

### **Access Control**

```sql
-- Función para verificar acceso de gestión
fn_can_user_manage_organization(@user_id, @organization_id)
-- Retorna 1 si:
--   - Usuario es admin_global y la org está gestionada por su Enterprise Pro
--   - Usuario es admin normal de la organización
```

## 📊 Vistas Útiles

### **vw_enterprise_pro_organizations**
Muestra todas las organizaciones Enterprise Pro con:
- Número de organizaciones gestionadas creadas
- Slots disponibles
- Estado de suscripción
- Número de admins globales

### **vw_managed_organizations**
Muestra todas las organizaciones gestionadas con:
- Información del Enterprise Pro que las gestiona
- Admin global que las creó
- Número de miembros
- Número de reportes

## 🎓 Diferencias con V1 (Modelo Jerárquico)

| Aspecto | V1 (Jerárquico) | V2 (Independiente) |
|---------|----------------|-------------------|
| **Modelo** | Padre/Hijo | Organizaciones independientes |
| **Campo** | `parent_organization_id` | `enterprise_pro_managed_organizations` |
| **Rol** | `admin` normal | `admin_global` nuevo |
| **Relación** | FK directa | Tabla de relación |
| **Flexibilidad** | Menos flexible | Más flexible |

## 💡 Ventajas del Modelo V2

1. **Organizaciones independientes**: No hay dependencia estructural
2. **Más flexible**: Una org puede cambiar de Enterprise Pro gestionador
3. **Rol explícito**: `admin_global` es claro y específico
4. **Sin jerarquía**: Modelo más simple y menos acoplado
5. **Auditoría mejor**: Tabla de relación permite tracking de quién gestiona qué

## 🚀 Próximos Pasos

1. ✅ Schema implementado
2. ✅ Funciones y procedimientos creados
3. ⏳ Integración con frontend (UI para crear orgs gestionadas)
4. ⏳ Dashboard para admins globales
5. ⏳ Reporting de uso por organización gestionada

## 📝 Notas de Implementación

- Las organizaciones gestionadas tienen sus propias suscripciones (inicialmente `free_trial`)
- Los usuarios pueden pertenecer a múltiples organizaciones con roles distintos
- La separación de metadata está garantizada por `organization_id` en todas las tablas
- El admin_global puede gestionar todas las orgs vinculadas a su Enterprise Pro
- Las organizaciones son completamente independientes - no hay jerarquía







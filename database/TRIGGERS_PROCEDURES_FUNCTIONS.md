# Triggers, Procedures y Funciones - Por Tabla

Documentación completa de todos los triggers, stored procedures y funciones organizados por tabla.

---

## 📋 TABLA: `plans`

### Triggers

#### `trg_plans_updated_at` (schema.sql)
- **Tipo**: AFTER UPDATE
- **Propósito**: Actualiza automáticamente `updated_at` cuando se modifica un plan
- **Se ejecuta**: Cada vez que se hace UPDATE en la tabla
- **Archivo**: schema.sql

---

## 👤 TABLA: `users`

### Triggers

#### `trg_users_updated_at` (schema.sql)
- **Tipo**: AFTER UPDATE
- **Propósito**: Actualiza automáticamente `updated_at` cuando se modifica un usuario
- **Se ejecuta**: Cada vez que se hace UPDATE en la tabla
- **Archivo**: schema.sql

### Stored Procedures

#### `sp_create_user` (state_machine_and_workflows.sql)
```sql
EXEC sp_create_user
    @email = 'usuario@example.com',
    @name = 'Nombre Usuario',
    @auth_provider = 'google',
    @auth_provider_id = 'google_12345',
    @avatar_url = NULL,
    @metadata = NULL
```
- **Propósito**: Crear un nuevo usuario después de OAuth
- **Valida**: Email único, auth_provider válido
- **Retorna**: user_id y mensaje de éxito/error
- **Archivo**: state_machine_and_workflows.sql

---

## 🏢 TABLA: `organizations`

### Triggers

#### `trg_organizations_updated_at` (schema.sql)
- **Tipo**: AFTER UPDATE
- **Propósito**: Actualiza automáticamente `updated_at` cuando se modifica una organización
- **Se ejecuta**: Cada vez que se hace UPDATE
- **Archivo**: schema.sql

#### `trg_organization_auto_assign_free_trial` (state_machine_and_workflows.sql)
- **Tipo**: AFTER INSERT
- **Propósito**: Asigna automáticamente plan `free_trial` cuando se crea una organización
- **Se ejecuta**: Cada vez que se inserta una organización nueva
- **Crea**: Registro en `subscriptions` con status=`trialing`, billing_cycle=NULL
- **Archivo**: state_machine_and_workflows.sql

#### `trg_organization_archive_members` (organization_workflows.sql)
- **Tipo**: AFTER UPDATE
- **Propósito**: Marca a todos los miembros como `left_at = NOW()` cuando se archiva una organización
- **Se ejecuta**: Cuando `is_archived` cambia de 0 a 1
- **Actualiza**: `organization_members.left_at`
- **Archivo**: organization_workflows.sql

### Stored Procedures

#### `sp_create_organization_with_user` (organization_workflows.sql)
```sql
EXEC sp_create_organization_with_user
    @organization_name = 'Mi Organización',
    @user_id = '<GUID>',
    @slug = NULL -- Auto-generado si NULL
```
- **Propósito**: Crear organización y asignar usuario como admin
- **Hace**:
  1. Crea organización
  2. Asigna usuario como admin (`role='admin'`, `is_primary=1`)
  3. Auto-asigna `free_trial` (vía trigger)
  4. Si el usuario ya tiene org primaria, la desmarca
- **Retorna**: organization_id, mensaje, needs_primary_selection
- **Archivo**: organization_workflows.sql

#### `sp_archive_organization` (useful_queries.sql)
```sql
EXEC sp_archive_organization
    @organization_id = '<GUID>',
    @archived_by_user_id = '<GUID>'
```
- **Propósito**: Archivar una organización (no se elimina, se oculta)
- **Hace**:
  1. Marca `is_archived = 1`
  2. Establece `archived_at` y `archived_by`
  3. Cancela suscripción activa
  4. Marca miembros como `left_at` (vía trigger)
- **Archivo**: useful_queries.sql

#### `sp_reactivate_organization` (organization_workflows.sql)
```sql
EXEC sp_reactivate_organization
    @organization_id = '<GUID>',
    @user_id = '<GUID>'
```
- **Propósito**: Reactivar una organización archivada
- **Hace**:
  1. Desmarca `is_archived`
  2. Limpia `archived_at`
  3. Crea nueva suscripción `free_trial`
  4. Re-activa miembros (limpia `left_at`)
- **Archivo**: organization_workflows.sql

### Funciones

#### `fn_can_user_create_organization` (organization_workflows.sql)
```sql
SELECT * FROM dbo.fn_can_user_create_organization('<GUID>')
```
- **Propósito**: Verificar si un usuario puede crear una organización
- **Retorna**: Tabla con información de organizaciones existentes del usuario
- **Campos**: can_create, existing_org_count, has_primary, primary_org_name, etc.
- **Archivo**: organization_workflows.sql

---

## 👥 TABLA: `organization_members`

### Triggers

#### `trg_org_members_updated_at` (schema.sql)
- **Tipo**: AFTER UPDATE
- **Propósito**: Actualiza automáticamente `updated_at`
- **Archivo**: schema.sql

#### `trg_validate_single_primary_organization` (organization_workflows.sql)
- **Tipo**: AFTER INSERT, UPDATE
- **Propósito**: Asegurar que un usuario tenga solo UNA organización primaria (`is_primary = 1`)
- **Valida**: Si un usuario intenta tener más de una org primaria, lanza error
- **Archivo**: organization_workflows.sql

#### `trg_organization_members_check_user_limit` (state_machine_and_workflows.sql)
- **Tipo**: BEFORE INSERT
- **Propósito**: Verificar que no se exceda el límite de usuarios del plan
- **Valida**: Llama a `fn_can_add_user()` antes de insertar
- **Bloquea**: Inserción si se excede el límite
- **Archivo**: state_machine_and_workflows.sql

### Stored Procedures

#### `sp_join_organization_by_invitation` (organization_workflows.sql)
```sql
EXEC sp_join_organization_by_invitation
    @user_id = '<GUID>',
    @invitation_token = 'token_abc123'
```
- **Propósito**: Unirse a una organización por invitación
- **Valida**: Token válido y no expirado
- **Hace**:
  1. Valida token (llama a `fn_validate_invitation_token`)
  2. Agrega usuario como miembro
  3. Asigna rol según invitación
  4. Marca invitación como usada
- **Retorna**: organization_id, nombre, role, had_existing_org
- **Archivo**: organization_workflows.sql

#### `sp_archive_and_join_organization` (organization_workflows.sql)
```sql
EXEC sp_archive_and_join_organization
    @user_id = '<GUID>',
    @current_organization_id = '<GUID>',
    @new_organization_id = '<GUID>'
```
- **Propósito**: Archivar organización actual y unirse a otra
- **Hace**:
  1. Archiva organización actual
  2. Desmarca como primaria
  3. Marca nueva organización como primaria
- **Archivo**: organization_workflows.sql

#### `sp_keep_both_set_new_primary` (organization_workflows.sql)
```sql
EXEC sp_keep_both_set_new_primary
    @user_id = '<GUID>',
    @new_organization_id = '<GUID>'
```
- **Propósito**: Mantener ambas organizaciones, establecer la nueva como primaria
- **Hace**:
  1. Desmarca org anterior como primaria
  2. Marca nueva org como primaria
- **Archivo**: organization_workflows.sql

#### `sp_change_primary_organization` (organization_workflows.sql)
```sql
EXEC sp_change_primary_organization
    @user_id = '<GUID>',
    @organization_id = '<GUID>'
```
- **Propósito**: Cambiar cuál es la organización primaria del usuario
- **Valida**: Usuario pertenece a la organización
- **Archivo**: organization_workflows.sql

#### `sp_create_invitation_token` (organization_workflows.sql)
```sql
EXEC sp_create_invitation_token
    @organization_id = '<GUID>',
    @invited_by = '<GUID>',
    @role = 'member',
    @email = 'invitado@example.com'
```
- **Propósito**: Crear token de invitación para invitar miembros
- **Crea**: Registro en `organization_members` con `invitation_token` y `invitation_expires_at`
- **Expira**: 7 días por defecto
- **Archivo**: organization_workflows.sql

### Funciones

#### `fn_validate_invitation_token` (organization_workflows.sql)
```sql
SELECT * FROM dbo.fn_validate_invitation_token('token_abc123')
```
- **Propósito**: Validar un token de invitación
- **Retorna**: organization_id, organization_name, role, invited_by_name, is_valid, etc.
- **Valida**: Token existe, no expiró, organización no archivada
- **Archivo**: organization_workflows.sql

#### `fn_get_user_organizations` (organization_workflows.sql)
```sql
SELECT * FROM dbo.fn_get_user_organizations('<GUID>')
```
- **Propósito**: Obtener todas las organizaciones de un usuario
- **Retorna**: Tabla con org_id, name, role, is_primary, is_archived, etc.
- **Archivo**: organization_workflows.sql

---

## 💳 TABLA: `subscriptions`

### Triggers

#### `trg_subscriptions_updated_at` (schema.sql)
- **Tipo**: AFTER UPDATE
- **Propósito**: Actualiza automáticamente `updated_at`
- **Archivo**: schema.sql

#### `trg_validate_billing_cycle_by_plan` (constraints_and_validations.sql)
- **Tipo**: BEFORE INSERT, UPDATE
- **Propósito**: Validar que `billing_cycle` sea NULL para `free_trial` y NOT NULL para planes pagos
- **Valida**: Llama a `fn_validate_billing_cycle_for_plan()`
- **Bloquea**: Inserción/actualización si `billing_cycle` es inválido
- **Archivo**: constraints_and_validations.sql

#### `trg_subscriptions_check_expiry` (state_machine_and_workflows.sql)
- **Tipo**: AFTER INSERT, UPDATE
- **Propósito**: Verificar si la suscripción expiró y cambiar estado automáticamente
- **Cambia**: Estado a `canceled` si `current_period_end` < NOW
- **Archivo**: state_machine_and_workflows.sql

### Stored Procedures

#### `sp_change_plan` (useful_queries.sql)
```sql
EXEC sp_change_plan
    @organization_id = '<GUID>',
    @new_plan_id = 'enterprise',
    @new_billing_cycle = 'yearly',
    @changed_by = '<GUID>'
```
- **Propósito**: Cambiar plan de suscripción (upgrade/downgrade)
- **Hace**:
  1. Actualiza `subscriptions.plan_id` y `billing_cycle`
  2. Registra cambio en `subscription_history`
  3. Actualiza `current_period_end` (prorratea)
- **Archivo**: useful_queries.sql

#### `sp_subscription_activate` (state_machine_and_workflows.sql)
```sql
EXEC sp_subscription_activate
    @subscription_id = '<GUID>',
    @stripe_subscription_id = 'sub_stripe_123',
    @stripe_price_id = 'price_123'
```
- **Propósito**: Activar suscripción después de pago exitoso en Stripe
- **Cambia**: status de `trialing` → `active`
- **Registra**: En `subscription_history`
- **Archivo**: state_machine_and_workflows.sql

#### `sp_subscription_cancel` (state_machine_and_workflows.sql)
```sql
EXEC sp_subscription_cancel
    @subscription_id = '<GUID>',
    @cancel_immediately = 0, -- 0 = al final del período, 1 = inmediato
    @canceled_by = '<GUID>'
```
- **Propósito**: Cancelar suscripción
- **Hace**:
  1. Si `cancel_immediately=1`: status → `canceled`
  2. Si `cancel_immediately=0`: marca `cancel_at_period_end = 1`
  3. Registra en `subscription_history`
- **Archivo**: state_machine_and_workflows.sql

#### `sp_subscription_mark_past_due` (state_machine_and_workflows.sql)
```sql
EXEC sp_subscription_mark_past_due
    @subscription_id = '<GUID>',
    @stripe_event_id = 'evt_stripe_123'
```
- **Propósito**: Marcar suscripción como vencida cuando falla el pago
- **Cambia**: status → `past_due`
- **Se ejecuta**: Desde webhook de Stripe cuando falla el pago
- **Archivo**: state_machine_and_workflows.sql

#### `sp_subscription_resolve_past_due` (state_machine_and_workflows.sql)
```sql
EXEC sp_subscription_resolve_past_due
    @subscription_id = '<GUID>',
    @stripe_event_id = 'evt_stripe_456'
```
- **Propósito**: Resolver pago vencido cuando se procesa exitosamente
- **Cambia**: status de `past_due` → `active`
- **Se ejecuta**: Desde webhook de Stripe cuando se resuelve el pago
- **Archivo**: state_machine_and_workflows.sql

#### `sp_subscription_finalize_cancellation` (state_machine_and_workflows.sql)
```sql
EXEC sp_subscription_finalize_cancellation
    @subscription_id = '<GUID>'
```
- **Propósito**: Finalizar cancelación al final del período
- **Se ejecuta**: Job programado o webhook de Stripe
- **Cambia**: status → `canceled`
- **Archivo**: state_machine_and_workflows.sql

#### `sp_update_subscription_plan` (state_machine_and_workflows.sql)
```sql
EXEC sp_update_subscription_plan
    @subscription_id = '<GUID>',
    @new_plan_id = 'enterprise',
    @new_billing_cycle = 'monthly',
    @changed_by = '<GUID>'
```
- **Propósito**: Actualizar plan de suscripción (similar a `sp_change_plan`)
- **Archivo**: state_machine_and_workflows.sql

### Funciones

#### `fn_validate_billing_cycle_for_plan` (constraints_and_validations.sql)
```sql
SELECT dbo.fn_validate_billing_cycle_for_plan('free_trial', NULL) -- Retorna 1
SELECT dbo.fn_validate_billing_cycle_for_plan('free_trial', 'monthly') -- Retorna 0
SELECT dbo.fn_validate_billing_cycle_for_plan('enterprise', 'monthly') -- Retorna 1
```
- **Propósito**: Validar que `billing_cycle` sea correcto según el plan
- **Reglas**:
  - `free_trial`: billing_cycle debe ser NULL
  - Otros planes: billing_cycle debe ser `monthly` o `yearly`
- **Retorna**: 1 (válido) o 0 (inválido)
- **Archivo**: constraints_and_validations.sql

---

## 📝 TABLA: `subscription_history`

### ¡No tiene triggers propios!
- Esta tabla es solo de lectura/inserción
- Se alimenta desde los procedures de `subscriptions`

---

## 📊 TABLA: `reports`

### Triggers

#### `trg_reports_updated_at` (schema.sql)
- **Tipo**: AFTER UPDATE
- **Propósito**: Actualiza automáticamente `updated_at`
- **Archivo**: schema.sql

#### `trg_reports_check_report_limit` (state_machine_and_workflows.sql)
- **Tipo**: BEFORE INSERT
- **Propósito**: Verificar que no se exceda el límite de reportes del plan
- **Valida**: Llama a `fn_can_add_report()` antes de insertar
- **Bloquea**: Inserción si se excede el límite
- **Archivo**: state_machine_and_workflows.sql

#### `trg_reports_validate_organization_for_user` (constraints_and_validations.sql)
- **Tipo**: BEFORE INSERT, UPDATE
- **Propósito**: Validar que `organization_id` sea correcto según el usuario
- **Valida**:
  - Si usuario tiene org: debe especificar `organization_id`
  - Si usuario NO tiene org (basic): `organization_id` debe ser NULL
- **Bloquea**: Inserción/actualización si es inválido
- **Archivo**: constraints_and_validations.sql

### Funciones

#### `fn_can_user_create_individual_report` (constraints_and_validations.sql)
```sql
SELECT dbo.fn_can_user_create_individual_report('<GUID>')
```
- **Propósito**: Verificar si un usuario puede crear reportes individuales (sin organización)
- **Retorna**: 1 (puede) o 0 (no puede)
- **Lógica**: Solo usuarios sin organización o con plan `basic` pueden
- **Archivo**: constraints_and_validations.sql

#### `fn_get_user_effective_plan` (constraints_and_validations.sql)
```sql
SELECT * FROM dbo.fn_get_user_effective_plan('<GUID>')
```
- **Propósito**: Obtener el plan efectivo de un usuario
- **Retorna**: Tabla con plan_id, plan_name, max_users, max_reports, etc.
- **Lógica**: Si tiene org primaria, retorna plan de esa org; sino retorna `basic`
- **Archivo**: constraints_and_validations.sql

---

## 🔗 TABLA: `enterprise_pro_managed_organizations`

### Triggers

#### `trg_ep_managed_check_limit` (enterprise_pro_plan_v2.sql)
- **Tipo**: BEFORE INSERT
- **Propósito**: Verificar que no se exceda el límite de organizaciones gestionadas (max 5)
- **Valida**: Llama a `fn_can_manage_more_organizations()`
- **Bloquea**: Inserción si se excede el límite
- **Archivo**: enterprise_pro_plan_v2.sql

### Stored Procedures

#### `sp_create_managed_organization` (enterprise_pro_plan_v2.sql)
```sql
EXEC sp_create_managed_organization
    @enterprise_pro_org_id = '<GUID>',
    @organization_name = 'Cliente Acme Corp',
    @created_by_user_id = '<GUID>' -- Debe ser admin_global
```
- **Propósito**: Crear una organización gestionada desde Enterprise Pro
- **Valida**:
  1. Usuario es `admin_global` de la org Enterprise Pro
  2. No se excede el límite de 5 orgs gestionadas
  3. Org Enterprise Pro tiene plan `enterprise_pro`
- **Hace**:
  1. Crea nueva organización
  2. Crea registro en `enterprise_pro_managed_organizations`
  3. Asigna `free_trial` a la nueva org
  4. Asigna creator como admin de la nueva org
- **Archivo**: enterprise_pro_plan_v2.sql

### Funciones

#### `fn_can_manage_more_organizations` (enterprise_pro_plan_v2.sql)
```sql
SELECT dbo.fn_can_manage_more_organizations('<GUID>')
```
- **Propósito**: Verificar si una org Enterprise Pro puede gestionar más organizaciones
- **Retorna**: 1 (puede) o 0 (no puede, límite alcanzado)
- **Límite**: 5 organizaciones gestionadas (según `plans.max_organizations`)
- **Archivo**: enterprise_pro_plan_v2.sql

#### `fn_get_managed_organizations_count` (enterprise_pro_plan_v2.sql)
```sql
SELECT dbo.fn_get_managed_organizations_count('<GUID>')
```
- **Propósito**: Contar cuántas organizaciones gestiona una org Enterprise Pro
- **Retorna**: Número entero (0-5)
- **Archivo**: enterprise_pro_plan_v2.sql

#### `fn_is_enterprise_pro_admin` (enterprise_pro_plan_v2.sql)
```sql
SELECT dbo.fn_is_enterprise_pro_admin('<user_id>', '<org_id>')
```
- **Propósito**: Verificar si un usuario es `admin_global` de una org Enterprise Pro
- **Retorna**: 1 (es admin_global) o 0 (no lo es)
- **Archivo**: enterprise_pro_plan_v2.sql

#### `fn_get_user_managed_organizations` (enterprise_pro_plan_v2.sql)
```sql
SELECT * FROM dbo.fn_get_user_managed_organizations('<GUID>')
```
- **Propósito**: Obtener todas las organizaciones gestionadas por un usuario
- **Retorna**: Tabla con org_id, name, admin_count, member_count, etc.
- **Archivo**: enterprise_pro_plan_v2.sql

#### `fn_can_user_manage_organization` (enterprise_pro_plan_v2.sql)
```sql
SELECT dbo.fn_can_user_manage_organization('<user_id>', '<org_id>')
```
- **Propósito**: Verificar si un usuario puede gestionar una organización
- **Retorna**: 1 (puede) o 0 (no puede)
- **Lógica**: Es `admin_global` y la org está gestionada por su Enterprise Pro, o es `admin` de la org
- **Archivo**: enterprise_pro_plan_v2.sql

---

## 🔄 PROCEDIMIENTOS GENERALES (Sin tabla específica)

### `sp_create_or_join_organization` (state_machine_and_workflows.sql)
```sql
EXEC sp_create_or_join_organization
    @user_id = '<GUID>',
    @organization_name = 'Nueva Org',
    @action = 'create' -- 'create' o 'join'
```
- **Propósito**: Procedimiento general para crear o unirse a organización
- **Hace**: Llama a `sp_create_organization_with_user` o `sp_join_organization_by_invitation`
- **Archivo**: state_machine_and_workflows.sql

### `sp_check_organization_limits` (state_machine_and_workflows.sql)
```sql
EXEC sp_check_organization_limits @organization_id = '<GUID>'
```
- **Propósito**: Verificar límites de usuarios y reportes de una organización
- **Retorna**: Tabla con límites, uso actual, y si puede agregar más
- **Archivo**: state_machine_and_workflows.sql

---

## 📊 RESUMEN POR TABLA

| Tabla | Triggers | Procedures | Funciones | Total |
|-------|----------|------------|-----------|-------|
| **plans** | 1 | 0 | 0 | 1 |
| **users** | 1 | 1 | 0 | 2 |
| **organizations** | 3 | 3 | 1 | 7 |
| **organization_members** | 3 | 5 | 2 | 10 |
| **subscriptions** | 3 | 6 | 1 | 10 |
| **subscription_history** | 0 | 0 | 0 | 0 |
| **reports** | 3 | 0 | 2 | 5 |
| **enterprise_pro_managed_orgs** | 1 | 1 | 5 | 7 |
| **TOTAL** | **15** | **16** | **11** | **42** |

---

## 🎯 Funciones de Validación (Resumen)

### **Límites de Plan**
- `fn_can_add_user(@organization_id)` - ¿Puede agregar usuarios?
- `fn_can_add_report(@organization_id)` - ¿Puede agregar reportes?

### **Organizaciones**
- `fn_can_user_create_organization(@user_id)` - ¿Puede crear org?
- `fn_validate_invitation_token(@token)` - ¿Token válido?
- `fn_get_user_organizations(@user_id)` - Obtener todas las orgs del usuario

### **Enterprise Pro**
- `fn_can_manage_more_organizations(@org_id)` - ¿Puede gestionar más orgs?
- `fn_is_enterprise_pro_admin(@user_id, @org_id)` - ¿Es admin_global?
- `fn_get_user_managed_organizations(@user_id)` - Orgs gestionadas
- `fn_can_user_manage_organization(@user_id, @org_id)` - ¿Puede gestionar esta org?

### **Suscripciones**
- `fn_validate_billing_cycle_for_plan(@plan_id, @billing_cycle)` - ¿Billing cycle válido?
- `fn_get_user_effective_plan(@user_id)` - Plan efectivo del usuario

### **Reportes**
- `fn_can_user_create_individual_report(@user_id)` - ¿Puede crear reportes sin org?

---

## 🔧 Triggers de Validación (Resumen)

### **Actualización Automática (6)**
Todas las tablas tienen trigger `trg_<tabla>_updated_at` que actualiza `updated_at` automáticamente.

### **Validación de Límites (2)**
- `trg_organization_members_check_user_limit` - Verifica límite de usuarios
- `trg_reports_check_report_limit` - Verifica límite de reportes

### **Validación de Business Logic (4)**
- `trg_validate_single_primary_organization` - Solo una org primaria por usuario
- `trg_validate_billing_cycle_by_plan` - Billing cycle correcto según plan
- `trg_organization_auto_assign_free_trial` - Auto-asigna free_trial al crear org
- `trg_organization_archive_members` - Marca miembros al archivar org

### **Validación Enterprise Pro (2)**
- `trg_ep_managed_check_limit` - Límite de 5 orgs gestionadas
- `trg_reports_validate_organization_for_user` - Validar org_id según usuario

---

## 📝 Nomenclatura

### **Triggers**
- `trg_<tabla>_updated_at` - Actualización automática de timestamp
- `trg_<tabla>_check_<validacion>` - Validación antes de INSERT/UPDATE
- `trg_<tabla>_validate_<regla>` - Validación de reglas de negocio
- `trg_<tabla>_auto_<accion>` - Acción automática después de INSERT/UPDATE

### **Procedures**
- `sp_create_<entidad>` - Crear nueva entidad
- `sp_<accion>_<entidad>` - Acción sobre entidad
- `sp_<tabla>_<accion>` - Acción específica de tabla

### **Funciones**
- `fn_can_<accion>` - Verificar si puede hacer acción (retorna BIT)
- `fn_get_<dato>` - Obtener dato específico (retorna valor o tabla)
- `fn_validate_<regla>` - Validar regla de negocio (retorna BIT)

---

## 🎓 Flujo de Ejecución Típico

### **Crear Organización**
```
1. sp_create_organization_with_user
   ↓
2. INSERT en organizations
   ↓
3. trg_organization_auto_assign_free_trial (trigger)
   ↓
4. INSERT en subscriptions (free_trial)
   ↓
5. INSERT en organization_members (admin)
```

### **Agregar Miembro**
```
1. sp_create_invitation_token
   ↓
2. Usuario hace click en invitación
   ↓
3. sp_join_organization_by_invitation
   ↓
4. BEFORE INSERT: trg_organization_members_check_user_limit
   ↓ (valida con fn_can_add_user)
5. INSERT en organization_members
```

### **Subir Reporte**
```
1. Frontend llama API de upload
   ↓
2. BEFORE INSERT: trg_reports_check_report_limit
   ↓ (valida con fn_can_add_report)
3. INSERT en reports
   ↓
4. Azure Blob Storage almacena archivo
```

### **Upgrade de Plan**
```
1. Usuario hace upgrade en frontend
   ↓
2. Stripe procesa pago
   ↓
3. Webhook de Stripe llama backend
   ↓
4. sp_change_plan
   ↓
5. UPDATE en subscriptions
   ↓
6. INSERT en subscription_history
```

---

## 📚 Archivos de Referencia

- **schema.sql** - Triggers básicos de updated_at, funciones de límites
- **organization_workflows.sql** - Procedures de creación/unión/archivo
- **state_machine_and_workflows.sql** - Triggers y procedures de suscripciones
- **enterprise_pro_plan_v2.sql** - Funciones y procedures de Enterprise Pro
- **constraints_and_validations.sql** - Triggers y funciones de validación
- **useful_queries.sql** - Procedures útiles (archive, change_plan)







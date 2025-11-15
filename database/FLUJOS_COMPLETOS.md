

# Flujos Completos - Report Tuner

Documentación de flujos principales del sistema con referencias explícitas a triggers, procedures, funciones y tablas involucradas.

---

## 🎯 FLUJO FELIZ: Usuario Nuevo Crea Organización

### Pasos del flujo

```
1. Usuario se registra con AzureAD/LinkedIn
   ↓
2. Usuario crea organización
   ↓
3. Sistema asigna Free Trial automáticamente
   ↓
4. Usuario invita colaboradores
   ↓
5. Usuario sube reportes
   ↓
6. Usuario hace upgrade a plan pago
```

### Paso 1: Registro de Usuario

**Backend llama:**
- `sp_create_user` (state_machine_and_workflows.sql)

**Tablas involucradas:**
- INSERT en `users`

**Triggers ejecutados:**
- Ninguno (es INSERT, updated_at no aplica)

**Resultado:**
- Nuevo registro en `users`
- `is_active = 1`, `is_email_verified = 1` (si es OAuth)

---

### Paso 2: Crear Organización

**Backend llama:**
- `sp_create_organization_with_user` (organization_workflows.sql)

**Proceso interno:**
1. INSERT en `organizations`
2. Trigger `trg_organization_auto_assign_free_trial` se ejecuta automáticamente
3. INSERT en `subscriptions` con plan=`free_trial`, status=`trialing`
4. INSERT en `organization_members` con role=`admin`, is_primary=1

**Tablas involucradas:**
- `organizations` (INSERT)
- `subscriptions` (INSERT automático vía trigger)
- `organization_members` (INSERT)
- `subscription_history` (INSERT para event_type='created')

**Triggers ejecutados:**
- `trg_organization_auto_assign_free_trial` (state_machine_and_workflows.sql)
  - Se ejecuta AFTER INSERT en `organizations`
  - Crea subscription con free_trial

**Funciones llamadas:**
- `fn_can_user_create_organization` (organization_workflows.sql)
  - Verifica si el usuario ya tiene org

**Resultado:**
- Nueva organización creada
- Usuario es admin con is_primary=1
- Subscription free_trial activa

---

### Paso 3: Invitar Colaboradores

**Backend llama:**
- `sp_create_invitation_token` (organization_workflows.sql)

**Proceso interno:**
1. INSERT en `organization_members` con `invitation_token` y `invitation_expires_at`
2. Email se envía desde backend (no desde DB)

**Tablas involucradas:**
- `organization_members` (INSERT con token)

**Triggers ejecutados:**
- Ninguno en INSERT (el trigger de límites solo aplica cuando joined_at es NOW)

**Resultado:**
- Registro pendiente en `organization_members`
- Token válido por 7 días

---

### Paso 4: Aceptar Invitación

**Backend llama:**
- `sp_join_organization_by_invitation` (organization_workflows.sql)

**Proceso interno:**
1. Valida token con `fn_validate_invitation_token`
2. UPDATE en `organization_members`: establece `joined_at = NOW()`
3. Trigger valida límite de usuarios

**Tablas involucradas:**
- `organization_members` (UPDATE)

**Triggers ejecutados:**
- `trg_organization_members_check_user_limit` (state_machine_and_workflows.sql)
  - Valida con `fn_can_add_user`
  - Si límite excedido, bloquea el UPDATE

**Funciones llamadas:**
- `fn_validate_invitation_token` (organization_workflows.sql)
- `fn_can_add_user` (schema.sql) - vía trigger

**Resultado:**
- Usuario agregado a la organización
- Token marcado como usado

---

### Paso 5: Subir Reporte

**Backend llama:**
- API de upload (no es procedure SQL, es código backend)
- INSERT directo en `reports`

**Proceso interno:**
1. BEFORE INSERT: trigger valida límite de reportes
2. INSERT en `reports`
3. Archivo se sube a Azure Blob Storage

**Tablas involucradas:**
- `reports` (INSERT)

**Triggers ejecutados:**
- `trg_reports_check_report_limit` (state_machine_and_workflows.sql)
  - Valida con `fn_can_add_report`
  - Si límite excedido, bloquea el INSERT
- `trg_reports_validate_organization_for_user` (constraints_and_validations.sql)
  - Valida que organization_id sea correcto

**Funciones llamadas:**
- `fn_can_add_report` (schema.sql) - vía trigger
- `fn_can_user_create_individual_report` (constraints_and_validations.sql) - vía trigger

**Resultado:**
- Reporte almacenado en `reports`
- status = 'uploaded'
- Archivo en Azure Blob Storage

---

### Paso 6: Upgrade de Plan

**Stripe procesa pago → Webhook llama backend → Backend llama:**
- `sp_change_plan` (useful_queries.sql)

**Proceso interno:**
1. UPDATE en `subscriptions`: cambia plan_id, billing_cycle, stripe_subscription_id
2. INSERT en `subscription_history` para registrar el cambio
3. Trigger valida que billing_cycle no sea NULL para planes pagos

**Tablas involucradas:**
- `subscriptions` (UPDATE)
- `subscription_history` (INSERT)

**Triggers ejecutados:**
- `trg_validate_billing_cycle_by_plan` (constraints_and_validations.sql)
  - Valida con `fn_validate_billing_cycle_for_plan`
- `trg_subscriptions_updated_at` (schema.sql)
  - Actualiza campo updated_at

**Funciones llamadas:**
- `fn_validate_billing_cycle_for_plan` (constraints_and_validations.sql) - vía trigger

**Resultado:**
- Plan actualizado
- Historial registrado
- Límites actualizados según nuevo plan

---

## 🔀 FLUJOS ALTERNATIVOS

### Alternativa 1: Usuario se Une a Organización Existente

**Pasos:**
1. Usuario recibe invitación (email)
2. Click en link → lleva a landing
3. Usuario hace login/registro
4. Backend llama `sp_join_organization_by_invitation`

**Elementos involucrados:**
- Procedure: `sp_join_organization_by_invitation`
- Función: `fn_validate_invitation_token`
- Trigger: `trg_organization_members_check_user_limit`
- Función (vía trigger): `fn_can_add_user`
- Tabla: `organization_members`

**Decisión: Usuario ya tiene organización propia**

Si el usuario es admin de otra org, frontend presenta opciones:
- **Opción A**: Archivar mi org y unirme a la nueva
  - Backend llama: `sp_archive_and_join_organization`
  - Tablas: `organizations` (is_archived=1), `organization_members` (is_primary actualizado)
  - Triggers: `trg_organization_archive_members` (marca left_at en miembros)

- **Opción B**: Mantener ambas organizaciones
  - Backend llama: `sp_keep_both_set_new_primary`
  - Tablas: `organization_members` (actualiza is_primary)
  - Triggers: `trg_validate_single_primary_organization` (valida una sola primaria)

---

### Alternativa 2: Pago Falla (Subscription Past Due)

**Flujo:**
1. Stripe webhook: `invoice.payment_failed`
2. Backend llama `sp_subscription_mark_past_due`
3. subscription.status → 'past_due'
4. Usuario recibe email (desde HubSpot)
5. Usuario actualiza método de pago
6. Stripe webhook: `invoice.payment_succeeded`
7. Backend llama `sp_subscription_resolve_past_due`
8. subscription.status → 'active'

**Elementos involucrados:**
- Procedures: `sp_subscription_mark_past_due`, `sp_subscription_resolve_past_due`
- Tabla: `subscriptions`, `subscription_history`
- Triggers: `trg_subscriptions_updated_at`

---

### Alternativa 3: Usuario Cancela Suscripción

**Flujo:**
1. Usuario click en "Cancelar" en settings
2. Frontend muestra modal: "Cancelar ahora" o "Al final del período"
3. Backend llama `sp_subscription_cancel` con parámetro `cancel_immediately`

**Si cancel_immediately = 0** (al final del período):
- subscription.cancel_at_period_end = 1
- subscription.status sigue en 'active'
- Job programado ejecuta `sp_subscription_finalize_cancellation` al final

**Si cancel_immediately = 1**:
- subscription.status → 'canceled'
- subscription.canceled_at = NOW()

**Elementos involucrados:**
- Procedures: `sp_subscription_cancel`, `sp_subscription_finalize_cancellation`
- Tabla: `subscriptions`, `subscription_history`
- Triggers: `trg_subscriptions_updated_at`

---

### Alternativa 4: Usuario Reactiva Organización Archivada

**Flujo:**
1. Usuario va a settings → "Mis organizaciones"
2. Ve org archivada
3. Click en "Reactivar"
4. Backend llama `sp_reactivate_organization`

**Proceso:**
1. organization.is_archived → 0
2. Crea nueva subscription free_trial
3. Re-activa miembros (left_at → NULL)

**Elementos involucrados:**
- Procedure: `sp_reactivate_organization`
- Trigger: `trg_organization_auto_assign_free_trial` (crea nueva subscription)
- Tablas: `organizations`, `subscriptions`, `organization_members`

---

## 🏢 FLUJO ENTERPRISE PRO: Gestionar Múltiples Organizaciones

### Paso 1: Admin Global Crea Organización para Cliente

**Backend llama:**
- `sp_create_managed_organization` (enterprise_pro_plan_v2.sql)

**Validaciones previas:**
1. Verifica que usuario es admin_global: `fn_is_enterprise_pro_admin`
2. Verifica límite de 5 orgs: `fn_can_manage_more_organizations`
3. Verifica que org tiene plan enterprise_pro

**Proceso:**
1. INSERT en `organizations` (org del cliente)
2. Trigger asigna free_trial
3. INSERT en `enterprise_pro_managed_organizations`
4. INSERT en `organization_members` (creator como admin de la org cliente)

**Triggers ejecutados:**
- `trg_organization_auto_assign_free_trial`
- `trg_ep_managed_check_limit` (valida límite de 5)

**Funciones llamadas:**
- `fn_is_enterprise_pro_admin`
- `fn_can_manage_more_organizations`

**Resultado:**
- Nueva organización gestionada
- Admin_global puede acceder a ella
- Cliente tiene su propia org con free_trial

---

## 📊 FLUJOS DE VALIDACIÓN AUTOMÁTICA

### Validación de Límites al Agregar Usuario

**Cuando:**
- Se intenta agregar miembro a organización

**Flujo:**
1. Backend llama `sp_join_organization_by_invitation` o INSERT manual
2. BEFORE INSERT: `trg_organization_members_check_user_limit`
3. Trigger llama `fn_can_add_user(@organization_id)`
4. Función consulta:
   - subscriptions.plan_id → plans.max_users
   - COUNT de organization_members activos
5. Si límite excedido: ROLLBACK + error
6. Si ok: continúa INSERT

**Elementos:**
- Trigger: `trg_organization_members_check_user_limit`
- Función: `fn_can_add_user`
- Tablas: `organization_members`, `subscriptions`, `plans`

---

### Validación de Límites al Subir Reporte

**Cuando:**
- Usuario sube reporte

**Flujo:**
1. Backend hace INSERT en `reports`
2. BEFORE INSERT: `trg_reports_check_report_limit`
3. Trigger llama `fn_can_add_report(@organization_id)`
4. Función consulta:
   - subscriptions.plan_id → plans.max_reports
   - COUNT de reports activos (is_deleted=0)
5. Si límite excedido: ROLLBACK + error
6. Si ok: continúa INSERT

**Elementos:**
- Trigger: `trg_reports_check_report_limit`
- Función: `fn_can_add_report`
- Tablas: `reports`, `subscriptions`, `plans`

---

### Validación de Billing Cycle

**Cuando:**
- Se crea o actualiza subscription

**Flujo:**
1. INSERT/UPDATE en `subscriptions`
2. AFTER INSERT/UPDATE: `trg_validate_billing_cycle_by_plan`
3. Trigger llama `fn_validate_billing_cycle_for_plan(@plan_id, @billing_cycle)`
4. Función valida:
   - Si plan = 'free_trial': billing_cycle debe ser NULL
   - Si plan != 'free_trial': billing_cycle debe ser 'monthly' o 'yearly'
5. Si inválido: ROLLBACK + error
6. Si ok: continúa

**Elementos:**
- Trigger: `trg_validate_billing_cycle_by_plan`
- Función: `fn_validate_billing_cycle_for_plan`
- Tabla: `subscriptions`

---

## 🔄 FLUJOS DE SUSCRIPCIÓN (Máquina de Estados)

### Estado: trialing → active

**Cuando:**
- Usuario completa pago en Stripe

**Flujo:**
1. Stripe webhook: `customer.subscription.created`
2. Backend llama `sp_subscription_activate`
3. UPDATE subscriptions: status → 'active', stripe_subscription_id, stripe_price_id
4. INSERT en subscription_history: event_type='updated'

**Elementos:**
- Procedure: `sp_subscription_activate`
- Trigger: `trg_subscriptions_updated_at`
- Tablas: `subscriptions`, `subscription_history`

---

### Estado: active → past_due

**Cuando:**
- Falla el pago (tarjeta rechazada, fondos insuficientes)

**Flujo:**
1. Stripe webhook: `invoice.payment_failed`
2. Backend llama `sp_subscription_mark_past_due`
3. UPDATE subscriptions: status → 'past_due'
4. INSERT en subscription_history: event_type='stripe_webhook'

**Elementos:**
- Procedure: `sp_subscription_mark_past_due`
- Tabla: `subscriptions`, `subscription_history`

---

### Estado: past_due → active

**Cuando:**
- Se resuelve el pago

**Flujo:**
1. Stripe webhook: `invoice.payment_succeeded`
2. Backend llama `sp_subscription_resolve_past_due`
3. UPDATE subscriptions: status → 'active'
4. INSERT en subscription_history: event_type='stripe_webhook'

**Elementos:**
- Procedure: `sp_subscription_resolve_past_due`
- Tabla: `subscriptions`, `subscription_history`

---

### Estado: active → canceled

**Cuando:**
- Usuario cancela

**Flujo Opción A (al final del período):**
1. Usuario click en "Cancelar"
2. Backend llama `sp_subscription_cancel` con cancel_immediately=0
3. UPDATE subscriptions: cancel_at_period_end = 1
4. Al final del período: Job ejecuta `sp_subscription_finalize_cancellation`
5. UPDATE subscriptions: status → 'canceled'

**Flujo Opción B (inmediato):**
1. Usuario click en "Cancelar ahora"
2. Backend llama `sp_subscription_cancel` con cancel_immediately=1
3. UPDATE subscriptions: status → 'canceled', canceled_at = NOW()

**Elementos:**
- Procedures: `sp_subscription_cancel`, `sp_subscription_finalize_cancellation`
- Tabla: `subscriptions`, `subscription_history`

---

## 🔀 FLUJOS DE ORGANIZACIONES

### Cambiar Organización Primaria

**Cuando:**
- Usuario pertenece a múltiples orgs y quiere cambiar cuál es la principal

**Flujo:**
1. Frontend muestra dropdown de organizaciones
2. Usuario selecciona otra
3. Backend llama `sp_change_primary_organization`
4. UPDATE en organization_members: desmarca todas, marca solo la seleccionada

**Elementos:**
- Procedure: `sp_change_primary_organization`
- Trigger: `trg_validate_single_primary_organization` (valida que solo haya una primaria)
- Tabla: `organization_members`

---

### Archivar y Unirse a Otra

**Cuando:**
- Usuario admin de org A recibe invitación a org B y elige archivar A

**Flujo:**
1. Usuario acepta invitación
2. Frontend muestra opciones (archivar vs mantener)
3. Usuario elige "Archivar mi org y unirme"
4. Backend llama `sp_archive_and_join_organization`
5. Proceso:
   - UPDATE organizations: is_archived=1, archived_at=NOW()
   - Trigger archiva miembros: `trg_organization_archive_members`
   - UPDATE organization_members: desmarca org A como primaria
   - UPDATE organization_members: marca org B como primaria
   - UPDATE subscriptions de org A: status='canceled'

**Elementos:**
- Procedure: `sp_archive_and_join_organization`
- Trigger: `trg_organization_archive_members`
- Función: `fn_validate_invitation_token`
- Tablas: `organizations`, `organization_members`, `subscriptions`

---

## 🎯 FLUJO COMPLETO: Upgrade de Plan

### Pasos detallados

**1. Usuario en settings click en "Upgrade a Enterprise"**

Frontend muestra pricing, usuario selecciona plan y billing_cycle.

**2. Frontend redirige a Stripe Checkout**

```javascript
const session = await stripe.checkout.sessions.create({
  customer: organization.stripe_customer_id,
  line_items: [{ price: 'price_enterprise_monthly', quantity: 1 }],
  mode: 'subscription'
});
```

**3. Usuario completa pago en Stripe**

Stripe procesa pago exitosamente.

**4. Stripe envía webhook: `checkout.session.completed`**

Backend recibe webhook.

**5. Backend llama procedure**

```sql
EXEC sp_change_plan
    @organization_id = '<GUID>',
    @new_plan_id = 'enterprise',
    @new_billing_cycle = 'monthly',
    @changed_by = '<user_id>';
```

**6. Proceso interno del procedure:**

```
sp_change_plan:
  ├── Obtiene subscription actual de la org
  ├── UPDATE subscriptions
  │     ├── plan_id → 'enterprise'
  │     ├── billing_cycle → 'monthly'
  │     ├── stripe_subscription_id → 'sub_xxx'
  │     ├── status → 'active'
  │     └── current_period_end → +30 días
  ├── INSERT en subscription_history
  │     ├── plan_id_old → 'free_trial'
  │     ├── plan_id_new → 'enterprise'
  │     ├── event_type → 'plan_changed'
  │     └── stripe_event_id → 'evt_xxx'
  └── Trigger: trg_subscriptions_updated_at
        └── updated_at → NOW()
```

**7. Frontend recibe confirmación**

Backend retorna success, frontend muestra mensaje de éxito.

**Elementos involucrados:**
- Procedure: `sp_change_plan` (useful_queries.sql)
- Trigger: `trg_subscriptions_updated_at` (schema.sql)
- Trigger: `trg_validate_billing_cycle_by_plan` (constraints_and_validations.sql)
- Función (vía trigger): `fn_validate_billing_cycle_for_plan`
- Tablas: `subscriptions`, `subscription_history`

---

## 🔄 FLUJO: Procesamiento de Reporte

### Pasos

**1. Usuario sube archivo .pbit**

Frontend llama API de upload con el archivo.

**2. Backend valida y crea registro**

```javascript
// Antes de INSERT
const canAdd = await db.query('SELECT dbo.fn_can_add_report(@org_id)');
if (!canAdd) {
  return res.status(400).json({ error: 'Límite de reportes alcanzado' });
}

// INSERT
await db.execute(`
  INSERT INTO reports (organization_id, user_id, name, original_filename, file_size_bytes, status)
  VALUES (@org_id, @user_id, @name, @filename, @size, 'uploaded')
`);
```

**3. Triggers se ejecutan automáticamente**

```
INSERT en reports
  ↓
Trigger: trg_reports_check_report_limit
  ├── Llama fn_can_add_report(@organization_id)
  ├── Valida límite del plan
  └── Si límite excedido → ROLLBACK
  ↓
Trigger: trg_reports_validate_organization_for_user
  ├── Valida que organization_id sea correcto
  └── Si usuario tiene org, debe especificar organization_id
  ↓
INSERT exitoso
```

**4. Backend procesa archivo**

```javascript
// Upload a Azure Blob Storage
const blobUrl = await uploadToAzure(file);

// UPDATE report
await db.execute(`
  UPDATE reports 
  SET status = 'processing', 
      file_url = @url, 
      blob_name = @blob,
      processing_started_at = GETUTCDATE()
  WHERE id = @report_id
`);

// Procesamiento (extraer metadata)
const metadata = await processP

BIT(file);

// UPDATE report
await db.execute(`
  UPDATE reports 
  SET status = 'processed', 
      metadata = @metadata,
      processing_completed_at = GETUTCDATE()
  WHERE id = @report_id
`);
```

**Elementos involucrados:**
- Triggers: `trg_reports_check_report_limit`, `trg_reports_validate_organization_for_user`, `trg_reports_updated_at`
- Funciones: `fn_can_add_report`, `fn_can_user_create_individual_report`
- Tabla: `reports`

---

## 📊 TABLA RESUMEN: Triggers, Procedures y Funciones por Flujo

| Flujo | Procedures | Funciones | Triggers | Tablas |
|-------|------------|-----------|----------|--------|
| **Crear organización** | `sp_create_organization_with_user` | `fn_can_user_create_organization` | `trg_organization_auto_assign_free_trial` | `organizations`, `subscriptions`, `organization_members`, `subscription_history` |
| **Invitar miembro** | `sp_create_invitation_token`, `sp_join_organization_by_invitation` | `fn_validate_invitation_token`, `fn_can_add_user` | `trg_organization_members_check_user_limit` | `organization_members` |
| **Subir reporte** | - | `fn_can_add_report`, `fn_can_user_create_individual_report` | `trg_reports_check_report_limit`, `trg_reports_validate_organization_for_user` | `reports` |
| **Upgrade plan** | `sp_change_plan` | `fn_validate_billing_cycle_for_plan` | `trg_validate_billing_cycle_by_plan`, `trg_subscriptions_updated_at` | `subscriptions`, `subscription_history` |
| **Cancelar subscription** | `sp_subscription_cancel`, `sp_subscription_finalize_cancellation` | - | `trg_subscriptions_updated_at` | `subscriptions`, `subscription_history` |
| **Archivar org** | `sp_archive_and_join_organization` | - | `trg_organization_archive_members` | `organizations`, `organization_members`, `subscriptions` |
| **Reactivar org** | `sp_reactivate_organization` | - | `trg_organization_auto_assign_free_trial` | `organizations`, `subscriptions`, `organization_members` |
| **Enterprise Pro: crear org gestionada** | `sp_create_managed_organization` | `fn_can_manage_more_organizations`, `fn_is_enterprise_pro_admin` | `trg_ep_managed_check_limit` | `enterprise_pro_managed_organizations`, `organizations`, `subscriptions` |

---

## 🎯 FLUJOS EN UN VISTAZO

### Flujo Feliz (Sin problemas)
```
Registro → Crear Org → Invite Miembros → Sube Reportes → Upgrade → Pago OK
```

### Flujos con Decisiones
```
Registro → Recibe Invitación → ¿Tiene org propia?
                                 ├─ NO → Se une directamente
                                 └─ SÍ → ¿Archivar o Mantener?
                                         ├─ Archivar → sp_archive_and_join_organization
                                         └─ Mantener → sp_keep_both_set_new_primary
```

### Flujos de Error/Recuperación
```
Active → Pago Falla → Past Due → ¿Resuelve pago?
                                  ├─ SÍ → sp_subscription_resolve_past_due → Active
                                  └─ NO → (después de X días) → Canceled
```

---

## 📚 Referencias

- **schema.sql** - Tablas base, triggers de updated_at, funciones de límites
- **organization_workflows.sql** - Procedures de creación/unión/archivo
- **state_machine_and_workflows.sql** - Procedures de suscripciones, triggers de límites
- **constraints_and_validations.sql** - Triggers y funciones de validación
- **enterprise_pro_plan_v2.sql** - Procedures y funciones de Enterprise Pro
- **useful_queries.sql** - Procedures útiles (archive, change_plan)







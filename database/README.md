# Report Tuner - Database Schema (Simplificado)

Esquema simplificado y modular para el sistema SaaS Report Tuner.

## 📋 Archivos SQL

### **Instalación** (en este orden)

1. **`schema.sql`** ⭐ - Schema principal con tablas, vistas, funciones básicas y triggers de updated_at
2. **`organization_workflows.sql`** - Procedures y funciones para creación/unión a organizaciones
3. **`state_machine_and_workflows.sql`** - Máquina de estados de suscripciones
4. **`constraints_and_validations.sql`** - Validaciones adicionales
5. **`documentation_procedures.sql`** - Procedures para gestionar documentación
6. **`settings_tables_and_procedures.sql`** - Tablas y procedures para Settings (Stripe invoices, integrations) (opcional)
7. **`enterprise_pro_plan_v2.sql`** - Enterprise Pro multi-organización (opcional)
8. **`useful_queries.sql`** - Procedures útiles (opcional)

### **Solo consulta** (no ejecutar)

- **`tables_only.sql`** - Solo definiciones de tablas (para referencia)

---

## 📚 Documentación

### **Guías Principales**
- **`README.md`** - Este archivo
- **`INSTALLATION_ORDER.md`** - Orden de ejecución paso a paso
- **`FLUJOS_COMPLETOS.md`** 🎯 - **Flujos del sistema con referencias a triggers/procedures/funciones**
- **`ARCHITECTURE_SIMPLE.md`** - Filosofía del diseño simplificado
- **`TRIGGERS_PROCEDURES_FUNCTIONS.md`** - Lista completa organizada por tabla

### **Guías Específicas**
- **`ENTERPRISE_PRO_V2_README.md`** - Documentación de Enterprise Pro
- **`DIAGRAM_PROMPT.md`** - Para generar diagrama UML/ER
- **`SAAS_TOOLS_AND_SYSTEMS.md`** - Herramientas externas (HubSpot, Stripe, etc.)
- **`SCHEMA_OVERVIEW.md`** - Resumen de cambios y simplificación

### **Excel**
- **`DATABASE_SIMPLE.xlsx`** 📊 - Todas las tablas con datos dummy relacionados (27 registros)

---

## 🚀 Quick Start

### Instalación Completa

```sql
-- 1. Schema base (OBLIGATORIO)
USE master;
GO
-- Ejecutar schema.sql

-- 2. Workflows (OBLIGATORIO)
USE empower_reports;
GO
-- Ejecutar organization_workflows.sql
-- Ejecutar state_machine_and_workflows.sql
-- Ejecutar constraints_and_validations.sql

-- 3. Enterprise Pro (OPCIONAL - solo si necesitas multi-org)
-- Ejecutar enterprise_pro_plan_v2.sql
```

### Instalación Mínima (Solo lo esencial)

```sql
-- Solo estos 4 archivos
1. schema.sql
2. organization_workflows.sql
3. state_machine_and_workflows.sql
4. constraints_and_validations.sql
```

---

## 📊 Contenido de Cada Archivo

| Archivo | Tablas | Triggers | Procedures | Funciones | Vistas |
|---------|--------|----------|------------|-----------|--------|
| **schema.sql** | 8 | 7 | 0 | 3 | 2 |
| **documentation_procedures.sql** | 0 | 0 | 3 | 0 | 0 |
| **organization_workflows.sql** | 0 | 2 | 6 | 3 | 1 |
| **state_machine_and_workflows.sql** | 0 | 4 | 8 | 0 | 2 |
| **constraints_and_validations.sql** | 0 | 2 | 0 | 3 | 0 |
| **settings_tables_and_procedures.sql** (opcional) | 2 | 2 | 11 | 4 | 2 |
| **enterprise_pro_plan_v2.sql** (opcional) | 1 | 1 | 1 | 5 | 2 |
| **useful_queries.sql** (opcional) | 0 | 0 | 2 | 0 | 0 |
| **TOTAL** | **11** | **17** | **31** | **17** | **9** |

---

## 📋 Estructura de Tablas

### Tablas Principales (8)

1. **`plans`** - Planes con límites (free_trial, basic, teams, enterprise, enterprise_pro)
2. **`users`** - Usuarios con OAuth (Google, LinkedIn, Azure AD) y auth local
3. **`organizations`** - Organizaciones donde colaboran usuarios
4. **`organization_documentation`** - URLs de documentación por organización (habilita botón "Ver documentación")
5. **`organization_members`** - Relación usuarios ↔ organizaciones con roles
6. **`subscriptions`** - Suscripciones activas (integración con Stripe)
7. **`subscription_history`** - Historial de cambios
8. **`reports`** - Reportes subidos (pueden ser de org o individuales)

### Tabla Enterprise Pro (1)

9. **`enterprise_pro_managed_organizations`** - Organizaciones gestionadas por Enterprise Pro

---

## 🎯 Elementos Principales

### Triggers (15)

**Actualización automática (6)**
- `trg_users_updated_at`
- `trg_organizations_updated_at`
- `trg_plans_updated_at`
- `trg_subscriptions_updated_at`
- `trg_org_members_updated_at`
- `trg_reports_updated_at`

**Validación de límites (2)**
- `trg_organization_members_check_user_limit`
- `trg_reports_check_report_limit`

**Validación de business logic (4)**
- `trg_validate_single_primary_organization`
- `trg_validate_billing_cycle_by_plan`
- `trg_organization_auto_assign_free_trial`
- `trg_organization_archive_members`

**Validaciones específicas (3)**
- `trg_reports_validate_organization_for_user`
- `trg_subscriptions_check_expiry`
- `trg_ep_managed_check_limit` (Enterprise Pro)

### Stored Procedures (17)

**Organizaciones (6)**
- `sp_create_organization_with_user`
- `sp_join_organization_by_invitation`
- `sp_archive_and_join_organization`
- `sp_keep_both_set_new_primary`
- `sp_change_primary_organization`
- `sp_reactivate_organization`

**Invitaciones (1)**
- `sp_create_invitation_token`

**Suscripciones (8)**
- `sp_subscription_activate`
- `sp_subscription_cancel`
- `sp_subscription_mark_past_due`
- `sp_subscription_resolve_past_due`
- `sp_subscription_finalize_cancellation`
- `sp_update_subscription_plan`
- `sp_change_plan`
- `sp_archive_organization`

**Usuarios (1)**
- `sp_create_user`

**Enterprise Pro (1)**
- `sp_create_managed_organization`

### Funciones (13)

**Validación de límites (2)**
- `fn_can_add_user(@organization_id)` - ¿Puede agregar usuarios?
- `fn_can_add_report(@organization_id)` - ¿Puede agregar reportes?

**Organizaciones (3)**
- `fn_can_user_create_organization(@user_id)` - ¿Puede crear org?
- `fn_validate_invitation_token(@token)` - Validar token de invitación
- `fn_get_user_organizations(@user_id)` - Obtener todas las orgs del usuario

**Suscripciones (1)**
- `fn_validate_billing_cycle_for_plan(@plan_id, @billing_cycle)` - Validar billing cycle

**Reportes (2)**
- `fn_can_user_create_individual_report(@user_id)` - ¿Puede crear reportes sin org?
- `fn_get_user_effective_plan(@user_id)` - Plan efectivo del usuario

**Enterprise Pro (5)**
- `fn_can_manage_more_organizations(@org_id)` - ¿Puede gestionar más orgs?
- `fn_get_managed_organizations_count(@org_id)` - Contar orgs gestionadas
- `fn_is_enterprise_pro_admin(@user_id, @org_id)` - ¿Es admin_global?
- `fn_get_user_managed_organizations(@user_id)` - Obtener orgs gestionadas
- `fn_can_user_manage_organization(@user_id, @org_id)` - ¿Puede gestionar esta org?

### Vistas (7)

- `vw_organizations_with_subscription` - Orgs con suscripciones activas
- `vw_users_with_primary_org` - Usuarios con org primaria
- `vw_user_organizations_dashboard` - Vista completa para dashboard
- `vw_organizations_usage_status` - Uso vs límites
- `vw_subscriptions_requiring_attention` - Suscripciones que requieren atención
- `vw_enterprise_pro_organizations` - Orgs Enterprise Pro (opcional)
- `vw_managed_organizations` - Orgs gestionadas (opcional)

---

## 📈 Planes y Límites

| Plan | Usuarios | Reportes | Storage | Precio/mes | Multi-Org |
|------|----------|----------|---------|------------|-----------|
| Free Trial | 10 | 100 | 5GB | Gratis | - |
| Basic | 1 | 30 | 1GB | $9.99 | - |
| Teams | 3 | 50 | 5GB | $29.99 | - |
| Enterprise | 10 | 300 | 50GB | $99.99 | - |
| Enterprise Pro | 50 | 1000 | 200GB | $199.99 | ✅ Hasta 5 |

---

## 💡 Filosofía: Simple y Delegado

**Lo que maneja la DB:**
- ✅ Usuarios y autenticación
- ✅ Organizaciones y membresías
- ✅ Planes y suscripciones
- ✅ Reportes y almacenamiento
- ✅ Validaciones de límites
- ✅ Historial de cambios

**Lo que se delega:**
- ❌ A/B Testing → HubSpot
- ❌ Pricing regional → Stripe + HubSpot
- ❌ Email marketing → HubSpot
- ❌ Analytics → HubSpot + Google Analytics
- ❌ Segmentación → HubSpot

---

## 🔍 Explorar el Sistema

### Para entender las tablas:
- Abrir **`DATABASE_SIMPLE.xlsx`** con datos dummy

### Para entender los flujos:
- Leer **`FLUJOS_COMPLETOS.md`**

### Para ver qué hace cada trigger/procedure/función:
- Leer **`TRIGGERS_PROCEDURES_FUNCTIONS.md`**

### Para instalar:
- Seguir **`INSTALLATION_ORDER.md`**

### Para entender la arquitectura:
- Leer **`ARCHITECTURE_SIMPLE.md`**

---

## 📝 Notas Importantes

1. **organization_id en reports puede ser NULL** - Para usuarios individuales (plan basic)
2. **billing_cycle en subscriptions puede ser NULL** - Solo para free_trial
3. **admin_global es un rol especial** - Solo en Enterprise Pro para gestionar múltiples orgs
4. **No hay jerarquía padre/hijo** - Las organizaciones son independientes
5. **Triggers automáticos** - Free trial se asigna automáticamente al crear org
6. **Validación en tiempo real** - Triggers bloquean si se exceden límites

---

## 🎓 Flujos Clave

Ver **`FLUJOS_COMPLETOS.md`** para:
- Flujo feliz completo paso a paso
- Flujos alternativos (archivar, mantener ambas, etc.)
- Flujos de error/recuperación (past_due, cancelación)
- Tabla resumen de elementos por flujo
- Diagramas de máquina de estados

---

## 📚 Más Información

- **Enterprise Pro**: Ver `ENTERPRISE_PRO_V2_README.md`
- **Herramientas SaaS**: Ver `SAAS_TOOLS_AND_SYSTEMS.md`
- **Diagrama UML**: Ver `DIAGRAM_PROMPT.md`

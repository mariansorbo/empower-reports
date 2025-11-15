# Report Tuner - Arquitectura Simplificada

## 🎯 Filosofía: Simple y Escalable

Este esquema mantiene solo lo esencial para el funcionamiento del sistema SaaS:
- **Usuarios y organizaciones**: Gestión de acceso y colaboración
- **Planes y suscripciones**: Facturación y límites
- **Reportes**: Almacenamiento y tracking
- **Enterprise Pro**: Multi-organización (opcional)

Todo lo demás se delega a herramientas especializadas:
- **A/B Testing**: HubSpot
- **Pricing complejo/regional**: Stripe + HubSpot
- **Email marketing**: HubSpot
- **Analytics avanzado**: HubSpot + Google Analytics

## 📊 Tablas del Sistema (7 principales)

### 1. `plans`
Define los 5 planes disponibles:
- `free_trial` (10 usuarios, 100 reportes)
- `basic` (1 usuario, 30 reportes)
- `teams` (3 usuarios, 50 reportes)
- `enterprise` (10 usuarios, 300 reportes)
- `enterprise_pro` (50 usuarios, 1000 reportes, 5 orgs gestionadas)

### 2. `users`
Usuarios con OAuth (Google, LinkedIn, Azure AD) o email/password.

### 3. `organizations`
Organizaciones donde colaboran usuarios.
- Vinculadas a Stripe
- Pueden ser archivadas
- Un usuario puede pertenecer a múltiples organizaciones

### 4. `organization_members`
Relación usuarios ↔ organizaciones con roles:
- `admin`: Administrador de la organización
- `admin_global`: Administrador Enterprise Pro (gestiona múltiples orgs)
- `member`: Miembro colaborador
- `viewer`: Solo lectura

### 5. `subscriptions`
Suscripciones activas a planes:
- Estados: `active`, `trialing`, `canceled`, `past_due`, `unpaid`, `incomplete`
- `billing_cycle`: `monthly`, `yearly`, o NULL (`free_trial`)
- Integración con Stripe

### 6. `subscription_history`
Historial de todos los cambios:
- Upgrades, downgrades
- Cancelaciones, reactivaciones
- Eventos de Stripe


### 8. `enterprise_pro_managed_organizations` (Opcional)
Solo para Enterprise Pro:
- Relaciona org Enterprise Pro con orgs gestionadas
- Organizaciones independientes (no jerarquía)

## 🔄 Flujo de Usuario

```
1. Usuario se registra → users
2. Usuario crea organización → organizations + organization_members (admin)
3. Se asigna free_trial → subscriptions (status=trialing)
4. Usuario invita miembros → organization_members
5. Usuario sube reportes → reports
6. Usuario hace upgrade → subscriptions (cambio de plan) + subscription_history
7. Stripe procesa pago → subscriptions (stripe_subscription_id)
```

## 💡 Integración con Herramientas Externas

### **HubSpot**
- A/B Testing de landing pages
- Email marketing
- Lead tracking
- CRM general

### **Stripe**
- Procesamiento de pagos
- Gestión de suscripciones
- Webhooks para actualizar estado

### **Azure Blob Storage**
- Almacenamiento de archivos .pbit
- URLs de reportes

### **Google Analytics / Mixpanel**
- Analytics de producto
- Tracking de eventos
- Funnel de conversión

## 🎯 ¿Qué se maneja en la Base de Datos?

✅ **Sí se maneja:**
- Usuarios y autenticación
- Organizaciones y membresías
- Planes y suscripciones
- Reportes subidos
- Límites por plan
- Historial de cambios

❌ **No se maneja (se delega):**
- A/B Testing → HubSpot
- Pricing dinámico por región → Stripe
- Segmentación de marketing → HubSpot
- Geolocalización → HubSpot / Analytics
- Email campaigns → HubSpot
- Landing pages → HubSpot

## 🔧 Procedimientos Principales

### Organizaciones
- `sp_create_organization_with_user` - Crear org + asignar free_trial
- `sp_join_organization_by_invitation` - Unirse por invitación
- `sp_archive_and_join_organization` - Archivar y unirse a otra
- `sp_reactivate_organization` - Reactivar archivada

### Suscripciones
- `sp_change_plan` - Cambiar plan (upgrade/downgrade)
- `sp_subscription_activate` - Activar suscripción
- `sp_subscription_cancel` - Cancelar suscripción

### Enterprise Pro
- `sp_create_managed_organization` - Crear org gestionada
- `fn_can_manage_more_organizations` - Verificar límite

## 📦 Archivos del Schema

1. **`schema.sql`** - Schema principal (¡EJECUTAR PRIMERO!)
2. **`organization_workflows.sql`** - Workflows de organizaciones
3. **`state_machine_and_workflows.sql`** - Estado de suscripciones
4. **`enterprise_pro_plan_v2.sql`** - Enterprise Pro (opcional)
5. **`useful_queries.sql`** - Queries útiles

## 🚀 Quick Start

```sql
-- 1. Crear base de datos y tablas
EXEC schema.sql

-- 2. Workflows
EXEC organization_workflows.sql
EXEC state_machine_and_workflows.sql

-- 3. Enterprise Pro (opcional)
EXEC enterprise_pro_plan_v2.sql
```

## 📝 Filosofía de Diseño

1. **Simple**: Solo lo necesario en la DB
2. **Escalable**: Fácil de mantener y extender
3. **Integrable**: Se conecta bien con Stripe, HubSpot, etc.
4. **Enfocado**: La DB hace lo que hace mejor (persistencia, relaciones, validaciones)
5. **Delegar**: Todo lo demás a herramientas especializadas

**Menos código = menos bugs = más fácil de mantener**







# 📊 SETTINGS DATABASE MAPPING - SIMPLIFIED VERSION

Análisis **simplificado** del panel de Settings después de eliminar Preferences y Avatar.

---

## 📑 CATEGORÍAS ACTUALES (5 secciones)

1. [👤 Profile](#1-profile)
2. [🔐 Security](#2-security)
3. [🏢 Organization](#3-organization)
4. [💳 Billing & Plan](#4-billing--plan)
5. [🔌 Integrations](#5-integrations)

---

## 1. 👤 PROFILE

**Descripción:** Configuración individual de la cuenta (SIMPLIFICADA)

| Campo | Tipo | Descripción | Tabla actual | Estado |
|-------|------|-------------|--------------|--------|
| **Full name** | String | Nombre completo del usuario | `users.name` | ✅ YA EXISTE |
| **Email** | String | Email (read-only de LinkedIn) | `users.email` | ✅ YA EXISTE |

**Estado:** ✅✅ COMPLETO (100%) - No necesita cambios

---

## 2. 🔐 SECURITY

**Descripción:** Cuentas vinculadas

| Campo | Tipo | Descripción | Tabla actual | Estado |
|-------|------|-------------|--------------|--------|
| **LinkedIn status** | Boolean | Cuenta LinkedIn conectada | `users.auth_provider` | ✅ YA EXISTE |

**Estado:** ✅✅ COMPLETO (100%) - No necesita cambios

---

## 3. 🏢 ORGANIZATION

**Descripción:** Configuración de la organización

| Campo | Tipo | Descripción | Tabla actual | Estado |
|-------|------|-------------|--------------|--------|
| **Organization name** | String | Nombre de la organización | `organizations.name` | ✅ YA EXISTE |
| **Logo URL** | String | URL del logo | `organizations.logo_url` | ✅ YA EXISTE |
| **Slug** | String | URL-friendly identifier | `organizations.slug` | ✅ YA EXISTE |
| **Website** | String | Sitio web corporativo | `organizations.website` | ✅ YA EXISTE |
| **Roles info** | Static | Documentación de roles | N/A (estático) | ✅ No requiere DB |
| **Delete org** | Action | Eliminar organización | `organizations.is_archived` | ✅ YA EXISTE |

**Estado:** ✅✅ COMPLETO (100%) - No necesita cambios

---

## 4. 💳 BILLING & PLAN

**Descripción:** Control de suscripción y pagos

### 4.1. **Plan actual** (100% cubierto)

| Campo | Tipo | Descripción | Tabla actual | Estado |
|-------|------|-------------|--------------|--------|
| **Plan name** | String | free_trial/basic/teams/enterprise | `subscriptions.plan_id` → `plans.name` | ✅ YA EXISTE |
| **Status** | String | active/trialing/canceled | `subscriptions.status` | ✅ YA EXISTE |
| **Period start/end** | DateTime | Período de facturación | `subscriptions.current_period_*` | ✅ YA EXISTE |
| **Billing cycle** | String | monthly/yearly | `subscriptions.billing_cycle` | ✅ YA EXISTE |

### 4.2. **Límites actuales** (100% cubierto)

| Campo | Tipo | Descripción | Tabla actual | Estado |
|-------|------|-------------|--------------|--------|
| **Max users** | Integer | Límite de usuarios | `plans.max_users` | ✅ YA EXISTE |
| **Max reports** | Integer | Límite de reportes | `plans.max_reports` | ✅ YA EXISTE |
| **Current users** | Integer (calc) | Usuarios actuales | Query: `organization_members` | ✅ YA EXISTE |
| **Current reports** | Integer (calc) | Reportes actuales | Query: `reports` | ✅ YA EXISTE |

### 4.3. **💳 Payment method** (FALTA - para cuando implementes Stripe)

| Campo | Tipo | Descripción | Tabla actual | Estado |
|-------|------|-------------|--------------|--------|
| **Stripe Customer ID** | String | Customer en Stripe | `organizations.stripe_customer_id` | ✅ YA EXISTE |
| **Default payment method** | String | pm_xxx (Stripe Payment Method ID) | ❌ NO EXISTE | 🟡 AGREGAR |
| **Card info** | JSON | Brand, last4, exp (opcional) | ❌ NO EXISTE | 🟢 OPCIONAL (via Stripe API) |

**Recomendación mínima:**
```sql
ALTER TABLE organizations 
ADD stripe_payment_method_id VARCHAR(255) NULL;

CREATE INDEX idx_organizations_payment_method 
ON organizations(stripe_payment_method_id) 
WHERE stripe_payment_method_id IS NOT NULL;
```

### 4.4. **📄 Billing history** (FALTA - para cuando implementes Stripe)

| Campo | Tipo | Descripción | Tabla actual | Estado |
|-------|------|-------------|--------------|--------|
| **Invoice list** | Array | Lista de facturas | ❌ NO EXISTE | 🔴 AGREGAR: Tabla `stripe_invoices` |

**Recomendación:**
```sql
CREATE TABLE stripe_invoices (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL,
    subscription_id UNIQUEIDENTIFIER NOT NULL,
    stripe_invoice_id VARCHAR(255) NOT NULL UNIQUE,
    stripe_invoice_pdf VARCHAR(500) NULL,
    stripe_hosted_invoice_url VARCHAR(500) NULL,
    amount_due DECIMAL(10, 2) NOT NULL,
    amount_paid DECIMAL(10, 2) NOT NULL,
    currency VARCHAR(3) NOT NULL DEFAULT 'USD',
    status VARCHAR(50) NOT NULL CHECK (status IN ('draft', 'open', 'paid', 'void', 'uncollectible')),
    period_start DATETIME2 NOT NULL,
    period_end DATETIME2 NOT NULL,
    paid_at DATETIME2 NULL,
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_invoices_org FOREIGN KEY (organization_id) REFERENCES organizations(id),
    CONSTRAINT fk_invoices_subscription FOREIGN KEY (subscription_id) REFERENCES subscriptions(id)
);
```

**Estado:** 🟡 NECESITA tabla nueva (solo cuando implementes Stripe)

---

## 5. 🔌 INTEGRATIONS

**Descripción:** Conexiones con herramientas externas

### Integraciones planificadas:

#### 5.1. **Power BI Service**

| Campo | Tipo | Descripción | Tabla actual | Estado |
|-------|------|-------------|--------------|--------|
| **Workspace ID** | String | ID del workspace de Power BI | ❌ NO EXISTE | 🔴 AGREGAR |
| **Connected** | Boolean | ¿Conectado? | ❌ NO EXISTE | 🔴 AGREGAR |
| **Access token** | Binary | Token OAuth (cifrado) | ❌ NO EXISTE | 🔴 AGREGAR |
| **Last sync** | DateTime | Última sincronización | ❌ NO EXISTE | 🔴 AGREGAR |

#### 5.2. **API - Upload .pbit files**

| Campo | Tipo | Descripción | Tabla actual | Estado |
|-------|------|-------------|--------------|--------|
| **API enabled** | Boolean | ¿API habilitada? | ❌ NO EXISTE | 🔴 AGREGAR |
| **API key** | String | Public key (pk_xxx) | ❌ NO EXISTE | 🔴 AGREGAR |
| **API secret** | String | Secret key (hashed) | ❌ NO EXISTE | 🔴 AGREGAR |
| **Rate limit** | Integer | Requests por hora | ❌ NO EXISTE | 🔴 AGREGAR |
| **Last used** | DateTime | Último uso del API | ❌ NO EXISTE | 🔴 AGREGAR |

#### 5.3. **API - Documentation endpoint**

| Campo | Tipo | Descripción | Tabla actual | Estado |
|-------|------|-------------|--------------|--------|
| **Documentation URL** | String | URL de la documentación | `organization_documentation.documentation_url` | ✅ YA EXISTE |
| **Endpoint enabled** | Boolean | ¿Endpoint público habilitado? | ❌ NO EXISTE | 🔴 AGREGAR |
| **Requires auth** | Boolean | ¿Requiere autenticación? | ❌ NO EXISTE | 🔴 AGREGAR |

**Recomendación:**
```sql
CREATE TABLE organization_integrations (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL UNIQUE,
    
    -- ===== POWER BI SERVICE =====
    powerbi_enabled BIT NOT NULL DEFAULT 0,
    powerbi_workspace_id VARCHAR(255) NULL,
    powerbi_tenant_id VARCHAR(255) NULL,
    powerbi_access_token_encrypted VARBINARY(MAX) NULL,
    powerbi_refresh_token_encrypted VARBINARY(MAX) NULL,
    powerbi_token_expires_at DATETIME2 NULL,
    powerbi_last_sync DATETIME2 NULL,
    
    -- ===== API ACCESS (Enterprise) =====
    api_enabled BIT NOT NULL DEFAULT 0,
    api_key VARCHAR(64) NULL UNIQUE, -- pk_live_xxx
    api_secret_hash VARCHAR(255) NULL, -- Hashed
    api_rate_limit_per_hour INT NOT NULL DEFAULT 1000,
    api_rate_limit_per_day INT NOT NULL DEFAULT 10000,
    api_last_used DATETIME2 NULL,
    api_total_calls INT NOT NULL DEFAULT 0,
    
    -- ===== DOCUMENTATION API =====
    documentation_api_enabled BIT NOT NULL DEFAULT 0,
    documentation_requires_auth BIT NOT NULL DEFAULT 1,
    documentation_public_key VARCHAR(64) NULL, -- Para autenticación pública
    
    -- ===== METADATA =====
    metadata JSON NULL,
    
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_integrations_org FOREIGN KEY (organization_id) REFERENCES organizations(id) ON DELETE CASCADE
);

CREATE INDEX idx_integrations_org ON organization_integrations(organization_id);
CREATE INDEX idx_integrations_api_key ON organization_integrations(api_key) WHERE api_key IS NOT NULL;

-- Trigger para updated_at
CREATE TRIGGER trg_org_integrations_updated_at
ON organization_integrations
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE organization_integrations
    SET updated_at = GETUTCDATE()
    FROM organization_integrations oi
    INNER JOIN inserted i ON oi.id = i.id;
END;
GO
```

**Estado:** 🔴 NECESITA tabla nueva

---

---

## 📊 RESUMEN FINAL (VERSIÓN SIMPLIFICADA)

### **Total de campos en Settings (después de simplificar):** ~25-30 campos

### **Cubiertos por DB actual:**
```
✅ Profile: 2/2 campos (100%)
✅ Security: 1/1 campo (100%)
✅ Organization: 6/6 campos (100%)
✅ Billing (core): 8/8 campos (100%)
❌ Billing (payment): 0/3 campos (0%)
❌ Billing (invoices): 0/5 campos (0%)
❌ Integrations: 0/15 campos (0%)
```

**TOTAL CUBIERTO:** ~17/30 = **57%** ✅

---

## 🎯 TABLAS NUEVAS NECESARIAS (VERSIÓN SIMPLIFICADA)

### **🔴 CRÍTICA - Para cuando implementes Stripe Checkout:**

#### **1. `stripe_invoices`** (5-8 campos mínimos)
```sql
-- Versión minimalista
CREATE TABLE stripe_invoices (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL,
    stripe_invoice_id VARCHAR(255) NOT NULL UNIQUE,
    stripe_invoice_url VARCHAR(500) NULL, -- PDF URL
    amount_paid DECIMAL(10, 2) NOT NULL,
    status VARCHAR(50) NOT NULL,
    paid_at DATETIME2 NULL,
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE()
);
```

**Cubre:** Billing > Invoice history

#### **2. `organization_integrations`** (15+ campos)
```sql
-- Tabla completa de integraciones
-- (ver arriba para DDL completo)
```

**Cubre:** Integrations > Todas las subsecciones

### **🟡 ÚTIL - Para mejorar tracking:**

#### **3. Modificar `organizations`** (1 campo)
```sql
ALTER TABLE organizations 
ADD stripe_payment_method_id VARCHAR(255) NULL;
```

**Cubre:** Billing > Payment method

#### **4. `checkout_sessions`** (opcional - tracking de conversiones)
```sql
-- Solo si querés analytics de checkout
-- No es crítico para la funcionalidad básica
```

---

## ✅ VEREDICTO SIMPLIFICADO

### **Situación actual:**
- ✅ Profile: **100% cubierto** (solo nombre y email)
- ✅ Security: **100% cubierto** (solo LinkedIn status)
- ✅ Organization: **100% cubierto** (nombre, logo, roles, delete)
- ✅ Billing (info básica): **100% cubierto** (plan, status, limits)
- 🔴 Billing (payment & invoices): **0% cubierto** → Necesita tablas nuevas
- 🔴 Integrations: **0% cubierto** → Necesita tabla nueva

### **Para que Settings esté 100% funcional:**

**AHORA (fase beta gratuita):**
- ✅ No necesitás cambiar nada
- Todo lo esencial ya está cubierto

**CUANDO IMPLEMENTES STRIPE:**
- 🔴 Agregar: `stripe_invoices` (historial de facturas)
- 🟡 Agregar a `organizations`: `stripe_payment_method_id`

**CUANDO HABILITES INTEGRACIONES:**
- 🔴 Agregar: `organization_integrations` (Power BI, API, etc.)

---

## 💡 PLAN DE ACCIÓN RECOMENDADO

### **FASE 0: Ahora (Beta gratuita)**
→ ✅ **No hacer nada**  
→ Settings funciona perfecto con tu DB actual  
→ Las secciones de Payment e Integrations muestran "No disponible" correctamente

### **FASE 1: Al implementar Stripe (1-2 meses)**
→ 🔴 Crear `stripe_invoices`  
→ 🟡 Agregar `stripe_payment_method_id` a `organizations`  
→ **Tiempo estimado:** 1 hora  
→ **Beneficio:** Historial de facturas funcional en Settings

### **FASE 2: Al habilitar API/Integraciones (3-6 meses)**
→ 🔴 Crear `organization_integrations`  
→ **Tiempo estimado:** 2 horas  
→ **Beneficio:** Power BI Service, API access, Documentation endpoint

---

## 🎯 CONCLUSIÓN

### **Tu DB actual cubre:**
- ✅ 3 de 5 categorías de Settings al 100%
- ✅ Todo lo esencial para fase beta
- ✅ Todo lo necesario para mostrar info de planes y billing

### **Lo que falta:**
- Solo necesitas tablas nuevas cuando **actives features nuevas**
- No es urgente para beta
- Podés agregarlas de forma incremental

### **Tu arquitectura actual es sólida y extensible** ✅

---

## 📁 Archivos relacionados

- `schema.sql` - Schema principal con tablas core
- `useful_queries.sql` - Queries para subscriptions
- `state_machine_and_workflows.sql` - Lógica de subscriptions
- `organization_workflows.sql` - Lógica de organizaciones
- `INSTALLATION_ORDER.md` - Orden de ejecución

---

**✅ Conclusión: Tu DB está lista para Stripe Checkout sin cambios urgentes.**  
**Solo necesitás agregar tablas cuando habilites features específicas.**







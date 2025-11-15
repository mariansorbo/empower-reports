# 📊 SETTINGS DATABASE MAPPING - COMPLETE ANALYSIS

Análisis exhaustivo de todas las categorías y campos del panel de Settings, 
mapeando dónde debe guardarse cada dato en la base de datos.

---

## 📑 ÍNDICE DE CATEGORÍAS

1. [👤 Profile](#1-profile)
2. [🔐 Security](#2-security)
3. [🏢 Organization](#3-organization)
4. [💳 Billing & Plan](#4-billing--plan)
5. [🔌 Integrations](#5-integrations)
6. [⚙️ Preferences](#6-preferences)

---

## 1. 👤 PROFILE

**Descripción:** Configuración individual de la cuenta del usuario

### Campos actuales:

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **Full name** | String | Nombre completo del usuario | `users.name` | ✅ YA EXISTE |
| **Email** | String | Email (read-only de LinkedIn) | `users.email` | ✅ YA EXISTE |
| **Photo / Avatar** | URL | URL de la imagen de perfil | ❌ NO EXISTE | 🔴 AGREGAR: `users.avatar_url` o `user_preferences.avatar_url` |

### Recomendación:
```sql
-- OPCIÓN 1: Agregar directamente a users (simple)
ALTER TABLE users 
ADD avatar_url VARCHAR(500) NULL;

-- OPCIÓN 2: Crear user_preferences (mejor para escalabilidad)
-- Ver más abajo en la propuesta completa
```

**Estado:** 🟡 Necesita 1 campo nuevo

---

## 2. 🔐 SECURITY

**Descripción:** Cuentas vinculadas y seguridad

### Subcategorías actuales:

#### 2.1. **Cuentas vinculadas**

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **LinkedIn status** | Boolean | ¿Cuenta LinkedIn conectada? | `users.auth_provider` | ✅ YA EXISTE (implícito) |
| **Provider type** | String | Google/LinkedIn/Azure AD | `users.auth_provider` | ✅ YA EXISTE |

**Estado:** ✅ COMPLETO (usando `users.auth_provider`)

---

## 3. 🏢 ORGANIZATION

**Descripción:** Configuración de la organización

### Subcategorías actuales:

#### 3.1. **Información básica**

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **Organization name** | String | Nombre de la organización | `organizations.name` | ✅ YA EXISTE |
| **Logo URL** | String | URL del logo | `organizations.logo_url` | ✅ YA EXISTE |
| **Slug** | String | URL-friendly identifier | `organizations.slug` | ✅ YA EXISTE |
| **Website** | String | Sitio web corporativo | `organizations.website` | ✅ YA EXISTE |

#### 3.2. **Roles y permisos** (info estática)
- No se guarda, es documentación de referencia

#### 3.3. **Danger Zone - Delete organization**
- No es un campo, es una acción (usa `organizations.is_archived`)

**Estado:** ✅ COMPLETO

---

## 4. 💳 BILLING & PLAN

**Descripción:** Control de suscripción, pagos y upgrades

### Subcategorías y campos:

#### 4.1. **Plan actual**

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **Plan name** | String | free_trial/basic/teams/enterprise | `subscriptions.plan_id` | ✅ YA EXISTE |
| **Status** | String | active/trialing/canceled/past_due | `subscriptions.status` | ✅ YA EXISTE |
| **Current period start** | DateTime | Inicio del período actual | `subscriptions.current_period_start` | ✅ YA EXISTE |
| **Current period end** | DateTime | Fin del período actual | `subscriptions.current_period_end` | ✅ YA EXISTE |
| **Billing cycle** | String | monthly/yearly | `subscriptions.billing_cycle` | ✅ YA EXISTE |

#### 4.2. **Límites actuales**

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **Max users** | Integer | Límite de usuarios del plan | `plans.max_users` | ✅ YA EXISTE |
| **Max reports** | Integer | Límite de reportes del plan | `plans.max_reports` | ✅ YA EXISTE |
| **Current users count** | Integer (calculado) | Usuarios actuales | Query a `organization_members` | ✅ YA EXISTE (via view) |
| **Current reports count** | Integer (calculado) | Reportes actuales | Query a `reports` | ✅ YA EXISTE (via view) |

#### 4.3. **Método de pago** (futuro)

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **Stripe Customer ID** | String | ID del customer en Stripe | `organizations.stripe_customer_id` | ✅ YA EXISTE |
| **Default payment method** | String | pm_xxx (Stripe) | ❌ NO EXISTE | 🟡 AGREGAR: `organizations.stripe_payment_method_id` |
| **Card brand** | String | Visa/Mastercard/Amex | ❌ NO EXISTE | 🟡 OPCIONAL (se obtiene de Stripe API on-demand) |
| **Last 4 digits** | String | **** 4242 | ❌ NO EXISTE | 🟡 OPCIONAL (se obtiene de Stripe API) |

**Recomendación:**
```sql
-- OPCIÓN A: Agregar a organizations (mínimo necesario)
ALTER TABLE organizations 
ADD stripe_payment_method_id VARCHAR(255) NULL;

-- OPCIÓN B: Tabla separada (si querés múltiples métodos de pago)
CREATE TABLE payment_methods (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL,
    stripe_payment_method_id VARCHAR(255) NOT NULL,
    card_brand VARCHAR(50) NULL, -- visa, mastercard, amex
    card_last4 VARCHAR(4) NULL,
    card_exp_month INT NULL,
    card_exp_year INT NULL,
    is_default BIT NOT NULL DEFAULT 0,
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_payment_methods_org FOREIGN KEY (organization_id) REFERENCES organizations(id)
);
```

#### 4.4. **Historial de facturación** (futuro)

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **Invoice ID** | String | ID de factura en Stripe | ❌ NO EXISTE | 🔴 AGREGAR: Tabla `stripe_invoices` |
| **Invoice PDF URL** | String | URL del PDF descargable | ❌ NO EXISTE | 🔴 AGREGAR: `stripe_invoices.invoice_url` |
| **Amount paid** | Decimal | Monto pagado | ❌ NO EXISTE | 🔴 AGREGAR: `stripe_invoices.amount_paid` |
| **Payment date** | DateTime | Fecha de pago | ❌ NO EXISTE | 🔴 AGREGAR: `stripe_invoices.paid_at` |
| **Status** | String | paid/open/void | ❌ NO EXISTE | 🔴 AGREGAR: `stripe_invoices.status` |

**Recomendación:**
```sql
CREATE TABLE stripe_invoices (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL,
    subscription_id UNIQUEIDENTIFIER NOT NULL,
    stripe_invoice_id VARCHAR(255) NOT NULL UNIQUE,
    stripe_invoice_url VARCHAR(500) NULL, -- PDF download URL
    stripe_hosted_url VARCHAR(500) NULL, -- Web view URL
    amount_due DECIMAL(10, 2) NOT NULL,
    amount_paid DECIMAL(10, 2) NOT NULL DEFAULT 0,
    currency VARCHAR(3) NOT NULL DEFAULT 'USD',
    status VARCHAR(50) NOT NULL CHECK (status IN ('draft', 'open', 'paid', 'void', 'uncollectible')),
    billing_reason VARCHAR(50) NULL, -- subscription_create, subscription_cycle, etc.
    period_start DATETIME2 NOT NULL,
    period_end DATETIME2 NOT NULL,
    paid_at DATETIME2 NULL,
    due_date DATETIME2 NULL,
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_invoices_organization FOREIGN KEY (organization_id) REFERENCES organizations(id),
    CONSTRAINT fk_invoices_subscription FOREIGN KEY (subscription_id) REFERENCES subscriptions(id)
);

CREATE INDEX idx_invoices_organization ON stripe_invoices(organization_id);
CREATE INDEX idx_invoices_status ON stripe_invoices(status);
CREATE INDEX idx_invoices_paid_at ON stripe_invoices(paid_at);
```

**Estado:** 🔴 NECESITA tabla nueva (para mostrar historial en UI)

---

## 5. 🔌 INTEGRATIONS

**Descripción:** Conexiones con herramientas externas

### Subcategorías actuales:

#### 5.1. **Power BI Service**

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **Workspace ID** | String | ID del workspace de Power BI | ❌ NO EXISTE | 🔴 AGREGAR |
| **Connection status** | Boolean | ¿Conectado o no? | ❌ NO EXISTE | 🔴 AGREGAR |
| **Last sync** | DateTime | Última sincronización | ❌ NO EXISTE | 🔴 AGREGAR |

#### 5.2. **API - Upload .pbit**

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **API Key** | String | Token para autenticación | ❌ NO EXISTE | 🔴 AGREGAR |
| **API secret** | String | Secret key (hashed) | ❌ NO EXISTE | 🔴 AGREGAR |
| **Enabled** | Boolean | ¿API habilitada? | ❌ NO EXISTE | 🔴 AGREGAR |

#### 5.3. **API - Documentation endpoint**

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **Endpoint URL** | String | URL del endpoint público | Ya existe en `organization_documentation` | ✅ YA EXISTE |
| **API enabled** | Boolean | ¿Endpoint activo? | ❌ NO EXISTE | 🟡 AGREGAR |

**Recomendación:**
```sql
-- Tabla unificada para todas las integraciones
CREATE TABLE organization_integrations (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL UNIQUE,
    
    -- Power BI Service
    powerbi_workspace_id VARCHAR(255) NULL,
    powerbi_connected BIT NOT NULL DEFAULT 0,
    powerbi_last_sync DATETIME2 NULL,
    powerbi_access_token_encrypted VARBINARY(MAX) NULL, -- Token cifrado
    
    -- API Access (Enterprise feature)
    api_enabled BIT NOT NULL DEFAULT 0,
    api_key VARCHAR(64) NULL, -- Public key
    api_secret_hash VARCHAR(255) NULL, -- Hashed secret
    api_rate_limit INT NOT NULL DEFAULT 1000, -- Requests per hour
    api_last_used DATETIME2 NULL,
    
    -- Documentation API
    documentation_api_enabled BIT NOT NULL DEFAULT 0,
    documentation_endpoint_public BIT NOT NULL DEFAULT 0,
    
    -- Slack (futuro)
    slack_webhook_url VARCHAR(500) NULL,
    slack_channel VARCHAR(100) NULL,
    
    -- Microsoft Teams (futuro)
    teams_webhook_url VARCHAR(500) NULL,
    
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_integrations_organization FOREIGN KEY (organization_id) REFERENCES organizations(id) ON DELETE CASCADE
);

CREATE INDEX idx_integrations_org ON organization_integrations(organization_id);
CREATE INDEX idx_integrations_api_key ON organization_integrations(api_key) WHERE api_key IS NOT NULL;
```

**Estado:** 🔴 NECESITA tabla nueva

---

## 6. ⚙️ PREFERENCES

**Descripción:** Personalización de la experiencia del usuario

### Subcategorías actuales:

#### 6.1. **🌐 Regionalization**

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **Date format** | Enum | dd/mm/yyyy, mm/dd/yyyy, yyyy-mm-dd | ❌ NO EXISTE | 🔴 AGREGAR: `user_preferences.date_format` |
| **Number format** | Enum | es (1.234,56) o en (1,234.56) | ❌ NO EXISTE | 🔴 AGREGAR: `user_preferences.number_format` |

#### 6.2. **📊 Report visualization**

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **Show automatic preview** | Boolean | Vista previa automática al cargar | ❌ NO EXISTE | 🔴 AGREGAR: `user_preferences.auto_preview` |
| **Highlight unused fields** | Boolean | Resaltar tablas/campos no usados | ❌ NO EXISTE | 🔴 AGREGAR: `user_preferences.highlight_unused_fields` |
| **Expand relationships** | Boolean | Expandir relaciones por defecto | ❌ NO EXISTE | 🔴 AGREGAR: `user_preferences.expand_relationships` |

#### 6.3. **💬 User experience**

| Campo | Tipo | Descripción | Tabla actual | Acción requerida |
|-------|------|-------------|--------------|------------------|
| **Show tips** | Boolean | Mostrar tips y sugerencias | ❌ NO EXISTE | 🔴 AGREGAR: `user_preferences.show_tips` |
| **Enable animations** | Boolean | Habilitar animaciones y transiciones | ❌ NO EXISTE | 🔴 AGREGAR: `user_preferences.enable_animations` |
| **Compact mode** | Boolean | Modo compacto (reduce espaciado) | ❌ NO EXISTE | 🔴 AGREGAR: `user_preferences.compact_mode` |

**Estado:** 🔴 NECESITA tabla `user_preferences` completa

---

---

## 📊 RESUMEN POR TABLA

### ✅ TABLAS EXISTENTES QUE YA CUBREN SETTINGS:

#### **`users`** (parcial)
- ✅ `name` → Profile: Full name
- ✅ `email` → Profile: Email
- ✅ `auth_provider` → Security: Linked accounts
- ❌ Falta: `avatar_url`

#### **`organizations`** (casi completo)
- ✅ `name` → Organization: Name
- ✅ `logo_url` → Organization: Logo
- ✅ `slug` → Organization: URL identifier
- ✅ `website` → Organization: Website
- ✅ `stripe_customer_id` → Billing: Customer
- ❌ Falta: `stripe_payment_method_id` (método de pago default)

#### **`subscriptions`** (completo para billing)
- ✅ `plan_id` → Billing: Current plan
- ✅ `status` → Billing: Status
- ✅ `billing_cycle` → Billing: Cycle
- ✅ `current_period_start/end` → Billing: Period
- ✅ `stripe_subscription_id` → Billing: Stripe sub ID
- ✅ `stripe_price_id` → Billing: Stripe price ID

#### **`plans`** (completo)
- ✅ `max_users` → Billing: User limit
- ✅ `max_reports` → Billing: Report limit
- ✅ `price_monthly/yearly` → Billing: Pricing
- ✅ `stripe_price_id_monthly/yearly` → Billing: Stripe prices

#### **`subscription_history`** (completo para auditoría)
- ✅ `event_type` → Billing: Change tracking
- ✅ `stripe_event_id` → Billing: Webhook events
- ✅ `plan_id_old/new` → Billing: Upgrades/downgrades

#### **`organization_documentation`** (completo)
- ✅ `documentation_url` → Integrations: Docs URL
- ✅ `is_active` → Integrations: Docs enabled

---

### 🔴 TABLAS NUEVAS NECESARIAS:

#### **1. `user_preferences`** (🔴 CRÍTICA)
```
Campos: 8 nuevos
- avatar_url
- date_format
- number_format
- auto_preview
- highlight_unused_fields
- expand_relationships
- show_tips
- enable_animations
- compact_mode
```

#### **2. `organization_integrations`** (🔴 CRÍTICA)
```
Campos: 10+ nuevos
- powerbi_workspace_id
- powerbi_connected
- powerbi_last_sync
- api_enabled
- api_key
- api_secret_hash
- documentation_api_enabled
- slack_webhook_url
- teams_webhook_url
```

#### **3. `stripe_invoices`** (🟡 IMPORTANTE pero puede esperar)
```
Campos: 12 nuevos
- stripe_invoice_id
- stripe_invoice_url
- amount_due
- amount_paid
- currency
- status
- period_start/end
- paid_at
- due_date
```

#### **4. `checkout_sessions`** (🟡 OPCIONAL pero útil)
```
Campos: 8 nuevos
- stripe_session_id
- plan_id
- billing_cycle
- amount
- status
- expires_at
- completed_at
```

#### **5. `payment_methods`** (🟡 OPCIONAL para múltiples tarjetas)
```
Campos: 9 nuevos
- stripe_payment_method_id
- card_brand
- card_last4
- card_exp_month
- card_exp_year
- is_default
```

---

## 📈 MATRIZ DE PRIORIDADES

| Tabla | Prioridad | Cuándo la necesitás | Impacto en Settings |
|-------|-----------|---------------------|---------------------|
| `user_preferences` | 🔴 ALTA | Ahora (Settings ya lo usa) | Sin ella, no guardás preferencias de UI |
| `organization_integrations` | 🔴 ALTA | Ahora (para API keys) | Sin ella, integraciones no funcionan |
| `stripe_invoices` | 🟡 MEDIA | Al implementar Stripe | Sin ella, no mostrás historial de facturas |
| `checkout_sessions` | 🟡 MEDIA | Al implementar Stripe | Sin ella, no trackeas conversiones |
| `payment_methods` | 🟢 BAJA | Solo si permitís múltiples tarjetas | Stripe Portal puede manejar esto |

---

## 🎯 PROPUESTA COMPLETA DE TABLAS

### **📋 1. user_preferences (CRÍTICA)**

```sql
CREATE TABLE user_preferences (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    user_id UNIQUEIDENTIFIER NOT NULL UNIQUE,
    
    -- ===== PROFILE =====
    avatar_url VARCHAR(500) NULL,
    
    -- ===== PREFERENCES > REGIONALIZATION =====
    date_format VARCHAR(20) NOT NULL DEFAULT 'dd/mm/yyyy' 
        CHECK (date_format IN ('dd/mm/yyyy', 'mm/dd/yyyy', 'yyyy-mm-dd')),
    number_format VARCHAR(10) NOT NULL DEFAULT 'es' 
        CHECK (number_format IN ('es', 'en')),
    
    -- ===== PREFERENCES > REPORT VISUALIZATION =====
    auto_preview BIT NOT NULL DEFAULT 1,
    highlight_unused_fields BIT NOT NULL DEFAULT 1,
    expand_relationships BIT NOT NULL DEFAULT 0,
    
    -- ===== PREFERENCES > USER EXPERIENCE =====
    show_tips BIT NOT NULL DEFAULT 1,
    enable_animations BIT NOT NULL DEFAULT 1,
    compact_mode BIT NOT NULL DEFAULT 0,
    
    -- ===== METADATA (flexible para futuro) =====
    metadata JSON NULL,
    
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_user_prefs_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);
```

**Cobertura:** 
- ✅ Profile: Avatar
- ✅ Preferences: Regionalization (2 campos)
- ✅ Preferences: Report visualization (3 campos)
- ✅ Preferences: User experience (3 campos)

---

### **📋 2. organization_integrations (CRÍTICA)**

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
    api_key VARCHAR(64) NULL UNIQUE, -- Public key: pk_live_xxx
    api_secret_hash VARCHAR(255) NULL, -- Hashed secret
    api_rate_limit_per_hour INT NOT NULL DEFAULT 1000,
    api_rate_limit_per_day INT NOT NULL DEFAULT 10000,
    api_last_used DATETIME2 NULL,
    api_total_calls INT NOT NULL DEFAULT 0,
    
    -- ===== DOCUMENTATION API =====
    documentation_api_enabled BIT NOT NULL DEFAULT 0,
    documentation_public BIT NOT NULL DEFAULT 0, -- ¿Endpoint público o requiere auth?
    
    -- ===== SLACK (futuro) =====
    slack_enabled BIT NOT NULL DEFAULT 0,
    slack_webhook_url VARCHAR(500) NULL,
    slack_channel VARCHAR(100) NULL,
    slack_connected_by UNIQUEIDENTIFIER NULL,
    
    -- ===== MICROSOFT TEAMS (futuro) =====
    teams_enabled BIT NOT NULL DEFAULT 0,
    teams_webhook_url VARCHAR(500) NULL,
    teams_channel VARCHAR(100) NULL,
    teams_connected_by UNIQUEIDENTIFIER NULL,
    
    -- ===== METADATA =====
    metadata JSON NULL,
    
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_integrations_org FOREIGN KEY (organization_id) REFERENCES organizations(id) ON DELETE CASCADE,
    CONSTRAINT fk_integrations_slack_user FOREIGN KEY (slack_connected_by) REFERENCES users(id),
    CONSTRAINT fk_integrations_teams_user FOREIGN KEY (teams_connected_by) REFERENCES users(id)
);
```

**Cobertura:**
- ✅ Integrations: Power BI Service (7 campos)
- ✅ Integrations: API Upload (7 campos)
- ✅ Integrations: Documentation API (2 campos)
- ✅ Integrations: Slack (4 campos)
- ✅ Integrations: Teams (4 campos)

---

### **📋 3. stripe_invoices (IMPORTANTE)**

```sql
CREATE TABLE stripe_invoices (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL,
    subscription_id UNIQUEIDENTIFIER NOT NULL,
    
    -- ===== STRIPE DATA =====
    stripe_invoice_id VARCHAR(255) NOT NULL UNIQUE,
    stripe_customer_id VARCHAR(255) NOT NULL,
    stripe_subscription_id VARCHAR(255) NOT NULL,
    stripe_invoice_pdf VARCHAR(500) NULL, -- PDF download URL
    stripe_hosted_invoice_url VARCHAR(500) NULL, -- Web view
    
    -- ===== AMOUNTS =====
    subtotal DECIMAL(10, 2) NOT NULL,
    tax DECIMAL(10, 2) NOT NULL DEFAULT 0,
    amount_due DECIMAL(10, 2) NOT NULL,
    amount_paid DECIMAL(10, 2) NOT NULL DEFAULT 0,
    amount_remaining DECIMAL(10, 2) NOT NULL DEFAULT 0,
    currency VARCHAR(3) NOT NULL DEFAULT 'USD',
    
    -- ===== STATUS & DATES =====
    status VARCHAR(50) NOT NULL CHECK (status IN ('draft', 'open', 'paid', 'void', 'uncollectible')),
    billing_reason VARCHAR(50) NULL CHECK (billing_reason IN ('subscription_create', 'subscription_cycle', 'subscription_update', 'manual')),
    period_start DATETIME2 NOT NULL,
    period_end DATETIME2 NOT NULL,
    due_date DATETIME2 NULL,
    paid_at DATETIME2 NULL,
    
    -- ===== METADATA =====
    description TEXT NULL,
    metadata JSON NULL,
    
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_invoices_org FOREIGN KEY (organization_id) REFERENCES organizations(id),
    CONSTRAINT fk_invoices_subscription FOREIGN KEY (subscription_id) REFERENCES subscriptions(id)
);

CREATE INDEX idx_invoices_org ON stripe_invoices(organization_id);
CREATE INDEX idx_invoices_subscription ON stripe_invoices(subscription_id);
CREATE INDEX idx_invoices_status ON stripe_invoices(status);
CREATE INDEX idx_invoices_paid_at ON stripe_invoices(paid_at);
CREATE INDEX idx_invoices_stripe_id ON stripe_invoices(stripe_invoice_id);
```

**Cobertura:**
- ✅ Billing: Historial completo de facturas (17 campos)

---

### **📋 4. checkout_sessions (ÚTIL para tracking)**

```sql
CREATE TABLE checkout_sessions (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL,
    user_id UNIQUEIDENTIFIER NOT NULL, -- Quién inició el checkout
    
    -- ===== STRIPE DATA =====
    stripe_session_id VARCHAR(255) NOT NULL UNIQUE,
    stripe_customer_id VARCHAR(255) NULL, -- Puede ser NULL si es nuevo customer
    
    -- ===== PLAN DATA =====
    plan_id VARCHAR(50) NOT NULL,
    billing_cycle VARCHAR(20) NOT NULL CHECK (billing_cycle IN ('monthly', 'yearly')),
    
    -- ===== AMOUNTS =====
    amount DECIMAL(10, 2) NOT NULL,
    currency VARCHAR(3) NOT NULL DEFAULT 'USD',
    
    -- ===== STATUS & TRACKING =====
    status VARCHAR(50) NOT NULL CHECK (status IN ('pending', 'completed', 'expired', 'canceled')) DEFAULT 'pending',
    payment_status VARCHAR(50) NULL CHECK (payment_status IN ('paid', 'unpaid', 'no_payment_required')),
    
    -- ===== URLs =====
    success_url VARCHAR(500) NOT NULL,
    cancel_url VARCHAR(500) NOT NULL,
    checkout_url VARCHAR(500) NULL, -- URL de Stripe Checkout
    
    -- ===== DATES =====
    expires_at DATETIME2 NOT NULL, -- Checkout sessions expire in 24h
    completed_at DATETIME2 NULL,
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    -- ===== METADATA =====
    metadata JSON NULL,
    
    CONSTRAINT fk_checkout_org FOREIGN KEY (organization_id) REFERENCES organizations(id),
    CONSTRAINT fk_checkout_user FOREIGN KEY (user_id) REFERENCES users(id),
    CONSTRAINT fk_checkout_plan FOREIGN KEY (plan_id) REFERENCES plans(id)
);

CREATE INDEX idx_checkout_stripe_session ON checkout_sessions(stripe_session_id);
CREATE INDEX idx_checkout_org ON checkout_sessions(organization_id);
CREATE INDEX idx_checkout_status ON checkout_sessions(status);
CREATE INDEX idx_checkout_expires ON checkout_sessions(expires_at);
```

**Cobertura:**
- ✅ Billing: Tracking de conversiones
- ✅ Billing: Debugging de pagos fallidos

---

## 📊 RESUMEN EJECUTIVO

### **Campos totales en Settings:** ~35-40 campos

### **Cubiertos por DB actual:** ~15 campos (43%)
- ✅ users: 3 campos
- ✅ organizations: 5 campos
- ✅ subscriptions: 10+ campos
- ✅ plans: 5+ campos

### **FALTAN:** ~20-25 campos (57%)

### **Distribución de campos faltantes:**

| Tabla nueva | Campos | Prioridad | Secciones que cubre |
|-------------|--------|-----------|---------------------|
| `user_preferences` | 9 | 🔴 ALTA | Profile (1), Preferences (8) |
| `organization_integrations` | 15+ | 🔴 ALTA | Integrations (15+) |
| `stripe_invoices` | 17 | 🟡 MEDIA | Billing: Invoice history |
| `checkout_sessions` | 14 | 🟡 MEDIA | Billing: Checkout tracking |
| `payment_methods` | 7 | 🟢 BAJA | Billing: Multiple cards (opcional) |

---

## ✅ RECOMENDACIÓN FINAL

### **Implementar YA (mínimo viable):**
1. 🔴 `user_preferences` - 9 campos
2. 🔴 `organization_integrations` - 15 campos
3. 🟡 Agregar a `organizations`: `stripe_payment_method_id`

### **Implementar cuando implementes Stripe:**
4. 🟡 `stripe_invoices` - 17 campos
5. 🟡 `checkout_sessions` - 14 campos

### **Implementar después (nice to have):**
6. 🟢 `payment_methods` - Solo si permitís múltiples tarjetas

---

## 🤔 ¿Quieres que cree estos scripts SQL?

Puedo generar:
1. ✅ Scripts SQL de las 2 tablas críticas (`user_preferences`, `organization_integrations`)
2. ✅ Scripts SQL de las tablas de Stripe (`stripe_invoices`, `checkout_sessions`)
3. ✅ Stored procedures para CRUD de cada tabla
4. ✅ Triggers de `updated_at`
5. ✅ Valores por defecto al crear usuario/organización
6. ✅ Actualizar `INSTALLATION_ORDER.md`
7. ✅ Actualizar documentación

**¿Arranco con las 2 tablas críticas primero?**






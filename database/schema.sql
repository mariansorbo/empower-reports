--============================================================================
-- REPORT TUNER - Database Schema (SQL Server / Azure SQL Database)
--============================================================================
-- Sistema SaaS para documentación de reportes de Power BI
-- Soporta modo Free Trial y planes comerciales (Basic, Teams, Enterprise, Enterprise Pro)
-- 
-- ESQUEMA SIMPLIFICADO: Solo lo esencial
-- A/B Testing: Se maneja con HubSpot
-- Pricing complejo: Se maneja con Stripe + HubSpot
-- 
-- Ejecutar también:
--   - organization_workflows.sql (procedimientos de creación/unión)
--   - state_machine_and_workflows.sql (validaciones y workflows)
--   - constraints_and_validations.sql (validaciones adicionales)
--   - documentation_procedures.sql (procedures para documentación)
--   - enterprise_pro_plan_v2.sql (solo si necesitas Enterprise Pro)
--============================================================================
--
-- ⚠️ IMPORTANTE PARA AZURE SQL DATABASE:
-- 
-- Si estás usando Azure SQL Database, la base de datos YA DEBE EXISTIR.
-- COMENTA o ELIMINA las siguientes líneas antes de ejecutar este script:
--   - USE master;
--   - CREATE DATABASE empower_reports;
--   - USE empower_reports;
--
-- En su lugar, asegúrate de estar conectado a tu base de datos de Azure SQL
-- y cambia el nombre de la base de datos en la línea USE si es necesario.
--
-- Para más detalles, consulta: DATABASE_DEPLOYMENT.md
--============================================================================

-- ============================================================================
-- CREACIÓN DE BASE DE DATOS (SOLO PARA SQL SERVER LOCAL/DOCKER)
-- ============================================================================
-- COMENTAR ESTAS LÍNEAS SI USAS AZURE SQL DATABASE
-- ============================================================================
USE master;
GO

-- Crear base de datos si no existe (solo para SQL Server local/Docker)
IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'empower_reports')
BEGIN
    CREATE DATABASE empower_reports;
END
GO

USE empower_reports;
GO
-- ============================================================================
-- FIN DE CREACIÓN DE BASE DE DATOS
-- ============================================================================
--
-- Para Azure SQL Database:
-- 1. Conéctate directamente a tu base de datos (ej: report-tuner-db)
-- 2. NO ejecutes las líneas de arriba (USE master, CREATE DATABASE, etc.)
-- 3. Ejecuta solo las tablas, vistas, funciones y triggers de abajo
-- ============================================================================

-- ============================================================================
-- TABLA: plans
-- ============================================================================
-- Define los planes disponibles con sus límites y características
-- ============================================================================

IF OBJECT_ID('plans', 'U') IS NOT NULL DROP TABLE plans;
GO

CREATE TABLE plans (
    id VARCHAR(50) PRIMARY KEY NOT NULL,
    name VARCHAR(100) NOT NULL,
    description TEXT,
    max_users INT NOT NULL DEFAULT 1,
    max_reports INT NOT NULL DEFAULT 10,
    max_storage_mb INT NOT NULL DEFAULT 100,
    features JSON, -- JSON con features habilitadas: {"api_access": true, "branding": false, "audit_log": true}
    price_monthly DECIMAL(10, 2) NULL, -- NULL para free_trial
    price_yearly DECIMAL(10, 2) NULL, -- NULL para free_trial
    stripe_price_id_monthly VARCHAR(255) NULL,
    stripe_price_id_yearly VARCHAR(255) NULL,
    max_organizations INT NULL, -- Máximo de organizaciones gestionadas (solo Enterprise Pro), NULL = no aplica
    is_active BIT NOT NULL DEFAULT 1,
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE()
);
GO

-- Índice para búsqueda por nombre
CREATE INDEX idx_plans_name ON plans(name);
GO

-- Comentarios en la tabla
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Tabla de planes disponibles. Define límites de usuarios, reportes y características por plan.', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'plans';
GO

-- ============================================================================
-- TABLA: users
-- ============================================================================
-- Usuarios del sistema con autenticación por Google, LinkedIn o Azure AD
-- ============================================================================

IF OBJECT_ID('users', 'U') IS NOT NULL DROP TABLE users;
GO

CREATE TABLE users (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    email VARCHAR(255) NOT NULL UNIQUE,
    name VARCHAR(255) NOT NULL,
    avatar_url VARCHAR(500) NULL,
    auth_provider VARCHAR(50) NOT NULL CHECK (auth_provider IN ('google', 'linkedin', 'azure_ad', 'email')), -- 'email' para auth local
    auth_provider_id VARCHAR(255) NULL, -- ID del usuario en el proveedor de auth
    password_hash VARCHAR(255) NULL, -- Solo para auth local, NULL para OAuth
    is_active BIT NOT NULL DEFAULT 1,
    is_email_verified BIT NOT NULL DEFAULT 0,
    last_login_at DATETIME2 NULL,
    metadata JSON NULL, -- Información adicional del proveedor OAuth
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE()
);
GO

-- Índices para búsqueda y performance
CREATE INDEX idx_users_email ON users(email);
CREATE INDEX idx_users_auth_provider ON users(auth_provider, auth_provider_id);
CREATE INDEX idx_users_is_active ON users(is_active);
GO

-- Comentarios
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Usuarios del sistema. Soporta autenticación OAuth (Google, LinkedIn, Azure AD) y auth local.', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'users';
GO

-- ============================================================================
-- TABLA: organizations
-- ============================================================================
-- Organizaciones donde los usuarios colaboran
-- ============================================================================

IF OBJECT_ID('organizations', 'U') IS NOT NULL DROP TABLE organizations;
GO

CREATE TABLE organizations (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    name VARCHAR(255) NOT NULL,
    slug VARCHAR(100) NULL UNIQUE, -- URL-friendly identifier
    logo_url VARCHAR(500) NULL,
    website VARCHAR(255) NULL,
    stripe_customer_id VARCHAR(255) NULL, -- ID del cliente en Stripe
    is_archived BIT NOT NULL DEFAULT 0,
    archived_at DATETIME2 NULL,
    metadata JSON NULL, -- Información adicional de la organización
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE()
);
GO

-- Índices
CREATE INDEX idx_organizations_slug ON organizations(slug);
CREATE INDEX idx_organizations_is_archived ON organizations(is_archived);
CREATE INDEX idx_organizations_stripe_customer_id ON organizations(stripe_customer_id);
GO

-- Comentarios
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Organizaciones donde los usuarios colaboran. Pueden tener múltiples usuarios y una suscripción activa.', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'organizations';
GO

-- ============================================================================
-- TABLA: organization_members
-- ============================================================================
-- Relación muchos a muchos entre usuarios y organizaciones (con roles)
-- ============================================================================

IF OBJECT_ID('organization_members', 'U') IS NOT NULL DROP TABLE organization_members;
GO

CREATE TABLE organization_members (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL,
    user_id UNIQUEIDENTIFIER NOT NULL,
    role VARCHAR(50) NOT NULL DEFAULT 'member' CHECK (role IN ('admin', 'admin_global', 'member', 'viewer')),
    is_primary BIT NOT NULL DEFAULT 0, -- Indica si esta es la organización principal del usuario
    invited_by UNIQUEIDENTIFIER NULL, -- Usuario que invitó
    invitation_token VARCHAR(255) NULL, -- Token para invitaciones pendientes
    invitation_expires_at DATETIME2 NULL,
    joined_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    left_at DATETIME2 NULL,
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    -- Constraint: un usuario solo puede tener una organización primaria
    CONSTRAINT fk_org_members_organization FOREIGN KEY (organization_id) REFERENCES organizations(id) ON DELETE CASCADE,
    CONSTRAINT fk_org_members_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
    CONSTRAINT fk_org_members_invited_by FOREIGN KEY (invited_by) REFERENCES users(id),
    
    -- Constraint: un usuario solo puede pertenecer una vez a la misma organización
    CONSTRAINT uk_org_members_org_user UNIQUE (organization_id, user_id)
);
GO

-- Índices
CREATE INDEX idx_org_members_user_id ON organization_members(user_id);
CREATE INDEX idx_org_members_organization_id ON organization_members(organization_id);
CREATE INDEX idx_org_members_role ON organization_members(role);
CREATE INDEX idx_org_members_is_primary ON organization_members(is_primary);
CREATE INDEX idx_org_members_invitation_token ON organization_members(invitation_token) WHERE invitation_token IS NOT NULL;
GO

-- Comentarios
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Relación entre usuarios y organizaciones con roles. Un usuario puede pertenecer a múltiples organizaciones pero solo una puede ser primaria.', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'organization_members';
GO

-- ============================================================================
-- TABLA: subscriptions
-- ============================================================================
-- Suscripciones activas de organizaciones a planes
-- ============================================================================

IF OBJECT_ID('subscriptions', 'U') IS NOT NULL DROP TABLE subscriptions;
GO

CREATE TABLE subscriptions (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL,
    plan_id VARCHAR(50) NOT NULL,
    status VARCHAR(50) NOT NULL CHECK (status IN ('active', 'trialing', 'canceled', 'past_due', 'unpaid', 'incomplete')),
    billing_cycle VARCHAR(20) NULL CHECK (billing_cycle IS NULL OR billing_cycle IN ('monthly', 'yearly')), -- NULL para free_trial
    current_period_start DATETIME2 NOT NULL,
    current_period_end DATETIME2 NOT NULL,
    cancel_at_period_end BIT NOT NULL DEFAULT 0, -- Si es true, se cancela al final del período
    canceled_at DATETIME2 NULL,
    trial_start DATETIME2 NULL,
    trial_end DATETIME2 NULL,
    stripe_subscription_id VARCHAR(255) NULL UNIQUE, -- ID de la suscripción en Stripe
    stripe_price_id VARCHAR(255) NULL, -- ID del precio en Stripe
    metadata JSON NULL, -- Información adicional de Stripe
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_subscriptions_organization FOREIGN KEY (organization_id) REFERENCES organizations(id) ON DELETE CASCADE,
    CONSTRAINT fk_subscriptions_plan FOREIGN KEY (plan_id) REFERENCES plans(id)
);
GO

-- Índices
CREATE INDEX idx_subscriptions_organization_id ON subscriptions(organization_id);
CREATE INDEX idx_subscriptions_plan_id ON subscriptions(plan_id);
CREATE INDEX idx_subscriptions_status ON subscriptions(status);
CREATE INDEX idx_subscriptions_stripe_subscription_id ON subscriptions(stripe_subscription_id);
CREATE INDEX idx_subscriptions_current_period_end ON subscriptions(current_period_end);
GO

-- Constraint: una organización solo puede tener una suscripción activa/trialing
CREATE UNIQUE INDEX uk_subscriptions_active_org 
ON subscriptions(organization_id) 
WHERE status IN ('active', 'trialing');
GO

-- Comentarios
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Suscripciones de organizaciones a planes. Refleja el estado real de Stripe y permite una suscripción activa por organización.', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'subscriptions';
GO

-- ============================================================================
-- TABLA: subscription_history
-- ============================================================================
-- Historial de cambios de planes, upgrades, downgrades y eventos de Stripe
-- ============================================================================

IF OBJECT_ID('subscription_history', 'U') IS NOT NULL DROP TABLE subscription_history;
GO

CREATE TABLE subscription_history (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    subscription_id UNIQUEIDENTIFIER NOT NULL,
    organization_id UNIQUEIDENTIFIER NOT NULL,
    plan_id_old VARCHAR(50) NULL,
    plan_id_new VARCHAR(50) NOT NULL,
    status_old VARCHAR(50) NULL,
    status_new VARCHAR(50) NOT NULL,
    event_type VARCHAR(50) NOT NULL CHECK (event_type IN ('created', 'updated', 'canceled', 'reactivated', 'plan_changed', 'billing_cycle_changed', 'stripe_webhook')),
    stripe_event_id VARCHAR(255) NULL, -- ID del evento en Stripe para tracking
    metadata JSON NULL, -- Información adicional del cambio
    changed_by UNIQUEIDENTIFIER NULL, -- Usuario que hizo el cambio (NULL si fue por webhook)
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_sub_history_subscription FOREIGN KEY (subscription_id) REFERENCES subscriptions(id) ON DELETE CASCADE,
    CONSTRAINT fk_sub_history_organization FOREIGN KEY (organization_id) REFERENCES organizations(id) ON DELETE CASCADE,
    CONSTRAINT fk_sub_history_plan_new FOREIGN KEY (plan_id_new) REFERENCES plans(id),
    CONSTRAINT fk_sub_history_plan_old FOREIGN KEY (plan_id_old) REFERENCES plans(id),
    CONSTRAINT fk_sub_history_changed_by FOREIGN KEY (changed_by) REFERENCES users(id)
);
GO

-- Índices
CREATE INDEX idx_sub_history_subscription_id ON subscription_history(subscription_id);
CREATE INDEX idx_sub_history_organization_id ON subscription_history(organization_id);
CREATE INDEX idx_sub_history_event_type ON subscription_history(event_type);
CREATE INDEX idx_sub_history_created_at ON subscription_history(created_at);
CREATE INDEX idx_sub_history_stripe_event_id ON subscription_history(stripe_event_id) WHERE stripe_event_id IS NOT NULL;
GO

-- Comentarios
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Historial completo de cambios en suscripciones. Rastrea upgrades, downgrades y eventos de Stripe.', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'subscription_history';
GO

-- ============================================================================
-- TABLA: reports
-- ============================================================================
-- Reportes (.pbit) subidos por los usuarios
-- ============================================================================

IF OBJECT_ID('reports', 'U') IS NOT NULL DROP TABLE reports;
GO

CREATE TABLE reports (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NULL, -- NULL si es usuario individual (plan basic)
    user_id UNIQUEIDENTIFIER NOT NULL, -- Usuario que subió el reporte
    name VARCHAR(255) NOT NULL,
    original_filename VARCHAR(255) NOT NULL,
    file_size_bytes BIGINT NOT NULL,
    file_url VARCHAR(500) NULL, -- URL en Azure Blob Storage
    blob_name VARCHAR(255) NULL, -- Nombre del blob en Azure
    status VARCHAR(50) NOT NULL DEFAULT 'uploaded' CHECK (status IN ('uploaded', 'processing', 'processed', 'failed', 'deleted')),
    processing_started_at DATETIME2 NULL,
    processing_completed_at DATETIME2 NULL,
    error_message TEXT NULL,
    metadata JSON NULL, -- Información extraída del .pbit (modelo, medidas, tablas, etc.)
    is_deleted BIT NOT NULL DEFAULT 0,
    deleted_at DATETIME2 NULL,
    deleted_by UNIQUEIDENTIFIER NULL,
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_reports_organization FOREIGN KEY (organization_id) REFERENCES organizations(id) ON DELETE SET NULL,
    CONSTRAINT fk_reports_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE,
    CONSTRAINT fk_reports_deleted_by FOREIGN KEY (deleted_by) REFERENCES users(id)
);
GO

-- Índices
CREATE INDEX idx_reports_organization_id ON reports(organization_id);
CREATE INDEX idx_reports_user_id ON reports(user_id);
CREATE INDEX idx_reports_status ON reports(status);
CREATE INDEX idx_reports_is_deleted ON reports(is_deleted);
CREATE INDEX idx_reports_created_at ON reports(created_at);
CREATE INDEX idx_reports_blob_name ON reports(blob_name) WHERE blob_name IS NOT NULL;
GO

-- Comentarios
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Reportes subidos por los usuarios. Pueden pertenecer a una organización o ser individuales (plan basic).', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'reports';
GO

-- ============================================================================
-- TABLA: organization_documentation
-- ============================================================================
-- URLs de documentación personalizada por organización
-- ============================================================================

IF OBJECT_ID('organization_documentation', 'U') IS NOT NULL DROP TABLE organization_documentation;
GO

CREATE TABLE organization_documentation (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL UNIQUE, -- Una organización solo puede tener una URL
    documentation_url VARCHAR(500) NOT NULL,
    description TEXT NULL, -- Descripción opcional de la documentación
    is_active BIT NOT NULL DEFAULT 1,
    created_by UNIQUEIDENTIFIER NULL, -- Usuario que configuró la URL
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_org_doc_organization FOREIGN KEY (organization_id) REFERENCES organizations(id) ON DELETE CASCADE,
    CONSTRAINT fk_org_doc_created_by FOREIGN KEY (created_by) REFERENCES users(id)
);
GO

CREATE INDEX idx_org_doc_organization ON organization_documentation(organization_id);
CREATE INDEX idx_org_doc_active ON organization_documentation(is_active);
GO

-- Comentarios
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'URLs de documentación personalizada por organización. Cada organización puede tener un link a su documentación.', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'organization_documentation';
GO

-- ============================================================================
-- DATOS INICIALES: plans
-- ============================================================================
-- Insertar los planes predefinidos con sus límites
-- ============================================================================

INSERT INTO plans (id, name, description, max_users, max_reports, max_storage_mb, max_organizations, features, price_monthly, price_yearly, is_active)
VALUES
    (
        'free_trial',
        'Free Trial',
        'Período de prueba gratuito con acceso completo a todas las funcionalidades',
        10, -- Máximo 10 usuarios durante el trial
        100, -- Máximo 100 reportes
        5000, -- 5GB de almacenamiento
        NULL, -- max_organizations: no aplica
        '{"api_access": false, "branding": false, "audit_log": false, "priority_support": false}',
        NULL, -- Gratis
        NULL, -- Gratis
        1
    ),
    (
        'basic',
        'Basic',
        'Plan individual para usuarios que trabajan solos',
        1, -- Solo 1 usuario
        30, -- Hasta 30 reportes
        1000, -- 1GB
        NULL, -- max_organizations: no aplica
        '{"api_access": false, "branding": false, "audit_log": false, "priority_support": false}',
        9.99, -- Ejemplo de precio mensual
        99.99, -- Ejemplo de precio anual
        1
    ),
    (
        'teams',
        'Teams',
        'Plan colaborativo para equipos pequeños',
        3, -- Hasta 3 usuarios
        50, -- Hasta 50 reportes compartidos
        5000, -- 5GB
        NULL, -- max_organizations: no aplica
        '{"api_access": false, "branding": false, "audit_log": false, "priority_support": false}',
        29.99,
        299.99,
        1
    ),
    (
        'enterprise',
        'Enterprise',
        'Plan completo para organizaciones grandes con todas las características',
        10, -- Hasta 10 usuarios
        300, -- Hasta 300 reportes
        50000, -- 50GB
        NULL, -- max_organizations: no aplica
        '{"api_access": true, "branding": true, "audit_log": true, "priority_support": true}',
        99.99,
        999.99,
        1
    ),
    (
        'enterprise_pro',
        'Enterprise Pro',
        'Plan para empresas de consultoría que gestionan múltiples clientes. Permite crear y gestionar hasta 5 organizaciones separadas con metadata confidencial.',
        50, -- Hasta 50 usuarios (cubrir múltiples equipos)
        1000, -- Hasta 1000 reportes (múltiples clientes)
        200000, -- 200GB de almacenamiento
        5, -- Máximo 5 organizaciones gestionadas
        '{"api_access": true, "branding": true, "audit_log": true, "priority_support": true, "multi_organization": true, "organization_isolation": true, "advanced_user_management": true, "global_admin_role": true}',
        199.99,
        1999.99,
        1
    );
GO

-- ============================================================================
-- TRIGGERS: Actualización automática de updated_at
-- ============================================================================
-- Triggers para actualizar automáticamente el campo updated_at
-- ============================================================================

-- Trigger para users
CREATE OR ALTER TRIGGER trg_users_updated_at
ON users
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE users
    SET updated_at = GETUTCDATE()
    WHERE id IN (SELECT id FROM inserted);
END;
GO

-- Trigger para organizations
CREATE OR ALTER TRIGGER trg_organizations_updated_at
ON organizations
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE organizations
    SET updated_at = GETUTCDATE()
    WHERE id IN (SELECT id FROM inserted);
END;
GO

-- Trigger para plans
CREATE OR ALTER TRIGGER trg_plans_updated_at
ON plans
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE plans
    SET updated_at = GETUTCDATE()
    WHERE id IN (SELECT id FROM inserted);
END;
GO

-- Trigger para subscriptions
CREATE OR ALTER TRIGGER trg_subscriptions_updated_at
ON subscriptions
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE subscriptions
    SET updated_at = GETUTCDATE()
    WHERE id IN (SELECT id FROM inserted);
END;
GO

-- Trigger para organization_members
CREATE OR ALTER TRIGGER trg_org_members_updated_at
ON organization_members
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE organization_members
    SET updated_at = GETUTCDATE()
    WHERE id IN (SELECT id FROM inserted);
END;
GO

-- Trigger para reports
CREATE OR ALTER TRIGGER trg_reports_updated_at
ON reports
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE reports
    SET updated_at = GETUTCDATE()
    WHERE id IN (SELECT id FROM inserted);
END;
GO

-- Trigger para organization_documentation
CREATE OR ALTER TRIGGER trg_org_documentation_updated_at
ON organization_documentation
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE organization_documentation
    SET updated_at = GETUTCDATE()
    WHERE id IN (SELECT id FROM inserted);
END;
GO

-- ============================================================================
-- VISTAS ÚTILES
-- ============================================================================
-- Vistas para consultas comunes
-- ============================================================================

-- Vista: Organizaciones con sus suscripciones activas
CREATE OR ALTER VIEW vw_organizations_with_subscription AS
SELECT 
    o.id,
    o.name,
    o.slug,
    o.stripe_customer_id,
    s.id AS subscription_id,
    s.plan_id,
    p.name AS plan_name,
    s.status AS subscription_status,
    s.current_period_end,
    s.trial_end,
    (SELECT COUNT(*) FROM organization_members om WHERE om.organization_id = o.id AND om.left_at IS NULL) AS current_users_count,
    (SELECT COUNT(*) FROM reports r WHERE r.organization_id = o.id AND r.is_deleted = 0) AS current_reports_count,
    p.max_users,
    p.max_reports,
    o.created_at
FROM organizations o
LEFT JOIN subscriptions s ON s.organization_id = o.id AND s.status IN ('active', 'trialing')
LEFT JOIN plans p ON p.id = s.plan_id
WHERE o.is_archived = 0;
GO

-- Vista: Usuarios con sus organizaciones principales
CREATE OR ALTER VIEW vw_users_with_primary_org AS
SELECT 
    u.id,
    u.email,
    u.name,
    u.auth_provider,
    u.last_login_at,
    om.organization_id AS primary_organization_id,
    o.name AS primary_organization_name,
    od.documentation_url AS organization_documentation_url,
    CASE WHEN od.documentation_url IS NOT NULL AND od.is_active = 1 THEN 1 ELSE 0 END AS has_documentation,
    u.created_at
FROM users u
LEFT JOIN organization_members om ON om.user_id = u.id AND om.is_primary = 1 AND om.left_at IS NULL
LEFT JOIN organizations o ON o.id = om.organization_id
LEFT JOIN organization_documentation od ON od.organization_id = o.id AND od.is_active = 1
WHERE u.is_active = 1;
GO

-- ============================================================================
-- FUNCIONES: Validación de límites
-- ============================================================================
-- Funciones para validar límites de planes
-- ============================================================================

-- Función: Verificar si una organización puede agregar más usuarios
CREATE OR ALTER FUNCTION fn_can_add_user(@organization_id UNIQUEIDENTIFIER)
RETURNS BIT
AS
BEGIN
    DECLARE @max_users INT;
    DECLARE @current_users INT;
    
    -- Obtener límite de usuarios del plan activo
    SELECT @max_users = p.max_users
    FROM subscriptions s
    INNER JOIN plans p ON p.id = s.plan_id
    WHERE s.organization_id = @organization_id
    AND s.status IN ('active', 'trialing');
    
    -- Si no hay suscripción activa, retornar 0
    IF @max_users IS NULL
        RETURN 0;
    
    -- Contar usuarios activos
    SELECT @current_users = COUNT(*)
    FROM organization_members
    WHERE organization_id = @organization_id
    AND left_at IS NULL;
    
    -- Verificar si puede agregar más
    IF @current_users < @max_users
        RETURN 1;
    
    RETURN 0;
END;
GO

-- Función: Verificar si una organización puede agregar más reportes
CREATE OR ALTER FUNCTION fn_can_add_report(@organization_id UNIQUEIDENTIFIER)
RETURNS BIT
AS
BEGIN
    DECLARE @max_reports INT;
    DECLARE @current_reports INT;
    
    -- Obtener límite de reportes del plan activo
    SELECT @max_reports = p.max_reports
    FROM subscriptions s
    INNER JOIN plans p ON p.id = s.plan_id
    WHERE s.organization_id = @organization_id
    AND s.status IN ('active', 'trialing');
    
    -- Si no hay suscripción activa, retornar 0
    IF @max_reports IS NULL
        RETURN 0;
    
    -- Contar reportes activos
    SELECT @current_reports = COUNT(*)
    FROM reports
    WHERE organization_id = @organization_id
    AND is_deleted = 0;
    
    -- Verificar si puede agregar más
    IF @current_reports < @max_reports
        RETURN 1;
    
    RETURN 0;
END;
GO

-- Función: Obtener URL de documentación de una organización
CREATE OR ALTER FUNCTION fn_get_organization_documentation_url(@organization_id UNIQUEIDENTIFIER)
RETURNS VARCHAR(500)
AS
BEGIN
    DECLARE @url VARCHAR(500);
    
    SELECT @url = documentation_url
    FROM organization_documentation
    WHERE organization_id = @organization_id
    AND is_active = 1;
    
    RETURN @url;
END;
GO

-- ============================================================================
-- ÍNDICES ADICIONALES PARA PERFORMANCE
-- ============================================================================

-- Índice compuesto para búsqueda rápida de reportes por organización y estado
CREATE INDEX idx_reports_org_status_deleted ON reports(organization_id, status, is_deleted) 
WHERE organization_id IS NOT NULL;
GO

-- Índice compuesto para organización members activos
CREATE INDEX idx_org_members_org_active ON organization_members(organization_id, left_at) 
WHERE left_at IS NULL;
GO

-- ============================================================================
-- COMENTARIOS FINALES
-- ============================================================================

PRINT 'Schema de base de datos creado exitosamente';
PRINT 'Tablas: users, organizations, organization_documentation, plans, subscriptions, subscription_history, reports, organization_members';
PRINT 'Planes: free_trial, basic, teams, enterprise, enterprise_pro';
PRINT 'Vistas: vw_organizations_with_subscription, vw_users_with_primary_org';
PRINT 'Funciones: fn_can_add_user, fn_can_add_report, fn_get_organization_documentation_url';
PRINT 'Triggers: updated_at automatico (7 tablas)';
PRINT '';
PRINT 'Proximos pasos:';
PRINT '   - Ejecutar organization_workflows.sql';
PRINT '   - Ejecutar state_machine_and_workflows.sql';
PRINT '   - Ejecutar enterprise_pro_plan_v2.sql (solo si necesitas Enterprise Pro)';
PRINT '';
PRINT 'A/B Testing: Se maneja con HubSpot';
PRINT 'Pricing complejo: Se maneja con Stripe + HubSpot';
GO

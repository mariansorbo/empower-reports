-- ============================================================================
-- REPORT TUNER - Settings Tables and Procedures
-- ============================================================================
-- Tablas, stored procedures, triggers y funciones para el panel de Settings
-- Basado en SETTINGS_DATABASE_MAPPING_SIMPLIFIED.md
-- ============================================================================
--
-- Este archivo agrega:
-- 1. Tabla stripe_invoices - Historial de facturas de Stripe
-- 2. Tabla organization_integrations - Integraciones (Power BI, API, etc.)
-- 3. Campo stripe_payment_method_id en organizations
-- 4. Stored procedures para CRUD de ambas tablas
-- 5. Triggers de updated_at
-- 6. Funciones útiles
-- ============================================================================

USE empower_reports;
GO

-- ============================================================================
-- MODIFICACIÓN A TABLA EXISTENTE: organizations
-- ============================================================================
-- Agregar campo para método de pago por defecto de Stripe
-- ============================================================================

IF NOT EXISTS (
    SELECT 1 FROM sys.columns 
    WHERE object_id = OBJECT_ID('organizations') 
    AND name = 'stripe_payment_method_id'
)
BEGIN
    ALTER TABLE organizations 
    ADD stripe_payment_method_id VARCHAR(255) NULL;
    
    CREATE INDEX idx_organizations_payment_method 
    ON organizations(stripe_payment_method_id) 
    WHERE stripe_payment_method_id IS NOT NULL;
    
    PRINT 'Campo stripe_payment_method_id agregado a organizations';
END
ELSE
BEGIN
    PRINT 'Campo stripe_payment_method_id ya existe en organizations';
END
GO

-- ============================================================================
-- TABLA: stripe_invoices
-- ============================================================================
-- Historial de facturas de Stripe para el panel de Settings > Billing
-- ============================================================================

IF OBJECT_ID('stripe_invoices', 'U') IS NOT NULL DROP TABLE stripe_invoices;
GO

CREATE TABLE stripe_invoices (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL,
    subscription_id UNIQUEIDENTIFIER NOT NULL,
    stripe_invoice_id VARCHAR(255) NOT NULL UNIQUE,
    stripe_invoice_pdf VARCHAR(500) NULL, -- URL del PDF de la factura
    stripe_hosted_invoice_url VARCHAR(500) NULL, -- URL de la factura en Stripe
    amount_due DECIMAL(10, 2) NOT NULL DEFAULT 0,
    amount_paid DECIMAL(10, 2) NOT NULL DEFAULT 0,
    currency VARCHAR(3) NOT NULL DEFAULT 'USD',
    status VARCHAR(50) NOT NULL CHECK (status IN ('draft', 'open', 'paid', 'void', 'uncollectible')),
    period_start DATETIME2 NOT NULL,
    period_end DATETIME2 NOT NULL,
    paid_at DATETIME2 NULL,
    metadata JSON NULL, -- Información adicional de Stripe
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_invoices_org FOREIGN KEY (organization_id) REFERENCES organizations(id) ON DELETE CASCADE,
    CONSTRAINT fk_invoices_subscription FOREIGN KEY (subscription_id) REFERENCES subscriptions(id) ON DELETE CASCADE
);
GO

-- Índices para performance
CREATE INDEX idx_invoices_organization_id ON stripe_invoices(organization_id);
CREATE INDEX idx_invoices_subscription_id ON stripe_invoices(subscription_id);
CREATE INDEX idx_invoices_stripe_invoice_id ON stripe_invoices(stripe_invoice_id);
CREATE INDEX idx_invoices_status ON stripe_invoices(status);
CREATE INDEX idx_invoices_period_end ON stripe_invoices(period_end);
CREATE INDEX idx_invoices_created_at ON stripe_invoices(created_at);
GO

-- Comentarios
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Historial de facturas de Stripe. Sincronizado desde webhooks de Stripe para mostrar en Settings > Billing.', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'stripe_invoices';
GO

-- ============================================================================
-- TABLA: organization_integrations
-- ============================================================================
-- Integraciones de organizaciones (Power BI, API, Documentation endpoint)
-- Para el panel de Settings > Integrations
-- ============================================================================

IF OBJECT_ID('organization_integrations', 'U') IS NOT NULL DROP TABLE organization_integrations;
GO

CREATE TABLE organization_integrations (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    organization_id UNIQUEIDENTIFIER NOT NULL UNIQUE,
    
    -- ===== POWER BI SERVICE =====
    powerbi_enabled BIT NOT NULL DEFAULT 0,
    powerbi_workspace_id VARCHAR(255) NULL,
    powerbi_tenant_id VARCHAR(255) NULL,
    powerbi_access_token_encrypted VARBINARY(MAX) NULL, -- Token OAuth cifrado
    powerbi_refresh_token_encrypted VARBINARY(MAX) NULL, -- Refresh token cifrado
    powerbi_token_expires_at DATETIME2 NULL,
    powerbi_last_sync DATETIME2 NULL,
    
    -- ===== API ACCESS (Enterprise) =====
    api_enabled BIT NOT NULL DEFAULT 0,
    api_key VARCHAR(64) NULL UNIQUE, -- pk_live_xxx o pk_test_xxx
    api_secret_hash VARCHAR(255) NULL, -- Secret key hasheado (bcrypt/argon2)
    api_rate_limit_per_hour INT NOT NULL DEFAULT 1000,
    api_rate_limit_per_day INT NOT NULL DEFAULT 10000,
    api_last_used DATETIME2 NULL,
    api_total_calls INT NOT NULL DEFAULT 0,
    
    -- ===== DOCUMENTATION API =====
    documentation_api_enabled BIT NOT NULL DEFAULT 0,
    documentation_requires_auth BIT NOT NULL DEFAULT 1,
    documentation_public_key VARCHAR(64) NULL, -- Para autenticación pública del endpoint
    
    -- ===== METADATA =====
    metadata JSON NULL, -- Información adicional de las integraciones
    
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_integrations_org FOREIGN KEY (organization_id) REFERENCES organizations(id) ON DELETE CASCADE
);
GO

-- Índices para performance
CREATE INDEX idx_integrations_org ON organization_integrations(organization_id);
CREATE INDEX idx_integrations_api_key ON organization_integrations(api_key) WHERE api_key IS NOT NULL;
CREATE INDEX idx_integrations_powerbi_enabled ON organization_integrations(powerbi_enabled) WHERE powerbi_enabled = 1;
CREATE INDEX idx_integrations_api_enabled ON organization_integrations(api_enabled) WHERE api_enabled = 1;
GO

-- Comentarios
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Integraciones de organizaciones: Power BI Service, API access, Documentation endpoint. Configurado desde Settings > Integrations.', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'organization_integrations';
GO

-- ============================================================================
-- TRIGGERS: Actualización automática de updated_at
-- ============================================================================

-- Trigger para stripe_invoices
CREATE OR ALTER TRIGGER trg_stripe_invoices_updated_at
ON stripe_invoices
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE stripe_invoices
    SET updated_at = GETUTCDATE()
    WHERE id IN (SELECT id FROM inserted);
END;
GO

-- Trigger para organization_integrations
CREATE OR ALTER TRIGGER trg_org_integrations_updated_at
ON organization_integrations
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE organization_integrations
    SET updated_at = GETUTCDATE()
    WHERE id IN (SELECT id FROM inserted);
END;
GO

-- ============================================================================
-- STORED PROCEDURES: stripe_invoices
-- ============================================================================

-- Procedure: Crear o actualizar factura de Stripe
CREATE OR ALTER PROCEDURE sp_upsert_stripe_invoice
    @organization_id UNIQUEIDENTIFIER,
    @subscription_id UNIQUEIDENTIFIER,
    @stripe_invoice_id VARCHAR(255),
    @stripe_invoice_pdf VARCHAR(500) = NULL,
    @stripe_hosted_invoice_url VARCHAR(500) = NULL,
    @amount_due DECIMAL(10, 2),
    @amount_paid DECIMAL(10, 2),
    @currency VARCHAR(3) = 'USD',
    @status VARCHAR(50),
    @period_start DATETIME2,
    @period_end DATETIME2,
    @paid_at DATETIME2 = NULL,
    @metadata JSON = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Validar que la organización existe
        IF NOT EXISTS (SELECT 1 FROM organizations WHERE id = @organization_id AND is_archived = 0)
        BEGIN
            THROW 50030, 'Organización no existe o está archivada', 1;
        END
        
        -- Validar que la suscripción existe
        IF NOT EXISTS (SELECT 1 FROM subscriptions WHERE id = @subscription_id)
        BEGIN
            THROW 50031, 'Suscripción no existe', 1;
        END
        
        -- Verificar si ya existe la factura
        IF EXISTS (SELECT 1 FROM stripe_invoices WHERE stripe_invoice_id = @stripe_invoice_id)
        BEGIN
            -- Actualizar existente
            UPDATE stripe_invoices
            SET 
                organization_id = @organization_id,
                subscription_id = @subscription_id,
                stripe_invoice_pdf = ISNULL(@stripe_invoice_pdf, stripe_invoice_pdf),
                stripe_hosted_invoice_url = ISNULL(@stripe_hosted_invoice_url, stripe_hosted_invoice_url),
                amount_due = @amount_due,
                amount_paid = @amount_paid,
                currency = @currency,
                status = @status,
                period_start = @period_start,
                period_end = @period_end,
                paid_at = ISNULL(@paid_at, paid_at),
                metadata = ISNULL(@metadata, metadata),
                updated_at = GETUTCDATE()
            WHERE stripe_invoice_id = @stripe_invoice_id;
            
            SELECT 1 AS success, 'Factura actualizada exitosamente' AS message;
        END
        ELSE
        BEGIN
            -- Insertar nueva
            INSERT INTO stripe_invoices (
                organization_id,
                subscription_id,
                stripe_invoice_id,
                stripe_invoice_pdf,
                stripe_hosted_invoice_url,
                amount_due,
                amount_paid,
                currency,
                status,
                period_start,
                period_end,
                paid_at,
                metadata
            )
            VALUES (
                @organization_id,
                @subscription_id,
                @stripe_invoice_id,
                @stripe_invoice_pdf,
                @stripe_hosted_invoice_url,
                @amount_due,
                @amount_paid,
                @currency,
                @status,
                @period_start,
                @period_end,
                @paid_at,
                @metadata
            );
            
            SELECT 1 AS success, 'Factura creada exitosamente' AS message;
        END
        
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedure: Obtener facturas de una organización
CREATE OR ALTER PROCEDURE sp_get_organization_invoices
    @organization_id UNIQUEIDENTIFIER,
    @status VARCHAR(50) = NULL,
    @limit INT = 50,
    @offset INT = 0
AS
BEGIN
    SET NOCOUNT ON;
    
    SELECT 
        si.id,
        si.stripe_invoice_id,
        si.stripe_invoice_pdf,
        si.stripe_hosted_invoice_url,
        si.amount_due,
        si.amount_paid,
        si.currency,
        si.status,
        si.period_start,
        si.period_end,
        si.paid_at,
        si.metadata,
        si.created_at,
        s.plan_id,
        p.name AS plan_name
    FROM stripe_invoices si
    INNER JOIN subscriptions s ON s.id = si.subscription_id
    LEFT JOIN plans p ON p.id = s.plan_id
    WHERE si.organization_id = @organization_id
    AND (@status IS NULL OR si.status = @status)
    ORDER BY si.created_at DESC
    OFFSET @offset ROWS
    FETCH NEXT @limit ROWS ONLY;
END;
GO

-- Procedure: Obtener una factura específica
CREATE OR ALTER PROCEDURE sp_get_stripe_invoice
    @stripe_invoice_id VARCHAR(255)
AS
BEGIN
    SET NOCOUNT ON;
    
    SELECT 
        si.*,
        o.name AS organization_name,
        s.plan_id,
        p.name AS plan_name
    FROM stripe_invoices si
    INNER JOIN organizations o ON o.id = si.organization_id
    INNER JOIN subscriptions s ON s.id = si.subscription_id
    LEFT JOIN plans p ON p.id = s.plan_id
    WHERE si.stripe_invoice_id = @stripe_invoice_id;
END;
GO

-- Procedure: Actualizar método de pago por defecto
CREATE OR ALTER PROCEDURE sp_update_organization_payment_method
    @organization_id UNIQUEIDENTIFIER,
    @stripe_payment_method_id VARCHAR(255)
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Validar que la organización existe
        IF NOT EXISTS (SELECT 1 FROM organizations WHERE id = @organization_id AND is_archived = 0)
        BEGIN
            THROW 50032, 'Organización no existe o está archivada', 1;
        END
        
        UPDATE organizations
        SET 
            stripe_payment_method_id = @stripe_payment_method_id,
            updated_at = GETUTCDATE()
        WHERE id = @organization_id;
        
        COMMIT TRANSACTION;
        
        SELECT 1 AS success, 'Método de pago actualizado exitosamente' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- ============================================================================
-- STORED PROCEDURES: organization_integrations
-- ============================================================================

-- Procedure: Crear o actualizar integraciones de organización
CREATE OR ALTER PROCEDURE sp_upsert_organization_integrations
    @organization_id UNIQUEIDENTIFIER,
    @powerbi_enabled BIT = NULL,
    @powerbi_workspace_id VARCHAR(255) = NULL,
    @powerbi_tenant_id VARCHAR(255) = NULL,
    @powerbi_access_token_encrypted VARBINARY(MAX) = NULL,
    @powerbi_refresh_token_encrypted VARBINARY(MAX) = NULL,
    @powerbi_token_expires_at DATETIME2 = NULL,
    @powerbi_last_sync DATETIME2 = NULL,
    @api_enabled BIT = NULL,
    @api_key VARCHAR(64) = NULL,
    @api_secret_hash VARCHAR(255) = NULL,
    @api_rate_limit_per_hour INT = NULL,
    @api_rate_limit_per_day INT = NULL,
    @api_last_used DATETIME2 = NULL,
    @documentation_api_enabled BIT = NULL,
    @documentation_requires_auth BIT = NULL,
    @documentation_public_key VARCHAR(64) = NULL,
    @metadata JSON = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Validar que la organización existe
        IF NOT EXISTS (SELECT 1 FROM organizations WHERE id = @organization_id AND is_archived = 0)
        BEGIN
            THROW 50033, 'Organización no existe o está archivada', 1;
        END
        
        -- Verificar si ya existe registro
        IF EXISTS (SELECT 1 FROM organization_integrations WHERE organization_id = @organization_id)
        BEGIN
            -- Actualizar existente
            UPDATE organization_integrations
            SET 
                powerbi_enabled = ISNULL(@powerbi_enabled, powerbi_enabled),
                powerbi_workspace_id = ISNULL(@powerbi_workspace_id, powerbi_workspace_id),
                powerbi_tenant_id = ISNULL(@powerbi_tenant_id, powerbi_tenant_id),
                powerbi_access_token_encrypted = ISNULL(@powerbi_access_token_encrypted, powerbi_access_token_encrypted),
                powerbi_refresh_token_encrypted = ISNULL(@powerbi_refresh_token_encrypted, powerbi_refresh_token_encrypted),
                powerbi_token_expires_at = ISNULL(@powerbi_token_expires_at, powerbi_token_expires_at),
                powerbi_last_sync = ISNULL(@powerbi_last_sync, powerbi_last_sync),
                api_enabled = ISNULL(@api_enabled, api_enabled),
                api_key = ISNULL(@api_key, api_key),
                api_secret_hash = ISNULL(@api_secret_hash, api_secret_hash),
                api_rate_limit_per_hour = ISNULL(@api_rate_limit_per_hour, api_rate_limit_per_hour),
                api_rate_limit_per_day = ISNULL(@api_rate_limit_per_day, api_rate_limit_per_day),
                api_last_used = ISNULL(@api_last_used, api_last_used),
                documentation_api_enabled = ISNULL(@documentation_api_enabled, documentation_api_enabled),
                documentation_requires_auth = ISNULL(@documentation_requires_auth, documentation_requires_auth),
                documentation_public_key = ISNULL(@documentation_public_key, documentation_public_key),
                metadata = ISNULL(@metadata, metadata),
                updated_at = GETUTCDATE()
            WHERE organization_id = @organization_id;
            
            SELECT 1 AS success, 'Integraciones actualizadas exitosamente' AS message;
        END
        ELSE
        BEGIN
            -- Insertar nueva
            INSERT INTO organization_integrations (
                organization_id,
                powerbi_enabled,
                powerbi_workspace_id,
                powerbi_tenant_id,
                powerbi_access_token_encrypted,
                powerbi_refresh_token_encrypted,
                powerbi_token_expires_at,
                powerbi_last_sync,
                api_enabled,
                api_key,
                api_secret_hash,
                api_rate_limit_per_hour,
                api_rate_limit_per_day,
                api_last_used,
                documentation_api_enabled,
                documentation_requires_auth,
                documentation_public_key,
                metadata
            )
            VALUES (
                @organization_id,
                ISNULL(@powerbi_enabled, 0),
                @powerbi_workspace_id,
                @powerbi_tenant_id,
                @powerbi_access_token_encrypted,
                @powerbi_refresh_token_encrypted,
                @powerbi_token_expires_at,
                @powerbi_last_sync,
                ISNULL(@api_enabled, 0),
                @api_key,
                @api_secret_hash,
                ISNULL(@api_rate_limit_per_hour, 1000),
                ISNULL(@api_rate_limit_per_day, 10000),
                @api_last_used,
                ISNULL(@documentation_api_enabled, 0),
                ISNULL(@documentation_requires_auth, 1),
                @documentation_public_key,
                @metadata
            );
            
            SELECT 1 AS success, 'Integraciones creadas exitosamente' AS message;
        END
        
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedure: Obtener integraciones de una organización
CREATE OR ALTER PROCEDURE sp_get_organization_integrations
    @organization_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    SELECT 
        oi.*,
        o.name AS organization_name
    FROM organization_integrations oi
    INNER JOIN organizations o ON o.id = oi.organization_id
    WHERE oi.organization_id = @organization_id;
END;
GO

-- Procedure: Actualizar última sincronización de Power BI
CREATE OR ALTER PROCEDURE sp_update_powerbi_last_sync
    @organization_id UNIQUEIDENTIFIER,
    @last_sync DATETIME2 = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE organization_integrations
    SET 
        powerbi_last_sync = ISNULL(@last_sync, GETUTCDATE()),
        updated_at = GETUTCDATE()
    WHERE organization_id = @organization_id
    AND powerbi_enabled = 1;
    
    IF @@ROWCOUNT > 0
        SELECT 1 AS success, 'Última sincronización actualizada' AS message;
    ELSE
        SELECT 0 AS success, 'No se encontró integración de Power BI para esta organización' AS message;
END;
GO

-- Procedure: Incrementar contador de llamadas API
CREATE OR ALTER PROCEDURE sp_increment_api_calls
    @organization_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE organization_integrations
    SET 
        api_total_calls = api_total_calls + 1,
        api_last_used = GETUTCDATE(),
        updated_at = GETUTCDATE()
    WHERE organization_id = @organization_id
    AND api_enabled = 1;
    
    IF @@ROWCOUNT > 0
        SELECT 1 AS success, 'Contador de API actualizado' AS message;
    ELSE
        SELECT 0 AS success, 'API no está habilitada para esta organización' AS message;
END;
GO

-- Procedure: Verificar si API key es válida
CREATE OR ALTER PROCEDURE sp_validate_api_key
    @api_key VARCHAR(64)
AS
BEGIN
    SET NOCOUNT ON;
    
    SELECT 
        oi.organization_id,
        o.name AS organization_name,
        oi.api_enabled,
        oi.api_rate_limit_per_hour,
        oi.api_rate_limit_per_day,
        oi.api_total_calls,
        oi.api_last_used,
        CASE 
            WHEN oi.api_enabled = 0 THEN 0
            WHEN o.is_archived = 1 THEN 0
            ELSE 1
        END AS is_valid
    FROM organization_integrations oi
    INNER JOIN organizations o ON o.id = oi.organization_id
    WHERE oi.api_key = @api_key;
END;
GO

-- Procedure: Deshabilitar Power BI
CREATE OR ALTER PROCEDURE sp_disable_powerbi_integration
    @organization_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE organization_integrations
    SET 
        powerbi_enabled = 0,
        powerbi_workspace_id = NULL,
        powerbi_tenant_id = NULL,
        powerbi_access_token_encrypted = NULL,
        powerbi_refresh_token_encrypted = NULL,
        powerbi_token_expires_at = NULL,
        powerbi_last_sync = NULL,
        updated_at = GETUTCDATE()
    WHERE organization_id = @organization_id;
    
    IF @@ROWCOUNT > 0
        SELECT 1 AS success, 'Integración de Power BI deshabilitada' AS message;
    ELSE
        SELECT 0 AS success, 'No se encontró integración de Power BI para esta organización' AS message;
END;
GO

-- Procedure: Deshabilitar API
CREATE OR ALTER PROCEDURE sp_disable_api_integration
    @organization_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE organization_integrations
    SET 
        api_enabled = 0,
        api_key = NULL,
        api_secret_hash = NULL,
        api_last_used = NULL,
        updated_at = GETUTCDATE()
    WHERE organization_id = @organization_id;
    
    IF @@ROWCOUNT > 0
        SELECT 1 AS success, 'API deshabilitada' AS message;
    ELSE
        SELECT 0 AS success, 'No se encontró configuración de API para esta organización' AS message;
END;
GO

-- ============================================================================
-- FUNCIONES ÚTILES
-- ============================================================================

-- Función: Verificar si una organización tiene Power BI conectado
CREATE OR ALTER FUNCTION fn_has_powerbi_integration(@organization_id UNIQUEIDENTIFIER)
RETURNS BIT
AS
BEGIN
    DECLARE @has_integration BIT = 0;
    
    SELECT @has_integration = CASE WHEN powerbi_enabled = 1 AND powerbi_access_token_encrypted IS NOT NULL THEN 1 ELSE 0 END
    FROM organization_integrations
    WHERE organization_id = @organization_id;
    
    RETURN ISNULL(@has_integration, 0);
END;
GO

-- Función: Verificar si una organización tiene API habilitada
CREATE OR ALTER FUNCTION fn_has_api_access(@organization_id UNIQUEIDENTIFIER)
RETURNS BIT
AS
BEGIN
    DECLARE @has_api BIT = 0;
    
    SELECT @has_api = api_enabled
    FROM organization_integrations
    WHERE organization_id = @organization_id;
    
    RETURN ISNULL(@has_api, 0);
END;
GO

-- Función: Obtener total de facturas pagadas de una organización
CREATE OR ALTER FUNCTION fn_get_total_paid_invoices(@organization_id UNIQUEIDENTIFIER)
RETURNS DECIMAL(10, 2)
AS
BEGIN
    DECLARE @total DECIMAL(10, 2) = 0;
    
    SELECT @total = SUM(amount_paid)
    FROM stripe_invoices
    WHERE organization_id = @organization_id
    AND status = 'paid';
    
    RETURN ISNULL(@total, 0);
END;
GO

-- Función: Obtener número de facturas de una organización
CREATE OR ALTER FUNCTION fn_get_invoice_count(@organization_id UNIQUEIDENTIFIER, @status VARCHAR(50) = NULL)
RETURNS INT
AS
BEGIN
    DECLARE @count INT = 0;
    
    SELECT @count = COUNT(*)
    FROM stripe_invoices
    WHERE organization_id = @organization_id
    AND (@status IS NULL OR status = @status);
    
    RETURN @count;
END;
GO

-- ============================================================================
-- VISTAS ÚTILES
-- ============================================================================

-- Vista: Organizaciones con sus integraciones
CREATE OR ALTER VIEW vw_organizations_with_integrations AS
SELECT 
    o.id AS organization_id,
    o.name AS organization_name,
    o.slug,
    oi.powerbi_enabled,
    oi.powerbi_workspace_id,
    oi.powerbi_last_sync,
    oi.api_enabled,
    oi.api_key,
    oi.api_total_calls,
    oi.api_last_used,
    oi.documentation_api_enabled,
    oi.documentation_requires_auth,
    o.stripe_payment_method_id,
    (SELECT COUNT(*) FROM stripe_invoices si WHERE si.organization_id = o.id AND si.status = 'paid') AS total_paid_invoices,
    (SELECT SUM(amount_paid) FROM stripe_invoices si WHERE si.organization_id = o.id AND si.status = 'paid') AS total_paid_amount,
    o.created_at
FROM organizations o
LEFT JOIN organization_integrations oi ON oi.organization_id = o.id
WHERE o.is_archived = 0;
GO

-- Vista: Resumen de facturas por organización
CREATE OR ALTER VIEW vw_invoices_summary AS
SELECT 
    si.organization_id,
    o.name AS organization_name,
    COUNT(*) AS total_invoices,
    SUM(CASE WHEN si.status = 'paid' THEN 1 ELSE 0 END) AS paid_invoices,
    SUM(CASE WHEN si.status = 'open' THEN 1 ELSE 0 END) AS open_invoices,
    SUM(CASE WHEN si.status = 'paid' THEN si.amount_paid ELSE 0 END) AS total_paid_amount,
    SUM(CASE WHEN si.status = 'open' THEN si.amount_due ELSE 0 END) AS total_open_amount,
    MAX(si.created_at) AS last_invoice_date
FROM stripe_invoices si
INNER JOIN organizations o ON o.id = si.organization_id
GROUP BY si.organization_id, o.name;
GO

-- ============================================================================
-- COMENTARIOS FINALES
-- ============================================================================

PRINT '✅ Tablas y procedures de Settings creados exitosamente';
PRINT '';
PRINT 'Tablas creadas:';
PRINT '  - stripe_invoices (historial de facturas)';
PRINT '  - organization_integrations (Power BI, API, Documentation)';
PRINT '';
PRINT 'Modificaciones:';
PRINT '  - organizations.stripe_payment_method_id (campo agregado)';
PRINT '';
PRINT 'Stored Procedures:';
PRINT '  - sp_upsert_stripe_invoice';
PRINT '  - sp_get_organization_invoices';
PRINT '  - sp_get_stripe_invoice';
PRINT '  - sp_update_organization_payment_method';
PRINT '  - sp_upsert_organization_integrations';
PRINT '  - sp_get_organization_integrations';
PRINT '  - sp_update_powerbi_last_sync';
PRINT '  - sp_increment_api_calls';
PRINT '  - sp_validate_api_key';
PRINT '  - sp_disable_powerbi_integration';
PRINT '  - sp_disable_api_integration';
PRINT '';
PRINT 'Funciones:';
PRINT '  - fn_has_powerbi_integration';
PRINT '  - fn_has_api_access';
PRINT '  - fn_get_total_paid_invoices';
PRINT '  - fn_get_invoice_count';
PRINT '';
PRINT 'Vistas:';
PRINT '  - vw_organizations_with_integrations';
PRINT '  - vw_invoices_summary';
PRINT '';
PRINT 'Triggers:';
PRINT '  - trg_stripe_invoices_updated_at';
PRINT '  - trg_org_integrations_updated_at';
PRINT '';
PRINT '✅ Settings tables and procedures instalado correctamente';
GO


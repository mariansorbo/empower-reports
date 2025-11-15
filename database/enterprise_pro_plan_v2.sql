-- ============================================================================
-- REPORT TUNER - Enterprise Pro Plan (V2)
-- ============================================================================
-- Plan para empresas de consultoría que necesitan gestionar múltiples 
-- organizaciones (clientes) con separación de metadata y confidencialidad
-- 
-- MODELO: Organizaciones independientes gestionadas por un admin global
-- NO hay jerarquía padre/hijo, solo relación de gestión
-- ============================================================================

USE empower_reports;
GO

-- ============================================================================
-- 1. AGREGAR CAMPO max_organizations A LA TABLA plans
-- ============================================================================

IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('plans') AND name = 'max_organizations')
BEGIN
    ALTER TABLE plans
    ADD max_organizations INT NULL;
    
    EXEC sp_addextendedproperty 
        @name = N'MS_Description', 
        @value = N'Máximo de organizaciones que puede gestionar un plan (solo para Enterprise Pro). NULL = no aplica.', 
        @level0type = N'SCHEMA', @level0name = N'dbo', 
        @level1type = N'TABLE', @level1name = N'plans',
        @level2type = N'COLUMN', @level2name = N'max_organizations';
END
GO

-- ============================================================================
-- 2. AGREGAR ROL 'admin_global' A organization_members
-- ============================================================================

-- Eliminar constraint existente
ALTER TABLE organization_members
DROP CONSTRAINT IF EXISTS CK_organization_members_role;
GO

-- Agregar nuevo constraint con admin_global
ALTER TABLE organization_members
ADD CONSTRAINT CK_organization_members_role 
CHECK (role IN ('admin', 'admin_global', 'member', 'viewer'));
GO

-- Comentario
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Roles: admin (admin normal), admin_global (puede gestionar múltiples orgs), member, viewer', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'organization_members',
    @level2type = N'COLUMN', @level2name = N'role';
GO

-- ============================================================================
-- 3. CREAR TABLA: enterprise_pro_managed_organizations
-- ============================================================================
-- Relaciona organizaciones Enterprise Pro con las organizaciones que gestionan
-- ============================================================================

IF OBJECT_ID('enterprise_pro_managed_organizations', 'U') IS NOT NULL DROP TABLE enterprise_pro_managed_organizations;
GO

CREATE TABLE enterprise_pro_managed_organizations (
    id UNIQUEIDENTIFIER PRIMARY KEY DEFAULT NEWID(),
    enterprise_pro_org_id UNIQUEIDENTIFIER NOT NULL, -- Organización con plan Enterprise Pro
    managed_organization_id UNIQUEIDENTIFIER NOT NULL, -- Organización gestionada
    admin_user_id UNIQUEIDENTIFIER NOT NULL, -- Usuario admin_global que gestiona
    created_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    updated_at DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    
    CONSTRAINT fk_ep_managed_ep_org FOREIGN KEY (enterprise_pro_org_id) REFERENCES organizations(id) ON DELETE CASCADE,
    CONSTRAINT fk_ep_managed_org FOREIGN KEY (managed_organization_id) REFERENCES organizations(id) ON DELETE CASCADE,
    CONSTRAINT fk_ep_managed_admin FOREIGN KEY (admin_user_id) REFERENCES users(id) ON DELETE CASCADE,
    
    -- Una organización solo puede ser gestionada por una Enterprise Pro
    CONSTRAINT uk_ep_managed_org UNIQUE (managed_organization_id),
    
    -- Un admin_global puede gestionar múltiples orgs, pero una org específica solo puede ser gestionada por un Enterprise Pro
    CONSTRAINT uk_ep_managed_ep_org_managed UNIQUE (enterprise_pro_org_id, managed_organization_id)
);
GO

-- Índices
CREATE INDEX idx_ep_managed_ep_org ON enterprise_pro_managed_organizations(enterprise_pro_org_id);
CREATE INDEX idx_ep_managed_org ON enterprise_pro_managed_organizations(managed_organization_id);
CREATE INDEX idx_ep_managed_admin ON enterprise_pro_managed_organizations(admin_user_id);
GO

-- Comentario
EXEC sp_addextendedproperty 
    @name = N'MS_Description', 
    @value = N'Relaciona organizaciones Enterprise Pro con las organizaciones que gestionan. Permite que un admin_global gestione múltiples organizaciones independientes.', 
    @level0type = N'SCHEMA', @level0name = N'dbo', 
    @level1type = N'TABLE', @level1name = N'enterprise_pro_managed_organizations';
GO

-- ============================================================================
-- 4. INSERTAR PLAN Enterprise Pro (si no existe)
-- ============================================================================

IF NOT EXISTS (SELECT 1 FROM plans WHERE id = 'enterprise_pro')
BEGIN
    INSERT INTO plans (
        id, 
        name, 
        description, 
        max_users, 
        max_reports, 
        max_storage_mb, 
        max_organizations,
        features, 
        price_monthly, 
        price_yearly, 
        is_active
    )
    VALUES (
        'enterprise_pro',
        'Enterprise Pro',
        'Plan para empresas de consultoría que gestionan múltiples clientes. Permite crear y gestionar hasta 5 organizaciones separadas con metadata confidencial.',
        50,  -- Más usuarios que Enterprise (10) para cubrir múltiples equipos
        1000,  -- Más reportes que Enterprise (300) para múltiples clientes
        200000,  -- 200GB de almacenamiento (vs 50GB de Enterprise)
        5,  -- Hasta 5 organizaciones gestionadas
        '{
            "api_access": true,
            "branding": true,
            "audit_log": true,
            "priority_support": true,
            "multi_organization": true,
            "organization_isolation": true,
            "advanced_user_management": true,
            "global_admin_role": true
        }',
        199.99,  -- Precio mensual
        1999.99,  -- Precio anual
        1
    );
END
GO

-- ============================================================================
-- 5. FUNCIÓN: Verificar si una organización Enterprise Pro puede gestionar más organizaciones
-- ============================================================================

CREATE OR ALTER FUNCTION fn_can_manage_more_organizations(@enterprise_pro_org_id UNIQUEIDENTIFIER)
RETURNS BIT
AS
BEGIN
    DECLARE @max_organizations INT;
    DECLARE @current_organizations INT;
    DECLARE @plan_id VARCHAR(50);
    
    -- Obtener plan activo de la organización Enterprise Pro
    SELECT @plan_id = s.plan_id
    FROM subscriptions s
    WHERE s.organization_id = @enterprise_pro_org_id
    AND s.status IN ('active', 'trialing');
    
    -- Si no hay suscripción activa, retornar 0
    IF @plan_id IS NULL
        RETURN 0;
    
    -- Verificar que es Enterprise Pro
    IF @plan_id != 'enterprise_pro'
        RETURN 0;
    
    -- Obtener límite de organizaciones del plan
    SELECT @max_organizations = max_organizations
    FROM plans
    WHERE id = @plan_id;
    
    -- Si el plan no permite gestionar organizaciones (NULL o 0), retornar 0
    IF @max_organizations IS NULL OR @max_organizations = 0
        RETURN 0;
    
    -- Contar organizaciones gestionadas activas (no archivadas)
    SELECT @current_organizations = COUNT(*)
    FROM enterprise_pro_managed_organizations epm
    INNER JOIN organizations o ON o.id = epm.managed_organization_id
    WHERE epm.enterprise_pro_org_id = @enterprise_pro_org_id
    AND o.is_archived = 0;
    
    -- Verificar si puede gestionar más
    IF @current_organizations < @max_organizations
        RETURN 1;
    
    RETURN 0;
END;
GO

-- ============================================================================
-- 6. FUNCIÓN: Obtener el número de organizaciones gestionadas
-- ============================================================================

CREATE OR ALTER FUNCTION fn_get_managed_organizations_count(@enterprise_pro_org_id UNIQUEIDENTIFIER)
RETURNS INT
AS
BEGIN
    DECLARE @count INT;
    
    SELECT @count = COUNT(*)
    FROM enterprise_pro_managed_organizations epm
    INNER JOIN organizations o ON o.id = epm.managed_organization_id
    WHERE epm.enterprise_pro_org_id = @enterprise_pro_org_id
    AND o.is_archived = 0;
    
    RETURN ISNULL(@count, 0);
END;
GO

-- ============================================================================
-- 7. FUNCIÓN: Verificar si un usuario es admin_global de una organización Enterprise Pro
-- ============================================================================

CREATE OR ALTER FUNCTION fn_is_enterprise_pro_admin(@user_id UNIQUEIDENTIFIER, @enterprise_pro_org_id UNIQUEIDENTIFIER)
RETURNS BIT
AS
BEGIN
    DECLARE @is_admin BIT = 0;
    
    -- Verificar si el usuario es admin_global en la organización Enterprise Pro
    SELECT @is_admin = CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END
    FROM organization_members
    WHERE organization_id = @enterprise_pro_org_id
    AND user_id = @user_id
    AND role = 'admin_global'
    AND left_at IS NULL;
    
    RETURN @is_admin;
END;
GO

-- ============================================================================
-- 8. TRIGGER: Validar límite de organizaciones gestionadas antes de crear relación
-- ============================================================================

CREATE OR ALTER TRIGGER trg_ep_managed_check_limit
ON enterprise_pro_managed_organizations
AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @enterprise_pro_org_id UNIQUEIDENTIFIER;
    DECLARE @can_manage BIT;
    
    SELECT @enterprise_pro_org_id = enterprise_pro_org_id FROM inserted;
    
    -- Verificar si puede gestionar más organizaciones
    SET @can_manage = dbo.fn_can_manage_more_organizations(@enterprise_pro_org_id);
    
    IF @can_manage = 0
    BEGIN
        DECLARE @max_orgs INT;
        DECLARE @current_orgs INT;
        
        -- Obtener límite y cantidad actual
        SELECT @max_orgs = p.max_organizations,
               @current_orgs = dbo.fn_get_managed_organizations_count(@enterprise_pro_org_id)
        FROM subscriptions s
        INNER JOIN plans p ON p.id = s.plan_id
        WHERE s.organization_id = @enterprise_pro_org_id
        AND s.status IN ('active', 'trialing');
        
        ROLLBACK TRANSACTION;
        THROW 50004, 
            CONCAT('Límite de organizaciones gestionadas excedido. Máximo permitido: ', @max_orgs, '. Actual: ', @current_orgs, '. Solo planes Enterprise Pro pueden gestionar múltiples organizaciones.'),
            1;
    END
END;
GO

-- ============================================================================
-- 9. PROCEDIMIENTO: Crear organización gestionada desde Enterprise Pro
-- ============================================================================

CREATE OR ALTER PROCEDURE sp_create_managed_organization
    @enterprise_pro_org_id UNIQUEIDENTIFIER,
    @organization_name VARCHAR(255),
    @organization_slug VARCHAR(100) = NULL,
    @created_by_user_id UNIQUEIDENTIFIER,
    @organization_id UNIQUEIDENTIFIER OUTPUT,
    @message VARCHAR(500) OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Validar que la organización Enterprise Pro existe y tiene plan Enterprise Pro
        DECLARE @plan_id VARCHAR(50);
        DECLARE @max_orgs INT;
        DECLARE @enterprise_pro_name VARCHAR(255);
        
        SELECT @plan_id = s.plan_id,
               @max_orgs = p.max_organizations,
               @enterprise_pro_name = o.name
        FROM organizations o
        INNER JOIN subscriptions s ON s.organization_id = o.id
        INNER JOIN plans p ON p.id = s.plan_id
        WHERE o.id = @enterprise_pro_org_id
        AND s.status IN ('active', 'trialing');
        
        IF @plan_id IS NULL
        BEGIN
            ROLLBACK TRANSACTION;
            SET @message = 'La organización Enterprise Pro no tiene una suscripción activa.';
            RETURN;
        END
        
        IF @plan_id != 'enterprise_pro'
        BEGIN
            ROLLBACK TRANSACTION;
            SET @message = 'Solo las organizaciones con plan Enterprise Pro pueden gestionar múltiples organizaciones.';
            RETURN;
        END
        
        -- Validar que el usuario es admin_global
        IF dbo.fn_is_enterprise_pro_admin(@created_by_user_id, @enterprise_pro_org_id) = 0
        BEGIN
            ROLLBACK TRANSACTION;
            SET @message = 'Solo los usuarios con rol admin_global pueden crear organizaciones gestionadas.';
            RETURN;
        END
        
        -- Validar límite de organizaciones gestionadas
        DECLARE @can_manage BIT;
        SET @can_manage = dbo.fn_can_manage_more_organizations(@enterprise_pro_org_id);
        
        IF @can_manage = 0
        BEGIN
            DECLARE @current_orgs INT;
            SET @current_orgs = dbo.fn_get_managed_organizations_count(@enterprise_pro_org_id);
            
            ROLLBACK TRANSACTION;
            SET @message = CONCAT('Límite de organizaciones gestionadas excedido. Máximo: ', @max_orgs, ', Actual: ', @current_orgs);
            RETURN;
        END
        
        -- Generar slug si no se proporciona
        DECLARE @final_slug VARCHAR(100);
        IF @organization_slug IS NULL OR @organization_slug = ''
        BEGIN
            SET @final_slug = LOWER(REPLACE(REPLACE(REPLACE(@organization_name, ' ', '-'), '.', ''), '_', '-'));
            
            -- Verificar unicidad y agregar sufijo si es necesario
            DECLARE @counter INT = 1;
            DECLARE @temp_slug VARCHAR(100) = @final_slug;
            
            WHILE EXISTS (SELECT 1 FROM organizations WHERE slug = @temp_slug)
            BEGIN
                SET @temp_slug = @final_slug + '-' + CAST(@counter AS VARCHAR);
                SET @counter = @counter + 1;
            END
            
            SET @final_slug = @temp_slug;
        END
        ELSE
        BEGIN
            SET @final_slug = @organization_slug;
            
            -- Verificar unicidad del slug
            IF EXISTS (SELECT 1 FROM organizations WHERE slug = @final_slug)
            BEGIN
                ROLLBACK TRANSACTION;
                SET @message = 'El slug ya está en uso. Por favor, elige otro.';
                RETURN;
            END
        END
        
        -- Crear la organización (independiente, sin jerarquía)
        SET @organization_id = NEWID();
        
        INSERT INTO organizations (
            id,
            name,
            slug,
            metadata,
            created_at,
            updated_at
        )
        VALUES (
            @organization_id,
            @organization_name,
            @final_slug,
            JSON_OBJECT('created_by': CAST(@created_by_user_id AS VARCHAR(36)), 'managed_by_enterprise_pro': CAST(@enterprise_pro_org_id AS VARCHAR(36))),
            GETUTCDATE(),
            GETUTCDATE()
        );
        
        -- Crear relación de gestión (no jerarquía)
        INSERT INTO enterprise_pro_managed_organizations (
            enterprise_pro_org_id,
            managed_organization_id,
            admin_user_id,
            created_at,
            updated_at
        )
        VALUES (
            @enterprise_pro_org_id,
            @organization_id,
            @created_by_user_id,
            GETUTCDATE(),
            GETUTCDATE()
        );
        
        -- Crear suscripción automática con plan free_trial para la organización gestionada
        INSERT INTO subscriptions (
            organization_id,
            plan_id,
            status,
            billing_cycle,
            current_period_start,
            current_period_end,
            metadata,
            created_at,
            updated_at
        )
        VALUES (
            @organization_id,
            'free_trial',
            'trialing',
            NULL,
            GETUTCDATE(),
            DATEADD(DAY, 30, GETUTCDATE()),
            JSON_OBJECT('managed_by_enterprise_pro': CAST(@enterprise_pro_org_id AS VARCHAR(36)), 'created_by': CAST(@created_by_user_id AS VARCHAR(36))),
            GETUTCDATE(),
            GETUTCDATE()
        );
        
        COMMIT TRANSACTION;
        SET @message = CONCAT('Organización "', @organization_name, '" creada exitosamente y gestionada por "', @enterprise_pro_name, '".');
        
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SET @message = ERROR_MESSAGE();
        SET @organization_id = NULL;
    END CATCH
END;
GO

-- ============================================================================
-- 10. VISTA: Organizaciones Enterprise Pro con sus organizaciones gestionadas
-- ============================================================================

CREATE OR ALTER VIEW vw_enterprise_pro_organizations AS
SELECT 
    o.id AS enterprise_pro_org_id,
    o.name AS enterprise_pro_name,
    o.slug AS enterprise_pro_slug,
    s.plan_id,
    s.status AS subscription_status,
    p.max_organizations,
    dbo.fn_get_managed_organizations_count(o.id) AS current_managed_organizations,
    p.max_organizations - dbo.fn_get_managed_organizations_count(o.id) AS remaining_slots,
    (
        SELECT COUNT(*) 
        FROM organization_members om 
        WHERE om.organization_id = o.id 
        AND om.role = 'admin_global' 
        AND om.left_at IS NULL
    ) AS admin_global_count,
    o.created_at AS enterprise_pro_created_at
FROM organizations o
INNER JOIN subscriptions s ON s.organization_id = o.id
INNER JOIN plans p ON p.id = s.plan_id
WHERE p.id = 'enterprise_pro'
AND s.status IN ('active', 'trialing')
AND o.is_archived = 0;
GO

-- ============================================================================
-- 11. VISTA: Organizaciones gestionadas con información del Enterprise Pro
-- ============================================================================

CREATE OR ALTER VIEW vw_managed_organizations AS
SELECT 
    managed.id AS managed_organization_id,
    managed.name AS managed_name,
    managed.slug AS managed_slug,
    ep_org.id AS enterprise_pro_org_id,
    ep_org.name AS enterprise_pro_name,
    ep_org.slug AS enterprise_pro_slug,
    epm.admin_user_id,
    u.name AS admin_user_name,
    u.email AS admin_user_email,
    managed.is_archived AS managed_is_archived,
    managed.created_at AS managed_created_at,
    (
        SELECT COUNT(*) 
        FROM organization_members om 
        WHERE om.organization_id = managed.id 
        AND om.left_at IS NULL
    ) AS managed_member_count,
    (
        SELECT COUNT(*) 
        FROM reports r 
        WHERE r.organization_id = managed.id 
        AND r.is_deleted = 0
    ) AS managed_reports_count
FROM organizations managed
INNER JOIN enterprise_pro_managed_organizations epm ON epm.managed_organization_id = managed.id
INNER JOIN organizations ep_org ON ep_org.id = epm.enterprise_pro_org_id
INNER JOIN users u ON u.id = epm.admin_user_id
WHERE managed.is_archived = 0;
GO

-- ============================================================================
-- 12. FUNCIÓN: Obtener organizaciones gestionadas por un usuario admin_global
-- ============================================================================

CREATE OR ALTER FUNCTION fn_get_user_managed_organizations(@user_id UNIQUEIDENTIFIER)
RETURNS TABLE
AS
RETURN
(
    SELECT 
        o.id AS organization_id,
        o.name AS organization_name,
        o.slug AS organization_slug,
        ep_org.id AS enterprise_pro_org_id,
        ep_org.name AS enterprise_pro_name
    FROM organizations o
    INNER JOIN enterprise_pro_managed_organizations epm ON epm.managed_organization_id = o.id
    INNER JOIN organizations ep_org ON ep_org.id = epm.enterprise_pro_org_id
    WHERE epm.admin_user_id = @user_id
    AND o.is_archived = 0
);
GO

-- ============================================================================
-- 13. FUNCIÓN: Verificar si un usuario puede gestionar una organización
-- ============================================================================

CREATE OR ALTER FUNCTION fn_can_user_manage_organization(
    @user_id UNIQUEIDENTIFIER,
    @organization_id UNIQUEIDENTIFIER
)
RETURNS BIT
AS
BEGIN
    -- Verificar si el usuario es admin_global y la organización está gestionada por su Enterprise Pro
    IF EXISTS (
        SELECT 1 
        FROM enterprise_pro_managed_organizations epm
        INNER JOIN organization_members om ON om.organization_id = epm.enterprise_pro_org_id
        WHERE epm.managed_organization_id = @organization_id
        AND epm.admin_user_id = @user_id
        AND om.user_id = @user_id
        AND om.role = 'admin_global'
        AND om.left_at IS NULL
    )
        RETURN 1;
    
    -- Verificar si el usuario es admin normal de la organización
    IF EXISTS (
        SELECT 1 
        FROM organization_members 
        WHERE organization_id = @organization_id 
        AND user_id = @user_id 
        AND role = 'admin'
        AND left_at IS NULL
    )
        RETURN 1;
    
    RETURN 0;
END;
GO

-- ============================================================================
-- 14. ÍNDICES ADICIONALES
-- ============================================================================

CREATE INDEX idx_ep_managed_ep_org_archived ON enterprise_pro_managed_organizations(enterprise_pro_org_id)
INCLUDE (managed_organization_id);
GO

-- ============================================================================
-- 15. COMENTARIOS Y DOCUMENTACIÓN
-- ============================================================================

PRINT '✅ Plan Enterprise Pro V2 configurado exitosamente';
PRINT '📊 Características:';
PRINT '   - Hasta 5 organizaciones gestionadas (no hijas)';
PRINT '   - 50 usuarios por organización Enterprise Pro';
PRINT '   - 1000 reportes totales';
PRINT '   - 200GB de almacenamiento';
PRINT '   - Separación completa de metadata por organización';
PRINT '   - Rol admin_global para gestionar múltiples organizaciones';
PRINT '';
PRINT '🔧 Funciones creadas:';
PRINT '   - fn_can_manage_more_organizations()';
PRINT '   - fn_get_managed_organizations_count()';
PRINT '   - fn_is_enterprise_pro_admin()';
PRINT '   - fn_get_user_managed_organizations()';
PRINT '   - fn_can_user_manage_organization()';
PRINT '';
PRINT '📋 Vistas creadas:';
PRINT '   - vw_enterprise_pro_organizations';
PRINT '   - vw_managed_organizations';
PRINT '';
PRINT '⚙️ Procedimientos creados:';
PRINT '   - sp_create_managed_organization()';
PRINT '';
PRINT '🔒 Seguridad:';
PRINT '   - Solo usuarios con rol admin_global pueden crear organizaciones gestionadas';
PRINT '   - Organizaciones son independientes, sin jerarquía';
PRINT '   - Metadata completamente separada por organization_id';
PRINT '   - Usuarios pueden pertenecer a múltiples organizaciones con roles distintos';
GO







-- ============================================================================
-- REPORT TUNER - Useful Queries
-- ============================================================================
-- Consultas útiles para operaciones comunes del sistema
-- ============================================================================

USE empower_reports;
GO

-- ============================================================================
-- CONSULTAS DE VALIDACIÓN DE LÍMITES
-- ============================================================================

-- Verificar si una organización puede agregar más usuarios
-- Ejemplo: SELECT dbo.fn_can_add_user('YOUR_ORG_ID')
-- Retorna 1 si puede, 0 si no puede

-- Verificar si una organización puede agregar más reportes
-- Ejemplo: SELECT dbo.fn_can_add_report('YOUR_ORG_ID')
-- Retorna 1 si puede, 0 si no puede

-- Contar usuarios activos de una organización
SELECT 
    COUNT(*) AS active_users_count
FROM organization_members
WHERE organization_id = 'YOUR_ORG_ID'
AND left_at IS NULL;
 
-- Contar reportes activos de una organización
SELECT 
    COUNT(*) AS active_reports_count
FROM reports
WHERE organization_id = 'YOUR_ORG_ID'
AND is_deleted = 0;

-- ============================================================================
-- CONSULTAS DE SUSCRIPCIONES
-- ============================================================================

-- Obtener todas las organizaciones con sus planes y límites
SELECT * FROM vw_organizations_with_subscription;

-- Obtener suscripciones que están por vencer en los próximos 7 días
SELECT 
    o.name AS organization_name,
    s.id AS subscription_id,
    s.plan_id,
    s.status,
    s.current_period_end,
    DATEDIFF(day, GETUTCDATE(), s.current_period_end) AS days_until_expiry
FROM subscriptions s
INNER JOIN organizations o ON o.id = s.organization_id
WHERE s.status = 'active'
AND s.current_period_end BETWEEN GETUTCDATE() AND DATEADD(day, 7, GETUTCDATE())
ORDER BY s.current_period_end;

-- Obtener suscripciones en período de prueba que están por vencer
SELECT 
    o.name AS organization_name,
    s.trial_end,
    DATEDIFF(day, GETUTCDATE(), s.trial_end) AS days_until_trial_end
FROM subscriptions s
INNER JOIN organizations o ON o.id = s.organization_id
WHERE s.status = 'trialing'
AND s.trial_end BETWEEN GETUTCDATE() AND DATEADD(day, 3, GETUTCDATE())
ORDER BY s.trial_end;

-- Historial completo de cambios de plan de una organización
SELECT 
    sh.event_type,
    p_old.name AS plan_from,
    p_new.name AS plan_to,
    sh.status_old,
    sh.status_new,
    u.name AS changed_by_user,
    sh.created_at
FROM subscription_history sh
LEFT JOIN plans p_old ON p_old.id = sh.plan_id_old
LEFT JOIN plans p_new ON p_new.id = sh.plan_id_new
LEFT JOIN users u ON u.id = sh.changed_by
WHERE sh.organization_id = 'YOUR_ORG_ID'
ORDER BY sh.created_at DESC;

-- ============================================================================
-- CONSULTAS DE USUARIOS Y ORGANIZACIONES
-- ============================================================================

-- Obtener todos los usuarios con sus organizaciones principales
SELECT * FROM vw_users_with_primary_org;

-- Obtener todos los miembros de una organización con sus roles
SELECT 
    u.email,
    u.name,
    om.role,
    om.is_primary,
    om.joined_at
FROM organization_members om
INNER JOIN users u ON u.id = om.user_id
WHERE om.organization_id = 'YOUR_ORG_ID'
AND om.left_at IS NULL
ORDER BY om.role DESC, om.joined_at;

-- Obtener todas las organizaciones de un usuario
SELECT 
    o.name AS organization_name,
    om.role,
    om.is_primary,
    s.plan_id,
    s.status AS subscription_status
FROM organization_members om
INNER JOIN organizations o ON o.id = om.organization_id
LEFT JOIN subscriptions s ON s.organization_id = o.id AND s.status IN ('active', 'trialing')
WHERE om.user_id = 'YOUR_USER_ID'
AND om.left_at IS NULL;

-- ============================================================================
-- CONSULTAS DE REPORTES
-- ============================================================================

-- Obtener todos los reportes de una organización con información del usuario
SELECT 
    r.name,
    r.original_filename,
    r.status,
    r.file_size_bytes,
    u.name AS uploaded_by,
    r.created_at
FROM reports r
INNER JOIN users u ON u.id = r.user_id
WHERE r.organization_id = 'YOUR_ORG_ID'
AND r.is_deleted = 0
ORDER BY r.created_at DESC;

-- Obtener reportes que están siendo procesados hace más de 30 minutos (posible error)
SELECT 
    r.id,
    r.name,
    r.organization_id,
    DATEDIFF(minute, r.processing_started_at, GETUTCDATE()) AS processing_time_minutes
FROM reports r
WHERE r.status = 'processing'
AND r.processing_started_at < DATEADD(minute, -30, GETUTCDATE());

-- Estadísticas de reportes por organización
SELECT 
    o.name AS organization_name,
    COUNT(*) AS total_reports,
    SUM(CASE WHEN r.status = 'processed' THEN 1 ELSE 0 END) AS processed_reports,
    SUM(CASE WHEN r.status = 'processing' THEN 1 ELSE 0 END) AS processing_reports,
    SUM(CASE WHEN r.status = 'failed' THEN 1 ELSE 0 END) AS failed_reports,
    SUM(r.file_size_bytes) / 1024.0 / 1024.0 AS total_size_mb
FROM reports r
INNER JOIN organizations o ON o.id = r.organization_id
WHERE r.is_deleted = 0
GROUP BY o.id, o.name
ORDER BY total_reports DESC;

-- ============================================================================
-- CONSULTAS DE AUDITORÍA Y MONITOREO
-- ============================================================================

-- Usuarios que no han iniciado sesión en los últimos 90 días
SELECT 
    u.email,
    u.name,
    u.last_login_at,
    DATEDIFF(day, u.last_login_at, GETUTCDATE()) AS days_since_last_login
FROM users u
WHERE u.is_active = 1
AND (u.last_login_at IS NULL OR u.last_login_at < DATEADD(day, -90, GETUTCDATE()))
ORDER BY u.last_login_at;

-- Organizaciones sin suscripción activa (necesitan suscripción)
SELECT 
    o.name,
    o.created_at,
    DATEDIFF(day, o.created_at, GETUTCDATE()) AS days_without_subscription
FROM organizations o
WHERE o.is_archived = 0
AND NOT EXISTS (
    SELECT 1 FROM subscriptions s 
    WHERE s.organization_id = o.id 
    AND s.status IN ('active', 'trialing')
);

-- Estadísticas de uso por plan
SELECT 
    p.name AS plan_name,
    COUNT(DISTINCT s.organization_id) AS total_organizations,
    COUNT(DISTINCT om.user_id) AS total_users,
    COUNT(DISTINCT r.id) AS total_reports,
    AVG((SELECT COUNT(*) FROM organization_members om2 WHERE om2.organization_id = s.organization_id AND om2.left_at IS NULL)) AS avg_users_per_org,
    AVG((SELECT COUNT(*) FROM reports r2 WHERE r2.organization_id = s.organization_id AND r2.is_deleted = 0)) AS avg_reports_per_org
FROM plans p
LEFT JOIN subscriptions s ON s.plan_id = p.id AND s.status IN ('active', 'trialing')
LEFT JOIN organization_members om ON om.organization_id = s.organization_id AND om.left_at IS NULL
LEFT JOIN reports r ON r.organization_id = s.organization_id AND r.is_deleted = 0
GROUP BY p.id, p.name
ORDER BY total_organizations DESC;

-- ============================================================================
-- PROCEDIMIENTOS ALMACENADOS ÚTILES
-- ============================================================================

-- Procedimiento: Archivar organización
CREATE OR ALTER PROCEDURE sp_archive_organization
    @organization_id UNIQUEIDENTIFIER,
    @archived_by UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Archivar organización
        UPDATE organizations
        SET is_archived = 1,
            archived_at = GETUTCDATE(),
            updated_at = GETUTCDATE()
        WHERE id = @organization_id;
        
        -- Cancelar suscripción si existe
        UPDATE subscriptions
        SET status = 'canceled',
            canceled_at = GETUTCDATE(),
            updated_at = GETUTCDATE()
        WHERE organization_id = @organization_id
        AND status IN ('active', 'trialing');
        
        -- Registrar en historial
        INSERT INTO subscription_history (
            subscription_id,
            organization_id,
            plan_id_new,
            status_new,
            event_type,
            changed_by,
            metadata
        )
        SELECT 
            id,
            @organization_id,
            plan_id,
            'canceled',
            'canceled',
            @archived_by,
            '{"reason": "organization_archived"}' AS JSON
        FROM subscriptions
        WHERE organization_id = @organization_id
        AND status = 'canceled';
        
        COMMIT TRANSACTION;
        
        SELECT 1 AS success, 'Organización archivada exitosamente' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: Cambiar plan de organización
CREATE OR ALTER PROCEDURE sp_change_plan
    @organization_id UNIQUEIDENTIFIER,
    @new_plan_id VARCHAR(50),
    @changed_by UNIQUEIDENTIFIER,
    @billing_cycle VARCHAR(20) = NULL -- NULL para free_trial, 'monthly' o 'yearly' para otros
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @subscription_id UNIQUEIDENTIFIER;
    DECLARE @old_plan_id VARCHAR(50);
    DECLARE @old_status VARCHAR(50);
    DECLARE @final_billing_cycle VARCHAR(20);
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Determinar billing_cycle según el plan
        IF @new_plan_id = 'free_trial'
        BEGIN
            SET @final_billing_cycle = NULL; -- free_trial no tiene billing_cycle
        END
        ELSE
        BEGIN
            -- Para planes pagos, usar el proporcionado o default 'monthly'
            SET @final_billing_cycle = ISNULL(@billing_cycle, 'monthly');
        END
        
        -- Obtener suscripción actual
        SELECT 
            @subscription_id = id,
            @old_plan_id = plan_id,
            @old_status = status
        FROM subscriptions
        WHERE organization_id = @organization_id
        AND status IN ('active', 'trialing');
        
        IF @subscription_id IS NULL
        BEGIN
            -- Crear nueva suscripción si no existe
            SET @subscription_id = NEWID();
            INSERT INTO subscriptions (
                id,
                organization_id,
                plan_id,
                status,
                billing_cycle,
                current_period_start,
                current_period_end
            )
            VALUES (
                @subscription_id,
                @organization_id,
                @new_plan_id,
                CASE WHEN @new_plan_id = 'free_trial' THEN 'trialing' ELSE 'active' END,
                @final_billing_cycle,
                GETUTCDATE(),
                CASE 
                    WHEN @final_billing_cycle = 'yearly' THEN DATEADD(year, 1, GETUTCDATE())
                    WHEN @final_billing_cycle = 'monthly' THEN DATEADD(month, 1, GETUTCDATE())
                    ELSE DATEADD(day, 30, GETUTCDATE()) -- free_trial: 30 días
                END
            );
            
            SET @old_plan_id = NULL;
            SET @old_status = NULL;
        END
        ELSE
        BEGIN
            -- Actualizar suscripción existente
            UPDATE subscriptions
            SET plan_id = @new_plan_id,
                billing_cycle = @final_billing_cycle,
                status = CASE WHEN @new_plan_id = 'free_trial' AND status = 'active' THEN 'trialing' ELSE status END,
                updated_at = GETUTCDATE()
            WHERE id = @subscription_id;
        END
        
        -- Registrar en historial
        DECLARE @metadata JSON = '{"billing_cycle": ' + 
            CASE WHEN @final_billing_cycle IS NULL THEN 'null' ELSE '"' + @final_billing_cycle + '"' END + '}' AS JSON;
        
        INSERT INTO subscription_history (
            subscription_id,
            organization_id,
            plan_id_old,
            plan_id_new,
            status_old,
            status_new,
            event_type,
            changed_by,
            metadata
        )
        VALUES (
            @subscription_id,
            @organization_id,
            @old_plan_id,
            @new_plan_id,
            @old_status,
            (SELECT status FROM subscriptions WHERE id = @subscription_id),
            'plan_changed',
            @changed_by,
            @metadata
        );
        
        COMMIT TRANSACTION;
        
        SELECT 1 AS success, 'Plan cambiado exitosamente' AS message, @subscription_id AS subscription_id;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message, NULL AS subscription_id;
    END CATCH
END;
GO

-- ============================================================================
-- ÍNDICES PARA MEJORAR PERFORMANCE DE CONSULTAS COMUNES
-- ============================================================================

-- Estos índices ya están incluidos en el schema principal, pero se listan aquí para referencia:
-- - idx_reports_org_status_deleted (reports)
-- - idx_org_members_org_active (organization_members)
-- - idx_subscriptions_current_period_end (subscriptions)

PRINT '✅ Consultas útiles y procedimientos almacenados creados';
GO


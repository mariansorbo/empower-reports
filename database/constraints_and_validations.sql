-- ============================================================================
-- REPORT TUNER - Additional Constraints and Validations
-- ============================================================================
-- Constraints adicionales y validaciones para garantizar integridad de datos
-- según las reglas de negocio
-- ============================================================================

USE empower_reports;
GO

-- ============================================================================
-- CONSTRAINT: Validar billing_cycle según plan_id
-- ============================================================================
-- Regla: free_trial debe tener billing_cycle = NULL
--        Otros planes deben tener billing_cycle definido
-- ============================================================================

-- Crear función de validación
CREATE OR ALTER FUNCTION fn_validate_billing_cycle_for_plan(@plan_id VARCHAR(50), @billing_cycle VARCHAR(20))
RETURNS BIT
AS
BEGIN
    -- free_trial debe tener billing_cycle = NULL
    IF @plan_id = 'free_trial' AND @billing_cycle IS NOT NULL
        RETURN 0;
    
    -- Otros planes deben tener billing_cycle definido
    IF @plan_id != 'free_trial' AND @billing_cycle IS NULL
        RETURN 0;
    
    -- Si billing_cycle está definido, debe ser válido
    IF @billing_cycle IS NOT NULL AND @billing_cycle NOT IN ('monthly', 'yearly')
        RETURN 0;
    
    RETURN 1;
END;
GO

-- ============================================================================
-- TRIGGER: Validar billing_cycle según plan_id
-- ============================================================================

CREATE OR ALTER TRIGGER trg_validate_billing_cycle_by_plan
ON subscriptions
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Validar cada fila insertada/actualizada
    IF EXISTS (
        SELECT 1
        FROM inserted i
        WHERE dbo.fn_validate_billing_cycle_for_plan(i.plan_id, i.billing_cycle) = 0
    )
    BEGIN
        DECLARE @error_msg VARCHAR(500);
        
        SELECT TOP 1 @error_msg = CASE
            WHEN plan_id = 'free_trial' AND billing_cycle IS NOT NULL
                THEN 'El plan free_trial debe tener billing_cycle = NULL (no tiene facturación)'
            WHEN plan_id != 'free_trial' AND billing_cycle IS NULL
                THEN 'Los planes pagos deben tener billing_cycle definido (monthly o yearly)'
            ELSE 'billing_cycle inválido para el plan especificado'
        END
        FROM inserted
        WHERE dbo.fn_validate_billing_cycle_for_plan(plan_id, billing_cycle) = 0;
        
        ROLLBACK TRANSACTION;
        THROW 50030, @error_msg, 1;
    END
END;
GO

-- ============================================================================
-- VALIDACIÓN: organization_id NULL para usuarios individuales
-- ============================================================================
-- Los reportes pueden tener organization_id = NULL para usuarios individuales
-- (plan basic). Esto ya está soportado en el schema, pero agregamos validación
-- ============================================================================

-- Función: Verificar si un usuario individual puede crear reporte sin organización
CREATE OR ALTER FUNCTION fn_can_user_create_individual_report(@user_id UNIQUEIDENTIFIER)
RETURNS BIT
AS
BEGIN
    DECLARE @has_org BIT = 0;
    DECLARE @plan_id VARCHAR(50);
    
    -- Verificar si tiene organización primaria activa
    SELECT @has_org = 1
    FROM organization_members om
    INNER JOIN organizations o ON o.id = om.organization_id
    WHERE om.user_id = @user_id
    AND om.is_primary = 1
    AND om.left_at IS NULL
    AND o.is_archived = 0;
    
    -- Si tiene organización activa, debe usar organization_id
    IF @has_org = 1
        RETURN 0;
    
    -- Si no tiene organización, puede ser usuario individual (plan basic)
    RETURN 1;
END;
GO

-- Trigger: Validar que usuarios con organización usen organization_id
CREATE OR ALTER TRIGGER trg_reports_validate_organization_for_user
ON reports
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @user_id UNIQUEIDENTIFIER;
    DECLARE @organization_id UNIQUEIDENTIFIER;
    DECLARE @has_org BIT;
    
    SELECT @user_id = user_id, @organization_id = organization_id FROM inserted;
    
    -- Verificar si el usuario tiene organización activa
    SELECT @has_org = CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END
    FROM organization_members om
    INNER JOIN organizations o ON o.id = om.organization_id
    WHERE om.user_id = @user_id
    AND om.is_primary = 1
    AND om.left_at IS NULL
    AND o.is_archived = 0;
    
    -- Si tiene organización activa pero reporte no tiene organization_id
    IF @has_org = 1 AND @organization_id IS NULL
    BEGIN
        DECLARE @org_name VARCHAR(255);
        SELECT TOP 1 @org_name = o.name
        FROM organization_members om
        INNER JOIN organizations o ON o.id = om.organization_id
        WHERE om.user_id = @user_id
        AND om.is_primary = 1
        AND om.left_at IS NULL
        AND o.is_archived = 0;
        
        ROLLBACK TRANSACTION;
        THROW 50032, 
            CONCAT('El usuario pertenece a la organización "', @org_name, '". El reporte debe tener organization_id asignado.'),
            1;
    END
    
    -- Si NO tiene organización pero reporte tiene organization_id (válido, puede ser invitado)
    -- Esto está permitido, no hacemos nada
    
    -- Si NO tiene organización y reporte NO tiene organization_id (usuario individual)
    -- Esto está permitido, no hacemos nada
END;
GO

-- ============================================================================
-- VALIDACIÓN: Usuarios individuales deben tener límites del plan basic
-- ============================================================================

-- Función: Obtener plan efectivo de un usuario (individual o de su organización)
CREATE OR ALTER FUNCTION fn_get_user_effective_plan(@user_id UNIQUEIDENTIFIER)
RETURNS TABLE
AS
RETURN (
    SELECT TOP 1
        p.id AS plan_id,
        p.name AS plan_name,
        p.max_users,
        p.max_reports,
        p.max_storage_mb,
        CASE WHEN om.organization_id IS NOT NULL THEN om.organization_id ELSE NULL END AS organization_id
    FROM users u
    LEFT JOIN organization_members om ON om.user_id = u.id 
        AND om.is_primary = 1 
        AND om.left_at IS NULL
    LEFT JOIN organizations o ON o.id = om.organization_id AND o.is_archived = 0
    LEFT JOIN subscriptions s ON s.organization_id = o.id AND s.status IN ('active', 'trialing')
    LEFT JOIN plans p ON p.id = ISNULL(s.plan_id, 'basic') -- Default a basic si no tiene org
    WHERE u.id = @user_id
    AND u.is_active = 1
    ORDER BY om.is_primary DESC
);
GO

PRINT '✅ Constraints y validaciones adicionales creadas exitosamente';
GO





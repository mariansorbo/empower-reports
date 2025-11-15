-- ============================================================================
-- REPORT TUNER - Organization Workflow Procedures & Triggers
-- ============================================================================
-- Procedimientos, triggers y funciones para el flujo UX de creación y unión
-- a organizaciones según el diseño especificado
-- ============================================================================

USE empower_reports;
GO

-- ============================================================================
-- FUNCIONES DE VALIDACIÓN Y CONSULTA
-- ============================================================================

-- Función: Verificar si un usuario puede crear una nueva organización
-- Valida si ya tiene una organización activa como admin
CREATE OR ALTER FUNCTION fn_can_user_create_organization(@user_id UNIQUEIDENTIFIER)
RETURNS @result TABLE (
    can_create BIT,
    has_existing_org BIT,
    existing_org_id UNIQUEIDENTIFIER,
    existing_org_name VARCHAR(255),
    reason VARCHAR(500)
)
AS
BEGIN
    DECLARE @has_org BIT = 0;
    DECLARE @org_id UNIQUEIDENTIFIER;
    DECLARE @org_name VARCHAR(255);
    
    -- Verificar si el usuario ya es admin de alguna organización activa
    SELECT TOP 1
        @has_org = 1,
        @org_id = o.id,
        @org_name = o.name
    FROM organization_members om
    INNER JOIN organizations o ON o.id = om.organization_id
    WHERE om.user_id = @user_id
    AND om.role = 'admin'
    AND om.left_at IS NULL
    AND o.is_archived = 0;
    
    IF @has_org = 1
    BEGIN
        INSERT INTO @result VALUES (1, 1, @org_id, @org_name, 'Usuario tiene organización existente pero puede crear otra o archivarla');
    END
    ELSE
    BEGIN
        INSERT INTO @result VALUES (1, 0, NULL, NULL, 'Usuario puede crear organización');
    END
    
    RETURN;
END;
GO

-- Función: Validar código de invitación y obtener información de la organización
CREATE OR ALTER FUNCTION fn_validate_invitation_token(@invitation_token VARCHAR(255))
RETURNS @result TABLE (
    is_valid BIT,
    organization_id UNIQUEIDENTIFIER,
    organization_name VARCHAR(255),
    admin_name VARCHAR(255),
    admin_email VARCHAR(255),
    member_count INT,
    plan_name VARCHAR(100),
    reason VARCHAR(500)
)
AS
BEGIN
    DECLARE @org_id UNIQUEIDENTIFIER;
    DECLARE @org_name VARCHAR(255);
    DECLARE @admin_id UNIQUEIDENTIFIER;
    DECLARE @expires_at DATETIME2;
    
    -- Buscar invitación válida
    SELECT 
        @org_id = om.organization_id,
        @org_name = o.name,
        @admin_id = om.invited_by,
        @expires_at = om.invitation_expires_at
    FROM organization_members om
    INNER JOIN organizations o ON o.id = om.organization_id
    WHERE om.invitation_token = @invitation_token
    AND om.invitation_expires_at > GETUTCDATE();
    
    IF @org_id IS NULL
    BEGIN
        INSERT INTO @result VALUES (0, NULL, NULL, NULL, NULL, NULL, NULL, 'Token de invitación inválido o expirado');
        RETURN;
    END
    
    IF EXISTS (SELECT 1 FROM organizations WHERE id = @org_id AND is_archived = 1)
    BEGIN
        INSERT INTO @result VALUES (0, @org_id, @org_name, NULL, NULL, NULL, NULL, 'La organización está archivada');
        RETURN;
    END
    
    -- Obtener información del admin
    DECLARE @admin_name VARCHAR(255);
    DECLARE @admin_email VARCHAR(255);
    
    SELECT 
        @admin_name = name,
        @admin_email = email
    FROM users
    WHERE id = @admin_id;
    
    -- Contar miembros activos
    DECLARE @member_count INT;
    SELECT @member_count = COUNT(*)
    FROM organization_members
    WHERE organization_id = @org_id
    AND left_at IS NULL;
    
    -- Obtener plan actual
    DECLARE @plan_name VARCHAR(100);
    SELECT TOP 1 @plan_name = p.name
    FROM subscriptions s
    INNER JOIN plans p ON p.id = s.plan_id
    WHERE s.organization_id = @org_id
    AND s.status IN ('active', 'trialing');
    
    INSERT INTO @result VALUES (
        1, 
        @org_id, 
        @org_name, 
        @admin_name, 
        @admin_email, 
        @member_count,
        ISNULL(@plan_name, 'free_trial'),
        'Invitación válida'
    );
    
    RETURN;
END;
GO

-- Función: Obtener todas las organizaciones de un usuario con su estado
CREATE OR ALTER FUNCTION fn_get_user_organizations(@user_id UNIQUEIDENTIFIER)
RETURNS TABLE
AS
RETURN (
    SELECT 
        o.id AS organization_id,
        o.name AS organization_name,
        o.slug,
        om.role,
        om.is_primary,
        o.is_archived,
        o.archived_at,
        s.plan_id,
        p.name AS plan_name,
        s.status AS subscription_status,
        s.current_period_end,
        (SELECT COUNT(*) FROM organization_members om2 WHERE om2.organization_id = o.id AND om2.left_at IS NULL) AS member_count,
        (SELECT COUNT(*) FROM reports r WHERE r.organization_id = o.id AND r.is_deleted = 0) AS reports_count,
        om.joined_at,
        CASE 
            WHEN o.is_archived = 1 THEN 'Archivada'
            WHEN s.status = 'active' THEN 'Activa'
            WHEN s.status = 'trialing' THEN 'En prueba'
            WHEN s.status = 'past_due' THEN 'Pago pendiente'
            ELSE 'Inactiva'
        END AS status_label
    FROM organization_members om
    INNER JOIN organizations o ON o.id = om.organization_id
    LEFT JOIN subscriptions s ON s.organization_id = o.id AND s.status IN ('active', 'trialing')
    LEFT JOIN plans p ON p.id = s.plan_id
    WHERE om.user_id = @user_id
    AND om.left_at IS NULL
);
GO

-- ============================================================================
-- PROCEDIMIENTOS PARA FLUJO UX
-- ============================================================================

-- Procedimiento: (2) Flujo A - Crear organización nueva
-- Crea organización, asigna plan free_trial y hace al usuario admin
CREATE OR ALTER PROCEDURE sp_create_organization_with_user
    @user_id UNIQUEIDENTIFIER,
    @organization_name VARCHAR(255),
    @make_primary BIT = 1 -- Si debe ser organización primaria
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @organization_id UNIQUEIDENTIFIER = NEWID();
    DECLARE @subscription_id UNIQUEIDENTIFIER = NEWID();
    DECLARE @slug VARCHAR(100);
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Validar que el usuario existe
        IF NOT EXISTS (SELECT 1 FROM users WHERE id = @user_id AND is_active = 1)
        BEGIN
            THROW 50020, 'Usuario no válido o inactivo', 1;
        END
        
        -- Generar slug único
        SET @slug = LOWER(REPLACE(REPLACE(REPLACE(REPLACE(@organization_name, ' ', '-'), '''', ''), '.', ''), 'áéíóúñ', 'aeioun'));
        
        -- Asegurar que el slug sea único
        DECLARE @slug_counter INT = 1;
        DECLARE @final_slug VARCHAR(100) = @slug;
        WHILE EXISTS (SELECT 1 FROM organizations WHERE slug = @final_slug)
        BEGIN
            SET @final_slug = @slug + '-' + CAST(@slug_counter AS VARCHAR(10));
            SET @slug_counter = @slug_counter + 1;
        END
        
        -- Crear organización
        INSERT INTO organizations (
            id,
            name,
            slug,
            is_archived
        )
        VALUES (
            @organization_id,
            @organization_name,
            @final_slug,
            0
        );
        
        -- El trigger trg_organization_auto_assign_free_trial creará automáticamente la suscripción
        -- Esperar un momento para que el trigger se ejecute
        WAITFOR DELAY '00:00:00.1';
        
        -- Obtener el subscription_id creado por el trigger
        SELECT @subscription_id = id 
        FROM subscriptions 
        WHERE organization_id = @organization_id;
        
        -- Si el usuario no tiene organización primaria, o si se solicita hacerla primaria
        DECLARE @has_primary BIT;
        SELECT @has_primary = CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END
        FROM organization_members
        WHERE user_id = @user_id AND is_primary = 1 AND left_at IS NULL;
        
        -- Agregar usuario como admin
        INSERT INTO organization_members (
            organization_id,
            user_id,
            role,
            is_primary,
            joined_at
        )
        VALUES (
            @organization_id,
            @user_id,
            'admin',
            CASE WHEN @has_primary = 0 OR @make_primary = 1 THEN 1 ELSE 0 END,
            GETUTCDATE()
        );
        
        -- Si se establece como primaria y ya había una, quitar primaria de la anterior
        IF @make_primary = 1 AND @has_primary = 1
        BEGIN
            UPDATE organization_members
            SET is_primary = 0,
                updated_at = GETUTCDATE()
            WHERE user_id = @user_id
            AND is_primary = 1
            AND organization_id != @organization_id
            AND left_at IS NULL;
        END
        
        COMMIT TRANSACTION;
        
        SELECT 
            1 AS success,
            @organization_id AS organization_id,
            @organization_name AS organization_name,
            @subscription_id AS subscription_id,
            'free_trial' AS plan_id,
            'Organización creada exitosamente' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 
            0 AS success,
            NULL AS organization_id,
            NULL AS organization_name,
            NULL AS subscription_id,
            NULL AS plan_id,
            ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: (3) Flujo B - Unirse a organización con código de invitación
CREATE OR ALTER PROCEDURE sp_join_organization_by_invitation
    @user_id UNIQUEIDENTIFIER,
    @invitation_token VARCHAR(255),
    @accept_invitation BIT = 1
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @organization_id UNIQUEIDENTIFIER;
    DECLARE @organization_name VARCHAR(255);
    DECLARE @invited_by UNIQUEIDENTIFIER;
    DECLARE @already_member BIT = 0;
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Validar token usando la función
        SELECT 
            @organization_id = organization_id,
            @organization_name = organization_name
        FROM fn_validate_invitation_token(@invitation_token)
        WHERE is_valid = 1;
        
        IF @organization_id IS NULL
        BEGIN
            SELECT 
                0 AS success,
                NULL AS organization_id,
                NULL AS organization_name,
                'Token de invitación inválido o expirado' AS message;
            RETURN;
        END
        
        -- Verificar si el usuario ya es miembro
        IF EXISTS (
            SELECT 1 FROM organization_members 
            WHERE organization_id = @organization_id 
            AND user_id = @user_id 
            AND left_at IS NULL
        )
        BEGIN
            SET @already_member = 1;
            SELECT 
                1 AS success,
                @organization_id AS organization_id,
                @organization_name AS organization_name,
                'Usuario ya es miembro de esta organización' AS message;
            COMMIT TRANSACTION;
            RETURN;
        END
        
        IF @accept_invitation = 1
        BEGIN
            -- Obtener quien invitó
            SELECT @invited_by = invited_by
            FROM organization_members
            WHERE invitation_token = @invitation_token
            AND organization_id = @organization_id;
            
            -- Verificar si usuario tiene organización existente
            DECLARE @has_existing_org BIT;
            DECLARE @existing_org_id UNIQUEIDENTIFIER;
            SELECT 
                @has_existing_org = CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END,
                @existing_org_id = MAX(o.id)
            FROM organization_members om
            INNER JOIN organizations o ON o.id = om.organization_id
            WHERE om.user_id = @user_id
            AND om.role = 'admin'
            AND om.left_at IS NULL
            AND o.is_archived = 0;
            
            -- Agregar usuario como miembro
            INSERT INTO organization_members (
                organization_id,
                user_id,
                role,
                is_primary,
                invited_by,
                invitation_token,
                joined_at
            )
            VALUES (
                @organization_id,
                @user_id,
                'member',
                0, -- No será primaria inicialmente (se manejará en el flujo de decisión)
                @invited_by,
                NULL, -- Limpiar token después de usar
                GETUTCDATE()
            );
            
            -- Limpiar token de invitación
            UPDATE organization_members
            SET invitation_token = NULL,
                invitation_expires_at = NULL
            WHERE invitation_token = @invitation_token;
            
            COMMIT TRANSACTION;
            
            SELECT 
                1 AS success,
                @organization_id AS organization_id,
                @organization_name AS organization_name,
                @has_existing_org AS has_existing_organization,
                @existing_org_id AS existing_organization_id,
                'Usuario unido a organización exitosamente' AS message;
        END
        ELSE
        BEGIN
            COMMIT TRANSACTION;
            SELECT 
                0 AS success,
                @organization_id AS organization_id,
                @organization_name AS organization_name,
                0 AS has_existing_organization,
                NULL AS existing_organization_id,
                'Invitación rechazada' AS message;
        END
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 
            0 AS success,
            NULL AS organization_id,
            NULL AS organization_name,
            0 AS has_existing_organization,
            NULL AS existing_organization_id,
            ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: (4) Flujo de decisión - Archivar organización actual y unirse a nueva
CREATE OR ALTER PROCEDURE sp_archive_and_join_organization
    @user_id UNIQUEIDENTIFIER,
    @old_organization_id UNIQUEIDENTIFIER,
    @new_organization_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Validar que el usuario es admin de la organización antigua
        IF NOT EXISTS (
            SELECT 1 FROM organization_members
            WHERE user_id = @user_id
            AND organization_id = @old_organization_id
            AND role = 'admin'
            AND left_at IS NULL
        )
        BEGIN
            THROW 50021, 'El usuario no es administrador de la organización antigua', 1;
        END
        
        -- Validar que el usuario es miembro de la nueva organización
        IF NOT EXISTS (
            SELECT 1 FROM organization_members
            WHERE user_id = @user_id
            AND organization_id = @new_organization_id
            AND left_at IS NULL
        )
        BEGIN
            THROW 50022, 'El usuario no es miembro de la nueva organización', 1;
        END
        
        -- Archivar organización antigua
        UPDATE organizations
        SET is_archived = 1,
            archived_at = GETUTCDATE(),
            updated_at = GETUTCDATE()
        WHERE id = @old_organization_id;
        
        -- Cancelar suscripción de la organización antigua si existe
        UPDATE subscriptions
        SET status = 'canceled',
            canceled_at = GETUTCDATE(),
            cancel_at_period_end = 0,
            updated_at = GETUTCDATE()
        WHERE organization_id = @old_organization_id
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
            @old_organization_id,
            plan_id,
            'canceled',
            'canceled',
            @user_id,
            '{"reason": "organization_archived_on_join", "new_organization_id": "' + CAST(@new_organization_id AS VARCHAR(36)) + '"}' AS JSON
        FROM subscriptions
        WHERE organization_id = @old_organization_id
        AND status = 'canceled';
        
        -- Establecer nueva organización como primaria
        -- Quitar primaria de todas las organizaciones del usuario
        UPDATE organization_members
        SET is_primary = 0,
            updated_at = GETUTCDATE()
        WHERE user_id = @user_id
        AND left_at IS NULL;
        
        -- Establecer nueva como primaria
        UPDATE organization_members
        SET is_primary = 1,
            updated_at = GETUTCDATE()
        WHERE user_id = @user_id
        AND organization_id = @new_organization_id;
        
        COMMIT TRANSACTION;
        
        SELECT 
            1 AS success,
            @new_organization_id AS new_organization_id,
            @old_organization_id AS archived_organization_id,
            'Organización archivada y usuario unido a nueva organización' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 
            0 AS success,
            NULL AS new_organization_id,
            NULL AS archived_organization_id,
            ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: (4) Mantener ambas organizaciones y establecer nueva como primaria
CREATE OR ALTER PROCEDURE sp_keep_both_set_new_primary
    @user_id UNIQUEIDENTIFIER,
    @new_organization_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Validar que el usuario es miembro de la nueva organización
        IF NOT EXISTS (
            SELECT 1 FROM organization_members
            WHERE user_id = @user_id
            AND organization_id = @new_organization_id
            AND left_at IS NULL
        )
        BEGIN
            THROW 50023, 'El usuario no es miembro de la nueva organización', 1;
        END
        
        -- Quitar primaria de todas las organizaciones del usuario
        UPDATE organization_members
        SET is_primary = 0,
            updated_at = GETUTCDATE()
        WHERE user_id = @user_id
        AND left_at IS NULL;
        
        -- Establecer nueva organización como primaria
        UPDATE organization_members
        SET is_primary = 1,
            updated_at = GETUTCDATE()
        WHERE user_id = @user_id
        AND organization_id = @new_organization_id;
        
        COMMIT TRANSACTION;
        
        SELECT 
            1 AS success,
            @new_organization_id AS primary_organization_id,
            'Nueva organización establecida como primaria' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 
            0 AS success,
            NULL AS primary_organization_id,
            ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: Cambiar organización primaria del usuario
CREATE OR ALTER PROCEDURE sp_change_primary_organization
    @user_id UNIQUEIDENTIFIER,
    @new_primary_org_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Validar que el usuario es miembro de la organización
        IF NOT EXISTS (
            SELECT 1 FROM organization_members
            WHERE user_id = @user_id
            AND organization_id = @new_primary_org_id
            AND left_at IS NULL
        )
        BEGIN
            THROW 50024, 'El usuario no es miembro de esta organización', 1;
        END
        
        -- Quitar primaria de todas
        UPDATE organization_members
        SET is_primary = 0,
            updated_at = GETUTCDATE()
        WHERE user_id = @user_id
        AND left_at IS NULL;
        
        -- Establecer nueva primaria
        UPDATE organization_members
        SET is_primary = 1,
            updated_at = GETUTCDATE()
        WHERE user_id = @user_id
        AND organization_id = @new_primary_org_id;
        
        COMMIT TRANSACTION;
        
        SELECT 1 AS success, 'Organización primaria cambiada exitosamente' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: Reactivar organización archivada
CREATE OR ALTER PROCEDURE sp_reactivate_organization
    @user_id UNIQUEIDENTIFIER,
    @organization_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Validar que el usuario es admin de la organización
        IF NOT EXISTS (
            SELECT 1 FROM organization_members
            WHERE user_id = @user_id
            AND organization_id = @organization_id
            AND role = 'admin'
            AND left_at IS NULL
        )
        BEGIN
            THROW 50025, 'El usuario no es administrador de esta organización', 1;
        END
        
        -- Reactivar organización
        UPDATE organizations
        SET is_archived = 0,
            archived_at = NULL,
            updated_at = GETUTCDATE()
        WHERE id = @organization_id;
        
        -- Si no tiene suscripción activa, crear una nueva con free_trial
        IF NOT EXISTS (
            SELECT 1 FROM subscriptions
            WHERE organization_id = @organization_id
            AND status IN ('active', 'trialing')
        )
        BEGIN
            DECLARE @subscription_id UNIQUEIDENTIFIER = NEWID();
            
            INSERT INTO subscriptions (
                id,
                organization_id,
                plan_id,
                status,
                billing_cycle,
                current_period_start,
                current_period_end,
                trial_start,
                trial_end
            )
            VALUES (
                @subscription_id,
                @organization_id,
                'free_trial',
                'trialing',
                NULL, -- free_trial no tiene billing_cycle (es gratuito)
                GETUTCDATE(),
                DATEADD(day, 30, GETUTCDATE()),
                GETUTCDATE(),
                DATEADD(day, 30, GETUTCDATE())
            );
            
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
            VALUES (
                @subscription_id,
                @organization_id,
                'free_trial',
                'trialing',
                'created',
                @user_id,
                '{"reason": "organization_reactivated"}' AS JSON
            );
        END
        
        COMMIT TRANSACTION;
        
        SELECT 1 AS success, 'Organización reactivada exitosamente' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: Crear código de invitación para organización
CREATE OR ALTER PROCEDURE sp_create_invitation_token
    @organization_id UNIQUEIDENTIFIER,
    @invited_by UNIQUEIDENTIFIER,
    @email VARCHAR(255) = NULL,
    @expires_in_days INT = 7
AS
BEGIN
    SET NOCOUNT ON;
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Validar que el invitador es admin de la organización
        IF NOT EXISTS (
            SELECT 1 FROM organization_members
            WHERE organization_id = @organization_id
            AND user_id = @invited_by
            AND role = 'admin'
            AND left_at IS NULL
        )
        BEGIN
            THROW 50026, 'Solo los administradores pueden crear invitaciones', 1;
        END
        
        -- Generar token único
        DECLARE @invitation_token VARCHAR(255) = 
            UPPER(SUBSTRING(CAST(NEWID() AS VARCHAR(36)), 1, 8) + '-' + 
                  SUBSTRING(CAST(NEWID() AS VARCHAR(36)), 1, 4) + '-' + 
                  SUBSTRING(CAST(NEWID() AS VARCHAR(36)), 1, 4));
        
        -- Crear registro temporal en organization_members para la invitación
        -- Si el email ya existe, actualizar el token
        IF EXISTS (
            SELECT 1 FROM organization_members
            WHERE organization_id = @organization_id
            AND invitation_token IS NOT NULL
            AND invitation_expires_at > GETUTCDATE()
            AND @email IS NOT NULL
        )
        BEGIN
            -- Buscar por usuario si existe
            DECLARE @existing_user_id UNIQUEIDENTIFIER;
            SELECT @existing_user_id = id FROM users WHERE email = @email;
            
            IF @existing_user_id IS NOT NULL AND NOT EXISTS (
                SELECT 1 FROM organization_members
                WHERE organization_id = @organization_id
                AND user_id = @existing_user_id
                AND left_at IS NULL
            )
            BEGIN
                -- Actualizar invitación existente
                UPDATE organization_members
                SET invitation_token = @invitation_token,
                    invitation_expires_at = DATEADD(day, @expires_in_days, GETUTCDATE()),
                    invited_by = @invited_by,
                    updated_at = GETUTCDATE()
                WHERE organization_id = @organization_id
                AND user_id = @existing_user_id
                AND left_at IS NULL;
            END
        END
        
        -- Si no existe registro, crear uno temporal (será activado cuando se acepte)
        -- Por ahora retornamos el token para que la aplicación lo maneje
        -- La aplicación debería guardar el token asociado al email
        
        COMMIT TRANSACTION;
        
        SELECT 
            1 AS success,
            @invitation_token AS invitation_token,
            DATEADD(day, @expires_in_days, GETUTCDATE()) AS expires_at,
            'Token de invitación creado exitosamente' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 
            0 AS success,
            NULL AS invitation_token,
            NULL AS expires_at,
            ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- ============================================================================
-- TRIGGERS ADICIONALES
-- ============================================================================

-- Trigger: Validar que solo haya una organización primaria por usuario
CREATE OR ALTER TRIGGER trg_validate_single_primary_organization
ON organization_members
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @user_id UNIQUEIDENTIFIER;
    DECLARE @is_primary BIT;
    
    SELECT @user_id = user_id, @is_primary = is_primary FROM inserted;
    
    IF @is_primary = 1
    BEGIN
        -- Verificar que no haya otra organización primaria para este usuario
        IF EXISTS (
            SELECT 1 FROM organization_members
            WHERE user_id = @user_id
            AND is_primary = 1
            AND left_at IS NULL
            AND id NOT IN (SELECT id FROM inserted)
        )
        BEGIN
            -- Auto-quitar primaria de las otras
            UPDATE organization_members
            SET is_primary = 0,
                updated_at = GETUTCDATE()
            WHERE user_id = @user_id
            AND is_primary = 1
            AND left_at IS NULL
            AND id NOT IN (SELECT id FROM inserted);
        END
    END
END;
GO

-- Trigger: Auto-archivar miembros cuando se archiva organización
-- No eliminar, solo marcar left_at para mantener historial
CREATE OR ALTER TRIGGER trg_organization_archive_members
ON organizations
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Si la organización fue archivada
    IF EXISTS (SELECT 1 FROM inserted WHERE is_archived = 1)
    AND EXISTS (SELECT 1 FROM deleted WHERE is_archived = 0)
    BEGIN
        DECLARE @org_id UNIQUEIDENTIFIER;
        SELECT @org_id = id FROM inserted WHERE is_archived = 1;
        
        -- Marcar left_at para todos los miembros (pero mantenerlos en la tabla para historial)
        UPDATE organization_members
        SET left_at = GETUTCDATE(),
            updated_at = GETUTCDATE()
        WHERE organization_id = @org_id
        AND left_at IS NULL;
    END
    
    -- Si la organización fue reactivada
    IF EXISTS (SELECT 1 FROM inserted WHERE is_archived = 0)
    AND EXISTS (SELECT 1 FROM deleted WHERE is_archived = 1)
    BEGIN
        DECLARE @org_id_reactivated UNIQUEIDENTIFIER;
        SELECT @org_id_reactivated = id FROM inserted WHERE is_archived = 0;
        
        -- Los miembros pueden ser reactivados manualmente si es necesario
        -- Por ahora, se mantienen con left_at para que el admin los reactive
    END
END;
GO

-- ============================================================================
-- VISTAS ADICIONALES
-- ============================================================================

-- Vista: Dashboard de organizaciones del usuario con toda la info necesaria para UI
CREATE OR ALTER VIEW vw_user_organizations_dashboard AS
SELECT 
    u.id AS user_id,
    u.email,
    o.id AS organization_id,
    o.name AS organization_name,
    o.slug,
    om.role,
    om.is_primary,
    o.is_archived,
    o.archived_at,
    s.plan_id,
    p.name AS plan_name,
    p.max_users,
    p.max_reports,
    s.status AS subscription_status,
    s.current_period_end,
    s.trial_end,
    (SELECT COUNT(*) FROM organization_members om2 WHERE om2.organization_id = o.id AND om2.left_at IS NULL) AS current_users,
    (SELECT COUNT(*) FROM reports r WHERE r.organization_id = o.id AND r.is_deleted = 0) AS current_reports,
    CASE 
        WHEN o.is_archived = 1 THEN 'Archivada'
        WHEN s.status = 'active' THEN 'Activa'
        WHEN s.status = 'trialing' THEN 'En prueba'
        WHEN s.status = 'past_due' THEN 'Pago pendiente'
        ELSE 'Inactiva'
    END AS status_label,
    CASE 
        WHEN o.is_archived = 1 THEN 0
        WHEN om.is_primary = 1 THEN 1
        ELSE 0
    END AS can_select,
    om.joined_at,
    o.created_at AS organization_created_at
FROM users u
INNER JOIN organization_members om ON om.user_id = u.id AND om.left_at IS NULL
INNER JOIN organizations o ON o.id = om.organization_id
LEFT JOIN subscriptions s ON s.organization_id = o.id AND s.status IN ('active', 'trialing')
LEFT JOIN plans p ON p.id = s.plan_id
WHERE u.is_active = 1;
GO

PRINT '✅ Procedimientos y funciones para flujo UX de organizaciones creados exitosamente';
GO


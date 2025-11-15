-- ============================================================================
-- REPORT TUNER - State Machine & Workflow Procedures
-- ============================================================================
-- Procedimientos almacenados, vistas y triggers para manejar el flujo completo
-- desde creación de usuario hasta gestión de suscripciones y límites
-- ============================================================================

USE empower_reports;
GO

-- ============================================================================
-- TRIGGERS PARA VALIDACIÓN Y AUDITORÍA
-- ============================================================================

-- Trigger: Auto-asignar plan free_trial al crear organización
-- Flujo: (2) JOIN_OR_CREATE_ORG → (3) ASSIGN_PLAN
CREATE OR ALTER TRIGGER trg_organization_auto_assign_free_trial
ON organizations
AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @organization_id UNIQUEIDENTIFIER;
    DECLARE @subscription_id UNIQUEIDENTIFIER;
    
    SELECT @organization_id = id FROM inserted;
    
    -- Crear suscripción automática al plan free_trial
    SET @subscription_id = NEWID();
    
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
        DATEADD(day, 30, GETUTCDATE()), -- 30 días de trial
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
        metadata
    )
    VALUES (
        @subscription_id,
        @organization_id,
        'free_trial',
        'trialing',
        'created',
        '{"auto_assigned": true, "trigger": "organization_created"}' AS JSON
    );
END;
GO

-- Trigger: Validar límites antes de insertar usuario en organización
-- Flujo: (4) USAGE_FLOW → (5) LIMIT_CHECK
CREATE OR ALTER TRIGGER trg_organization_members_check_user_limit
ON organization_members
AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @organization_id UNIQUEIDENTIFIER;
    DECLARE @max_users INT;
    DECLARE @current_users INT;
    
    SELECT @organization_id = organization_id FROM inserted;
    
    -- Obtener límite del plan activo
    SELECT @max_users = p.max_users
    FROM subscriptions s
    INNER JOIN plans p ON p.id = s.plan_id
    WHERE s.organization_id = @organization_id
    AND s.status IN ('active', 'trialing');
    
    -- Si no hay suscripción activa, no permitir
    IF @max_users IS NULL
    BEGIN
        ROLLBACK TRANSACTION;
        THROW 50001, 'La organización no tiene una suscripción activa', 1;
    END
    
    -- Contar usuarios actuales
    SELECT @current_users = COUNT(*)
    FROM organization_members
    WHERE organization_id = @organization_id
    AND left_at IS NULL;
    
    -- Validar límite
    IF @current_users > @max_users
    BEGIN
        ROLLBACK TRANSACTION;
        THROW 50002, 
            CONCAT('Límite de usuarios excedido. Máximo permitido: ', @max_users, '. Actual: ', @current_users),
            1;
    END
END;
GO

-- Trigger: Validar límites antes de insertar reporte
-- Flujo: (4) USAGE_FLOW → (5) LIMIT_CHECK
CREATE OR ALTER TRIGGER trg_reports_check_report_limit
ON reports
AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @organization_id UNIQUEIDENTIFIER;
    DECLARE @max_reports INT;
    DECLARE @current_reports INT;
    
    SELECT @organization_id = organization_id FROM inserted;
    
    -- Si es reporte individual (plan basic), validar límite del usuario
    IF @organization_id IS NULL
    BEGIN
        DECLARE @user_id UNIQUEIDENTIFIER;
        DECLARE @user_plan_id VARCHAR(50);
        
        SELECT @user_id = user_id FROM inserted;
        
        -- Obtener plan del usuario (buscar suscripción individual o de su org primaria)
        SELECT TOP 1 @user_plan_id = s.plan_id
        FROM users u
        LEFT JOIN organization_members om ON om.user_id = u.id AND om.is_primary = 1 AND om.left_at IS NULL
        LEFT JOIN subscriptions s ON s.organization_id = om.organization_id AND s.status IN ('active', 'trialing')
        WHERE u.id = @user_id;
        
        IF @user_plan_id IS NULL
            SET @user_plan_id = 'free_trial';
        
        SELECT @max_reports = max_reports FROM plans WHERE id = @user_plan_id;
        
        -- Contar reportes del usuario sin organización
        SELECT @current_reports = COUNT(*)
        FROM reports
        WHERE user_id = @user_id
        AND organization_id IS NULL
        AND is_deleted = 0;
    END
    ELSE
    BEGIN
        -- Obtener límite del plan de la organización
        SELECT @max_reports = p.max_reports
        FROM subscriptions s
        INNER JOIN plans p ON p.id = s.plan_id
        WHERE s.organization_id = @organization_id
        AND s.status IN ('active', 'trialing');
        
        IF @max_reports IS NULL
        BEGIN
            ROLLBACK TRANSACTION;
            THROW 50003, 'La organización no tiene una suscripción activa', 1;
        END
        
        -- Contar reportes de la organización
        SELECT @current_reports = COUNT(*)
        FROM reports
        WHERE organization_id = @organization_id
        AND is_deleted = 0;
    END
    
    -- Validar límite
    IF @current_reports > @max_reports
    BEGIN
        ROLLBACK TRANSACTION;
        THROW 50004, 
            CONCAT('Límite de reportes excedido. Máximo permitido: ', @max_reports, '. Actual: ', @current_reports),
            1;
    END
END;
GO

-- Trigger: Actualizar estado de suscripción cuando vence el período
-- Estados: Trialing → Active, Active → PastDue
CREATE OR ALTER TRIGGER trg_subscriptions_check_expiry
ON subscriptions
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Verificar si el período venció
    UPDATE s
    SET s.status = CASE 
        WHEN s.current_period_end < GETUTCDATE() AND s.status = 'active' 
        THEN 'past_due'
        WHEN s.trial_end < GETUTCDATE() AND s.status = 'trialing'
        THEN 'past_due'
        ELSE s.status
    END,
    s.updated_at = GETUTCDATE()
    FROM subscriptions s
    INNER JOIN inserted i ON i.id = s.id
    WHERE (s.current_period_end < GETUTCDATE() OR s.trial_end < GETUTCDATE())
    AND s.status IN ('active', 'trialing');
END;
GO

-- ============================================================================
-- PROCEDIMIENTOS PARA TRANSICIONES DE ESTADO DE SUSCRIPCIÓN
-- ============================================================================

-- Procedimiento: Trialing → Active (upgrade/checkout success)
-- Flujo: (3) ASSIGN_PLAN → Estado Active
CREATE OR ALTER PROCEDURE sp_subscription_activate
    @subscription_id UNIQUEIDENTIFIER,
    @stripe_subscription_id VARCHAR(255) = NULL,
    @stripe_price_id VARCHAR(255) = NULL,
    @period_start DATETIME2 = NULL,
    @period_end DATETIME2 = NULL,
    @activated_by UNIQUEIDENTIFIER = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @organization_id UNIQUEIDENTIFIER;
    DECLARE @old_status VARCHAR(50);
    DECLARE @plan_id VARCHAR(50);
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Obtener datos actuales
        SELECT 
            @organization_id = organization_id,
            @old_status = status,
            @plan_id = plan_id
        FROM subscriptions
        WHERE id = @subscription_id;
        
        IF @organization_id IS NULL
        BEGIN
            THROW 50005, 'Suscripción no encontrada', 1;
        END
        
        -- Actualizar suscripción a Active
        UPDATE subscriptions
        SET status = 'active',
            stripe_subscription_id = ISNULL(@stripe_subscription_id, stripe_subscription_id),
            stripe_price_id = ISNULL(@stripe_price_id, stripe_price_id),
            current_period_start = ISNULL(@period_start, current_period_start),
            current_period_end = ISNULL(@period_end, current_period_end),
            trial_end = CASE WHEN status = 'trialing' THEN GETUTCDATE() ELSE trial_end END,
            updated_at = GETUTCDATE()
        WHERE id = @subscription_id;
        
        -- Registrar en historial
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
            @plan_id,
            @plan_id,
            @old_status,
            'active',
            'updated',
            @activated_by,
            '{"stripe_subscription_id": "' + ISNULL(@stripe_subscription_id, '') + '", "transition": "trialing_to_active"}' AS JSON
        );
        
        COMMIT TRANSACTION;
        
        SELECT 1 AS success, 'Suscripción activada exitosamente' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: Active → Canceled (cancel)
-- Flujo: Cancelación de suscripción
CREATE OR ALTER PROCEDURE sp_subscription_cancel
    @subscription_id UNIQUEIDENTIFIER,
    @cancel_at_period_end BIT = 1,
    @canceled_by UNIQUEIDENTIFIER = NULL,
    @reason TEXT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @organization_id UNIQUEIDENTIFIER;
    DECLARE @old_status VARCHAR(50);
    DECLARE @plan_id VARCHAR(50);
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Obtener datos actuales
        SELECT 
            @organization_id = organization_id,
            @old_status = status,
            @plan_id = plan_id
        FROM subscriptions
        WHERE id = @subscription_id;
        
        IF @organization_id IS NULL
        BEGIN
            THROW 50006, 'Suscripción no encontrada', 1;
        END
        
        IF @cancel_at_period_end = 1
        BEGIN
            -- Cancelar al final del período
            UPDATE subscriptions
            SET cancel_at_period_end = 1,
                updated_at = GETUTCDATE()
            WHERE id = @subscription_id;
        END
        ELSE
        BEGIN
            -- Cancelar inmediatamente
            UPDATE subscriptions
            SET status = 'canceled',
                cancel_at_period_end = 0,
                canceled_at = GETUTCDATE(),
                updated_at = GETUTCDATE()
            WHERE id = @subscription_id;
        END
        
        -- Registrar en historial
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
            @plan_id,
            @plan_id,
            @old_status,
            CASE WHEN @cancel_at_period_end = 1 THEN @old_status ELSE 'canceled' END,
            'canceled',
            @canceled_by,
            '{"cancel_at_period_end": ' + CAST(@cancel_at_period_end AS VARCHAR(1)) + ', "reason": "' + ISNULL(@reason, '') + '"}' AS JSON
        );
        
        COMMIT TRANSACTION;
        
        SELECT 1 AS success, 
            CASE WHEN @cancel_at_period_end = 1 
                THEN 'Cancelación programada al final del período' 
                ELSE 'Suscripción cancelada exitosamente' 
            END AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: Active → PastDue (payment failed)
-- Flujo: Cuando falla el pago en Stripe
CREATE OR ALTER PROCEDURE sp_subscription_mark_past_due
    @subscription_id UNIQUEIDENTIFIER,
    @stripe_event_id VARCHAR(255) = NULL,
    @metadata JSON = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @organization_id UNIQUEIDENTIFIER;
    DECLARE @old_status VARCHAR(50);
    DECLARE @plan_id VARCHAR(50);
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Obtener datos actuales
        SELECT 
            @organization_id = organization_id,
            @old_status = status,
            @plan_id = plan_id
        FROM subscriptions
        WHERE id = @subscription_id;
        
        IF @organization_id IS NULL
        BEGIN
            THROW 50007, 'Suscripción no encontrada', 1;
        END
        
        -- Actualizar a PastDue
        UPDATE subscriptions
        SET status = 'past_due',
            updated_at = GETUTCDATE()
        WHERE id = @subscription_id;
        
        -- Registrar en historial
        INSERT INTO subscription_history (
            subscription_id,
            organization_id,
            plan_id_old,
            plan_id_new,
            status_old,
            status_new,
            event_type,
            stripe_event_id,
            metadata
        )
        VALUES (
            @subscription_id,
            @organization_id,
            @plan_id,
            @plan_id,
            @old_status,
            'past_due',
            'stripe_webhook',
            @stripe_event_id,
            ISNULL(@metadata, '{"reason": "payment_failed"}' AS JSON)
        );
        
        COMMIT TRANSACTION;
        
        SELECT 1 AS success, 'Suscripción marcada como PastDue' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: PastDue → Active (payment resolved)
-- Flujo: Cuando se resuelve el pago pendiente
CREATE OR ALTER PROCEDURE sp_subscription_resolve_past_due
    @subscription_id UNIQUEIDENTIFIER,
    @stripe_event_id VARCHAR(255) = NULL,
    @new_period_end DATETIME2 = NULL,
    @metadata JSON = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @organization_id UNIQUEIDENTIFIER;
    DECLARE @old_status VARCHAR(50);
    DECLARE @plan_id VARCHAR(50);
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Obtener datos actuales
        SELECT 
            @organization_id = organization_id,
            @old_status = status,
            @plan_id = plan_id
        FROM subscriptions
        WHERE id = @subscription_id;
        
        IF @organization_id IS NULL
        BEGIN
            THROW 50008, 'Suscripción no encontrada', 1;
        END
        
        IF @old_status != 'past_due'
        BEGIN
            THROW 50009, 'La suscripción no está en estado PastDue', 1;
        END
        
        -- Actualizar a Active y extender período si es necesario
        UPDATE subscriptions
        SET status = 'active',
            current_period_end = ISNULL(@new_period_end, current_period_end),
            updated_at = GETUTCDATE()
        WHERE id = @subscription_id;
        
        -- Registrar en historial
        INSERT INTO subscription_history (
            subscription_id,
            organization_id,
            plan_id_old,
            plan_id_new,
            status_old,
            status_new,
            event_type,
            stripe_event_id,
            metadata
        )
        VALUES (
            @subscription_id,
            @organization_id,
            @plan_id,
            @plan_id,
            @old_status,
            'active',
            'stripe_webhook',
            @stripe_event_id,
            ISNULL(@metadata, '{"reason": "payment_resolved"}' AS JSON)
        );
        
        COMMIT TRANSACTION;
        
        SELECT 1 AS success, 'Pago resuelto, suscripción reactivada' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: Manejar cancelación al final del período
-- Flujo: Canceled → [*] (finalizar suscripción)
CREATE OR ALTER PROCEDURE sp_subscription_finalize_cancellation
    @subscription_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @organization_id UNIQUEIDENTIFIER;
    DECLARE @plan_id VARCHAR(50);
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Obtener datos actuales
        SELECT 
            @organization_id = organization_id,
            @plan_id = plan_id
        FROM subscriptions
        WHERE id = @subscription_id
        AND cancel_at_period_end = 1
        AND current_period_end < GETUTCDATE();
        
        IF @organization_id IS NULL
        BEGIN
            SELECT 0 AS success, 'La suscripción no está programada para cancelarse o el período aún no ha terminado' AS message;
            RETURN;
        END
        
        -- Finalizar cancelación
        UPDATE subscriptions
        SET status = 'canceled',
            canceled_at = GETUTCDATE(),
            cancel_at_period_end = 0,
            updated_at = GETUTCDATE()
        WHERE id = @subscription_id;
        
        -- Registrar en historial
        INSERT INTO subscription_history (
            subscription_id,
            organization_id,
            plan_id_old,
            plan_id_new,
            status_new,
            event_type,
            metadata
        )
        VALUES (
            @subscription_id,
            @organization_id,
            @plan_id,
            @plan_id,
            'canceled',
            'canceled',
            '{"finalized": true, "reason": "period_end"}' AS JSON
        );
        
        COMMIT TRANSACTION;
        
        SELECT 1 AS success, 'Cancelación finalizada' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- ============================================================================
-- PROCEDIMIENTOS PARA EL FLUJO PRINCIPAL
-- ============================================================================

-- Procedimiento: (1) USER_CREATED - Crear usuario después de OAuth
CREATE OR ALTER PROCEDURE sp_create_user
    @email VARCHAR(255),
    @name VARCHAR(255),
    @auth_provider VARCHAR(50),
    @auth_provider_id VARCHAR(255),
    @avatar_url VARCHAR(500) = NULL,
    @metadata JSON = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @user_id UNIQUEIDENTIFIER = NEWID();
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        -- Verificar si el usuario ya existe
        IF EXISTS (SELECT 1 FROM users WHERE email = @email OR (auth_provider = @auth_provider AND auth_provider_id = @auth_provider_id))
        BEGIN
            SELECT id, email, name, is_active
            FROM users
            WHERE email = @email OR (auth_provider = @auth_provider AND auth_provider_id = @auth_provider_id);
            RETURN;
        END
        
        -- Crear usuario
        INSERT INTO users (
            id,
            email,
            name,
            avatar_url,
            auth_provider,
            auth_provider_id,
            is_email_verified,
            last_login_at,
            metadata
        )
        VALUES (
            @user_id,
            @email,
            @name,
            @avatar_url,
            @auth_provider,
            @auth_provider_id,
            CASE WHEN @auth_provider != 'email' THEN 1 ELSE 0 END,
            GETUTCDATE(),
            @metadata
        );
        
        COMMIT TRANSACTION;
        
        SELECT @user_id AS user_id, 'Usuario creado exitosamente' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT NULL AS user_id, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: (2) JOIN_OR_CREATE_ORG - Crear o unirse a organización
CREATE OR ALTER PROCEDURE sp_create_or_join_organization
    @user_id UNIQUEIDENTIFIER,
    @organization_name VARCHAR(255),
    @organization_id UNIQUEIDENTIFIER = NULL, -- Si se proporciona, unirse a esa org
    @invitation_token VARCHAR(255) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @new_organization_id UNIQUEIDENTIFIER;
    DECLARE @existing_org_id UNIQUEIDENTIFIER;
    DECLARE @role VARCHAR(50) = 'admin';
    
    BEGIN TRANSACTION;
    
    BEGIN TRY
        IF @organization_id IS NOT NULL
        BEGIN
            -- Unirse a organización existente
            SET @new_organization_id = @organization_id;
            
            -- Validar invitación si hay token
            IF @invitation_token IS NOT NULL
            BEGIN
                SELECT @existing_org_id = organization_id
                FROM organization_members
                WHERE invitation_token = @invitation_token
                AND invitation_expires_at > GETUTCDATE();
                
                IF @existing_org_id IS NULL
                BEGIN
                    THROW 50010, 'Token de invitación inválido o expirado', 1;
                END
                
                SET @new_organization_id = @existing_org_id;
                SET @role = 'member';
            END
            
            -- Verificar si el usuario ya es miembro
            IF EXISTS (SELECT 1 FROM organization_members WHERE organization_id = @new_organization_id AND user_id = @user_id AND left_at IS NULL)
            BEGIN
                SELECT @new_organization_id AS organization_id, 'Usuario ya es miembro de esta organización' AS message;
                COMMIT TRANSACTION;
                RETURN;
            END
        END
        ELSE
        BEGIN
            -- Crear nueva organización
            SET @new_organization_id = NEWID();
            
            INSERT INTO organizations (
                id,
                name,
                slug,
                is_archived
            )
            VALUES (
                @new_organization_id,
                @organization_name,
                LOWER(REPLACE(REPLACE(REPLACE(@organization_name, ' ', '-'), '''', ''), '.', '')),
                0
            );
            
            -- El trigger trg_organization_auto_assign_free_trial asignará automáticamente el plan free_trial
        END
        
        -- Agregar usuario a la organización
        -- Si es el primer miembro o el creador, será admin y primaria
        DECLARE @is_first_member BIT;
        SELECT @is_first_member = CASE WHEN COUNT(*) = 0 THEN 1 ELSE 0 END
        FROM organization_members
        WHERE organization_id = @new_organization_id AND left_at IS NULL;
        
        -- Si el usuario ya tiene una organización primaria, no hacer esta primaria automáticamente
        DECLARE @has_primary BIT;
        SELECT @has_primary = CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END
        FROM organization_members
        WHERE user_id = @user_id AND is_primary = 1 AND left_at IS NULL;
        
        INSERT INTO organization_members (
            organization_id,
            user_id,
            role,
            is_primary,
            joined_at
        )
        VALUES (
            @new_organization_id,
            @user_id,
            CASE WHEN @is_first_member = 1 THEN 'admin' ELSE @role END,
            CASE WHEN @has_primary = 0 THEN 1 ELSE 0 END,
            GETUTCDATE()
        );
        
        COMMIT TRANSACTION;
        
        SELECT @new_organization_id AS organization_id, 'Organización creada/unida exitosamente' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT NULL AS organization_id, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- Procedimiento: (5) LIMIT_CHECK - Verificar y notificar límites alcanzados
CREATE OR ALTER PROCEDURE sp_check_organization_limits
    @organization_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    SELECT 
        o.name AS organization_name,
        s.plan_id,
        p.name AS plan_name,
        s.status AS subscription_status,
        
        -- Usuarios
        (SELECT COUNT(*) FROM organization_members om WHERE om.organization_id = @organization_id AND om.left_at IS NULL) AS current_users,
        p.max_users AS max_users,
        CASE 
            WHEN (SELECT COUNT(*) FROM organization_members om WHERE om.organization_id = @organization_id AND om.left_at IS NULL) >= p.max_users 
            THEN 1 ELSE 0 
        END AS users_limit_reached,
        CAST((SELECT COUNT(*) FROM organization_members om WHERE om.organization_id = @organization_id AND om.left_at IS NULL) AS FLOAT) / p.max_users * 100 AS users_usage_percent,
        
        -- Reportes
        (SELECT COUNT(*) FROM reports r WHERE r.organization_id = @organization_id AND r.is_deleted = 0) AS current_reports,
        p.max_reports AS max_reports,
        CASE 
            WHEN (SELECT COUNT(*) FROM reports r WHERE r.organization_id = @organization_id AND r.is_deleted = 0) >= p.max_reports 
            THEN 1 ELSE 0 
        END AS reports_limit_reached,
        CAST((SELECT COUNT(*) FROM reports r WHERE r.organization_id = @organization_id AND r.is_deleted = 0) AS FLOAT) / p.max_reports * 100 AS reports_usage_percent,
        
        -- Almacenamiento (si implementado)
        (SELECT ISNULL(SUM(file_size_bytes), 0) / 1024.0 / 1024.0 FROM reports r WHERE r.organization_id = @organization_id AND r.is_deleted = 0) AS current_storage_mb,
        p.max_storage_mb AS max_storage_mb,
        CASE 
            WHEN (SELECT ISNULL(SUM(file_size_bytes), 0) FROM reports r WHERE r.organization_id = @organization_id AND r.is_deleted = 0) / 1024.0 / 1024.0 >= p.max_storage_mb 
            THEN 1 ELSE 0 
        END AS storage_limit_reached
        
    FROM organizations o
    INNER JOIN subscriptions s ON s.organization_id = o.id AND s.status IN ('active', 'trialing')
    INNER JOIN plans p ON p.id = s.plan_id
    WHERE o.id = @organization_id;
END;
GO

-- Procedimiento: (6) PLAN_UPDATE - Upgrade/Downgrade de plan con validación
CREATE OR ALTER PROCEDURE sp_update_subscription_plan
    @subscription_id UNIQUEIDENTIFIER,
    @new_plan_id VARCHAR(50),
    @billing_cycle VARCHAR(20) = NULL, -- NULL para free_trial, 'monthly' o 'yearly' para otros
    @stripe_price_id VARCHAR(255) = NULL,
    @stripe_subscription_id VARCHAR(255) = NULL,
    @updated_by UNIQUEIDENTIFIER = NULL,
    @force_downgrade BIT = 0 -- Si es 1, permite downgrade incluso si excede límites
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @organization_id UNIQUEIDENTIFIER;
    DECLARE @old_plan_id VARCHAR(50);
    DECLARE @new_max_users INT;
    DECLARE @new_max_reports INT;
    DECLARE @current_users INT;
    DECLARE @current_reports INT;
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
        
        -- Obtener datos actuales
        SELECT 
            @organization_id = organization_id,
            @old_plan_id = plan_id
        FROM subscriptions
        WHERE id = @subscription_id;
        
        IF @organization_id IS NULL
        BEGIN
            THROW 50011, 'Suscripción no encontrada', 1;
        END
        
        -- Obtener límites del nuevo plan
        SELECT 
            @new_max_users = max_users,
            @new_max_reports = max_reports
        FROM plans
        WHERE id = @new_plan_id;
        
        IF @new_max_users IS NULL
        BEGIN
            THROW 50012, 'Plan no encontrado', 1;
        END
        
        -- Contar uso actual
        SELECT @current_users = COUNT(*)
        FROM organization_members
        WHERE organization_id = @organization_id AND left_at IS NULL;
        
        SELECT @current_reports = COUNT(*)
        FROM reports
        WHERE organization_id = @organization_id AND is_deleted = 0;
        
        -- Validar límites (a menos que sea downgrade forzado)
        IF @force_downgrade = 0
        BEGIN
            IF @current_users > @new_max_users
            BEGIN
                THROW 50013, 
                    CONCAT('No se puede cambiar al plan ', @new_plan_id, '. La organización tiene ', @current_users, ' usuarios pero el plan permite máximo ', @new_max_users),
                    1;
            END
            
            IF @current_reports > @new_max_reports
            BEGIN
                THROW 50014, 
                    CONCAT('No se puede cambiar al plan ', @new_plan_id, '. La organización tiene ', @current_reports, ' reportes pero el plan permite máximo ', @new_max_reports),
                    1;
            END
        END
        
        -- Actualizar suscripción
        UPDATE subscriptions
        SET plan_id = @new_plan_id,
            billing_cycle = @final_billing_cycle,
            stripe_price_id = ISNULL(@stripe_price_id, stripe_price_id),
            stripe_subscription_id = ISNULL(@stripe_subscription_id, stripe_subscription_id),
            status = CASE WHEN @new_plan_id = 'free_trial' AND status = 'active' THEN 'trialing' ELSE status END,
            updated_at = GETUTCDATE()
        WHERE id = @subscription_id;
        
        -- Registrar en historial
        DECLARE @metadata JSON = '{"billing_cycle": ' + 
            CASE WHEN @final_billing_cycle IS NULL THEN 'null' ELSE '"' + @final_billing_cycle + '"' END + 
            ', "force_downgrade": ' + CAST(@force_downgrade AS VARCHAR(1)) + '}' AS JSON;
        
        INSERT INTO subscription_history (
            subscription_id,
            organization_id,
            plan_id_old,
            plan_id_new,
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
            (SELECT status FROM subscriptions WHERE id = @subscription_id),
            'plan_changed',
            @updated_by,
            @metadata
        );
        
        COMMIT TRANSACTION;
        
        SELECT 1 AS success, 'Plan actualizado exitosamente' AS message;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        SELECT 0 AS success, ERROR_MESSAGE() AS message;
    END CATCH
END;
GO

-- ============================================================================
-- VISTAS PARA MONITOREO Y DASHBOARD
-- ============================================================================

-- Vista: Estado actual de todas las organizaciones con sus límites y uso
CREATE OR ALTER VIEW vw_organizations_usage_status AS
SELECT 
    o.id AS organization_id,
    o.name AS organization_name,
    o.slug,
    s.id AS subscription_id,
    s.status AS subscription_status,
    p.id AS plan_id,
    p.name AS plan_name,
    s.current_period_end,
    s.trial_end,
    
    -- Conteos actuales
    (SELECT COUNT(*) FROM organization_members om WHERE om.organization_id = o.id AND om.left_at IS NULL) AS current_users,
    (SELECT COUNT(*) FROM reports r WHERE r.organization_id = o.id AND r.is_deleted = 0) AS current_reports,
    
    -- Límites
    p.max_users,
    p.max_reports,
    
    -- Porcentajes de uso
    CAST((SELECT COUNT(*) FROM organization_members om WHERE om.organization_id = o.id AND om.left_at IS NULL) AS FLOAT) / NULLIF(p.max_users, 0) * 100 AS users_usage_percent,
    CAST((SELECT COUNT(*) FROM reports r WHERE r.organization_id = o.id AND r.is_deleted = 0) AS FLOAT) / NULLIF(p.max_reports, 0) * 100 AS reports_usage_percent,
    
    -- Flags de límite alcanzado
    CASE WHEN (SELECT COUNT(*) FROM organization_members om WHERE om.organization_id = o.id AND om.left_at IS NULL) >= p.max_users THEN 1 ELSE 0 END AS users_limit_reached,
    CASE WHEN (SELECT COUNT(*) FROM reports r WHERE r.organization_id = o.id AND r.is_deleted = 0) >= p.max_reports THEN 1 ELSE 0 END AS reports_limit_reached,
    
    o.created_at,
    o.is_archived
FROM organizations o
LEFT JOIN subscriptions s ON s.organization_id = o.id AND s.status IN ('active', 'trialing')
LEFT JOIN plans p ON p.id = s.plan_id
WHERE o.is_archived = 0;
GO

-- Vista: Suscripciones que requieren atención
CREATE OR ALTER VIEW vw_subscriptions_requiring_attention AS
SELECT 
    s.id AS subscription_id,
    o.name AS organization_name,
    s.plan_id,
    p.name AS plan_name,
    s.status,
    s.current_period_end,
    s.trial_end,
    DATEDIFF(day, GETUTCDATE(), s.current_period_end) AS days_until_expiry,
    DATEDIFF(day, GETUTCDATE(), s.trial_end) AS days_until_trial_end,
    CASE 
        WHEN s.status = 'past_due' THEN 'Pago pendiente'
        WHEN s.status = 'trialing' AND s.trial_end < DATEADD(day, 3, GETUTCDATE()) THEN 'Trial por vencer'
        WHEN s.cancel_at_period_end = 1 AND s.current_period_end < DATEADD(day, 7, GETUTCDATE()) THEN 'Cancelación próxima'
        WHEN s.current_period_end < DATEADD(day, 7, GETUTCDATE()) THEN 'Renovación próxima'
        ELSE NULL
    END AS attention_reason
FROM subscriptions s
INNER JOIN organizations o ON o.id = s.organization_id
INNER JOIN plans p ON p.id = s.plan_id
WHERE s.status IN ('active', 'trialing', 'past_due')
AND (
    s.status = 'past_due'
    OR (s.status = 'trialing' AND s.trial_end < DATEADD(day, 3, GETUTCDATE()))
    OR (s.cancel_at_period_end = 1 AND s.current_period_end < DATEADD(day, 7, GETUTCDATE()))
    OR (s.current_period_end < DATEADD(day, 7, GETUTCDATE()))
)
ORDER BY 
    CASE s.status
        WHEN 'past_due' THEN 1
        WHEN 'trialing' THEN 2
        ELSE 3
    END,
    s.current_period_end;
GO

-- Vista: Historial de cambios de estado de suscripciones
CREATE OR ALTER VIEW vw_subscription_state_history AS
SELECT 
    sh.id,
    sh.subscription_id,
    o.name AS organization_name,
    p_old.name AS plan_from,
    p_new.name AS plan_to,
    sh.status_old,
    sh.status_new,
    sh.event_type,
    u.name AS changed_by_user,
    sh.stripe_event_id,
    sh.created_at,
    DATEDIFF(minute, LAG(sh.created_at) OVER (PARTITION BY sh.subscription_id ORDER BY sh.created_at), sh.created_at) AS minutes_since_last_change
FROM subscription_history sh
INNER JOIN organizations o ON o.id = sh.organization_id
LEFT JOIN plans p_old ON p_old.id = sh.plan_id_old
LEFT JOIN plans p_new ON p_new.id = sh.plan_id_new
LEFT JOIN users u ON u.id = sh.changed_by
ORDER BY sh.subscription_id, sh.created_at DESC;
GO

PRINT '✅ Procedimientos, triggers y vistas de flujo y estado creados exitosamente';
GO


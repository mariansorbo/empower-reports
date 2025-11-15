-- ============================================================================
-- REPORT TUNER - Documentation Procedures
-- ============================================================================
-- Stored procedures para gestionar documentación de organizaciones
-- ============================================================================

USE empower_reports;
GO

-- ============================================================================
-- PROCEDURE: Establecer o actualizar URL de documentación
-- ============================================================================

CREATE OR ALTER PROCEDURE sp_set_organization_documentation
    @organization_id UNIQUEIDENTIFIER,
    @documentation_url VARCHAR(500),
    @description TEXT = NULL,
    @created_by UNIQUEIDENTIFIER = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Validar que la organización existe y no está archivada
    IF NOT EXISTS (
        SELECT 1 FROM organizations 
        WHERE id = @organization_id 
        AND is_archived = 0
    )
    BEGIN
        SELECT 0 AS success, 'Organización no existe o está archivada' AS message;
        RETURN;
    END;
    
    -- Validar que el usuario pertenece a la organización (si se proporciona)
    IF @created_by IS NOT NULL
    BEGIN
        IF NOT EXISTS (
            SELECT 1 FROM organization_members
            WHERE organization_id = @organization_id
            AND user_id = @created_by
            AND role IN ('admin', 'admin_global')
            AND left_at IS NULL
        )
        BEGIN
            SELECT 0 AS success, 'Usuario no es admin de la organización' AS message;
            RETURN;
        END;
    END;
    
    -- Verificar si ya existe registro
    IF EXISTS (SELECT 1 FROM organization_documentation WHERE organization_id = @organization_id)
    BEGIN
        -- Actualizar existente
        UPDATE organization_documentation
        SET 
            documentation_url = @documentation_url,
            description = ISNULL(@description, description),
            is_active = 1,
            updated_at = GETUTCDATE()
        WHERE organization_id = @organization_id;
        
        SELECT 1 AS success, 'URL de documentación actualizada' AS message;
    END
    ELSE
    BEGIN
        -- Insertar nueva
        INSERT INTO organization_documentation (
            organization_id,
            documentation_url,
            description,
            created_by
        )
        VALUES (
            @organization_id,
            @documentation_url,
            @description,
            @created_by
        );
        
        SELECT 1 AS success, 'URL de documentación configurada' AS message;
    END;
END;
GO

-- ============================================================================
-- PROCEDURE: Desactivar URL de documentación
-- ============================================================================

CREATE OR ALTER PROCEDURE sp_disable_organization_documentation
    @organization_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE organization_documentation
    SET 
        is_active = 0,
        updated_at = GETUTCDATE()
    WHERE organization_id = @organization_id;
    
    IF @@ROWCOUNT > 0
        SELECT 1 AS success, 'Documentación desactivada' AS message;
    ELSE
        SELECT 0 AS success, 'No se encontró documentación para esta organización' AS message;
END;
GO

-- ============================================================================
-- PROCEDURE: Eliminar URL de documentación
-- ============================================================================

CREATE OR ALTER PROCEDURE sp_remove_organization_documentation
    @organization_id UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    
    DELETE FROM organization_documentation
    WHERE organization_id = @organization_id;
    
    IF @@ROWCOUNT > 0
        SELECT 1 AS success, 'Documentación eliminada' AS message;
    ELSE
        SELECT 0 AS success, 'No se encontró documentación para esta organización' AS message;
END;
GO

PRINT 'Procedures de documentacion creados exitosamente';
PRINT 'Procedures: sp_set_organization_documentation, sp_disable_organization_documentation, sp_remove_organization_documentation';
GO







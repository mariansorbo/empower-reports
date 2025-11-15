#!/bin/bash

# Script de inicialización de la base de datos SQL Server
# Este script inicia SQL Server y luego ejecuta los scripts de inicialización

echo "Iniciando SQL Server..."

# Iniciar SQL Server en segundo plano
/opt/mssql/bin/sqlservr &

# Esperar a que SQL Server esté listo
echo "Esperando a que SQL Server esté listo..."
sleep 30s

# Intentar conectarse hasta que esté disponible
until /opt/mssql-tools/bin/sqlcmd -S localhost -U sa -P ${MSSQL_SA_PASSWORD} -Q "SELECT 1" &> /dev/null
do
  echo "SQL Server aún no está listo, esperando..."
  sleep 5s
done

echo "SQL Server está listo. Ejecutando scripts de inicialización..."

# Ejecutar scripts en orden
echo "1. Creando esquema base..."
/opt/mssql-tools/bin/sqlcmd -S localhost -U sa -P ${MSSQL_SA_PASSWORD} -i /usr/src/app/schema.sql

echo "2. Creando workflows de organizaciones..."
/opt/mssql-tools/bin/sqlcmd -S localhost -U sa -P ${MSSQL_SA_PASSWORD} -i /usr/src/app/organization_workflows.sql

echo "3. Creando state machines y workflows..."
/opt/mssql-tools/bin/sqlcmd -S localhost -U sa -P ${MSSQL_SA_PASSWORD} -i /usr/src/app/state_machine_and_workflows.sql

echo "4. Aplicando constraints y validaciones..."
/opt/mssql-tools/bin/sqlcmd -S localhost -U sa -P ${MSSQL_SA_PASSWORD} -i /usr/src/app/constraints_and_validations.sql

echo "5. Aplicando migraciones..."
if [ -f "/usr/src/app/migrations/001_fix_billing_cycle_and_organization_null.sql" ]; then
    /opt/mssql-tools/bin/sqlcmd -S localhost -U sa -P ${MSSQL_SA_PASSWORD} -i /usr/src/app/migrations/001_fix_billing_cycle_and_organization_null.sql
fi

echo "¡Base de datos inicializada correctamente!"
echo "================================"
echo "Servidor: localhost,1433"
echo "Usuario: sa"
echo "Contraseña: ${MSSQL_SA_PASSWORD}"
echo "Base de datos: empower_reports"
echo "================================"

# Mantener el contenedor corriendo
wait







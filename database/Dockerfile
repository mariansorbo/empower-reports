# Dockerfile para SQL Server con Empower Reports Database
FROM mcr.microsoft.com/mssql/server:2022-latest

# Variables de entorno
ENV ACCEPT_EULA=Y
ENV MSSQL_SA_PASSWORD=YourStrong!Passw0rd
ENV MSSQL_PID=Developer

# Crear directorio para scripts
USER root
RUN mkdir -p /usr/src/app
WORKDIR /usr/src/app

# Copiar scripts SQL
COPY schema.sql ./
COPY organization_workflows.sql ./
COPY state_machine_and_workflows.sql ./
COPY constraints_and_validations.sql ./
COPY migrations/ ./migrations/
COPY init-db.sh ./

# Dar permisos de ejecución
RUN chmod +x init-db.sh

# Exponer puerto SQL Server
EXPOSE 1433

# Punto de entrada
ENTRYPOINT ["/usr/src/app/init-db.sh"]


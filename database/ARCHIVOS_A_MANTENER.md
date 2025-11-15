# Archivos a Mantener en /database

## ✅ Archivos NECESARIOS

### SQL (6 archivos)
- ✅ **schema.sql** - Schema principal (OBLIGATORIO)
- ✅ **organization_workflows.sql** - Workflows de organizaciones (OBLIGATORIO)
- ✅ **state_machine_and_workflows.sql** - Máquina de estados (OBLIGATORIO)
- ✅ **constraints_and_validations.sql** - Validaciones (OBLIGATORIO)
- ✅ **enterprise_pro_plan_v2.sql** - Enterprise Pro (OPCIONAL)
- ✅ **useful_queries.sql** - Queries útiles (OPCIONAL)

### SQL Referencia (1 archivo)
- ✅ **tables_only.sql** - Solo tablas (para referencia, no ejecutar)

### Documentación (9 archivos)
- ✅ **README.md** - Guía principal
- ✅ **INSTALLATION_ORDER.md** - Orden de ejecución
- ✅ **FLUJOS_COMPLETOS.md** - Flujos con referencias
- ✅ **TRIGGERS_PROCEDURES_FUNCTIONS.md** - Lista completa
- ✅ **ARCHITECTURE_SIMPLE.md** - Filosofía del diseño
- ✅ **SCHEMA_OVERVIEW.md** - Resumen de cambios
- ✅ **ENTERPRISE_PRO_V2_README.md** - Enterprise Pro
- ✅ **SAAS_TOOLS_AND_SYSTEMS.md** - Herramientas externas
- ✅ **DIAGRAM_PROMPT.md** - Para generar UML

### Excel (1 archivo)
- ✅ **DATABASE_SIMPLE.xlsx** - Datos dummy (MANTENER)

### Migrations (1 carpeta)
- ✅ **migrations/** - Carpeta con migraciones

---

## ❌ Archivos OBSOLETOS (Eliminar)

Los siguientes archivos están obsoletos y pueden eliminarse:

- ❌ **EMPOWER_REPORTS_DATABASE_SCHEMA.xlsx** (Excel viejo - cerrar y eliminar)
- ❌ **EMPOWER_REPORTS_SCHEMA.xlsx** (Excel viejo - cerrar y eliminar)

**NOTA**: Estos archivos no se pudieron eliminar porque están abiertos en Excel. Ciérralos y elimínalos manualmente.

---

## 📊 Total de Archivos

**Total a mantener**: 17 archivos + 1 carpeta
- 7 archivos SQL (6 para ejecutar + 1 referencia)
- 9 archivos de documentación (.md)
- 1 archivo Excel (DATABASE_SIMPLE.xlsx)
- 1 carpeta migrations/

**Total obsoleto**: 2 archivos Excel viejos (eliminar después de cerrarlos)

---

## 🗂️ Organización Final

```
database/
├── 📄 SQL PARA EJECUTAR (orden de instalación)
│   ├── schema.sql                              ⬅ 1. Ejecutar primero
│   ├── organization_workflows.sql              ⬅ 2. Ejecutar segundo
│   ├── state_machine_and_workflows.sql         ⬅ 3. Ejecutar tercero
│   ├── constraints_and_validations.sql         ⬅ 4. Ejecutar cuarto
│   ├── enterprise_pro_plan_v2.sql              ⬅ 5. OPCIONAL
│   └── useful_queries.sql                      ⬅ 6. OPCIONAL
│
├── 📄 SQL REFERENCIA (no ejecutar)
│   └── tables_only.sql                         ⬅ Solo para consulta
│
├── 📋 DOCUMENTACIÓN
│   ├── README.md                               ⬅ Guía principal
│   ├── INSTALLATION_ORDER.md                   ⬅ Cómo instalar
│   ├── FLUJOS_COMPLETOS.md                     ⬅ Flujos paso a paso
│   ├── TRIGGERS_PROCEDURES_FUNCTIONS.md        ⬅ Lista completa
│   ├── ARCHITECTURE_SIMPLE.md                  ⬅ Filosofía
│   ├── SCHEMA_OVERVIEW.md                      ⬅ Resumen
│   ├── ENTERPRISE_PRO_V2_README.md             ⬅ Enterprise Pro
│   ├── DIAGRAM_PROMPT.md                       ⬅ Generar UML
│   └── SAAS_TOOLS_AND_SYSTEMS.md               ⬅ Herramientas
│
├── 📊 EXCEL
│   └── DATABASE_SIMPLE.xlsx                    ⬅ Datos dummy
│
└── 📁 MIGRATIONS
    └── 001_fix_billing_cycle_and_organization_null.sql
```

---

## 🎯 Archivos por Propósito

### Para instalar la DB:
1. INSTALLATION_ORDER.md (leer primero)
2. Ejecutar los 6 archivos SQL en orden

### Para entender el sistema:
1. DATABASE_SIMPLE.xlsx (ver datos dummy)
2. FLUJOS_COMPLETOS.md (ver flujos)
3. TRIGGERS_PROCEDURES_FUNCTIONS.md (ver elementos)

### Para desarrollo:
- schema.sql (modificar tablas)
- organization_workflows.sql (agregar procedures de orgs)
- state_machine_and_workflows.sql (agregar procedures de subs)

### Para Enterprise Pro:
- ENTERPRISE_PRO_V2_README.md (documentación)
- enterprise_pro_plan_v2.sql (instalación)







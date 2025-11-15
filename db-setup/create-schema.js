import sql from 'mssql';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { readFileSync } from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Cargar variables de entorno
dotenv.config({ path: join(__dirname, '.env') });

const config = {
  server: process.env.DB_SERVER,
  database: process.env.DB_DATABASE,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  port: parseInt(process.env.DB_PORT || '1433'),
  options: {
    encrypt: process.env.DB_ENCRYPT === 'true',
    trustServerCertificate: process.env.DB_TRUST_SERVER_CERTIFICATE === 'true',
    enableArithAbort: true,
    connectionTimeout: 30000,
    requestTimeout: 60000 // Aumentado para scripts largos
  }
};

// Función para dividir SQL en batches (separados por GO)
function splitSQLBatches(sqlContent) {
  return sqlContent
    .split(/\nGO\s*\n/gi)
    .map(batch => batch.trim())
    .filter(batch => batch.length > 0 && !batch.startsWith('USE master') && !batch.startsWith('USE empower_reports'));
}

async function executeSQLFile(pool, filePath, fileName) {
  console.log(`\n📄 Ejecutando: ${fileName}...`);
  
  try {
    const sqlContent = readFileSync(filePath, 'utf8');
    const batches = splitSQLBatches(sqlContent);
    
    console.log(`   Batches encontrados: ${batches.length}`);
    
    let executedCount = 0;
    let skippedCount = 0;
    
    for (let i = 0; i < batches.length; i++) {
      const batch = batches[i];
      
      // Saltar ciertos comandos que no son compatibles o necesarios
      if (
        batch.includes('CREATE DATABASE') ||
        batch.includes('IF NOT EXISTS (SELECT * FROM sys.databases')
      ) {
        skippedCount++;
        continue;
      }
      
      try {
        await pool.request().query(batch);
        executedCount++;
        process.stdout.write(`\r   Progreso: ${executedCount}/${batches.length - skippedCount} batches ejecutados`);
      } catch (err) {
        // Ignorar errores de objetos que ya existen
        if (err.message.includes('already exists') || 
            err.message.includes('There is already an object')) {
          skippedCount++;
          continue;
        }
        throw err;
      }
    }
    
    console.log(`\n   ✅ ${fileName} ejecutado: ${executedCount} batches, ${skippedCount} omitidos`);
    return { success: true, executed: executedCount, skipped: skippedCount };
  } catch (err) {
    console.error(`\n   ❌ Error en ${fileName}:`);
    console.error(`      ${err.message}`);
    return { success: false, error: err.message };
  }
}

async function createSchema() {
  console.log('🚀 Iniciando creación del schema de Report Tuner\n');
  
  try {
    console.log('⏳ Conectando a Azure SQL Database...');
    const pool = await sql.connect(config);
    console.log('✅ Conectado exitosamente\n');

    // Verificar estado inicial
    const initialCheck = await pool.request().query(`
      SELECT COUNT(*) as TableCount 
      FROM INFORMATION_SCHEMA.TABLES 
      WHERE TABLE_TYPE = 'BASE TABLE'
    `);
    
    console.log(`📊 Estado inicial: ${initialCheck.recordset[0].TableCount} tablas existentes`);

    // Archivos SQL a ejecutar en orden
    const sqlFiles = [
      { path: '../database/schema.sql', name: 'schema.sql (Tablas principales)' },
      { path: '../database/organization_workflows.sql', name: 'organization_workflows.sql' },
      { path: '../database/state_machine_and_workflows.sql', name: 'state_machine_and_workflows.sql' },
      { path: '../database/constraints_and_validations.sql', name: 'constraints_and_validations.sql' }
    ];

    const results = [];

    // Ejecutar cada archivo
    for (const file of sqlFiles) {
      const filePath = join(__dirname, file.path);
      const result = await executeSQLFile(pool, filePath, file.name);
      results.push({ file: file.name, ...result });
    }

    // Verificar estado final
    const finalCheck = await pool.request().query(`
      SELECT COUNT(*) as TableCount 
      FROM INFORMATION_SCHEMA.TABLES 
      WHERE TABLE_TYPE = 'BASE TABLE'
    `);

    // Listar tablas creadas
    const tables = await pool.request().query(`
      SELECT TABLE_NAME 
      FROM INFORMATION_SCHEMA.TABLES 
      WHERE TABLE_TYPE = 'BASE TABLE'
      ORDER BY TABLE_NAME
    `);

    console.log('\n\n📊 Resumen de ejecución:');
    console.log('═══════════════════════════════════════════════════\n');
    
    results.forEach(result => {
      const status = result.success ? '✅' : '❌';
      console.log(`${status} ${result.file}`);
      if (result.success) {
        console.log(`   Ejecutados: ${result.executed}, Omitidos: ${result.skipped}`);
      } else {
        console.log(`   Error: ${result.error}`);
      }
    });

    console.log('\n📋 Tablas en la base de datos:');
    tables.recordset.forEach((table, index) => {
      console.log(`   ${index + 1}. ${table.TABLE_NAME}`);
    });

    console.log(`\n✨ Total de tablas: ${finalCheck.recordset[0].TableCount}`);
    console.log('\n🎉 ¡Schema creado exitosamente!');

    await pool.close();
    process.exit(0);
  } catch (err) {
    console.error('\n❌ Error fatal:');
    console.error(`   ${err.message}`);
    process.exit(1);
  }
}

createSchema();







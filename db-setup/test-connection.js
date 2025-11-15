import sql from 'mssql';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

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
    requestTimeout: 30000
  }
};

async function testConnection() {
  console.log('🔍 Probando conexión a Azure SQL Database...\n');
  console.log('📊 Configuración:');
  console.log(`   Servidor: ${config.server}`);
  console.log(`   Base de datos: ${config.database}`);
  console.log(`   Usuario: ${config.user}`);
  console.log(`   Puerto: ${config.port}`);
  console.log(`   Encrypt: ${config.options.encrypt}`);
  console.log('');

  try {
    // Intentar conectar
    console.log('⏳ Conectando...');
    const pool = await sql.connect(config);
    console.log('✅ ¡Conexión exitosa!\n');

    // Verificar la base de datos
    console.log('🔍 Verificando información de la base de datos...');
    const result = await pool.request().query(`
      SELECT 
        DB_NAME() as DatabaseName,
        SYSTEM_USER as CurrentUser,
        @@VERSION as SQLVersion,
        (SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE') as TableCount
    `);

    console.log('📊 Información de la base de datos:');
    console.log(`   Base de datos: ${result.recordset[0].DatabaseName}`);
    console.log(`   Usuario actual: ${result.recordset[0].CurrentUser}`);
    console.log(`   Tablas existentes: ${result.recordset[0].TableCount}`);
    console.log('');

    // Listar tablas existentes
    const tables = await pool.request().query(`
      SELECT TABLE_SCHEMA, TABLE_NAME 
      FROM INFORMATION_SCHEMA.TABLES 
      WHERE TABLE_TYPE = 'BASE TABLE'
      ORDER BY TABLE_SCHEMA, TABLE_NAME
    `);

    if (tables.recordset.length > 0) {
      console.log('📋 Tablas existentes:');
      tables.recordset.forEach(table => {
        console.log(`   - ${table.TABLE_SCHEMA}.${table.TABLE_NAME}`);
      });
    } else {
      console.log('ℹ️  No hay tablas creadas aún en la base de datos.');
    }

    console.log('\n✨ Test de conexión completado con éxito!');
    
    await pool.close();
    process.exit(0);
  } catch (err) {
    console.error('\n❌ Error al conectar:');
    console.error(`   Tipo: ${err.name}`);
    console.error(`   Mensaje: ${err.message}`);
    
    if (err.code) {
      console.error(`   Código: ${err.code}`);
    }

    console.log('\n💡 Posibles soluciones:');
    console.log('   1. Verifica que la contraseña sea correcta');
    console.log('   2. Verifica que el firewall de Azure SQL permita tu IP');
    console.log('   3. Ve a Azure Portal → SQL Server → Networking → Firewall rules');
    console.log('   4. Agrega tu IP pública o activa "Allow Azure services"');
    
    process.exit(1);
  }
}

testConnection();







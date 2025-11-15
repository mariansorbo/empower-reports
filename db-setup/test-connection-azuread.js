import sql from 'mssql';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { DefaultAzureCredential } from '@azure/identity';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Cargar variables de entorno
dotenv.config({ path: join(__dirname, '.env') });

async function getAzureToken() {
  try {
    const credential = new DefaultAzureCredential();
    const token = await credential.getToken('https://database.windows.net/');
    return token.token;
  } catch (err) {
    console.error('Error obteniendo token de Azure AD:', err.message);
    throw err;
  }
}

const config = {
  server: process.env.DB_SERVER,
  database: process.env.DB_DATABASE,
  port: parseInt(process.env.DB_PORT || '1433'),
  authentication: {
    type: 'azure-active-directory-default'
  },
  options: {
    encrypt: true,
    trustServerCertificate: false,
    enableArithAbort: true,
    connectionTimeout: 30000,
    requestTimeout: 30000
  }
};

async function testConnection() {
  console.log('🔍 Probando conexión con Azure AD Authentication...\n');
  console.log('📊 Configuración:');
  console.log(`   Servidor: ${config.server}`);
  console.log(`   Base de datos: ${config.database}`);
  console.log(`   Autenticación: Azure Active Directory Default`);
  console.log('');

  try {
    console.log('⏳ Conectando...');
    const pool = await sql.connect(config);
    console.log('✅ ¡Conexión exitosa con Azure AD!\n');

    // Verificar la base de datos
    const result = await pool.request().query(`
      SELECT 
        DB_NAME() as DatabaseName,
        SYSTEM_USER as CurrentUser,
        USER_NAME() as UserName,
        (SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE') as TableCount
    `);

    console.log('📊 Información de la base de datos:');
    console.log(`   Base de datos: ${result.recordset[0].DatabaseName}`);
    console.log(`   Usuario actual: ${result.recordset[0].CurrentUser}`);
    console.log(`   Tablas existentes: ${result.recordset[0].TableCount}`);
    console.log('\n✨ Test de conexión completado con éxito!');
    
    await pool.close();
    process.exit(0);
  } catch (err) {
    console.error('\n❌ Error al conectar:');
    console.error(`   ${err.message}`);
    
    console.log('\n💡 Asegúrate de:');
    console.log('   1. Estar logueado en Azure CLI: az login');
    console.log('   2. Tener permisos en la base de datos');
    console.log('   3. Tu usuario debe estar configurado como Azure AD admin del servidor');
    
    process.exit(1);
  }
}

testConnection();







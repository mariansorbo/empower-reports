import { BlobServiceClient } from '@azure/storage-blob'

// Configuración desde variables de entorno
const accountName = import.meta.env.VITE_AZURE_ACCOUNT_NAME
const sasToken = import.meta.env.VITE_AZURE_SAS_TOKEN
const connectionString = import.meta.env.VITE_AZURE_CONNECTION_STRING
const containerName = import.meta.env.VITE_CONTAINER_NAME

/**
 * Crea el cliente de BlobService usando SAS Token o Connection String
 */
const getBlobServiceClient = () => {
  // Opción 1: SAS Token (RECOMENDADO)
  if (accountName && sasToken) {
    const serviceUrl = `https://${accountName}.blob.core.windows.net?${sasToken}`
    return new BlobServiceClient(serviceUrl)
  }
  
  // Opción 2: Connection String (SOLO DESARROLLO)
  if (connectionString) {
    return BlobServiceClient.fromConnectionString(connectionString)
  }
  
  throw new Error('Missing Azure Storage configuration. Check VITE_AZURE_ACCOUNT_NAME and VITE_AZURE_SAS_TOKEN or VITE_AZURE_CONNECTION_STRING')
}

/**
 * Lista todos los archivos .pbit del container
 * @returns {Promise<Array>} Lista de archivos con metadata
 */
export const listReports = async () => {
  try {
    const blobServiceClient = getBlobServiceClient()
    const containerClient = blobServiceClient.getContainerClient(containerName)
    
    const reports = []
    
    // Iterar sobre todos los blobs en el container (sin prefix para leer todo el container)
    for await (const blob of containerClient.listBlobsFlat()) {
      // Solo incluir archivos .pbit
      if (blob.name.toLowerCase().endsWith('.pbit')) {
        // Extraer el nombre del archivo sin carpetas (si hay)
        const fileName = blob.name.split('/').pop()
        
        reports.push({
          id: blob.name, // Usar el nombre completo del blob como ID único para eliminación
          name: fileName, // Mostrar solo el nombre del archivo
          fullPath: blob.name, // Guardar el path completo por si es necesario
          date: blob.properties.lastModified?.toISOString().split('T')[0] || 'Unknown',
          size: formatBytes(blob.properties.contentLength || 0),
          sizeBytes: blob.properties.contentLength || 0,
          uploader: 'Unknown', // Azure Blob no guarda el uploader por defecto
          lastModified: blob.properties.lastModified,
          contentType: blob.properties.contentType || 'application/octet-stream'
        })
      }
    }
    
    // Ordenar por fecha de modificación (más reciente primero)
    reports.sort((a, b) => {
      if (!a.lastModified) return 1
      if (!b.lastModified) return -1
      return b.lastModified - a.lastModified
    })
    
    return reports
  } catch (error) {
    console.error('Error listing reports:', error)
    throw new Error(`Failed to load reports: ${error.message}`)
  }
}

/**
 * Elimina uno o más archivos del container
 * @param {string[]} blobNames - Array de nombres de blobs a eliminar
 * @returns {Promise<Object>} Resultado de la eliminación
 */
export const deleteReports = async (blobNames) => {
  try {
    const blobServiceClient = getBlobServiceClient()
    const containerClient = blobServiceClient.getContainerClient(containerName)
    
    const results = {
      success: [],
      failed: []
    }
    
    // Eliminar cada blob
    for (const blobName of blobNames) {
      try {
        const blockBlobClient = containerClient.getBlockBlobClient(blobName)
        await blockBlobClient.delete()
        results.success.push(blobName)
      } catch (error) {
        console.error(`Error deleting ${blobName}:`, error)
        results.failed.push({ name: blobName, error: error.message })
      }
    }
    
    return results
  } catch (error) {
    console.error('Error deleting reports:', error)
    throw new Error(`Failed to delete reports: ${error.message}`)
  }
}

/**
 * Descarga un archivo del container
 * @param {string} blobName - Nombre del blob a descargar
 * @returns {Promise<Blob>} Blob del archivo
 */
export const downloadReport = async (blobName) => {
  try {
    const blobServiceClient = getBlobServiceClient()
    const containerClient = blobServiceClient.getContainerClient(containerName)
    const blockBlobClient = containerClient.getBlockBlobClient(blobName)
    
    const downloadResponse = await blockBlobClient.download()
    return await downloadResponse.blobBody
  } catch (error) {
    console.error('Error downloading report:', error)
    throw new Error(`Failed to download report: ${error.message}`)
  }
}

/**
 * Formatea bytes a formato legible
 * @param {number} bytes - Cantidad de bytes
 * @returns {string} Tamaño formateado (ej: "2.5 MB")
 */
const formatBytes = (bytes) => {
  if (bytes === 0) return '0 Bytes'
  const k = 1024
  const sizes = ['Bytes', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
}


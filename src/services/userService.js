// Servicio para manejar usuarios
// Usa localStorage para simular persistencia del CSV
// En un entorno real, esto sería una API REST

const CSV_PATH = '/data/usuarios.csv'
const STORAGE_KEY = 'empower_reports_users'

// Función para parsear CSV
const parseCSV = (csvText) => {
  const lines = csvText.trim().split('\n')
  const headers = lines[0].split(',')
  const users = []
  
  for (let i = 1; i < lines.length; i++) {
    const values = lines[i].split(',')
    if (values.length === headers.length) {
      const user = {}
      headers.forEach((header, index) => {
        user[header.trim()] = values[index].trim()
      })
      users.push(user)
    }
  }
  
  return users
}

// Función para convertir usuarios a CSV
const usersToCSV = (users) => {
  if (users.length === 0) return 'email,password\n'
  
  const headers = Object.keys(users[0])
  const csvLines = [headers.join(',')]
  
  users.forEach(user => {
    const values = headers.map(header => user[header] || '')
    csvLines.push(values.join(','))
  })
  
  return csvLines.join('\n')
}

// Cargar usuarios desde localStorage o CSV inicial
const loadUsersFromStorage = () => {
  try {
    const stored = localStorage.getItem(STORAGE_KEY)
    if (stored) {
      return JSON.parse(stored)
    }
  } catch (error) {
    console.error('Error cargando usuarios desde localStorage:', error)
  }
  return null
}

// Guardar usuarios en localStorage
const saveUsersToStorage = (users) => {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(users))
    return true
  } catch (error) {
    console.error('Error guardando usuarios en localStorage:', error)
    return false
  }
}

// Cargar usuarios (primero desde localStorage, luego desde CSV)
export const loadUsers = async () => {
  try {
    // Primero intentar cargar desde localStorage
    const storedUsers = loadUsersFromStorage()
    if (storedUsers && storedUsers.length > 0) {
      console.log('Usuarios cargados desde localStorage:', storedUsers.length)
      return storedUsers
    }

    // Si no hay usuarios en localStorage, cargar desde CSV inicial
    const response = await fetch(CSV_PATH)
    if (!response.ok) {
      throw new Error('No se pudo cargar el archivo de usuarios')
    }
    const csvText = await response.text()
    const users = parseCSV(csvText)
    
    // Guardar en localStorage para futuras cargas
    saveUsersToStorage(users)
    console.log('Usuarios iniciales cargados desde CSV:', users.length)
    
    return users
  } catch (error) {
    console.error('Error cargando usuarios:', error)
    // Retornar usuarios por defecto si no se puede cargar nada
    const defaultUsers = [
      { email: 'admin@reporttuner.com', password: 'admin123' },
      { email: 'demo@reporttuner.com', password: 'demo123' }
    ]
    saveUsersToStorage(defaultUsers)
    return defaultUsers
  }
}

// Validar credenciales de usuario
export const validateUser = async (email, password) => {
  try {
    const users = await loadUsers()
    const user = users.find(u => u.email.toLowerCase() === email.toLowerCase())
    
    if (user && user.password === password) {
      return { success: true, user: { email: user.email, name: user.email.split('@')[0] } }
    } else {
      return { success: false, error: 'Credenciales inválidas' }
    }
  } catch (error) {
    console.error('Error validando usuario:', error)
    return { success: false, error: 'Error al validar credenciales' }
  }
}

// Verificar si un email ya existe
export const emailExists = async (email) => {
  try {
    const users = await loadUsers()
    return users.some(u => u.email.toLowerCase() === email.toLowerCase())
  } catch (error) {
    console.error('Error verificando email:', error)
    return false
  }
}

// Agregar nuevo usuario
export const addUser = async (email, password) => {
  try {
    const users = await loadUsers()
    
    // Verificar si el email ya existe
    const exists = await emailExists(email)
    if (exists) {
      return { success: false, error: 'El email ya está registrado' }
    }
    
    // Agregar nuevo usuario
    const newUser = { email, password }
    users.push(newUser)
    
    // Guardar en localStorage
    const saved = saveUsersToStorage(users)
    if (!saved) {
      return { success: false, error: 'Error al guardar el usuario' }
    }
    
    console.log('Usuario agregado y guardado:', newUser)
    console.log('Total usuarios:', users.length)
    
    return { success: true, user: { email, name: email.split('@')[0] } }
  } catch (error) {
    console.error('Error agregando usuario:', error)
    return { success: false, error: 'Error al crear cuenta' }
  }
}

// Obtener todos los usuarios (para debugging)
export const getAllUsers = async () => {
  try {
    return await loadUsers()
  } catch (error) {
    console.error('Error obteniendo usuarios:', error)
    return []
  }
}

// Limpiar usuarios del localStorage (para debugging)
export const clearUsers = () => {
  try {
    localStorage.removeItem(STORAGE_KEY)
    console.log('Usuarios limpiados del localStorage')
    return true
  } catch (error) {
    console.error('Error limpiando usuarios:', error)
    return false
  }
}

// Reiniciar usuarios a los valores por defecto del CSV
export const resetUsers = async () => {
  try {
    clearUsers()
    const response = await fetch(CSV_PATH)
    if (response.ok) {
      const csvText = await response.text()
      const users = parseCSV(csvText)
      saveUsersToStorage(users)
      console.log('Usuarios reiniciados a valores por defecto')
      return true
    }
    return false
  } catch (error) {
    console.error('Error reiniciando usuarios:', error)
    return false
  }
}

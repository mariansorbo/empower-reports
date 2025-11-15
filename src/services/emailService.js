import emailjs from '@emailjs/browser'

// Configuración de EmailJS
// Obtén estas credenciales desde https://www.emailjs.com/
const EMAILJS_SERVICE_ID = import.meta.env.VITE_EMAILJS_SERVICE_ID || ''
const EMAILJS_TEMPLATE_ID = import.meta.env.VITE_EMAILJS_TEMPLATE_ID || ''
const EMAILJS_PUBLIC_KEY = import.meta.env.VITE_EMAILJS_PUBLIC_KEY || ''

/**
 * Envía un correo de contacto/feedback usando EmailJS
 * @param {Object} formData - Datos del formulario
 * @param {string} formData.name - Nombre del usuario
 * @param {string} formData.email - Email del usuario
 * @param {string} formData.message - Mensaje/experiencia del usuario
 * @returns {Promise<Object>} Resultado del envío
 */
export const sendContactEmail = async (formData) => {
  try {
    // Validar que las credenciales estén configuradas
    if (!EMAILJS_SERVICE_ID || !EMAILJS_TEMPLATE_ID || !EMAILJS_PUBLIC_KEY) {
      throw new Error('EmailJS credentials not configured. Please check your .env.local file.')
    }

    // Preparar los datos del template
    const templateParams = {
      from_name: formData.name,
      email: formData.email, // Para reply-to
      message: formData.message,
      title: 'New Feedback', // Título del correo
      time: new Date().toLocaleString('es-AR'), // Timestamp
    }

    // Enviar el correo usando EmailJS
    const response = await emailjs.send(
      EMAILJS_SERVICE_ID,
      EMAILJS_TEMPLATE_ID,
      templateParams,
      EMAILJS_PUBLIC_KEY
    )

    console.log('Email sent successfully:', response)
    return {
      success: true,
      message: 'Your message has been sent successfully!',
      response
    }
  } catch (error) {
    console.error('Error sending email:', error)
    return {
      success: false,
      message: `Failed to send message: ${error.text || error.message}`,
      error
    }
  }
}

/**
 * Envía un correo directamente desde un elemento form
 * @param {Event} event - Evento del formulario
 * @returns {Promise<Object>} Resultado del envío
 */
export const sendContactEmailFromForm = async (event) => {
  try {
    event.preventDefault()

    // Validar que las credenciales estén configuradas
    if (!EMAILJS_SERVICE_ID || !EMAILJS_TEMPLATE_ID || !EMAILJS_PUBLIC_KEY) {
      throw new Error('EmailJS credentials not configured. Please check your .env.local file.')
    }

    // EmailJS puede leer directamente del formulario
    const response = await emailjs.sendForm(
      EMAILJS_SERVICE_ID,
      EMAILJS_TEMPLATE_ID,
      event.target,
      EMAILJS_PUBLIC_KEY
    )

    console.log('Email sent successfully:', response)
    return {
      success: true,
      message: 'Your message has been sent successfully!',
      response
    }
  } catch (error) {
    console.error('Error sending email:', error)
    return {
      success: false,
      message: `Failed to send message: ${error.text || error.message}`,
      error
    }
  }
}





import React, { useState, useEffect } from 'react'
import { useAuth } from '../contexts/AuthContext'
import { useOrganization } from '../contexts/OrganizationContext'
import './AuthModal.css'

const AuthModal = ({ isOpen, onClose, onAuthSuccess }) => {
  const [step, setStep] = useState('auth') // 'auth', 'org-setup', 'create-org', 'join-org', 'join-confirm'
  const [authMode, setAuthMode] = useState('login') // Cambio de isLogin a authMode
  const [formData, setFormData] = useState({
    email: '',
    password: '',
    confirmPassword: ''
  })
  const [error, setError] = useState('')
  const [success, setSuccess] = useState('')
  const [orgName, setOrgName] = useState('')
  const [invitationCode, setInvitationCode] = useState('')
  const [invitationPreview, setInvitationPreview] = useState(null)
  
  const { login, register, loading, isAuthenticated } = useAuth()
  const { createOrganization, joinOrganization, archiveAndJoin, keepBothOrganizations, needsSetup, loading: orgLoading } = useOrganization()
  
  // Handler for LinkedIn OAuth (simulated - automatic login)
  const handleLinkedInLogin = async () => {
    setError('')

    try {
      // Generate unique email based on timestamp to simulate different users
      const timestamp = Date.now()
      const linkedInUserData = {
        email: `user${timestamp}@linkedin.com`,
        name: `User ${timestamp}`,
        provider: 'linkedin'
      }
      
      // Simulated automatic login
      const result = await register(
        linkedInUserData.email,
        'linkedin_auto',
        'linkedin_auto'
      )

      if (!result.success) {
        setError(result.error || 'Error authenticating with LinkedIn')
      }
      // useEffect will handle transition to org-setup if needed
    } catch (err) {
      setError('Error connecting with LinkedIn')
    }
  }

  // Handler for Azure AD OAuth (simulated - automatic login)
  const handleAzureADLogin = async () => {
    setError('')

    try {
      // Generate unique email based on timestamp to simulate different users
      const timestamp = Date.now()
      const azureUserData = {
        email: `user${timestamp}@company.com`,
        name: `Corporate User ${timestamp}`,
        provider: 'azure_ad'
      }
      
      // Simulated automatic login
      const result = await register(
        azureUserData.email,
        'azure_auto',
        'azure_auto'
      )

      if (!result.success) {
        setError(result.error || 'Error authenticating with Azure AD')
      }
      // useEffect will handle transition to org-setup if needed
    } catch (err) {
      setError('Error connecting with Azure AD')
    }
  }

  // Reset when opening/closing the modal
  useEffect(() => {
    if (isOpen) {
      setStep('auth')
      setError('')
      setSuccess('')
      setFormData({ email: '', password: '', confirmPassword: '' })
      setOrgName('')
      setInvitationCode('')
      setInvitationPreview(null)
    }
  }, [isOpen])

  // If authenticated and needs setup, show organization options
  useEffect(() => {
    if (isAuthenticated && needsSetup && !orgLoading && step === 'auth') {
      setStep('org-setup')
    }
    // If not authenticated, always show auth step
    if (!isAuthenticated && step !== 'auth') {
      setStep('auth')
    }
  }, [isAuthenticated, needsSetup, orgLoading, step])

  const handleInputChange = (e) => {
    const { name, value } = e.target
    setFormData(prev => ({
      ...prev,
      [name]: value
    }))
    if (error) setError('')
  }

  const handleAuthSubmit = async (e) => {
    e.preventDefault()
    setError('')
    setSuccess('')

    try {
      let result
      if (isLogin) {
        result = await login(formData.email, formData.password)
      } else {
        result = await register(formData.email, formData.password, formData.confirmPassword)
      }

      if (result.success) {
        setSuccess(authMode === 'login' ? '¡Inicio de sesión exitoso!' : '¡Cuenta creada exitosamente!')
        // No cerrar el modal, esperar a que el useEffect detecte needsSetup
      } else {
        setError(result.error)
      }
    } catch (err) {
      setError('Error inesperado. Intenta nuevamente.')
    }
  }

  const handleCreateOrg = async (e) => {
    e.preventDefault()
    setError('')

    // Validate authentication
    if (!isAuthenticated) {
      setError('You must be authenticated to create an organization')
      setStep('auth')
      return
    }

    if (!orgName.trim()) {
      setError('Organization name is required')
      return
    }

    if (orgName.trim().length < 3) {
      setError('Name must be at least 3 characters')
      return
    }

    const result = await createOrganization(orgName.trim())
    
    if (result.success) {
      if (onAuthSuccess) {
        onAuthSuccess()
      }
      onClose()
    } else {
      setError(result.error || 'Error creating organization')
    }
  }

  const handleJoinValidate = async (e) => {
    e.preventDefault()
    setError('')
    setInvitationPreview(null)

    if (!invitationCode.trim()) {
      setError('Por favor ingresá el código de invitación')
      return
    }

    const result = await joinOrganization(invitationCode.trim())
    
    if (result.success) {
      if (result.hasExistingOrganization) {
        // Mostrar confirmación
        setInvitationPreview(result)
        setStep('join-confirm')
      } else {
        // Unirse directamente
        if (onAuthSuccess) {
          onAuthSuccess()
        }
        onClose()
      }
    } else {
      setError(result.error || 'Código de invitación inválido')
    }
  }

  const handleJoinDecision = async (action) => {
    let result
    if (action === 'archive') {
      result = await archiveAndJoin(
        invitationPreview.existingOrganization.id,
        invitationPreview.organization.id
      )
    } else {
      result = await keepBothOrganizations(invitationPreview.organization.id)
    }
    
    if (result.success) {
      if (onAuthSuccess) {
        onAuthSuccess()
      }
      onClose()
    } else {
      setError(result.error || 'Error al procesar la unión')
    }
  }

  const toggleMode = () => {
    setAuthMode(authMode === 'login' ? 'register' : 'login')
    setError('')
    setSuccess('')
    setFormData({ email: '', password: '', confirmPassword: '' })
  }

  if (!isOpen) return null

  // Paso 1: Autenticación (Login/Registro)
  if (step === 'auth') {
    return (
      <div className="auth-modal-overlay" onClick={onClose}>
        <div className="auth-modal" onClick={(e) => e.stopPropagation()}>
          <button 
            className="auth-modal-close" 
            onClick={(e) => {
              e.stopPropagation()
              onClose()
            }}
            type="button"
          >
            ×
          </button>
          
          <div className="auth-modal-header">
            <h2>Acceder a Report Tuner</h2>
            <p className="auth-modal-subtitle">
              Elegí tu método de autenticación preferido
            </p>
          </div>

          <div className="auth-modal-content">
            <div className="auth-options-container">
              {error && (
                <div className="auth-error">
                  ❌ {error}
                </div>
              )}

              <div className="auth-providers-row">
                {/* LinkedIn Option */}
                <div className="auth-option-card">
                  <div className="auth-option-header">
                    <div className="auth-provider-logo linkedin-logo-small">in</div>
                    <h3>LinkedIn</h3>
                  </div>
                  <p className="auth-option-description">
                    Accede con tu cuenta profesional de LinkedIn
                  </p>
                  <button 
                    type="button"
                    className="btn btn-linkedin full" 
                    onClick={handleLinkedInLogin}
                    disabled={loading}
                  >
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                      <path d="M19 3a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h14m-.5 15.5v-5.3a3.26 3.26 0 0 0-3.26-3.26c-.85 0-1.84.52-2.32 1.3v-1.11h-2.79v8.37h2.79v-4.93c0-.77.62-1.4 1.39-1.4a1.4 1.4 0 0 1 1.4 1.4v4.93h2.79M6.88 8.56a1.68 1.68 0 0 0 1.68-1.68c0-.93-.75-1.69-1.68-1.69a1.69 1.69 0 0 0-1.69 1.69c0 .93.76 1.68 1.69 1.68m1.39 9.94v-8.37H5.5v8.37h2.77z"/>
                    </svg>
                    {loading ? 'Conectando...' : 'Continuar con LinkedIn'}
                  </button>
                </div>

                {/* Azure AD Option */}
                <div className="auth-option-card">
                  <div className="auth-option-header">
                    <div className="auth-provider-logo azure-logo-small">
                      <svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor">
                        <path d="M0 0h11.377v11.372H0zm12.623 0H24v11.372H12.623zM0 12.623h11.377V24H0zm12.623 0H24V24H12.623z"/>
                      </svg>
                    </div>
                    <h3>Cuenta Empresarial</h3>
                  </div>
                  <p className="auth-option-description">
                    Accede con tu cuenta corporativa de Microsoft
                  </p>
                  <button 
                    type="button"
                    className="btn btn-azure full" 
                    onClick={handleAzureADLogin}
                    disabled={loading}
                  >
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                      <path d="M0 0h11.377v11.372H0zm12.623 0H24v11.372H12.623zM0 12.623h11.377V24H0zm12.623 0H24V24H12.623z"/>
                    </svg>
                    {loading ? 'Conectando...' : 'Continuar con cuenta corporativa'}
                  </button>
                </div>
              </div>

              <div className="auth-privacy-notice">
                <p>🔒 Tu información está protegida. Solo usamos estos métodos para autenticación segura.</p>
              </div>
            </div>
          </div>
        </div>
      </div>
    )
  }

  // Paso 2: Setup de organización (opciones)
  if (step === 'org-setup') {
    // Si no está autenticado, volver a auth
    if (!isAuthenticated) {
      setStep('auth')
      return null
    }
    return (
      <div className="auth-modal-overlay" onClick={onClose}>
        <div className="auth-modal auth-modal-org-setup" onClick={(e) => e.stopPropagation()}>
          <button 
            className="auth-modal-close" 
            onClick={(e) => {
              e.stopPropagation()
              onClose()
            }}
            type="button"
          >
            ×
          </button>
          
          <div className="auth-modal-header">
            <h2>Welcome to Report Tuner</h2>
            <p className="auth-modal-subtitle">
              Create your workspace or join your team's
            </p>
          </div>

          <div className="org-setup-options-container">
            <button 
              className="org-setup-option-btn org-setup-primary"
              onClick={() => setStep('create-org')}
            >
              <div className="org-setup-icon">➕</div>
              <div className="org-setup-content">
                <h3>Create new organization</h3>
                <p>Start your own workspace from scratch</p>
              </div>
            </button>

            <button 
              className="org-setup-option-btn org-setup-secondary"
              onClick={() => setStep('join-org')}
            >
              <div className="org-setup-icon">🔗</div>
              <div className="org-setup-content">
                <h3>Join an existing organization</h3>
                <p>You have an invitation code or invitation link</p>
              </div>
            </button>
          </div>
        </div>
      </div>
    )
  }

  // Paso 3: Crear organización
  if (step === 'create-org') {
    // Si no está autenticado, volver a auth
    if (!isAuthenticated) {
      setStep('auth')
      return null
    }
    return (
      <div className="auth-modal-overlay" onClick={onClose}>
        <div className="auth-modal" onClick={(e) => e.stopPropagation()}>
          <button 
            className="auth-modal-back" 
            onClick={(e) => {
              e.stopPropagation()
              setStep('org-setup')
            }}
            type="button"
            title="Volver"
          >
            ←
          </button>
          
          <button 
            className="auth-modal-close" 
            onClick={(e) => {
              e.stopPropagation()
              onClose()
            }}
            type="button"
          >
            ×
          </button>
          
          <div className="auth-modal-header">
            <h2>Create Organization</h2>
            <p className="auth-modal-subtitle">
              Start your workspace from scratch
            </p>
          </div>

          <form onSubmit={handleCreateOrg} className="auth-form">
            <div className="form-group">
              <label htmlFor="org-name">Organization name</label>
              <input
                type="text"
                id="org-name"
                value={orgName}
                onChange={(e) => {
                  setOrgName(e.target.value)
                  setError('')
                }}
                placeholder="Ex: My Company"
                required
                minLength={3}
                autoFocus
              />
            </div>

            {error && (
              <div className="auth-error">
                ❌ {error}
              </div>
            )}

            <div className="auth-form-actions-center">
              <button 
                type="submit" 
                className="auth-submit-btn"
                disabled={orgLoading || !orgName.trim()}
              >
                {orgLoading ? 'Creating...' : 'Create Organization'}
              </button>
            </div>
          </form>
        </div>
      </div>
    )
  }

  // Paso 4: Unirse a organización (ingresar código)
  if (step === 'join-org') {
    // Si no está autenticado, volver a auth
    if (!isAuthenticated) {
      setStep('auth')
      return null
    }
    return (
      <div className="auth-modal-overlay" onClick={onClose}>
        <div className="auth-modal" onClick={(e) => e.stopPropagation()}>
          <button 
            className="auth-modal-back" 
            onClick={(e) => {
              e.stopPropagation()
              setStep('org-setup')
              setInvitationCode('')
              setError('')
            }}
            type="button"
            title="Volver"
          >
            ←
          </button>
          
          <button 
            className="auth-modal-close" 
            onClick={(e) => {
              e.stopPropagation()
              onClose()
            }}
            type="button"
          >
            ×
          </button>
          
          <div className="auth-modal-header">
            <h2>Join an Organization</h2>
            <p className="auth-modal-subtitle">
              Enter the invitation code you received
            </p>
          </div>

          <form onSubmit={handleJoinValidate} className="auth-form">
            <div className="form-group">
              <label htmlFor="invitation-code">Invitation code</label>
              <input
                type="text"
                id="invitation-code"
                value={invitationCode}
                onChange={(e) => {
                  setInvitationCode(e.target.value.toUpperCase())
                  setError('')
                }}
                placeholder="Ex: DATA-LATAM-2024"
                required
                autoFocus
              />
              <small className="input-hint">
                You can also paste a full invitation link
              </small>
            </div>

            {error && (
              <div className="auth-error">
                ❌ {error}
              </div>
            )}

            <div className="auth-form-actions-center">
              <button 
                type="submit" 
                className="auth-submit-btn"
                disabled={orgLoading || !invitationCode.trim()}
              >
                {orgLoading ? 'Validating...' : 'Validate Code'}
              </button>
            </div>
          </form>
        </div>
      </div>
    )
  }

  // Paso 5: Confirmación de unión (cuando tiene organización existente)
  if (step === 'join-confirm' && invitationPreview) {
    // Si no está autenticado, volver a auth
    if (!isAuthenticated) {
      setStep('auth')
      return null
    }
    const { organization, existingOrganization } = invitationPreview
    
    return (
      <div className="auth-modal-overlay" onClick={onClose}>
        <div className="auth-modal auth-modal-large" onClick={(e) => e.stopPropagation()}>
          <button 
            className="auth-modal-close" 
            onClick={(e) => {
              e.stopPropagation()
              setStep('join-org')
              setInvitationPreview(null)
            }}
            type="button"
          >
            ×
          </button>
          
          <div className="auth-modal-header">
            <h2>You Already Belong to Another Organization</h2>
            <p className="auth-modal-subtitle">
              You are currently an admin of the organization <strong>'{existingOrganization.name}'</strong>.
              <br />
              If you join <strong>'{organization.name}'</strong>, you can archive your current organization.
              <br />
              <span style={{ fontSize: '13px', color: '#999', fontStyle: 'italic' }}>
                Your reports and settings will not be lost.
              </span>
            </p>
          </div>

          <div className="join-confirmation-org-preview">
            <div className="org-preview-card org-preview-new">
              <div className="org-preview-label">New organization</div>
              <div className="org-preview-name">{organization.name}</div>
              {organization.admin && (
                <div className="org-preview-detail">👤 Admin: {organization.admin}</div>
              )}
              {organization.members && (
                <div className="org-preview-detail">👥 Members: {organization.members}</div>
              )}
            </div>
            
            <div className="org-preview-card org-preview-existing">
              <div className="org-preview-label">Your current organization</div>
              <div className="org-preview-name">{existingOrganization.name}</div>
              <div className="org-preview-detail">📊 Status: Active</div>
            </div>
          </div>

          <div className="join-confirmation-actions">
            <button
              className="auth-submit-btn auth-submit-primary"
              onClick={() => handleJoinDecision('archive')}
              disabled={orgLoading}
            >
              ✅ Join and archive my current organization
            </button>
            <button
              className="auth-secondary-btn"
              onClick={() => handleJoinDecision('keep')}
              disabled={orgLoading}
            >
              ⚙️ Keep both organizations
            </button>
            <button
              className="auth-cancel-btn"
              onClick={() => {
                setStep('join-org')
                setInvitationPreview(null)
                setInvitationCode('')
              }}
            >
              Cancel
            </button>
          </div>
        </div>
      </div>
    )
  }

  return null
}

export default AuthModal

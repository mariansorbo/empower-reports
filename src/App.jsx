import React, { useState, useEffect } from 'react'
import FileUpload from './components/FileUpload'
import AuthModal from './components/AuthModal'
import ReportsModal from './components/ReportsModal'
import WelcomeSetup from './components/WelcomeSetup'
import CreateOrganizationModal from './components/CreateOrganizationModal'
import OrganizationSelector from './components/OrganizationSelector'
import OrganizationsSettingsPanel from './components/OrganizationsSettingsPanel'
import InviteMembersModal from './components/InviteMembersModal'
import SettingsPanel from './components/SettingsPanel'
import FAQs from './components/FAQs'
import { AuthProvider, useAuth } from './contexts/AuthContext'
import { OrganizationProvider, useOrganization } from './contexts/OrganizationContext'
import { useDocumentation } from './hooks/useDocumentation'
import { sendContactEmailFromForm } from './services/emailService'
import './App.css'

const AppContent = () => {
  const [currentView, setCurrentView] = useState('home') // 'home', 'faqs'
  const [showAuthModal, setShowAuthModal] = useState(false)
  const [showReportsModal, setShowReportsModal] = useState(false)
  const [showCreateOrgModal, setShowCreateOrgModal] = useState(false)
  const [showOrgSettings, setShowOrgSettings] = useState(false)
  const [showInviteModal, setShowInviteModal] = useState(false)
  const [notification, setNotification] = useState(null)
  const [formSubmitting, setFormSubmitting] = useState(false)
  const { user, logout, isAuthenticated } = useAuth()
  const { needsSetup, currentOrganization, loading: orgLoading, refreshOrganizations, userOrganizations } = useOrganization()
  const { documentationUrl, hasDocumentation, openDocumentation } = useDocumentation()

  // Handle hash routing
  useEffect(() => {
    const handleHashChange = () => {
      const hash = window.location.hash
      if (hash === '#faqs') {
        setCurrentView('faqs')
      } else {
        setCurrentView('home')
      }
    }

    // Check initial hash
    handleHashChange()

    // Listen for hash changes
    window.addEventListener('hashchange', handleHashChange)
    return () => window.removeEventListener('hashchange', handleHashChange)
  }, [])

  useEffect(() => {
    if (notification) {
      const timer = setTimeout(() => setNotification(null), 5000)
      return () => clearTimeout(timer)
    }
  }, [notification])

  const handleAuthClick = () => {
    setShowAuthModal(true)
  }

  const handleAuthSuccess = async () => {
    // Auth modal handles organization setup internally
    // We just need to refresh organizations
    await refreshOrganizations()
  }

  const handleLogout = () => {
    logout()
  }

  const handleOrgCreated = (organization) => {
    setNotification({
      type: 'success',
      message: `✅ Your organization "${organization.name}" was successfully created.`
    })
  }

  const handleContactFormSubmit = async (event) => {
    event.preventDefault()
    setFormSubmitting(true)

    try {
      const result = await sendContactEmailFromForm(event)
      
      if (result.success) {
        setNotification({
          type: 'success',
          message: '✅ Thank you! Your message has been sent successfully.'
        })
        // Limpiar el formulario
        event.target.reset()
      } else {
        setNotification({
          type: 'error',
          message: `❌ ${result.message}`
        })
      }
    } catch (error) {
      setNotification({
        type: 'error',
        message: '❌ An error occurred while sending your message. Please try again.'
      })
    } finally {
      setFormSubmitting(false)
    }
  }


  // If needs setup and authenticated, show WelcomeSetup as full screen
  // Only if not showing the setup modal (to avoid conflict)
  if (isAuthenticated && needsSetup && !orgLoading && !showAuthModal && !showCreateOrgModal && currentView === 'home') {
    return (
      <WelcomeSetup
        onCreateOrganization={() => setShowCreateOrgModal(true)}
        onJoinOrganization={() => {}}
      />
    )
  }

  // Show FAQs page
  if (currentView === 'faqs') {
    return (
      <>
        <FAQs onBack={() => {
          setCurrentView('home')
          window.location.hash = ''
        }} />
      </>
    )
  }

  return (
    <div className="site">
      <div className="powerbi-banner">⚠️ Only works for users with Power BI Pro or PPU</div>
      <div className="beta-banner">🧪 Beta version in testing. Results and times may vary.</div>

      {notification && (
        <div className={`notification-banner ${notification.type}`}>
          {notification.message}
        </div>
      )}

      <header className="site-header">
        <div className="site-header__left">
          <div className="brand">
            <span className="brand__logo">📊</span>
            <span className="brand__name">Report <span>Tuner</span></span>
          </div>
        </div>
        <nav className="site-nav">
          <a 
            className="nav-link" 
            href="#faqs"
            onClick={(e) => {
              e.preventDefault()
              setCurrentView('faqs')
              window.location.hash = 'faqs'
            }}
          >
            FAQs
          </a>
          <a 
            className="nav-link" 
            href="#about"
            onClick={(e) => {
              e.preventDefault()
              if (currentView === 'faqs') {
                setCurrentView('home')
                window.location.hash = ''
                setTimeout(() => {
                  const element = document.getElementById('about')
                  if (element) {
                    element.scrollIntoView({ behavior: 'smooth', block: 'start' })
                  }
                }, 100)
              } else {
                const element = document.getElementById('about')
                if (element) {
                  element.scrollIntoView({ behavior: 'smooth', block: 'start' })
                }
              }
            }}
          >
            Quiénes somos
          </a>
          <a 
            className="nav-link" 
            href="#contact"
            onClick={(e) => {
              e.preventDefault()
              if (currentView === 'faqs') {
                setCurrentView('home')
                window.location.hash = ''
                setTimeout(() => {
                  const element = document.getElementById('contact')
                  if (element) {
                    element.scrollIntoView({ behavior: 'smooth', block: 'start' })
                  }
                }, 100)
              } else {
                const element = document.getElementById('contact')
                if (element) {
                  element.scrollIntoView({ behavior: 'smooth', block: 'start' })
                }
              }
            }}
          >
            Contacto
          </a>
        </nav>
        <div className="site-header__actions">
          {isAuthenticated && currentOrganization && (
            <button
              className="btn btn-secondary"
              onClick={() => setShowInviteModal(true)}
              style={{ marginRight: '12px' }}
            >
              👥 Invite Team
            </button>
          )}
          <a className="btn btn-secondary" href="#docs">View Documentation</a>
          {isAuthenticated ? (
            <div className="user-menu">
              {currentOrganization && <OrganizationSelector />}
              <button
                className="btn btn-secondary"
                onClick={() => setShowOrgSettings(true)}
                title="Settings"
              >
                ⚙️
              </button>
              <span className="user-greeting">Hello, {user?.name}</span>
              <button className="btn btn-primary-light" onClick={handleLogout}>
                Logout
              </button>
            </div>
          ) : (
            <button className="btn btn-primary-light" onClick={handleAuthClick}>
              → Sign In
            </button>
          )}
        </div>
      </header>

      <main className="hero">
        <section className="hero__left">
          <h1 className="hero__title">Report <span>Tuner</span></h1>
          <p className="hero__subtitle">Document the internal logic of Power BI reports in a clear and navigable way. Empower assessments, analysis and new developments.</p>

          <ul className="hero__bullets">
            <li className="bullet bad">No more depending on the original developer.</li>
            <li className="bullet bad">No more navigating Power Query like a black box.</li>
            <li className="bullet bad">No more manual documentation in Excel or Notion.</li>
            <li className="bullet good">Reverse engineering is fast and visual.</li>
            <li className="bullet good">Drives new developments with consistency.</li>
            <li className="bullet good">Promotes DAX and model standardization.</li>
            <li className="bullet good">Improves collaborative work.</li>
          </ul>
        </section>

        <section className="hero__right">
          <div className="card info">
            <h3>About the .pbit file</h3>
            <p>The .pbit file is the report template, it contains the model structure but not the data. This way, Report Tuner analyzes your logic without accessing sensitive information.</p>
          </div>

          <div className="card upload">
            <div className="upload__title">Drag your .pbit file here</div>
            <FileUpload compact={true} onAuthRequired={handleAuthClick} />
            <button 
              className="btn btn-primary-light full"
              onClick={() => setShowReportsModal(true)}
              style={{ marginTop: '12px' }}
            >
              📋 View Reports
            </button>
            <button 
              className={`btn full ${hasDocumentation ? 'btn-documentation-enabled' : 'btn-documentation-disabled'}`}
              onClick={openDocumentation}
              disabled={!hasDocumentation}
              title={hasDocumentation ? 'View organization documentation' : 'No documentation configured'}
            >
              📚 View Documentation
            </button>
          </div>

          <div className="card help" id="faqs">
            <h3>How to get your .pbit?</h3>
            <ol>
              <li>Open the .pbix in Power BI Desktop.</li>
              <li>File → Export → Power BI Template (.pbit).</li>
              <li>Save and drag here.</li>
            </ol>
          </div>
        </section>
      </main>

      <section id="about" className="about-section">
        <div className="about-container">
          <h2 className="about-title">¿Quiénes somos?</h2>
          <div className="about-content">
            <div className="about-card">
              <div className="about-icon">🎯</div>
              <h3>Nuestra Misión</h3>
              <p>Report Tuner nace de la experiencia directa con los desafíos de mantener y comprender modelos complejos de Power BI.</p>
            </div>
            <div className="about-card">
              <div className="about-icon">👥</div>
              <h3>Nuestro Equipo</h3>
              <p>Somos un equipo de desarrolladores y analistas que creemos que la documentación debe ser una herramienta de crecimiento, no un obstáculo.</p>
            </div>
            <div className="about-card">
              <div className="about-icon">⚡</div>
              <h3>Nuestra Visión</h3>
              <p>Creamos esta plataforma para hacer visible la lógica detrás de cada modelo, acelerar la colaboración y facilitar el trabajo técnico de quienes construyen reportes día a día.</p>
            </div>
          </div>
        </div>
      </section>

      <section id="contact" className="feedback-section">
        <div className="feedback-container">
          <div className="feedback-header">
            <div className="feedback-icon-large">💬</div>
            <h2 className="feedback-title">We Want to Hear About Your Experience</h2>
            <p className="feedback-description">
              This project is in testing phase. Tell us what you think, what we could improve, or how it should evolve.
              <br/>
              <strong>Every comment helps us build a more useful tool for Power BI users.</strong>
            </p>
          </div>
          
          <form className="feedback-form" onSubmit={handleContactFormSubmit}>
            <div className="form-group">
              <label htmlFor="nombre">Name *</label>
              <input 
                type="text" 
                id="nombre" 
                name="from_name" 
                required 
                disabled={formSubmitting}
              />
            </div>
            
            <div className="form-group">
              <label htmlFor="email">Email *</label>
              <input 
                type="email" 
                id="email" 
                name="email" 
                required 
                disabled={formSubmitting}
              />
            </div>
            
            <div className="form-group">
              <label htmlFor="experiencia">Your Experience</label>
              <textarea 
                id="experiencia" 
                name="message" 
                rows="4" 
                placeholder="Tell us what you thought, suggestions, ideas...."
                disabled={formSubmitting}
              ></textarea>
            </div>
            
            <button 
              type="submit" 
              className="btn btn-feedback"
              disabled={formSubmitting}
            >
              {formSubmitting ? 'Sending...' : 'Send Feedback'}
            </button>
          </form>
          
          <div className="feedback-footer">
            <p>🧡 Thank you for helping us improve 💛</p>
            <p>Report Tuner is in beta: your input has a real impact.</p>
          </div>
        </div>
      </section>

      <footer className="site-footer">
        <div className="footer__brand">Report Tuner</div>
        <div className="footer__legal">© 2024 Report Tuner. All rights reserved.</div>
      </footer>

      <AuthModal 
        isOpen={showAuthModal} 
        onClose={() => setShowAuthModal(false)}
        onAuthSuccess={handleAuthSuccess}
      />

      <ReportsModal
        isOpen={showReportsModal}
        onClose={() => setShowReportsModal(false)}
      />


      <SettingsPanel
        isOpen={showOrgSettings}
        onClose={() => setShowOrgSettings(false)}
      />

      <InviteMembersModal
        isOpen={showInviteModal}
        onClose={() => setShowInviteModal(false)}
        onSave={(data) => {
          console.log('Invitations saved:', data)
          setNotification({
            type: 'success',
            message: `✅ Invitations sent to ${data.members.length} team member${data.members.length > 1 ? 's' : ''}`
          })
        }}
      />
    </div>
  )
}

function App() {
  return (
    <AuthProvider>
      <OrganizationProvider>
        <AppContent />
      </OrganizationProvider>
    </AuthProvider>
  )
}

export default App

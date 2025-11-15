import React, { useState } from 'react'
import { useAuth } from '../contexts/AuthContext'
import { useOrganization } from '../contexts/OrganizationContext'
import './SettingsPanel.css'

const SettingsPanel = ({ isOpen, onClose }) => {
  const [activeSection, setActiveSection] = useState('profile')
  const { user, logout } = useAuth()
  const { currentOrganization } = useOrganization()

  // Perfil state
  const [profileData, setProfileData] = useState({
    name: user?.name || '',
    email: user?.email || ''
  })

  if (!isOpen) return null

  const sections = [
    { id: 'profile', label: 'Profile', icon: '👤' },
    { id: 'security', label: 'Security', icon: '🔐' },
    { id: 'organization', label: 'Organization', icon: '🏢' },
    { id: 'billing', label: 'Billing & Plan', icon: '💳' },
    { id: 'integrations', label: 'Integrations', icon: '🔌' }
  ]

  const handleSaveProfile = () => {
    // Aquí guardarías los cambios
    console.log('Saving profile:', profileData)
  }

  const renderContent = () => {
    switch (activeSection) {
      case 'profile':
        return (
          <div className="settings-section">
            <h2>Profile</h2>
            <p className="settings-description">Individual account settings</p>
            
            <form className="settings-form">
              <div className="form-group">
                <label>Full name</label>
                <input
                  type="text"
                  value={profileData.name}
                  onChange={(e) => setProfileData({ ...profileData, name: e.target.value })}
                  placeholder="John Doe"
                />
                <small>Displayed in reports and comments</small>
              </div>

              <div className="form-group">
                <label>Email</label>
                <input
                  type="email"
                  value={profileData.email}
                  disabled
                  className="disabled-input"
                />
                <small>Read-only (provided by LinkedIn)</small>
              </div>

              <div className="form-actions">
                <button type="button" className="btn-secondary">Undo</button>
                <button type="button" className="btn-primary" onClick={handleSaveProfile}>
                  Save changes
                </button>
              </div>
            </form>
          </div>
        )

      case 'security':
        return (
          <div className="settings-section">
            <h2>Security</h2>
            <p className="settings-description">Linked accounts and security settings</p>

            <div className="security-card">
              <h3>Linked accounts</h3>
              <div className="linked-account">
                <span>💼 LinkedIn</span>
                <span className="account-status">Connected</span>
              </div>
              <button type="button" className="btn-link" disabled style={{ opacity: 0.5, cursor: 'not-allowed' }}>
                Unlink account
              </button>
              <small className="settings-note">For now, LinkedIn is the only available authentication method</small>
            </div>
          </div>
        )

      case 'organization':
        return (
          <div className="settings-section">
            <h2>Organization</h2>
            {currentOrganization ? (
              <>
                <div className="form-group">
                  <label>Organization name</label>
                  <input type="text" value={currentOrganization.name || ''} />
                  <button type="button" className="btn-primary" style={{ marginTop: '12px' }}>
                    Update name
                  </button>
                </div>

                <div className="organization-section">
                  <h3>Roles and permissions</h3>
                  <p className="settings-description">Permission structure within your organization</p>
                  <div className="role-info">
                    <div className="role-item">
                      <strong>Admin</strong>
                      <p>Full access, can invite and remove members, manage settings</p>
                    </div>
                    <div className="role-item">
                      <strong>Member</strong>
                      <p>Can create, edit, and delete reports</p>
                    </div>
                    <div className="role-item">
                      <strong>Viewer</strong>
                      <p>Read-only access to reports, no editing permissions</p>
                    </div>
                  </div>
                </div>

                <div className="organization-section danger-zone">
                  <h3>⚠️ Danger Zone</h3>
                  <p>This action cannot be undone. All reports and data will be permanently deleted.</p>
                  <div className="danger-zone-actions">
                    <button type="button" className="btn-danger">Delete this organization</button>
                  </div>
                </div>
              </>
            ) : (
              <div className="no-org-message">
                <p>📭 You don't belong to an organization.</p>
                <p>Create one or join an existing one to start collaborating.</p>
                <button type="button" className="btn-primary">Create organization</button>
              </div>
            )}
          </div>
        )

      case 'billing':
        return (
          <div className="settings-section">
            <h2>Billing & Plan</h2>
            <p className="settings-description">Manage your current subscription</p>

            <div className="beta-notice">
              <h3>🎉 Free Beta Version</h3>
              <p>
                You are currently using Report Tuner for free during the beta phase.
                No payment method is required and there are no charges.
              </p>
            </div>

            <div className="billing-card">
              <div className="billing-header">
                <div>
                  <h3>Current plan</h3>
                  <p className="plan-name">Free Trial</p>
                </div>
                <span className="badge-active">Active</span>
              </div>
              <div className="billing-details">
                <div>
                  <small>Status</small>
                  <p>Free beta · No time limit</p>
                </div>
                <div>
                  <small>Current limits</small>
                  <p>10 users · 100 reports</p>
                </div>
              </div>
            </div>

            <div className="billing-card disabled-card">
              <h3>💳 Payment method</h3>
              <p className="disabled-text">
                Not available during free beta version.
                <br />
                <small>Will be enabled when we launch paid plans.</small>
              </p>
            </div>

            <div className="billing-card disabled-card">
              <h3>📄 Billing history</h3>
              <p className="disabled-text">
                Not available during free beta version.
                <br />
                <small>You'll see your invoices here when you start paying.</small>
              </p>
            </div>

            <p className="billing-note">
              💡 <strong>Coming soon:</strong> We'll launch Basic, Teams, and Enterprise plans with additional features.
            </p>
          </div>
        )

      case 'integrations':
        return (
          <div className="settings-section">
            <h2>Integrations</h2>
            <p className="settings-description">Connect Report Tuner with other tools (in development)</p>

            <div className="integration-list">
              <div className="integration-item integration-disabled">
                <div>
                  <strong>🔌 Power BI Service</strong>
                  <p>Connect directly to your Power BI workspace</p>
                  <span className="status-badge status-soon">Coming soon</span>
                </div>
                <button type="button" className="btn-secondary" disabled>Not available</button>
              </div>

              <div className="integration-item integration-disabled">
                <div>
                  <strong>📡 API - Upload .pbit files</strong>
                  <p>Automate report uploads via REST API</p>
                  <span className="status-badge status-soon">In development</span>
                </div>
                <button type="button" className="btn-secondary" disabled>Not available</button>
              </div>

              <div className="integration-item integration-disabled">
                <div>
                  <strong>📚 API - Documentation endpoint</strong>
                  <p>Access generated documentation via endpoint</p>
                  <span className="status-badge status-soon">In development</span>
                </div>
                <button type="button" className="btn-secondary" disabled>Not available</button>
              </div>
            </div>

            <div className="integration-note">
              <p>
                💡 <strong>Need another integration?</strong>
                <br />
                Send us your suggestion through the feedback form on the main page.
              </p>
            </div>
          </div>
        )

      default:
        return null
    }
  }

  return (
    <div className="settings-overlay" onClick={onClose}>
      <div className="settings-panel" onClick={(e) => e.stopPropagation()}>
        <div className="settings-sidebar">
          <div className="settings-header">
            <h2>⚙️ Settings</h2>
            <button className="settings-close" onClick={onClose}>×</button>
          </div>
          <nav className="settings-nav">
            {sections.map((section) => (
              <button
                key={section.id}
                className={`settings-nav-item ${activeSection === section.id ? 'active' : ''}`}
                onClick={() => setActiveSection(section.id)}
              >
                <span className="nav-icon">{section.icon}</span>
                <span className="nav-label">{section.label}</span>
              </button>
            ))}
          </nav>
        </div>
        <div className="settings-content">
          {renderContent()}
        </div>
      </div>
    </div>
  )
}

export default SettingsPanel



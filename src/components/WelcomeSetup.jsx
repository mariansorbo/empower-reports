import React from 'react'
import './WelcomeSetup.css'

const WelcomeSetup = ({ onCreateOrganization }) => {
  return (
    <div className="welcome-setup-overlay">
      <div className="welcome-setup">
        <div className="welcome-setup-header">
          <div className="welcome-logo">📊</div>
          <h1>Welcome to Report Tuner</h1>
          <p className="welcome-subtitle">
            Create your workspace to get started
          </p>
        </div>

        <div className="welcome-setup-options">
          <button 
            className="welcome-option-btn welcome-option-primary"
            onClick={onCreateOrganization}
            style={{ width: '100%', maxWidth: '400px', margin: '0 auto' }}
          >
            <div className="welcome-option-icon">➕</div>
            <div className="welcome-option-content">
              <h3>Create new organization</h3>
              <p>Start your own workspace from scratch</p>
            </div>
          </button>
        </div>
      </div>
    </div>
  )
}

export default WelcomeSetup

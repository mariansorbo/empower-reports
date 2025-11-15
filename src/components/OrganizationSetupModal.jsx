import React from 'react'
import './OrganizationSetupModal.css'

const OrganizationSetupModal = ({ isOpen, onClose, onCreateOrganization }) => {
  if (!isOpen) return null

  return (
    <div className="org-setup-modal-overlay" onClick={onClose}>
      <div className="org-setup-modal" onClick={(e) => e.stopPropagation()}>
        <button className="org-setup-modal-close" onClick={onClose}>
          ×
        </button>
        
        <div className="org-setup-header">
          <div className="org-setup-logo">📊</div>
          <h2>Bienvenido a Report Tuner</h2>
          <p className="org-setup-subtitle">
            Creá tu espacio de trabajo para comenzar
          </p>
        </div>

        <div className="org-setup-options">
          <button 
            className="org-setup-option-btn org-setup-option-primary"
            onClick={() => {
              onCreateOrganization()
              onClose()
            }}
          >
            <div className="org-setup-option-icon">➕</div>
            <div className="org-setup-option-content">
              <h3>Crear nueva organización</h3>
              <p>Empeza tu propio espacio de trabajo desde cero</p>
            </div>
          </button>
        </div>
      </div>
    </div>
  )
}

export default OrganizationSetupModal


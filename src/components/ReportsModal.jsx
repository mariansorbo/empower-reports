import React, { useState, useEffect } from 'react'
import { useAuth } from '../contexts/AuthContext'
import { listReports, deleteReports } from '../services/azureStorageService'
import './ReportsModal.css'

const ReportsModal = ({ isOpen, onClose }) => {
  const { user } = useAuth()
  const [reports, setReports] = useState([])
  const [selectedReports, setSelectedReports] = useState(new Set())
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false)
  const [pendingDelete, setPendingDelete] = useState(new Set())
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [deleteMessage, setDeleteMessage] = useState('')

  // Cargar reportes al abrir el modal
  useEffect(() => {
    if (isOpen) {
      loadReports()
    }
  }, [isOpen])

  const loadReports = async () => {
    setLoading(true)
    setError('')
    try {
      const data = await listReports()
      setReports(data)
    } catch (err) {
      console.error('Error loading reports:', err)
      setError(`Failed to load reports: ${err.message}`)
    } finally {
      setLoading(false)
    }
  }

  const handleCheckboxChange = (reportId) => {
    setSelectedReports(prev => {
      const newSet = new Set(prev)
      if (newSet.has(reportId)) {
        newSet.delete(reportId)
      } else {
        newSet.add(reportId)
      }
      return newSet
    })
  }

  const handleDeleteClick = () => {
    if (selectedReports.size === 0) return
    
    // Save IDs to be deleted
    setPendingDelete(new Set(selectedReports))
    // Show confirmation modal
    setShowDeleteConfirm(true)
  }

  const handleConfirmDelete = async () => {
    setLoading(true)
    setDeleteMessage('')
    try {
      // Convertir Set de IDs (nombres de blob) a array
      const blobNames = Array.from(pendingDelete)
      const results = await deleteReports(blobNames)
      
      if (results.failed.length === 0) {
        setDeleteMessage(`✅ Successfully deleted ${results.success.length} report(s)`)
        // Recargar la lista de reportes
        await loadReports()
      } else {
        setDeleteMessage(
          `⚠️ Deleted ${results.success.length} report(s). Failed: ${results.failed.length}`
        )
      }
      
      setSelectedReports(new Set())
      setPendingDelete(new Set())
      setShowDeleteConfirm(false)
      
      // Limpiar mensaje después de 3 segundos
      setTimeout(() => setDeleteMessage(''), 3000)
    } catch (err) {
      console.error('Error deleting reports:', err)
      setError(`Failed to delete reports: ${err.message}`)
      setShowDeleteConfirm(false)
    } finally {
      setLoading(false)
    }
  }

  const handleCancelDelete = () => {
    setPendingDelete(new Set())
    setShowDeleteConfirm(false)
  }

  if (!isOpen) return null

  const selectedCount = selectedReports.size

  return (
    <>
      <div className="reports-modal-overlay" onClick={onClose}>
        <div className="reports-modal" onClick={(e) => e.stopPropagation()}>
          <button 
            className="reports-modal-close" 
            onClick={onClose}
            type="button"
          >
            ×
          </button>
          
          <div className="reports-modal-header">
            <div className="reports-header-top">
              <div>
                <h2>Reports</h2>
                <p className="reports-modal-subtitle">
                  Manage your uploaded reports
                </p>
              </div>
              {selectedCount > 0 && (
                <button 
                  className="btn-delete-top"
                  onClick={handleDeleteClick}
                  title={`Delete ${selectedCount} report${selectedCount > 1 ? 's' : ''}`}
                >
                  🗑️ Delete ({selectedCount})
                </button>
              )}
            </div>
          </div>

          <div className="reports-modal-content">
            {/* Mensajes de error y éxito */}
            {error && (
              <div className="error-message" style={{ marginBottom: '16px' }}>
                ❌ {error}
              </div>
            )}
            {deleteMessage && (
              <div className="success-message" style={{ marginBottom: '16px' }}>
                {deleteMessage}
              </div>
            )}

            {loading ? (
              <div className="reports-loading">
                <p>Loading reports...</p>
              </div>
            ) : reports.length === 0 ? (
              <div className="reports-empty">
                <p>No reports available</p>
              </div>
            ) : (
              <div className="reports-list">
                {reports.map(report => (
                  <div key={report.id} className="report-item">
                    <label className="report-checkbox">
                      <input
                        type="checkbox"
                        checked={selectedReports.has(report.id)}
                        onChange={() => handleCheckboxChange(report.id)}
                      />
                    </label>
                    <div className="report-info">
                      <div className="report-name">{report.name}</div>
                      <div className="report-meta">
                        <span className="report-date">📅 {report.date}</span>
                        <span className="report-size">💾 {report.size}</span>
                        <span className="report-uploader">👤 {report.uploader}</span>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Delete confirmation modal */}
      {showDeleteConfirm && (
        <div className="delete-confirm-overlay" onClick={handleCancelDelete}>
          <div className="delete-confirm-modal" onClick={(e) => e.stopPropagation()}>
            <div className="delete-confirm-header">
              <h3>Delete reports</h3>
            </div>
            <div className="delete-confirm-content">
              <p>
                Are you sure you want to delete {pendingDelete.size} report{pendingDelete.size > 1 ? 's' : ''}?
              </p>
              <p className="delete-confirm-warning">
                This action cannot be undone.
              </p>
            </div>
            <div className="delete-confirm-actions">
              <button
                type="button"
                className="btn-cancel"
                onClick={handleCancelDelete}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn-save-delete"
                onClick={handleConfirmDelete}
                disabled={loading}
              >
                {loading ? 'Deleting...' : 'Delete'}
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  )
}

export default ReportsModal

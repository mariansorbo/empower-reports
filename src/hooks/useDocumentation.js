/**
 * Hook para gestionar la documentación de la organización
 * Habilita el botón "Ver documentación" si la org tiene URL configurada
 */

import { useState, useEffect } from 'react';
import { useOrganization } from '../contexts/OrganizationContext';
import { useAuth } from '../contexts/AuthContext';

export function useDocumentation() {
  const { currentOrganization } = useOrganization();
  const { isAuthenticated } = useAuth();
  const [documentationUrl, setDocumentationUrl] = useState(null);
  const [hasDocumentation, setHasDocumentation] = useState(false);
  const [loading, setLoading] = useState(true);
  
  useEffect(() => {
    if (isAuthenticated && currentOrganization?.id) {
      fetchDocumentationUrl();
    } else {
      setDocumentationUrl(null);
      setHasDocumentation(false);
      setLoading(false);
    }
  }, [isAuthenticated, currentOrganization]);
  
  const fetchDocumentationUrl = async () => {
    setLoading(true);
    try {
      // En un sistema real, esto sería una llamada a la API
      // const response = await fetch(`/api/organizations/${currentOrganization.id}/documentation`);
      // const data = await response.json();
      
      // Por ahora, simular desde localStorage de la organización
      const orgData = localStorage.getItem(`org_${currentOrganization.id}`);
      if (orgData) {
        const parsed = JSON.parse(orgData);
        setDocumentationUrl(parsed.documentation_url || null);
        setHasDocumentation(!!parsed.documentation_url);
      }
    } catch (error) {
      console.error('Error fetching documentation URL:', error);
      setDocumentationUrl(null);
      setHasDocumentation(false);
    } finally {
      setLoading(false);
    }
  };
  
  const openDocumentation = () => {
    if (documentationUrl) {
      window.open(documentationUrl, '_blank', 'noopener,noreferrer');
    }
  };
  
  return {
    documentationUrl,
    hasDocumentation,
    loading,
    openDocumentation
  };
}







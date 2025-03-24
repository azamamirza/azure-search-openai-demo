// src/hooks/useAuthContext.ts
import { useState, useEffect, createContext, useContext } from 'react';
import { useMsal } from '@azure/msal-react';
import { getToken } from '../authConfig';

interface AuthContextType {
  idToken?: string;
  isAuthenticated: boolean;
  isLoading: boolean;
}

const AuthContext = createContext<AuthContextType>({
  idToken: undefined,
  isAuthenticated: false,
  isLoading: true
});

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { instance } = useMsal();
  const [authState, setAuthState] = useState<AuthContextType>({
    idToken: undefined,
    isAuthenticated: false,
    isLoading: true
  });

  useEffect(() => {
    const fetchToken = async () => {
      try {
        const token = await getToken(instance);
        
        setAuthState({
          idToken: token,
          isAuthenticated: !!token,
          isLoading: false
        });
      } catch (error) {
        console.error('Failed to get token:', error);
        setAuthState({
          idToken: undefined,
          isAuthenticated: false,
          isLoading: false
        });
      }
    };

    fetchToken();
  }, [instance]);

  return (
    <AuthContext.Provider value={authState}>
      {children}
    </AuthContext.Provider>
  );
};

export const useAuthContext = () => useContext(AuthContext);
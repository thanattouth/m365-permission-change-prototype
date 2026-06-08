import React from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider, useMsal } from "@azure/msal-react";
import { msalConfig, loginRequest } from "./authConfig";
import "./App.css";

const msalInstance = new PublicClientApplication(msalConfig);

function MainApp() {
  const { instance, accounts } = useMsal();
  const [isLoggingIn, setIsLoggingIn] = React.useState(false);
  const [userData, setUserData] = React.useState(null);
  const [isLoadingData, setIsLoadingData] = React.useState(false);
  const [error, setError] = React.useState(null);

  const handleLogin = async () => {
    if (isLoggingIn) return;
    setIsLoggingIn(true);
    setError(null);
    try {
      await instance.loginRedirect(loginRequest);
    } catch (error) {
      console.error("Login error:", error);
      if (error.errorCode !== "block_nested_popups") {
        setError(error.message || "Login failed");
      }
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    await instance.logoutPopup();
    setUserData(null);
    setError(null);
  };

  const callGraph = async () => {
    if (isLoadingData) return;
    setIsLoadingData(true);
    setError(null);
    try {
      const response = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });

      const res = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: {
          Authorization: `Bearer ${response.accessToken}`,
        },
      });

      if (!res.ok) {
        throw new Error(`API returned ${res.status}`);
      }

      const data = await res.json();
      setUserData(data);
    } catch (error) {
      console.error("Graph API error:", error);
      setError(error.message || "Failed to fetch user data");
    } finally {
      setIsLoadingData(false);
    }
  };

  return (
    <div className="container">
      <div className="card">
        {!accounts.length ? (
          <div className="auth-section">
            <div className="header">
              <h1>M365</h1>
              <p className="subtitle">Permission Manager</p>
            </div>
            <p className="description">Sign in to manage your Microsoft 365 permissions securely.</p>
            {error && <div className="error-message">{error}</div>}
            <button 
              className="btn btn-primary" 
              onClick={handleLogin} 
              disabled={isLoggingIn}
            >
              {isLoggingIn ? (
                <>
                  <span className="spinner"></span>
                  Signing in...
                </>
              ) : (
                "Sign in with Microsoft"
              )}
            </button>
          </div>
        ) : (
          <div className="user-section">
            <div className="header">
              <h1>Dashboard</h1>
              <p className="subtitle">Welcome back</p>
            </div>
            
            {!userData ? (
              <>
                <p className="description">Load your profile information</p>
                {error && <div className="error-message">{error}</div>}
                <button 
                  className="btn btn-primary" 
                  onClick={callGraph}
                  disabled={isLoadingData}
                >
                  {isLoadingData ? (
                    <>
                      <span className="spinner"></span>
                      Loading...
                    </>
                  ) : (
                    "Load Profile"
                  )}
                </button>
              </>
            ) : (
              <div className="user-info">
                <div className="user-avatar">
                  {userData.displayName?.charAt(0).toUpperCase() || "U"}
                </div>
                <h2>{userData.displayName}</h2>
                <p className="email">{userData.mail || userData.userPrincipalName}</p>
                <div className="user-details">
                  {userData.jobTitle && (
                    <div className="detail-item">
                      <span className="label">Job Title:</span>
                      <span className="value">{userData.jobTitle}</span>
                    </div>
                  )}
                  {userData.officeLocation && (
                    <div className="detail-item">
                      <span className="label">Location:</span>
                      <span className="value">{userData.officeLocation}</span>
                    </div>
                  )}
                </div>
                <button 
                  className="btn btn-secondary"
                  onClick={callGraph}
                  disabled={isLoadingData}
                >
                  {isLoadingData ? "Refreshing..." : "Refresh"}
                </button>
              </div>
            )}
            
            <button 
              className="btn btn-outline"
              onClick={handleLogout}
            >
              Sign Out
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

export default function App() {
  return (
    <MsalProvider instance={msalInstance}>
      <MainApp />
    </MsalProvider>
  );
}
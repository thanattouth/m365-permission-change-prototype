import React, { useEffect, useState, useRef } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider, useMsal } from "@azure/msal-react";
import { msalConfig, loginRequest } from "./authConfig";
import ItemsList from "./components/ItemsList";
import "./App.css";

const msalInstance = new PublicClientApplication(msalConfig);

function MainApp() {
  const { instance, accounts } = useMsal();
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [userData, setUserData] = useState(null);
  const [isLoadingProfile, setIsLoadingProfile] = useState(false);
  const [error, setError] = useState(null);
  const [showDropdown, setShowDropdown] = useState(false);
  const dropdownRef = useRef(null);

  // Auto-fetch profile once logged in
  useEffect(() => {
    const fetchProfile = async () => {
      if (accounts.length > 0 && !userData && !isLoadingProfile) {
        setIsLoadingProfile(true);
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
        } catch (err) {
          console.error("Microsoft Graph profile error:", err);
          setError(err.message || "Failed to load profile details in background");
        } finally {
          setIsLoadingProfile(false);
        }
      }
    };

    fetchProfile();
  }, [accounts, userData, instance, isLoadingProfile]);

  // Click outside listener to close the profile dropdown
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setShowDropdown(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);

  const handleLogin = async () => {
    if (isLoggingIn) return;
    setIsLoggingIn(true);
    setError(null);
    try {
      await instance.loginRedirect(loginRequest);
    } catch (err) {
      console.error("Login error:", err);
      if (err.errorCode !== "block_nested_popups") {
        setError(err.message || "Login failed");
      }
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    await instance.logoutPopup();
    setUserData(null);
    setError(null);
    setShowDropdown(false);
  };

  const isLoggedIn = accounts.length > 0;

  return (
    <div className={`container ${isLoggedIn ? "full-layout" : ""}`}>
      {!isLoggedIn ? (
        <div className="card">
          <div className="auth-section">
            <div className="header" style={{ textAlign: "center" }}>
              <span className="shield-icon">🛡️</span>
              <h1>M365</h1>
              <p className="subtitle">Permission Manager</p>
            </div>
            <p className="description" style={{ textAlign: "center", marginBottom: "24px" }}>
              Sign in to manage Microsoft 365 SharePoint access across site contents, folders, and files.
            </p>
            {error && <div className="error-message" style={{ marginBottom: "16px" }}>{error}</div>}
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
        </div>
      ) : (
        /* Large Portal Shell Container matching the mockup layout */
        <div className="portal-shell">
          
          {/* Main Top Header Navbar */}
          <header className="portal-navbar">
            <div className="brand-section">
              <span className="brand-logo">🛠️</span>
              <span>M365 Portal</span>
            </div>
            
            <div className="nav-user-actions" ref={dropdownRef}>
              <span className="site-indicator">
                🌐 Connected to SharePoint Site
              </span>
              
              <button 
                className="avatar-btn" 
                onClick={() => setShowDropdown(!showDropdown)}
                title={userData ? userData.displayName : "Profile Info"}
              >
                {userData && userData.displayName 
                  ? userData.displayName.charAt(0).toUpperCase() 
                  : (accounts[0].name ? accounts[0].name.charAt(0).toUpperCase() : "U")}
              </button>
              
              {/* Profile details dropdown */}
              {showDropdown && (
                <div className="profile-dropdown">
                  <div className="dropdown-user-info">
                    <p className="dropdown-user-name">
                      {userData ? userData.displayName : accounts[0].name || "User"}
                    </p>
                    <p className="dropdown-user-email">
                      {userData ? userData.mail || userData.userPrincipalName : accounts[0].username}
                    </p>
                  </div>
                  
                  {userData && (userData.jobTitle || userData.officeLocation) && (
                    <div className="dropdown-details">
                      {userData.jobTitle && (
                        <div className="dropdown-detail-item">
                          <span className="label">Job Title</span>
                          <span className="value">{userData.jobTitle}</span>
                        </div>
                      )}
                      {userData.officeLocation && (
                        <div className="dropdown-detail-item">
                          <span className="label">Location</span>
                          <span className="value">{userData.officeLocation}</span>
                        </div>
                      )}
                    </div>
                  )}
                  
                  <button 
                    className="btn btn-secondary btn-sm" 
                    onClick={handleLogout}
                    style={{ width: "100%", marginTop: "4px" }}
                  >
                    Sign Out
                  </button>
                </div>
              )}
            </div>
          </header>

          {/* Main Items Listing Page Content */}
          <main className="portal-content">
            <ItemsList instance={instance} accounts={accounts} />
          </main>
          
        </div>
      )}
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

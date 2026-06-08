import React from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider, useMsal } from "@azure/msal-react";
import { msalConfig, loginRequest } from "./authConfig";

const msalInstance = new PublicClientApplication(msalConfig);

function MainApp() {
  const { instance, accounts } = useMsal();

  const handleLogin = () => {
    instance.loginPopup(loginRequest);
  };

  const callGraph = async () => {
    const response = await instance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0],
    });

    const res = await fetch("https://graph.microsoft.com/v1.0/me", {
      headers: {
        Authorization: `Bearer ${response.accessToken}`,
      },
    });

    const data = await res.json();
    console.log(data);
    alert(`Hello ${data.displayName}`);
  };

  return (
    <div style={{ padding: 20 }}>
      {!accounts.length ? (
        <button onClick={handleLogin}>Login</button>
      ) : (
        <button onClick={callGraph}>Call Graph</button>
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
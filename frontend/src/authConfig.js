// SharePoint Site Configuration
export const sharePointConfig = {
  siteUrl: "https://devm365th.sharepoint.com/sites/DocumentManagement",
  // Example: "https://tenant.sharepoint.com/sites/yoursitename"
};

export const msalConfig = {
  auth: {
    clientId: "5e597123-d103-4c16-a7f3-cb99da285561",
    authority: "https://login.microsoftonline.com/0f3101bc-add7-42aa-a041-4b5648c7bacf",
    redirectUri: window.location.origin
  }
};

export const loginRequest = {
  scopes: [
    "User.Read",
    "Sites.ReadWrite.All",      // Read/Write SharePoint sites and content
    "Files.ReadWrite.All",       // Read/Write files in SharePoint
    "Directory.ReadWrite.All"    // Manage users/groups for permissions
  ]
};

// For getting access token with specific scopes
export const tokenRequest = {
  scopes: [
    "Sites.ReadWrite.All",
    "Files.ReadWrite.All",
    "Directory.ReadWrite.All"
  ]
};

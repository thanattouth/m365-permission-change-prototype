import { sharePointConfig } from "../authConfig";

/**
 * Get SharePoint site ID from URL
 */
export const getSiteId = async (instance) => {
  try {
    const token = await getAccessToken(instance);
    const siteUrl = sharePointConfig.siteUrl;
    
    // Extract hostname and path from SharePoint URL
    const url = new URL(siteUrl);
    const hostname = url.hostname;
    const sitePath = url.pathname.replace(/\/$/, ""); // Remove trailing slash
    
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`,
      {
        headers: { Authorization: `Bearer ${token}` },
      }
    );
    
    if (!response.ok) {
      throw new Error(`Failed to get site ID: ${response.status}`);
    }
    
    const data = await response.json();
    return data.id;
  } catch (error) {
    console.error("Error getting site ID:", error);
    throw error;
  }
};

/**
 * Get access token for API calls
 */
export const getAccessToken = async (instance) => {
  const response = await instance.acquireTokenSilent({
    scopes: [
      "Sites.ReadWrite.All",
      "Files.ReadWrite.All",
      "Directory.ReadWrite.All"
    ]
  });
  return response.accessToken;
};

/**
 * Get all items (files and folders) from root of SharePoint site
 */
export const getItemsList = async (instance) => {
  try {
    const siteId = await getSiteId(instance);
    const token = await getAccessToken(instance);
    
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children?$select=id,name,size,webUrl,folder,file,lastModifiedDateTime,lastModifiedBy`,
      {
        headers: { Authorization: `Bearer ${token}` },
      }
    );
    
    if (!response.ok) {
      throw new Error(`Failed to fetch items: ${response.status}`);
    }
    
    const data = await response.json();
    return data.value || [];
  } catch (error) {
    console.error("Error fetching items:", error);
    throw error;
  }
};

/**
 * Get permissions of a specific item
 */
export const getItemPermissions = async (instance, itemId) => {
  try {
    const siteId = await getSiteId(instance);
    const token = await getAccessToken(instance);
    
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/permissions?$select=id,grantedTo,roles`,
      {
        headers: { Authorization: `Bearer ${token}` },
      }
    );
    
    if (!response.ok) {
      throw new Error(`Failed to fetch permissions: ${response.status}`);
    }
    
    const data = await response.json();
    return data.value || [];
  } catch (error) {
    console.error("Error fetching permissions:", error);
    throw error;
  }
};

/**
 * Add permission to an item
 * @param {string} itemId - Item ID
 * @param {string} userEmail - User email or group email
 * @param {string} role - 'read', 'write', 'owner'
 */
export const addItemPermission = async (instance, itemId, userEmail, role = "read") => {
  try {
    const siteId = await getSiteId(instance);
    const token = await getAccessToken(instance);
    
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/invite`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          recipients: [{ email: userEmail }],
          requireSignIn: true,
          sendInvitation: true,
          roles: [role],
          message: `You have been granted ${role} access to this file/folder.`
        }),
      }
    );
    
    if (!response.ok) {
      throw new Error(`Failed to add permission: ${response.status}`);
    }
    
    const data = await response.json();
    return data;
  } catch (error) {
    console.error("Error adding permission:", error);
    throw error;
  }
};

/**
 * Remove permission from an item
 * @param {string} itemId - Item ID
 * @param {string} permissionId - Permission ID
 */
export const removeItemPermission = async (instance, itemId, permissionId) => {
  try {
    const siteId = await getSiteId(instance);
    const token = await getAccessToken(instance);
    
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/permissions/${permissionId}`,
      {
        method: "DELETE",
        headers: { Authorization: `Bearer ${token}` },
      }
    );
    
    if (!response.ok) {
      throw new Error(`Failed to remove permission: ${response.status}`);
    }
    
    return { success: true };
  } catch (error) {
    console.error("Error removing permission:", error);
    throw error;
  }
};

/**
 * Search for users in Azure AD
 */
export const searchUsers = async (instance, searchTerm) => {
  try {
    const token = await getAccessToken(instance);
    
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${searchTerm}') or startswith(userPrincipalName,'${searchTerm}')&$select=id,displayName,userPrincipalName,mail`,
      {
        headers: { Authorization: `Bearer ${token}` },
      }
    );
    
    if (!response.ok) {
      throw new Error(`Failed to search users: ${response.status}`);
    }
    
    const data = await response.json();
    return data.value || [];
  } catch (error) {
    console.error("Error searching users:", error);
    throw error;
  }
};

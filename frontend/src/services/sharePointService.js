import { sharePointConfig } from "../authConfig";

/**
 * Get access token for API calls
 */
export const getAccessToken = async (instance, account) => {
  const response = await instance.acquireTokenSilent({
    scopes: [
      "Sites.ReadWrite.All",
      "Files.ReadWrite.All",
      "Directory.ReadWrite.All"
    ],
    account: account
  });
  return response.accessToken;
};

/**
 * Get SharePoint site ID from URL
 */
export const getSiteId = async (instance, account) => {
  try {
    const token = await getAccessToken(instance, account);
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
 * Get all items (files and folders) from root of SharePoint site
 */
export const getItemsList = async (instance, account, folderId = "root") => {
  try {
    const siteId = await getSiteId(instance, account);
    const token = await getAccessToken(instance, account);
    
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${folderId}/children?$select=id,name,size,webUrl,folder,file,lastModifiedDateTime,lastModifiedBy`,
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
export const getItemPermissions = async (instance, account, itemId) => {
  try {
    const siteId = await getSiteId(instance, account);
    const token = await getAccessToken(instance, account);
    
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
 * Map UI role names to Microsoft Graph API role values
 * The /invite endpoint accepts: 'read', 'write'
 * SharePoint-specific roles use the 'sp.' prefix
 */
const mapRoleToApiRole = (role) => {
  const roleMap = {
    read: "read",
    contribute: "sp.contribute",
    write: "write",
  };
  return roleMap[role] || role;
};

/**
 * Add permission to an item
 * @param {string} itemId - Item ID
 * @param {string} userEmail - User email or group email
 * @param {string} role - 'read', 'contribute', 'write'
 */
export const addItemPermission = async (instance, account, itemId, userEmail, role = "read") => {
  try {
    const siteId = await getSiteId(instance, account);
    const token = await getAccessToken(instance, account);
    
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
          roles: [mapRoleToApiRole(role)],
          message: `You have been granted access to this file/folder.`
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
export const removeItemPermission = async (instance, account, itemId, permissionId) => {
  try {
    const siteId = await getSiteId(instance, account);
    const token = await getAccessToken(instance, account);
    
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
 * Update permission role for an item
 * Updates the role for an existing permission
 */
export const updateItemPermission = async (instance, account, itemId, permissionId, newRole) => {
  try {
    const siteId = await getSiteId(instance, account);
    const token = await getAccessToken(instance, account);
    
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/permissions/${permissionId}`,
      {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          roles: [mapRoleToApiRole(newRole)],
        }),
      }
    );
    
    if (!response.ok) {
      throw new Error(`Failed to update permission: ${response.status}`);
    }
    
    const data = await response.json();
    return data;
  } catch (error) {
    console.error("Error updating permission:", error);
    throw error;
  }
};

/**
 * Search for users in Azure AD
 */
export const searchUsers = async (instance, account, searchTerm) => {
  try {
    const token = await getAccessToken(instance, account);
    
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

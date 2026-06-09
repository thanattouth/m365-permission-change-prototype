import { sharePointConfig } from "../authConfig";

/**
 * Get access token for Microsoft Graph API calls
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
 * Get access token for SharePoint REST API.
 * Uses the AllSites.FullControl scope (consented at login via authConfig.js).
 * Silent-only — no popup needed since scope is pre-consented.
 */
export const getSharePointToken = async (instance, account) => {
  const siteHostname = new URL(sharePointConfig.siteUrl).hostname;
  const response = await instance.acquireTokenSilent({
    scopes: [`https://${siteHostname}/AllSites.FullControl`],
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
    const permissions = data.value || [];

    /* Commented out Restrict Editor (Contribute) role merging temporarily
    try {
      const itemRes = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}?$select=id,webUrl,folder,file`,
        {
          headers: { Authorization: `Bearer ${token}` },
        }
      );
      if (itemRes.ok) {
        const itemData = await itemRes.json();
        const webUrl = itemData.webUrl;
        const isFolder = !!itemData.folder;
        
        if (webUrl) {
          const spToken = await getSharePointToken(instance, account);
          const siteUrl = sharePointConfig.siteUrl.replace(/\/$/, "");
          
          const url = new URL(webUrl);
          const itemServerRelativeUrl = decodeURIComponent(url.pathname);
          
          const encodedUrl = `'${encodeURIComponent(itemServerRelativeUrl)}'`;
          const itemBaseUrl = isFolder
            ? `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(${encodedUrl})/ListItemAllFields`
            : `${siteUrl}/_api/web/GetFileByServerRelativeUrl(${encodedUrl})/ListItemAllFields`;
            
          const spRes = await fetch(
            `${itemBaseUrl}/roleassignments?$expand=Member,RoleDefinitionBindings`,
            {
              headers: {
                Authorization: `Bearer ${spToken}`,
                Accept: "application/json;odata=verbose",
              }
            }
          );
          
          if (spRes.ok) {
            const spData = await spRes.json();
            const roleAssignments = spData.d?.results || [];
            
            permissions.forEach(perm => {
              const matchingAssignment = roleAssignments.find(assign => {
                const member = assign.Member;
                if (!member) return false;
                
                if (perm.grantedTo?.user) {
                  const user = perm.grantedTo.user;
                  const email = (user.email || "").toLowerCase();
                  const upn = (user.userPrincipalName || "").toLowerCase();
                  const displayName = (user.displayName || "").toLowerCase();
                  
                  const spEmail = (member.Email || "").toLowerCase();
                  const spLogin = (member.LoginName || "").toLowerCase();
                  const spTitle = (member.Title || "").toLowerCase();
                  
                  return (
                    (email && spEmail === email) ||
                    (email && spLogin.includes(email)) ||
                    (upn && spLogin.includes(upn)) ||
                    (displayName && spTitle === displayName)
                  );
                }
                
                if (perm.grantedTo?.group) {
                  const group = perm.grantedTo.group;
                  const displayName = (group.displayName || "").toLowerCase();
                  const spTitle = (member.Title || "").toLowerCase();
                  
                  return displayName && spTitle === displayName;
                }
                
                return false;
              });
              
              if (matchingAssignment && matchingAssignment.RoleDefinitionBindings?.results) {
                const spRoles = matchingAssignment.RoleDefinitionBindings.results.map(r => r.Name);
                const mappedRoles = [];
                
                if (spRoles.includes("Full Control")) {
                  mappedRoles.push("owner");
                } else if (spRoles.includes("Edit")) {
                  mappedRoles.push("write");
                } else if (spRoles.includes("Contribute")) {
                  mappedRoles.push("contribute");
                } else if (spRoles.includes("Read")) {
                  mappedRoles.push("read");
                }
                
                if (mappedRoles.length > 0) {
                  perm.roles = mappedRoles;
                }
              }
            });
          } else {
            console.warn(`SharePoint REST role assignment fetch failed with status: ${spRes.status}`);
          }
        }
      }
    } catch (spError) {
      console.error("Error fetching or merging SharePoint REST roles:", spError);
    }
    */

    return permissions;
  } catch (error) {
    console.error("Error fetching permissions:", error);
    throw error;
  }
};

/**
 * Map UI role names to Microsoft Graph API role values.
 * NOTE: Graph API /invite only supports 'read' and 'write'.
 * 'contribute' is handled separately via SharePoint REST API.
 */
const mapRoleToGraphRole = (role) => {
  const roleMap = {
    read: "read",
    write: "write",
  };
  return roleMap[role] || "read";
};

/**
 * Add 'Contribute' (Restrict Editor) permission via SharePoint REST API
 * This bypasses the Graph API limitation (only read/write supported)
 *
 * Flow:
 *  1. Get SharePoint REST token (different scope from Graph)
 *  2. Look up SharePoint user ID by email
 *  3. Find the 'Contribute' role definition ID on this site
 *  4. Break item permission inheritance (if needed) and add role assignment
 *
 * @param {string} itemServerRelativeUrl - e.g. /sites/DocumentManagement/Shared Documents/file.docx
 * @param {string} userEmail
 * @param {boolean} isFolder - true if item is a folder, false if file
 * @param {object} instance - MSAL instance
 * @param {object} account - MSAL account
 */
export const addContributePermissionViaRestApi = async (instance, account, itemServerRelativeUrl, userEmail, isFolder = false) => {
  const spToken = await getSharePointToken(instance, account);
  const siteUrl = sharePointConfig.siteUrl.replace(/\/$/, "");

  const spHeaders = {
    Authorization: `Bearer ${spToken}`,
    Accept: "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
  };

  // Choose correct SharePoint REST endpoint based on item type
  // Files: GetFileByServerRelativeUrl | Folders: GetFolderByServerRelativeUrl
  const encodedUrl = `'${encodeURIComponent(itemServerRelativeUrl)}'`;
  const itemBaseUrl = isFolder
    ? `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(${encodedUrl})/ListItemAllFields`
    : `${siteUrl}/_api/web/GetFileByServerRelativeUrl(${encodedUrl})/ListItemAllFields`;

  // Step 1: Ensure unique permissions on the item (break inheritance)
  const breakRes = await fetch(
    `${itemBaseUrl}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`,
    { method: "POST", headers: spHeaders }
  );
  if (!breakRes.ok) {
    throw new Error(`SharePoint REST: Failed to break permission inheritance (${breakRes.status}). Item may not exist at: ${itemServerRelativeUrl}`);
  }

  // Step 2: Get user's SharePoint principal ID
  const userRes = await fetch(
    `${siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(userEmail)}')`,
    { headers: spHeaders }
  );
  if (!userRes.ok) {
    throw new Error(`SharePoint REST: Failed to get user by email (${userRes.status}). Make sure the user exists in SharePoint.`);
  }
  const userData = await userRes.json();
  const principalId = userData.d?.Id;
  if (!principalId) throw new Error("SharePoint REST: Could not find user principal ID");

  // Step 3: Find the 'Contribute' role definition ID
  const rolesRes = await fetch(
    `${siteUrl}/_api/web/roledefinitions`,
    { headers: spHeaders }
  );
  if (!rolesRes.ok) throw new Error(`SharePoint REST: Failed to get role definitions (${rolesRes.status})`);
  const rolesData = await rolesRes.json();
  const contributeRole = (rolesData.d?.results || []).find(r => r.Name === "Contribute");
  if (!contributeRole) throw new Error("SharePoint REST: Could not find 'Contribute' role definition on this site");
  const roleDefId = contributeRole.Id;

  // Step 4: Assign the Contribute role to the user on this item
  const assignRes = await fetch(
    `${itemBaseUrl}/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleDefId})`,
    { method: "POST", headers: spHeaders }
  );
  if (!assignRes.ok) {
    throw new Error(`SharePoint REST: Failed to assign Contribute role (${assignRes.status})`);
  }

  return { success: true, method: "SharePoint REST API", role: "Contribute", userId: principalId, roleDefId };
};

/**
 * Add permission to an item
 * @param {string} itemId - Item ID (Graph)
 * @param {string} userEmail - User email
 * @param {string} role - 'read' | 'contribute' | 'write'
 * @param {string|null} itemServerRelativeUrl - Server-relative URL (for SharePoint REST, used for Contribute)
 * @param {boolean} isFolder - true if item is a folder (affects SharePoint REST endpoint)
 */
export const addItemPermission = async (instance, account, itemId, userEmail, role = "read", itemServerRelativeUrl = null, isFolder = false) => {
  // 'contribute' uses SharePoint REST API (Graph API does not support it)
  if (role === "contribute") {
    if (!itemServerRelativeUrl) {
      throw new Error("Restrict Editor requires itemServerRelativeUrl to use SharePoint REST API");
    }
    return addContributePermissionViaRestApi(instance, account, itemServerRelativeUrl, userEmail, isFolder);
  }

  // 'read' and 'write' use Microsoft Graph API
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
          roles: [mapRoleToGraphRole(role)],
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
 * NOTE: Graph API only supports 'read' and 'write' for PATCH.
 * 'contribute' update path is not yet supported (would need remove + re-add via REST API).
 */
export const updateItemPermission = async (instance, account, itemId, permissionId, newRole) => {
  if (newRole === "contribute") {
    throw new Error("Updating to Restrict Editor is not supported.");
  }
  try {
    const siteId = await getSiteId(instance, account);
    const token = await getAccessToken(instance, account);
    
    // Standard Graph API PATCH
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/permissions/${permissionId}`,
      {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          roles: [mapRoleToGraphRole(newRole)],
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

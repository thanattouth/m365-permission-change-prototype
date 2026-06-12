import React, { useState, useEffect } from "react";
import {
  getItemPermissions,
  addItemPermission,
  removeItemPermission,
  updateItemPermission,
  searchUsers,
} from "../services/sharePointService";
import Icon from "./Icon";
// File Classification is parked for this prototype phase.
// The current customer use case only needs admin-driven permission management
// across site contents, folders, and files.
// import FileClassification from "./FileClassification";
import "./ItemPermissions.css";

function ItemPermissions({ instance, item, onPermissionChanged, onClose, account, selectedSite }) {
  const [permissions, setPermissions] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [notice, setNotice] = useState(null);
  const [userSearchResults, setUserSearchResults] = useState([]);
  const [searchInput, setSearchInput] = useState("");
  const [isSearching, setIsSearching] = useState(false);
  const [userRoleMap, setUserRoleMap] = useState({});
  const [addingPermissionUserId, setAddingPermissionUserId] = useState(null);
  // const [classification, setClassification] = useState(null);
  const [editingPermissionId, setEditingPermissionId] = useState(null);
  const [editingNewRole, setEditingNewRole] = useState(null);
  const [isUpdatingPermission, setIsUpdatingPermission] = useState(false);

  const loadPermissions = React.useCallback(async () => {
    setIsLoading(true);
    setError(null);
    try {
      const perms = await getItemPermissions(instance, account, item.id, !!item.isLibrary, item.driveId, selectedSite?.url);
      setPermissions(perms);
    } catch (err) {
      setError(err.message || "Failed to load permissions");
    } finally {
      setIsLoading(false);
    }
  }, [instance, account, item.id, item.isLibrary, item.driveId, selectedSite]);

  useEffect(() => {
    loadPermissions();
  }, [item, loadPermissions]);

  const handleSearchUsers = async (e) => {
    const value = e.target.value;
    setSearchInput(value);

    if (value.length < 2) {
      setUserSearchResults([]);
      return;
    }

    setIsSearching(true);
    try {
      const results = await searchUsers(instance, account, value);
      setUserSearchResults(results);
    } catch (err) {
      console.error("Search error:", err);
      setUserSearchResults([]);
    } finally {
      setIsSearching(false);
    }
  };

  const getUserRole = (userId) => userRoleMap[userId] || "read";
  const setUserRole = (userId, role) =>
    setUserRoleMap((prev) => ({ ...prev, [userId]: role }));
  const isAddingPermission = addingPermissionUserId !== null;
  const isAddingUser = (userId) => addingPermissionUserId === userId;

  const handleAddPermission = async (userEmail, userId) => {
    const role = getUserRole(userId);
    setAddingPermissionUserId(userId);
    setError(null);
    setNotice(null);
    console.info("Grant access clicked", {
      userEmail,
      role,
      itemId: item.id,
      itemName: item.name,
      isLibrary: !!item.isLibrary,
      driveId: item.driveId,
      siteUrl: selectedSite?.url,
    });
    try {
      let itemServerRelativeUrl = null;
      if (item.webUrl) {
        try {
          const url = new URL(item.webUrl);
          itemServerRelativeUrl = decodeURIComponent(url.pathname);
        } catch (e) {
          console.warn("Could not parse webUrl for relative path", e);
        }
      }

      const result = await addItemPermission(
        instance, 
        account, 
        item.id, 
        userEmail, 
        role, 
        itemServerRelativeUrl, 
        !!item.folder, 
        !!item.isLibrary,
        item.driveId,
        selectedSite?.url
      );
      console.info("Grant access completed", {
        userEmail,
        role,
        result,
      });
      setSearchInput("");
      setUserSearchResults([]);
      setUserRoleMap({});
      await loadPermissions();
      setNotice(`Access request completed for ${userEmail}. If this is an external user, the invitation may appear after redemption or be blocked by tenant mail flow policy.`);
    } catch (err) {
      console.error("Grant access failed", err);
      setError(err.message || "Failed to add permission");
    } finally {
      setAddingPermissionUserId(null);
    }
  };

  const handleRemovePermission = async (permissionId) => {
    if (!window.confirm("Remove this permission?")) return;

    try {
      await removeItemPermission(instance, account, item.id, permissionId, !!item.isLibrary, item.driveId, selectedSite?.url);
      await loadPermissions();
    } catch (err) {
      setError(err.message || "Failed to remove permission");
    }
  };

  const handleSavePermissionEdit = async (permissionId, currentRole) => {
    if (editingNewRole === currentRole) {
      setEditingPermissionId(null);
      setEditingNewRole(null);
      return;
    }

    setIsUpdatingPermission(true);
    setError(null);
    try {
      await updateItemPermission(instance, account, item.id, permissionId, editingNewRole, !!item.isLibrary, item.driveId, selectedSite?.url);
      setEditingPermissionId(null);
      setEditingNewRole(null);
      await loadPermissions();
    } catch (err) {
      setError(err.message || "Failed to update permission");
    } finally {
      setIsUpdatingPermission(false);
    }
  };

  const getRoleLabel = (roles) => {
    if (!roles || roles.length === 0) return "Viewer";
    const role = roles[0].toLowerCase();
    const roleMap = {
      read: "Viewer",
      contribute: "Custom Role",
      write: "Editor",
      owner: "Owner",
    };
    return roleMap[role] || role;
  };

  const getRoleClass = (roles) => {
    if (!roles || roles.length === 0) return "view";
    const role = roles[0].toLowerCase();
    const roleMap = {
      read: "view",
      contribute: "restrict-edit",
      write: "edit",
      owner: "owner",
    };
    return roleMap[role] || "view";
  };

  const getPermissionPrincipal = (permission) => {
    const user = permission.grantedTo?.user || permission.grantedToV2?.user;
    if (user) {
      return {
        type: "user",
        name: user.displayName || user.email || user.userPrincipalName || "User",
        email: user.email || user.userPrincipalName || permission.grantedToV2?.siteUser?.loginName || "User",
        avatar: user.displayName?.charAt(0).toUpperCase() || "U",
      };
    }

    const group = permission.grantedTo?.group || permission.grantedToV2?.group;
    if (group) {
      return {
        type: "group",
        name: group.displayName || "Group",
        email: "Group",
        avatar: "G",
      };
    }

    if (permission.invitation?.email) {
      return {
        type: "invitation",
        name: permission.invitation.email,
        email: "Invitation pending",
        avatar: "E",
      };
    }

    if (permission.link) {
      return {
        type: "link",
        name: permission.link.type ? `${permission.link.type} link` : "Sharing link",
        email: permission.link.scope || "Link access",
        avatar: "L",
      };
    }

    return {
      type: "unknown",
      name: "Unknown",
      email: "Permission details unavailable",
      avatar: "?",
    };
  };

  // Helper for External Email
  const isEmailFormat = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(searchInput);
  const linkPermission = permissions.find((permission) => permission.link);

  return (
    <div className="permissions-container">
      <div className="permission-header">
        <div className="permission-title-row">
          <div className="permission-item-icon">
            <Icon
              name={item.isLibrary ? "library" : (item.folder ? "folder" : "file")}
              size={26}
            />
          </div>
          <div className="permission-item-info">
            <h2>Manage Access</h2>
            <p>{item.name}</p>
          </div>
        </div>
        <button className="permission-close-btn" onClick={onClose} aria-label="Close manage access">
          x
        </button>
      </div>

      <div className="permission-body">
        {error && (
          <div className="error-message">
            {error}
            <button onClick={() => setError(null)} className="close-error">x</button>
          </div>
        )}

        {notice && (
          <div className="info-message">
            {notice}
            <button onClick={() => setNotice(null)} className="close-info">x</button>
          </div>
        )}

        <div className="link-settings-panel">
          <div className="link-settings-icon">
            <Icon name={linkPermission ? "link" : "lock"} size={18} />
          </div>
          <div>
            <p>{linkPermission ? `${linkPermission.link.scope || "Sharing"} link is active` : "Direct access only"}</p>
            <span>
              {linkPermission
                ? `${linkPermission.link.type || "Link"} permission detected from SharePoint`
                : "No sharing link permission returned by Microsoft Graph"}
            </span>
          </div>
        </div>

        {/*
        <FileClassification
          instance={instance}
          account={account}
          item={item}
          currentClassification={classification}
          onClassificationChanged={(newId) => {
            setClassification(newId);
            loadPermissions();
          }}
        />
        */}

        <div className="add-permission-section">
        <h3>Add people, groups, or roles</h3>
        <div className="search-box">
          <Icon name="personAdd" className="search-input-icon" size={18} />
          <input
            type="text"
            placeholder="Name, email, or group"
            value={searchInput}
            onChange={handleSearchUsers}
            disabled={isAddingPermission}
            className="search-input"
          />
          {isSearching && <span className="searching">Searching...</span>}
        </div>

        {userSearchResults.length === 0 && isEmailFormat && !isSearching && (
          <div className="search-results">
            <div className="user-item">
              <div className="user-info">
                <p className="user-name">External User</p>
                <p className="user-email">{searchInput}</p>
              </div>
              <div className="user-action">
                <select
                  value={getUserRole("external")}
                  onChange={(e) => setUserRole("external", e.target.value)}
                  className="role-select"
                  disabled={isAddingPermission}
                >
                  <option value="read">Viewer</option>
                  <option value="write">Editor</option>
                </select>
                <button
                  className="btn btn-primary btn-sm"
                  onClick={() => handleAddPermission(searchInput, "external")}
                  disabled={isAddingPermission}
                >
                  {isAddingUser("external") ? "Inviting..." : "Invite External"}
                </button>
              </div>
            </div>
          </div>
        )}

        {userSearchResults.length > 0 && (
          <div className="search-results">
            {userSearchResults.map((user) => (
              <div key={user.id} className="user-item">
                <div className="user-info">
                  <p className="user-name">{user.displayName}</p>
                  <p className="user-email">{user.mail || user.userPrincipalName}</p>
                </div>
                <div className="user-action">
                  <select
                    value={getUserRole(user.id)}
                    onChange={(e) => setUserRole(user.id, e.target.value)}
                    className="role-select"
                    disabled={isAddingPermission}
                  >
                    <option value="read">Viewer</option>
                    <option value="write">Editor</option>
                  </select>
                  <button
                    className="btn btn-primary btn-sm"
                    onClick={() => handleAddPermission(user.mail || user.userPrincipalName, user.id)}
                    disabled={isAddingPermission}
                  >
                    {isAddingUser(user.id) ? "Adding..." : "Add"}
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

        <div className="current-permissions-section">
        <div className="section-heading-row">
          <div>
            <h3>People with access</h3>
            <p className="section-kicker">Direct permissions returned by SharePoint</p>
          </div>
          {!isLoading && (
            <span className="access-count">
              {permissions.length} {permissions.length === 1 ? "entry" : "entries"}
            </span>
          )}
        </div>
        {isLoading ? (
          <div className="loading"><span className="spinner"></span><p>Loading permissions...</p></div>
        ) : permissions.length === 0 ? (
          <p className="no-permissions">No direct permissions set</p>
        ) : (
          <div className="permissions-list">
            {permissions.map((perm) => {
              const principal = getPermissionPrincipal(perm);
              return (
              <div key={perm.id} className="permission-item">
                <div className="permission-user">
                  <div className={`user-avatar ${principal.type === "group" ? "group" : ""} ${getRoleClass(perm.roles)}`}>
                    {principal.avatar}
                  </div>
                  <div className="user-details">
                    <p className="user-name">{principal.name}</p>
                    <p className="user-email">{principal.email}</p>
                  </div>
                </div>
                {editingPermissionId === perm.id ? (
                  <div className="permission-actions edit-mode">
                    <select
                      value={editingNewRole}
                      onChange={(e) => setEditingNewRole(e.target.value)}
                      className="role-select-edit"
                    >
                      <option value="read">Viewer</option>
                      <option value="write">Editor</option>
                    </select>
                    <button
                      className="btn btn-primary btn-sm"
                      onClick={() => handleSavePermissionEdit(perm.id, perm.roles[0])}
                      disabled={isUpdatingPermission}
                    >
                      {isUpdatingPermission ? "Saving..." : "Save"}
                    </button>
                    <button
                      className="btn btn-outline btn-sm"
                      onClick={() => setEditingPermissionId(null)}
                      disabled={isUpdatingPermission}
                    >
                      Cancel
                    </button>
                  </div>
                ) : (
                  <div className="permission-actions">
                    <span className={`role-badge ${getRoleClass(perm.roles)}`}>{getRoleLabel(perm.roles)}</span>
                    {!perm.roles?.includes("owner") && (
                      <>
                        <button
                          className="permission-action-btn edit-action"
                          onClick={() => { setEditingPermissionId(perm.id); setEditingNewRole(perm.roles[0]); }}
                          title="Edit role"
                        >
                          Edit role
                        </button>
                        <button
                          className="permission-action-btn remove-action"
                          onClick={() => handleRemovePermission(perm.id)}
                          title="Remove access"
                        >
                          Remove access
                        </button>
                      </>
                    )}
                  </div>
                )}
              </div>
            );
            })}
          </div>
        )}
      </div>
      </div>

    </div>
  );
}

export default ItemPermissions;

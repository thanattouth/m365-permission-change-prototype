import React, { useState, useEffect } from "react";
import {
  getItemPermissions,
  addItemPermission,
  removeItemPermission,
  updateItemPermission,
  searchUsers,
} from "../services/sharePointService";
// File Classification is parked for this prototype phase.
// The current customer use case only needs admin-driven permission management
// across site contents, folders, and files.
// import FileClassification from "./FileClassification";
import "./ItemPermissions.css";

function ItemPermissions({ instance, item, onPermissionChanged, account }) {
  const [permissions, setPermissions] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [userSearchResults, setUserSearchResults] = useState([]);
  const [searchInput, setSearchInput] = useState("");
  const [isSearching, setIsSearching] = useState(false);
  const [userRoleMap, setUserRoleMap] = useState({});
  const [isAddingPermission, setIsAddingPermission] = useState(false);
  // const [classification, setClassification] = useState(null);
  const [editingPermissionId, setEditingPermissionId] = useState(null);
  const [editingNewRole, setEditingNewRole] = useState(null);
  const [isUpdatingPermission, setIsUpdatingPermission] = useState(false);

  const loadPermissions = React.useCallback(async () => {
    setIsLoading(true);
    setError(null);
    try {
      const perms = await getItemPermissions(instance, account, item.id, !!item.isLibrary);
      setPermissions(perms);
    } catch (err) {
      setError(err.message || "Failed to load permissions");
    } finally {
      setIsLoading(false);
    }
  }, [instance, account, item.id, item.isLibrary]);

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

  const handleAddPermission = async (userEmail, userId) => {
    const role = getUserRole(userId);
    setIsAddingPermission(true);
    setError(null);
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

      await addItemPermission(
        instance, 
        account, 
        item.id, 
        userEmail, 
        role, 
        itemServerRelativeUrl, 
        !!item.folder, 
        !!item.isLibrary
      );
      setSearchInput("");
      setUserSearchResults([]);
      setUserRoleMap({});
      await loadPermissions();
    } catch (err) {
      setError(err.message || "Failed to add permission");
    } finally {
      setIsAddingPermission(false);
    }
  };

  const handleRemovePermission = async (permissionId) => {
    if (!window.confirm("Remove this permission?")) return;

    try {
      await removeItemPermission(instance, account, item.id, permissionId, !!item.isLibrary);
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
      await updateItemPermission(instance, account, item.id, permissionId, editingNewRole, !!item.isLibrary);
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
      contribute: "Restrict Editor",
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

  // Helper for External Email
  const isEmailFormat = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(searchInput);

  return (
    <div className="permissions-container">
      <div className="permission-header">
        <div className="permission-item-info">
          <h2>{item.name}</h2>
          <p className="item-type">
            {item.isLibrary ? "📚 Library (Site Content)" : (item.folder ? "📁 Folder" : "📄 File")}
          </p>
        </div>
      </div>

      {error && (
        <div className="error-message">
          {error}
          <button onClick={() => setError(null)} className="close-error">×</button>
        </div>
      )}

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
        <h3>Grant Access</h3>
        <div className="search-box">
          <input
            type="text"
            placeholder="Search by name or email..."
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
                >
                  <option value="read">Viewer</option>
                  <option value="write">Editor</option>
                </select>
                <button
                  className="btn btn-primary btn-sm"
                  onClick={() => handleAddPermission(searchInput, "external")}
                  disabled={isAddingPermission}
                >
                  {isAddingPermission ? "Inviting..." : "Invite External"}
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
                  >
                    <option value="read">Viewer</option>
                    {!item.isLibrary && <option value="contribute">Restrict Editor</option>}
                    <option value="write">Editor</option>
                  </select>
                  <button
                    className="btn btn-primary btn-sm"
                    onClick={() => handleAddPermission(user.mail || user.userPrincipalName, user.id)}
                    disabled={isAddingPermission}
                  >
                    {isAddingPermission ? "Adding..." : "Add"}
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="current-permissions-section">
        <h3>Current Access</h3>
        {isLoading ? (
          <div className="loading"><span className="spinner"></span><p>Loading permissions...</p></div>
        ) : permissions.length === 0 ? (
          <p className="no-permissions">No direct permissions set</p>
        ) : (
          <div className="permissions-list">
            {permissions.map((perm) => (
              <div key={perm.id} className="permission-item">
                <div className="permission-user">
                  {perm.grantedTo?.user ? (
                    <>
                      <div className={`user-avatar ${getRoleClass(perm.roles)}`}>
                        {perm.grantedTo.user.displayName?.charAt(0).toUpperCase() || "U"}
                      </div>
                      <div className="user-details">
                        <p className="user-name">{perm.grantedTo.user.displayName}</p>
                        <p className="user-email">{perm.grantedTo.user.email}</p>
                      </div>
                    </>
                  ) : perm.grantedTo?.group ? (
                    <>
                      <div className={`user-avatar group ${getRoleClass(perm.roles)}`}>G</div>
                      <div className="user-details">
                        <p className="user-name">{perm.grantedTo.group.displayName}</p>
                        <p className="user-email">Group</p>
                      </div>
                    </>
                  ) : (
                    <div className="user-details"><p className="user-name">Unknown</p></div>
                  )}
                </div>
                {editingPermissionId === perm.id ? (
                  <div className="permission-actions edit-mode">
                    <select
                      value={editingNewRole}
                      onChange={(e) => setEditingNewRole(e.target.value)}
                      className="role-select-edit"
                    >
                      <option value="read">Viewer</option>
                      {!item.isLibrary && <option value="contribute">Restrict Editor</option>}
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
                          className="btn btn-outline btn-sm"
                          onClick={() => { setEditingPermissionId(perm.id); setEditingNewRole(perm.roles[0]); }}
                        >
                          Edit
                        </button>
                        <button
                          className="btn btn-outline btn-sm"
                          onClick={() => handleRemovePermission(perm.id)}
                        >
                          Remove
                        </button>
                      </>
                    )}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

export default ItemPermissions;

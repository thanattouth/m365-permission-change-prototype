import React, { useState, useEffect } from "react";
import {
  getItemPermissions,
  addItemPermission,
  removeItemPermission,
  searchUsers,
} from "../services/sharePointService";
import "./ItemPermissions.css";

function ItemPermissions({ instance, item, onPermissionChanged }) {
  const [permissions, setPermissions] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [userSearchResults, setUserSearchResults] = useState([]);
  const [searchInput, setSearchInput] = useState("");
  const [isSearching, setIsSearching] = useState(false);
  const [selectedRole, setSelectedRole] = useState("read");
  const [isAddingPermission, setIsAddingPermission] = useState(false);

  const loadPermissions = React.useCallback(async () => {
    setIsLoading(true);
    setError(null);
    try {
      const perms = await getItemPermissions(instance, item.id);
      setPermissions(perms);
    } catch (err) {
      setError(err.message || "Failed to load permissions");
    } finally {
      setIsLoading(false);
    }
  }, [instance, item.id]);

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
      const results = await searchUsers(instance, value);
      setUserSearchResults(results);
    } catch (err) {
      console.error("Search error:", err);
      setUserSearchResults([]);
    } finally {
      setIsSearching(false);
    }
  };

  const handleAddPermission = async (userEmail) => {
    setIsAddingPermission(true);
    setError(null);
    try {
      await addItemPermission(instance, item.id, userEmail, selectedRole);
      setSearchInput("");
      setUserSearchResults([]);
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
      await removeItemPermission(instance, item.id, permissionId);
      await loadPermissions();
    } catch (err) {
      setError(err.message || "Failed to remove permission");
    }
  };

  const getRoleLabel = (roles) => {
    if (!roles || roles.length === 0) return "View";
    const role = roles[0].toLowerCase();
    const roleMap = {
      read: "View",
      write: "Edit",
      owner: "Owner",
    };
    return roleMap[role] || role;
  };

  const canRemovePermission = (permission) => {
    // Don't allow removing owner permissions (optional)
    return permission.roles && !permission.roles.includes("owner");
  };

  return (
    <div className="permissions-container">
      <div className="permission-header">
        <div className="permission-item-info">
          <h2>{item.name}</h2>
          <p className="item-type">
            {item.folder ? "📁 Folder" : "📄 File"}
          </p>
        </div>
      </div>

      {error && (
        <div className="error-message">
          {error}
          <button
            onClick={() => setError(null)}
            className="close-error"
          >
            ×
          </button>
        </div>
      )}

      {/* Add Permission Section */}
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
                    value={selectedRole}
                    onChange={(e) => setSelectedRole(e.target.value)}
                    className="role-select"
                  >
                    <option value="read">View</option>
                    <option value="write">Edit</option>
                    <option value="owner">Owner</option>
                  </select>
                  <button
                    className="btn btn-primary btn-sm"
                    onClick={() =>
                      handleAddPermission(user.mail || user.userPrincipalName)
                    }
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

      {/* Current Permissions Section */}
      <div className="current-permissions-section">
        <h3>Current Access</h3>

        {isLoading ? (
          <div className="loading">
            <span className="spinner"></span>
            <p>Loading permissions...</p>
          </div>
        ) : permissions.length === 0 ? (
          <p className="no-permissions">No direct permissions set</p>
        ) : (
          <div className="permissions-list">
            {permissions.map((perm) => (
              <div key={perm.id} className="permission-item">
                <div className="permission-user">
                  {perm.grantedTo?.user ? (
                    <>
                      <div className="user-avatar">
                        {perm.grantedTo.user.displayName?.charAt(0).toUpperCase() ||
                          "U"}
                      </div>
                      <div className="user-details">
                        <p className="user-name">
                          {perm.grantedTo.user.displayName}
                        </p>
                        <p className="user-email">
                          {perm.grantedTo.user.email}
                        </p>
                      </div>
                    </>
                  ) : perm.grantedTo?.group ? (
                    <>
                      <div className="user-avatar group">G</div>
                      <div className="user-details">
                        <p className="user-name">
                          {perm.grantedTo.group.displayName}
                        </p>
                        <p className="user-email">Group</p>
                      </div>
                    </>
                  ) : (
                    <>
                      <div className="user-avatar">?</div>
                      <div className="user-details">
                        <p className="user-name">Unknown</p>
                      </div>
                    </>
                  )}
                </div>
                <div className="permission-role">
                  <span className="role-badge">{getRoleLabel(perm.roles)}</span>
                </div>
                {canRemovePermission(perm) && (
                  <button
                    className="btn btn-outline btn-sm remove-btn"
                    onClick={() => handleRemovePermission(perm.id)}
                  >
                    Remove
                  </button>
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

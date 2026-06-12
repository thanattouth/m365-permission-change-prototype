import React, { useState, useEffect } from "react";
import { getAllSiteContents, getDriveItems, getDriveIdFromListId } from "../services/sharePointService";
import ItemPermissions from "./ItemPermissions";
import Icon from "./Icon";
import "./ItemsList.css";

function ItemsList({ instance, accounts, selectedSite }) {
  const [items, setItems] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [selectedItem, setSelectedItem] = useState(null);
  const [activeTab, setActiveTab] = useState("all");
  
  // Navigation state
  // level: 'site' | 'library' | 'folder'
  const [navLevel, setNavLevel] = useState("site");
  const [currentDriveId, setCurrentDriveId] = useState(null);
  const [currentFolderId, setCurrentFolderId] = useState("root");
  const [navHistory, setNavHistory] = useState([{ level: "site", id: "site", name: "Site Contents" }]);

  useEffect(() => {
    setItems([]);
    setSelectedItem(null);
    setActiveTab("all");
    setNavLevel("site");
    setCurrentDriveId(null);
    setCurrentFolderId("root");
    setNavHistory([{ level: "site", id: "site", name: selectedSite?.label || "Site Contents" }]);
  }, [selectedSite]);

  const loadItems = React.useCallback(async () => {
    setIsLoading(true);
    setError(null);
    try {
      if (navLevel === "site") {
        const contents = await getAllSiteContents(instance, accounts[0], selectedSite.url);
        // Map lists to a common item format
        setItems(contents.map(l => ({
          id: l.id,
          name: l.displayName,
          isLibrary: true,
          webUrl: l.webUrl,
          lastModifiedDateTime: new Date().toISOString(), // Lists don't have this easily in the summary
        })));
      } else {
        const driveItems = await getDriveItems(instance, accounts[0], currentDriveId, currentFolderId);
        setItems(driveItems.map((item) => ({
          ...item,
          driveId: currentDriveId,
        })));
      }
    } catch (err) {
      setError(err.message || "Failed to load items");
      console.error("Error loading items:", err);
    } finally {
      setIsLoading(false);
    }
  }, [instance, accounts, navLevel, currentDriveId, currentFolderId, selectedSite]);

  useEffect(() => {
    if (accounts.length > 0) {
      loadItems();
    }
  }, [accounts, loadItems]);

  const handleLibraryClick = async (library) => {
    setIsLoading(true);
    try {
      const driveId = await getDriveIdFromListId(instance, accounts[0], library.id, selectedSite.url);
      setCurrentDriveId(driveId);
      setCurrentFolderId("root");
      setNavLevel("library");
      setNavHistory((prev) => [...prev, { level: "library", id: library.id, driveId, name: library.name }]);
      setActiveTab("all");
    } catch (err) {
      setError("Failed to access library: " + err.message);
    } finally {
      setIsLoading(false);
    }
  };

  const handleFolderClick = (folder) => {
    setCurrentFolderId(folder.id);
    setNavLevel("folder");
    setNavHistory((prev) => [...prev, { level: "folder", id: folder.id, name: folder.name }]);
    setActiveTab("all");
  };

  const handleBreadcrumbClick = (index) => {
    const newHistory = navHistory.slice(0, index + 1);
    const target = newHistory[index];
    
    setNavHistory(newHistory);
    setNavLevel(target.level);
    
    if (target.level === "site") {
      setCurrentDriveId(null);
      setCurrentFolderId("root");
    } else if (target.level === "library") {
      setCurrentDriveId(target.driveId);
      setCurrentFolderId("root");
    } else {
      setCurrentFolderId(target.id);
    }
    setActiveTab("all");
  };

  const getItemIconName = (item) => {
    if (item.isLibrary) return "library";
    if (item.folder) return "folder";
    const ext = item.name.split(".").pop()?.toLowerCase();
    const iconMap = {
      pdf: "doc", doc: "doc", docx: "doc", xlsx: "sheet", xls: "sheet",
      pptx: "slides", ppt: "slides", txt: "doc", jpg: "image", jpeg: "image",
      png: "image", gif: "image",
    };
    return iconMap[ext] || "paperclip";
  };

  const getPillLabel = (item) => {
    if (item.isLibrary) return "Library";
    if (item.folder) return "Folder";
    const ext = item.name.split(".").pop()?.toUpperCase() || "FILE";
    return `${ext} File`;
  };

  const formatFileSize = (bytes) => {
    if (!bytes || bytes === 0) return "0 B";
    const k = 1024;
    const sizes = ["B", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + " " + sizes[i];
  };

  const formatDate = (dateString) => {
    return new Date(dateString).toLocaleDateString("en-US", {
      year: "numeric", month: "short", day: "numeric",
    });
  };

  const cardTints = ["green", "blue", "purple", "yellow", "orange", "teal"];
  const filteredItems = items.filter(item => {
    if (activeTab === "folders") return item.folder;
    if (activeTab === "files") return !item.folder && !item.isLibrary;
    return true;
  });
  const currentNode = navHistory[navHistory.length - 1];
  const pageTitle = navLevel === "site" ? `${selectedSite.label} Contents` : currentNode.name;
  const pageDescription = navLevel === "site"
    ? "Review document libraries available in this SharePoint site."
    : "Browse folders and files, then manage direct access for the selected item.";
  const libraryCount = items.filter((item) => item.isLibrary).length;
  const folderCount = items.filter((item) => item.folder).length;
  const fileCount = items.filter((item) => !item.folder && !item.isLibrary).length;
  const managementScope = navLevel === "site"
    ? "Library level"
    : navLevel === "library"
      ? "Library root"
      : "Folder level";

  if (isLoading) {
    return (
      <div className="items-loading">
        <div className="items-loading-card">
          <span className="spinner"></span>
          <p>Loading {navLevel === "site" ? "Site Contents" : "Items"}...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="items-error-container">
        <div className="error-message" style={{ marginBottom: "16px" }}>{error}</div>
        <button className="btn btn-primary" onClick={loadItems} style={{ width: "auto" }}>
          Retry
        </button>
      </div>
    );
  }

  return (
    <div className="items-container">
      {selectedItem ? (
        <div className="permissions-view">
          <ItemPermissions
            instance={instance}
            item={selectedItem}
            selectedSite={selectedSite}
            account={accounts[0]}
            onClose={() => setSelectedItem(null)}
            onPermissionChanged={() => {
              setSelectedItem(null);
              loadItems();
            }}
          />
        </div>
      ) : (
        <>
          <div className="items-page-header">
            <div>
              <span className="section-kicker">SharePoint workspace</span>
              <h1>{pageTitle}</h1>
              <p>{pageDescription}</p>
            </div>
          </div>

          <div className="portal-subbar">
            <div className="tab-links">
              <button 
                className={`tab-link ${activeTab === "all" ? "active" : ""}`}
                onClick={() => setActiveTab("all")}
              >
                All {navLevel === "site" ? "Contents" : "Files"}
              </button>
              {navLevel !== "site" && (
                <>
                  <button 
                    className={`tab-link ${activeTab === "folders" ? "active" : ""}`}
                    onClick={() => setActiveTab("folders")}
                  >
                    Folders
                  </button>
                  <button 
                    className={`tab-link ${activeTab === "files" ? "active" : ""}`}
                    onClick={() => setActiveTab("files")}
                  >
                    Files
                  </button>
                </>
              )}
            </div>
            
            <div className="subbar-actions">
              <span className="item-count">{items.length} items</span>
              <button className="icon-btn" onClick={loadItems} title="Refresh items list">
                <Icon name="refresh" size={16} />
              </button>
            </div>
          </div>

          <div className="breadcrumbs">
            {navHistory.map((node, index) => (
              <span key={`${node.level}-${node.id}`} className="breadcrumb-node">
                {index > 0 && <span className="breadcrumb-separator">/</span>}
                <button 
                  className={`breadcrumb-item ${index === navHistory.length - 1 ? "active" : ""}`}
                  onClick={() => handleBreadcrumbClick(index)}
                  disabled={index === navHistory.length - 1}
                >
                  {node.level === "site" ? (
                    <>
                      <Icon name="home" className="breadcrumb-icon" size={14} />
                      Site Contents
                    </>
                  ) : node.level === "library" ? (
                    <>
                      <Icon name="library" className="breadcrumb-icon" size={14} />
                      {node.name}
                    </>
                  ) : (
                    node.name
                  )}
                </button>
              </span>
            ))}
          </div>

          <div className="items-workspace">
            <div className="items-main">
              {items.length === 0 ? (
                <div className="empty-state">
                  <p>No items found in this {navLevel === "site" ? "site" : "directory"}.</p>
                </div>
              ) : (
                <div className="items-list">
                  <div className="items-list-header">
                    <span>Name</span>
                    <span>Type</span>
                    <span>Modified</span>
                    <span>Actions</span>
                  </div>
                  {filteredItems.map((item, index) => {
                    const tint = cardTints[index % cardTints.length];
                    return (
                      <div
                        key={item.id}
                        className={`item-card tint-${tint}`}
                        onClick={() => {
                          if (item.isLibrary) {
                            handleLibraryClick(item);
                          } else if (item.folder) {
                            handleFolderClick(item);
                          } else {
                            setSelectedItem(item);
                          }
                        }}
                      >
                        <div className="card-badge-header">
                          <div className="card-icon-badge">
                            <Icon name={getItemIconName(item)} className="file-icon-symbol" size={22} />
                          </div>
                          <div className="card-title-block">
                            <h3 className="card-item-title" title={item.name}>{item.name}</h3>
                            <p className="card-item-desc">
                              {item.isLibrary ? "Document Library (Site Content)" : 
                               (item.folder ? "Folder containing document assets." : `File asset. Size: ${formatFileSize(item.size)}`)}
                            </p>
                          </div>
                        </div>
                        
                        <div className="card-type-pill">
                          {getPillLabel(item)}
                        </div>

                        <div className="card-meta-row">
                          <span className="card-date-modified">
                            {item.isLibrary ? "Site Content Level" : `Modified ${formatDate(item.lastModifiedDateTime)}`}
                          </span>
                        </div>

                        <div className="card-action-bar">
                          <button 
                            className="card-btn card-btn-secondary"
                            onClick={(e) => { e.stopPropagation(); setSelectedItem(item); }}
                          >
                            <Icon name="shield" size={15} />
                            Manage Access
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            <aside className="site-insights-panel" aria-label="Site access overview">
              <div className="insights-card">
                <div className="insights-card-header">
                  <div className="insights-icon">
                    <Icon name="shield" size={18} />
                  </div>
                  <div>
                    <h2>Access Overview</h2>
                    <p>{managementScope}</p>
                  </div>
                </div>

                <div className="insights-metrics">
                  <div>
                    <strong>{navLevel === "site" ? libraryCount : folderCount}</strong>
                    <span>{navLevel === "site" ? "Libraries" : "Folders"}</span>
                  </div>
                  <div>
                    <strong>{fileCount}</strong>
                    <span>Files</span>
                  </div>
                </div>

                <div className="insights-note">
                  <Icon name="users" size={16} />
                  <span>Open Manage Access on any row to review the exact people and roles returned by SharePoint.</span>
                </div>
              </div>

              <div className="insights-card compact">
                <h3>Permission Scope</h3>
                <ul className="scope-list">
                  <li>
                    <span>Site</span>
                    <strong>{selectedSite.label}</strong>
                  </li>
                  <li>
                    <span>Current view</span>
                    <strong>{currentNode.name}</strong>
                  </li>
                  <li>
                    <span>Managed action</span>
                    <strong>Direct permissions</strong>
                  </li>
                </ul>
              </div>
            </aside>
          </div>
        </>
      )}
    </div>
  );
}

export default ItemsList;

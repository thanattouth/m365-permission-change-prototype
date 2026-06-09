import React, { useState, useEffect } from "react";
import { getItemsList } from "../services/sharePointService";
import ItemPermissions from "./ItemPermissions";
import "./ItemsList.css";

function ItemsList({ instance, accounts }) {
  const [items, setItems] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [selectedItem, setSelectedItem] = useState(null);
  const [activeTab, setActiveTab] = useState("all");

  const loadItems = React.useCallback(async () => {
    setIsLoading(true);
    setError(null);
    try {
      const itemsList = await getItemsList(instance, accounts[0]);
      setItems(itemsList);
    } catch (err) {
      setError(err.message || "Failed to load items");
      console.error("Error loading items:", err);
    } finally {
      setIsLoading(false);
    }
  }, [instance, accounts]);

  useEffect(() => {
    if (accounts.length > 0) {
      loadItems();
    }
  }, [accounts, loadItems]);

  const getItemIcon = (item) => {
    if (item.folder) {
      return "📁";
    }
    const ext = item.name.split(".").pop()?.toLowerCase();
    const iconMap = {
      pdf: "📄",
      doc: "📝",
      docx: "📝",
      xlsx: "📊",
      xls: "📊",
      pptx: "🎯",
      ppt: "🎯",
      txt: "📄",
      jpg: "🖼️",
      jpeg: "🖼️",
      png: "🖼️",
      gif: "🖼️",
    };
    return iconMap[ext] || "📎";
  };

  const getPillLabel = (item) => {
    if (item.folder) {
      return "Folder";
    }
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
    return new Date(dateString).toLocaleDateString("th-TH", {
      year: "numeric",
      month: "short",
      day: "numeric",
    });
  };

  // Card background tint classes matching the color mockups
  const cardTints = ["green", "blue", "purple", "yellow", "orange", "teal"];

  if (isLoading) {
    return (
      <div className="loading">
        <span className="spinner"></span>
        <p>Loading items...</p>
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
          <button
            className="btn btn-outline back-btn"
            onClick={() => setSelectedItem(null)}
          >
            ← Back to All Files
          </button>
          <ItemPermissions
            instance={instance}
            item={selectedItem}
            account={accounts[0]}
            onPermissionChanged={() => {
              setSelectedItem(null);
            }}
          />
        </div>
      ) : (
        <>
          {/* Sub-Navbar containing Categories/Tabs and quick actions */}
          <div className="portal-subbar">
            <div className="tab-links">
              <button 
                className={`tab-link ${activeTab === "all" ? "active" : ""}`}
                onClick={() => setActiveTab("all")}
              >
                All Files
              </button>
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
            </div>
            
            <div className="subbar-actions">
              <span className="item-count">{items.length} items</span>
              <button className="icon-btn" onClick={loadItems} title="Refresh items list">
                🔄
              </button>
            </div>
          </div>

          {/* Grid Layout of Items */}
          {items.length === 0 ? (
            <div className="empty-state">
              <p>No items found in the root directory of this SharePoint site.</p>
            </div>
          ) : (
            <div className="items-grid">
              {items
                .filter(item => {
                  if (activeTab === "folders") return item.folder;
                  if (activeTab === "files") return !item.folder;
                  return true;
                })
                .map((item, index) => {
                  const tint = cardTints[index % cardTints.length];
                  return (
                    <div
                      key={item.id}
                      className={`item-card tint-${tint}`}
                      onClick={() => setSelectedItem(item)}
                    >
                      {/* Top badge indicators on the card */}
                      <div className="card-badge-header">
                        <div className="card-icon-badge">
                          <span className="file-icon-symbol">{getItemIcon(item)}</span>
                        </div>
                        <div className="card-dot-badge"></div>
                      </div>

                      {/* Card Content */}
                      <h3 className="card-item-title">{item.name}</h3>
                      
                      <p className="card-item-desc">
                        {item.folder ? "Folder containing document assets." : `File asset. Size: ${formatFileSize(item.size)}`}
                      </p>

                      <div className="card-meta-row">
                        <span className="card-date-modified">
                          Modified {formatDate(item.lastModifiedDateTime)}
                        </span>
                        {item.lastModifiedBy?.user && (
                          <span className="card-modified-by" title={item.lastModifiedBy.user.displayName}>
                            by {item.lastModifiedBy.user.displayName.split(" ")[0]}
                          </span>
                        )}
                      </div>

                      {/* Bottom-right type pill badge matching the mockup */}
                      <div className="card-type-pill">
                        {getPillLabel(item)}
                      </div>
                    </div>
                  );
                })}
            </div>
          )}
        </>
      )}
    </div>
  );
}

export default ItemsList;

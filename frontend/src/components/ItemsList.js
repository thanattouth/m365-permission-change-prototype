import React, { useState, useEffect } from "react";
import { getItemsList } from "../services/sharePointService";
import ItemPermissions from "./ItemPermissions";
import "./ItemsList.css";

function ItemsList({ instance, accounts }) {
  const [items, setItems] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [selectedItem, setSelectedItem] = useState(null);

  useEffect(() => {
    if (accounts.length > 0) {
      loadItems();
    }
  }, [accounts]);

  const loadItems = async () => {
    setIsLoading(true);
    setError(null);
    try {
      const itemsList = await getItemsList(instance);
      setItems(itemsList);
    } catch (err) {
      setError(err.message || "Failed to load items");
      console.error("Error loading items:", err);
    } finally {
      setIsLoading(false);
    }
  };

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

  const formatFileSize = (bytes) => {
    if (bytes === 0) return "0 B";
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

  if (isLoading) {
    return (
      <div className="items-container">
        <div className="loading">
          <span className="spinner"></span>
          <p>Loading items...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="items-container">
        <div className="error-message">{error}</div>
        <button className="btn btn-primary" onClick={loadItems}>
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
            ← Back to Items
          </button>
          <ItemPermissions
            instance={instance}
            item={selectedItem}
            onPermissionChanged={() => {
              setSelectedItem(null);
            }}
          />
        </div>
      ) : (
        <>
          <div className="items-header">
            <h2>Files & Folders</h2>
            <span className="item-count">{items.length} items</span>
            <button className="btn btn-secondary" onClick={loadItems}>
              🔄 Refresh
            </button>
          </div>

          {items.length === 0 ? (
            <div className="empty-state">
              <p>No items found in this location</p>
            </div>
          ) : (
            <div className="items-list">
              {items.map((item) => (
                <div
                  key={item.id}
                  className="item-card"
                  onClick={() => setSelectedItem(item)}
                >
                  <div className="item-icon">{getItemIcon(item)}</div>
                  <div className="item-info">
                    <h3 className="item-name">{item.name}</h3>
                    <div className="item-meta">
                      <span className="item-type">
                        {item.folder ? "Folder" : "File"}
                      </span>
                      {item.size && (
                        <span className="item-size">
                          {formatFileSize(item.size)}
                        </span>
                      )}
                      <span className="item-date">
                        {formatDate(item.lastModifiedDateTime)}
                      </span>
                    </div>
                    {item.lastModifiedBy?.user && (
                      <p className="item-modified-by">
                        by {item.lastModifiedBy.user.displayName}
                      </p>
                    )}
                  </div>
                  <div className="item-action">
                    <button className="btn btn-primary btn-sm">
                      Manage Permissions
                    </button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </>
      )}
    </div>
  );
}

export default ItemsList;

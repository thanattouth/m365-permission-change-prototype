import React from "react";
import Icon from "./Icon";
import "./SitePicker.css";

const getInitials = (label) => {
  const words = label.split(/\s+/).filter(Boolean);
  if (words.length === 1) return words[0].slice(0, 2).toUpperCase();
  return words.slice(0, 2).map((word) => word[0]).join("").toUpperCase();
};

function SitePicker({ sites, onSelect }) {
  return (
    <div className="site-picker">
      <div className="site-picker-hero">
        <div>
          <span className="section-kicker">Workspace selection</span>
          <h1>Sites</h1>
          <p>Choose the SharePoint workspace where you want to review libraries, folders, files, and access.</p>
        </div>
      </div>

      <div className="site-card-grid">
        {sites.map((site, index) => (
          <button
            key={site.url}
            className={`site-card site-card-${index % 4}`}
            onClick={() => onSelect(site)}
          >
            <div className="site-card-band">
              <Icon name="shield" className="site-card-star" size={18} />
            </div>
            <div className="site-card-body">
              <div className="site-card-avatar">{getInitials(site.label)}</div>
              <div className="site-card-copy">
                <h2>{site.label}</h2>
                <p>{new URL(site.url).hostname}</p>
              </div>
              <div className="site-card-meta">
                <span><Icon name="library" size={14} /> Libraries</span>
                <span><Icon name="users" size={14} /> Permissions</span>
              </div>
              <span className="site-card-action">
                <Icon name="folder" size={16} />
                Open site
              </span>
            </div>
          </button>
        ))}
      </div>
    </div>
  );
}

export default SitePicker;

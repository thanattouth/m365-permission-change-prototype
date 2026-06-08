import React, { useState } from "react";
import {
  CLASSIFICATION_OPTIONS,
  getClassification,
  getClassificationLabel,
} from "../config/classificationConfig";
import { applyClassificationPermissions } from "../services/classificationService";
import "./FileClassification.css";

function FileClassification({
  instance,
  account,
  item,
  currentClassification,
  onClassificationChanged,
}) {
  const [selectedClassification, setSelectedClassification] = useState(
    currentClassification || ""
  );
  const [isApplying, setIsApplying] = useState(false);
  const [applyResult, setApplyResult] = useState(null);
  const [error, setError] = useState(null);

  const handleClassificationChange = async (newClassificationId) => {
    if (!newClassificationId) return;

    setSelectedClassification(newClassificationId);
    setError(null);
    setApplyResult(null);

    const classification = getClassification(newClassificationId);

    // Show confirmation if requires approval
    if (classification.requiresApproval) {
      const confirmed = window.confirm(
        `This classification requires ${classification.approverRole} approval.\n\nProceed with applying default permissions?`
      );
      if (!confirmed) {
        setSelectedClassification(currentClassification || "");
        return;
      }
    }

    // Apply permissions automatically
    setIsApplying(true);
    try {
      const result = await applyClassificationPermissions(
        instance,
        account,
        item.id,
        newClassificationId
      );

      setApplyResult({
        success: true,
        classification: result.classification,
        added: result.added,
        failed: result.failed,
      });

      // Notify parent component
      if (onClassificationChanged) {
        onClassificationChanged(newClassificationId, result);
      }
    } catch (err) {
      setError(err.message || "Failed to apply classification");
      setSelectedClassification(currentClassification || "");
    } finally {
      setIsApplying(false);
    }
  };

  const getClassificationBgColor = (classificationId) => {
    const classification = getClassification(classificationId);
    return classification ? classification.bgColor : "#f0f0f0";
  };

  return (
    <div className="file-classification-section">
      <h3>File Classification</h3>

      <div className="classification-selector">
        <label htmlFor="classification-select">Classification Level:</label>
        <select
          id="classification-select"
          value={selectedClassification}
          onChange={(e) => handleClassificationChange(e.target.value)}
          disabled={isApplying}
          className="classification-select"
        >
          <option value="">Select classification...</option>
          {CLASSIFICATION_OPTIONS.map((option) => (
            <option key={option.value} value={option.value}>
              {option.label}
            </option>
          ))}
        </select>
      </div>

      {selectedClassification && (
        <div
          className="classification-badge"
          style={{
            backgroundColor: getClassificationBgColor(selectedClassification),
          }}
        >
          <div className="badge-content">
            <div className="badge-label">
              {getClassificationLabel(selectedClassification)}
            </div>
            {(() => {
              const classification = getClassification(selectedClassification);
              return (
                <>
                  <p className="badge-description">{classification.description}</p>
                  {classification.requiresApproval && (
                    <p className="badge-approval">
                      ⚠️ Requires {classification.approverRole} approval
                    </p>
                  )}
                </>
              );
            })()}
          </div>
        </div>
      )}

      {isApplying && (
        <div className="applying-state">
          <span className="spinner-small"></span>
          <p>Applying classification permissions...</p>
        </div>
      )}

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

      {applyResult && applyResult.success && (
        <div className="success-message">
          <div className="success-header">
            ✓ Classification applied successfully
          </div>
          {applyResult.added.length > 0 && (
            <div className="permissions-added">
              <h4>Permissions Added:</h4>
              <ul>
                {applyResult.added.map((perm, idx) => (
                  <li key={idx}>
                    {perm.email} - {perm.role}
                  </li>
                ))}
              </ul>
            </div>
          )}
          {applyResult.failed.length > 0 && (
            <div className="permissions-failed">
              <h4>Failed to apply:</h4>
              <ul>
                {applyResult.failed.map((perm, idx) => (
                  <li key={idx}>
                    {perm.email} - {perm.error}
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
      )}

      {currentClassification && !selectedClassification && (
        <p className="info-text">Current classification: {getClassificationLabel(currentClassification)}</p>
      )}
    </div>
  );
}

export default FileClassification;

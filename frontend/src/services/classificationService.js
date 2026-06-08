import { getClassification } from "../config/classificationConfig";
import { addItemPermission } from "./sharePointService";

/**
 * Apply classification permissions to an item
 * This will set default permissions based on classification level
 */
export const applyClassificationPermissions = async (
  instance,
  account,
  itemId,
  classificationId
) => {
  try {
    const classification = getClassification(classificationId);
    
    if (!classification) {
      throw new Error(`Invalid classification: ${classificationId}`);
    }

    // Get current permissions to clear them first
    // Note: In production, you might want to be more selective about what to remove
    
    const results = {
      added: [],
      failed: [],
      classification: classification
    };

    // Apply owner permissions
    for (const ownerGroup of classification.defaultGroups.owner) {
      try {
        await addItemPermission(
          instance,
          account,
          itemId,
          `${ownerGroup}@m365.co.th`, // Adjust email domain as needed
          "owner"
        );
        results.added.push({
          email: `${ownerGroup}@m365.co.th`,
          role: "owner"
        });
      } catch (err) {
        results.failed.push({
          email: `${ownerGroup}@m365.co.th`,
          role: "owner",
          error: err.message
        });
      }
    }

    // Apply editor permissions
    for (const editorGroup of classification.defaultGroups.editor) {
      try {
        await addItemPermission(
          instance,
          account,
          itemId,
          `${editorGroup}@m365.co.th`,
          "write"
        );
        results.added.push({
          email: `${editorGroup}@m365.co.th`,
          role: "write"
        });
      } catch (err) {
        results.failed.push({
          email: `${editorGroup}@m365.co.th`,
          role: "write",
          error: err.message
        });
      }
    }

    return results;
  } catch (error) {
    console.error("Error applying classification permissions:", error);
    throw error;
  }
};

/**
 * Get classification from item metadata
 * For now, this is a placeholder that can be expanded later
 */
export const getItemClassification = async (itemId) => {
  // TODO: Fetch classification from SharePoint item properties
  // For now, return null - can be implemented later with custom properties
  return null;
};

/**
 * Set classification for an item
 * Stores classification in item metadata/properties
 */
export const setItemClassification = async (
  instance,
  account,
  itemId,
  classificationId
) => {
  try {
    // TODO: Update SharePoint item properties with classification
    // For now, this is a placeholder
    return { success: true, classificationId };
  } catch (error) {
    console.error("Error setting item classification:", error);
    throw error;
  }
};

/**
 * Validate if user can view item based on classification
 */
export const canViewItem = (userRole, classificationId) => {
  const classification = getClassification(classificationId);
  
  if (!classification) return false;
  
  // Owner can always view
  if (userRole === "owner") return true;
  
  // Editor can view
  if (userRole === "editor" || userRole === "write") return true;
  
  // Viewer can view
  if (userRole === "viewer" || userRole === "read") return true;
  
  return false;
};

/**
 * Validate if user can edit item based on classification
 */
export const canEditItem = (userRole, classificationId) => {
  const classification = getClassification(classificationId);
  
  if (!classification) return false;
  
  // Only owner and editor can edit
  return userRole === "owner" || userRole === "editor" || userRole === "write";
};

/**
 * Validate if user can delete item based on classification
 */
export const canDeleteItem = (userRole, classificationId) => {
  // Only owner can delete
  return userRole === "owner";
};

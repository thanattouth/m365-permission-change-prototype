/**
 * File Classification Configuration
 * Defines classification levels and their corresponding permissions
 */

export const CLASSIFICATIONS = {
  INTERNAL: {
    id: "internal",
    label: "Internal (DGCS)",
    color: "#FFD700",
    bgColor: "#FFFACD",
    description: "Executive committee access only",
    owner: "DGCS",
    permissions: {
      owner: ["owner"],
      editor: ["write"],
      viewer: ["read"]
    },
    defaultGroups: {
      owner: ["DGCS"],
      editor: [],
      viewer: []
    }
  },
  CONFIDENTIAL: {
    id: "confidential",
    label: "Confidential - Custom (DGVP)",
    color: "#FF8C00",
    bgColor: "#FFE4B5",
    description: "Requires approval from VP level or above",
    owner: "DGVP",
    permissions: {
      owner: ["owner"],
      editor: ["write"],
      viewer: ["read"]
    },
    defaultGroups: {
      owner: ["DGVP"],
      editor: [],
      viewer: []
    },
    requiresApproval: true,
    approverRole: "VP"
  },
  SECRET: {
    id: "secret",
    label: "Secret - Custom (EVPN)",
    color: "#FF0000",
    bgColor: "#FFB6C6",
    description: "Requires approval from EVP level or above",
    owner: "EVPN",
    permissions: {
      owner: ["owner"],
      editor: ["write"],
      viewer: ["read"]
    },
    defaultGroups: {
      owner: ["EVP"],
      editor: [],
      viewer: []
    },
    requiresApproval: true,
    approverRole: "EVP"
  }
};

export const CLASSIFICATION_OPTIONS = Object.values(CLASSIFICATIONS).map(c => ({
  value: c.id,
  label: c.label,
  color: c.color,
  bgColor: c.bgColor
}));

/**
 * Get classification by ID
 */
export const getClassification = (classificationId) => {
  return Object.values(CLASSIFICATIONS).find(c => c.id === classificationId);
};

/**
 * Get classification color
 */
export const getClassificationColor = (classificationId) => {
  const classification = getClassification(classificationId);
  return classification ? classification.color : "#666";
};

/**
 * Get classification label
 */
export const getClassificationLabel = (classificationId) => {
  const classification = getClassification(classificationId);
  return classification ? classification.label : "Unknown";
};

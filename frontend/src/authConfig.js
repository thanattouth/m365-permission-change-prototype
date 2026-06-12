const sharePointSiteUrl =
  process.env.REACT_APP_SHAREPOINT_SITE_URL ||
  "https://devm365th.sharepoint.com/sites/DocumentManagement";

const parseSharePointSites = (rawSites, fallbackUrl) => {
  const entries = rawSites
    ? rawSites.split(",")
    : [];

  if (fallbackUrl) {
    entries.unshift(`Primary Site|${fallbackUrl}`);
  }

  const seen = new Set();

  return entries
    .map((entry) => {
      const [label, url] = entry.split("|").map((part) => part?.trim());
      const siteUrl = url || label;
      if (!siteUrl) return null;
      const siteLabel = url ? label : siteUrl.split("/").filter(Boolean).pop();
      return { label: siteLabel || siteUrl, url: siteUrl };
    })
    .filter((site) => {
      if (!site || seen.has(site.url)) return false;
      seen.add(site.url);
      return true;
    })
    .filter(Boolean);
};

export const sharePointSites = parseSharePointSites(
  process.env.REACT_APP_SHAREPOINT_SITES,
  sharePointSiteUrl
);

const tenantId =
  process.env.REACT_APP_TENANT_ID ||
  "0f3101bc-add7-42aa-a041-4b5648c7bacf";

const clientId =
  process.env.REACT_APP_CLIENT_ID ||
  "5e597123-d103-4c16-a7f3-cb99da285561";

const redirectUri =
  process.env.REACT_APP_REDIRECT_URI ||
  window.location.origin;

const sharePointHostname = new URL(sharePointSiteUrl).hostname;
const sharePointRestScope =
  process.env.REACT_APP_SHAREPOINT_REST_SCOPE ||
  `https://${sharePointHostname}/AllSites.FullControl`;

const graphScopes = [
  "User.Read",
  "Sites.ReadWrite.All",
  "Files.ReadWrite.All",
  "Directory.ReadWrite.All",
];

// SharePoint Site Configuration
export const sharePointConfig = {
  siteUrl: sharePointSiteUrl,
};

export const msalConfig = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri,
  }
};

export const loginRequest = {
  scopes: [
    ...graphScopes,
    sharePointRestScope,
  ]
};

// For getting access token with specific scopes
export const tokenRequest = {
  scopes: [
    ...graphScopes.filter((scope) => scope !== "User.Read"),
    sharePointRestScope,
  ]
};

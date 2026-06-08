export const msalConfig = {
  auth: {
    clientId: "5e597123-d103-4c16-a7f3-cb99da285561",
    authority: "https://login.microsoftonline.com/0f3101bc-add7-42aa-a041-4b5648c7bacf",
    redirectUri: window.location.origin
  }
};

export const loginRequest = {
  scopes: ["User.Read", "Sites.ReadWrite.All"]
};

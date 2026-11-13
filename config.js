// Public settings for the phone app. Nothing here is secret: the client ID only identifies the app
// to Microsoft. Your data and the Excel link are never stored in this file; each phone keeps them locally.
window.APP_CONFIG = {
  // Application (client) ID from the Microsoft Entra app registration (see README_MOBILE.md).
  clientId: "PASTE-YOUR-CLIENT-ID-HERE",
  // "consumers" = personal Microsoft accounts (Outlook/Hotmail/Live). Use "common" if the registration
  // also allows work/school accounts.
  tenant: "consumers",
};

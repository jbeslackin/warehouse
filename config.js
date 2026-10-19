// Warehouse Commander: Ranked Play settings. IT fills in the first four values (see README.md, steps 1–3).
window.DQB_CONFIG = {
  // From the Entra ID app registration's Overview page
  TENANT_ID: "PASTE-DIRECTORY-TENANT-ID",
  CLIENT_ID: "PASTE-APPLICATION-CLIENT-ID",

  // The SharePoint site that holds the data file, e.g. https://wyanokegroup.sharepoint.com/sites/Warehouse
  SITE_HOSTNAME: "wyanokegroup.sharepoint.com",
  SITE_PATH: "/sites/Warehouse",

  // Path to the data file inside that site's "Documents" library
  DATA_FILE: "Warehouse Commander/warehouse-commander.json",

  // Admins' private list (quests from their own inbox). Lives in each admin's OneDrive and is created automatically.
  PERSONAL_FILE: "Warehouse Commander/my-personal-quests.json",

  // People who see the admin tools (post, assign and delete quests; remove players)
  ADMIN_EMAILS: ["jslack@wyanokegroup.com"],

  // How often open boards check for changes made by others
  REFRESH_SECONDS: 20,
};

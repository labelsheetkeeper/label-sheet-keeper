// Label Sheet Keeper — values set once for everyone.
//
// msClientId: the app's "Application (client) ID" from its Microsoft
// registration (not a secret). While it's empty, signing in and sync stay off
// and hidden. The same ID goes in worker/wrangler.toml as MS_CLIENT_ID.
window.LSK_CONFIG = {
  msClientId: ''
};

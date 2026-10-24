// Desk dashboard config. Copy this file to "config.js" and set your endpoint.
//
// The endpoint is your Apps Script web app deployment URL. It is NOT a
// secret by itself (it is unguessable and every request is key-checked),
// but treat it like an address you don't post publicly.
//
// The shared KEY is never written here. You type it once on the tablet
// itself and the page stores it on that device only.
window.DASH_CONFIG = {
  endpoint: 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec',
  // Weather location: Sunnyvale, CA 94089 (America/Los_Angeles)
  latitude: 37.36883,
  longitude: -122.03635,
  locationLabel: 'Sunnyvale · 94089',
  timezone: 'America/Los_Angeles'
};

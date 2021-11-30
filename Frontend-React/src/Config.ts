const Config = {
  uri: ["https://graph.microsoft.com"],
  versions: ["stagingbeta", "stagingv1.0", "beta", "v1.0"],
  endpoints: ["me","groups"],
  appId: '8a792f49-ae0d-4b9b-92d2-614fcba43bea',
  redirectUri: 'http://localhost:8080/',
  scopes: [
    'user.read',
    'mailboxsettings.read',
    'calendars.readwrite'
  ]
};

export default Config;
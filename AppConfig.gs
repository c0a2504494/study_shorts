const APP_CONFIG_ = Object.freeze({
  appName: 'Study Shorts',
  version: '0.3.0',
  schemaVersion: 1,
  bookSchemaVersion: 1,
});

function getAppBootstrap() {
  return {
    config: { ...APP_CONFIG_ },
    books: getSheetNames(),
  };
}

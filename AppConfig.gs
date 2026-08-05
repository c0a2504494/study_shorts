const APP_CONFIG_ = Object.freeze({
  appName: 'Study Shorts',
  version: '0.6.0',
  schemaVersion: 1,
  bookSchemaVersion: 1,
  storage: 'indexeddb',
});

function getAppBootstrap() {
  return {
    config: { ...APP_CONFIG_ },
    books: getSheetNames(),
  };
}

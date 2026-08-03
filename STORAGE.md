# Storage architecture

Study Shorts 0.4 uses the device database (`IndexedDB`) as its primary storage.

## Stored locally

Each Book is stored as a versioned JSON-compatible document containing:

- Book name
- Question and answer rows
- Study count and rank
- Comments
- Background settings
- Daily and total study statistics
- Last login date and streak

The learning screen, Book editor, rank updates, study counts, and settings all read and write this local database. Normal use no longer requires Google Sheets.

## Moving existing data

When the app is still deployed through Google Apps Script, the start screen can show **Import Google Sheets once**. This copies every existing sheet into IndexedDB. After import, learning and editing use only the local copy.

The import is intentionally one-way. Later changes made in Google Sheets are not automatically synchronized.

## Backup and restore

Use **Export Backup** to download all Books and progress as a JSON file. Use **Import Backup** to restore that file on another browser or device.

Users should export backups regularly because browser storage can be removed when:

- The browser or app data is cleared
- The app is uninstalled
- The device is lost or reset
- Private browsing storage expires

## Native app migration

The Book document format is independent of IndexedDB. A future iOS or Android client can replace the storage adapter with SQLite while keeping the same Book structure and learning logic.

Cloud synchronization should be added separately after authentication and user-data isolation are implemented. A cloud database must not be introduced before each record can be securely associated with its owner.

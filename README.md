# Lenovo Case Tracker

A small desktop app I built to track Lenovo repair cases, parts, statuses, and notes.

This started as a simple CSV-based tracker for day-to-day field tech work. It has grown into a more useful repair tracking tool with search, filters, status updates, CSV import/export, and now SQLite storage.

## Current Branch

`sqlite-ui`

This branch moves the app from CSV storage to a local SQLite database and improves the layout so the UI works better when the window size or screen resolution changes.

## What It Does

- Tracks Lenovo cases by work order and serial number
- Stores repair status, parts, notes, and timestamps
- Supports searching and filtering entries
- Lets entries be added, edited, updated, and deleted
- Imports old CSV records into SQLite
- Exports current records back to CSV
- Uses a local SQLite database instead of relying on the CSV as the main data source
- Keeps database and CSV files out of Git so real repair data is not uploaded

## Why I Changed It

The original version worked, but CSV storage was starting to feel limiting. I had hundreds of entries from this year alone, and I wanted the app to behave more like a real database-backed application.

SQLite made sense because it is lightweight, local, easy to ship with a desktop app, and still lets me practice real SQL/database development.

## Tech Used

- Python
- PySide6
- SQLite
- CSV import/export
- Git/GitHub

## Notes

This is still a practical internal workflow tool first. I am using it as a way to improve the app while also practicing better software development habits: version control, database-backed storage, cleaner imports, and safer UI changes.

## Roadmap

- Clean up older CSV-related code
- Improve error handling during imports
- Add better duplicate detection
- Add status history
- Improve dashboard/reporting
- Add screenshots to the README
- Package the SQLite version as a new release
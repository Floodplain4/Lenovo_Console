# Lenovo Case Tracker

A desktop utility I built for tracking Lenovo repair cases, statuses, parts, notes, and repeat hardware issues.

This originally started as a basic CSV tracker for day-to-day field tech work and slowly turned into a more complete repair tracking application. The current version uses SQLite instead of relying entirely on CSV storage and includes a more responsive UI layout, repeat serial tracking, filtering, import/export support, and repair workflow management.

## Current Development Branch

`sqlite-ui`

This branch contains the ongoing SQLite/database migration work and UI cleanup before merging back into main.

## Features

- Track Lenovo repair cases by:
  - Work Order
  - Serial Number
  - Status
  - Parts
  - Notes
  - Timestamp

- Add, edit, update, and delete repair entries
- Search across all visible case data
- Filter by:
  - Status
  - Parts

- CSV import/export support
- SQLite database backend
- Responsive UI improvements for smaller/windowed displays
- Repeat serial number detection
- Dashboard statistics
- Follow-up tracking
- Email triage tools for ticket review workflow

## Repeat Serial Detection

One of the newer additions is repeat serial detection.

The app tracks devices that appear multiple times in the repair log and allows quick review of:
- Repeat repair counts
- Previous work orders
- Repair statuses
- Latest timestamps

This was added because recurring hardware issues became difficult to track manually once the dataset grew larger.

## Why SQLite

The original CSV version worked fine for smaller datasets, but it started becoming harder to manage once the repair log reached several hundred entries.

Migrating to SQLite made it easier to:
- Search and filter efficiently
- Handle larger datasets
- Reduce duplicate handling issues
- Move toward a more maintainable application structure
- Practice real database-backed application development

## Tech Used

- Python
- PySide6
- SQLite
- CSV import/export
- Git/GitHub

## Notes

This project is still primarily a practical internal workflow tool. Most changes are driven by real-world repair tracking and field tech workflow issues rather than building features just for the sake of adding them.

At the same time, I have been using the project to improve:
- Database design
- Version control workflow
- UI structure
- Application organization
- General software development practices

## Roadmap

- Clean up remaining CSV-era logic
- Improve database structure and normalization
- Add status history tracking
- Improve error handling and import validation
- Add screenshots/GIF demos to README
- Improve dashboard analytics
- Package updated SQLite version into a stable EXE release
- Explore possible future web-based version

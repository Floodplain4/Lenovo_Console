# Changelog

## [v2.4.6] - 2026-05-06

### Added
- Email Triage system for reviewing unread Outlook emails
- Pasted-text analysis for emails, tickets, and messages
- Support-likelihood scoring for triage results
- Detected part tags and suggested issue checks
- Ability to create case entries directly from emails
- Follow-up tracking for aging cases
- Manual follow-up marking
- 24-hour follow-up snooze
- Quick Stats dashboard panel
- Multi-case follow-up review workflow
- Compact application header

### Improved
- Dashboard layout and spacing
- Overall UI organization and workflow
- Email parsing and detection logic
- Outlook inbox scanning reliability
- Follow-up visibility and management
- Case review workflow
- Taskbar and application icon handling
- PyInstaller executable compatibility

### Fixed
- Fixed follow-up indicators interfering with CSV data
- Fixed right-click follow-up actions
- Fixed Outlook scanning issues in compiled executable builds
- Fixed taskbar/header icon behavior
- Fixed layout compression and spacing issues
- Fixed PyInstaller packaging issues with Outlook COM integration

---

## [v2.3] - 2026-04-30

### Added
- Improved Lenovo case tracking workflow
- Dashboard statistics panel
- Part tracking and status management
- Import/export support for CSV logs
- Duplicate entry detection
- Automatic backup behavior before imports
- Search and filtering functionality
- LCD script helper window
- Copy Case Summary functionality
- Double-click entry editing
- Multi-entry management workflow
- Context menu actions for serials and work orders

### Improved
- UI styling and dark theme consistency
- Workflow speed for field technician use
- CSV logging reliability
- Case status color coding
- Search behavior and match cycling
- Serial number normalization and case-insensitive parsing
- Settings persistence using QSettings

### Fixed
- Fixed duplicate entry handling
- Fixed timestamp update behavior
- Fixed PyInstaller taskbar icon issues
- Fixed settings persistence inconsistencies

---

## [v2.2] - 2026-04-20

### Added
- Transition to PySide6 desktop interface
- Improved dashboard layout
- Enhanced filtering and search tools
- Status tracking improvements
- About dialog and application versioning
- Resource path handling for packaged builds

### Improved
- General UI responsiveness
- Application layout density and spacing
- Packaging support for one-file executable builds

---

## [v2.1] - 2026-04-13

### Added
- Early PySide6 rewrite of the original tracker
- CSV-based local data storage
- Work order and serial number tracking
- Part selection system
- Status update workflow
- Timestamp logging
- Initial dashboard statistics
- Basic import/export support

### Improved
- Replaced earlier Tkinter-based interface with modern Qt UI
- Improved workflow organization for repair tracking

---

## [v1.0] - Initial Release

### Added
- Initial Lenovo repair case tracking workflow
- Simple CSV logging system
- Basic desktop interface
- Internal workflow tracking for Lenovo repair cases

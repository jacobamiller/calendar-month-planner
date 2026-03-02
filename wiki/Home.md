# Month Planner - Architecture Documentation

A standalone web app that displays a custom vertical month planner where each day is a row and each Google Calendar is a column. Includes country holiday tracking, trip idea auto-detection, reserved status management, and 2-way sync via Google Calendar.

## Table of Contents

- [Architecture Overview](Architecture-Overview)
- [Authentication & OAuth Flow](Authentication-OAuth-Flow)
- [Data Model & State Management](Data-Model-State-Management)
- [Sync Engine](Sync-Engine)
- [Rendering Pipeline](Rendering-Pipeline)
- [Column System](Column-System)
- [API Reference](API-Reference)

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Frontend | Vanilla HTML/CSS/JS (no framework) |
| Auth | Google Identity Services (GIS) |
| API | Google Calendar API v3 (REST) |
| Storage | localStorage + Google Calendar sync |
| Hosting | Any static server (`python3 -m http.server 8000`) |

## File Structure

```
calendar-view/
├── index.html    # Layout, CSS styles, HTML structure
├── app.js        # Auth, API, rendering, sync logic (~1050 lines)
├── config.js     # OAuth Client ID and scopes
└── wiki/         # This documentation
```

# Architecture Overview

## High-Level System Diagram

```mermaid
graph TB
    subgraph Browser
        UI[index.html<br>Layout & CSS]
        APP[app.js<br>Core Logic]
        CFG[config.js<br>OAuth Config]
        LS[(localStorage<br>Preferences & Cache)]
    end

    subgraph Google Cloud
        GIS[Google Identity Services<br>OAuth 2.0]
        GCAL[Google Calendar API v3]
        subgraph User Calendars
            UC1[Personal Calendar]
            UC2[Work Calendar]
            UC3[Other Calendars]
        end
        subgraph Holiday Calendars
            HC1[Vietnam Holidays]
            HC2[Thailand Holidays]
            HC3[China Holidays]
            HC4[Mexico Holidays]
            HC5[US Holidays]
        end
        SYNC[Month Planner Sync<br>Calendar]
    end

    UI --> APP
    CFG --> APP
    APP <--> LS
    APP --> GIS
    GIS --> APP
    APP -->|READ| GCAL
    APP <-->|READ/WRITE| SYNC
    GCAL --> UC1 & UC2 & UC3
    GCAL --> HC1 & HC2 & HC3 & HC4 & HC5
```

## Application Flow

```mermaid
sequenceDiagram
    participant U as User
    participant A as App (app.js)
    participant G as Google OAuth
    participant C as Calendar API
    participant S as Sync Calendar
    participant L as localStorage

    U->>A: Click "Sign in"
    A->>G: requestAccessToken({prompt: 'consent'})
    G-->>U: Consent screen
    U-->>G: Approve
    G-->>A: access_token
    A->>C: GET /calendarList
    C-->>A: All calendars
    A->>A: Match holiday calendars
    A->>S: Find or create "Month Planner Sync"
    S-->>A: syncCalId
    A->>A: renderCalendarCheckboxes()
    A->>A: renderColumnCheckboxes()
    A->>A: loadMonth()
    A->>S: GET /events (sync data)
    S-->>A: Reserved levels & notes
    A->>L: Overwrite localStorage
    A->>C: GET /events (holidays, parallel)
    A->>C: GET /events (user calendars, parallel)
    C-->>A: Events data
    A->>A: Scan for "trip idea" events
    A->>A: renderGrid()
    A->>A: renderLegend()
    A-->>U: Month grid displayed

    Note over U,L: User edits reserved/notes
    U->>A: Click reserved cell → Save
    A->>L: Save to localStorage
    A->>S: PATCH/POST sync event
    S-->>A: Confirmed
    A-->>U: "Synced" status
```

## Component Architecture

```mermaid
graph LR
    subgraph Initialization
        INIT[init] --> AUTH[initAuth]
        AUTH --> TOKEN[onTokenResponse]
        TOKEN --> FETCH_CAL[fetchCalendars]
    end

    subgraph Calendar Setup
        FETCH_CAL --> RESOLVE[Resolve Holiday IDs]
        FETCH_CAL --> ENSURE_SYNC[ensureSyncCalendar]
        FETCH_CAL --> RENDER_CB[renderCalendarCheckboxes]
        FETCH_CAL --> RENDER_COL[renderColumnCheckboxes]
    end

    subgraph Month Loading
        FETCH_CAL --> LOAD[loadMonth]
        LOAD --> FETCH_SYNC[fetchSyncEvents]
        LOAD --> FETCH_HOL[fetchHolidayEvents]
        LOAD --> FETCH_EV[fetchEvents]
        LOAD --> SCAN_TRIP[Scan Trip Ideas]
        LOAD --> GRID[renderGrid]
        LOAD --> LEGEND[renderLegend]
    end

    subgraph User Interactions
        GRID --> PICKER[Reserved Picker]
        GRID --> NOTES[Notes Input]
        PICKER --> UPSERT[syncUpsertEvent]
        NOTES --> UPSERT
    end
```

## Design Principles

1. **No build step** — Vanilla JS, load directly in browser
2. **Minimal permissions** — Only requests what's needed from Google
3. **Offline-capable** — localStorage provides immediate data, sync is additive
4. **Single-page** — All rendering happens client-side via DOM manipulation
5. **Progressive enhancement** — Works read-only without sync; sync adds persistence

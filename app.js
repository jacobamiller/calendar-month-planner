// ── State ──
let accessToken = null;
let tokenClient = null;
let allCalendars = [];       // from API
let selectedCalendarIds = []; // user-chosen subset
let currentYear, currentMonth; // 0-indexed month
let eventsCache = {};         // "YYYY-MM" → { calId → { "YYYY-MM-DD" → [events] } }
let holidayCache = {};        // "YYYY-MM-calId" → { "YYYY-MM-DD" → ["name"] }
let tripIdeaDates = {};       // "YYYY-MM-DD" → true (auto-detected from events)
let syncCalId = null;         // ID of the "Month Planner Sync" calendar
let syncEventIds = {};        // "YYYY-MM-DD" → event ID on the sync calendar
let syncReady = false;        // true once sync calendar is found/created

// Country columns: name + possible Google Calendar ID patterns
const COUNTRY_COLUMNS = [
  { name: 'Vietnam',  match: ['vietnam', 'vietnamese'], fallbackId: 'en.vietnamese#holiday@group.v.calendar.google.com' },
  { name: 'Thailand', match: ['thailand', 'thai'],      fallbackId: 'en.thai#holiday@group.v.calendar.google.com' },
  { name: 'China',    match: ['china', 'chinese'],      fallbackId: 'en.china#holiday@group.v.calendar.google.com' },
  { name: 'Mexico',   match: ['mexico', 'mexican'],     fallbackId: 'en.mexican#holiday@group.v.calendar.google.com' },
  { name: 'US',       match: ['usa', ' us', 'united states', 'american'], fallbackId: 'en.usa#holiday@group.v.calendar.google.com' },
];

// Resolved at runtime after fetching calendar list
let HOLIDAY_CAL_IDS = new Set();

// Reserved column opacity levels: click cycles through these
const RESERVED_LEVELS = [0, 0.25, 0.50, 0.75, 1.0];
const RESERVED_LABELS = ['', 'Planning', 'Considering', 'Confident', 'Reserved'];

// ── DOM refs ──
const monthTitle = document.getElementById('month-title');
const mainContent = document.getElementById('main-content');
const signInPrompt = document.getElementById('sign-in-prompt');
const authBtn = document.getElementById('auth-btn');
const prevBtn = document.getElementById('prev-btn');
const nextBtn = document.getElementById('next-btn');
const todayBtn = document.getElementById('today-btn');
const columnsToggle = document.getElementById('columns-toggle');
const columnsPanel = document.getElementById('columns-panel');
const columnCheckboxes = document.getElementById('column-checkboxes');
const settingsToggle = document.getElementById('settings-toggle');
const settingsPanel = document.getElementById('settings-panel');
const calendarCheckboxes = document.getElementById('calendar-checkboxes');
const legendEl = document.getElementById('legend');
const tooltipEl = document.getElementById('tooltip');

// Hidden columns set — persisted to localStorage
let hiddenColumns = new Set(JSON.parse(localStorage.getItem('mp_hiddenCols') || '[]'));

// Column order — persisted to localStorage
let columnOrder = JSON.parse(localStorage.getItem('mp_colOrder') || '[]');

// ── Init ──
(function init() {
  const saved = localStorage.getItem('mp_lastMonth');
  if (saved) {
    const [y, m] = saved.split('-').map(Number);
    currentYear = y;
    currentMonth = m;
  } else {
    const now = new Date();
    currentYear = now.getFullYear();
    currentMonth = now.getMonth();
  }
  updateTitle();

  window.addEventListener('load', () => {
    if (typeof google !== 'undefined' && google.accounts) {
      initAuth();
    } else {
      const interval = setInterval(() => {
        if (typeof google !== 'undefined' && google.accounts) {
          clearInterval(interval);
          initAuth();
        }
      }, 200);
    }
  });
})();

// ── Auth ──
function initAuth() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: onTokenResponse,
  });
  authBtn.onclick = () => {
    if (accessToken) {
      signOut();
    } else {
      tokenClient.requestAccessToken({ prompt: 'consent' });
    }
  };
}

function onTokenResponse(resp) {
  if (resp.error) {
    console.error('Auth error:', resp);
    return;
  }
  accessToken = resp.access_token;
  authBtn.textContent = 'Sign out';
  signInPrompt.style.display = 'none';
  fetchCalendars();
}

function signOut() {
  if (accessToken) {
    google.accounts.oauth2.revoke(accessToken);
  }
  accessToken = null;
  allCalendars = [];
  selectedCalendarIds = [];
  eventsCache = {};
  tripIdeaDates = {};
  syncCalId = null;
  syncEventIds = {};
  syncReady = false;
  updateSyncStatus('');
  authBtn.textContent = 'Sign in with Google';
  settingsPanel.classList.remove('open');
  calendarCheckboxes.innerHTML = '';
  legendEl.innerHTML = '';
  mainContent.innerHTML = '';
  signInPrompt.style.display = '';
  mainContent.appendChild(signInPrompt);
}

// ── Fetch Calendars ──
async function fetchCalendars() {
  const resp = await apiFetch('https://www.googleapis.com/calendar/v3/users/me/calendarList');
  if (!resp) return;
  const allItems = (resp.items || []);

  // Resolve country holiday calendar IDs from user's subscribed calendars
  HOLIDAY_CAL_IDS = new Set();
  COUNTRY_COLUMNS.forEach(cc => {
    // Search user's calendar list for a matching holiday calendar
    const found = allItems.find(c => {
      const id = (c.id || '').toLowerCase();
      const name = (c.summary || '').toLowerCase();
      return cc.match.some(m => id.includes(m) || name.includes(m));
    });
    if (found) {
      cc.calId = found.id;
      console.log(`Matched ${cc.name} → ${found.id} (${found.summary})`);
    } else {
      cc.calId = cc.fallbackId;
      console.log(`No match for ${cc.name}, using fallback: ${cc.fallbackId}`);
    }
    HOLIDAY_CAL_IDS.add(cc.calId);
  });

  allCalendars = allItems
    .filter(c => c.accessRole !== 'freeBusyReader')
    .sort((a, b) => (a.summary || '').localeCompare(b.summary || ''));

  const saved = localStorage.getItem('mp_selectedCals');
  if (saved) {
    const ids = JSON.parse(saved);
    selectedCalendarIds = ids.filter(id => allCalendars.some(c => c.id === id));
  } else {
    selectedCalendarIds = allCalendars.map(c => c.id);
  }

  await ensureSyncCalendar(allItems);

  renderCalendarCheckboxes();
  renderColumnCheckboxes();
  loadMonth();
}

async function ensureSyncCalendar(calendarItems) {
  const cached = localStorage.getItem('mp_syncCalId');
  console.log('ensureSyncCalendar: cached ID =', cached, ', calendars count =', calendarItems.length);

  const found = calendarItems.find(c =>
    c.id === cached ||
    (c.summary === 'Month Planner Sync' && (c.description || '').includes('Month Planner app'))
  );

  if (found) {
    syncCalId = found.id;
    localStorage.setItem('mp_syncCalId', syncCalId);
    syncReady = true;
    updateSyncStatus('synced');
    console.log('Sync calendar found:', syncCalId);
    return;
  }

  console.log('ensureSyncCalendar: not found, creating...');
  try {
    const resp = await fetch('https://www.googleapis.com/calendar/v3/calendars', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + accessToken,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        summary: 'Month Planner Sync',
        description: 'Auto-created by Month Planner app. Stores reserved levels and notes.',
      }),
    });
    if (resp.ok) {
      const cal = await resp.json();
      syncCalId = cal.id;
      localStorage.setItem('mp_syncCalId', syncCalId);
      syncReady = true;
      updateSyncStatus('synced');
      console.log('Sync calendar created:', syncCalId);
    } else {
      const errText = await resp.text();
      console.error('Failed to create sync calendar:', resp.status, errText);
      updateSyncStatus('error');
    }
  } catch (err) {
    console.error('Error creating sync calendar:', err);
    updateSyncStatus('error');
  }
}

function renderCalendarCheckboxes() {
  calendarCheckboxes.innerHTML = '';
  allCalendars.filter(c => !HOLIDAY_CAL_IDS.has(c.id) && !isHolidayCalendar(c)).forEach(cal => {
    const label = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = selectedCalendarIds.includes(cal.id);
    cb.onchange = () => {
      if (cb.checked) {
        if (!selectedCalendarIds.includes(cal.id)) selectedCalendarIds.push(cal.id);
      } else {
        selectedCalendarIds = selectedCalendarIds.filter(id => id !== cal.id);
      }
      localStorage.setItem('mp_selectedCals', JSON.stringify(selectedCalendarIds));
      loadMonth();
    };

    const dot = document.createElement('span');
    dot.className = 'cal-color-dot';
    dot.style.background = cal.backgroundColor || '#4285f4';

    label.appendChild(cb);
    label.appendChild(dot);
    label.appendChild(document.createTextNode(' ' + (cal.summary || cal.id)));
    calendarCheckboxes.appendChild(label);
  });
}

function getDefaultColumnList() {
  return [
    { key: 'col_reserved', label: 'Reserved', type: 'fixed' },
    ...COUNTRY_COLUMNS.map(cc => ({ key: 'col_' + cc.name, label: cc.name, type: 'country' })),
    ...allCalendars
      .filter(c => !HOLIDAY_CAL_IDS.has(c.id) && !isHolidayCalendar(c))
      .map(c => ({ key: 'col_cal_' + c.id, label: c.summaryOverride || c.summary || c.id, type: 'calendar' })),
    { key: 'col_notes', label: 'Notes', type: 'fixed' },
  ];
}

function getOrderedColumns() {
  const defaults = getDefaultColumnList();
  if (columnOrder.length === 0) return defaults;

  // Build ordered list from saved order, add any new columns at the end
  const byKey = {};
  defaults.forEach(c => { byKey[c.key] = c; });
  const ordered = [];
  columnOrder.forEach(key => {
    if (byKey[key]) {
      ordered.push(byKey[key]);
      delete byKey[key];
    }
  });
  // Append any new columns not in saved order
  Object.values(byKey).forEach(c => ordered.push(c));
  return ordered;
}

function saveColumnOrder(columns) {
  columnOrder = columns.map(c => c.key);
  localStorage.setItem('mp_colOrder', JSON.stringify(columnOrder));
}

function renderColumnCheckboxes() {
  columnCheckboxes.innerHTML = '';
  const columns = getOrderedColumns();

  columns.forEach((col, i) => {
    const row = document.createElement('div');
    row.style.cssText = 'display:flex;align-items:center;gap:4px;padding:2px 0;';

    const upBtn = document.createElement('button');
    upBtn.textContent = '\u25B2';
    upBtn.style.cssText = 'font-size:9px;padding:0 3px;cursor:pointer;border:1px solid #ccc;background:#fff;border-radius:2px;';
    upBtn.disabled = i === 0;
    upBtn.onclick = () => { moveColumn(i, -1); };

    const downBtn = document.createElement('button');
    downBtn.textContent = '\u25BC';
    downBtn.style.cssText = 'font-size:9px;padding:0 3px;cursor:pointer;border:1px solid #ccc;background:#fff;border-radius:2px;';
    downBtn.disabled = i === columns.length - 1;
    downBtn.onclick = () => { moveColumn(i, 1); };

    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = !hiddenColumns.has(col.key);
    cb.onchange = () => {
      if (cb.checked) {
        hiddenColumns.delete(col.key);
      } else {
        hiddenColumns.add(col.key);
      }
      localStorage.setItem('mp_hiddenCols', JSON.stringify([...hiddenColumns]));
      loadMonth();
    };

    const lbl = document.createElement('span');
    lbl.textContent = col.label;
    lbl.style.fontSize = '13px';

    row.appendChild(upBtn);
    row.appendChild(downBtn);
    row.appendChild(cb);
    row.appendChild(lbl);
    columnCheckboxes.appendChild(row);
  });
}

function moveColumn(index, direction) {
  const columns = getOrderedColumns();
  const newIndex = index + direction;
  if (newIndex < 0 || newIndex >= columns.length) return;
  const temp = columns[index];
  columns[index] = columns[newIndex];
  columns[newIndex] = temp;
  saveColumnOrder(columns);
  renderColumnCheckboxes();
  loadMonth();
}

// ── Month Navigation ──
prevBtn.onclick = () => { shiftMonth(-1); };
nextBtn.onclick = () => { shiftMonth(1); };
todayBtn.onclick = () => {
  const now = new Date();
  currentYear = now.getFullYear();
  currentMonth = now.getMonth();
  updateTitle();
  if (accessToken) loadMonth();
};

columnsToggle.onclick = () => {
  columnsPanel.classList.toggle('open');
  settingsPanel.classList.remove('open');
};

settingsToggle.onclick = () => {
  settingsPanel.classList.toggle('open');
  columnsPanel.classList.remove('open');
};

function shiftMonth(delta) {
  currentMonth += delta;
  if (currentMonth < 0) { currentMonth = 11; currentYear--; }
  if (currentMonth > 11) { currentMonth = 0; currentYear++; }
  updateTitle();
  if (accessToken) loadMonth();
}

function updateTitle() {
  const names = ['January','February','March','April','May','June',
    'July','August','September','October','November','December'];
  monthTitle.textContent = `${names[currentMonth]} ${currentYear}`;
  localStorage.setItem('mp_lastMonth', `${currentYear}-${currentMonth}`);
}

// ── Holidays (via Google Calendar) ──
async function fetchHolidayEvents(calId, timeMin, timeMax) {
  const cacheKey = `${timeMin}-${calId}`;
  if (holidayCache[cacheKey]) return holidayCache[cacheKey];

  const events = await fetchEvents(calId, timeMin, timeMax);
  console.log(`Holidays for ${calId}: ${events.length} events`);
  const map = {};
  events.forEach(ev => {
    const start = ev.start.date || ev.start.dateTime.split('T')[0];
    const end = ev.end.date || ev.end.dateTime.split('T')[0];
    const name = ev.summary || 'Holiday';

    // Expand multi-day holidays
    const d = new Date(start + 'T00:00:00');
    const endDate = new Date(end + 'T00:00:00');
    // end date is exclusive for all-day events
    while (d < endDate) {
      const dk = dateKey(d);
      if (!map[dk]) map[dk] = [];
      map[dk].push(name);
      d.setDate(d.getDate() + 1);
    }
  });

  holidayCache[cacheKey] = map;
  return map;
}

// ── Sync: fetch remote data for a month ──
async function fetchSyncEvents(timeMin, timeMax) {
  if (!syncReady || !syncCalId) return;

  syncEventIds = {};

  const params = new URLSearchParams({
    timeMin,
    timeMax,
    singleEvents: 'true',
    maxResults: '250',
    privateExtendedProperty: 'mpApp=monthplanner',
  });

  const url = `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(syncCalId)}/events?${params}`;
  const data = await apiFetch(url);
  if (!data || !data.items) return;

  data.items.forEach(ev => {
    const dk = ev.start.date;
    if (!dk) return;

    const props = (ev.extendedProperties && ev.extendedProperties.private) || {};
    const reserved = parseInt(props.mpReserved || '0', 10);
    const note = props.mpNote || '';

    syncEventIds[dk] = ev.id;

    if (reserved > 0) {
      localStorage.setItem('mp_reserved_' + dk, reserved);
    } else {
      localStorage.removeItem('mp_reserved_' + dk);
    }

    if (note) {
      localStorage.setItem('mp_note_' + dk, note);
    } else {
      localStorage.removeItem('mp_note_' + dk);
    }
  });

  console.log(`Synced ${Object.keys(syncEventIds).length} events from remote`);
}

// ── Sync: write-through on edit ──
async function syncUpsertEvent(dk) {
  if (!syncReady || !syncCalId) return;
  updateSyncStatus('saving');

  const reserved = parseInt(localStorage.getItem('mp_reserved_' + dk) || '0', 10);
  const note = localStorage.getItem('mp_note_' + dk) || '';

  if (reserved === 0 && !note) {
    await syncDeleteEvent(dk);
    return;
  }

  let summary = '';
  if (reserved > 0) summary = 'R' + reserved + ': ' + RESERVED_LABELS[reserved];
  if (note) summary = summary ? summary + ' | ' + note.substring(0, 40) : 'Note';

  const eventBody = {
    summary,
    description: note || '',
    start: { date: dk },
    end: { date: nextDay(dk) },
    extendedProperties: {
      private: {
        mpApp: 'monthplanner',
        mpReserved: String(reserved),
        mpNote: note,
      },
    },
  };

  const existingId = syncEventIds[dk];

  try {
    let resp;
    if (existingId) {
      resp = await fetch(
        `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(syncCalId)}/events/${encodeURIComponent(existingId)}`,
        {
          method: 'PATCH',
          headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
          body: JSON.stringify(eventBody),
        }
      );
    } else {
      resp = await fetch(
        `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(syncCalId)}/events`,
        {
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
          body: JSON.stringify(eventBody),
        }
      );
    }

    if (resp.ok) {
      const ev = await resp.json();
      syncEventIds[dk] = ev.id;
      updateSyncStatus('synced');
    } else {
      console.error('Sync write failed:', resp.status, await resp.text());
      updateSyncStatus('error');
    }
  } catch (err) {
    console.error('Sync write error:', err);
    updateSyncStatus('error');
  }
}

async function syncDeleteEvent(dk) {
  if (!syncReady || !syncCalId) return;
  const existingId = syncEventIds[dk];
  if (!existingId) { updateSyncStatus('synced'); return; }

  try {
    await fetch(
      `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(syncCalId)}/events/${encodeURIComponent(existingId)}`,
      {
        method: 'DELETE',
        headers: { 'Authorization': 'Bearer ' + accessToken },
      }
    );
    delete syncEventIds[dk];
    updateSyncStatus('synced');
  } catch (err) {
    console.error('Sync delete error:', err);
    updateSyncStatus('error');
  }
}

function nextDay(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() + 1);
  return dateKey(d);
}

function updateSyncStatus(state) {
  const el = document.getElementById('sync-status');
  if (!el) return;
  if (state === 'saving') { el.textContent = 'Saving...'; el.style.color = '#f59e0b'; }
  else if (state === 'synced') { el.textContent = 'Synced'; el.style.color = '#16a34a'; }
  else if (state === 'error') { el.textContent = 'Sync error'; el.style.color = '#dc2626'; }
  else { el.textContent = ''; }
}

// ── Load Events for Month ──
async function loadMonth() {
  const key = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;
  // Filter out holiday calendars from the main columns
  const selectedCals = allCalendars
    .filter(c => selectedCalendarIds.includes(c.id))
    .filter(c => !HOLIDAY_CAL_IDS.has(c.id) && !isHolidayCalendar(c));

  mainContent.innerHTML = '<div class="loading">Loading events...</div>';

  const timeMin = new Date(currentYear, currentMonth, 1).toISOString();
  const timeMax = new Date(currentYear, currentMonth + 1, 0, 23, 59, 59).toISOString();

  // Pull sync data from remote (remote wins)
  await fetchSyncEvents(timeMin, timeMax);

  // Fetch holidays from Google Calendar holiday calendars (parallel)
  const holidayMaps = {};
  await Promise.all(COUNTRY_COLUMNS.map(async cc => {
    holidayMaps[cc.calId] = await fetchHolidayEvents(cc.calId, timeMin, timeMax);
  }));

  // Fetch user calendar events (use cache if available)
  let calEvents;
  if (eventsCache[key]) {
    calEvents = eventsCache[key];
  } else {
    calEvents = {};
    await Promise.all(selectedCals.map(async cal => {
      const events = await fetchEvents(cal.id, timeMin, timeMax);
      calEvents[cal.id] = indexEventsByDate(events, currentYear, currentMonth);
    }));

    eventsCache[key] = calEvents;
  }

  // Scan all events for "trip idea" → extract % for Reserved, note title on first day
  tripIdeaDates = {};   // dk → { level: 1-4, title: string, isFirstDay: bool }
  Object.values(calEvents).forEach(dateMap => {
    Object.entries(dateMap).forEach(([dk, events]) => {
      events.forEach(ev => {
        if (!ev.summary) return;
        if (!ev.summary.toLowerCase().includes('trip idea')) return;

        // Parse percentage from anywhere in title (e.g. "trip ideas - 75%" or "50% trip idea Tokyo")
        const pctMatch = ev.summary.match(/(\d+)\s*%/);
        let level = 1; // default 25% Planning
        if (pctMatch) {
          const pct = parseInt(pctMatch[1], 10);
          if (pct >= 100) level = 4;
          else if (pct >= 75) level = 3;
          else if (pct >= 50) level = 2;
          else level = 1;
        }

        // Determine if this is the first day of the event
        const evStart = ev.start.date
          ? ev.start.date
          : ev.start.dateTime.split('T')[0];
        const isFirstDay = (dk === evStart);

        // Keep highest level if multiple trip ideas on same day
        if (!tripIdeaDates[dk] || level > tripIdeaDates[dk].level) {
          tripIdeaDates[dk] = { level, title: ev.summary, isFirstDay };
        } else if (isFirstDay && !tripIdeaDates[dk].isFirstDay) {
          // Prefer the entry that is the first day for notes
          tripIdeaDates[dk] = { level: Math.max(level, tripIdeaDates[dk].level), title: ev.summary, isFirstDay };
        }
      });
    });
  });

  renderGrid(selectedCals, calEvents, holidayMaps);
  renderLegend(selectedCals);
}

// Detect any holiday calendar by ID pattern or name
function isHolidayCalendar(cal) {
  const id = (cal.id || '').toLowerCase();
  return id.includes('#holiday@group.v.calendar.google.com');
}

async function fetchEvents(calendarId, timeMin, timeMax) {
  const params = new URLSearchParams({
    timeMin, timeMax,
    singleEvents: 'true',
    orderBy: 'startTime',
    maxResults: '2500',
  });
  const url = `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(calendarId)}/events?${params}`;
  const data = await apiFetch(url);
  return data ? (data.items || []) : [];
}

function indexEventsByDate(events, year, month) {
  const map = {};
  const monthStart = new Date(year, month, 1);
  const monthEnd = new Date(year, month + 1, 0);

  events.forEach(ev => {
    const start = ev.start.dateTime ? new Date(ev.start.dateTime) : new Date(ev.start.date + 'T00:00:00');
    const end = ev.end.dateTime ? new Date(ev.end.dateTime) : new Date(ev.end.date + 'T00:00:00');

    const endAdjusted = ev.end.date && !ev.end.dateTime
      ? new Date(end.getTime() - 86400000)
      : end;

    const d = new Date(Math.max(start.getTime(), monthStart.getTime()));
    const last = new Date(Math.min(endAdjusted.getTime(), monthEnd.getTime()));

    while (d <= last) {
      const key = dateKey(d);
      if (!map[key]) map[key] = [];
      map[key].push(ev);
      d.setDate(d.getDate() + 1);
    }
  });
  return map;
}

// ── Render Grid ──
function renderGrid(calendars, calEvents, holidayMaps) {
  const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
  const dayNames = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

  // Build ordered visible columns
  const orderedCols = getOrderedColumns().filter(c => !hiddenColumns.has(c.key));

  // Sizing: fixed cols (reserved/country) together = 1 share, each calendar = 1 share, notes = 2 shares
  const fixedCount = orderedCols.filter(c => c.type === 'fixed' && c.key !== 'col_notes' || c.type === 'country').length;
  const calCount = orderedCols.filter(c => c.type === 'calendar').length;
  const hasNotes = orderedCols.some(c => c.key === 'col_notes');
  const hasReserved = orderedCols.some(c => c.key === 'col_reserved');

  const totalShares = (fixedCount > 0 ? 1 : 0) + calCount + (hasNotes ? 2 : 0);
  const oneShare = totalShares > 0 ? (100 / totalShares) : 10;
  const fixedGroup = Math.max(3, Math.min(20, oneShare));
  const remaining = 100 - fixedGroup;
  const calShares = calCount + (hasNotes ? 2 : 0);
  const calPct = calShares > 0 ? (remaining / calShares) : 5;
  const notesPct = calPct * 2;

  // Split fixed group
  const countryCount = orderedCols.filter(c => c.type === 'country').length || 1;
  const countryShare = (fixedGroup * 0.5) / countryCount;
  const datePct = fixedGroup * 0.25;
  const reservedPct = fixedGroup * 0.25;

  // Build calendarById lookup
  const calById = {};
  calendars.forEach(c => { calById[c.id] = c; });
  const countryByKey = {};
  COUNTRY_COLUMNS.forEach(cc => { countryByKey['col_' + cc.name] = cc; });

  let html = '<div class="grid-container"><table class="month-grid">';

  // Colgroup
  html += '<colgroup>';
  html += `<col style="width:${datePct}%">`; // date always first
  orderedCols.forEach(col => {
    if (col.key === 'col_reserved') html += `<col style="width:${reservedPct}%">`;
    else if (col.type === 'country') html += `<col style="width:${countryShare}%">`;
    else if (col.key === 'col_notes') html += `<col style="width:${notesPct}%">`;
    else if (col.type === 'calendar') html += `<col style="width:${calPct}%">`;
  });
  html += '</colgroup>';

  // Header row
  html += '<thead><tr class="header-row">';
  html += '<th class="date-header">Date</th>';
  orderedCols.forEach(col => {
    if (col.key === 'col_reserved')
      html += '<th class="fixed-col-header reserved-header"><span class="angled-header">Reserved</span></th>';
    else if (col.type === 'country')
      html += `<th class="fixed-col-header"><span class="angled-header">${esc(col.label)}</span></th>`;
    else if (col.key === 'col_notes')
      html += '<th class="notes-header">Notes</th>';
    else if (col.type === 'calendar')
      html += `<th class="cal-col"><span class="angled-header" title="${esc(col.label)}">${esc(col.label)}</span></th>`;
  });
  html += '</tr></thead>';

  // Body rows
  html += '<tbody>';
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(currentYear, currentMonth, day);
    const dow = date.getDay();
    const isWeekend = dow === 0 || dow === 6;
    const dk = dateKey(date);
    const tripInfo = tripIdeaDates[dk];

    html += `<tr class="${isWeekend ? 'weekend' : ''}">`;
    html += `<td class="date-cell">${dayNames[dow]} ${String(day).padStart(2, '\u00A0')}</td>`;

    orderedCols.forEach(col => {
      if (col.key === 'col_reserved') {
        const savedLevel = localStorage.getItem('mp_reserved_' + dk);
        let level = savedLevel !== null ? parseInt(savedLevel, 10) : (tripInfo ? tripInfo.level : 0);
        const opacity = RESERVED_LEVELS[level];
        const label = RESERVED_LABELS[level];
        html += `<td class="fixed-col reserved-cell" data-date="${dk}" data-level="${level}" data-tip="${label}" style="cursor:pointer">`;
        if (opacity > 0) html += `<div class="reserved-block" style="background:rgba(0,0,0,${opacity})"></div>`;
        html += '</td>';
      } else if (col.type === 'country') {
        const cc = countryByKey[col.key];
        const holidays = cc ? (holidayMaps[cc.calId] || {}) : {};
        const names = holidays[dk];
        if (names && names.length > 0) {
          html += `<td class="fixed-col holiday-cell" data-tip="${esc(names.join(', '))}"><div class="holiday-block"></div></td>`;
        } else {
          html += '<td class="fixed-col"></td>';
        }
      } else if (col.key === 'col_notes') {
        const manualNote = localStorage.getItem('mp_note_' + dk);
        let noteVal = manualNote || '';
        if (!manualNote && tripInfo && tripInfo.isFirstDay) noteVal = tripInfo.title;
        const isAuto = !manualNote && noteVal;
        const noteStyle = isAuto ? 'color:#999;font-style:italic' : 'color:#333';
        html += `<td class="notes-cell"><input type="text" value="${esc(noteVal)}" data-date="${dk}" data-manual="${manualNote ? '1' : '0'}" style="${noteStyle}" /></td>`;
      } else if (col.type === 'calendar') {
        const calId = col.key.replace('col_cal_', '');
        const cal = calById[calId];
        const events = (calEvents[calId] && calEvents[calId][dk]) || [];
        const color = cal ? (cal.backgroundColor || '#4285f4') : '#4285f4';
        if (events.length > 0) {
          const titles = events.map(e => {
            let t = e.summary || '(No title)';
            if (e.start.dateTime) t += ' \u2022 ' + formatTime(e.start.dateTime);
            return t;
          }).join('\n');
          html += `<td class="event-cell" data-tip="${esc(titles)}"><div class="event-block" style="background:${color}"></div></td>`;
        } else {
          html += '<td class="event-cell"></td>';
        }
      }
    });

    html += '</tr>';
  }
  html += '</tbody></table></div>';

  mainContent.innerHTML = html;

  // Bind Reserved column — click to open picker popup with save
  mainContent.querySelectorAll('.reserved-cell').forEach(cell => {
    cell.addEventListener('click', (e) => {
      e.stopPropagation();
      const old = document.getElementById('reserved-picker');
      if (old) old.remove();

      const dk = cell.dataset.date;
      const currentLevel = parseInt(cell.dataset.level, 10);
      const tripInfo = tripIdeaDates[dk];
      const autoLevel = tripInfo ? tripInfo.level : 0;
      const hasSavedOverride = localStorage.getItem('mp_reserved_' + dk) !== null;
      let selectedLevel = currentLevel;

      const picker = document.createElement('div');
      picker.id = 'reserved-picker';
      picker.className = 'reserved-picker';

      // Info header: show auto-detected and current
      const info = document.createElement('div');
      info.className = 'picker-info';
      const autoLabel = autoLevel > 0 ? RESERVED_LABELS[autoLevel] + ' (' + (RESERVED_LEVELS[autoLevel] * 100) + '%)' : 'None';
      info.innerHTML = '<span style="color:#888">Auto: ' + autoLabel + '</span>';
      if (hasSavedOverride) {
        info.innerHTML += '<br><span style="color:#1a73e8">Saved: ' + RESERVED_LABELS[currentLevel] + '</span>';
      }
      picker.appendChild(info);

      const options = [
        { lvl: 0, label: 'Clear', opacity: 0 },
        { lvl: 1, label: 'Planning 25%', opacity: 0.25 },
        { lvl: 2, label: 'Considering 50%', opacity: 0.50 },
        { lvl: 3, label: 'Confident 75%', opacity: 0.75 },
        { lvl: 4, label: 'Reserved 100%', opacity: 1.0 },
      ];

      const buttons = [];
      options.forEach(opt => {
        const btn = document.createElement('button');
        btn.className = 'picker-btn';
        if (currentLevel === opt.lvl) btn.classList.add('picker-active');
        const swatch = document.createElement('span');
        swatch.className = 'picker-swatch';
        swatch.style.background = opt.opacity > 0 ? `rgba(0,0,0,${opt.opacity})` : '#fff';
        swatch.style.border = '1px solid #999';
        btn.appendChild(swatch);
        btn.appendChild(document.createTextNode(' ' + opt.label));
        btn.onclick = (ev) => {
          ev.stopPropagation();
          selectedLevel = opt.lvl;
          buttons.forEach(b => b.classList.remove('picker-selected'));
          btn.classList.add('picker-selected');
          saveBtn.disabled = false;
          saveBtn.style.opacity = '1';
        };
        picker.appendChild(btn);
        buttons.push(btn);
      });

      // Action buttons row
      const actions = document.createElement('div');
      actions.className = 'picker-actions';

      // Reset to auto button
      if (hasSavedOverride && autoLevel !== currentLevel) {
        const resetBtn = document.createElement('button');
        resetBtn.className = 'picker-action-btn';
        resetBtn.textContent = 'Reset to auto';
        resetBtn.onclick = (ev) => {
          ev.stopPropagation();
          localStorage.removeItem('mp_reserved_' + dk);
          applyReservedLevel(cell, autoLevel);
          picker.remove();
          syncUpsertEvent(dk);
        };
        actions.appendChild(resetBtn);
      }

      // Save button
      const saveBtn = document.createElement('button');
      saveBtn.className = 'picker-action-btn picker-save-btn';
      saveBtn.textContent = 'Save';
      saveBtn.disabled = true;
      saveBtn.style.opacity = '0.4';
      saveBtn.onclick = (ev) => {
        ev.stopPropagation();
        if (selectedLevel === 0) {
          localStorage.removeItem('mp_reserved_' + dk);
        } else {
          localStorage.setItem('mp_reserved_' + dk, selectedLevel);
        }
        applyReservedLevel(cell, selectedLevel);
        picker.remove();
        syncUpsertEvent(dk);
      };
      actions.appendChild(saveBtn);

      picker.appendChild(actions);

      // Position near the cell
      const rect = cell.getBoundingClientRect();
      picker.style.top = (rect.bottom + window.scrollY + 2) + 'px';
      picker.style.left = (rect.left + window.scrollX) + 'px';
      document.body.appendChild(picker);

      // Close on outside click (no save)
      setTimeout(() => {
        document.addEventListener('click', function closePicker() {
          picker.remove();
          document.removeEventListener('click', closePicker);
        }, { once: true });
      }, 0);
    });
  });

  // Bind note saving — style manual notes differently
  mainContent.querySelectorAll('.notes-cell input').forEach(input => {
    input.addEventListener('input', () => {
      const dk = input.dataset.date;
      if (input.value) {
        localStorage.setItem('mp_note_' + dk, input.value);
      } else {
        localStorage.removeItem('mp_note_' + dk);
      }
      // Mark as manually edited
      input.dataset.manual = '1';
      input.style.color = '#333';
      input.style.fontStyle = 'normal';
      // Debounced sync write
      clearTimeout(input._syncTimeout);
      input._syncTimeout = setTimeout(() => syncUpsertEvent(dk), 1000);
    });
  });

  // Bind tooltips
  mainContent.querySelectorAll('[data-tip]').forEach(cell => {
    cell.addEventListener('mouseenter', e => {
      if (!cell.dataset.tip) return;
      tooltipEl.textContent = cell.dataset.tip;
      tooltipEl.style.display = 'block';
      positionTooltip(e);
    });
    cell.addEventListener('mousemove', positionTooltip);
    cell.addEventListener('mouseleave', () => {
      tooltipEl.style.display = 'none';
    });
  });
}

function positionTooltip(e) {
  tooltipEl.style.left = (e.clientX + 12) + 'px';
  tooltipEl.style.top = (e.clientY + 12) + 'px';
}

function applyReservedLevel(cell, lvl) {
  cell.dataset.level = lvl;
  cell.dataset.tip = lvl > 0 ? RESERVED_LABELS[lvl] : '';
  const opacity = RESERVED_LEVELS[lvl];
  let block = cell.querySelector('.reserved-block');
  if (opacity > 0) {
    if (!block) {
      block = document.createElement('div');
      block.className = 'reserved-block';
      cell.appendChild(block);
    }
    block.style.background = `rgba(0,0,0,${opacity})`;
  } else if (block) {
    block.remove();
  }
}

function renderLegend(calendars) {
  legendEl.innerHTML = '';

  // Reserved legend
  const levels = [
    { opacity: 0.25, label: 'Planning (25%)' },
    { opacity: 0.50, label: 'Considering (50%)' },
    { opacity: 0.75, label: 'Confident (75%)' },
    { opacity: 1.0,  label: 'Reserved (100%)' },
  ];
  levels.forEach(l => {
    const item = document.createElement('span');
    item.className = 'legend-item';
    const swatch = document.createElement('span');
    swatch.className = 'legend-color';
    swatch.style.background = `rgba(0,0,0,${l.opacity})`;
    item.appendChild(swatch);
    item.appendChild(document.createTextNode(' ' + l.label));
    legendEl.appendChild(item);
  });

  // Holiday legend
  const hItem = document.createElement('span');
  hItem.className = 'legend-item';
  const hSwatch = document.createElement('span');
  hSwatch.className = 'legend-color';
  hSwatch.style.background = '#e53935';
  hItem.appendChild(hSwatch);
  hItem.appendChild(document.createTextNode(' Holiday'));
  legendEl.appendChild(hItem);

  // Calendar legends
  calendars.forEach(cal => {
    const item = document.createElement('span');
    item.className = 'legend-item';
    const swatch = document.createElement('span');
    swatch.className = 'legend-color';
    swatch.style.background = cal.backgroundColor || '#4285f4';
    item.appendChild(swatch);
    item.appendChild(document.createTextNode(' ' + (cal.summaryOverride || cal.summary || cal.id)));
    legendEl.appendChild(item);
  });
}

// ── API Helper ──
async function apiFetch(url) {
  try {
    const resp = await fetch(url, {
      headers: { Authorization: 'Bearer ' + accessToken },
    });
    if (resp.status === 401) {
      tokenClient.requestAccessToken();
      return null;
    }
    if (!resp.ok) {
      console.error('API error', resp.status, await resp.text());
      return null;
    }
    return await resp.json();
  } catch (err) {
    console.error('Fetch error:', err);
    return null;
  }
}

// ── Utilities ──
function dateKey(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function formatTime(isoStr) {
  const d = new Date(isoStr);
  let h = d.getHours();
  const m = String(d.getMinutes()).padStart(2, '0');
  const ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  return `${h}:${m} ${ampm}`;
}

function esc(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML.replace(/"/g, '&quot;');
}

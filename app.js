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
let currentView = 'grid';    // 'grid', 'gantt', 'addtrip', 'summary'

function reloadView() {
  if (currentView === 'gantt') loadGantt();
  else if (currentView === 'addtrip') renderAddTripForm();
  else if (currentView === 'summary') renderSummaryList();
  else loadMonth();
}

// Frequent cities grouped by country — used for trip form location picker
const LOCATIONS = [
  { country: 'Vietnam', cities: ['Ho Chi Minh City', 'Hanoi', 'Da Nang', 'Hai Phong', 'Nha Trang', 'Vung Tau'] },
  { country: 'Thailand', cities: ['Bangkok', 'Chiang Mai', 'Pattaya', 'Phuket', 'Hua Hin'] },
  { country: 'China', cities: ['Shanghai', 'Xiamen', 'Shenzhen', 'Guangzhou', 'Beijing', 'Wuxi', 'Jiaxing'] },
  { country: 'Cambodia', cities: ['Phnom Penh', 'Siem Reap'] },
  { country: 'Singapore', cities: ['Singapore'] },
  { country: 'Mexico', cities: ['Mexico City', 'Guadalajara', 'Cancun'] },
  { country: 'USA', cities: ['San Francisco', 'Los Angeles', 'New York', 'Phoenix'] },
  { country: 'Japan', cities: ['Tokyo', 'Osaka'] },
  { country: 'South Korea', cities: ['Seoul', 'Busan'] },
  { country: 'Hong Kong', cities: ['Hong Kong'] },
  { country: 'Taiwan', cities: ['Taipei'] },
  { country: 'Malaysia', cities: ['Kuala Lumpur'] },
  { country: 'Philippines', cities: ['Manila', 'Cebu'] },
  { country: 'Indonesia', cities: ['Jakarta', 'Bali'] },
];

const TRIP_TYPES = ['My Trip', 'Friend Visit', 'Client Visit', 'Event', 'Conference'];

// Trip Ideas calendar ID — found at runtime
let tripIdeasCalId = null;

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
const RESERVED_COLORS = ['transparent', '#4CAF50', '#FFC107', '#FF9800', '#F44336']; // green→yellow→orange→red
const RESERVED_TEXT_COLORS = ['#333', '#fff', '#333', '#fff', '#fff'];

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
  resyncBtn.style.display = 'inline-block';
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
  resyncBtn.style.display = 'none';
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
  await ensureTripIdeasCalendar(allItems);

  renderCalendarCheckboxes();
  renderColumnCheckboxes();
  reloadView();
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

// Find or create the "Trip Ideas" calendar
async function ensureTripIdeasCalendar(calendarItems) {
  const cached = localStorage.getItem('mp_tripIdeasCalId');
  const found = calendarItems.find(c =>
    c.id === cached ||
    (c.summary && c.summary.toLowerCase().includes('trip idea'))
  );
  if (found) {
    tripIdeasCalId = found.id;
    localStorage.setItem('mp_tripIdeasCalId', tripIdeasCalId);
    return;
  }
  // Create it
  try {
    const resp = await fetch('https://www.googleapis.com/calendar/v3/calendars', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        summary: 'Trip Ideas',
        description: 'Trip planning calendar managed by Month Planner app.',
      }),
    });
    if (resp.ok) {
      const cal = await resp.json();
      tripIdeasCalId = cal.id;
      localStorage.setItem('mp_tripIdeasCalId', tripIdeasCalId);
    }
  } catch(e) { console.error('Failed to create Trip Ideas calendar:', e); }
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
      reloadView();
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
      reloadView();
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
  reloadView();
}

// ── Month Navigation ──
prevBtn.onclick = () => { shiftMonth(-1); };
nextBtn.onclick = () => { shiftMonth(1); };
todayBtn.onclick = () => {
  const now = new Date();
  currentYear = now.getFullYear();
  currentMonth = now.getMonth();
  updateTitle();
  if (accessToken) reloadView();
};

// Re-sync all reserved days to Google Calendar with current format
const resyncBtn = document.getElementById('resync-btn');
resyncBtn.onclick = async () => {
  if (!syncReady || !syncCalId) return;
  resyncBtn.disabled = true;
  resyncBtn.textContent = 'Clearing old...';

  // Step 1: Delete ALL sync events for the month (catches orphaned old-format events)
  const timeMin = new Date(currentYear, currentMonth, 1).toISOString();
  const timeMax = new Date(currentYear, currentMonth + 1, 0, 23, 59, 59).toISOString();
  const params = new URLSearchParams({
    timeMin, timeMax, singleEvents: 'true', maxResults: '250',
    privateExtendedProperty: 'mpApp=monthplanner',
  });
  const url = `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(syncCalId)}/events?${params}`;
  const data = await apiFetch(url);
  if (data && data.items) {
    for (let di = 0; di < data.items.length; di++) {
      try {
        await fetch(
          `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(syncCalId)}/events/${encodeURIComponent(data.items[di].id)}`,
          { method: 'DELETE', headers: { 'Authorization': 'Bearer ' + accessToken } }
        );
      } catch(e) {}
    }
  }
  // Clear local tracking
  syncEventIds = {};

  // Step 2: Create fresh events for all days with data
  resyncBtn.textContent = 'Syncing...';
  const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(currentYear, currentMonth, day);
    const dk = dateKey(date);
    const reserved = parseInt(localStorage.getItem('mp_reserved_' + dk) || '0', 10);
    const note = localStorage.getItem('mp_note_' + dk) || '';
    const tripInfo = tripIdeaDates[dk];
    if (reserved > 0 || note || (tripInfo && tripInfo.trips.length > 0)) {
      await syncUpsertEvent(dk);
    }
  }
  resyncBtn.disabled = false;
  resyncBtn.textContent = 'Re-sync All';
  updateSyncStatus('synced');
};

// View toggle
const viewToggle = document.getElementById('view-toggle');
viewToggle.onclick = () => {
  const views = ['grid', 'gantt'];
  const idx = views.indexOf(currentView);
  currentView = views[(idx + 1) % views.length];
  viewToggle.textContent = currentView === 'grid' ? 'Gantt View' : 'Month View';
  if (accessToken) reloadView();
};

// Add Trip button
document.getElementById('add-trip-btn').onclick = () => {
  currentView = 'addtrip';
  if (accessToken) renderAddTripForm();
};

// Summary button
document.getElementById('summary-btn').onclick = () => {
  currentView = 'summary';
  if (accessToken) renderSummaryList();
};

// Legend toggle
const legendToggle = document.getElementById('legend-toggle');
let legendVisible = true;
legendToggle.onclick = () => {
  legendVisible = !legendVisible;
  legendEl.style.display = legendVisible ? '' : 'none';
  legendToggle.textContent = legendVisible ? 'Key' : 'Show Key';
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
  if (accessToken) reloadView();
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

    if (!syncEventIds[dk]) syncEventIds[dk] = [];
    syncEventIds[dk].push(ev.id);

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
// Creates one Google Calendar event per trip on a given day.
// If no trips, creates a single event with the reserved label or manual note.
async function syncUpsertEvent(dk) {
  if (!syncReady || !syncCalId) return;
  updateSyncStatus('saving');

  const note = localStorage.getItem('mp_note_' + dk) || '';
  const tripInfo = tripIdeaDates[dk];

  if (!note && (!tripInfo || tripInfo.trips.length === 0)) {
    await syncDeleteAllEvents(dk);
    return;
  }

  // Build list of summaries — one per trip, each with its own % from title
  const summaries = [];
  if (tripInfo && tripInfo.trips.length > 0 && !note) {
    // Sort: first day of event first, then last day, then middle days
    tripInfo.trips.slice().sort(function(a, b) {
      const aFirst = a.dayNum === 1 ? 0 : (a.dayNum === a.totalDays ? 1 : 2);
      const bFirst = b.dayNum === 1 ? 0 : (b.dayNum === b.totalDays ? 1 : 2);
      return aFirst - bFirst;
    }).forEach(function(t) {
      const dayLabel = '(' + t.dayNum + '/' + t.totalDays + ') ';
      const tripMatch = t.title.match(/^trip ideas?\s*-\s*/i);
      const desc = tripMatch ? t.title.substring(tripMatch[0].length) : t.title;
      const cleanDesc = desc.replace(/\d+\s*%\s*/, '').trim();
      const tripPct = Math.round(RESERVED_LEVELS[t.level] * 100) + '%';
      summaries.push(tripPct + ' ' + dayLabel + cleanDesc);
    });
  } else if (note) {
    const cleanNote = note.replace(/^trip ideas?\s*-\s*/i, '').replace(/\d+\s*%\s*/, '').trim();
    summaries.push(cleanNote.substring(0, 60));
  }

  // Delete old events for this day first, then create fresh ones
  await syncDeleteAllEvents(dk);

  const newIds = [];
  for (let si = 0; si < summaries.length; si++) {
    const eventBody = {
      summary: summaries[si],
      description: note || '',
      start: { date: dk },
      end: { date: nextDay(dk) },
      extendedProperties: {
        private: {
          mpApp: 'monthplanner',
          mpReserved: '0',
          mpNote: note,
        },
      },
    };

    try {
      const resp = await fetch(
        `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(syncCalId)}/events`,
        {
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
          body: JSON.stringify(eventBody),
        }
      );
      if (resp.ok) {
        const ev = await resp.json();
        newIds.push(ev.id);
      } else {
        console.error('Sync write failed:', resp.status, await resp.text());
      }
    } catch (err) {
      console.error('Sync write error:', err);
    }
  }

  syncEventIds[dk] = newIds;
  updateSyncStatus(newIds.length > 0 ? 'synced' : 'error');
}

async function syncDeleteAllEvents(dk) {
  if (!syncReady || !syncCalId) return;
  const existingIds = syncEventIds[dk];
  if (!existingIds || existingIds.length === 0) { updateSyncStatus('synced'); return; }

  for (let i = 0; i < existingIds.length; i++) {
    try {
      await fetch(
        `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(syncCalId)}/events/${encodeURIComponent(existingIds[i])}`,
        {
          method: 'DELETE',
          headers: { 'Authorization': 'Bearer ' + accessToken },
        }
      );
    } catch (err) {
      console.error('Sync delete error:', err);
    }
  }
  delete syncEventIds[dk];
  updateSyncStatus('synced');
}

// Keep old name as alias for places that call syncDeleteEvent
async function syncDeleteEvent(dk) { return syncDeleteAllEvents(dk); }

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

  // Scan all events for "trip idea" → extract day counter, per-trip per-day % levels
  // tripIdeaDates[dk] = { level: max level of day, trips: [{title, tripKey, dayNum, totalDays, level}] }
  // Each trip's % is stored per-day in localStorage: mp_trip_pct_{dk}_{tripKey}
  // tripKey = cleaned event summary for consistent storage keys
  tripIdeaDates = {};
  Object.entries(calEvents).forEach(([calId, dateMap]) => {
    if (calId === syncCalId) return;
    Object.entries(dateMap).forEach(([dk, events]) => {
      events.forEach(ev => {
        if (!ev.summary) return;
        if (!ev.summary.toLowerCase().includes('trip idea')) return;

        // Parse default % from title (used as initial value if no per-day override)
        const pctMatch = ev.summary.match(/(\d+)\s*%/);
        let defaultLevel = 1; // 25% Planning
        if (pctMatch) {
          const pct = parseInt(pctMatch[1], 10);
          if (pct >= 100) defaultLevel = 4;
          else if (pct >= 75) defaultLevel = 3;
          else if (pct >= 50) defaultLevel = 2;
          else defaultLevel = 1;
        }

        // Build a stable key from the event summary (strip % and "Trip Ideas -" prefix)
        const tripKey = ev.summary.replace(/^trip ideas?\s*-\s*/i, '').replace(/\d+\s*%\s*/, '').trim().substring(0, 40);

        const evStart = ev.start.date ? ev.start.date : ev.start.dateTime.split('T')[0];
        const evEnd = ev.end.date ? ev.end.date : ev.end.dateTime.split('T')[0];
        const startDate = new Date(evStart + 'T00:00:00');
        const endDate = new Date(evEnd + 'T00:00:00');
        const totalDays = Math.max(1, Math.round((endDate - startDate) / 86400000));
        const currentDate = new Date(dk + 'T00:00:00');
        const dayNum = Math.round((currentDate - startDate) / 86400000) + 1;

        // Use per-day stored level if available, otherwise use default from title
        const storedLevel = localStorage.getItem('mp_trip_pct_' + dk + '_' + tripKey);
        const level = storedLevel !== null ? parseInt(storedLevel, 10) : defaultLevel;

        if (!tripIdeaDates[dk]) tripIdeaDates[dk] = { level: 0, trips: [] };
        const isDup = tripIdeaDates[dk].trips.some(t => t.tripKey === tripKey);
        if (!isDup) {
          tripIdeaDates[dk].trips.push({ title: ev.summary, tripKey, dayNum, totalDays, level });
          if (level > tripIdeaDates[dk].level) tripIdeaDates[dk].level = level;
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

// Build the trip note text for a given day's tripInfo
function buildTripNoteText(tripInfo) {
  if (!tripInfo || tripInfo.trips.length === 0) return '';
  const sorted = tripInfo.trips.slice().sort(function(a, b) {
    const aFirst = a.dayNum === 1 ? 0 : (a.dayNum === a.totalDays ? 1 : 2);
    const bFirst = b.dayNum === 1 ? 0 : (b.dayNum === b.totalDays ? 1 : 2);
    return aFirst - bFirst;
  });
  const stars = sorted.length > 1 ? '*'.repeat(sorted.length) + ' ' : '';
  const parts = sorted.map(function(t) {
    const dayLabel = '(' + t.dayNum + '/' + t.totalDays + ')';
    const tripPct = Math.round(RESERVED_LEVELS[t.level] * 100) + '%';
    const tripMatch = t.title.match(/^trip ideas?\s*-\s*/i);
    const desc = tripMatch ? t.title.substring(tripMatch[0].length) : t.title;
    const cleanDesc = desc.replace(/\d+\s*%\s*/, '').trim();
    const prefix = tripMatch ? tripMatch[0] : '';
    return prefix + tripPct + ' ' + dayLabel + ' ' + cleanDesc;
  });
  return stars + parts.join(' | ');
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
    const isToday = (currentYear === new Date().getFullYear() && currentMonth === new Date().getMonth() && day === new Date().getDate());

    html += `<tr class="${isWeekend ? 'weekend' : ''}${isToday ? ' today-row' : ''}">`;
    html += `<td class="date-cell">${dayNames[dow]} ${String(day).padStart(2, '\u00A0')}</td>`;

    orderedCols.forEach(col => {
      if (col.key === 'col_reserved') {
        // Auto-calculated from highest trip level on this day (read-only)
        let level = tripInfo ? tripInfo.level : 0;
        const opacity = RESERVED_LEVELS[level];
        const label = RESERVED_LABELS[level];
        const pct = Math.round(opacity * 100);
        const bgColor = RESERVED_COLORS[level];
        const txtColor = RESERVED_TEXT_COLORS[level];
        html += `<td class="fixed-col reserved-cell" data-date="${dk}" data-level="${level}" data-tip="${label} (${pct}%)">`;
        if (level > 0) html += `<div class="reserved-block" style="background:${bgColor};color:${txtColor};display:flex;align-items:center;justify-content:center;font-size:8px;font-weight:600;overflow:hidden;">${pct}%</div>`;
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
        if (!manualNote && tripInfo && tripInfo.trips.length > 0) {
          noteVal = buildTripNoteText(tripInfo);
        }
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
  // Reserved column — click to set % per trip per day
  mainContent.querySelectorAll('.reserved-cell').forEach(cell => {
    cell.style.cursor = 'pointer';
    cell.addEventListener('click', (e) => {
      e.stopPropagation();
      const old = document.getElementById('reserved-picker');
      if (old) old.remove();

      const dk = cell.dataset.date;
      const tripInfo = tripIdeaDates[dk];
      if (!tripInfo || tripInfo.trips.length === 0) return;

      const picker = document.createElement('div');
      picker.id = 'reserved-picker';
      picker.className = 'reserved-picker';

      const header = document.createElement('div');
      header.style.cssText = 'font-weight:600;margin-bottom:6px;font-size:12px;';
      header.textContent = dk + ' — Set % per trip';
      picker.appendChild(header);

      tripInfo.trips.forEach(t => {
        const row = document.createElement('div');
        row.style.cssText = 'margin-bottom:8px;padding:4px;border:1px solid #eee;border-radius:4px;';

        // Trip name (cleaned)
        const label = document.createElement('div');
        label.style.cssText = 'font-size:11px;color:#555;margin-bottom:4px;';
        const cleanName = t.title.replace(/^trip ideas?\s*-\s*/i, '').replace(/\d+\s*%\s*/, '').trim();
        label.textContent = '(' + t.dayNum + '/' + t.totalDays + ') ' + cleanName.substring(0, 50);
        row.appendChild(label);

        // Level buttons
        const btnRow = document.createElement('div');
        btnRow.style.cssText = 'display:flex;gap:3px;';
        [
          { lvl: 1, lbl: '25%' },
          { lvl: 2, lbl: '50%' },
          { lvl: 3, lbl: '75%' },
          { lvl: 4, lbl: '100%' },
        ].forEach(opt => {
          const btn = document.createElement('button');
          btn.style.cssText = 'padding:2px 8px;border:1px solid #ccc;border-radius:3px;font-size:10px;cursor:pointer;background:' + RESERVED_COLORS[opt.lvl] + ';color:' + RESERVED_TEXT_COLORS[opt.lvl] + ';';
          if (t.level === opt.lvl) btn.style.outline = '2px solid #333';
          btn.textContent = opt.lbl;
          btn.onclick = (ev) => {
            ev.stopPropagation();
            // Save per-trip per-day level
            localStorage.setItem('mp_trip_pct_' + dk + '_' + t.tripKey, opt.lvl);
            t.level = opt.lvl;
            // Recalculate day max level
            let maxLvl = 0;
            tripInfo.trips.forEach(tr => { if (tr.level > maxLvl) maxLvl = tr.level; });
            tripInfo.level = maxLvl;
            // Update reserved cell display
            applyReservedLevel(cell, maxLvl);
            // Update note display
            const noteInput = mainContent.querySelector('.notes-cell input[data-date="' + dk + '"]');
            if (noteInput && noteInput.dataset.manual === '0') {
              noteInput.value = buildTripNoteText(tripInfo);
              noteInput.style.color = '#999';
              noteInput.style.fontStyle = 'italic';
            }
            // Sync to Google Calendar
            syncUpsertEvent(dk);
            picker.remove();
          };
          btnRow.appendChild(btn);
        });
        row.appendChild(btnRow);
        picker.appendChild(row);
      });

      const rect = cell.getBoundingClientRect();
      picker.style.top = (rect.bottom + window.scrollY + 2) + 'px';
      picker.style.left = (rect.left + window.scrollX) + 'px';
      document.body.appendChild(picker);

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
  const pct = Math.round(RESERVED_LEVELS[lvl] * 100);
  cell.dataset.tip = lvl > 0 ? RESERVED_LABELS[lvl] + ' (' + pct + '%)' : '';
  let block = cell.querySelector('.reserved-block');
  if (lvl > 0) {
    if (!block) {
      block = document.createElement('div');
      block.className = 'reserved-block';
      cell.appendChild(block);
    }
    block.style.background = RESERVED_COLORS[lvl];
    block.style.color = RESERVED_TEXT_COLORS[lvl];
    block.style.display = 'flex';
    block.style.alignItems = 'center';
    block.style.justifyContent = 'center';
    block.style.fontSize = '8px';
    block.style.fontWeight = '600';
    block.textContent = pct + '%';
  } else if (block) {
    block.remove();
  }
}

function renderLegend(calendars) {
  legendEl.innerHTML = '';

  // Reserved legend — stacked vertically, 100% on top, aligned left
  const reservedSection = document.createElement('div');
  reservedSection.style.cssText = 'display:flex;flex-direction:column;gap:1px;margin-right:12px;';
  const levels = [
    { color: RESERVED_COLORS[4], textColor: RESERVED_TEXT_COLORS[4], label: 'Reserved (100%)' },
    { color: RESERVED_COLORS[3], textColor: RESERVED_TEXT_COLORS[3], label: 'Confident (75%)' },
    { color: RESERVED_COLORS[2], textColor: RESERVED_TEXT_COLORS[2], label: 'Considering (50%)' },
    { color: RESERVED_COLORS[1], textColor: RESERVED_TEXT_COLORS[1], label: 'Planning (25%)' },
  ];
  levels.forEach(l => {
    const item = document.createElement('div');
    item.style.cssText = 'display:flex;align-items:center;gap:4px;';
    const swatch = document.createElement('span');
    swatch.className = 'legend-color';
    swatch.style.background = l.color;
    item.appendChild(swatch);
    item.appendChild(document.createTextNode(l.label));
    reservedSection.appendChild(item);
  });
  legendEl.appendChild(reservedSection);

  // Holiday legends — stacked vertically like reserved
  const holidaySection = document.createElement('div');
  holidaySection.style.cssText = 'display:flex;flex-direction:column;gap:1px;margin-right:12px;';
  COUNTRY_COLUMNS.forEach(cc => {
    if (hiddenColumns.has('col_' + cc.name)) return;
    const item = document.createElement('div');
    item.style.cssText = 'display:flex;align-items:center;gap:4px;';
    const swatch = document.createElement('span');
    swatch.className = 'legend-color';
    swatch.style.background = '#e53935';
    item.appendChild(swatch);
    item.appendChild(document.createTextNode(cc.name + ' Holiday'));
    holidaySection.appendChild(item);
  });
  if (holidaySection.children.length > 0) legendEl.appendChild(holidaySection);

  // Calendar legends — only show calendars that are visible (not hidden in columns)
  calendars.forEach(cal => {
    if (hiddenColumns.has('col_cal_' + cal.id)) return;
    const item = document.createElement('span');
    item.className = 'legend-item';
    const swatch = document.createElement('span');
    swatch.className = 'legend-color';
    swatch.style.background = cal.backgroundColor || '#4285f4';
    item.appendChild(swatch);
    item.appendChild(document.createTextNode(' ' + (cal.summaryOverride || cal.summary || cal.id)));
    legendEl.appendChild(item);
  });

  // Respect legend visibility toggle
  legendEl.style.display = legendVisible ? '' : 'none';
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

// ── Gantt View ──
// Shows trip events as horizontal bars across a 3-month timeline.
// X-axis: days (1 month back, current month, 2 months ahead)
// Y-axis: each unique trip event
// Bars colored by per-trip per-day % level. Click to change %.

async function loadGantt() {
  mainContent.innerHTML = '<div class="loading">Loading Gantt view...</div>';

  const today = new Date();
  const startMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const endMonth = new Date(today.getFullYear(), today.getMonth() + 3, 0);

  const timeMin = startMonth.toISOString();
  const timeMax = new Date(endMonth.getFullYear(), endMonth.getMonth(), endMonth.getDate(), 23, 59, 59).toISOString();

  const selectedCals = allCalendars
    .filter(c => selectedCalendarIds.includes(c.id))
    .filter(c => !HOLIDAY_CAL_IDS.has(c.id) && !isHolidayCalendar(c));

  const allTripEvents = [];
  await Promise.all(selectedCals.map(async cal => {
    if (cal.id === syncCalId) return;
    const events = await fetchEvents(cal.id, timeMin, timeMax);
    events.forEach(ev => {
      if (!ev.summary) return;
      if (!ev.summary.toLowerCase().includes('trip idea')) return;
      allTripEvents.push(ev);
    });
  }));

  // Fetch holidays for Gantt
  const ganttHolidays = {};
  await Promise.all(COUNTRY_COLUMNS.map(async cc => {
    if (!cc.calId) return;
    const events = await fetchEvents(cc.calId, timeMin, timeMax);
    ganttHolidays[cc.name] = {};
    events.forEach(ev => {
      const s = new Date((ev.start.date || ev.start.dateTime.split('T')[0]) + 'T00:00:00');
      const e = new Date((ev.end.date || ev.end.dateTime.split('T')[0]) + 'T00:00:00');
      const d = new Date(s);
      while (d < e) {
        const dk = dateKey(d);
        if (!ganttHolidays[cc.name][dk]) ganttHolidays[cc.name][dk] = [];
        ganttHolidays[cc.name][dk].push(ev.summary);
        d.setDate(d.getDate() + 1);
      }
    });
  }));

  renderGantt(allTripEvents, startMonth, endMonth, today, ganttHolidays);
}

function renderGantt(tripEvents, startDate, endDate, today, ganttHolidays) {
  // Build list of all days in range
  const days = [];
  const d = new Date(startDate);
  while (d <= endDate) {
    days.push(new Date(d));
    d.setDate(d.getDate() + 1);
  }
  const totalDays = days.length;

  // Build unique trips: group by cleaned title (dedup multi-day same event)
  const tripMap = {}; // tripKey → { title, tripKey, startDk, endDk, days: Set of dk }
  tripEvents.forEach(ev => {
    const tripMatch = ev.summary.match(/^trip ideas?\s*-\s*/i);
    const desc = tripMatch ? ev.summary.substring(tripMatch[0].length) : ev.summary;
    const tripKey = desc.replace(/\d+\s*%\s*/, '').trim().substring(0, 40);

    const evStart = ev.start.date || ev.start.dateTime.split('T')[0];
    const evEnd = ev.end.date || ev.end.dateTime.split('T')[0];
    const s = new Date(evStart + 'T00:00:00');
    const e = new Date(evEnd + 'T00:00:00');

    if (!tripMap[tripKey]) {
      tripMap[tripKey] = { title: ev.summary, tripKey, startDk: evStart, endDk: evEnd, days: new Set() };
    }

    // Expand all days of this event
    const cur = new Date(s);
    while (cur < e) {
      tripMap[tripKey].days.add(dateKey(cur));
      cur.setDate(cur.getDate() + 1);
    }
    // Update start/end bounds
    if (evStart < tripMap[tripKey].startDk) tripMap[tripKey].startDk = evStart;
    if (evEnd > tripMap[tripKey].endDk) tripMap[tripKey].endDk = evEnd;
  });

  const trips = Object.values(tripMap).sort((a, b) => a.startDk.localeCompare(b.startDk));
  const todayDk = dateKey(today);
  const dayNames = ['S','M','T','W','T','F','S'];

  // Hidden trips set — persisted to localStorage
  const hiddenTrips = new Set(JSON.parse(localStorage.getItem('mp_hiddenTrips') || '[]'));

  // Filter trips to only those with at least 1 day in the viewed range
  const daySet = new Set(days.map(d => dateKey(d)));
  const visibleTrips = trips.filter(trip => {
    for (const dk of trip.days) { if (daySet.has(dk)) return true; }
    return false;
  });

  let html = '<div class="gantt-container"><div class="gantt-chart">';

  // Month labels row
  html += '<div class="gantt-month-row">';
  html += '<div class="gantt-label-col" style="background:#f0f0f0"></div>';
  let mi = 0;
  while (mi < days.length) {
    const m = days[mi].getMonth();
    let span = 0;
    while (mi + span < days.length && days[mi + span].getMonth() === m) span++;
    const mName = days[mi].toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    html += '<div class="gantt-month-cell" style="flex:' + span + ';min-width:' + (span * 22) + 'px">' + mName + '</div>';
    mi += span;
  }
  html += '<div class="gantt-end-col"></div>';
  html += '</div>';

  // Day headers row
  html += '<div class="gantt-header">';
  html += '<div class="gantt-label-col">Trip</div>';
  html += '<div class="gantt-timeline">';
  days.forEach(day => {
    const dow = day.getDay();
    const isToday = dateKey(day) === todayDk;
    const isWeekend = dow === 0 || dow === 6;
    const isMonthStart = day.getDate() === 1;
    let cls = 'gantt-day-header';
    if (isWeekend) cls += ' weekend';
    if (isToday) cls += ' today';
    if (isMonthStart) cls += ' month-start';
    html += '<div class="' + cls + '"><div>' + day.getDate() + '</div><div>' + dayNames[dow] + '</div></div>';
  });
  html += '</div></div>';

  // Reserved row (highest % per day across all trips)
  if (!hiddenColumns.has('col_reserved')) {
  html += '<div class="gantt-row" style="min-height:28px;">';
  html += '<div class="gantt-row-label section-label">Reserved</div>';
  html += '<div class="gantt-row-timeline">';
  days.forEach(day => {
    const dk = dateKey(day);
    const dow = day.getDay();
    // Find max level across all trips for this day
    let maxLevel = 0;
    trips.forEach(trip => {
      if (trip.days.has(dk)) {
        const stored = localStorage.getItem('mp_trip_pct_' + dk + '_' + trip.tripKey);
        const pctMatch = trip.title.match(/(\d+)\s*%/);
        let defLvl = 1;
        if (pctMatch) { const p = parseInt(pctMatch[1],10); if (p>=100) defLvl=4; else if (p>=75) defLvl=3; else if (p>=50) defLvl=2; }
        const lvl = stored !== null ? parseInt(stored, 10) : defLvl;
        if (lvl > maxLevel) maxLevel = lvl;
      }
    });
    if (maxLevel > 0) {
      const pctVal = Math.round(RESERVED_LEVELS[maxLevel] * 100);
      html += '<div class="gantt-cell' + (dow===0||dow===6 ? ' weekend' : '') + '"><div class="gantt-bar" style="background:' + RESERVED_COLORS[maxLevel] + ';color:' + RESERVED_TEXT_COLORS[maxLevel] + '">' + pctVal + '%</div></div>';
    } else {
      html += '<div class="gantt-cell' + (dow===0||dow===6 ? ' weekend' : '') + '"></div>';
    }
  });
  html += '</div></div>';
  } // end if !hiddenColumns reserved

  // Holiday rows per country (half height, respect hidden columns)
  if (ganttHolidays) {
    COUNTRY_COLUMNS.forEach(cc => {
      if (hiddenColumns.has('col_' + cc.name)) return;
      const holidays = ganttHolidays[cc.name] || {};
      html += '<div class="gantt-row holiday-row">';
      html += '<div class="gantt-row-label section-label">' + esc(cc.name) + '</div>';
      html += '<div class="gantt-row-timeline">';
      days.forEach((day, dayIdx) => {
        const dk = dateKey(day);
        const dow = day.getDay();
        const names = holidays[dk];
        if (names && names.length > 0) {
          // Check prev/next day for any holiday in this country to merge into one visual bar
          const prevDk = dayIdx > 0 ? dateKey(days[dayIdx - 1]) : null;
          const nextDk = dayIdx < days.length - 1 ? dateKey(days[dayIdx + 1]) : null;
          const hasPrev = prevDk && holidays[prevDk] && holidays[prevDk].length > 0;
          const hasNext = nextDk && holidays[nextDk] && holidays[nextDk].length > 0;
          let dotCls = 'gantt-holiday-dot';
          if (!hasPrev && !hasNext) dotCls += ' gantt-holiday-single';
          else if (!hasPrev) dotCls += ' gantt-holiday-first';
          else if (!hasNext) dotCls += ' gantt-holiday-last';
          html += '<div class="gantt-cell' + (dow===0||dow===6 ? ' weekend' : '') + '" data-country="' + esc(cc.name) + '" data-dk="' + dk + '" title="' + esc(names.join(', ')) + '"><div class="' + dotCls + '"></div></div>';
        } else {
          html += '<div class="gantt-cell' + (dow===0||dow===6 ? ' weekend' : '') + '"></div>';
        }
      });
      html += '</div></div>';
    });
  }

  // Dark separator between holidays and trips — built as a row with matching cells
  const hiddenCount = visibleTrips.filter(t => hiddenTrips.has(t.tripKey)).length;
  const sepHeight = hiddenCount > 0 ? '12px' : '4px';
  html += '<div class="gantt-row" style="min-height:' + sepHeight + ';background:#555;">';
  html += '<div class="gantt-row-label" style="min-height:' + sepHeight + ';padding:0;background:#555;font-size:8px;color:#ccc;display:flex;align-items:center;">';
  if (hiddenCount > 0) {
    html += '<button id="gantt-show-all" style="background:none;border:none;color:#ccc;font-size:8px;cursor:pointer;padding:0 4px;line-height:1;">Show ' + hiddenCount + ' hidden</button>';
  }
  html += '</div>';
  html += '<div class="gantt-row-timeline" style="background:#555;">';
  days.forEach(function() { html += '<div class="gantt-cell" style="background:#555;min-height:' + sepHeight + ';border:none;"></div>'; });
  html += '</div>';
  html += '<div class="gantt-end-col" style="background:#555;"></div>';
  html += '</div>';

  // Trip rows — only visible trips that have days in the current range, sorted by first day
  visibleTrips.forEach(trip => {
    if (hiddenTrips.has(trip.tripKey)) return;
    const cleanName = trip.title.replace(/^trip ideas?\s*-\s*/i, '').replace(/\d+\s*%\s*/, '').trim();
    html += '<div class="gantt-row">';
    html += '<div class="gantt-row-label trip-label" title="' + esc(cleanName) + '"><button class="gantt-hide-btn" data-hide-trip="' + esc(trip.tripKey) + '">✕</button><span class="trip-name-text">' + esc(cleanName) + '</span></div>';
    html += '<div class="gantt-row-timeline">';

    // Pre-compute levels for border gradient between days
    const pctMatch = trip.title.match(/(\d+)\s*%/);
    let defaultLevel = 1;
    if (pctMatch) {
      const p = parseInt(pctMatch[1], 10);
      if (p >= 100) defaultLevel = 4; else if (p >= 75) defaultLevel = 3; else if (p >= 50) defaultLevel = 2;
    }

    days.forEach((day, dayIdx) => {
      const dk = dateKey(day);
      const isInTrip = trip.days.has(dk);
      const dow = day.getDay();
      const isWeekend = dow === 0 || dow === 6;

      if (isInTrip) {
        const storedLevel = localStorage.getItem('mp_trip_pct_' + dk + '_' + trip.tripKey);
        const level = storedLevel !== null ? parseInt(storedLevel, 10) : defaultLevel;
        const pctVal = Math.round(RESERVED_LEVELS[level] * 100);
        const bgColor = RESERVED_COLORS[level];
        const txtColor = RESERVED_TEXT_COLORS[level];

        // Check if prev/next days are in trip for border styling
        const prevDk = dayIdx > 0 ? dateKey(days[dayIdx - 1]) : null;
        const nextDk = dayIdx < days.length - 1 ? dateKey(days[dayIdx + 1]) : null;
        const isFirst = !prevDk || !trip.days.has(prevDk);
        const isLast = !nextDk || !trip.days.has(nextDk);

        // Border: top and bottom in trip's color, left only on first, right only on last
        let borderStyle = 'border-top:2px solid ' + bgColor + ';border-bottom:2px solid ' + bgColor + ';';
        if (isFirst) borderStyle += 'border-left:2px solid ' + bgColor + ';border-radius:4px 0 0 4px;';
        if (isLast) borderStyle += 'border-right:2px solid ' + bgColor + ';border-radius:0 4px 4px 0;';
        if (isFirst && isLast) borderStyle += 'border-radius:4px;';

        html += '<div class="gantt-cell' + (isWeekend ? ' weekend' : '') + '" data-dk="' + dk + '" data-tripkey="' + esc(trip.tripKey) + '" data-level="' + level + '" data-title="' + esc(cleanName) + '" style="' + borderStyle + '">';
        html += '<div class="gantt-bar" style="background:' + bgColor + ';color:' + txtColor + '">' + pctVal + '%</div>';
        html += '</div>';
      } else {
        html += '<div class="gantt-cell' + (isWeekend ? ' weekend' : '') + '"></div>';
      }
    });

    html += '</div></div>';
  });

  // Today line — positioned via JS after render
  const todayIdx = days.findIndex(d => dateKey(d) === todayDk);
  if (todayIdx >= 0) {
    html += '<div class="gantt-today-line" data-today-idx="' + todayIdx + '" data-total="' + totalDays + '"></div>';
  }

  html += '</div></div>';
  mainContent.innerHTML = html;

  // Add + Month button positioned in the end column, spanning trip rows
  const chart = mainContent.querySelector('.gantt-chart');
  const tripRowEls = mainContent.querySelectorAll('.gantt-row .trip-label');
  const endCols = mainContent.querySelectorAll('.gantt-end-col');
  if (tripRowEls.length > 0 && endCols.length > 0) {
    const firstTripRow = tripRowEls[0].closest('.gantt-row');
    const lastTripRow = tripRowEls[tripRowEls.length - 1].closest('.gantt-row');
    const topOffset = firstTripRow.offsetTop;
    const height = lastTripRow.offsetTop + lastTripRow.offsetHeight - topOffset;
    const lastEndCol = endCols[0];
    const leftOffset = lastEndCol.offsetLeft;

    const btnWrap = document.createElement('div');
    btnWrap.style.cssText = 'position:absolute;top:' + topOffset + 'px;left:' + leftOffset + 'px;width:28px;height:' + height + 'px;z-index:20;';
    const btn = document.createElement('button');
    btn.id = 'gantt-load-more';
    btn.style.cssText = 'writing-mode:vertical-rl;text-orientation:mixed;background:#dc3545;color:#fff;border:none;border-radius:4px;padding:8px 4px;font-size:9px;font-weight:600;cursor:pointer;letter-spacing:1px;width:100%;height:100%;';
    btn.textContent = '+ Month';
    btnWrap.appendChild(btn);
    chart.appendChild(btnWrap);
  }

  // Load more click — extend endDate by 1 month and re-render
  const loadMoreBtn = document.getElementById('gantt-load-more');
  if (loadMoreBtn) {
    loadMoreBtn.onclick = async () => {
      const newEnd = new Date(endDate.getFullYear(), endDate.getMonth() + 2, 0);
      loadMoreBtn.textContent = 'Loading...';
      // Fetch additional month's events
      const extraTimeMin = new Date(endDate.getFullYear(), endDate.getMonth() + 1, 1).toISOString();
      const extraTimeMax = new Date(newEnd.getFullYear(), newEnd.getMonth(), newEnd.getDate(), 23, 59, 59).toISOString();
      const selectedCals = allCalendars
        .filter(c => selectedCalendarIds.includes(c.id))
        .filter(c => !HOLIDAY_CAL_IDS.has(c.id) && !isHolidayCalendar(c));
      const extraTrips = [];
      await Promise.all(selectedCals.map(async cal => {
        if (cal.id === syncCalId) return;
        const events = await fetchEvents(cal.id, extraTimeMin, extraTimeMax);
        events.forEach(ev => {
          if (ev.summary && ev.summary.toLowerCase().includes('trip idea')) extraTrips.push(ev);
        });
      }));
      const extraHolidays = {};
      await Promise.all(COUNTRY_COLUMNS.map(async cc => {
        if (!cc.calId) return;
        const events = await fetchEvents(cc.calId, extraTimeMin, extraTimeMax);
        extraHolidays[cc.name] = {};
        events.forEach(ev => {
          const s2 = new Date((ev.start.date || ev.start.dateTime.split('T')[0]) + 'T00:00:00');
          const e2 = new Date((ev.end.date || ev.end.dateTime.split('T')[0]) + 'T00:00:00');
          const d2 = new Date(s2);
          while (d2 < e2) {
            const dk2 = dateKey(d2);
            if (!extraHolidays[cc.name][dk2]) extraHolidays[cc.name][dk2] = [];
            extraHolidays[cc.name][dk2].push(ev.summary);
            d2.setDate(d2.getDate() + 1);
          }
        });
      }));
      // Merge holidays
      COUNTRY_COLUMNS.forEach(cc => {
        if (extraHolidays[cc.name]) {
          Object.entries(extraHolidays[cc.name]).forEach(([dk2, names]) => {
            if (!ganttHolidays[cc.name]) ganttHolidays[cc.name] = {};
            ganttHolidays[cc.name][dk2] = names;
          });
        }
      });
      const allTrips = tripEvents.concat(extraTrips);
      renderGantt(allTrips, startDate, newEnd, today, ganttHolidays);
    };
  }

  // Position today line using actual DOM measurements
  const todayLine = mainContent.querySelector('.gantt-today-line');
  if (todayLine) {
    const firstTimeline = mainContent.querySelector('.gantt-row-timeline');
    if (firstTimeline) {
      const cells = firstTimeline.children;
      const idx = parseInt(todayLine.dataset.todayIdx, 10);
      if (idx >= 0 && idx < cells.length) {
        const cell = cells[idx];
        const chartRect = mainContent.querySelector('.gantt-chart').getBoundingClientRect();
        const cellRect = cell.getBoundingClientRect();
        todayLine.style.left = (cellRect.left - chartRect.left + cellRect.width / 2) + 'px';
      }
    }
  }

  // Bind click handlers for % editing on trip bars
  mainContent.querySelectorAll('.gantt-cell[data-tripkey]').forEach(cell => {
    cell.addEventListener('click', () => {
      const dk = cell.dataset.dk;
      const tripKey = cell.dataset.tripkey;
      const currentLevel = parseInt(cell.dataset.level, 10);
      // Cycle: 1→2→3→4→1
      const nextLevel = currentLevel >= 4 ? 1 : currentLevel + 1;

      localStorage.setItem('mp_trip_pct_' + dk + '_' + tripKey, nextLevel);

      // Update bar display
      const bar = cell.querySelector('.gantt-bar');
      const pctVal = Math.round(RESERVED_LEVELS[nextLevel] * 100);
      bar.style.background = RESERVED_COLORS[nextLevel];
      bar.style.color = RESERVED_TEXT_COLORS[nextLevel];
      bar.textContent = pctVal + '%';
      cell.dataset.level = nextLevel;

      // Sync to Google Calendar
      syncUpsertEvent(dk);
    });
  });

  // Bind click handlers for holiday info popup
  // Holiday click — show grouped list of consecutive holiday dates
  mainContent.querySelectorAll('.holiday-row .gantt-cell[data-country]').forEach(cell => {
    cell.addEventListener('click', (e) => {
      const country = cell.dataset.country;
      const dk = cell.dataset.dk;
      if (!country || !dk || !ganttHolidays[country]) return;
      const holidays = ganttHolidays[country];

      // Walk backwards and forwards to find all consecutive holiday dates
      const group = [];
      // Walk back
      let d = new Date(dk + 'T00:00:00');
      while (true) {
        d.setDate(d.getDate() - 1);
        const prevDk = dateKey(d);
        if (holidays[prevDk] && holidays[prevDk].length > 0) {
          group.unshift({ dk: prevDk, names: holidays[prevDk] });
        } else break;
      }
      // Add clicked day
      group.push({ dk: dk, names: holidays[dk] });
      // Walk forward
      d = new Date(dk + 'T00:00:00');
      while (true) {
        d.setDate(d.getDate() + 1);
        const nextDk = dateKey(d);
        if (holidays[nextDk] && holidays[nextDk].length > 0) {
          group.push({ dk: nextDk, names: holidays[nextDk] });
        } else break;
      }

      const old = document.getElementById('holiday-popup');
      if (old) old.remove();
      const popup = document.createElement('div');
      popup.id = 'holiday-popup';
      popup.style.cssText = 'position:fixed;background:#fff;border:1px solid #ccc;border-radius:6px;padding:10px 14px;font-size:12px;box-shadow:0 4px 12px rgba(0,0,0,0.15);z-index:200;max-width:350px;';
      let popupHtml = '<div style="font-weight:600;margin-bottom:6px;">' + esc(country) + ' Holidays</div>';
      group.forEach(g => {
        const date = new Date(g.dk + 'T00:00:00');
        const dateStr = date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
        g.names.forEach(name => {
          popupHtml += '<div style="margin:2px 0;">' + dateStr + ' — ' + esc(name) + '</div>';
        });
      });
      popup.innerHTML = popupHtml;
      popup.style.left = (e.clientX + 10) + 'px';
      popup.style.top = (e.clientY + 10) + 'px';
      document.body.appendChild(popup);
      setTimeout(() => {
        document.addEventListener('click', function closePopup() {
          popup.remove();
          document.removeEventListener('click', closePopup);
        }, { once: true });
      }, 0);
    });
  });

  // Click trip label to show full text as overlay popup
  mainContent.querySelectorAll('.trip-label').forEach(label => {
    label.addEventListener('click', (e) => {
      if (e.target.classList.contains('gantt-hide-btn')) return;
      const old = document.querySelector('.gantt-trip-popup');
      if (old) old.remove();
      const popup = document.createElement('div');
      popup.className = 'gantt-trip-popup';
      popup.textContent = label.getAttribute('title');
      popup.style.left = (e.clientX + 10) + 'px';
      popup.style.top = (e.clientY - 10) + 'px';
      document.body.appendChild(popup);
      setTimeout(() => {
        document.addEventListener('click', function close() {
          popup.remove();
          document.removeEventListener('click', close);
        }, { once: true });
      }, 0);
    });
  });

  // Hide trip buttons
  mainContent.querySelectorAll('.gantt-hide-btn[data-hide-trip]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const key = btn.dataset.hideTrip;
      hiddenTrips.add(key);
      localStorage.setItem('mp_hiddenTrips', JSON.stringify([...hiddenTrips]));
      renderGantt(tripEvents, startDate, endDate, today, ganttHolidays);
    });
  });

  // Show all hidden trips
  const showAllBtn = document.getElementById('gantt-show-all');
  if (showAllBtn) {
    showAllBtn.addEventListener('click', () => {
      hiddenTrips.clear();
      localStorage.setItem('mp_hiddenTrips', '[]');
      renderGantt(tripEvents, startDate, endDate, today, ganttHolidays);
    });
  }
}

// ── Add Trip Form ──
function renderAddTripForm() {
  let locationOptions = '';
  LOCATIONS.forEach(loc => {
    locationOptions += '<optgroup label="' + esc(loc.country) + '">';
    loc.cities.forEach(city => {
      locationOptions += '<option value="' + esc(city + ', ' + loc.country) + '">' + esc(city) + '</option>';
    });
    locationOptions += '</optgroup>';
  });

  mainContent.innerHTML = `
    <div class="trip-form">
      <h2>Add Trip Idea</h2>
      <label>Trip Name</label>
      <input type="text" id="trip-name" placeholder="e.g. Visit Jared in Bangkok">
      <div class="form-row">
        <div>
          <label>Start Date</label>
          <input type="date" id="trip-start">
        </div>
        <div>
          <label>End Date</label>
          <input type="date" id="trip-end">
        </div>
      </div>
      <label>Location</label>
      <select id="trip-location">
        <option value="">Select city...</option>
        ${locationOptions}
        <option value="__custom__">Other (type below)</option>
      </select>
      <input type="text" id="trip-location-custom" placeholder="City, Country" style="display:none;margin-top:4px;">
      <label>Who (optional)</label>
      <input type="text" id="trip-who" placeholder="e.g. Jared Stevens">
      <label>Type</label>
      <div class="type-btns" id="trip-type-btns">
        ${TRIP_TYPES.map((t, i) => '<button class="type-btn' + (i === 0 ? ' active' : '') + '" data-type="' + esc(t) + '">' + esc(t) + '</button>').join('')}
      </div>
      <label>Likelihood</label>
      <div class="type-btns" id="trip-pct-btns">
        <button class="type-btn active" data-pct="25" style="background:#4CAF50;color:#fff;border-color:#4CAF50;">25%</button>
        <button class="type-btn" data-pct="50" style="background:#FFC107;color:#333;border-color:#FFC107;">50%</button>
        <button class="type-btn" data-pct="75" style="background:#FF9800;color:#fff;border-color:#FF9800;">75%</button>
        <button class="type-btn" data-pct="100" style="background:#F44336;color:#fff;border-color:#F44336;">100%</button>
      </div>
      <label>Notes (optional)</label>
      <textarea id="trip-notes" placeholder="Any details..."></textarea>
      <button class="submit-btn" id="trip-submit">Add Trip</button>
      <div class="form-status" id="trip-status"></div>
    </div>
  `;

  // Location custom toggle
  const locSelect = document.getElementById('trip-location');
  const locCustom = document.getElementById('trip-location-custom');
  locSelect.onchange = () => {
    locCustom.style.display = locSelect.value === '__custom__' ? 'block' : 'none';
  };

  // Type buttons
  let selectedType = TRIP_TYPES[0];
  document.getElementById('trip-type-btns').addEventListener('click', (e) => {
    if (!e.target.dataset.type) return;
    selectedType = e.target.dataset.type;
    document.querySelectorAll('#trip-type-btns .type-btn').forEach(b => b.classList.remove('active'));
    e.target.classList.add('active');
  });

  // Pct buttons
  let selectedPct = 25;
  document.getElementById('trip-pct-btns').addEventListener('click', (e) => {
    if (!e.target.dataset.pct) return;
    selectedPct = parseInt(e.target.dataset.pct, 10);
    document.querySelectorAll('#trip-pct-btns .type-btn').forEach(b => b.classList.remove('active'));
    e.target.classList.add('active');
  });

  // Submit
  document.getElementById('trip-submit').onclick = async () => {
    const name = document.getElementById('trip-name').value.trim();
    const start = document.getElementById('trip-start').value;
    const end = document.getElementById('trip-end').value;
    const location = locSelect.value === '__custom__' ? locCustom.value.trim() : locSelect.value;
    const who = document.getElementById('trip-who').value.trim();
    const notes = document.getElementById('trip-notes').value.trim();
    const status = document.getElementById('trip-status');

    if (!name || !start || !end) {
      status.textContent = 'Please fill in name, start and end dates.';
      status.style.color = '#dc3545';
      return;
    }

    if (!tripIdeasCalId) {
      status.textContent = 'Trip Ideas calendar not found. Please reload.';
      status.style.color = '#dc3545';
      return;
    }

    // Build event summary: "Trip Ideas - 25% Visit Jared in Bangkok"
    const summary = 'Trip Ideas - ' + selectedPct + '% ' + name;

    // Build description with structured data
    let desc = '';
    if (who) desc += 'Who: ' + who + '\n';
    desc += 'Type: ' + selectedType + '\n';
    if (location) desc += 'Location: ' + location + '\n';
    desc += 'Likelihood: ' + selectedPct + '%\n';
    if (notes) desc += 'Notes: ' + notes;

    // End date for all-day event is exclusive (next day)
    const endDate = new Date(end + 'T00:00:00');
    endDate.setDate(endDate.getDate() + 1);
    const endStr = endDate.toISOString().split('T')[0];

    const submitBtn = document.getElementById('trip-submit');
    submitBtn.disabled = true;
    submitBtn.textContent = 'Creating...';

    try {
      const resp = await fetch(
        'https://www.googleapis.com/calendar/v3/calendars/' + encodeURIComponent(tripIdeasCalId) + '/events',
        {
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
          body: JSON.stringify({
            summary: summary,
            description: desc,
            location: location,
            start: { date: start },
            end: { date: endStr },
          }),
        }
      );

      if (resp.ok) {
        status.textContent = 'Trip added! Switching to Gantt view...';
        status.style.color = '#28a745';
        // Clear cache so it reloads
        eventsCache = {};
        setTimeout(() => {
          currentView = 'gantt';
          viewToggle.textContent = 'Month View';
          loadGantt();
        }, 1000);
      } else {
        const err = await resp.text();
        status.textContent = 'Error: ' + err;
        status.style.color = '#dc3545';
        submitBtn.disabled = false;
        submitBtn.textContent = 'Add Trip';
      }
    } catch(e) {
      status.textContent = 'Error: ' + e.message;
      status.style.color = '#dc3545';
      submitBtn.disabled = false;
      submitBtn.textContent = 'Add Trip';
    }
  };

  // Set default dates
  const today = new Date();
  document.getElementById('trip-start').value = dateKey(today);
  const nextWeek = new Date(today);
  nextWeek.setDate(nextWeek.getDate() + 3);
  document.getElementById('trip-end').value = dateKey(nextWeek);
}

// ── Summary List ──
async function renderSummaryList() {
  mainContent.innerHTML = '<div class="loading">Loading trips...</div>';

  const today = new Date();
  const startMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const endMonth = new Date(today.getFullYear(), today.getMonth() + 6, 0);
  const timeMin = startMonth.toISOString();
  const timeMax = new Date(endMonth.getFullYear(), endMonth.getMonth(), endMonth.getDate(), 23, 59, 59).toISOString();

  // Fetch from all calendars (not just trip ideas) to catch existing events
  const selectedCals = allCalendars
    .filter(c => selectedCalendarIds.includes(c.id))
    .filter(c => !HOLIDAY_CAL_IDS.has(c.id) && !isHolidayCalendar(c));

  const trips = [];
  await Promise.all(selectedCals.map(async cal => {
    if (cal.id === syncCalId) return;
    const events = await fetchEvents(cal.id, timeMin, timeMax);
    events.forEach(ev => {
      if (!ev.summary || !ev.summary.toLowerCase().includes('trip idea')) return;
      // Dedup by summary + start date
      const startDk = ev.start.date || ev.start.dateTime.split('T')[0];
      if (trips.some(t => t.summary === ev.summary && t.startDk === startDk)) return;

      const endDk = ev.end.date || ev.end.dateTime.split('T')[0];
      const s = new Date(startDk + 'T00:00:00');
      const e = new Date(endDk + 'T00:00:00');
      const days = Math.max(1, Math.round((e - s) / 86400000));

      // Parse structured data from description — handle both real newlines and literal \n
      const desc = (ev.description || '').replace(/\\n/g, '\n');
      const descLines = desc.split('\n');
      let who = '', type = '', location = ev.location || '', notes = '';
      descLines.forEach(line => {
        const l = line.trim();
        if (l.match(/^Who:\s*/i)) who = l.replace(/^Who:\s*/i, '');
        else if (l.match(/^Type:\s*/i)) type = l.replace(/^Type:\s*/i, '');
        else if (l.match(/^Location:\s*/i) && !location) location = l.replace(/^Location:\s*/i, '');
        else if (l.match(/^Notes:\s*/i)) notes = l.replace(/^Notes:\s*/i, '');
      });

      // Parse % from title
      const pctMatch = ev.summary.match(/(\d+)\s*%/);
      const pct = pctMatch ? parseInt(pctMatch[1], 10) : 25;
      const level = pct >= 100 ? 4 : pct >= 75 ? 3 : pct >= 50 ? 2 : 1;

      const cleanName = ev.summary.replace(/^trip ideas?\s*-\s*/i, '').replace(/\d+\s*%\s*/, '').trim();

      trips.push({ summary: ev.summary, cleanName, startDk, endDk, days, who, type, location, notes, pct, level, eventId: ev.id, calId: ev.organizer ? ev.organizer.email : '' });
    });
  }));

  // Sort by start date
  trips.sort((a, b) => a.startDk.localeCompare(b.startDk));

  // Get unique types for filter buttons
  const types = ['All'];
  trips.forEach(t => { if (t.type && types.indexOf(t.type) === -1) types.push(t.type); });

  let html = '<div class="summary-container">';
  html += '<div class="summary-header">';
  html += '<h2>Trips (' + trips.length + ')</h2>';
  html += '<div class="summary-filters">';
  types.forEach(t => {
    html += '<button class="summary-filter-btn' + (t === 'All' ? ' active' : '') + '" data-filter="' + esc(t) + '">' + esc(t) + '</button>';
  });
  html += '</div></div>';

  html += '<div class="summary-cards" id="summary-cards">';
  trips.forEach((trip, idx) => {
    const bgColor = RESERVED_COLORS[trip.level];
    const startDate = new Date(trip.startDk + 'T00:00:00');
    const endDate = new Date(trip.endDk + 'T00:00:00');
    endDate.setDate(endDate.getDate() - 1);
    const startStr = startDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
    const endStr = endDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
    const dateRange = trip.days === 1 ? startStr : startStr + ' – ' + endStr;

    html += '<div class="summary-card" data-type="' + esc(trip.type) + '" data-idx="' + idx + '" style="cursor:pointer;">';
    html += '<div class="summary-card-pct" style="background:' + bgColor + '">' + trip.pct + '%</div>';
    html += '<div class="summary-card-body">';
    html += '<div class="summary-card-title">' + esc(trip.cleanName) + '</div>';
    html += '<div class="summary-card-meta">';
    html += '<span>' + dateRange + ' (' + trip.days + 'd)</span>';
    if (trip.location) html += '<span>' + esc(trip.location) + '</span>';
    if (trip.who) html += '<span>' + esc(trip.who) + '</span>';
    if (trip.type) html += '<span class="summary-card-type">' + esc(trip.type) + '</span>';
    html += '</div>';
    if (trip.notes) html += '<div style="font-size:11px;color:#999;margin-top:3px;">' + esc(trip.notes) + '</div>';
    html += '</div></div>';
  });
  html += '</div></div>';

  mainContent.innerHTML = html;

  // Click card to edit
  mainContent.querySelectorAll('.summary-card').forEach(card => {
    card.addEventListener('click', (e) => {
      if (e.target.closest('.summary-filter-btn')) return;
      const idx = parseInt(card.dataset.idx, 10);
      const trip = trips[idx];
      if (!trip) return;

      const old = document.getElementById('trip-edit-popup');
      if (old) old.remove();

      const popup = document.createElement('div');
      popup.id = 'trip-edit-popup';
      popup.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:#fff;border:1px solid #ccc;border-radius:12px;padding:16px 20px;font-size:13px;box-shadow:0 8px 24px rgba(0,0,0,0.2);z-index:200;width:360px;max-width:90vw;';

      let phtml = '<div style="font-weight:700;font-size:15px;margin-bottom:10px;">' + esc(trip.cleanName) + '</div>';
      phtml += '<div style="color:#666;font-size:12px;margin-bottom:6px;">' + trip.startDk + ' to ' + trip.endDk + ' (' + trip.days + ' days)</div>';
      if (trip.location) phtml += '<div style="color:#666;font-size:12px;">Location: ' + esc(trip.location) + '</div>';
      if (trip.who) phtml += '<div style="color:#666;font-size:12px;">Who: ' + esc(trip.who) + '</div>';
      if (trip.type) phtml += '<div style="color:#666;font-size:12px;">Type: ' + esc(trip.type) + '</div>';
      if (trip.notes) phtml += '<div style="color:#666;font-size:12px;margin-top:4px;">Notes: ' + esc(trip.notes) + '</div>';

      // % buttons
      phtml += '<div style="margin-top:12px;font-size:11px;font-weight:600;color:#888;">CHANGE LIKELIHOOD</div>';
      phtml += '<div style="display:flex;gap:6px;margin-top:6px;">';
      [25, 50, 75, 100].forEach(p => {
        const lvl = p >= 100 ? 4 : p >= 75 ? 3 : p >= 50 ? 2 : 1;
        const active = trip.pct === p ? 'outline:2px solid #333;' : '';
        phtml += '<button class="edit-pct-btn" data-pct="' + p + '" style="flex:1;padding:8px;border:none;border-radius:6px;cursor:pointer;font-weight:700;font-size:13px;background:' + RESERVED_COLORS[lvl] + ';color:' + RESERVED_TEXT_COLORS[lvl] + ';' + active + '">' + p + '%</button>';
      });
      phtml += '</div>';

      phtml += '<div style="display:flex;gap:8px;margin-top:14px;">';
      phtml += '<button id="edit-close" style="flex:1;padding:8px;border:1px solid #ccc;border-radius:6px;cursor:pointer;background:#fff;font-size:12px;">Close</button>';
      phtml += '<button id="edit-gantt" style="flex:1;padding:8px;border:none;border-radius:6px;cursor:pointer;background:#0a66c2;color:#fff;font-size:12px;">View in Gantt</button>';
      phtml += '</div>';

      popup.innerHTML = phtml;
      document.body.appendChild(popup);

      // Backdrop
      const backdrop = document.createElement('div');
      backdrop.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.3);z-index:199;';
      document.body.appendChild(backdrop);

      const closePopup = () => { popup.remove(); backdrop.remove(); };
      backdrop.onclick = closePopup;
      popup.querySelector('#edit-close').onclick = closePopup;

      popup.querySelector('#edit-gantt').onclick = () => {
        closePopup();
        currentView = 'gantt';
        viewToggle.textContent = 'Month View';
        loadGantt();
      };

      // % change buttons — update Google Calendar event title
      popup.querySelectorAll('.edit-pct-btn').forEach(btn => {
        btn.addEventListener('click', async () => {
          const newPct = parseInt(btn.dataset.pct, 10);
          const newSummary = 'Trip Ideas - ' + newPct + '% ' + trip.cleanName;
          const calId = trip.calId || tripIdeasCalId;

          btn.textContent = '...';
          try {
            const resp = await fetch(
              'https://www.googleapis.com/calendar/v3/calendars/' + encodeURIComponent(calId) + '/events/' + encodeURIComponent(trip.eventId),
              {
                method: 'PATCH',
                headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
                body: JSON.stringify({ summary: newSummary }),
              }
            );
            if (resp.ok) {
              eventsCache = {};
              closePopup();
              renderSummaryList();
            } else {
              btn.textContent = 'Error';
            }
          } catch(err) {
            btn.textContent = 'Error';
          }
        });
      });
    });
  });

  // Filter buttons
  mainContent.querySelectorAll('.summary-filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const filter = btn.dataset.filter;
      mainContent.querySelectorAll('.summary-filter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      mainContent.querySelectorAll('.summary-card').forEach(card => {
        if (filter === 'All' || card.dataset.type === filter) {
          card.style.display = '';
        } else {
          card.style.display = 'none';
        }
      });
    });
  });
}

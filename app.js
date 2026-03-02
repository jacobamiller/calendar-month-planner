// ── State ──
let accessToken = null;
let tokenClient = null;
let allCalendars = [];       // from API
let selectedCalendarIds = []; // user-chosen subset
let currentYear, currentMonth; // 0-indexed month
let eventsCache = {};         // "YYYY-MM" → { calId → { "YYYY-MM-DD" → [events] } }

// ── DOM refs ──
const monthTitle = document.getElementById('month-title');
const mainContent = document.getElementById('main-content');
const signInPrompt = document.getElementById('sign-in-prompt');
const authBtn = document.getElementById('auth-btn');
const prevBtn = document.getElementById('prev-btn');
const nextBtn = document.getElementById('next-btn');
const todayBtn = document.getElementById('today-btn');
const settingsToggle = document.getElementById('settings-toggle');
const settingsPanel = document.getElementById('settings-panel');
const calendarCheckboxes = document.getElementById('calendar-checkboxes');
const legendEl = document.getElementById('legend');
const tooltipEl = document.getElementById('tooltip');

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

  // Wait for GIS library to load
  window.addEventListener('load', () => {
    if (typeof google !== 'undefined' && google.accounts) {
      initAuth();
    } else {
      // GIS script may still be loading
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
      tokenClient.requestAccessToken();
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
  allCalendars = (resp.items || [])
    .filter(c => c.accessRole !== 'freeBusyReader')
    .sort((a, b) => (a.summary || '').localeCompare(b.summary || ''));

  // Restore saved selection or default to all
  const saved = localStorage.getItem('mp_selectedCals');
  if (saved) {
    const ids = JSON.parse(saved);
    selectedCalendarIds = ids.filter(id => allCalendars.some(c => c.id === id));
  } else {
    selectedCalendarIds = allCalendars.map(c => c.id);
  }

  renderCalendarCheckboxes();
  loadMonth();
}

function renderCalendarCheckboxes() {
  calendarCheckboxes.innerHTML = '';
  allCalendars.forEach(cal => {
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

settingsToggle.onclick = () => {
  settingsPanel.classList.toggle('open');
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

// ── Load Events for Month ──
async function loadMonth() {
  const key = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;
  const selectedCals = allCalendars.filter(c => selectedCalendarIds.includes(c.id));

  if (selectedCals.length === 0) {
    renderGrid(selectedCals, {});
    return;
  }

  // Check cache
  if (eventsCache[key]) {
    renderGrid(selectedCals, eventsCache[key]);
    renderLegend(selectedCals);
    return;
  }

  mainContent.innerHTML = '<div class="loading">Loading events...</div>';

  const timeMin = new Date(currentYear, currentMonth, 1).toISOString();
  const timeMax = new Date(currentYear, currentMonth + 1, 0, 23, 59, 59).toISOString();

  const calEvents = {};
  await Promise.all(selectedCals.map(async cal => {
    const events = await fetchEvents(cal.id, timeMin, timeMax);
    calEvents[cal.id] = indexEventsByDate(events, currentYear, currentMonth);
  }));

  eventsCache[key] = calEvents;
  renderGrid(selectedCals, calEvents);
  renderLegend(selectedCals);
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
  const map = {}; // "YYYY-MM-DD" → [event]
  const monthStart = new Date(year, month, 1);
  const monthEnd = new Date(year, month + 1, 0);

  events.forEach(ev => {
    const start = ev.start.dateTime ? new Date(ev.start.dateTime) : new Date(ev.start.date + 'T00:00:00');
    const end = ev.end.dateTime ? new Date(ev.end.dateTime) : new Date(ev.end.date + 'T00:00:00');

    // For all-day events, the end date is exclusive, so subtract 1 day
    const endAdjusted = ev.end.date && !ev.end.dateTime
      ? new Date(end.getTime() - 86400000)
      : end;

    // Iterate each day the event spans within this month
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
function renderGrid(calendars, calEvents) {
  const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
  const dayNames = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

  let html = '<div class="grid-container"><table class="month-grid">';

  // Header row
  html += '<thead><tr class="header-row">';
  html += '<th class="date-header">Date</th>';
  calendars.forEach(cal => {
    const name = cal.summaryOverride || cal.summary || cal.id;
    html += `<th><span class="angled-header" title="${esc(name)}">${esc(name)}</span></th>`;
  });
  html += '<th class="notes-header">Notes</th>';
  html += '</tr></thead>';

  // Body rows
  html += '<tbody>';
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(currentYear, currentMonth, day);
    const dow = date.getDay();
    const isWeekend = dow === 0 || dow === 6;
    const dk = dateKey(date);

    html += `<tr class="${isWeekend ? 'weekend' : ''}">`;
    html += `<td class="date-cell">${dayNames[dow]} ${String(day).padStart(2, '\u00A0')}</td>`;

    calendars.forEach(cal => {
      const events = (calEvents[cal.id] && calEvents[cal.id][dk]) || [];
      const color = cal.backgroundColor || '#4285f4';

      if (events.length > 0) {
        const titles = events.map(e => {
          let t = e.summary || '(No title)';
          if (e.start.dateTime) {
            t += ' \u2022 ' + formatTime(e.start.dateTime);
          }
          return t;
        }).join('\n');
        html += `<td class="event-cell" data-tip="${esc(titles)}"><div class="event-block" style="background:${color}"></div></td>`;
      } else {
        html += '<td class="event-cell"></td>';
      }
    });

    // Notes
    const noteVal = localStorage.getItem('mp_note_' + dk) || '';
    html += `<td class="notes-cell"><input type="text" value="${esc(noteVal)}" data-date="${dk}" /></td>`;
    html += '</tr>';
  }
  html += '</tbody></table></div>';

  mainContent.innerHTML = html;

  // Bind note saving
  mainContent.querySelectorAll('.notes-cell input').forEach(input => {
    input.addEventListener('input', () => {
      const dk = input.dataset.date;
      if (input.value) {
        localStorage.setItem('mp_note_' + dk, input.value);
      } else {
        localStorage.removeItem('mp_note_' + dk);
      }
    });
  });

  // Bind tooltips
  mainContent.querySelectorAll('[data-tip]').forEach(cell => {
    cell.addEventListener('mouseenter', e => {
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

function renderLegend(calendars) {
  legendEl.innerHTML = '';
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
      // Token expired — request a new one
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

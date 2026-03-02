// Google Cloud OAuth 2.0 Client ID
// To set up:
// 1. Go to https://console.cloud.google.com/
// 2. Create a new project (or select existing)
// 3. Enable the "Google Calendar API" under APIs & Services > Library
// 4. Go to APIs & Services > Credentials
// 5. Create an OAuth 2.0 Client ID (type: Web application)
// 6. Add http://localhost:8000 to "Authorized JavaScript origins"
// 7. Copy the Client ID below
const CLIENT_ID = 'YOUR_CLIENT_ID_HERE';
const SCOPES = 'https://www.googleapis.com/auth/calendar.readonly';

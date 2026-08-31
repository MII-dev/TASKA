// ==========================================
// ====== Microsoft Teams calendar sync =====
// ==========================================
//
// Microsoft Graph API (the normal way to read Teams/Outlook data) needs a
// full OAuth app registration in the organization's Azure AD tenant, which
// isn't available here. Power Automate sidesteps that entirely: it already
// has an approved Office 365 Outlook connector under the user's own
// account, and can push events to this script on a schedule via a plain
// HTTP POST — no app registration, no OAuth flow, just a shared secret.

const TEAMS_WEBHOOK_SECRET_PROP = 'TEAMS_WEBHOOK_SECRET';
const TEAMS_EVENTS_JSON_PROP = 'TEAMS_EVENTS_JSON';
const TEAMS_EVENTS_SYNCED_AT_PROP = 'TEAMS_EVENTS_SYNCED_AT';

/**
 * Receives a bulk sync of upcoming Teams/Outlook events from the Power
 * Automate flow. Replaces the whole stored set on every call — a
 * recurrence-triggered "send everything upcoming" flow is simpler and more
 * self-healing than trying to reconcile individual create/update/delete
 * webhooks, and it means a missed run just gets caught by the next one.
 * Expects JSON body: {secret, events: [{id, title, start, end, isAllDay, location, organizer, joinUrl}]}
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const expectedSecret = getTeamsWebhookSecret();

    if (!expectedSecret || body.secret !== expectedSecret) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const events = Array.isArray(body.events) ? body.events : [];
    PropertiesService.getScriptProperties().setProperty(TEAMS_EVENTS_JSON_PROP, JSON.stringify(events));
    PropertiesService.getScriptProperties().setProperty(TEAMS_EVENTS_SYNCED_AT_PROP, new Date().toISOString());

    return ContentService.createTextOutput(JSON.stringify({ ok: true, count: events.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Returns (creating on first use) the shared secret Power Automate must send
 * on every sync — this is a public, anonymously-accessible web app, so
 * without this check anyone who found the URL could inject fake events.
 */
function getTeamsWebhookSecret() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty(TEAMS_WEBHOOK_SECRET_PROP);
  if (!secret) {
    secret = Utilities.getUuid();
    props.setProperty(TEAMS_WEBHOOK_SECRET_PROP, secret);
  }
  return secret;
}

/**
 * Info the client needs to configure the Power Automate flow: this
 * deployment's own URL (the POST target) plus the shared secret, and a
 * quick status readout so the settings UI can show whether a sync has
 * actually landed yet.
 */
function getTeamsWebhookInfo() {
  return {
    url: ScriptApp.getService().getUrl(),
    secret: getTeamsWebhookSecret(),
    lastSyncedAt: PropertiesService.getScriptProperties().getProperty(TEAMS_EVENTS_SYNCED_AT_PROP) || null,
    eventCount: getTeamsEvents().length
  };
}

/**
 * Returns the most recently synced Teams events, normalized into the same
 * shape fetchCalendarEvents() produces for Google Calendar — the client's
 * existing rendering, meeting-link detection and "AI Бриф" button all key
 * off that shape already, so merging Teams events into the same widget
 * needs no changes there.
 */
function getTeamsEvents() {
  const raw = PropertiesService.getScriptProperties().getProperty(TEAMS_EVENTS_JSON_PROP);
  if (!raw) return [];

  let events;
  try {
    events = JSON.parse(raw);
  } catch (e) {
    return [];
  }
  if (!Array.isArray(events)) return [];

  return events
    .filter(function (ev) { return ev && ev.start && ev.end; })
    .map(function (ev) {
      return {
        title: ev.title || 'Подія Teams',
        start: ev.start,
        end: ev.end,
        location: ev.location || '',
        description: ev.organizer ? ('Організатор: ' + ev.organizer) : '',
        htmlDescription: '',
        hangoutLink: ev.joinUrl || '',
        conferenceName: ev.joinUrl ? 'Microsoft Teams' : '',
        isAllDay: !!ev.isAllDay,
        color: '#6264A7',
        calendarName: 'Microsoft Teams',
        eventId: 'teams_' + (ev.id || (ev.title + '_' + ev.start))
      };
    });
}

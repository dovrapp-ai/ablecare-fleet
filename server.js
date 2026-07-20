const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = Number(process.env.PORT || 5173);
const ROOT = __dirname;
const SQUARE_VERSION = '2026-01-22';
const RINGCENTRAL_TOKEN_URL = '/restapi/oauth/token';
const MS_GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
const FIREBASE_DATABASE_URL = process.env.FIREBASE_DATABASE_URL
  || 'https://ablecare-fleet-default-rtdb.firebaseio.com';
let ringCentralToken = null;
let microsoftGraphToken = null;
let reminderSchedulerTimer = null;
let reminderSchedulerLastRunKey = '';

loadEnv(path.join(ROOT, '.env'));

function loadEnv(filePath) {
  if (!fs.existsSync(filePath)) return;
  const lines = fs.readFileSync(filePath, 'utf8').split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#') || !trimmed.includes('=')) continue;
    const index = trimmed.indexOf('=');
    const key = trimmed.slice(0, index).trim();
    const value = trimmed.slice(index + 1).trim();
    if (key && process.env[key] === undefined) process.env[key] = value;
  }
}

function squareBaseUrl() {
  return process.env.SQUARE_ENVIRONMENT === 'production'
    ? 'https://connect.squareup.com'
    : 'https://connect.squareupsandbox.com';
}

function sendJson(res, status, body) {
  const json = JSON.stringify(body);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
  });
  res.end(json);
}

function readRequestBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (chunk) => chunks.push(chunk));
    req.on('end', () => {
      const raw = Buffer.concat(chunks).toString('utf8');
      if (!raw) {
        resolve({});
        return;
      }
      try {
        resolve(JSON.parse(raw));
      } catch (error) {
        error.status = 400;
        error.message = 'Request body must be valid JSON.';
        reject(error);
      }
    });
    req.on('error', reject);
  });
}

function parseDateRange(searchParams) {
  const from = searchParams.get('from');
  const to = searchParams.get('to');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from || '') || !/^\d{4}-\d{2}-\d{2}$/.test(to || '')) {
    throw new Error('Use from and to dates in YYYY-MM-DD format.');
  }

  const begin = new Date(`${from}T00:00:00`);
  const end = new Date(`${to}T00:00:00`);
  end.setDate(end.getDate() + 1);

  if (Number.isNaN(begin.getTime()) || Number.isNaN(end.getTime()) || begin >= end) {
    throw new Error('Invalid date range.');
  }

  return {
    from,
    to,
    beginTime: begin.toISOString(),
    endTime: end.toISOString(),
  };
}

async function squareGet(endpoint, params = {}) {
  const token = process.env.SQUARE_ACCESS_TOKEN;
  if (!token || token.includes('paste_your') || token.includes('replace_with')) {
    const err = new Error('Square access token is missing from .env.');
    err.status = 400;
    throw err;
  }

  const url = new URL(endpoint, squareBaseUrl());
  for (const [key, value] of Object.entries(params)) {
    if (value !== undefined && value !== null && value !== '') url.searchParams.set(key, value);
  }

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Square-Version': SQUARE_VERSION,
      Accept: 'application/json',
      'Content-Type': 'application/json',
    },
  });

  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = body.errors?.map((e) => e.detail || e.code).filter(Boolean).join('; ')
      || `Square API returned HTTP ${response.status}`;
    const err = new Error(message);
    err.status = response.status;
    err.squareErrors = body.errors;
    throw err;
  }

  return body;
}

function ringCentralBaseUrl() {
  return (process.env.RC_SERVER_URL || 'https://platform.ringcentral.com').replace(/\/+$/, '');
}

function assertRingCentralConfig() {
  const missing = ['RC_CLIENT_ID', 'RC_CLIENT_SECRET', 'RC_JWT']
    .filter((key) => !process.env[key] || process.env[key].includes('replace_with'));
  if (missing.length) {
    const err = new Error(`RingCentral config is missing: ${missing.join(', ')}.`);
    err.status = 400;
    throw err;
  }
}

async function getRingCentralToken() {
  assertRingCentralConfig();
  if (ringCentralToken && ringCentralToken.expiresAt > Date.now() + 60000) {
    return ringCentralToken.accessToken;
  }

  const auth = Buffer.from(`${process.env.RC_CLIENT_ID}:${process.env.RC_CLIENT_SECRET}`).toString('base64');
  const body = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    assertion: process.env.RC_JWT,
  });

  const response = await fetch(`${ringCentralBaseUrl()}${RINGCENTRAL_TOKEN_URL}`, {
    method: 'POST',
    headers: {
      Authorization: `Basic ${auth}`,
      'Content-Type': 'application/x-www-form-urlencoded',
      Accept: 'application/json',
    },
    body,
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const err = new Error(payload.error_description || payload.message || `RingCentral auth failed with HTTP ${response.status}`);
    err.status = response.status;
    err.ringCentral = payload;
    throw err;
  }

  ringCentralToken = {
    accessToken: payload.access_token,
    expiresAt: Date.now() + (Number(payload.expires_in || 3600) * 1000),
  };
  return ringCentralToken.accessToken;
}

async function ringCentralRequest(method, endpoint, body) {
  const token = await getRingCentralToken();
  const response = await fetch(`${ringCentralBaseUrl()}${endpoint}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'Content-Type': 'application/json',
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const err = new Error(payload.message || payload.error_description || `RingCentral API returned HTTP ${response.status}`);
    err.status = response.status;
    err.ringCentral = payload;
    throw err;
  }
  return payload;
}

function configuredEnvValue(key) {
  const value = String(process.env[key] || '').trim();
  if (!value || value.includes('replace_with')) return '';
  return value;
}

function boolEnv(key, fallback = false) {
  const value = String(process.env[key] || '').trim().toLowerCase();
  if (!value) return fallback;
  return ['1', 'true', 'yes', 'on'].includes(value);
}

function numberEnv(key, fallback) {
  const value = Number(process.env[key]);
  return Number.isFinite(value) ? value : fallback;
}

function assertMicrosoftGraphConfig() {
  const missing = ['MS_GRAPH_TENANT_ID', 'MS_GRAPH_CLIENT_ID', 'MS_GRAPH_CLIENT_SECRET', 'OUTLOOK_SENDER_EMAIL']
    .filter((key) => !configuredEnvValue(key));
  if (missing.length) {
    const err = new Error(`Outlook/Microsoft Graph config is missing: ${missing.join(', ')}.`);
    err.status = 400;
    throw err;
  }
}

async function getMicrosoftGraphToken() {
  assertMicrosoftGraphConfig();
  if (microsoftGraphToken && microsoftGraphToken.expiresAt > Date.now() + 60000) {
    return microsoftGraphToken.accessToken;
  }

  const tenantId = configuredEnvValue('MS_GRAPH_TENANT_ID');
  const body = new URLSearchParams({
    client_id: configuredEnvValue('MS_GRAPH_CLIENT_ID'),
    client_secret: configuredEnvValue('MS_GRAPH_CLIENT_SECRET'),
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const response = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: {
      Accept: 'application/json',
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body,
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const err = new Error(payload.error_description || payload.error || `Microsoft Graph auth failed with HTTP ${response.status}`);
    err.status = response.status;
    err.microsoftGraph = payload;
    throw err;
  }

  microsoftGraphToken = {
    accessToken: payload.access_token,
    expiresAt: Date.now() + (Number(payload.expires_in || 3600) * 1000),
  };
  return microsoftGraphToken.accessToken;
}

async function microsoftGraphRequest(method, endpoint, body) {
  const token = await getMicrosoftGraphToken();
  const url = endpoint.startsWith('https://') ? endpoint : `${MS_GRAPH_BASE_URL}${endpoint}`;
  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'Content-Type': 'application/json',
      Prefer: 'outlook.timezone="America/New_York"',
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (response.status === 202 || response.status === 204) return {};
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const err = new Error(payload.error?.message || payload.error_description || `Microsoft Graph API returned HTTP ${response.status}`);
    err.status = response.status;
    err.microsoftGraph = payload;
    throw err;
  }
  return payload;
}

async function firebaseGetJson(pathname) {
  const cleanPath = String(pathname || '').replace(/^\/+/, '');
  const url = `${FIREBASE_DATABASE_URL.replace(/\/+$/, '')}/${cleanPath}.json`;
  const response = await fetch(url, { headers: { Accept: 'application/json' } });
  const body = await response.json().catch(() => null);
  if (!response.ok) {
    const err = new Error(`Firebase API returned HTTP ${response.status}`);
    err.status = response.status;
    throw err;
  }
  return body;
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function localDateString(date = new Date()) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function tomorrowDateString() {
  return localDateString(addDays(new Date(), 1));
}

function validateDateString(value, fallback = tomorrowDateString()) {
  const date = String(value || fallback).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    const err = new Error('Use date in YYYY-MM-DD format.');
    err.status = 400;
    throw err;
  }
  return date;
}

function isCancelledRide(ride) {
  const status = String(ride?.status || '').toLowerCase();
  return status.includes('cancel') || status === 'noshow' || status === 'no_show';
}

function cleanEmail(value) {
  const email = String(value || '').trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : '';
}

function firstEmail(...values) {
  for (const value of values) {
    const email = cleanEmail(value);
    if (email) return email;
  }
  return '';
}

function normalizeNameKey(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function settingsFacilityList(settings) {
  const raw = settings?.facilities || settings?.facilityList || [];
  if (Array.isArray(raw)) return raw;
  if (raw && typeof raw === 'object') return Object.values(raw);
  return [];
}

function facilityEmailForRide(settings, ride) {
  const facilityKey = normalizeNameKey(ride?.facility || '');
  if (!facilityKey) return '';
  const facility = settingsFacilityList(settings).find((item) => normalizeNameKey(item?.name || item?.facility || '') === facilityKey);
  if (!facility) return '';
  return firstEmail(
    facility.email,
    facility.facilityEmail,
    facility.contactEmail,
    facility.schedulerEmail,
    facility.billingEmail,
    facility.requesterEmail
  );
}

function normalizeCommunicationSource(value) {
  const raw = String(value || '').toLowerCase().trim();
  if (!raw) return '';
  if (['ringcentral_sms', 'ringcentral', 'sms', 'text'].includes(raw)) return 'ringcentral_sms';
  if (['outlook_email', 'outlook', 'email', 'mail'].includes(raw)) return 'outlook_email';
  if (['outlook_calendar_only', 'calendar', 'calendar_import', 'outlook_calendar'].includes(raw)) return 'outlook_calendar_only';
  if (['manual', 'phone', 'unknown'].includes(raw)) return raw;
  return raw;
}

function inferCommunicationSource(ride, channel = '') {
  const explicit = normalizeCommunicationSource(ride?.communicationSource || ride?.contactSource || ride?.originalCommunicationSource || '');
  if (explicit) return explicit;

  const source = [
    ride?.entrySource,
    ride?.importSource,
    ride?.source,
    ride?.sourceImportKey,
    ride?.sourceEmailSubject,
  ].filter(Boolean).join(' ').toLowerCase();

  if (source.includes('ringcentral') || source.includes('sms') || source.includes('text')) return 'ringcentral_sms';
  if (source.includes('email') || source.includes('mail')) return 'outlook_email';
  if (source.includes('outlook') || source.includes('calendar')) return channel === 'email' ? 'outlook_email' : 'outlook_calendar_only';
  if (channel === 'sms') return 'ringcentral_sms';
  if (channel === 'email') return 'outlook_email';
  return 'manual';
}

function inferReminderChannel(ride) {
  const explicit = String(ride?.reminderMethod || ride?.communicationChannel || ride?.contactMethod || '').toLowerCase().trim();
  if (['sms', 'text', 'ringcentral'].includes(explicit)) return 'sms';
  if (['email', 'outlook'].includes(explicit)) return 'email';
  if (['off', 'none', 'no'].includes(explicit)) return 'off';

  const source = [
    ride?.entrySource,
    ride?.importSource,
    ride?.source,
    ride?.sourceImportKey,
    ride?.sourceEmailSubject,
  ].filter(Boolean).join(' ').toLowerCase();
  const hasOutlookEmailThread = Boolean(ride?.sourceOutlookMessageId || ride?.outlookMessageId);
  const hasCalendarOnlySource = source.includes('outlook-ical')
    || source.includes('outlook_calendar_only')
    || source.includes('outlook calendar')
    || source.includes('calendar_import');

  if (source.includes('ringcentral') || source.includes('sms') || source.includes('text')) return 'sms';
  if (source.includes('email') || source.includes('mail') || hasOutlookEmailThread) return 'email';
  if (cleanPhone(ride?.clientPhone || ride?.phone || ride?.requesterPhone)) return 'sms';
  if (firstEmail(ride?.clientEmail, ride?.requesterEmail, ride?.email, ride?.facilityEmail)) return 'email';
  if (hasCalendarOnlySource) return 'review';
  if (source.includes('outlook')) return 'review';
  return 'review';
}

function displayTime(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const match = raw.match(/^(\d{1,2}):(\d{2})/);
  if (!match) return raw;
  let hour = Number(match[1]);
  const minute = match[2];
  const suffix = hour >= 12 ? 'PM' : 'AM';
  hour = hour % 12 || 12;
  return `${hour}:${minute} ${suffix}`;
}

function displayDate(value) {
  const raw = String(value || '').trim();
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return raw;
  return `${match[2]}/${match[3]}/${match[1]}`;
}

function buildReminderText(ride) {
  const client = String(ride.client || 'your ride').trim();
  const facility = String(ride.facility || '').trim();
  const isPrivatePay = /^private\s*pay$/i.test(facility);
  const rideLabel = facility && !isPrivatePay && client && client.toLowerCase() !== facility.toLowerCase()
    ? `${facility} patient ${client}`
    : client;
  const time = displayTime(ride.time);
  const pickup = String(ride.pickup || '').trim();
  const dropoff = String(ride.dropoff || '').trim();
  const parts = [
    `Reminder from AbleCare Mobility: ${rideLabel}${time ? ` is scheduled for pickup tomorrow at ${time}` : ' has transportation scheduled for tomorrow'}.`,
  ];
  if (pickup) parts.push(`Pickup: ${pickup}.`);
  if (dropoff) parts.push(`Drop off: ${dropoff}.`);
  parts.push('Reply here or call AbleCare if anything changed.');
  return parts.join(' ');
}

function buildRideConfirmationText(ride) {
  const client = String(ride.client || ride.facility || 'the rider').trim();
  const date = displayDate(ride.date);
  const time = displayTime(ride.time) || String(ride.time || '').trim() || 'not set';
  const pickup = String(ride.pickup || '').trim() || 'not saved';
  const dropoff = String(ride.dropoff || '').trim() || 'not saved';
  return [
    'Is this all correct?',
    '',
    `Name: ${client}`,
    `Pickup date: ${date || 'not set'}`,
    `Pickup time: ${time}`,
    `Pickup address: ${pickup}`,
    `Dropoff address: ${dropoff}`,
    '',
    'Please reply yes if everything is correct, or send any changes.'
  ].join('\n');
}

function reminderLogKey(date, rideKey) {
  return `${date}_${safeFirebaseKey(rideKey)}`;
}

async function patchFirebase(pathname, value) {
  const url = `${FIREBASE_DATABASE_URL.replace(/\/+$/, '')}/${pathname.replace(/^\/+/, '')}.json`;
  const response = await fetch(url, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(value),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const err = new Error(payload.error || `Firebase update failed with HTTP ${response.status}`);
    err.status = response.status;
    throw err;
  }
  return payload;
}

async function listReminderCandidates(date) {
  const [ridesRaw, sentRaw, settingsRaw] = await Promise.all([
    firebaseGetJson('rides').catch(() => ({})),
    firebaseGetJson(`reminderLogs/${date}`).catch(() => ({})),
    firebaseGetJson('settings').catch(() => ({})),
  ]);
  const sent = sentRaw || {};
  const rides = Object.entries(ridesRaw || {})
    .filter(([, ride]) => ride && ride.date === date && !isCancelledRide(ride))
    .map(([key, ride]) => {
      const facilityEmail = facilityEmailForRide(settingsRaw, ride);
      const enrichedRide = { ...ride, facilityEmail };
      const channel = inferReminderChannel(enrichedRide);
      const communicationSource = inferCommunicationSource(enrichedRide, channel);
      const toPhone = cleanPhone(ride.clientPhone || ride.phone || ride.requesterPhone || '');
      const toEmail = firstEmail(ride.clientEmail, ride.requesterEmail, ride.email, facilityEmail);
      const sourceOutlookMessageId = ride.sourceOutlookMessageId || ride.outlookMessageId || '';
      const sourceOutlookConversationId = ride.sourceOutlookConversationId || ride.outlookConversationId || '';
      const canReplyToOutlookThread = channel === 'email' && Boolean(sourceOutlookMessageId);
      const sentRecord = sent[reminderLogKey(date, key)] || null;
      const blockedReason = channel === 'off'
        ? 'Reminder off for this ride.'
        : channel === 'sms' && !toPhone
          ? 'SMS selected but no client phone is saved.'
          : channel === 'email' && !toEmail && !canReplyToOutlookThread
            ? 'Email selected but no client email is saved and no original Outlook thread is linked.'
            : channel === 'review'
              ? (communicationSource === 'outlook_calendar_only'
                ? 'Imported from Outlook Calendar only; add/link an email or phone before sending.'
                : 'Choose email or SMS before sending.')
              : '';
      return {
        rideKey: key,
        date,
        channel,
        communicationSource,
        sourceOutlookMessageId,
        sourceOutlookConversationId,
        toPhone,
        toEmail,
        text: buildReminderText(enrichedRide),
        sent: Boolean(sentRecord),
        sentRecord,
        blockedReason,
        ride: {
          facility: ride.facility || '',
          client: ride.client || '',
          time: ride.time || '',
          pickup: ride.pickup || '',
          dropoff: ride.dropoff || '',
          driver: ride.driver || '',
          facilityEmail,
          entrySource: ride.entrySource || '',
          communicationSource,
          sourceMessageId: ride.sourceMessageId || '',
          sourceOutlookMessageId,
          sourceOutlookConversationId,
        },
      };
    })
    .sort((a, b) => String(a.ride.time || '').localeCompare(String(b.ride.time || '')));
  return rides;
}

async function sendReminderEmail(candidate) {
  const sender = configuredEnvValue('OUTLOOK_SENDER_EMAIL');
  const subjectClient = candidate.ride?.client || candidate.ride?.facility || 'your ride';
  const sourceMessageId = String(candidate.sourceOutlookMessageId || candidate.ride?.sourceOutlookMessageId || '').trim();
  if (sourceMessageId) {
    await microsoftGraphRequest('POST', `/users/${encodeURIComponent(sender)}/messages/${encodeURIComponent(sourceMessageId)}/reply`, {
      comment: candidate.text,
    });
    return {
      id: sourceMessageId,
      subject: candidate.ride?.sourceEmailSubject || '',
      to: candidate.toEmail,
      sender,
      replyToMessageId: sourceMessageId,
    };
  }

  const message = {
    message: {
      subject: `AbleCare Mobility ride reminder for tomorrow - ${subjectClient}`,
      body: {
        contentType: 'Text',
        content: candidate.text,
      },
      toRecipients: [
        { emailAddress: { address: candidate.toEmail } },
      ],
    },
    saveToSentItems: true,
  };
  await microsoftGraphRequest('POST', `/users/${encodeURIComponent(sender)}/sendMail`, message);
  return {
    id: '',
    subject: message.message.subject,
    to: candidate.toEmail,
    sender,
    replyToMessageId: '',
  };
}

async function sendConfirmationEmail({ ride, toEmail, text }) {
  const sender = configuredEnvValue('OUTLOOK_SENDER_EMAIL');
  const subjectClient = ride?.client || ride?.facility || 'ride';
  const sourceMessageId = String(ride?.sourceOutlookMessageId || ride?.outlookMessageId || '').trim();
  if (sourceMessageId) {
    await microsoftGraphRequest('POST', `/users/${encodeURIComponent(sender)}/messages/${encodeURIComponent(sourceMessageId)}/reply`, {
      comment: text,
    });
    return {
      id: sourceMessageId,
      subject: ride?.sourceEmailSubject || '',
      to: toEmail,
      sender,
      replyToMessageId: sourceMessageId,
    };
  }

  const message = {
    message: {
      subject: `AbleCare Mobility ride confirmation - ${subjectClient}`,
      body: {
        contentType: 'Text',
        content: text,
      },
      toRecipients: [
        { emailAddress: { address: toEmail } },
      ],
    },
    saveToSentItems: true,
  };
  await microsoftGraphRequest('POST', `/users/${encodeURIComponent(sender)}/sendMail`, message);
  return {
    id: '',
    subject: message.message.subject,
    to: toEmail,
    sender,
    replyToMessageId: '',
  };
}

function parseGraphDateRange(searchParams) {
  const from = validateDateString(searchParams.get('from'), localDateString(new Date()));
  const to = validateDateString(searchParams.get('to'), from);
  const begin = new Date(`${from}T00:00:00`);
  const end = new Date(`${to}T00:00:00`);
  end.setDate(end.getDate() + 1);
  if (begin > end) {
    const err = new Error('Invalid date range.');
    err.status = 400;
    throw err;
  }
  return {
    from,
    to,
    startDateTime: begin.toISOString(),
    endDateTime: end.toISOString(),
  };
}

function normalizeOutlookEvent(event) {
  return {
    id: event.id || '',
    subject: event.subject || '',
    start: event.start || {},
    end: event.end || {},
    location: event.location?.displayName || '',
    bodyPreview: event.bodyPreview || '',
    organizer: event.organizer?.emailAddress || null,
    webLink: event.webLink || '',
    isCancelled: Boolean(event.isCancelled),
  };
}

function normalizeOutlookMessage(message) {
  return {
    id: message.id || '',
    conversationId: message.conversationId || '',
    subject: message.subject || '',
    from: message.from?.emailAddress || null,
    receivedDateTime: message.receivedDateTime || '',
    bodyPreview: message.bodyPreview || '',
    webLink: message.webLink || '',
  };
}

function outlookMessageImportFields(message) {
  const normalized = normalizeOutlookMessage(message || {});
  return {
    sourceOutlookMessageId: normalized.id,
    sourceOutlookConversationId: normalized.conversationId,
    requesterEmail: normalized.from?.address || '',
    requesterName: normalized.from?.name || '',
    subject: normalized.subject,
    rawEmail: normalized.bodyPreview || '',
    source: 'outlook-email',
  };
}

async function sendReminder(candidate, requestedChannel) {
  if (!candidate) {
    const err = new Error('Ride is not eligible for reminders on that date.');
    err.status = 404;
    throw err;
  }
  if (candidate.sent) {
    const err = new Error('Reminder was already sent for this ride/date.');
    err.status = 409;
    throw err;
  }

  const channel = requestedChannel || candidate.channel;
  if (channel === 'sms') {
    if (!candidate.toPhone) {
      const err = new Error('Cannot send SMS reminder because no client phone is saved.');
      err.status = 400;
      throw err;
    }
    const from = cleanPhone(process.env.RC_FROM_NUMBER);
    if (!from) {
      const err = new Error('RC_FROM_NUMBER is missing from .env.');
      err.status = 400;
      throw err;
    }
    const sms = await ringCentralRequest('POST', '/restapi/v1.0/account/~/extension/~/sms', {
      from: { phoneNumber: from },
      to: [{ phoneNumber: candidate.toPhone }],
      text: candidate.text,
    });
    return { channel, provider: 'ringcentral', providerResult: sms };
  }

  if (channel === 'email') {
    const sourceMessageId = String(candidate.sourceOutlookMessageId || candidate.ride?.sourceOutlookMessageId || '').trim();
    if (!candidate.toEmail && !sourceMessageId) {
      const err = new Error('Cannot send email reminder because no client email is saved.');
      err.status = 400;
      throw err;
    }
    const email = await sendReminderEmail(candidate);
    return { channel, provider: 'outlook', providerResult: email };
  }

  const err = new Error('Choose SMS or email before sending this reminder.');
  err.status = 400;
  throw err;
}

async function writeReminderLog(date, rideKey, candidate, result, source = 'manual') {
  const logKey = reminderLogKey(date, rideKey);
  const log = {
    rideKey,
    date,
    channel: result.channel,
    provider: result.provider,
    communicationSource: candidate.communicationSource || inferCommunicationSource(candidate.ride, result.channel),
    toPhone: candidate.toPhone || '',
    toEmail: candidate.toEmail || '',
    text: candidate.text,
    sentAt: Date.now(),
    source,
    providerResult: result.providerResult || null,
  };
  await writeFirebase(`reminderLogs/${date}/${logKey}`, log);
  await patchFirebase(`rides/${rideKey}`, {
    lastReminderAt: log.sentAt,
    lastReminderChannel: result.channel,
    lastReminderProvider: result.provider,
    communicationSource: log.communicationSource,
  });
  return log;
}

async function runReminderBatch(date, options = {}) {
  const source = options.source || 'manual';
  const dryRun = Boolean(options.dryRun);
  const candidates = await listReminderCandidates(date);
  const ready = candidates.filter((item) => !item.sent && !item.blockedReason);
  const sent = [];
  const failed = [];

  if (!dryRun) {
    for (const candidate of ready) {
      try {
        const result = await sendReminder(candidate, null);
        const log = await writeReminderLog(date, candidate.rideKey, candidate, result, source);
        sent.push({ rideKey: candidate.rideKey, channel: result.channel, log });
      } catch (error) {
        failed.push({
          rideKey: candidate.rideKey,
          channel: candidate.channel,
          error: error.message || 'Reminder send failed.',
        });
      }
    }
  }

  const runLog = {
    date,
    source,
    dryRun,
    checkedAt: Date.now(),
    candidateCount: candidates.length,
    readyCount: ready.length,
    sentCount: sent.length,
    failedCount: failed.length,
    sent,
    failed,
    blocked: candidates
      .filter((item) => item.blockedReason)
      .map((item) => ({
        rideKey: item.rideKey,
        channel: item.channel,
        reason: item.blockedReason,
      })),
  };

  await writeFirebase(`reminderRuns/${date}/${safeFirebaseKey(`${source}_${Date.now()}`)}`, runLog).catch(() => null);
  return runLog;
}

function shouldRunReminderScheduler(now = new Date()) {
  if (!boolEnv('REMINDER_AUTOSEND_ENABLED', false)) return false;
  const hour = numberEnv('REMINDER_SEND_HOUR', 17);
  const minute = numberEnv('REMINDER_SEND_MINUTE', 0);
  return now.getHours() === hour && now.getMinutes() === minute;
}

async function tickReminderScheduler() {
  const now = new Date();
  if (!shouldRunReminderScheduler(now)) return;
  const targetDate = localDateString(addDays(now, numberEnv('REMINDER_LOOKAHEAD_DAYS', 1)));
  const runKey = `${localDateString(now)}_${now.getHours()}_${now.getMinutes()}_${targetDate}`;
  if (reminderSchedulerLastRunKey === runKey) return;
  reminderSchedulerLastRunKey = runKey;
  try {
    await runReminderBatch(targetDate, { source: 'auto-scheduler' });
  } catch (error) {
    await writeFirebase(`reminderRuns/${targetDate}/auto_error_${Date.now()}`, {
      date: targetDate,
      source: 'auto-scheduler',
      checkedAt: Date.now(),
      error: error.message || 'Reminder scheduler failed.',
    }).catch(() => null);
  }
}

function startReminderScheduler() {
  if (reminderSchedulerTimer) return;
  reminderSchedulerTimer = setInterval(tickReminderScheduler, 60 * 1000);
  tickReminderScheduler().catch(() => {});
}

async function sendExpoPushNotification(message) {
  const response = await fetch('https://exp.host/--/api/v2/push/send', {
    method: 'POST',
    headers: {
      Accept: 'application/json',
      'Accept-Encoding': 'gzip, deflate',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(message),
  });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    const err = new Error(body?.errors?.[0]?.message || `Expo push API returned HTTP ${response.status}`);
    err.status = response.status;
    err.expo = body;
    throw err;
  }
  return body;
}

function cleanPhone(value) {
  return String(value || '').replace(/[^\d+]/g, '');
}

function safeFirebaseKey(value) {
  return String(value || `rc-${Date.now()}`).replace(/[.#$/\[\]\s]/g, '_').slice(0, 180);
}

function normalizeRingCentralMessage(record) {
  const from = record.from?.phoneNumber || record.from?.name || '';
  const to = (record.to || []).map((item) => item.phoneNumber || item.name || '').filter(Boolean).join(', ');
  const text = record.subject || record.text || record.message || '';
  return {
    status: 'pending',
    source: 'ringcentral',
    sourceMessageId: String(record.id || record.uri || ''),
    createdAt: Date.now(),
    receivedAt: record.creationTime || new Date().toISOString(),
    requesterPhone: cleanPhone(from),
    requesterName: record.from?.name || '',
    requesterEmail: '',
    rawMessage: text,
    ringCentral: {
      id: record.id || '',
      uri: record.uri || '',
      direction: record.direction || '',
      type: record.type || record.messageType || 'SMS',
      from,
      to,
      subject: record.subject || '',
      creationTime: record.creationTime || '',
      readStatus: record.readStatus || '',
      messageStatus: record.messageStatus || '',
    },
    ride: {
      facility: '',
      client: '',
      date: '',
      time: '',
      clientPhone: cleanPhone(from),
      clientEmail: '',
      reminderMethod: 'sms',
      communicationSource: 'ringcentral_sms',
      reminderContactName: record.from?.name || '',
      pickup: '',
      dropoff: '',
      ridetype: 'wc',
      triptype: '1way',
      notes: text ? `RingCentral text from ${from}: ${text}` : `RingCentral message from ${from}`,
    },
  };
}

function normalizeRingCentralCall(record) {
  const from = record.from?.phoneNumber || record.from?.name || '';
  const to = record.to?.phoneNumber || record.to?.name || '';
  const result = record.result || record.action || '';
  return {
    status: 'pending',
    source: 'ringcentral-call',
    sourceMessageId: String(record.id || record.sessionId || record.telephonySessionId || ''),
    createdAt: Date.now(),
    receivedAt: record.startTime || new Date().toISOString(),
    requesterPhone: cleanPhone(from),
    requesterName: record.from?.name || '',
    requesterEmail: '',
    rawMessage: `RingCentral call from ${from}${to ? ` to ${to}` : ''}${result ? ` (${result})` : ''}`,
    ringCentral: {
      id: record.id || '',
      sessionId: record.sessionId || record.telephonySessionId || '',
      direction: record.direction || '',
      type: record.type || 'Voice',
      from,
      to,
      startTime: record.startTime || '',
      duration: record.duration || 0,
      result,
    },
    ride: {
      facility: '',
      client: '',
      date: '',
      time: '',
      clientPhone: cleanPhone(from),
      clientEmail: '',
      reminderMethod: 'sms',
      communicationSource: 'ringcentral_sms',
      reminderContactName: record.from?.name || '',
      pickup: '',
      dropoff: '',
      ridetype: 'wc',
      triptype: '1way',
      notes: `Follow up on RingCentral call from ${from}${result ? ` (${result})` : ''}.`,
    },
  };
}

async function writeFirebase(pathname, value) {
  const url = `${FIREBASE_DATABASE_URL.replace(/\/+$/, '')}/${pathname.replace(/^\/+/, '')}.json`;
  const response = await fetch(url, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(value),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const err = new Error(payload.error || `Firebase write failed with HTTP ${response.status}`);
    err.status = response.status;
    throw err;
  }
  return payload;
}

async function importRingCentralMessage(record) {
  const request = normalizeRingCentralMessage(record);
  const sourceId = request.sourceMessageId || `${request.requesterPhone}-${request.receivedAt}`;
  const key = `ringcentral_${safeFirebaseKey(sourceId)}`;
  await writeFirebase(`rideRequests/${key}`, request);
  return { key, request };
}

async function importRingCentralCall(record) {
  const request = normalizeRingCentralCall(record);
  const sourceId = request.sourceMessageId || `${request.requesterPhone}-${request.receivedAt}`;
  const key = `ringcentral_call_${safeFirebaseKey(sourceId)}`;
  await writeFirebase(`rideRequests/${key}`, request);
  return { key, request };
}

async function listPayments(beginTime, endTime) {
  const payments = [];
  let cursor = '';

  do {
    const body = await squareGet('/v2/payments', {
      begin_time: beginTime,
      end_time: endTime,
      sort_order: 'DESC',
      limit: 100,
      cursor,
    });
    payments.push(...(body.payments || []));
    cursor = body.cursor || '';
  } while (cursor);

  return payments;
}

function moneyAmount(money) {
  return Number(money?.amount || 0) / 100;
}

function summarizePayments(payments) {
  const completed = payments.filter((payment) => payment.status === 'COMPLETED');
  const total = completed.reduce((sum, payment) => sum + moneyAmount(payment.total_money), 0);
  const refunded = completed.reduce((sum, payment) => sum + moneyAmount(payment.refunded_money), 0);
  const fees = completed.reduce((sum, payment) => {
    return sum + (payment.processing_fee || []).reduce((feeSum, fee) => feeSum + moneyAmount(fee.amount_money), 0);
  }, 0);

  const bySource = {};
  for (const payment of completed) {
    const source = payment.source_type || 'UNKNOWN';
    if (!bySource[source]) bySource[source] = { count: 0, total: 0 };
    bySource[source].count += 1;
    bySource[source].total += moneyAmount(payment.total_money);
  }

  return {
    count: completed.length,
    total: Number(total.toFixed(2)),
    refunded: Number(refunded.toFixed(2)),
    net: Number((total - refunded).toFixed(2)),
    fees: Number(fees.toFixed(2)),
    bySource,
    recent: completed.slice(0, 10).map((payment) => ({
      id: payment.id,
      createdAt: payment.created_at,
      status: payment.status,
      sourceType: payment.source_type || '',
      amount: moneyAmount(payment.total_money),
      refunded: moneyAmount(payment.refunded_money),
      receiptUrl: payment.receipt_url || '',
      note: payment.note || '',
    })),
  };
}

function serveStatic(req, res, pathname) {
  const requestedPath = pathname === '/' ? '/index.html' : decodeURIComponent(pathname);
  const normalized = path.normalize(requestedPath).replace(/^(\.\.[/\\])+/, '');
  const filePath = path.join(ROOT, normalized);

  if (!filePath.startsWith(ROOT) || path.basename(filePath).startsWith('.')) {
    res.writeHead(404);
    res.end('Not found');
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found');
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    const types = {
      '.html': 'text/html; charset=utf-8',
      '.js': 'text/javascript; charset=utf-8',
      '.css': 'text/css; charset=utf-8',
      '.json': 'application/json; charset=utf-8',
      '.png': 'image/png',
      '.jpg': 'image/jpeg',
      '.jpeg': 'image/jpeg',
      '.svg': 'image/svg+xml',
      '.ico': 'image/x-icon',
    };
    res.writeHead(200, {
      'Content-Type': types[ext] || 'application/octet-stream',
      'Cache-Control': 'no-store, no-cache, must-revalidate, proxy-revalidate',
      Pragma: 'no-cache',
      Expires: '0',
    });
    res.end(data);
  });
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);

  try {
    if (url.pathname === '/api/health') {
      sendJson(res, 200, {
        ok: true,
        service: 'ablecare-fleet-dashboard',
        checkedAt: new Date().toISOString(),
      });
      return;
    }

    if (url.pathname === '/api/square/status') {
      const locations = await squareGet('/v2/locations');
      sendJson(res, 200, {
        ok: true,
        environment: process.env.SQUARE_ENVIRONMENT || 'sandbox',
        locations: locations.locations?.map((location) => ({
          id: location.id,
          name: location.name,
        })) || [],
      });
      return;
    }

    if (url.pathname === '/api/square/sales-summary') {
      const range = parseDateRange(url.searchParams);
      const payments = await listPayments(range.beginTime, range.endTime);
      sendJson(res, 200, {
        ok: true,
        environment: process.env.SQUARE_ENVIRONMENT || 'sandbox',
        ...range,
        ...summarizePayments(payments),
      });
      return;
    }

    if (url.pathname === '/api/ringcentral/status') {
      const extension = await ringCentralRequest('GET', '/restapi/v1.0/account/~/extension/~');
      sendJson(res, 200, {
        ok: true,
        server: ringCentralBaseUrl(),
        extension: {
          id: extension.id,
          extensionNumber: extension.extensionNumber,
          name: [extension.contact?.firstName, extension.contact?.lastName].filter(Boolean).join(' '),
          email: extension.contact?.email || '',
        },
      });
      return;
    }

    if (url.pathname === '/api/outlook/status') {
      const sender = configuredEnvValue('OUTLOOK_SENDER_EMAIL');
      const mailbox = await microsoftGraphRequest('GET', `/users/${encodeURIComponent(sender)}/mailFolders/inbox?$select=id,displayName,totalItemCount`);
      sendJson(res, 200, {
        ok: true,
        sender,
        mailbox: {
          id: mailbox.id || '',
          displayName: mailbox.displayName || 'Inbox',
          totalItemCount: Number(mailbox.totalItemCount || 0),
        },
      });
      return;
    }

    if (url.pathname === '/api/outlook/calendar') {
      const sender = configuredEnvValue('OUTLOOK_SENDER_EMAIL');
      const range = parseGraphDateRange(url.searchParams);
      const params = new URLSearchParams({
        startDateTime: range.startDateTime,
        endDateTime: range.endDateTime,
        '$select': 'id,subject,start,end,location,bodyPreview,organizer,webLink,isCancelled',
        '$orderby': 'start/dateTime',
        '$top': String(Math.min(Number(url.searchParams.get('top') || 50), 100)),
      });
      const events = await microsoftGraphRequest('GET', `/users/${encodeURIComponent(sender)}/calendarView?${params}`);
      sendJson(res, 200, {
        ok: true,
        ...range,
        events: (events.value || []).map(normalizeOutlookEvent),
      });
      return;
    }

    if (url.pathname === '/api/outlook/messages') {
      const sender = configuredEnvValue('OUTLOOK_SENDER_EMAIL');
      const search = String(url.searchParams.get('q') || '').trim();
      const top = String(Math.min(Number(url.searchParams.get('top') || 25), 50));
      const params = new URLSearchParams({
        '$select': 'id,conversationId,subject,from,receivedDateTime,bodyPreview,webLink',
        '$orderby': 'receivedDateTime desc',
        '$top': top,
      });
      if (search) params.set('$search', `"${search.replace(/"/g, '\\"')}"`);
      const messages = await microsoftGraphRequest('GET', `/users/${encodeURIComponent(sender)}/messages?${params}`);
      sendJson(res, 200, {
        ok: true,
        query: search,
        messages: (messages.value || []).map(normalizeOutlookMessage),
      });
      return;
    }

    if (url.pathname === '/api/reminders/preview') {
      const date = validateDateString(url.searchParams.get('date'));
      const candidates = await listReminderCandidates(date);
      sendJson(res, 200, {
        ok: true,
        date,
        count: candidates.length,
        readyCount: candidates.filter((item) => !item.sent && !item.blockedReason).length,
        sentCount: candidates.filter((item) => item.sent).length,
        candidates,
      });
      return;
    }

    if (url.pathname === '/api/reminders/contact' && req.method === 'POST') {
      const body = await readRequestBody(req);
      const rideKey = String(body.rideKey || '').trim();
      const method = String(body.method || '').toLowerCase().trim();
      const contact = String(body.contact || '').trim();
      if (!rideKey) {
        sendJson(res, 400, { ok: false, error: 'rideKey is required.' });
        return;
      }
      if (!['sms', 'email'].includes(method)) {
        sendJson(res, 400, { ok: false, error: 'Choose Text or Email.' });
        return;
      }
      if (!contact) {
        sendJson(res, 400, { ok: false, error: method === 'email' ? 'Enter an email address.' : 'Enter a phone number.' });
        return;
      }

      const patch = {
        updatedAt: Date.now(),
        updatedBy: 'Dashboard',
        updatedByRole: 'owner',
        missingReminderContact: null,
        reviewFlag: null,
        contactSource: 'manual-reminder-entry',
      };
      if (method === 'email') {
        const email = cleanEmail(contact);
        if (!email) {
          sendJson(res, 400, { ok: false, error: 'Enter a valid email address.' });
          return;
        }
        patch.clientEmail = email;
        patch.reminderMethod = 'email';
        patch.communicationSource = 'manual_email';
      } else {
        const digits = cleanPhone(contact);
        if (digits.length < 10) {
          sendJson(res, 400, { ok: false, error: 'Enter a valid phone number.' });
          return;
        }
        patch.clientPhone = digits.length === 11 && digits.startsWith('1') ? digits.slice(1) : digits;
        patch.reminderMethod = 'sms';
        patch.communicationSource = 'ringcentral_sms';
      }

      await patchFirebase(`rides/${safeFirebaseKey(rideKey)}`, patch);
      sendJson(res, 200, { ok: true, rideKey, method, contact: method === 'email' ? patch.clientEmail : patch.clientPhone });
      return;
    }

    if (url.pathname === '/api/ride-confirmations/send' && req.method === 'POST') {
      const body = await readRequestBody(req);
      const rideKey = String(body.rideKey || '').trim();
      const channel = String(body.channel || '').toLowerCase().trim();
      const contact = String(body.contact || '').trim();
      const sentBy = String(body.sentBy || 'Dashboard').trim();
      const text = String(body.text || '').trim();
      if (!rideKey) {
        sendJson(res, 400, { ok: false, error: 'rideKey is required.' });
        return;
      }
      if (!['sms', 'email'].includes(channel)) {
        sendJson(res, 400, { ok: false, error: 'Choose Text or Email.' });
        return;
      }
      const ride = await firebaseGetJson(`rides/${safeFirebaseKey(rideKey)}`);
      if (!ride) {
        sendJson(res, 404, { ok: false, error: 'Ride was not found.' });
        return;
      }
      const messageText = text || buildRideConfirmationText(ride);
      let result;
      let toPhone = '';
      let toEmail = '';
      if (channel === 'sms') {
        toPhone = cleanPhone(contact || ride.clientPhone || ride.phone || ride.requesterPhone || '');
        const from = cleanPhone(process.env.RC_FROM_NUMBER);
        if (!toPhone) {
          sendJson(res, 400, { ok: false, error: 'Enter a valid phone number.' });
          return;
        }
        if (!from) {
          sendJson(res, 400, { ok: false, error: 'RC_FROM_NUMBER is missing from .env.' });
          return;
        }
        result = {
          channel,
          provider: 'ringcentral',
          providerResult: await ringCentralRequest('POST', '/restapi/v1.0/account/~/extension/~/sms', {
            from: { phoneNumber: from },
            to: [{ phoneNumber: toPhone }],
            text: messageText,
          }),
        };
      } else {
        toEmail = firstEmail(contact, ride.clientEmail, ride.requesterEmail, ride.email, ride.facilityEmail);
        const sourceMessageId = String(ride.sourceOutlookMessageId || ride.outlookMessageId || '').trim();
        if (!toEmail && !sourceMessageId) {
          sendJson(res, 400, { ok: false, error: 'Enter a valid email address or link the original Outlook thread.' });
          return;
        }
        result = {
          channel,
          provider: 'outlook',
          providerResult: await sendConfirmationEmail({ ride, toEmail, text: messageText }),
        };
      }
      const sentAt = Date.now();
      const logKey = safeFirebaseKey(`${rideKey}_${sentAt}`);
      const log = {
        rideKey,
        date: ride.date || '',
        channel,
        provider: result.provider,
        toPhone,
        toEmail,
        text: messageText,
        sentAt,
        sentBy,
        providerResult: result.providerResult || null,
      };
      await writeFirebase(`rideConfirmationLogs/${ride.date || 'unknown'}/${logKey}`, log);
      await patchFirebase(`rides/${safeFirebaseKey(rideKey)}`, {
        lastConfirmationAt: sentAt,
        lastConfirmationChannel: channel,
        lastConfirmationProvider: result.provider,
        lastConfirmationBy: sentBy,
        ...(channel === 'sms' ? { clientPhone: toPhone, reminderMethod: 'sms', communicationSource: 'ringcentral_sms' } : {}),
        ...(channel === 'email' && toEmail ? { clientEmail: toEmail, reminderMethod: 'email', communicationSource: 'outlook_email' } : {}),
      });
      sendJson(res, 200, { ok: true, log });
      return;
    }

    if (url.pathname === '/api/reminders/send' && req.method === 'POST') {
      const body = await readRequestBody(req);
      const date = validateDateString(body.date);
      const rideKey = String(body.rideKey || '').trim();
      const requestedChannel = String(body.channel || '').toLowerCase().trim();
      if (!rideKey) {
        sendJson(res, 400, { ok: false, error: 'rideKey is required.' });
        return;
      }
      const candidates = await listReminderCandidates(date);
      const candidate = candidates.find((item) => item.rideKey === rideKey);
      const result = await sendReminder(candidate, requestedChannel || null);
      const log = await writeReminderLog(date, rideKey, candidate, result, 'manual');
      sendJson(res, 200, { ok: true, log });
      return;
    }

    if (url.pathname === '/api/reminders/run' && req.method === 'POST') {
      const body = await readRequestBody(req);
      const date = validateDateString(body.date);
      const dryRun = Boolean(body.dryRun);
      const run = await runReminderBatch(date, { source: 'manual-run', dryRun });
      sendJson(res, 200, { ok: true, run });
      return;
    }

    if (url.pathname === '/api/ringcentral/messages') {
      const params = new URLSearchParams({
        messageType: 'SMS',
        direction: url.searchParams.get('direction') || 'Inbound',
        perPage: url.searchParams.get('perPage') || '20',
      });
      if (url.searchParams.get('dateFrom')) params.set('dateFrom', url.searchParams.get('dateFrom'));
      if (url.searchParams.get('dateTo')) params.set('dateTo', url.searchParams.get('dateTo'));
      const messages = await ringCentralRequest('GET', `/restapi/v1.0/account/~/extension/~/message-store?${params}`);
      sendJson(res, 200, {
        ok: true,
        records: messages.records || [],
        navigation: messages.navigation || {},
      });
      return;
    }

    if (url.pathname === '/api/ringcentral/phone-numbers') {
      const numbers = await ringCentralRequest('GET', '/restapi/v1.0/account/~/extension/~/phone-number');
      sendJson(res, 200, {
        ok: true,
        records: numbers.records || [],
      });
      return;
    }

    if (url.pathname === '/api/ringcentral/import-messages' && req.method === 'POST') {
      const body = await readRequestBody(req);
      const records = Array.isArray(body.records) ? body.records : [];
      const imported = [];
      for (const record of records) {
        imported.push(await importRingCentralMessage(record));
      }
      sendJson(res, 200, { ok: true, importedCount: imported.length, imported });
      return;
    }

    if (url.pathname === '/api/ringcentral/calls') {
      const params = new URLSearchParams({
        perPage: url.searchParams.get('perPage') || '20',
      });
      if (url.searchParams.get('dateFrom')) params.set('dateFrom', url.searchParams.get('dateFrom'));
      if (url.searchParams.get('dateTo')) params.set('dateTo', url.searchParams.get('dateTo'));
      const calls = await ringCentralRequest('GET', `/restapi/v1.0/account/~/extension/~/call-log?${params}`);
      sendJson(res, 200, {
        ok: true,
        records: calls.records || [],
        navigation: calls.navigation || {},
      });
      return;
    }

    if (url.pathname === '/api/ringcentral/import-calls' && req.method === 'POST') {
      const body = await readRequestBody(req);
      const records = Array.isArray(body.records) ? body.records : [];
      const imported = [];
      for (const record of records) {
        imported.push(await importRingCentralCall(record));
      }
      sendJson(res, 200, { ok: true, importedCount: imported.length, imported });
      return;
    }

    if (url.pathname === '/api/ringcentral/sms' && req.method === 'POST') {
      const body = await readRequestBody(req);
      const to = cleanPhone(body.to);
      const text = String(body.text || '').trim();
      const from = cleanPhone(body.from || process.env.RC_FROM_NUMBER);
      if (!to || !text) {
        sendJson(res, 400, { ok: false, error: 'Send SMS requires to and text.' });
        return;
      }
      if (!from) {
        sendJson(res, 400, { ok: false, error: 'RC_FROM_NUMBER is missing from .env.' });
        return;
      }
      const sms = await ringCentralRequest('POST', '/restapi/v1.0/account/~/extension/~/sms', {
        from: { phoneNumber: from },
        to: [{ phoneNumber: to }],
        text,
      });
      sendJson(res, 200, { ok: true, message: sms });
      return;
    }

    if (url.pathname === '/api/push/return-ride' && req.method === 'POST') {
      const body = await readRequestBody(req);
      const driver = String(body.driver || '').trim();
      const rideKey = String(body.rideKey || '').trim();
      if (!driver || !rideKey) {
        sendJson(res, 400, { ok: false, error: 'Return ride push requires driver and rideKey.' });
        return;
      }

      const driverKey = safeFirebaseKey(driver);
      const [driverRecord, tokenRecord] = await Promise.all([
        firebaseGetJson(`drivers/${driverKey}`).catch(() => null),
        firebaseGetJson(`driverPushTokens/${driverKey}`).catch(() => null),
      ]);
      const token = body.pushToken || tokenRecord?.token || driverRecord?.pushToken;
      if (!token) {
        sendJson(res, 200, { ok: false, sent: false, reason: `No push token saved for ${driver}. Ask the driver to open the updated app and allow notifications.` });
        return;
      }

      const client = String(body.client || 'Client').trim();
      const time = String(body.time || '').trim();
      const pushResult = await sendExpoPushNotification({
        to: token,
        sound: 'default',
        title: body.title || 'New Return Ride',
        body: body.message || `${client} is ready for return pickup${time ? ` at ${time}` : ''}.`,
        priority: 'high',
        channelId: 'return-rides',
        data: {
          type: 'return-ride',
          rideKey,
          driver,
          date: body.date || '',
        },
      });

      sendJson(res, 200, { ok: true, sent: true, result: pushResult });
      return;
    }

    if (url.pathname === '/api/ringcentral/webhook' && req.method === 'POST') {
      const validationToken = req.headers['validation-token'];
      if (validationToken) {
        res.setHeader('Validation-Token', validationToken);
      }
      const body = await readRequestBody(req);
      const notifications = Array.isArray(body) ? body : [body];
      const imported = [];
      for (const notification of notifications) {
        const record = notification?.body || notification;
        if (record && (record.id || record.from || record.subject || record.messageStatus)) {
          imported.push(await importRingCentralMessage(record));
        }
      }
      sendJson(res, 200, { ok: true, importedCount: imported.length, imported });
      return;
    }

    if (url.pathname === '/api/ringcentral/webhook/subscribe' && req.method === 'POST') {
      const body = await readRequestBody(req);
      const deliveryAddress = body.deliveryAddress || process.env.RC_WEBHOOK_URL;
      if (!deliveryAddress) {
        sendJson(res, 400, { ok: false, error: 'Provide deliveryAddress or set RC_WEBHOOK_URL in .env.' });
        return;
      }
      const subscription = await ringCentralRequest('POST', '/restapi/v1.0/subscription', {
        eventFilters: ['/restapi/v1.0/account/~/extension/~/message-store/instant?type=SMS'],
        deliveryMode: {
          transportType: 'WebHook',
          address: deliveryAddress,
        },
      });
      sendJson(res, 200, { ok: true, subscription });
      return;
    }

    serveStatic(req, res, url.pathname);
  } catch (error) {
    sendJson(res, error.status || 500, {
      ok: false,
      error: error.message || 'Unexpected server error',
      squareErrors: error.squareErrors,
      ringCentral: error.ringCentral,
    });
  }
});

server.listen(PORT, () => {
  console.log(`AbleCare dashboard running at http://localhost:${PORT}`);
  startReminderScheduler();
});

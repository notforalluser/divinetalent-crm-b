require('dotenv').config();
const express = require('express');
const http = require('http');
const cors = require('cors');
const XLSX = require('xlsx');
const { Server } = require('socket.io');
const mongoose = require('mongoose');

const PORT = process.env.PORT || 5000;
const MONGO_URI = process.env.MONGO_URI || '';

// Same ceiling the frontend already enforces (SettingsContext.jsx) — kept
// here too so a direct API call (or a bug on the client) can never push a
// score above it.
const MAX_ATS_SCORE = 94;

function clampScore(n, fallback) {
  const num = Number(n);
  const value = Number.isFinite(num) ? num : fallback;
  return Math.min(MAX_ATS_SCORE, Math.max(0, value));
}

const EMPTY_PROFILE = {
  jobRoles: [],
  experienceYears: '',
  location: '',
  name: '',
  mobile: '',
  email: '',
  skills: [],
};

function defaultAtsSettings() {
  return {
    atsScoreMin: 55,
    atsScoreMax: 68,
    atsEligibilityThreshold: 75,
    atsGeneralProfile: { ...EMPTY_PROFILE },
    atsFilenameOverrides: [],
  };
}

function sanitizeAtsSettings(raw) {
  const p = raw || {};
  return {
    atsScoreMin: clampScore(p.atsScoreMin, 55),
    atsScoreMax: clampScore(p.atsScoreMax, 68),
    atsEligibilityThreshold: clampScore(p.atsEligibilityThreshold, 65),
    atsGeneralProfile: { ...EMPTY_PROFILE, ...(p.atsGeneralProfile || {}) },
    atsFilenameOverrides: Array.isArray(p.atsFilenameOverrides)
      ? p.atsFilenameOverrides
          .map((o) => ({
            id: o.id || `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
            filename: (o.filename || '').trim(),
            minScore: clampScore(o.minScore, 55),
            profile: { ...EMPTY_PROFILE, ...(o.profile || {}) },
          }))
          .filter((o) => o.filename)
      : [],
  };
}

// ---- In-memory state (always the source of truth the app reads/writes to) ----
let atsSettings = defaultAtsSettings();

// ---- Optional MongoDB setup (ATS settings only) ----
let AtsSettingsModel = null;
let useMongo = false;

async function setupMongo() {
  if (!MONGO_URI) {
    console.log('MONGO_URI not set — running with in-memory storage only.');
    return;
  }
  try {
    await mongoose.connect(MONGO_URI);
    const schema = new mongoose.Schema(
      {
        _id: { type: String, default: 'singleton' },
        atsScoreMin: { type: Number, default: 55 },
        atsScoreMax: { type: Number, default: 68 },
        atsEligibilityThreshold: { type: Number, default: 75 },
        atsGeneralProfile: { type: mongoose.Schema.Types.Mixed, default: () => ({ ...EMPTY_PROFILE }) },
        atsFilenameOverrides: { type: mongoose.Schema.Types.Mixed, default: () => [] },
      },
      { versionKey: false }
    );
    AtsSettingsModel = mongoose.model('AtsSettings', schema);

    let doc = await AtsSettingsModel.findById('singleton');
    if (!doc) {
      doc = await AtsSettingsModel.create({ _id: 'singleton', ...atsSettings });
    }
    atsSettings = sanitizeAtsSettings(doc.toObject());
    useMongo = true;
    console.log('Connected to MongoDB — ATS settings will persist across restarts.');
  } catch (err) {
    console.error('MongoDB connection failed, falling back to in-memory only:', err.message);
  }
}

async function persist() {
  if (!useMongo) return;
  try {
    await AtsSettingsModel.updateOne({ _id: 'singleton' }, atsSettings, { upsert: true });
  } catch (err) {
    console.error('Failed to persist ATS settings to MongoDB:', err.message);
  }
}

// ============================================================================
// CRM dataset (crm-data.xlsx)
//
// The RAW FILE now lives on Hostinger (persistent PHP storage), not on
// Render's disk (which is wiped on every restart/redeploy). Node's job is
// unchanged otherwise: pull the bytes, parse them ONCE with the "xlsx"
// package, cache the JSON in memory, and broadcast it over Socket.IO. The
// browser never parses anything — it only ever receives this cached JSON.
//
// Flow:
//  - On boot, Node pulls the current file from Hostinger to warm the cache.
//  - Hostinger's crm-upload.php / crm-reset.php ping POST /api/crm-data/sync
//    after every change, which makes Node re-pull + re-parse + rebroadcast.
//  - GET /api/crm-data always just serves whatever is currently cached —
//    fast, no network hop, no parsing — same as before.
// ============================================================================

const CRM_SHEETS = ['Candidates', 'Jobs', 'Recruiters', 'Interviews', 'TechnicalHelp', 'Activity', 'MarketingActivity'];
const EMPTY_CRM_SHEETS = Object.fromEntries(CRM_SHEETS.map((s) => [s, []]));

// Hostinger's crm-data-file.php — the endpoint Node fetches the raw .xlsx from.
const HOSTINGER_FILE_URL = process.env.HOSTINGER_FILE_URL || 'https://YOUR-DOMAIN.com/api/crm-data-file.php';
// Must match SYNC_SECRET in Hostinger's config.php exactly.
const CRM_SYNC_SECRET = process.env.CRM_SYNC_SECRET || '';

let crmData = { ...EMPTY_CRM_SHEETS, fetchedAt: null, source: null };

/**
 * Pulls the current .xlsx from Hostinger, parses it, and updates the cache.
 * IMPORTANT: on failure, the existing cache is left untouched — a Hostinger
 * hiccup should never wipe a dashboard that was working a second ago.
 * Returns true on success, false on failure.
 */
async function pullCrmDataFromHostinger() {
  if (!CRM_SYNC_SECRET) {
    console.error('CRM_SYNC_SECRET is not set — refusing to pull CRM data.');
    return false;
  }
  try {
    const res = await fetch(HOSTINGER_FILE_URL, {
      headers: { 'X-Crm-Secret': CRM_SYNC_SECRET },
    });
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      throw new Error(`Hostinger returned ${res.status}: ${body.slice(0, 200)}`);
    }

    const sourceType = res.headers.get('x-source-type') || 'uploaded';
    const sourceName = decodeURIComponent(res.headers.get('x-source-name') || 'crm-data.xlsx');
    const sourceUploadedAt = res.headers.get('x-source-uploaded-at')
      ? decodeURIComponent(res.headers.get('x-source-uploaded-at'))
      : undefined;

    const buffer = Buffer.from(await res.arrayBuffer());
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });

    const parsed = {};
    for (const sheetName of CRM_SHEETS) {
      const sheet = workbook.Sheets[sheetName];
      parsed[sheetName] = sheet ? XLSX.utils.sheet_to_json(sheet, { defval: '' }) : [];
    }

    crmData = {
      ...parsed,
      fetchedAt: new Date().toISOString(),
      source: sourceType === 'bundled' ? { type: 'bundled', name: sourceName } : { type: 'uploaded', name: sourceName, uploadedAt: sourceUploadedAt },
    };

    console.log(`CRM dataset refreshed from Hostinger (${sourceType}: ${sourceName}).`);
    return true;
  } catch (err) {
    console.error('Failed to pull CRM dataset from Hostinger — keeping last known cache:', err.message);
    return false;
  }
}

// ---- Express + Socket.IO setup ----
const app = express();
app.use(cors());
app.use(express.json());

const server = http.createServer(app);
const io = new Server(server, {
  cors: { origin: '*' }, // allow phone/laptop/any origin to connect for testing
});

function broadcast() {
  io.emit('atsSettingsUpdate', atsSettings); // pushed to every connected device instantly
}

function broadcastCrmData() {
  io.emit('crmDataUpdate', crmData); // pushed to every connected device instantly
}

// GET current settings.
app.get('/api/ats-settings', (req, res) => {
  res.json(atsSettings);
});

app.put('/api/ats-settings/field', async (req, res) => {
  const { key, value } = req.body || {};
  const allowed = ['atsScoreMin', 'atsScoreMax', 'atsEligibilityThreshold'];
  if (!allowed.includes(key)) {
    return res.status(400).json({ error: 'Invalid field' });
  }
  atsSettings = { ...atsSettings, [key]: clampScore(value, atsSettings[key]) };
  await persist();
  broadcast();
  res.json(atsSettings);
});

app.put('/api/ats-settings/general-profile', async (req, res) => {
  const { partial } = req.body || {};
  atsSettings = {
    ...atsSettings,
    atsGeneralProfile: { ...atsSettings.atsGeneralProfile, ...(partial || {}) },
  };
  await persist();
  broadcast();
  res.json(atsSettings);
});

app.post('/api/ats-settings/overrides', async (req, res) => {
  const { id, filename, minScore, profile } = req.body || {};
  const trimmed = (filename || '').trim();
  if (!trimmed) {
    return res.status(400).json({ error: 'filename is required' });
  }
  const clampedMinScore = clampScore(minScore, 55);
  const others = (atsSettings.atsFilenameOverrides || []).filter(
    (o) => o.id !== id && o.filename.trim().toLowerCase() !== trimmed.toLowerCase()
  );
  const newId = id || `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
  atsSettings = {
    ...atsSettings,
    atsFilenameOverrides: [
      ...others,
      {
        id: newId,
        filename: trimmed,
        minScore: clampedMinScore,
        profile: { ...EMPTY_PROFILE, ...(profile || {}) },
      },
    ],
  };
  await persist();
  broadcast();
  res.json(atsSettings);
});

app.delete('/api/ats-settings/overrides/:id', async (req, res) => {
  const { id } = req.params;
  atsSettings = {
    ...atsSettings,
    atsFilenameOverrides: (atsSettings.atsFilenameOverrides || []).filter((o) => o.id !== id),
  };
  await persist();
  broadcast();
  res.json(atsSettings);
});

app.post('/api/ats-settings/reset', async (req, res) => {
  atsSettings = defaultAtsSettings();
  await persist();
  broadcast();
  res.json(atsSettings);
});

app.post('/api/ats-settings/replace', async (req, res) => {
  const { atsSettings: incoming } = req.body || {};
  atsSettings = sanitizeAtsSettings(incoming);
  await persist();
  broadcast();
  res.json(atsSettings);
});

// ---- CRM dataset routes ----

// GET the currently cached, already-parsed dataset — this is the only route
// the frontend's normal read path ever calls.
app.get('/api/crm-data', (req, res) => {
  res.json(crmData);
});

// Webhook: Hostinger's crm-upload.php / crm-reset.php call this after every
// change. Re-pulls the file, re-parses it, and broadcasts the result.
app.post('/api/crm-data/sync', async (req, res) => {
  const given = req.headers['x-crm-secret'] || '';
  if (given !== CRM_SYNC_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  const ok = await pullCrmDataFromHostinger();
  if (ok) broadcastCrmData();
  res.status(ok ? 200 : 502).json({ ok, source: crmData.source });
});

// Manual/admin re-sync — handy for debugging without touching Hostinger.
app.get('/api/crm-data/resync', async (req, res) => {
  const given = req.headers['x-crm-secret'] || '';
  if (given !== CRM_SYNC_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  const ok = await pullCrmDataFromHostinger();
  if (ok) broadcastCrmData();
  res.status(ok ? 200 : 502).json({ ok, source: crmData.source });
});

// Socket.IO: realtime channel
io.on('connection', (socket) => {
  console.log('Client connected:', socket.id);
  socket.emit('atsSettingsUpdate', atsSettings);
  socket.emit('crmDataUpdate', crmData);

  socket.on('disconnect', () => {
    console.log('Client disconnected:', socket.id);
  });
});

async function start() {
  await setupMongo();
  await pullCrmDataFromHostinger(); // warm the cache on boot — if this fails, we start empty and wait for the next upload's webhook ping (or hit /api/crm-data/resync manually)
  server.listen(PORT, '0.0.0.0', () => {
    console.log(`Backend running on http://0.0.0.0:${PORT}`);
  });
}

start();

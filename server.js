/**
 * server.js — realtime ATS settings backend
 *
 * - Express REST API for reading/writing the ATS score simulation settings
 *   (atsScoreMin, atsScoreMax, atsEligibilityThreshold, atsGeneralProfile,
 *   atsFilenameOverrides).
 * - Socket.IO pushes every change to all connected devices instantly — set
 *   a score on your phone, it shows up on your laptop with no refresh.
 * - MongoDB persistence is OPTIONAL: set MONGO_URI in .env to enable it.
 *   Without it, settings just live in memory (reset when the server restarts).
 *
 * Run:
 *   npm install
 *   node server.js
 *
 * Env vars (optional, put in a .env file next to this):
 *   PORT=5000
 *   MONGO_URI=mongodb://127.0.0.1:27017/crm-ats-settings
 *
 * On the frontend, point VITE_ATS_SERVER_URL at wherever this is running,
 * e.g. http://192.168.1.23:5000 so your phone and laptop both reach it.
 */

require('dotenv').config();
const express = require('express');
const http = require('http');
const cors = require('cors');
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

// ---- Optional MongoDB setup ----
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

    // load existing doc, or create it
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

// GET current settings — used on first load, and as a fallback if a client
// ever needs to re-sync outside of the socket connection.
app.get('/api/ats-settings', (req, res) => {
  res.json(atsSettings);
});

// PUT a single top-level field: atsScoreMin | atsScoreMax | atsEligibilityThreshold
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

// PUT the general (no filename override) simulated profile
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

// POST create or update a filename override (e.g. "resume.pdf" -> 79)
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

// DELETE a filename override
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

// POST reset everything back to defaults
app.post('/api/ats-settings/reset', async (req, res) => {
  atsSettings = defaultAtsSettings();
  await persist();
  broadcast();
  res.json(atsSettings);
});

// POST wholesale replace — used by "Import settings" on the Settings page
app.post('/api/ats-settings/replace', async (req, res) => {
  const { atsSettings: incoming } = req.body || {};
  atsSettings = sanitizeAtsSettings(incoming);
  await persist();
  broadcast();
  res.json(atsSettings);
});

// Socket.IO: realtime channel
io.on('connection', (socket) => {
  console.log('Client connected:', socket.id);

  // send current state immediately on connect
  socket.emit('atsSettingsUpdate', atsSettings);

  socket.on('disconnect', () => {
    console.log('Client disconnected:', socket.id);
  });
});

setupMongo().finally(() => {
  server.listen(PORT, '0.0.0.0', () => {
    console.log(`ATS settings server running on http://0.0.0.0:${PORT}`);
    console.log('On your phone, use your laptop\'s LAN IP, e.g. http://192.168.x.x:' + PORT);
  });
});

// bot.js - WhatsApp & Email Job Application Bot (Baileys + MongoDB)
// Kolom spreadsheet (Sheet1):
//   A: COMPANY | B: POSITION | C: CONTACT | D: REQUIREMENTS
//   E: STATUS  | F: FILE_PATH | G: SUBJECT | H: BACKUP_ONLY
//
// BACKUP_ONLY:
//   "1"  → backup only (hanya kirim file bersih ke galehforjob@gmail.com)
//   kosong / selain "1" → default (kirim ke target + backup)

const fs = require('fs');
const path = require('path');
const chalk = require('chalk');
const ora = require('ora');
const mongoose = require('mongoose');
const pino = require('pino');
const qrcode = require('qrcode-terminal');
const nodemailer = require('nodemailer');
const dotenv = require('dotenv');
const { google } = require('googleapis');
const { execFile } = require('child_process');
const { promisify } = require('util');
const dns = require('dns');

let PDFDocument = null, StandardFonts = null, rgbFn = null, degreesFn = null;
try {
    const pdfLib = require('pdf-lib');
    PDFDocument = pdfLib.PDFDocument;
    StandardFonts = pdfLib.StandardFonts;
    rgbFn = pdfLib.rgb;
    degreesFn = pdfLib.degrees;
} catch {
    console.log('⚠️ pdf-lib tidak terinstall. Jalankan: npm install pdf-lib');
}

const execFileAsync = promisify(execFile);
try { dns.setDefaultResultOrder('ipv4first'); } catch {}
dotenv.config();

// ============================================================
// FILTER LOG NOISE
// ============================================================
(function installConsoleNoiseFilter() {
    const NOISE = ['Closing session', 'Opening session', 'Session already closed',
        'SessionEntry', '"class":"baileys"', 'failed to decrypt message',
        'sent retry receipt', 'Bad MAC', 'Closing open session'];
    const isNoiseText = (s) => {
        if (typeof s !== 'string' || !s.length) return false;
        for (const p of NOISE) if (s.includes(p)) return true;
        return false;
    };
    const isNoiseArgs = (args) => {
        if (!args || !args.length) return false;
        for (const a of args) {
            if (typeof a === 'string' && isNoiseText(a)) return true;
            if (a instanceof Error && isNoiseText(a.stack || a.message || '')) return true;
        }
        return false;
    };
    const wrap = (fn) => (...args) => { if (isNoiseArgs(args)) return; return fn.apply(console, args); };
    console.log = wrap(console.log); console.info = wrap(console.info);
    console.debug = wrap(console.debug); console.warn = wrap(console.warn);
    console.error = wrap(console.error);
    const wrapStream = (stream) => {
        const orig = stream.write.bind(stream);
        stream.write = function (chunk, encoding, cb) {
            let s = '';
            try {
                if (typeof chunk === 'string') s = chunk;
                else if (Buffer.isBuffer(chunk)) s = chunk.toString('utf8');
                else s = String(chunk);
            } catch {}
            if (isNoiseText(s)) {
                if (typeof encoding === 'function') { encoding(); return true; }
                if (typeof cb === 'function') { cb(); return true; }
                return true;
            }
            return orig(chunk, encoding, cb);
        };
    };
    try { wrapStream(process.stdout); } catch {}
    try { wrapStream(process.stderr); } catch {}
})();

const DEBUG_WA = process.env.DEBUG_WA === '1';

// ============================================================
// GOOGLE SHEETS
// ============================================================
const CREDENTIALS_PATH = path.join(__dirname, 'credentials.json');
if (!fs.existsSync(CREDENTIALS_PATH)) { console.error(chalk.red('❌ credentials.json hilang!')); process.exit(1); }
const credentials = require(CREDENTIALS_PATH);
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// ============================================================
// IMPORT modules
// ============================================================
const { generateApplicationLetter, CONFIG: AI_CONFIG } = require('./lamaran-ai.js');
const { mergeAndStamp, sanitizeName, FINAL_DIR, IJAZAH_PATH, STAMP: TM_STAMP } = require('./test-merge.js');

// ============================================================
// KONFIGURASI BOT
// ============================================================
const BOT_CONFIG = {
    APPLICANT_NAME: 'Galeh Riski Prasetio',
    APPLICANT_PHONE: '0851-9419-4479',
    APPLICANT_EMAIL: 'galehforjob@gmail.com',
    CV_PATH: path.join(__dirname, 'cv', 'CV-Galeh.pdf'),

    // ---- SHEET ----
    SHEET_NAME: 'Sheet1',

    DELAY_BETWEEN_SENDS: 8000,
    DELAY_BETWEEN_TARGETS_WA_MIN: 60000,
    DELAY_BETWEEN_TARGETS_WA_MAX: 100000,
    DELAY_BETWEEN_TARGETS_EMAIL_MIN: 1000,
    DELAY_BETWEEN_TARGETS_EMAIL_MAX: 3000,
    MAX_RETRIES: 3,
    WA_SYNC_DELAY_MS: 8000,
    MIN_PDF_BYTES: 1000,
    ACK_LOG_TIMEOUT_MS: 5000,

    MERGE_WITH_IJAZAH: true,
    FINAL_DIR: FINAL_DIR,
    IJAZAH_PATH: IJAZAH_PATH,

    FLATTEN_LETTER: true,
    FLATTEN_CV: false,
    FLATTEN_IJAZAH: true,
    FLATTEN_DPI: 300,
    FLATTEN_FORMAT: 'jpeg',
    FLATTEN_JPEG_QUALITY: 92,

    BLOCK_PRINT: true,
    ALLOW_COPY: true,
    ALLOW_ANNOTATION: true,
    GS_OWNER_PASSWORD: process.env.GS_OWNER_PASSWORD || '',

    WATERMARK_ENABLED: true,
    WATERMARK_MODE: 'company',
    WATERMARK_OPACITY: 0.10,
    WATERMARK_FONT_SIZE: 32,
    WATERMARK_DIAGONAL: true,
    WATERMARK_ANGLE: 45,
    WATERMARK_COLOR: [0.55, 0.55, 0.55],

    BACKUP_ENABLED: true,
    BACKUP_EMAIL: 'galehforjob@gmail.com',
    BACKUP_SUBJECT_PREFIX: 'Backup',

    DELETE_AFTER_SENT: true,
    CACHE_RASTERIZED_IJAZAH: true,
    REQUIRE_GHOSTSCRIPT: false,
};

const TEMP_DIR = path.join(__dirname, 'temp');

function buildDefaultSubject(position) { return `Galeh Riski Prasetio - ${position}`; }

// ============================================================
// MONGODB
// ============================================================
const MONGODB_URI = process.env.MONGODB || process.env.MONGODB_URI;
const SESSION_ID = process.env.SESSION_ID || 'whatsapp_session';
const PHONE_NUMBER = process.env.WA_PHONE_NUMBER || '';
if (!MONGODB_URI) { console.error(chalk.red('❌ MONGODB tidak diatur')); process.exit(1); }

const { model, Schema } = mongoose;
const sessionSchema = new Schema({
    sessionId: { type: String, required: true, unique: true },
    session: { type: String }, state: { type: Object },
    updatedAt: { type: Date, default: Date.now }
}, { strict: false, collection: 'sessionschemas' });
const SessionModel = model('WhatsAppSession', sessionSchema);

// ============================================================
// CEK GHOSTSCRIPT
// ============================================================
let HAS_GHOSTSCRIPT = false;
async function checkGhostscript() {
    try {
        const { stdout } = await execFileAsync('gs', ['--version'], { timeout: 5000 });
        HAS_GHOSTSCRIPT = true;
        return (stdout || '').trim();
    } catch { HAS_GHOSTSCRIPT = false; return null; }
}

// ============================================================
// HELPER
// ============================================================
function delay(ms) { return new Promise(r => setTimeout(r, ms)); }
function jitterDelay(baseMs) { const v = baseMs * 0.3; return Math.round(baseMs - v + Math.random() * (v * 2)); }
function isEmail(contact) { return contact && contact.includes('@'); }
function safeUnlink(p) { try { if (p && fs.existsSync(p)) { fs.unlinkSync(p); return true; } } catch {} return false; }
function isUsableFile(filePath, minBytes = BOT_CONFIG.MIN_PDF_BYTES) {
    try { return !!filePath && fs.existsSync(filePath) && fs.statSync(filePath).size > minBytes; } catch { return false; }
}
function getGreetingByTime() {
    const h = parseInt(new Date().toLocaleString('en-US', { timeZone: 'Asia/Jakarta', hour: '2-digit', hour12: false }), 10);
    if (h >= 5 && h < 11) return 'Selamat pagi';
    if (h >= 11 && h < 15) return 'Selamat siang';
    if (h >= 15 && h < 18) return 'Selamat sore';
    return 'Selamat malam';
}
function buildOutputFileName(position) {
    const s1 = sanitizeName(BOT_CONFIG.APPLICANT_NAME || 'Pelamar');
    const s2 = sanitizeName(position || 'Posisi');
    return `${s1}_${s2}.pdf`;
}
function buildWatermarkText(target, trackingId) {
    if (BOT_CONFIG.WATERMARK_MODE === 'company') return target.company;
    const parts = [BOT_CONFIG.APPLICANT_NAME, target.company, target.contact,
        new Date().toISOString().replace('T', ' ').substring(0, 19), trackingId || ''].filter(Boolean);
    return parts.join(' | ');
}

// ============================================================
// HTML EMAIL (profesional)
// ============================================================
function buildEmailHtml(company, position) {
    const name = BOT_CONFIG.APPLICANT_NAME;
    const phone = BOT_CONFIG.APPLICANT_PHONE;
    const email = BOT_CONFIG.APPLICANT_EMAIL;
    return `<!DOCTYPE html>
<html lang="id"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>Lamaran Kerja</title></head>
<body style="margin:0;padding:0;background-color:#f3f4f6;font-family:Arial,Helvetica,sans-serif;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f3f4f6;padding:24px 0;">
<tr><td align="center">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.06);">
<tr><td style="background-color:#1e3a8a;padding:24px 32px;">
<h1 style="margin:0;color:#ffffff;font-size:20px;font-weight:700;letter-spacing:0.3px;">${name}</h1>
<p style="margin:4px 0 0;color:#bfdbfe;font-size:12px;letter-spacing:1.5px;text-transform:uppercase;font-weight:500;">Job Application</p>
</td></tr>
<tr><td style="padding:32px;">
<p style="margin:0 0 16px;color:#1f2937;font-size:14px;line-height:1.65;">Kepada Yth.<br><strong style="color:#111827;">HRD ${company}</strong></p>
<p style="margin:0 0 16px;color:#1f2937;font-size:14px;line-height:1.65;">Dengan hormat,</p>
<p style="margin:0 0 16px;color:#1f2937;font-size:14px;line-height:1.65;">Perkenalkan, saya <strong>${name}</strong>. Sehubungan dengan informasi lowongan posisi <strong style="color:#1e3a8a;">${position}</strong> di <strong>${company}</strong>, saya bermaksud mengajukan lamaran untuk posisi tersebut.</p>
<p style="margin:0 0 16px;color:#1f2937;font-size:14px;line-height:1.65;">Bersama email ini, saya lampirkan CV dan Surat Lamaran Kerja sebagai bahan pertimbangan awal. Besar harapan saya untuk dapat diberikan kesempatan wawancara.</p>
<p style="margin:0 0 16px;color:#1f2937;font-size:14px;line-height:1.65;">Terima kasih atas waktu dan perhatian Bapak/Ibu.</p>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-top:32px;padding-top:20px;border-top:1px solid #e5e7eb;"><tr><td>
<p style="margin:0 0 8px;color:#1f2937;font-size:14px;">Hormat saya,</p>
<p style="margin:0;color:#1e3a8a;font-size:15px;font-weight:700;">${name}</p>
<p style="margin:6px 0 0;color:#6b7280;font-size:13px;">📱 ${phone}</p>
<p style="margin:2px 0 0;color:#6b7280;font-size:13px;">✉️ ${email}</p>
</td></tr></table>
</td></tr>
<tr><td style="background-color:#f9fafb;padding:16px 32px;text-align:center;border-top:1px solid #e5e7eb;">
<p style="margin:0;color:#9ca3af;font-size:11px;">Email ini dikirim secara otomatis sebagai bagian dari proses lamaran kerja.</p>
</td></tr>
</table></td></tr></table></body></html>`;
}

function buildEmailText(company, position) {
    return `Yth. HRD ${company},

Dengan hormat,

Perkenalkan, saya ${BOT_CONFIG.APPLICANT_NAME}. Sehubungan dengan informasi lowongan posisi ${position} di ${company}, saya bermaksud mengajukan lamaran untuk posisi tersebut.

Bersama email ini, saya lampirkan CV dan Surat Lamaran Kerja sebagai bahan pertimbangan awal. Besar harapan saya untuk dapat diberikan kesempatan wawancara.

Terima kasih atas waktu dan perhatian Bapak/Ibu.

Hormat saya,
${BOT_CONFIG.APPLICANT_NAME}
${BOT_CONFIG.APPLICANT_PHONE}
${BOT_CONFIG.APPLICANT_EMAIL}`;
}

// ============================================================
// MONGODB
// ============================================================
async function connectMongo() {
    if (mongoose.connection.readyState === 1) return;
    try { await mongoose.connect(MONGODB_URI); console.log(chalk.green('✅ MongoDB')); }
    catch (e) { console.error(chalk.red('❌ MongoDB:'), e.message); process.exit(1); }
}
let _baileysModule = null;
async function getBaileys() { if (!_baileysModule) _baileysModule = await import('@whiskeysockets/baileys'); return _baileysModule; }

function looksLikeByteArrayObject(val) {
    if (!val || typeof val !== 'object' || Array.isArray(val) || Buffer.isBuffer(val)) return false;
    const keys = Object.keys(val); if (!keys.length) return false;
    for (let i = 0; i < keys.length; i++) {
        if (keys[i] !== String(i)) return false;
        const v = val[keys[i]];
        if (typeof v !== 'number' || !Number.isInteger(v) || v < 0 || v > 255) return false;
    } return true;
}
function deepRepairBuffers(node, rc = { n: 0 }) {
    if (node === null || typeof node !== 'object' || Buffer.isBuffer(node)) return node;
    if (Array.isArray(node)) { for (let i = 0; i < node.length; i++) node[i] = deepRepairBuffers(node[i], rc); return node; }
    if (node._bsontype === 'Binary' && node.buffer) { rc.n++; return Buffer.isBuffer(node.buffer) ? node.buffer : Buffer.from(node.buffer); }
    if (node.type === 'Buffer' && node.data !== undefined) {
        const val = node.data; rc.n++;
        if (typeof val === 'string') return Buffer.from(val, 'base64');
        if (Array.isArray(val)) return Buffer.from(val);
        if (looksLikeByteArrayObject(val)) return Buffer.from(Object.values(val));
        return Buffer.alloc(0);
    }
    if (looksLikeByteArrayObject(node)) { rc.n++; return Buffer.from(Object.values(node)); }
    for (const k of Object.keys(node)) node[k] = deepRepairBuffers(node[k], rc);
    return node;
}
async function loadSessionFromMongo() {
    try {
        const doc = await SessionModel.findOne({ sessionId: SESSION_ID });
        if (!doc) return null;
        console.log(chalk.green('✅ Session ditemukan'));
        if (doc.session) {
            try {
                const { BufferJSON } = await getBaileys();
                const parsed = JSON.parse(doc.session, BufferJSON.reviver);
                if (parsed.creds && parsed.keys) {
                    const rc = { n: 0 };
                    const repaired = deepRepairBuffers(parsed, rc);
                    if (rc.n > 0) await saveSessionToMongo(repaired.creds, repaired.keys);
                    return repaired;
                }
            } catch (e) { console.error(chalk.red('❌ Parse session:'), e.message); }
        }
        if (doc.state && doc.state.creds && doc.state.keys) {
            const rc = { n: 0 };
            return deepRepairBuffers(doc.state.toObject ? doc.state.toObject() : doc.state, rc);
        }
        return null;
    } catch (e) { console.error(chalk.red('❌ Load session:'), e.message); return null; }
}
function makeMongoKeyStore(keysData, onChanged) {
    return {
        get: async (type, ids) => { const d = {}; for (const id of ids) { const v = keysData?.[type]?.[id]; if (v) d[id] = v; } return d; },
        set: async (data) => {
            for (const t in data) { keysData[t] = keysData[t] || {};
                for (const id in data[t]) { const v = data[t][id]; if (v) keysData[t][id] = v; else delete keysData[t][id]; } }
            if (onChanged) onChanged();
        }
    };
}
async function saveSessionToMongo(creds, keysData) {
    try {
        if (!creds || !keysData) return;
        const state = { creds, keys: keysData };
        const { BufferJSON } = await getBaileys();
        await SessionModel.findOneAndUpdate({ sessionId: SESSION_ID },
            { sessionId: SESSION_ID, state, session: JSON.stringify(state, BufferJSON.replacer), updatedAt: new Date() },
            { upsert: true });
        console.log(chalk.gray('💾 Session disimpan'));
    } catch (e) { console.error(chalk.yellow('⚠️ Gagal simpan session:'), e.message); }
}
function makeDebouncedSaver(getCreds, keysData, delayMs = 2000) {
    let timer = null;
    return () => {
        if (timer) clearTimeout(timer);
        timer = setTimeout(() => { timer = null; const c = getCreds(); if (c) saveSessionToMongo(c, keysData).catch(() => {}); }, delayMs);
    };
}
async function deleteSessionFromMongo() {
    try { await SessionModel.deleteOne({ sessionId: SESSION_ID }); console.log(chalk.gray('🗑️ Session dihapus')); }
    catch (e) { console.error(chalk.yellow('⚠️'), e.message); }
}

// ============================================================
// WA STATE
// ============================================================
function createWaState() { return { sock: null, isOpen: false, pendingAcks: new Map(), recentMessages: new Map(), reconnectLog: [], sentTargets: new Set() }; }
function makeGetMessage(s) { return async (key) => { const e = s.recentMessages.get(key.id); return e ? e.message : undefined; }; }
function makeSimpleCacheStore() { const m = new Map(); return { get: (k) => m.get(k), set: (k, v) => { m.set(k, v); return true; }, del: (k) => m.delete(k), flushAll: () => m.clear() }; }
function attachAckTracking(sock, s) {
    sock.ev.on('messages.update', (ups) => {
        for (const { key, update } of ups) {
            if (key?.id && typeof update?.status === 'number' && s.pendingAcks.has(key.id)) {
                const cb = s.pendingAcks.get(key.id); s.pendingAcks.delete(key.id); cb(update.status);
            }
        }
    });
    sock.ev.on('message-receipt.update', (ups) => {
        for (const u of ups) { const id = u?.key?.id; if (id && s.pendingAcks.has(id)) { const cb = s.pendingAcks.get(id); s.pendingAcks.delete(id); cb(3); } }
    });
}
function waitForOpenConnection(s, timeoutMs = 90000) {
    if (s.isOpen && s.sock) return Promise.resolve(true);
    return new Promise((resolve) => {
        const start = Date.now();
        const iv = setInterval(() => {
            if (s.isOpen && s.sock) { clearInterval(iv); resolve(true); }
            else if (Date.now() - start > timeoutMs) { clearInterval(iv); resolve(false); }
        }, 1000);
    });
}
async function sendMessageConfirmed(s, jid, content, opts = {}, ackTimeout = BOT_CONFIG.ACK_LOG_TIMEOUT_MS) {
    const sock = s.sock; if (!sock) return { success: false, reason: 'no-socket' };
    let r; try { r = await sock.sendMessage(jid, content, opts); } catch (e) { return { success: false, reason: 'send-error', error: e.message }; }
    const id = r?.key?.id; if (!id) return { success: false, reason: 'no-id' };
    s.recentMessages.set(id, { key: r.key, message: r.message || content });
    if (s.recentMessages.size > 200) s.recentMessages.delete(s.recentMessages.keys().next().value);
    let ackStatus = null;
    try {
        ackStatus = await new Promise((resolve) => {
            const t = setTimeout(() => { s.pendingAcks.delete(id); resolve(null); }, ackTimeout);
            s.pendingAcks.set(id, (st) => { clearTimeout(t); resolve(st); });
        });
    } catch {}
    return { success: true, id, ackStatus };
}
function noteReconnectAndMaybeCooldown(s) {
    const now = Date.now(); s.reconnectLog.push(now);
    s.reconnectLog = s.reconnectLog.filter(t => now - t < 10 * 60 * 1000);
    if (s.reconnectLog.length >= 4) return 10 * 60 * 1000;
    return 0;
}

// ============================================================
// CONNECT WHATSAPP
// ============================================================
async function connectToWhatsApp(waState) {
    await connectMongo();
    let state = await loadSessionFromMongo();
    if (state && state.creds && state.keys) console.log(chalk.green('✅ Session valid'));
    else { console.log(chalk.yellow('ℹ️ Minta pairing/QR...')); state = { creds: {}, keys: {} }; }

    const { default: makeWASocket, DisconnectReason, fetchLatestBaileysVersion, initAuthCreds, makeCacheableSignalKeyStore } = await getBaileys();
    if (!state.creds || !state.creds.noiseKey) state = { creds: initAuthCreds(), keys: state.keys || {} };
    const { version } = await fetchLatestBaileysVersion();
    console.log(chalk.gray(`ℹ️ Baileys v${version.join('.')}`));

    const logger = pino({ level: DEBUG_WA ? 'debug' : 'silent' });
    const keysData = state.keys || {};
    let latestCreds = state.creds;
    const debounced = makeDebouncedSaver(() => latestCreds, keysData);
    let pendingSave = Promise.resolve();

    const sock = makeWASocket({
        logger,
        auth: { creds: state.creds, keys: makeCacheableSignalKeyStore(makeMongoKeyStore(keysData, debounced), logger) },
        version, browser: ['Ubuntu', 'Chrome', '20.0.00'],
        connectTimeoutMs: 60000, qrTimeout: 60000, markOnlineOnConnect: true,
        emitOwnEvents: true, fireInitQueries: true, generateHighQualityLinkPreview: true,
        syncFullHistory: false, retryRequestDelayMs: 1000,
        getMessage: makeGetMessage(waState), msgRetryCounterCache: makeSimpleCacheStore(),
    });
    waState.sock = sock; waState.isOpen = false; attachAckTracking(sock, waState);
    let attempts = 0, pairingReq = false;

    sock.ev.on('creds.update', async () => {
        if (waState.sock !== sock) return;
        if (sock.authState && sock.authState.creds) {
            latestCreds = sock.authState.creds;
            pendingSave = saveSessionToMongo(sock.authState.creds, keysData);
            await pendingSave;
        }
    });

    sock.ev.on('connection.update', async (u) => {
        if (waState.sock !== sock) return;
        const { qr } = u;
        if (qr && PHONE_NUMBER && !pairingReq && !sock.authState.creds.registered) {
            pairingReq = true;
            try { const code = await sock.requestPairingCode(PHONE_NUMBER.trim()); console.log(chalk.bold.yellow(`\n   Pairing: ${code}\n`)); }
            catch (e) { console.log(chalk.yellow(`⚠️ Pairing: ${e.message}`)); pairingReq = false; }
        } else if (qr && !PHONE_NUMBER) { qrcode.generate(qr, { small: true }); }
    });

    sock.ev.on('connection.update', async (u) => {
        if (waState.sock !== sock) return;
        const { connection, lastDisconnect } = u;
        if (connection === 'close') {
            waState.isOpen = false;
            const sc = lastDisconnect?.error?.output?.statusCode;
            if (sc === DisconnectReason.loggedOut) { await deleteSessionFromMongo(); process.exit(1); }
            else if (sc === DisconnectReason.badSession) {
                await deleteSessionFromMongo();
                const cd = noteReconnectAndMaybeCooldown(waState);
                await pendingSave.catch(() => {});
                setTimeout(() => connectToWhatsApp(waState).catch(console.error), cd || 3000);
            } else {
                attempts++;
                if (attempts > 10) process.exit(1);
                const wait = noteReconnectAndMaybeCooldown(waState) || 3000;
                console.log(chalk.yellow(`🔄 Reconnect (${attempts}/10) dalam ${Math.round(wait/1000)}s...`));
                await pendingSave.catch(() => {});
                setTimeout(() => connectToWhatsApp(waState).catch(console.error), wait);
            }
        } else if (connection === 'open') {
            attempts = 0; waState.isOpen = true;
            console.log(chalk.green('✅ WhatsApp terhubung!'));
            if (sock.user?.id) console.log(chalk.cyan(`   👤 ${sock.user.id}`));
        }
    });

    const opened = await waitForOpenConnection(waState, 180000);
    if (!opened) throw new Error('Timeout koneksi');
    return sock;
}

// ============================================================
// FOLDER INIT
// ============================================================
function initializeFolders() {
    ['output', 'output/final', 'output/preview', 'templates', 'cv', 'temp', 'temp/cache'].forEach(d => {
        const p = path.join(__dirname, d);
        if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true });
    });
}

// ============================================================
// DOKUMEN META
// ============================================================
function resolveDocMeta(fp) {
    const ext = (path.extname(fp) || '').toLowerCase();
    if (ext === '.pdf') return { ext: '.pdf', mimetype: 'application/pdf', label: 'PDF', isPdf: true };
    if (ext === '.docx') return { ext: '.docx', mimetype: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', label: 'DOCX', isPdf: false };
    if (ext === '.doc') return { ext: '.doc', mimetype: 'application/msword', label: 'DOC', isPdf: false };
    return { ext: ext || '.pdf', mimetype: 'application/pdf', label: (ext.replace('.', '') || 'PDF').toUpperCase(), isPdf: ext === '.pdf' };
}
function moveToTemp(srcPath) {
    if (!isUsableFile(srcPath)) return null;
    if (path.resolve(path.dirname(srcPath)) === path.resolve(TEMP_DIR)) return srcPath;
    try { if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true }); } catch {}
    const dst = path.join(TEMP_DIR, path.basename(srcPath));
    try { fs.renameSync(srcPath, dst); return dst; }
    catch { try { fs.copyFileSync(srcPath, dst); fs.unlinkSync(srcPath); return dst; } catch (e2) { console.log(chalk.yellow(`⚠️ Temp: ${e2.message}`)); return null; } }
}

// ============================================================
// LIBREOFFICE
// ============================================================
function buildConverterCandidates() {
    const isWin = process.platform === 'win32', isMac = process.platform === 'darwin';
    const arr = [];
    if (isWin) { ['soffice.exe', 'soffice', 'libreoffice'].forEach(c => arr.push({ cmd: c, args: ['--version'] })); }
    else { ['libreoffice', 'soffice'].forEach(c => arr.push({ cmd: c, args: ['--version'] })); }
    if (!isWin && !isMac) {
        ['/usr/bin/libreoffice', '/usr/bin/soffice', '/usr/local/bin/libreoffice', '/usr/local/bin/soffice', '/snap/bin/libreoffice'].forEach(c => arr.push({ cmd: c, args: ['--version'] }));
        try { const o = '/opt'; if (fs.existsSync(o)) for (const n of fs.readdirSync(o)) if (n.toLowerCase().startsWith('libreoffice')) arr.push({ cmd: path.join(o, n, 'program', 'soffice'), args: ['--version'] }); } catch {}
    }
    if (isWin) {
        const roots = [process.env['PROGRAMFILES'], process.env['PROGRAMFILES(X86)'], 'C:\\Program Files', 'C:\\Program Files (x86)'].filter(Boolean);
        for (const r of roots) for (const s of ['LibreOffice', 'LibreOffice 7', 'LibreOffice 24', 'LibreOffice 25', 'LibreOffice\\program']) {
            const exe = s.toLowerCase().endsWith('program') ? path.join(r, s, 'soffice.exe') : path.join(r, s, 'program', 'soffice.exe');
            arr.push({ cmd: exe, args: ['--version'] });
        }
    }
    if (isMac) arr.push({ cmd: '/Applications/LibreOffice.app/Contents/MacOS/soffice', args: ['--version'] });
    const seen = new Set(); const uniq = [];
    for (const c of arr) { const k = isWin ? c.cmd.toLowerCase() : c.cmd; if (seen.has(k)) continue; seen.add(k); uniq.push(c); }
    return uniq;
}
let _convCache;
async function detectConverter() {
    if (_convCache !== undefined) return _convCache;
    for (const c of buildConverterCandidates()) {
        try { await execFileAsync(c.cmd, c.args, { timeout: 15000, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] }); _convCache = c; return c; } catch {}
    }
    _convCache = null; return null;
}
async function convertDocxToPdf(srcPath, outDirOverride) {
    if (!isUsableFile(srcPath)) return null;
    if (path.extname(srcPath).toLowerCase() === '.pdf') return srcPath;
    const outDir = outDirOverride || path.dirname(srcPath);
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
    const inTemp = path.resolve(path.dirname(srcPath)) === path.resolve(TEMP_DIR);
    let working = srcPath;
    if (!inTemp) { const m = moveToTemp(srcPath); if (!m) return null; working = m; }
    const pdfName = path.basename(working, path.extname(working)) + '.pdf';
    const target = path.join(outDir, pdfName);
    if (isUsableFile(target)) { safeUnlink(working); return target; }
    const conv = await detectConverter();
    if (!conv) return working;
    const profileDir = path.join(TEMP_DIR, '.lo_profile');
    if (!fs.existsSync(profileDir)) fs.mkdirSync(profileDir, { recursive: true });
    const profileUri = 'file:///' + profileDir.replace(/\\/g, '/').replace(/^\//, '');
    try {
        await execFileAsync(conv.cmd, ['--headless', '--invisible', '--nologo', '--nofirststartwizard', '--norestore', '--nolockcheck', '--nodefault',
            `-env:UserInstallation=${profileUri}`, '--convert-to', 'pdf', '--outdir', outDir, working],
            { timeout: 90000, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
    } catch { return working; }
    if (isUsableFile(target)) { safeUnlink(working); return target; }
    return working;
}

// ============================================================
// RASTERIZE
// ============================================================
async function rasterizePdfToImagePdf(inputPath, outputPath, dpi, format = 'jpeg', quality = 92) {
    if (!HAS_GHOSTSCRIPT || !PDFDocument || !isUsableFile(inputPath)) return { success: false, error: 'precondition' };
    const workDir = path.join(TEMP_DIR, `rast_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`);
    fs.mkdirSync(workDir, { recursive: true });
    const isJpeg = format === 'jpeg', dev = isJpeg ? 'jpeg' : 'png16m', ext = isJpeg ? 'jpg' : 'png';
    try {
        const args = ['-q', '-dNOPAUSE', '-dBATCH', '-dSAFER', `-sDEVICE=${dev}`, `-r${dpi}`,
            '-dTextAlphaBits=4', '-dGraphicsAlphaBits=4', '-dAutoRotatePages=/None', '-dUseCropBox',
            `-sOutputFile=${path.join(workDir, `page-%04d.${ext}`)}`];
        if (isJpeg) args.push(`-dJPEGQ=${quality}`);
        args.push(inputPath);
        await execFileAsync('gs', args, { timeout: 300000, windowsHide: true });
        const files = fs.readdirSync(workDir).filter(f => f.endsWith(`.${ext}`)).sort();
        if (!files.length) throw new Error('Tidak ada halaman');
        const newPdf = await PDFDocument.create();
        for (const f of files) {
            const b = fs.readFileSync(path.join(workDir, f));
            const img = isJpeg ? await newPdf.embedJpg(b) : await newPdf.embedPng(b);
            const p = newPdf.addPage([img.width, img.height]);
            p.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
        }
        fs.writeFileSync(outputPath, await newPdf.save());
        for (const f of files) safeUnlink(path.join(workDir, f));
        try { fs.rmdirSync(workDir); } catch {}
        return { success: true, path: outputPath, pages: files.length };
    } catch (e) {
        try { for (const f of fs.readdirSync(workDir)) safeUnlink(path.join(workDir, f)); fs.rmdirSync(workDir); } catch {}
        return { success: false, error: e.message };
    }
}
let _cachedIjazahRasterPath = null;
async function getRasterizedIjazah() {
    if (!HAS_GHOSTSCRIPT || !isUsableFile(IJAZAH_PATH)) return null;
    if (_cachedIjazahRasterPath && isUsableFile(_cachedIjazahRasterPath)) return _cachedIjazahRasterPath;
    const cacheDir = path.join(TEMP_DIR, 'cache');
    if (!fs.existsSync(cacheDir)) fs.mkdirSync(cacheDir, { recursive: true });
    const cachePath = path.join(cacheDir, `ijazah_dpi${BOT_CONFIG.FLATTEN_DPI}_${BOT_CONFIG.FLATTEN_FORMAT}.pdf`);
    if (BOT_CONFIG.CACHE_RASTERIZED_IJAZAH && isUsableFile(cachePath)) {
        console.log(chalk.gray(`   💾 Cache ijazah (disk)`)); _cachedIjazahRasterPath = cachePath; return cachePath;
    }
    console.log(chalk.cyan(`   🖼️  Rasterize Ijazah DPI ${BOT_CONFIG.FLATTEN_DPI}...`));
    const r = await rasterizePdfToImagePdf(IJAZAH_PATH, cachePath, BOT_CONFIG.FLATTEN_DPI, BOT_CONFIG.FLATTEN_FORMAT, BOT_CONFIG.FLATTEN_JPEG_QUALITY);
    if (!r.success) return null;
    console.log(chalk.gray(`   ↳ Ijazah: ${r.pages}h (cache)`));
    _cachedIjazahRasterPath = cachePath; return cachePath;
}

// ============================================================
// WATERMARK
// ============================================================
async function applyWatermark(pdfPath, outputPath, opts = {}) {
    if (!PDFDocument || !StandardFonts || !isUsableFile(pdfPath)) return { success: false, error: 'precondition' };
    const { text, fontSize = 32, opacity = 0.10, color = [0.55, 0.55, 0.55], diagonal = true, angle = 45 } = opts;
    try {
        const doc = await PDFDocument.load(fs.readFileSync(pdfPath), { ignoreEncryption: true });
        const font = await doc.embedFont(StandardFonts.HelveticaBold);
        const c = rgbFn(color[0], color[1], color[2]);
        for (const page of doc.getPages()) {
            const { width, height } = page.getSize();
            const tw = font.widthOfTextAtSize(text, fontSize);
            if (diagonal) {
                const rot = degreesFn(angle), stepX = tw + 90, stepY = fontSize * 5;
                for (let y = -height; y < height * 2; y += stepY)
                    for (let x = -width; x < width * 2; x += stepX)
                        page.drawText(text, { x, y, size: fontSize, font, color: c, opacity, rotate: rot });
            } else {
                const th = font.heightAtSize(fontSize);
                page.drawText(text, { x: (width - tw) / 2, y: (height - th) / 2, size: fontSize, font, color: c, opacity });
            }
        }
        fs.writeFileSync(outputPath, await doc.save());
        return { success: true, pages: doc.getPageCount() };
    } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// MERGE
// ============================================================
async function mergePdfsFixed(paths, outputPath) {
    if (!PDFDocument) return { success: false, reason: 'pdf-lib-missing' };
    const order = ['letter', 'cv', 'ijazah'];
    const srcMap = { letter: paths.letterPath, cv: paths.cvPath, ijazah: paths.ijazahPath };
    try {
        const merged = await PDFDocument.create();
        const counts = {}, order2 = [];
        for (const k of order) {
            const src = srcMap[k];
            if (!isUsableFile(src)) continue;
            const doc = await PDFDocument.load(fs.readFileSync(src), { ignoreEncryption: true });
            const pages = await merged.copyPages(doc, doc.getPageIndices());
            pages.forEach(p => merged.addPage(p));
            counts[k] = doc.getPageCount();
            order2.push(k);
        }
        if (!order2.length) return { success: false, reason: 'no-docs' };
        fs.writeFileSync(outputPath, await merged.save());
        return { success: true, path: outputPath, pageCounts: counts, addedOrder: order2, totalPages: merged.getPageCount() };
    } catch (e) { return { success: false, reason: e.message }; }
}

// ============================================================
// ENCRYPT
// ============================================================
async function encryptPdf(inputPath, outputPath, opts = {}) {
    if (!HAS_GHOSTSCRIPT) return { success: false, error: 'gs-missing' };
    const { blockPrint = true, allowCopy = true, allowAnnotation = true,
        ownerPassword = BOT_CONFIG.GS_OWNER_PASSWORD || ('owner-' + Date.now()) } = opts;
    let perm = -3904;
    perm += 512;
    if (!blockPrint) perm += 4 + 2048;
    if (allowCopy) perm += 16;
    if (allowAnnotation) perm += 32;
    try {
        await execFileAsync('gs', ['-q', '-dNOPAUSE', '-dBATCH', '-dSAFER', '-sDEVICE=pdfwrite',
            '-dCompatibilityLevel=1.7', '-dEncryptionR=3', '-dKeyLength=128',
            `-sOwnerPassword=${ownerPassword}`, '-sUserPassword=', `-dPermissions=${perm}`,
            '-dAutoRotatePages=/None', `-sOutputFile=${outputPath}`, inputPath],
            { timeout: 120000, windowsHide: true });
        return { success: true, path: outputPath, permissions: perm };
    } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// CLEAN PDF — Surat + CV saja (untuk backup)
// ============================================================
async function buildCleanPdf(letterPath, target, safeName) {
    const cleanPath = path.join(TEMP_DIR, `clean_${safeName}_${Date.now()}_${Math.random().toString(36).slice(2, 6)}.pdf`);
    const cvPath = isUsableFile(BOT_CONFIG.CV_PATH) ? BOT_CONFIG.CV_PATH : null;
    const res = await mergePdfsFixed({ letterPath, cvPath, ijazahPath: null }, cleanPath);
    if (!res.success) { console.log(chalk.yellow(`   ⚠️ Clean PDF gagal: ${res.reason}`)); return null; }
    const parts = res.addedOrder.map(k => `${k}:${res.pageCounts[k]}h`).join(' + ');
    console.log(chalk.gray(`   📄 Clean PDF (Surat+CV): ${parts} → ${res.totalPages}h`));
    return cleanPath;
}

// ============================================================
// EMAIL TRANSPORTER
// ============================================================
let _emailTransporter = null;
function getEmailTransporter() {
    if (_emailTransporter) return _emailTransporter;
    _emailTransporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            type: 'OAuth2',
            user: process.env.GOOGLE_USER || process.env.EMAIL_USER,
            clientId: process.env.GOOGLE_CLIENT_ID || '407408718192.apps.googleusercontent.com',
            clientSecret: process.env.GOOGLE_CLIENT_SECRET || '',
            refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
        },
    });
    return _emailTransporter;
}
async function ensureEmailAuth() {
    const o = new google.auth.OAuth2(
        process.env.GOOGLE_CLIENT_ID || '407408718192.apps.googleusercontent.com',
        process.env.GOOGLE_CLIENT_SECRET || '', 'https://developers.google.com/oauthplayground');
    if (!process.env.GOOGLE_REFRESH_TOKEN) { const e = new Error('REFRESH_TOKEN hilang'); e.isConfigError = true; throw e; }
    o.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
    try { await o.getAccessToken(); } catch { const e = new Error('Refresh token invalid'); e.isConfigError = true; throw e; }
}

// ============================================================
// GOOGLE SHEETS — LOAD (A:H, kolom H = BACKUP_ONLY)
// ============================================================
async function loadTargetsFromSheet() {
    try {
        const auth = new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
        const sheets = google.sheets({ version: 'v4', auth });
        const SHEET_NAME = BOT_CONFIG.SHEET_NAME;

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${SHEET_NAME}!A:H`,   // A:H — H = BACKUP_ONLY
        });
        const rows = response.data.values;
        if (!rows || rows.length < 2) return [];

        const headers = rows[0].map(h => h.trim().toUpperCase());
        const companyIdx = headers.indexOf('COMPANY');
        const positionIdx = headers.indexOf('POSITION');
        const contactIdx = headers.indexOf('CONTACT');
        const requirementsIdx = headers.indexOf('REQUIREMENTS');
        const statusIdx = headers.indexOf('STATUS');
        const filePathIdx = headers.indexOf('FILE_PATH');
        const subjectIdx = headers.indexOf('SUBJECT');
        const backupOnlyIdx = headers.indexOf('BACKUP_ONLY');   // ← kolom H

        if (companyIdx === -1 || positionIdx === -1 || statusIdx === -1) {
            console.error(chalk.red('❌ Header wajib: COMPANY, POSITION, STATUS'));
            return null;
        }

        const allRows = [];
        for (let i = 1; i < rows.length; i++) {
            const r = rows[i] || [];
            const backupOnlyRaw = backupOnlyIdx !== -1 ? (r[backupOnlyIdx]?.trim() || '') : '';
            // "1" → backup only; kosong / lainnya → default (kirim + backup)
            const backupOnly = backupOnlyRaw === '1';
            allRows.push({
                sheetRow: i + 1,
                company: r[companyIdx]?.trim() || '',
                position: r[positionIdx]?.trim() || '',
                contact: contactIdx !== -1 ? (r[contactIdx]?.trim() || '') : '',
                requirements: requirementsIdx !== -1 ? (r[requirementsIdx]?.trim() || '') : '',
                status: r[statusIdx]?.trim().toLowerCase() || '',
                filePath: filePathIdx !== -1 ? (r[filePathIdx]?.trim() || '') : '',
                subject: subjectIdx !== -1 ? (r[subjectIdx]?.trim() || '') : '',
                backupOnly,
            });
        }

        const sentKeys = new Set();
        for (const r of allRows) if ((r.status === 'sent' || r.status === 'skip') && r.company && r.position)
            sentKeys.add(`${r.company.toLowerCase()}|||${r.position.toLowerCase()}`);

        const seen = new Set(), toDelete = [], targets = [];
        for (const r of allRows) {
            if (!r.company || !r.position) continue;
            if (r.status === 'sent' || r.status === 'skip') continue;
            const key = `${r.company.toLowerCase()}|||${r.position.toLowerCase()}`;
            if (sentKeys.has(key)) { toDelete.push(r.sheetRow); console.log(chalk.yellow(`🗑️ [${r.sheetRow}] sudah pernah dikirim`)); continue; }
            if (seen.has(key)) { toDelete.push(r.sheetRow); continue; }
            seen.add(key);
            targets.push(r);
        }

        if (toDelete.length) {
            try { await deleteSheetRowsByIndex(sheets, SHEET_NAME, toDelete); console.log(chalk.green(`✅ ${toDelete.length} duplikat dihapus.`)); }
            catch (e) { console.log(chalk.yellow(`⚠️ ${e.message}`)); }
        }
        return targets;
    } catch (e) { console.error(chalk.red('❌ Load sheet:'), e.message); return null; }
}

async function deleteSheetRowsByIndex(sheets, sheetTitle, rowNumbers) {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID, fields: 'sheets.properties' });
    const sheet = (meta.data.sheets || []).find(s => s.properties?.title === sheetTitle);
    if (!sheet) throw new Error(`Sheet tidak ditemukan`);
    const sheetId = sheet.properties.sheetId;
    const sorted = [...new Set(rowNumbers)].sort((a, b) => b - a);
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        resource: { requests: sorted.map(n => ({ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: n - 1, endIndex: n } } })) },
    });
}

async function updateSheetStatus(company, position, status = 'sent', filePath = '') {
    try {
        const auth = new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
        const sheets = google.sheets({ version: 'v4', auth });
        const SHEET_NAME = BOT_CONFIG.SHEET_NAME;
        const res = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `${SHEET_NAME}!A:H` });
        const rows = res.data.values; if (!rows) return false;
        const headers = rows[0].map(h => h.trim().toUpperCase());
        const companyIdx = headers.indexOf('COMPANY');
        const positionIdx = headers.indexOf('POSITION');
        const statusIdx = headers.indexOf('STATUS');
        const filePathIdx = headers.indexOf('FILE_PATH');
        let targetRow = -1;
        for (let i = 1; i < rows.length; i++) {
            if ((rows[i]?.[companyIdx]?.trim() || '') === company && (rows[i]?.[positionIdx]?.trim() || '') === position) { targetRow = i + 1; break; }
        }
        if (targetRow === -1) return false;
        const statusCol = String.fromCharCode(65 + statusIdx);
        await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID, range: `${SHEET_NAME}!${statusCol}${targetRow}`,
            valueInputOption: 'RAW', resource: { values: [[status]] },
        });
        if (filePath && filePathIdx !== -1) {
            const fpCol = String.fromCharCode(65 + filePathIdx);
            await sheets.spreadsheets.values.update({
                spreadsheetId: SPREADSHEET_ID, range: `${SHEET_NAME}!${fpCol}${targetRow}`,
                valueInputOption: 'RAW', resource: { values: [[filePath]] },
            });
        }
        return true;
    } catch (e) { console.error(chalk.red('❌ Update Sheet:'), e.message); return false; }
}

// ============================================================
// EMAIL — KIRIM KE TARGET
// ============================================================
async function attemptSendEmailToTarget(contact, company, position, letterPath, customSubject, letterFileName) {
    await ensureEmailAuth();
    const transporter = getEmailTransporter();
    const subject = (customSubject && customSubject.trim()) ? customSubject.trim() : buildDefaultSubject(position);
    const attachments = [];
    if (isUsableFile(letterPath)) {
        const fname = letterFileName && letterFileName.trim() ? letterFileName.trim() : path.basename(letterPath);
        attachments.push({ filename: fname, path: letterPath });
    }
    await transporter.sendMail({
        from: `"${BOT_CONFIG.APPLICANT_NAME}" <${process.env.GOOGLE_USER || process.env.EMAIL_USER}>`,
        to: contact, subject,
        text: buildEmailText(company, position),
        html: buildEmailHtml(company, position),
        attachments,
    });
}
async function sendEmailToTarget(contact, company, position, letterPath, customSubject, letterFileName) {
    let attempt = 0;
    while (true) {
        attempt++;
        try {
            console.log(chalk.gray(`   ⏳ Kirim ke ${contact}...`));
            await attemptSendEmailToTarget(contact, company, position, letterPath, customSubject, letterFileName);
            console.log(chalk.green(`📧 Email → target: ${contact}${attempt > 1 ? ` (try ${attempt})` : ''}`));
            return true;
        } catch (e) {
            console.error(chalk.red(`❌ Gagal (${attempt}):`), e.message);
            if (e.isConfigError) return false;
            await delay(Math.min(5000 * attempt, 5 * 60 * 1000));
        }
    }
}

// ============================================================
// EMAIL BACKUP
// ============================================================
async function sendBackupEmail(position, cleanPath) {
    if (!BOT_CONFIG.BACKUP_ENABLED) return { success: false, reason: 'disabled' };
    if (!isUsableFile(cleanPath)) return { success: false, reason: 'clean-missing' };
    try {
        await ensureEmailAuth();
        const transporter = getEmailTransporter();
        const fileName = path.basename(cleanPath);
        const subject = `${BOT_CONFIG.BACKUP_SUBJECT_PREFIX} - ${position}`;
        await transporter.sendMail({
            from: `"${BOT_CONFIG.APPLICANT_NAME}" <${process.env.GOOGLE_USER || process.env.EMAIL_USER}>`,
            to: BOT_CONFIG.BACKUP_EMAIL,
            subject, text: '',
            attachments: [{ filename: fileName, path: cleanPath }],
        });
        console.log(chalk.green(`📧 Backup → ${BOT_CONFIG.BACKUP_EMAIL} (${fileName})`));
        return { success: true };
    } catch (e) {
        console.error(chalk.red(`❌ Gagal backup: ${e.message}`));
        return { success: false, reason: e.message };
    }
}

// ============================================================
// WHATSAPP
// ============================================================
function toJid(number) {
    let raw = number.replace(/\D/g, '');
    if (raw.startsWith('0')) raw = '62' + raw.slice(1);
    else if (!raw.startsWith('62')) raw = '62' + raw;
    return raw + '@s.whatsapp.net';
}
async function sendWhatsApp(waState, contact, company, position, letterPath, letterFileName, maxRetries = 3) {
    const key = `${contact}|${company}|${position}`;
    if (waState.sentTargets.has(key)) return true;
    const jid = toJid(contact);
    const msgText = `${getGreetingByTime()}, Bapak/Ibu HRD ${company},\n\nPerkenalkan, saya ${BOT_CONFIG.APPLICANT_NAME}. Sehubungan dengan informasi lowongan posisi ${position} di ${company}, saya bermaksud mengajukan lamaran untuk posisi tersebut.\n\nBersama pesan ini, saya lampirkan CV dan Surat Lamaran Kerja sebagai bahan pertimbangan awal. Besar harapan saya untuk dapat diberikan kesempatan wawancara.\n\nTerima kasih atas waktu dan perhatian Bapak/Ibu.\n\nHormat saya,\n${BOT_CONFIG.APPLICANT_NAME}\n${BOT_CONFIG.APPLICANT_PHONE}`;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        const ok = await waitForOpenConnection(waState, 90000);
        if (!ok) { if (attempt < maxRetries) await delay(5000 * attempt); continue; }
        const sock = waState.sock;
        try {
            let tJid = jid;
            try {
                const [r] = await sock.onWhatsApp(jid);
                if (!r || !r.exists) { console.log(chalk.yellow(`⚠️ ${contact} tidak terdaftar`)); return false; }
                if (r.jid && r.jid !== jid) tJid = r.jid;
            } catch {}
            const tr = await sendMessageConfirmed(waState, tJid, { text: msgText });
            if (!tr.success) throw new Error(tr.reason);
            await delay(jitterDelay(BOT_CONFIG.DELAY_BETWEEN_SENDS));
            if (isUsableFile(letterPath)) {
                const buf = fs.readFileSync(letterPath);
                const meta = resolveDocMeta(letterPath);
                const fname = letterFileName && letterFileName.trim() ? letterFileName.trim() : `Surat-${company}${meta.ext}`;
                const dr = await sendMessageConfirmed(waState, tJid, {
                    document: buf, mimetype: meta.mimetype, fileName: fname,
                    caption: `Lamaran - ${company} (${position})`,
                });
                if (!dr.success) throw new Error(dr.reason);
                await delay(jitterDelay(BOT_CONFIG.DELAY_BETWEEN_SENDS));
            }
            console.log(chalk.green(`📱 WhatsApp → ${contact}`));
            waState.sentTargets.add(key);
            return true;
        } catch (e) {
            console.error(chalk.red(`❌ WA (${attempt}/${maxRetries}):`), e.message);
            if (attempt < maxRetries) await delay(5000 * attempt);
        }
    }
    return false;
}

// ============================================================
// GENERATE SURAT
// ============================================================
async function getOrCreateLetterPdf(target, safeName) {
    const outPdf = path.join(AI_CONFIG.OUTPUT_DIR, `${target.position} - ${safeName}.pdf`);
    const outDocx = path.join(AI_CONFIG.OUTPUT_DIR, `${target.position} - ${safeName}.docx`);

    if (isUsableFile(target.filePath)) {
        const m = resolveDocMeta(target.filePath);
        if (m.isPdf) { console.log(chalk.green(`✅ PDF tersimpan: ${path.basename(target.filePath)}`)); return { pdfPath: target.filePath, fromCache: true }; }
        const p = await convertDocxToPdf(target.filePath, path.dirname(target.filePath));
        if (p) return { pdfPath: p, fromCache: true };
    }
    if (isUsableFile(outPdf)) { console.log(chalk.green(`✅ PDF output: ${path.basename(outPdf)}`)); return { pdfPath: outPdf, fromCache: true }; }
    if (isUsableFile(outDocx)) { const p = await convertDocxToPdf(outDocx, AI_CONFIG.OUTPUT_DIR); if (p) return { pdfPath: p, fromCache: true }; }

    const spinner = ora('Generate surat via AI...').start();
    let result;
    try { result = await generateApplicationLetter(target.company, target.position, outPdf, target.requirements || ''); }
    catch (e) { spinner.fail('Error: ' + e.message); return null; }
    if (!result || !result.success || !result.actualPath) { spinner.fail('Gagal'); return null; }
    let genPath = result.actualPath;
    if (!isUsableFile(genPath)) { spinner.fail('File kosong'); return null; }
    const meta = resolveDocMeta(genPath);
    if (meta.isPdf) { spinner.succeed(`✅ Surat: ${path.basename(genPath)}`); return { pdfPath: genPath, fromCache: false }; }
    spinner.text = `Konversi ${meta.label} → PDF...`;
    const pdf = await convertDocxToPdf(genPath, AI_CONFIG.OUTPUT_DIR);
    if (!pdf) { spinner.fail('Konversi gagal'); return null; }
    spinner.succeed(`✅ Surat PDF: ${path.basename(pdf)}`);
    return { pdfPath: pdf, fromCache: false };
}

// ============================================================
// BUILD TARGET PDF
// ============================================================
async function getOrCreateTargetPdf(letterPath, target, safeName) {
    const letterDir = path.resolve(path.dirname(letterPath));
    const officialOut = path.resolve(path.join(__dirname, 'output'));

    if (letterDir === path.resolve(FINAL_DIR)) return { path: letterPath, merged: false, fileName: path.basename(letterPath) };
    if (!BOT_CONFIG.MERGE_WITH_IJAZAH) return { path: letterPath, merged: false, fileName: path.basename(letterPath) };

    const fileName = buildOutputFileName(target.position);
    const finalPath = path.join(FINAL_DIR, fileName);
    const cvPath = BOT_CONFIG.CV_PATH;
    const hasCv = isUsableFile(cvPath);
    const hasIjazah = isUsableFile(IJAZAH_PATH);

    if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
    const wid = `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;

    let pLetter = letterPath, pCv = hasCv ? cvPath : null, pIjazah = hasIjazah ? IJAZAH_PATH : null;
    const tasks = [];

    if (BOT_CONFIG.FLATTEN_LETTER && HAS_GHOSTSCRIPT) {
        const rp = path.join(TEMP_DIR, `letter_flat_${wid}.pdf`);
        tasks.push(rasterizePdfToImagePdf(letterPath, rp, BOT_CONFIG.FLATTEN_DPI, BOT_CONFIG.FLATTEN_FORMAT, BOT_CONFIG.FLATTEN_JPEG_QUALITY)
            .then(r => { if (r.success) { pLetter = rp; console.log(chalk.gray(`   ↳ Surat: ${r.pages}h img`)); } }));
    }
    if (BOT_CONFIG.FLATTEN_IJAZAH && hasIjazah && HAS_GHOSTSCRIPT) {
        tasks.push(getRasterizedIjazah().then(p => { if (p) pIjazah = p; }));
    }
    if (tasks.length) { console.log(chalk.cyan(`   🖼️  Rasterize (DPI ${BOT_CONFIG.FLATTEN_DPI})...`)); await Promise.all(tasks); }

    if (hasCv && !BOT_CONFIG.FLATTEN_CV) console.log(chalk.cyan(`   📄 CV: teks (hyperlink aktif)`));

    const mergedPath = path.join(TEMP_DIR, `merged_${wid}.pdf`);
    const mr = await mergePdfsFixed({ letterPath: pLetter, cvPath: pCv, ijazahPath: pIjazah }, mergedPath);
    if (!mr.success) {
        console.log(chalk.yellow(`   ⚠️ Merge gagal: ${mr.reason}`));
        if (letterDir === officialOut) safeUnlink(letterPath);
        return { path: letterPath, merged: false, fileName: path.basename(letterPath) };
    }
    console.log(chalk.cyan(`   📎 Merge ${mr.addedOrder.map(k => `${k}:${mr.pageCounts[k]}h`).join(' + ')} → ${mr.totalPages}h`));

    let wmPath = mergedPath, wmApplied = false;
    if (BOT_CONFIG.WATERMARK_ENABLED) {
        const tid = (Date.now().toString(36) + Math.random().toString(36).slice(2, 7)).toUpperCase();
        const wmText = buildWatermarkText(target, tid);
        const wp = path.join(TEMP_DIR, `wm_${wid}.pdf`);
        console.log(chalk.cyan(`   💧 Watermark: "${wmText}"`));
        const wmR = await applyWatermark(mergedPath, wp, {
            text: wmText, fontSize: BOT_CONFIG.WATERMARK_FONT_SIZE, opacity: BOT_CONFIG.WATERMARK_OPACITY,
            color: BOT_CONFIG.WATERMARK_COLOR, diagonal: BOT_CONFIG.WATERMARK_DIAGONAL, angle: BOT_CONFIG.WATERMARK_ANGLE,
        });
        if (wmR.success) { wmPath = wp; wmApplied = true; console.log(chalk.gray(`   ↳ ${wmR.pages}h`)); }
        else console.log(chalk.yellow(`   ⚠️ WM gagal: ${wmR.error}`));
        safeUnlink(mergedPath);
    }

    const stampPath = path.join(TEMP_DIR, `stamp_${wid}.pdf`);
    console.log(chalk.cyan(`   🔗 Stamp + footer...`));
    let stR;
    try {
        stR = await mergeAndStamp({
            letterPath: wmPath, ijazahPath: null,
            companyName: target.company, position: target.position,
            outputPath: stampPath, tmpDir: TEMP_DIR,
            stampOnly: TM_STAMP.stampOnly, stampCorner: TM_STAMP.corner,
            stampFontSize: TM_STAMP.fontSize, stampOpacity: TM_STAMP.opacity,
            stampPadding: TM_STAMP.padding, stampColor: TM_STAMP.color,
            stampFontStyle: TM_STAMP.fontStyle, stampBg: TM_STAMP.bg,
            stampBgColor: TM_STAMP.bgColor, stampBgOpacity: TM_STAMP.bgOpacity,
            stampBgPadding: TM_STAMP.bgPadding, stampBorderColor: TM_STAMP.borderColor,
            stampBorderWidth: TM_STAMP.borderWidth, stampOutline: TM_STAMP.outline,
            stampOutlineColor: TM_STAMP.outlineColor, stampOutlineWidth: TM_STAMP.outlineWidth,
            enableFooter: true, enableMetadata: true, noEncrypt: true,
        });
    } catch (e) {
        safeUnlink(wmPath); if (letterDir === officialOut) safeUnlink(letterPath);
        return { path: letterPath, merged: false, fileName: path.basename(letterPath) };
    }
    safeUnlink(wmPath);
    if (!stR.success) {
        safeUnlink(stampPath); if (letterDir === officialOut) safeUnlink(letterPath);
        return { path: letterPath, merged: false, fileName: path.basename(letterPath) };
    }

    let finalUsed = stampPath, enc = false;
    if (HAS_GHOSTSCRIPT) {
        const er = await encryptPdf(stampPath, finalPath, {
            blockPrint: BOT_CONFIG.BLOCK_PRINT, allowCopy: BOT_CONFIG.ALLOW_COPY,
            allowAnnotation: BOT_CONFIG.ALLOW_ANNOTATION,
        });
        if (er.success) { finalUsed = finalPath; enc = true; safeUnlink(stampPath); console.log(chalk.green(`   🔒 Encrypt OK`)); }
        else { console.log(chalk.yellow(`   ⚠️ Encrypt gagal: ${er.error}`)); safeUnlink(finalPath); try { fs.renameSync(stampPath, finalPath); finalUsed = finalPath; } catch {} }
    } else { safeUnlink(finalPath); try { fs.renameSync(stampPath, finalPath); finalUsed = finalPath; } catch {} }

    if (pLetter !== letterPath && path.resolve(path.dirname(pLetter)) === path.resolve(TEMP_DIR)) safeUnlink(pLetter);

    const tags = [];
    if (hasCv && !BOT_CONFIG.FLATTEN_CV) tags.push('CV teks');
    if (hasIjazah) tags.push('Ijazah');
    if (wmApplied) tags.push('WM');
    if (enc) tags.push('🔒');
    console.log(chalk.green(`   ✅ Target: ${fileName} (${stR.totalPages}h${tags.length ? ', ' + tags.join(', ') : ''})`));

    if (letterDir === officialOut) safeUnlink(letterPath);

    return { path: finalUsed, merged: true, fileName, totalPages: stR.totalPages, encrypted: enc, watermarked: wmApplied };
}

// ============================================================
// PROCESS TARGET
// ============================================================
async function processTarget(target, waState) {
    const sendToTarget = !target.backupOnly;   // backupOnly=true → skip kirim target
    const safeName = sanitizeName(target.company);
    const label = sendToTarget ? 'TARGET' : 'BACKUP-ONLY';
    console.log(chalk.cyan(`\n▶ [${label}] ${target.company} — ${target.position}`));

    // 1. Surat
    const letterInfo = await getOrCreateLetterPdf(target, safeName);
    if (!letterInfo) {
        await updateSheetStatus(target.company, target.position, 'failed');
        return { success: false, cacheHit: false, aiCall: false };
    }
    let letterPath = letterInfo.pdfPath;
    const cacheHit = letterInfo.fromCache;
    if (cacheHit) console.log(chalk.gray(`   💾 Cache hit`));

    // 2. Clean PDF (Surat + CV)
    let cleanPath = null;
    if (BOT_CONFIG.BACKUP_ENABLED) cleanPath = await buildCleanPdf(letterPath, target, safeName);

    // 3. Kirim ke target (kalau bukan backup-only)
    let targetSendOk = true;
    let targetPath = null;
    let letterFileName = null;

    if (sendToTarget) {
        const mergedInfo = await getOrCreateTargetPdf(letterPath, target, safeName);
        targetPath = mergedInfo.path;
        letterFileName = mergedInfo.fileName || path.basename(targetPath);

        if (isEmail(target.contact)) {
            const subj = (target.subject && target.subject.trim()) ? target.subject.trim() : buildDefaultSubject(target.position);
            console.log(chalk.gray(`   ✉️  Subjek: "${subj}"${(target.subject && target.subject.trim()) ? ' (custom)' : ' (default)'}`));
            console.log(chalk.gray(`   📄 File  : ${letterFileName}`));
            targetSendOk = await sendEmailToTarget(target.contact, target.company, target.position, targetPath, target.subject, letterFileName);
        } else if (target.contact && waState) {
            console.log(chalk.gray(`   📄 File  : ${letterFileName}`));
            targetSendOk = await sendWhatsApp(waState, target.contact, target.company, target.position, targetPath, letterFileName);
        } else if (!target.contact) {
            console.log(chalk.yellow(`   ⚠️ Tidak ada CONTACT — kirim ke target dilewati`));
            targetSendOk = false;
        } else {
            console.log(chalk.red(`   ❌ WA state tidak tersedia`));
            targetSendOk = false;
        }
    } else {
        console.log(chalk.gray(`   ℹ️  Mode BACKUP_ONLY — hanya kirim file bersih ke ${BOT_CONFIG.BACKUP_EMAIL}`));
    }

    // 4. Backup (kalau target berhasil ATAU mode backup-only)
    let backupOk = false;
    if (targetSendOk && BOT_CONFIG.BACKUP_ENABLED && cleanPath) {
        const br = await sendBackupEmail(target.position, cleanPath);
        backupOk = br.success;
        if (!backupOk) console.log(chalk.yellow(`   ⚠️ Backup gagal: ${br.reason}`));
    }

    // 5. Update sheet + cleanup
    if (targetSendOk) {
        await updateSheetStatus(target.company, target.position, 'sent', targetPath || cleanPath || '');
        if (BOT_CONFIG.DELETE_AFTER_SENT) {
            if (targetPath && safeUnlink(targetPath)) console.log(chalk.gray(`   🗑️  Target: ${path.basename(targetPath)}`));
            if (cleanPath) {
                if (backupOk || !BOT_CONFIG.BACKUP_ENABLED) {
                    if (safeUnlink(cleanPath)) console.log(chalk.gray(`   🗑️  Clean: ${path.basename(cleanPath)}`));
                } else {
                    console.log(chalk.yellow(`   ⏸️  Clean disimpan (retry backup): ${cleanPath}`));
                }
            }
        }
        return { success: true, cacheHit, aiCall: !cacheHit, backupOk, backupOnly: !sendToTarget };
    } else {
        await updateSheetStatus(target.company, target.position, 'failed', targetPath || '');
        if (cleanPath) console.log(chalk.gray(`   ℹ️  Clean disimpan: ${cleanPath}`));
        if (targetPath) console.log(chalk.gray(`   ℹ️  Target disimpan: ${targetPath}`));
        return { success: false, cacheHit, aiCall: !cacheHit };
    }
}

// ============================================================
// MAIN
// ============================================================
async function main() {
    console.log(chalk.bold.cyan('\n' + '='.repeat(64)));
    console.log(chalk.bold.cyan('📱 WHATSAPP & EMAIL JOB APPLICATION BOT'));
    console.log(chalk.bold.cyan(`   Sheet            : ${BOT_CONFIG.SHEET_NAME} (kolom A:H)`));
    console.log(chalk.bold.cyan(`   Kolom H          : BACKUP_ONLY ("1" = backup saja)`));
    console.log(chalk.bold.cyan(`   Backup →         : ${BOT_CONFIG.BACKUP_EMAIL} (Surat + CV)`));
    console.log(chalk.bold.cyan(`   PDF              : Surat→CV→Ijazah | WM: ${BOT_CONFIG.WATERMARK_MODE}`));
    console.log(chalk.bold.cyan(`   Permission       : Print=${BOT_CONFIG.BLOCK_PRINT ? 'BLOCKED' : 'OK'} | Copy=${BOT_CONFIG.ALLOW_COPY ? 'OK' : 'BLOCKED'}`));
    console.log(chalk.bold.cyan(`   Email template   : HTML profesional`));

    const gs = await checkGhostscript();
    if (gs) console.log(chalk.bold.green(`   ✅ Ghostscript v${gs}`));
    else { console.log(chalk.bold.yellow(`   ⚠️  Ghostscript tidak ada`)); if (BOT_CONFIG.REQUIRE_GHOSTSCRIPT) process.exit(1); }
    if (!PDFDocument) console.log(chalk.bold.yellow('   ⚠️  pdf-lib tidak ada!'));
    console.log(chalk.bold.cyan('='.repeat(64) + '\n'));

    initializeFolders();

    if (!fs.existsSync(AI_CONFIG.TEMPLATE_PATH)) { console.log(chalk.red(`❌ Template: ${AI_CONFIG.TEMPLATE_PATH}`)); process.exit(1); }
    if (!fs.existsSync(BOT_CONFIG.CV_PATH)) { console.log(chalk.red(`❌ CV: ${BOT_CONFIG.CV_PATH}`)); process.exit(1); }
    if (!SPREADSHEET_ID) { console.log(chalk.red('❌ SPREADSHEET_ID belum diatur')); process.exit(1); }

    const targets = await loadTargetsFromSheet();
    if (!targets) { console.log(chalk.red('❌ Gagal load sheet')); process.exit(1); }
    if (!targets.length) { console.log(chalk.green('\n✅ Tidak ada target.')); process.exit(0); }

    const normalTargets = targets.filter(t => !t.backupOnly);
    const backupOnlyTargets = targets.filter(t => t.backupOnly);

    console.log(chalk.yellow(`\n📋 ${targets.length} target akan diproses:`));
    if (normalTargets.length) {
        console.log(chalk.cyan(`\n  📤 Kirim + backup (${normalTargets.length}):`));
        normalTargets.forEach((t, i) => {
            const ch = isEmail(t.contact) ? '📧' : '📱';
            const subj = (t.subject && t.subject.trim()) || buildDefaultSubject(t.position);
            console.log(`    ${i+1}. ${ch} ${t.company} - ${t.position}`);
            if (ch === '📧') console.log(chalk.gray(`        subjek: "${subj}"`));
        });
    }
    if (backupOnlyTargets.length) {
        console.log(chalk.magenta(`\n  💾 Backup only (${backupOnlyTargets.length}):`));
        backupOnlyTargets.forEach((t, i) => console.log(`    ${i+1}. ${t.company} - ${t.position}`));
    }

    // Connect WA kalau ada target WA di normalTargets
    const hasWA = normalTargets.some(t => t.contact && !isEmail(t.contact));
    let waState = null;
    if (hasWA) {
        console.log(chalk.cyan('\n📱 Menghubungkan WhatsApp...'));
        waState = createWaState();
        await connectToWhatsApp(waState);
        console.log(chalk.green(`✅ WhatsApp siap, tunggu ${BOT_CONFIG.WA_SYNC_DELAY_MS / 1000}s...`));
        await delay(BOT_CONFIG.WA_SYNC_DELAY_MS);
    }

    let ok = 0, fail = 0, cacheHits = 0, aiCalls = 0, backups = 0;
    const t0 = Date.now();

    for (let i = 0; i < targets.length; i++) {
        const t = targets[i];
        console.log(chalk.gray(`\n[${i+1}/${targets.length}]`));
        const r = await processTarget(t, waState);
        if (r.success) ok++; else fail++;
        if (r.cacheHit) cacheHits++; else if (r.aiCall) aiCalls++;
        if (r.backupOk) backups++;

        if (i < targets.length - 1) {
            const dt = (!t.backupOnly && isEmail(t.contact))
                ? BOT_CONFIG.DELAY_BETWEEN_TARGETS_EMAIL_MIN + Math.floor(Math.random() * (BOT_CONFIG.DELAY_BETWEEN_TARGETS_EMAIL_MAX - BOT_CONFIG.DELAY_BETWEEN_TARGETS_EMAIL_MIN))
                : BOT_CONFIG.DELAY_BETWEEN_TARGETS_WA_MIN + Math.floor(Math.random() * (BOT_CONFIG.DELAY_BETWEEN_TARGETS_WA_MAX - BOT_CONFIG.DELAY_BETWEEN_TARGETS_WA_MIN));
            console.log(chalk.gray(`⏳ Delay ${dt/1000}s...`));
            await delay(dt);
        }
    }

    const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
    console.log(chalk.bold.green('\n📊 SUMMARY'));
    console.log(chalk.green(`✅ Berhasil   : ${ok}`));
    console.log(chalk.red(`❌ Gagal      : ${fail}`));
    console.log(chalk.cyan(`💾 Cache hit  : ${cacheHits}`));
    console.log(chalk.cyan(`🤖 AI calls   : ${aiCalls}`));
    console.log(chalk.cyan(`📤 Backup     : ${backups}`));
    console.log(chalk.cyan(`⏱️  Waktu      : ${elapsed}s`));

    process.exit(0);
}

main().catch(e => { console.error(chalk.red('\n❌ Error:'), e.message); process.exit(1); });

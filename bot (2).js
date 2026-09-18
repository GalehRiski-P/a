// bot.js - WhatsApp & Email Job Application Bot (Baileys + MongoDB)
// Alur:
//   1. Ambil target dari Sheet
//   2. Generate surat via lamaran-ai.js (cache kalau sudah ada)
//   3. Prepare tiap komponen: Surat & Ijazah → rasterize; CV → tetap teks
//   4. Merge (Surat → CV → Ijazah) dengan pdf-lib
//   5. Watermark dinamis (nama perusahaan, kontak, timestamp, tracking ID)
//   6. Stamp + footer via test-merge.js
//   7. Encrypt (block print, allow copy) via Ghostscript
//   8. Kirim dengan nama file: {Nama}_{Posisi}.pdf
//   9. Update Sheet: STATUS + FILE_PATH

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

// pdf-lib
let PDFDocument = null;
let StandardFonts = null;
let rgbFn = null;
let degreesFn = null;
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

try { dns.setDefaultResultOrder('ipv4first'); } catch { /* abaikan */ }
dotenv.config();

// ============================================================
// FILTER LOG NOISE
// ============================================================
(function installConsoleNoiseFilter() {
    const NOISE = [
        'Closing session', 'Opening session', 'Session already closed',
        'SessionEntry', '"class":"baileys"', 'failed to decrypt message',
        'sent retry receipt', 'Bad MAC', 'Closing open session',
    ];
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
    console.log = wrap(console.log);
    console.info = wrap(console.info);
    console.debug = wrap(console.debug);
    console.warn = wrap(console.warn);
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
if (!fs.existsSync(CREDENTIALS_PATH)) {
    console.error(chalk.red('❌ File credentials.json tidak ditemukan!'));
    process.exit(1);
}
const credentials = require(CREDENTIALS_PATH);
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// ============================================================
// IMPORT lamaran-ai.js & test-merge.js
// ============================================================
const { generateApplicationLetter, CONFIG: AI_CONFIG } = require('./lamaran-ai.js');
const {
    mergeAndStamp,
    sanitizeName,
    FINAL_DIR,
    IJAZAH_PATH,
    STAMP: TM_STAMP,
} = require('./test-merge.js');

// ============================================================
// KONFIGURASI BOT
// ============================================================
const BOT_CONFIG = {
    APPLICANT_NAME: 'Galeh Riski Prasetio',
    CV_PATH: path.join(__dirname, 'cv', 'CV-Galeh.pdf'),

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

    // ---- Flatten per komponen ----
    // Surat & Ijazah → gambar (tidak bisa copy)
    // CV → tetap teks (bisa copy + hyperlink tetap aktif)
    FLATTEN_LETTER: true,
    FLATTEN_CV: false,           // ← CV tetap teks
    FLATTEN_IJAZAH: true,

    // DPI tinggi biar TIDAK BLUR. 300 = standar cetak.
    // Kalau file terlalu besar, turunkan ke 200.
    FLATTEN_DPI: 300,
    FLATTEN_FORMAT: 'jpeg',      // 'jpeg' lebih kecil dari 'png'
    FLATTEN_JPEG_QUALITY: 92,    // 0-100, makin tinggi makin tajam & besar

    // ---- Permission PDF ----
    // Block print tapi ALLOW copy (biar CV bisa di-copy)
    BLOCK_PRINT: true,
    ALLOW_COPY: true,            // ← diizinkan supaya CV bisa dicopy
    ALLOW_ANNOTATION: true,      // ← hyperlink butuh ini
    GS_OWNER_PASSWORD: process.env.GS_OWNER_PASSWORD || '',

    // ---- Watermark dinamis ----
    // Ditempel di setiap halaman, diagonal berulang
    // Isi: nama perusahaan, kontak tujuan, timestamp, tracking ID
    WATERMARK_ENABLED: true,
    WATERMARK_OPACITY: 0.10,     // 0.0-1.0 — kecil biar tidak ganggu baca
    WATERMARK_FONT_SIZE: 32,
    WATERMARK_DIAGONAL: true,
    WATERMARK_ANGLE: 45,
    WATERMARK_COLOR: [0.55, 0.55, 0.55],  // light gray

    REQUIRE_GHOSTSCRIPT: false,
};

const TEMP_DIR = path.join(__dirname, 'temp');

function buildDefaultSubject(position) {
    return `Galeh Riski Prasetio - ${position}`;
}

// ============================================================
// KONFIGURASI MONGODB
// ============================================================
const MONGODB_URI = process.env.MONGODB || process.env.MONGODB_URI;
const SESSION_ID = process.env.SESSION_ID || 'whatsapp_session';
const PHONE_NUMBER = process.env.WA_PHONE_NUMBER || '';

if (!MONGODB_URI) {
    console.error(chalk.red('❌ MONGODB / MONGODB_URI tidak diatur di .env'));
    process.exit(1);
}

const { model, Schema } = mongoose;
const sessionSchema = new Schema({
    sessionId: { type: String, required: true, unique: true },
    session: { type: String },
    state: { type: Object },
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
// MONGODB HELPERS
// ============================================================
async function connectMongo() {
    if (mongoose.connection.readyState === 1) return;
    try {
        await mongoose.connect(MONGODB_URI);
        console.log(chalk.green('✅ Terhubung ke MongoDB'));
    } catch (error) {
        console.error(chalk.red('❌ Gagal koneksi MongoDB:'), error.message);
        process.exit(1);
    }
}

let _baileysModule = null;
async function getBaileys() {
    if (!_baileysModule) _baileysModule = await import('@whiskeysockets/baileys');
    return _baileysModule;
}

function looksLikeByteArrayObject(val) {
    if (!val || typeof val !== 'object' || Array.isArray(val) || Buffer.isBuffer(val)) return false;
    const keys = Object.keys(val);
    if (!keys.length) return false;
    for (let i = 0; i < keys.length; i++) {
        if (keys[i] !== String(i)) return false;
        const v = val[keys[i]];
        if (typeof v !== 'number' || !Number.isInteger(v) || v < 0 || v > 255) return false;
    }
    return true;
}

function deepRepairBuffers(node, repairedCount = { n: 0 }) {
    if (node === null || typeof node !== 'object' || Buffer.isBuffer(node)) return node;
    if (Array.isArray(node)) {
        for (let i = 0; i < node.length; i++) node[i] = deepRepairBuffers(node[i], repairedCount);
        return node;
    }
    if (node._bsontype === 'Binary' && node.buffer) {
        repairedCount.n++;
        return Buffer.isBuffer(node.buffer) ? node.buffer : Buffer.from(node.buffer);
    }
    if (node.type === 'Buffer' && node.data !== undefined) {
        const val = node.data;
        repairedCount.n++;
        if (typeof val === 'string') return Buffer.from(val, 'base64');
        if (Array.isArray(val)) return Buffer.from(val);
        if (looksLikeByteArrayObject(val)) return Buffer.from(Object.values(val));
        return Buffer.alloc(0);
    }
    if (looksLikeByteArrayObject(node)) {
        repairedCount.n++;
        return Buffer.from(Object.values(node));
    }
    for (const key of Object.keys(node)) node[key] = deepRepairBuffers(node[key], repairedCount);
    return node;
}

async function loadSessionFromMongo() {
    try {
        const doc = await SessionModel.findOne({ sessionId: SESSION_ID });
        if (!doc) { console.log(chalk.yellow('⚠️ Session tidak ditemukan')); return null; }
        console.log(chalk.green('✅ Dokumen session ditemukan'));
        if (doc.session) {
            console.log(chalk.green('✅ Menggunakan field session (string)'));
            try {
                const { BufferJSON } = await getBaileys();
                const parsed = JSON.parse(doc.session, BufferJSON.reviver);
                if (parsed.creds && parsed.keys) {
                    const rc = { n: 0 };
                    const repaired = deepRepairBuffers(parsed, rc);
                    if (rc.n > 0) {
                        console.log(chalk.yellow(`🔧 Perbaiki ${rc.n} field Buffer...`));
                        await saveSessionToMongo(repaired.creds, repaired.keys);
                    }
                    return repaired;
                }
            } catch (e) { console.error(chalk.red('❌ Parse session:'), e.message); }
        }
        if (doc.state && doc.state.creds && doc.state.keys) {
            console.log(chalk.yellow('⚠️ Fallback ke field state.'));
            const rc = { n: 0 };
            const repaired = deepRepairBuffers(doc.state.toObject ? doc.state.toObject() : doc.state, rc);
            return repaired;
        }
        return null;
    } catch (err) { console.error(chalk.red('❌ Load session:'), err.message); return null; }
}

function makeMongoKeyStore(keysData, onKeysChanged) {
    return {
        get: async (type, ids) => {
            const data = {};
            for (const id of ids) {
                let value = keysData?.[type]?.[id];
                if (value) data[id] = value;
            }
            return data;
        },
        set: async (data) => {
            for (const type in data) {
                keysData[type] = keysData[type] || {};
                for (const id in data[type]) {
                    const value = data[type][id];
                    if (value) keysData[type][id] = value;
                    else delete keysData[type][id];
                }
            }
            if (onKeysChanged) onKeysChanged();
        }
    };
}

async function saveSessionToMongo(creds, keysData) {
    try {
        if (!creds || !keysData) return;
        const state = { creds, keys: keysData };
        const { BufferJSON } = await getBaileys();
        await SessionModel.findOneAndUpdate(
            { sessionId: SESSION_ID },
            { sessionId: SESSION_ID, state, session: JSON.stringify(state, BufferJSON.replacer), updatedAt: new Date() },
            { upsert: true }
        );
        console.log(chalk.gray('💾 Session disimpan'));
    } catch (err) { console.error(chalk.yellow('⚠️ Gagal simpan session:'), err.message); }
}

function makeDebouncedSaver(getCreds, keysData, delayMs = 2000) {
    let timer = null;
    return () => {
        if (timer) clearTimeout(timer);
        timer = setTimeout(() => {
            timer = null;
            const creds = getCreds();
            if (creds) saveSessionToMongo(creds, keysData).catch(() => {});
        }, delayMs);
    };
}

async function deleteSessionFromMongo() {
    try {
        await SessionModel.deleteOne({ sessionId: SESSION_ID });
        console.log(chalk.gray('🗑️ Session dihapus'));
    } catch (err) { console.error(chalk.yellow('⚠️ Gagal hapus session:'), err.message); }
}

// ============================================================
// WA STATE & ACK
// ============================================================
function createWaState() {
    return { sock: null, isOpen: false, pendingAcks: new Map(), recentMessages: new Map(), reconnectLog: [], sentTargets: new Set() };
}
function makeGetMessage(waState) {
    return async (key) => {
        const entry = waState.recentMessages.get(key.id);
        return entry ? entry.message : undefined;
    };
}
function makeSimpleCacheStore() {
    const store = new Map();
    return { get: (k) => store.get(k), set: (k, v) => { store.set(k, v); return true; }, del: (k) => store.delete(k), flushAll: () => store.clear() };
}
function attachAckTracking(sock, waState) {
    sock.ev.on('messages.update', (updates) => {
        for (const { key, update } of updates) {
            if (key?.fromMe && typeof update?.status === 'number') console.log(chalk.gray(`   [ACK] id=${key.id} status=${update.status}`));
            if (key?.id && typeof update?.status === 'number' && waState.pendingAcks.has(key.id)) {
                const cb = waState.pendingAcks.get(key.id);
                waState.pendingAcks.delete(key.id);
                cb(update.status);
            }
        }
    });
    sock.ev.on('message-receipt.update', (updates) => {
        for (const u of updates) {
            const id = u?.key?.id;
            if (id && waState.pendingAcks.has(id)) {
                const cb = waState.pendingAcks.get(id);
                waState.pendingAcks.delete(id);
                cb(3);
            }
        }
    });
}
function waitForOpenConnection(waState, timeoutMs = 90000) {
    if (waState.isOpen && waState.sock) return Promise.resolve(true);
    return new Promise((resolve) => {
        const start = Date.now();
        const interval = setInterval(() => {
            if (waState.isOpen && waState.sock) { clearInterval(interval); resolve(true); }
            else if (Date.now() - start > timeoutMs) { clearInterval(interval); resolve(false); }
        }, 1000);
    });
}
async function sendMessageConfirmed(waState, jid, content, options = {}, ackLogTimeoutMs = BOT_CONFIG.ACK_LOG_TIMEOUT_MS) {
    const sock = waState.sock;
    if (!sock) return { success: false, id: null, reason: 'no-socket' };
    let sendResult;
    try { sendResult = await sock.sendMessage(jid, content, options); }
    catch (err) { return { success: false, id: null, reason: 'send-error', error: err.message }; }
    const id = sendResult?.key?.id;
    if (!id) return { success: false, id: null, reason: 'no-id' };
    waState.recentMessages.set(id, { key: sendResult.key, message: sendResult.message || content });
    if (waState.recentMessages.size > 200) {
        const oldest = waState.recentMessages.keys().next().value;
        waState.recentMessages.delete(oldest);
    }
    let ackStatus = null;
    try {
        ackStatus = await new Promise((resolve) => {
            const timer = setTimeout(() => { waState.pendingAcks.delete(id); resolve(null); }, ackLogTimeoutMs);
            waState.pendingAcks.set(id, (status) => { clearTimeout(timer); resolve(status); });
        });
    } catch {}
    return { success: true, id, ackStatus };
}
function noteReconnectAndMaybeCooldown(waState) {
    const now = Date.now();
    waState.reconnectLog.push(now);
    waState.reconnectLog = waState.reconnectLog.filter(t => now - t < 10 * 60 * 1000);
    if (waState.reconnectLog.length >= 4) return 10 * 60 * 1000;
    return 0;
}

// ============================================================
// KONEKSI WHATSAPP
// ============================================================
async function connectToWhatsApp(waState) {
    await connectMongo();
    let state = await loadSessionFromMongo();
    if (state && state.creds && state.keys) console.log(chalk.green('✅ Session valid'));
    else { console.log(chalk.yellow('ℹ️ Tidak ada session, minta pairing/QR...')); state = { creds: {}, keys: {} }; }

    const { default: makeWASocket, DisconnectReason, fetchLatestBaileysVersion, initAuthCreds, makeCacheableSignalKeyStore } = await getBaileys();
    if (!state.creds || !state.creds.noiseKey) state = { creds: initAuthCreds(), keys: state.keys || {} };
    const { version } = await fetchLatestBaileysVersion();
    console.log(chalk.gray(`ℹ️ Baileys v${version.join('.')}`));

    const logger = pino({ level: DEBUG_WA ? 'debug' : 'silent' });
    const keysData = state.keys || {};
    let latestCreds = state.creds;
    const debouncedSaveKeys = makeDebouncedSaver(() => latestCreds, keysData);
    let pendingCredsSave = Promise.resolve();

    const sock = makeWASocket({
        logger,
        auth: { creds: state.creds, keys: makeCacheableSignalKeyStore(makeMongoKeyStore(keysData, debouncedSaveKeys), logger) },
        version, browser: ['Ubuntu', 'Chrome', '20.0.00'],
        connectTimeoutMs: 60000, qrTimeout: 60000, markOnlineOnConnect: true,
        emitOwnEvents: true, fireInitQueries: true, generateHighQualityLinkPreview: true,
        syncFullHistory: false, retryRequestDelayMs: 1000,
        getMessage: makeGetMessage(waState), msgRetryCounterCache: makeSimpleCacheStore(),
    });
    waState.sock = sock;
    waState.isOpen = false;
    attachAckTracking(sock, waState);

    let reconnectAttempts = 0;
    let pairingRequested = false;

    sock.ev.on('creds.update', async () => {
        if (waState.sock !== sock) return;
        if (sock.authState && sock.authState.creds) {
            latestCreds = sock.authState.creds;
            pendingCredsSave = saveSessionToMongo(sock.authState.creds, keysData);
            await pendingCredsSave;
        }
    });

    sock.ev.on('connection.update', async (update) => {
        if (waState.sock !== sock) return;
        const { qr } = update;
        if (qr && PHONE_NUMBER && !pairingRequested && !sock.authState.creds.registered) {
            pairingRequested = true;
            try {
                console.log(chalk.cyan('🔑 Meminta pairing code...'));
                const code = await sock.requestPairingCode(PHONE_NUMBER.trim());
                console.log(chalk.bold.yellow(`   ${code}\n`));
            } catch (err) { console.log(chalk.yellow(`⚠️ Pairing: ${err.message}`)); pairingRequested = false; }
        } else if (qr && !PHONE_NUMBER) {
            console.log(chalk.yellow('\n📱 SCAN QR CODE:'));
            qrcode.generate(qr, { small: true });
        }
    });

    sock.ev.on('connection.update', async (update) => {
        if (waState.sock !== sock) return;
        const { connection, lastDisconnect } = update;
        if (connection === 'close') {
            waState.isOpen = false;
            const statusCode = lastDisconnect?.error?.output?.statusCode;
            if (statusCode === DisconnectReason.loggedOut) {
                console.log(chalk.red('❌ Logged out.')); await deleteSessionFromMongo(); process.exit(1);
            } else if (statusCode === DisconnectReason.badSession) {
                console.log(chalk.red('❌ Bad session.'));
                await deleteSessionFromMongo();
                const cooldown = noteReconnectAndMaybeCooldown(waState);
                await pendingCredsSave.catch(() => {});
                setTimeout(() => connectToWhatsApp(waState).catch(console.error), cooldown || 3000);
            } else {
                reconnectAttempts++;
                if (reconnectAttempts > 10) { console.log(chalk.red('❌ Keluar.')); process.exit(1); }
                const waitMs = noteReconnectAndMaybeCooldown(waState) || 3000;
                console.log(chalk.yellow(`🔄 Reconnect (${reconnectAttempts}/10) dalam ${Math.round(waitMs/1000)}s...`));
                await pendingCredsSave.catch(() => {});
                setTimeout(() => connectToWhatsApp(waState).catch(console.error), waitMs);
            }
        } else if (connection === 'open') {
            reconnectAttempts = 0;
            waState.isOpen = true;
            console.log(chalk.green('✅ WhatsApp terhubung!'));
            const u = sock.user;
            if (u && u.id) console.log(chalk.cyan(`   👤 ${u.id} (${u.name || 'tanpa nama'})`));
            else console.log(chalk.red('   ⚠️ sock.user kosong!'));
        }
    });

    const opened = await waitForOpenConnection(waState, 180000);
    if (!opened) throw new Error('Timeout koneksi (180 detik)');
    return sock;
}

// ============================================================
// FOLDER INIT
// ============================================================
function initializeFolders() {
    ['output', 'output/final', 'output/preview', 'templates', 'cv', 'temp'].forEach(dir => {
        const p = path.join(__dirname, dir);
        if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true });
    });
}

// ============================================================
// DOKUMEN META
// ============================================================
function resolveDocMeta(filePath) {
    const ext = (path.extname(filePath) || '').toLowerCase();
    if (ext === '.pdf') return { ext: '.pdf', mimetype: 'application/pdf', label: 'PDF', isPdf: true };
    if (ext === '.docx') return { ext: '.docx', mimetype: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', label: 'DOCX', isPdf: false };
    if (ext === '.doc') return { ext: '.doc', mimetype: 'application/msword', label: 'DOC', isPdf: false };
    return { ext: ext || '.pdf', mimetype: 'application/pdf', label: (ext.replace('.', '') || 'PDF').toUpperCase(), isPdf: ext === '.pdf' };
}

function isUsableFile(filePath, minBytes = BOT_CONFIG.MIN_PDF_BYTES) {
    try { return !!filePath && fs.existsSync(filePath) && fs.statSync(filePath).size > minBytes; }
    catch { return false; }
}

function moveToTemp(srcPath) {
    if (!isUsableFile(srcPath)) return null;
    const srcDir = path.resolve(path.dirname(srcPath));
    if (srcDir === path.resolve(TEMP_DIR)) return srcPath;
    try { if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true }); } catch {}
    const destPath = path.join(TEMP_DIR, path.basename(srcPath));
    try { fs.renameSync(srcPath, destPath); return destPath; }
    catch {
        try { fs.copyFileSync(srcPath, destPath); fs.unlinkSync(srcPath); return destPath; }
        catch (e2) { console.log(chalk.yellow(`⚠️ Gagal pindah temp: ${e2.message}`)); return null; }
    }
}

function deleteTempFile(filePath) {
    try {
        if (!filePath) return;
        if (path.resolve(path.dirname(filePath)) !== path.resolve(TEMP_DIR)) return;
        if (fs.existsSync(filePath)) { fs.unlinkSync(filePath); console.log(chalk.gray(`🗑️ Temp: ${path.basename(filePath)}`)); }
    } catch (e) { console.log(chalk.gray(`   (Gagal hapus temp: ${e.message})`)); }
}

// ============================================================
// DETEKSI KONVERTER LIBREOFFICE
// ============================================================
function buildConverterCandidates() {
    const isWin = process.platform === 'win32';
    const isMac = process.platform === 'darwin';
    const candidates = [];
    if (isWin) {
        candidates.push({ cmd: 'soffice.exe', args: ['--version'], kind: 'libreoffice' });
        candidates.push({ cmd: 'soffice', args: ['--version'], kind: 'libreoffice' });
        candidates.push({ cmd: 'libreoffice', args: ['--version'], kind: 'libreoffice' });
    } else {
        candidates.push({ cmd: 'libreoffice', args: ['--version'], kind: 'libreoffice' });
        candidates.push({ cmd: 'soffice', args: ['--version'], kind: 'libreoffice' });
    }
    if (!isWin && !isMac) {
        ['/usr/bin/libreoffice', '/usr/bin/soffice', '/usr/local/bin/libreoffice', '/usr/local/bin/soffice', '/snap/bin/libreoffice'].forEach(c => candidates.push({ cmd: c, args: ['--version'], kind: 'libreoffice' }));
        try {
            const optDir = '/opt';
            if (fs.existsSync(optDir)) {
                for (const name of fs.readdirSync(optDir)) {
                    if (name.toLowerCase().startsWith('libreoffice')) candidates.push({ cmd: path.join(optDir, name, 'program', 'soffice'), args: ['--version'], kind: 'libreoffice' });
                }
            }
        } catch {}
    }
    if (isWin) {
        const roots = [process.env['PROGRAMFILES'], process.env['PROGRAMFILES(X86)'], process.env['ProgramW6432'], 'C:\\Program Files', 'C:\\Program Files (x86)'].filter(Boolean);
        for (const root of roots) {
            for (const sub of ['LibreOffice', 'LibreOffice 7', 'LibreOffice 24', 'LibreOffice 25', 'LibreOffice\\program']) {
                const exe = sub.toLowerCase().endsWith('program') ? path.join(root, sub, 'soffice.exe') : path.join(root, sub, 'program', 'soffice.exe');
                candidates.push({ cmd: exe, args: ['--version'], kind: 'libreoffice' });
            }
        }
    }
    if (isMac) {
        candidates.push({ cmd: '/Applications/LibreOffice.app/Contents/MacOS/soffice', args: ['--version'], kind: 'libreoffice' });
        candidates.push({ cmd: path.join(process.env.HOME || '', 'Applications/LibreOffice.app/Contents/MacOS/soffice'), args: ['--version'], kind: 'libreoffice' });
    }
    const seen = new Set(); const unique = [];
    for (const c of candidates) {
        const key = (isWin ? c.cmd.toLowerCase() : c.cmd);
        if (seen.has(key)) continue;
        seen.add(key);
        unique.push(c);
    }
    return unique;
}

let _converterCache;
async function detectConverter() {
    if (_converterCache !== undefined) return _converterCache;
    for (const c of buildConverterCandidates()) {
        try {
            await execFileAsync(c.cmd, c.args, { timeout: 15000, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
            console.log(chalk.gray(`ℹ️ Konverter: ${c.cmd}`));
            _converterCache = c; return c;
        } catch {}
    }
    console.log(chalk.red('\n❌ Tidak ada konverter DOCX→PDF.\n'));
    _converterCache = null; return null;
}

async function convertDocxToPdf(srcPath, pdfOutputDir) {
    if (!isUsableFile(srcPath)) return null;
    if (path.extname(srcPath).toLowerCase() === '.pdf') return srcPath;
    const outDir = pdfOutputDir || path.dirname(srcPath);
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    const wasInTempBefore = path.resolve(path.dirname(srcPath)) === path.resolve(TEMP_DIR);
    let workingPath = srcPath;
    if (!wasInTempBefore) {
        const moved = moveToTemp(srcPath);
        if (!moved) { console.log(chalk.yellow(`⚠️ Gagal pindah ke temp.`)); return null; }
        workingPath = moved;
    }
    const pdfBasename = path.basename(workingPath, path.extname(workingPath)) + '.pdf';
    const targetPdfPath = path.join(outDir, pdfBasename);
    if (isUsableFile(targetPdfPath)) { deleteTempFile(workingPath); return targetPdfPath; }

    const converter = await detectConverter();
    if (!converter) return workingPath;

    const profileDir = path.join(TEMP_DIR, '.lo_profile');
    if (!fs.existsSync(profileDir)) fs.mkdirSync(profileDir, { recursive: true });
    const profileUri = 'file:///' + profileDir.replace(/\\/g, '/').replace(/^\//, '');
    const args = ['--headless', '--invisible', '--nologo', '--nofirststartwizard', '--norestore', '--nolockcheck', '--nodefault', `-env:UserInstallation=${profileUri}`, '--convert-to', 'pdf', '--outdir', outDir, workingPath];
    try {
        await execFileAsync(converter.cmd, args, { timeout: 90000, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
    } catch (e) { console.log(chalk.yellow(`⚠️ Konversi gagal: ${(e.stderr || e.message || '').toString().trim()}`)); return workingPath; }

    if (isUsableFile(targetPdfPath)) { deleteTempFile(workingPath); return targetPdfPath; }
    return workingPath;
}

// ============================================================
// RASTERIZE PDF → IMAGE-ONLY PDF (canvas rendering)
// ------------------------------------------------------------
// Pakai Ghostscript + pdf-lib untuk embed JPEG hasil render
// Tidak ada teks → tidak bisa di-copy
// ============================================================
async function rasterizePdfToImagePdf(inputPath, outputPath, dpi, format = 'jpeg', quality = 92) {
    if (!HAS_GHOSTSCRIPT) return { success: false, error: 'gs-missing' };
    if (!PDFDocument) return { success: false, error: 'pdf-lib-missing' };
    if (!isUsableFile(inputPath)) return { success: false, error: 'input-missing' };

    const workDir = path.join(TEMP_DIR, `rast_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`);
    fs.mkdirSync(workDir, { recursive: true });

    const isJpeg = format === 'jpeg';
    const device = isJpeg ? 'jpeg' : 'png16m';
    const ext = isJpeg ? 'jpg' : 'png';

    try {
        // ---- STEP 1: rasterize via Ghostscript (canvas) ----
        const gsArgs = [
            '-q', '-dNOPAUSE', '-dBATCH', '-dSAFER',
            `-sDEVICE=${device}`,
            `-r${dpi}`,
            '-dTextAlphaBits=4',           // anti-aliasing teks
            '-dGraphicsAlphaBits=4',        // anti-aliasing gambar
            '-dAutoRotatePages=/None',
            '-dUseCropBox',
            `-sOutputFile=${path.join(workDir, `page-%04d.${ext}`)}`,
        ];
        if (isJpeg) gsArgs.push(`-dJPEGQ=${quality}`);
        gsArgs.push(inputPath);

        await execFileAsync('gs', gsArgs, { timeout: 300000, windowsHide: true });

        const files = fs.readdirSync(workDir).filter(f => f.endsWith(`.${ext}`)).sort();
        if (!files.length) throw new Error('Tidak ada halaman ter-render');

        // ---- STEP 2: embed sebagai image-only PDF ----
        const newPdf = await PDFDocument.create();
        for (const f of files) {
            const imgBytes = fs.readFileSync(path.join(workDir, f));
            const img = isJpeg ? await newPdf.embedJpg(imgBytes) : await newPdf.embedPng(imgBytes);
            const page = newPdf.addPage([img.width, img.height]);
            page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
        }

        const outBytes = await newPdf.save();
        fs.writeFileSync(outputPath, outBytes);

        // cleanup
        for (const f of files) { try { fs.unlinkSync(path.join(workDir, f)); } catch {} }
        try { fs.rmdirSync(workDir); } catch {}

        return { success: true, path: outputPath, pages: files.length };
    } catch (e) {
        try {
            if (fs.existsSync(workDir)) {
                for (const f of fs.readdirSync(workDir)) { try { fs.unlinkSync(path.join(workDir, f)); } catch {} }
                try { fs.rmdirSync(workDir); } catch {}
            }
        } catch {}
        return { success: false, error: e.message };
    }
}

// ============================================================
// WATERMARK DINAMIS (pdf-lib)
// ------------------------------------------------------------
// Ditempel di setiap halaman. Teks berisi info target supaya
// kalau bocor, sumbernya bisa dilacak.
// ============================================================
async function applyWatermark(pdfPath, outputPath, options = {}) {
    if (!PDFDocument || !StandardFonts) return { success: false, error: 'pdf-lib-missing' };
    if (!isUsableFile(pdfPath)) return { success: false, error: 'input-missing' };

    const {
        text = 'CONFIDENTIAL',
        fontSize = 32,
        opacity = 0.10,
        color = [0.55, 0.55, 0.55],
        diagonal = true,
        angle = 45,
    } = options;

    try {
        const bytes = fs.readFileSync(pdfPath);
        const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });
        const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
        const colorRgb = rgbFn(color[0], color[1], color[2]);
        const pages = pdfDoc.getPages();

        for (const page of pages) {
            const { width, height } = page.getSize();
            const textWidth = font.widthOfTextAtSize(text, fontSize);

            if (diagonal) {
                const rot = degreesFn(angle);
                const stepX = textWidth + 90;
                const stepY = fontSize * 5;
                // mulai dari negatif supaya nutup seluruh halaman
                for (let y = -height; y < height * 2; y += stepY) {
                    for (let x = -width; x < width * 2; x += stepX) {
                        page.drawText(text, {
                            x, y, size: fontSize, font,
                            color: colorRgb, opacity, rotate: rot,
                        });
                    }
                }
            } else {
                const textHeight = font.heightAtSize(fontSize);
                page.drawText(text, {
                    x: (width - textWidth) / 2,
                    y: (height - textHeight) / 2,
                    size: fontSize, font, color: colorRgb, opacity,
                });
            }
        }

        const out = await pdfDoc.save();
        fs.writeFileSync(outputPath, out);
        return { success: true, path: outputPath, pages: pages.length };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ============================================================
// GABUNG BEBERAPA PDF → 1 PDF (fixed order)
// ------------------------------------------------------------
// URUTAN TETAP: Surat → CV → Ijazah
// pdf-lib copyPages mempertahankan hyperlink di CV
// ============================================================
async function mergePdfsFixed(paths, outputPath) {
    if (!PDFDocument) return { success: false, reason: 'pdf-lib-missing' };

    const order = ['letter', 'cv', 'ijazah'];
    const sourceMap = {
        letter: paths.letterPath,
        cv: paths.cvPath,
        ijazah: paths.ijazahPath,
    };

    try {
        const merged = await PDFDocument.create();
        const pageCounts = {};
        const addedOrder = [];

        for (const key of order) {
            const srcPath = sourceMap[key];
            if (!isUsableFile(srcPath)) {
                console.log(chalk.gray(`   ℹ️ [${key}] tidak tersedia, dilewati`));
                continue;
            }
            const bytes = fs.readFileSync(srcPath);
            const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
            // copyPages mempertahankan halaman; pdf-lib juga menyalin link annotations
            const pages = await merged.copyPages(doc, doc.getPageIndices());
            pages.forEach(p => merged.addPage(p));
            pageCounts[key] = doc.getPageCount();
            addedOrder.push(key);
        }

        if (!addedOrder.length) return { success: false, reason: 'tidak-ada-dokumen' };

        const out = await merged.save();
        fs.writeFileSync(outputPath, out);
        return { success: true, path: outputPath, pageCounts, addedOrder, totalPages: merged.getPageCount() };
    } catch (e) {
        return { success: false, reason: e.message };
    }
}

// ============================================================
// ENCRYPT PDF — permission: block print, ALLOW copy & annot
// ------------------------------------------------------------
// Bit values: 4=print, 8=modify, 16=copy, 32=annot, 256=forms,
//             512=accessibility, 1024=assemble, 2048=print-hi
// ============================================================
async function encryptPdf(inputPath, outputPath, options = {}) {
    if (!HAS_GHOSTSCRIPT) return { success: false, error: 'gs-missing' };

    const {
        blockPrint = true,
        allowCopy = true,
        allowAnnotation = true,
        ownerPassword = BOT_CONFIG.GS_OWNER_PASSWORD || ('owner-' + Date.now()),
    } = options;

    // Base -3904 = deny all, reserved bits sudah benar
    let permissions = -3904;
    permissions += 512;                                 // accessibility selalu
    if (!blockPrint) permissions += 4 + 2048;           // print + print-hi
    if (allowCopy) permissions += 16;                   // copy/extract
    if (allowAnnotation) permissions += 32;             // annot/link

    try {
        await execFileAsync('gs', [
            '-q', '-dNOPAUSE', '-dBATCH', '-dSAFER',
            '-sDEVICE=pdfwrite',
            '-dCompatibilityLevel=1.7',
            '-dEncryptionR=3',
            '-dKeyLength=128',
            `-sOwnerPassword=${ownerPassword}`,
            '-sUserPassword=',
            `-dPermissions=${permissions}`,
            '-dAutoRotatePages=/None',
            `-sOutputFile=${outputPath}`,
            inputPath,
        ], { timeout: 120000, windowsHide: true });
        return { success: true, path: outputPath, permissions };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ============================================================
// RESOLVE EXISTING LETTER PDF (cache)
// ============================================================
async function resolveExistingLetterPdf(target, safeName) {
    const outputPdfPath = path.join(AI_CONFIG.OUTPUT_DIR, `${target.position} - ${safeName}.pdf`);
    const outputDocxPath = path.join(AI_CONFIG.OUTPUT_DIR, `${target.position} - ${safeName}.docx`);

    if (isUsableFile(target.filePath)) {
        const meta = resolveDocMeta(target.filePath);
        if (meta.isPdf) {
            console.log(chalk.green(`✅ Pakai PDF tersimpan (FILE_PATH): ${path.basename(target.filePath)}`));
            return target.filePath;
        }
        const pdf = await convertDocxToPdf(target.filePath, path.dirname(target.filePath));
        if (pdf) return pdf;
    }
    if (isUsableFile(outputPdfPath)) {
        console.log(chalk.green(`✅ Pakai PDF di output: ${path.basename(outputPdfPath)}`));
        return outputPdfPath;
    }
    if (isUsableFile(outputDocxPath)) {
        const pdf = await convertDocxToPdf(outputDocxPath, AI_CONFIG.OUTPUT_DIR);
        if (pdf) return pdf;
    }
    return null;
}

// ============================================================
// GOOGLE SHEETS
// ============================================================
async function loadTargetsFromGoogleSheets() {
    try {
        const auth = new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
        const sheets = google.sheets({ version: 'v4', auth });
        const SHEET_NAME = 'Sheet1';
        const response = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `${SHEET_NAME}!A:G` });
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

        if (companyIdx === -1 || positionIdx === -1 || contactIdx === -1 || statusIdx === -1) {
            console.error(chalk.red('❌ Header tidak lengkap. Perlu: COMPANY, POSITION, CONTACT, STATUS'));
            return null;
        }

        const allRows = [];
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i] || [];
            allRows.push({
                sheetRow: i + 1,
                company: row[companyIdx]?.trim() || '',
                position: row[positionIdx]?.trim() || '',
                contact: row[contactIdx]?.trim() || '',
                requirements: requirementsIdx !== -1 ? (row[requirementsIdx]?.trim() || '') : '',
                status: row[statusIdx]?.trim().toLowerCase() || '',
                filePath: filePathIdx !== -1 ? (row[filePathIdx]?.trim() || '') : '',
                subject: subjectIdx !== -1 ? (row[subjectIdx]?.trim() || '') : '',
            });
        }

        const sentKeys = new Set();
        for (const r of allRows) {
            if ((r.status === 'sent' || r.status === 'skip') && r.company && r.position) {
                sentKeys.add(`${r.company.toLowerCase()}|||${r.position.toLowerCase()}`);
            }
        }

        const seenInQueue = new Set();
        const rowsToDelete = [];
        const targets = [];

        for (const r of allRows) {
            if (!r.company || !r.position || !r.contact) continue;
            if (r.status === 'sent' || r.status === 'skip') continue;
            const key = `${r.company.toLowerCase()}|||${r.position.toLowerCase()}`;
            if (sentKeys.has(key)) {
                rowsToDelete.push(r.sheetRow);
                console.log(chalk.yellow(`🗑️ [${r.sheetRow}] ${r.company} — ${r.position} → sudah pernah dikirim`));
                continue;
            }
            if (seenInQueue.has(key)) {
                rowsToDelete.push(r.sheetRow);
                console.log(chalk.yellow(`🗑️ [${r.sheetRow}] ${r.company} — ${r.position} → duplikat`));
                continue;
            }
            seenInQueue.add(key);
            targets.push({ company: r.company, position: r.position, contact: r.contact, requirements: r.requirements, filePath: r.filePath, subject: r.subject });
        }

        if (rowsToDelete.length) {
            try {
                await deleteSheetRowsByIndex(sheets, SHEET_NAME, rowsToDelete);
                console.log(chalk.green(`✅ ${rowsToDelete.length} baris duplikat dihapus.`));
            } catch (e) { console.log(chalk.yellow(`⚠️ Gagal hapus: ${e.message}`)); }
        }
        return targets;
    } catch (error) {
        console.error(chalk.red('❌ Gagal baca Sheets:'), error.message);
        return null;
    }
}

async function deleteSheetRowsByIndex(sheets, sheetTitle, rowNumbers) {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID, fields: 'sheets.properties' });
    const sheet = (meta.data.sheets || []).find(s => s.properties?.title === sheetTitle);
    if (!sheet) throw new Error(`Sheet "${sheetTitle}" tidak ditemukan`);
    const sheetId = sheet.properties.sheetId;
    const sortedDesc = [...new Set(rowNumbers)].sort((a, b) => b - a);
    const requests = sortedDesc.map(rowNum => ({
        deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: rowNum - 1, endIndex: rowNum } },
    }));
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, resource: { requests } });
}

async function updateGoogleSheetStatus(company, position, status = 'sent', filePath = '') {
    try {
        const auth = new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
        const sheets = google.sheets({ version: 'v4', auth });
        const response = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Sheet1!A:G' });
        const rows = response.data.values;
        if (!rows) return false;
        const headers = rows[0].map(h => h.trim().toUpperCase());
        const companyIdx = headers.indexOf('COMPANY');
        const positionIdx = headers.indexOf('POSITION');
        const statusIdx = headers.indexOf('STATUS');
        const filePathIdx = headers.indexOf('FILE_PATH');
        let targetRow = -1;
        for (let i = 1; i < rows.length; i++) {
            if ((rows[i]?.[companyIdx]?.trim() || '') === company && (rows[i]?.[positionIdx]?.trim() || '') === position) {
                targetRow = i + 1; break;
            }
        }
        if (targetRow === -1) return false;
        const statusCol = String.fromCharCode(65 + statusIdx);
        await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID, range: `Sheet1!${statusCol}${targetRow}`,
            valueInputOption: 'RAW', resource: { values: [[status]] },
        });
        if (filePath && filePathIdx !== -1) {
            const fpCol = String.fromCharCode(65 + filePathIdx);
            await sheets.spreadsheets.values.update({
                spreadsheetId: SPREADSHEET_ID, range: `Sheet1!${fpCol}${targetRow}`,
                valueInputOption: 'RAW', resource: { values: [[filePath]] },
            });
        }
        return true;
    } catch (error) { console.error(chalk.red('❌ Update Sheets:'), error.message); return false; }
}

// ============================================================
// EMAIL
// ============================================================
async function attemptSendEmail(contact, company, position, cvPath, letterPath, customSubject = '', letterFileName = '') {
    const oAuth2Client = new google.auth.OAuth2(
        process.env.GOOGLE_CLIENT_ID || '407408718192.apps.googleusercontent.com',
        process.env.GOOGLE_CLIENT_SECRET || '',
        'https://developers.google.com/oauthplayground'
    );
    if (!process.env.GOOGLE_REFRESH_TOKEN) { const e = new Error('REFRESH_TOKEN hilang'); e.isConfigError = true; throw e; }
    oAuth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
    try { await oAuth2Client.getAccessToken(); }
    catch { const e = new Error('Refresh token invalid'); e.isConfigError = true; throw e; }

    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            type: 'OAuth2',
            user: process.env.GOOGLE_USER || process.env.EMAIL_USER,
            clientId: process.env.GOOGLE_CLIENT_ID || '407408718192.apps.googleusercontent.com',
            clientSecret: process.env.GOOGLE_CLIENT_SECRET || '',
            refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
        },
    });

    const subject = (customSubject && customSubject.trim()) ? customSubject.trim() : buildDefaultSubject(position);
    const text = `
Yth. HRD ${company},

Perkenalkan, saya Galeh Riski Prasetio. Sehubungan dengan informasi lowongan posisi ${position} di ${company}, saya bermaksud mengajukan lamaran untuk posisi tersebut.

Bersama email ini, saya lampirkan CV dan Surat Lamaran Kerja sebagai bahan pertimbangan awal. Besar harapan saya untuk dapat diberikan kesempatan wawancara.

Terima kasih atas waktu dan perhatian Bapak/Ibu.

Hormat saya,
Galeh Riski Prasetio
0851-9419-4479
    `;

    const attachments = [];
    // CV tidak dilampirkan terpisah — sudah digabung
    if (isUsableFile(letterPath)) {
        const fname = letterFileName && letterFileName.trim() ? letterFileName.trim() : path.basename(letterPath);
        attachments.push({ filename: fname, path: letterPath });
    }

    await transporter.sendMail({
        from: process.env.GOOGLE_USER || process.env.EMAIL_USER,
        to: contact, subject, text, attachments,
    });
}

async function sendEmail(contact, company, position, cvPath, letterPath, customSubject = '', letterFileName = '') {
    let attempt = 0;
    while (true) {
        attempt++;
        try {
            console.log(chalk.gray(`   ⏳ Mengirim email ke ${contact}...`));
            await attemptSendEmail(contact, company, position, cvPath, letterPath, customSubject, letterFileName);
            console.log(chalk.green(`📧 Email terkirim ke ${contact}${attempt > 1 ? ` (percobaan ${attempt})` : ''}`));
            return true;
        } catch (error) {
            console.error(chalk.red(`❌ Gagal kirim email (${attempt}):`), error.message);
            if (error.isConfigError) return false;
            const waitMs = Math.min(5000 * attempt, 5 * 60 * 1000);
            await delay(waitMs);
        }
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

async function sendWhatsApp(waState, contact, company, position, cvPath, letterPath, letterFileName = '', maxRetries = 3) {
    const targetKey = `${contact}|${company}|${position}`;
    if (waState.sentTargets.has(targetKey)) return true;

    const jid = toJid(contact);
    const messageText = `
${getGreetingByTime()}, Bapak/Ibu HRD ${company},

Perkenalkan, saya Galeh Riski Prasetio. Sehubungan dengan informasi lowongan posisi ${position} di ${company}, saya bermaksud mengajukan lamaran untuk posisi tersebut.

Bersama pesan ini, saya lampirkan CV dan Surat Lamaran Kerja sebagai bahan pertimbangan awal. Besar harapan saya untuk dapat diberikan kesempatan wawancara.

Terima kasih atas waktu dan perhatian Bapak/Ibu.

Hormat saya,
Galeh Riski Prasetio
0851-9419-4479
    `;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        const connected = await waitForOpenConnection(waState, 90000);
        if (!connected) { if (attempt < maxRetries) await delay(5000 * attempt); continue; }
        const sock = waState.sock;
        try {
            let targetJid = jid;
            try {
                const [result] = await sock.onWhatsApp(jid);
                if (!result || !result.exists) { console.log(chalk.yellow(`⚠️ Nomor ${contact} tidak terdaftar.`)); return false; }
                if (result.jid && result.jid !== jid) targetJid = result.jid;
            } catch {}

            const textResult = await sendMessageConfirmed(waState, targetJid, { text: messageText });
            if (!textResult.success) throw new Error(`Teks gagal (${textResult.reason})`);
            await delay(jitterDelay(BOT_CONFIG.DELAY_BETWEEN_SENDS));

            if (isUsableFile(letterPath)) {
                const letterBuffer = fs.readFileSync(letterPath);
                const letterMeta = resolveDocMeta(letterPath);
                const fname = letterFileName && letterFileName.trim() ? letterFileName.trim() : `Surat-${company}${letterMeta.ext}`;
                const letterResult = await sendMessageConfirmed(waState, targetJid, {
                    document: letterBuffer,
                    mimetype: letterMeta.mimetype,
                    fileName: fname,
                    caption: `Lamaran - ${company} (${position})`,
                });
                if (!letterResult.success) throw new Error(`Dokumen gagal (${letterResult.reason})`);
                await delay(jitterDelay(BOT_CONFIG.DELAY_BETWEEN_SENDS));
            }

            console.log(chalk.green(`📱 WhatsApp terkirim ke ${contact}`));
            waState.sentTargets.add(targetKey);
            return true;
        } catch (error) {
            console.error(chalk.red(`❌ WA gagal (${attempt}/${maxRetries}):`), error.message);
            if (attempt < maxRetries) await delay(5000 * attempt);
        }
    }
    return false;
}

// ============================================================
// HELPER
// ============================================================
function getGreetingByTime() {
    const wibHour = parseInt(new Date().toLocaleString('en-US', { timeZone: 'Asia/Jakarta', hour: '2-digit', hour12: false }), 10);
    if (wibHour >= 5 && wibHour < 11) return 'Selamat pagi';
    if (wibHour >= 11 && wibHour < 15) return 'Selamat siang';
    if (wibHour >= 15 && wibHour < 18) return 'Selamat sore';
    return 'Selamat malam';
}
function delay(ms) { return new Promise(r => setTimeout(r, ms)); }
function jitterDelay(baseMs) { const v = baseMs * 0.3; return Math.round(baseMs - v + Math.random() * (v * 2)); }
function isEmail(contact) { return contact && contact.includes('@'); }

function buildOutputFileName(position) {
    const safeApplicant = sanitizeName(BOT_CONFIG.APPLICANT_NAME || 'Pelamar');
    const safePosition = sanitizeName(position || 'Posisi');
    return `${safeApplicant}_${safePosition}.pdf`;
}

function buildWatermarkText(target, trackingId) {
    // Komposisi watermark: perusahaan | kontak | timestamp | trackingID
    const parts = [
        BOT_CONFIG.APPLICANT_NAME,
        target.company,
        target.contact,
        new Date().toISOString().replace('T', ' ').substring(0, 19),
        trackingId || '',
    ].filter(Boolean);
    return parts.join(' | ');
}

// ============================================================
// GENERATE / AMBIL SURAT
// ============================================================
async function getOrCreateLetterPdf(target, safeName) {
    const existing = await resolveExistingLetterPdf(target, safeName);
    if (existing) return { pdfPath: existing, fromCache: true };

    const outputPdfPath = path.join(AI_CONFIG.OUTPUT_DIR, `${target.position} - ${safeName}.pdf`);
    const spinner = ora('Membuat surat via AI...').start();

    let result;
    try {
        result = await generateApplicationLetter(target.company, target.position, outputPdfPath, target.requirements || '');
    } catch (error) { spinner.fail('Error: ' + error.message); return null; }

    if (!result || !result.success || !result.actualPath) { spinner.fail('Gagal generate'); return null; }

    let genPath = result.actualPath;
    if (!isUsableFile(genPath)) { spinner.fail(`File kosong: ${genPath}`); return null; }

    const genMeta = resolveDocMeta(genPath);
    if (genMeta.isPdf) { spinner.succeed(`✅ Surat PDF: ${path.basename(genPath)}`); return { pdfPath: genPath, fromCache: false }; }

    spinner.text = `Konversi ${genMeta.label} → PDF...`;
    const pdfPath = await convertDocxToPdf(genPath, AI_CONFIG.OUTPUT_DIR);
    if (!pdfPath) { spinner.fail('Konversi gagal'); return null; }
    spinner.succeed(`✅ Surat PDF: ${path.basename(pdfPath)}`);
    return { pdfPath, fromCache: false };
}

// ============================================================
// MERGE FINAL
// ------------------------------------------------------------
// 1. Prepare tiap komponen:
//    - Surat: rasterize (kecuali FLATTEN_LETTER=false)
//    - CV: biarkan teks (FLATTEN_CV=false) → hyperlink & copy jalan
//    - Ijazah: rasterize
// 2. Merge → Surat → CV → Ijazah
// 3. Watermark dinamis di semua halaman
// 4. Stamp + footer (test-merge.js, tanpa encrypt)
// 5. Encrypt (block print, allow copy)
// ============================================================
async function getOrCreateMergedPdf(letterPath, target, safeName) {
    const letterDir = path.resolve(path.dirname(letterPath));
    const officialOutputDir = path.resolve(path.join(__dirname, 'output'));

    if (letterDir === path.resolve(FINAL_DIR)) {
        console.log(chalk.gray(`   💾 Surat sudah merged (final/)`));
        return { path: letterPath, merged: false, fromFinal: true, fileName: path.basename(letterPath) };
    }
    if (!BOT_CONFIG.MERGE_WITH_IJAZAH) {
        return { path: letterPath, merged: false, fromFinal: false, fileName: path.basename(letterPath) };
    }

    const finalFileName = buildOutputFileName(target.position);
    const finalPath = path.join(FINAL_DIR, finalFileName);
    const cvPath = BOT_CONFIG.CV_PATH;
    const hasCv = isUsableFile(cvPath);
    const hasIjazah = isUsableFile(IJAZAH_PATH);

    // Prepare workspace
    if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
    const workId = `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;

    // ---- STEP 1: siapkan tiap komponen ----
    let preparedLetter = letterPath;
    let preparedCv = hasCv ? cvPath : null;
    let preparedIjazah = hasIjazah ? IJAZAH_PATH : null;

    // Surat → rasterize
    if (BOT_CONFIG.FLATTEN_LETTER && HAS_GHOSTSCRIPT) {
        const rastPath = path.join(TEMP_DIR, `letter_flat_${workId}.pdf`);
        console.log(chalk.cyan(`   🖼️  Rasterize Surat (DPI ${BOT_CONFIG.FLATTEN_DPI}, ${BOT_CONFIG.FLATTEN_FORMAT.toUpperCase()})...`));
        const res = await rasterizePdfToImagePdf(letterPath, rastPath, BOT_CONFIG.FLATTEN_DPI, BOT_CONFIG.FLATTEN_FORMAT, BOT_CONFIG.FLATTEN_JPEG_QUALITY);
        if (res.success) {
            preparedLetter = rastPath;
            console.log(chalk.gray(`   ↳ Surat: ${res.pages} halaman jadi gambar`));
        } else {
            console.log(chalk.yellow(`   ⚠️ Rasterize surat gagal: ${res.error}`));
        }
    }

    // CV → keep teks (biar hyperlink aktif)
    if (BOT_CONFIG.FLATTEN_CV && HAS_GHOSTSCRIPT) {
        const rastPath = path.join(TEMP_DIR, `cv_flat_${workId}.pdf`);
        console.log(chalk.cyan(`   🖼️  Rasterize CV...`));
        const res = await rasterizePdfToImagePdf(cvPath, rastPath, BOT_CONFIG.FLATTEN_DPI, BOT_CONFIG.FLATTEN_FORMAT, BOT_CONFIG.FLATTEN_JPEG_QUALITY);
        if (res.success) { preparedCv = rastPath; console.log(chalk.gray(`   ↳ CV: ${res.pages} halaman jadi gambar`)); }
    } else if (hasCv) {
        console.log(chalk.cyan(`   📄 CV: dipertahankan sebagai teks (hyperlink & copy aktif)`));
    }

    // Ijazah → rasterize
    if (BOT_CONFIG.FLATTEN_IJAZAH && hasIjazah && HAS_GHOSTSCRIPT) {
        const rastPath = path.join(TEMP_DIR, `ijazah_flat_${workId}.pdf`);
        console.log(chalk.cyan(`   🖼️  Rasterize Ijazah...`));
        const res = await rasterizePdfToImagePdf(IJAZAH_PATH, rastPath, BOT_CONFIG.FLATTEN_DPI, BOT_CONFIG.FLATTEN_FORMAT, BOT_CONFIG.FLATTEN_JPEG_QUALITY);
        if (res.success) { preparedIjazah = rastPath; console.log(chalk.gray(`   ↳ Ijazah: ${res.pages} halaman jadi gambar`)); }
    }

    // ---- STEP 2: merge fixed order: Surat → CV → Ijazah ----
    const mergedPath = path.join(TEMP_DIR, `merged_${workId}.pdf`);
    const mergeRes = await mergePdfsFixed(
        { letterPath: preparedLetter, cvPath: preparedCv, ijazahPath: preparedIjazah },
        mergedPath
    );
    if (!mergeRes.success) {
        console.log(chalk.yellow(`   ⚠️ Merge gagal: ${mergeRes.reason}`));
        if (letterDir === officialOutputDir) { try { fs.unlinkSync(letterPath); } catch {} }
        return { path: letterPath, merged: false, fromFinal: false, error: mergeRes.reason, fileName: path.basename(letterPath) };
    }
    const parts = mergeRes.addedOrder.map(k => `${k}:${mergeRes.pageCounts[k]}h`).join(' + ');
    console.log(chalk.cyan(`   📎 Merge ${parts} → ${mergeRes.totalPages}h`));

    // ---- STEP 3: watermark dinamis ----
    let watermarkedPath = mergedPath;
    let watermarkApplied = false;

    if (BOT_CONFIG.WATERMARK_ENABLED) {
        // Tracking ID & watermark text
        const trackingId = (Date.now().toString(36) + Math.random().toString(36).slice(2, 7)).toUpperCase();
        const wmText = buildWatermarkText(target, trackingId);
        const wmPath = path.join(TEMP_DIR, `wm_${workId}.pdf`);
        console.log(chalk.cyan(`   💧 Watermark: "${wmText.substring(0, 60)}${wmText.length > 60 ? '...' : ''}"`));

        const wmRes = await applyWatermark(mergedPath, wmPath, {
            text: wmText,
            fontSize: BOT_CONFIG.WATERMARK_FONT_SIZE,
            opacity: BOT_CONFIG.WATERMARK_OPACITY,
            color: BOT_CONFIG.WATERMARK_COLOR,
            diagonal: BOT_CONFIG.WATERMARK_DIAGONAL,
            angle: BOT_CONFIG.WATERMARK_ANGLE,
        });

        if (wmRes.success) {
            watermarkedPath = wmPath;
            watermarkApplied = true;
            console.log(chalk.gray(`   ↳ Watermark di ${wmRes.pages} halaman`));
        } else {
            console.log(chalk.yellow(`   ⚠️ Watermark gagal: ${wmRes.error}`));
        }
        try { fs.unlinkSync(mergedPath); } catch {}
    }

    // ---- STEP 4: stamp + footer via test-merge.js ----
    const stampedPath = path.join(TEMP_DIR, `stamped_${workId}.pdf`);
    console.log(chalk.cyan(`   🔗 Stamp + footer...`));

    let stampResult;
    try {
        stampResult = await mergeAndStamp({
            letterPath: watermarkedPath,
            ijazahPath: null,  // sudah digabung
            companyName: target.company,
            position: target.position,
            outputPath: stampedPath,
            tmpDir: TEMP_DIR,
            stampOnly: TM_STAMP.stampOnly,
            stampCorner: TM_STAMP.corner,
            stampFontSize: TM_STAMP.fontSize,
            stampOpacity: TM_STAMP.opacity,
            stampPadding: TM_STAMP.padding,
            stampColor: TM_STAMP.color,
            stampFontStyle: TM_STAMP.fontStyle,
            stampBg: TM_STAMP.bg,
            stampBgColor: TM_STAMP.bgColor,
            stampBgOpacity: TM_STAMP.bgOpacity,
            stampBgPadding: TM_STAMP.bgPadding,
            stampBorderColor: TM_STAMP.borderColor,
            stampBorderWidth: TM_STAMP.borderWidth,
            stampOutline: TM_STAMP.outline,
            stampOutlineColor: TM_STAMP.outlineColor,
            stampOutlineWidth: TM_STAMP.outlineWidth,
            enableFooter: true,
            enableMetadata: true,
            noEncrypt: true,   // encrypt di step berikutnya
        });
    } catch (e) {
        console.log(chalk.yellow(`   ⚠️ Stamp error: ${e.message}`));
        try { fs.unlinkSync(watermarkedPath); } catch {}
        if (letterDir === officialOutputDir) { try { fs.unlinkSync(letterPath); } catch {} }
        return { path: letterPath, merged: false, fromFinal: false, error: e.message, fileName: path.basename(letterPath) };
    }
    try { fs.unlinkSync(watermarkedPath); } catch {}

    if (!stampResult.success) {
        console.log(chalk.yellow(`   ⚠️ Stamp gagal: ${stampResult.error}`));
        if (letterDir === officialOutputDir) { try { fs.unlinkSync(letterPath); } catch {} }
        return { path: letterPath, merged: false, fromFinal: false, error: stampResult.error, fileName: path.basename(letterPath) };
    }

    // ---- STEP 5: encrypt (block print, allow copy) ----
    let finalUsedPath = stampedPath;
    let encrypted = false;
    let permissionsUsed = null;

    if (HAS_GHOSTSCRIPT) {
        const encRes = await encryptPdf(stampedPath, finalPath, {
            blockPrint: BOT_CONFIG.BLOCK_PRINT,
            allowCopy: BOT_CONFIG.ALLOW_COPY,
            allowAnnotation: BOT_CONFIG.ALLOW_ANNOTATION,
        });
        if (encRes.success) {
            finalUsedPath = finalPath;
            encrypted = true;
            permissionsUsed = encRes.permissions;
            try { fs.unlinkSync(stampedPath); } catch {}
            console.log(chalk.green(
                `   🔒 Encrypt: print=${BOT_CONFIG.BLOCK_PRINT ? 'BLOCKED' : 'OK'}, ` +
                `copy=${BOT_CONFIG.ALLOW_COPY ? 'OK' : 'BLOCKED'}, ` +
                `annot=${BOT_CONFIG.ALLOW_ANNOTATION ? 'OK' : 'BLOCKED'}`
            ));
        } else {
            console.log(chalk.yellow(`   ⚠️ Encrypt gagal: ${encRes.error} — pakai PDF tanpa encrypt`));
            try { if (fs.existsSync(finalPath)) fs.unlinkSync(finalPath); fs.renameSync(stampedPath, finalPath); finalUsedPath = finalPath; } catch {}
        }
    } else {
        try { if (fs.existsSync(finalPath)) fs.unlinkSync(finalPath); fs.renameSync(stampedPath, finalPath); finalUsedPath = finalPath; } catch {}
    }

    // ---- Cleanup temp raster files ----
    for (const p of [preparedLetter, preparedCv, preparedIjazah]) {
        if (p && p !== letterPath && p !== cvPath && p !== IJAZAH_PATH && path.resolve(path.dirname(p)) === path.resolve(TEMP_DIR)) {
            try { if (fs.existsSync(p)) fs.unlinkSync(p); } catch {}
        }
    }

    const tagList = [];
    if (hasCv && !BOT_CONFIG.FLATTEN_CV) tagList.push('CV teks ✓');
    if (hasCv && BOT_CONFIG.FLATTEN_CV) tagList.push('CV gambar');
    if (hasIjazah) tagList.push('Ijazah ✓');
    if (watermarkApplied) tagList.push('💧 WM');
    if (encrypted) tagList.push('🔒');

    console.log(chalk.green(
        `   ✅ Final: ${finalFileName} (${stampResult.totalPages}h${tagList.length ? ', ' + tagList.join(', ') : ''})`
    ));

    if (letterDir === officialOutputDir) {
        try { fs.unlinkSync(letterPath); console.log(chalk.gray(`   🗑️ Surat asli dihapus`)); } catch {}
    }

    return {
        path: finalUsedPath,
        merged: true,
        fromFinal: false,
        fileName: finalFileName,
        trackingId: stampResult.trackingId,
        totalPages: stampResult.totalPages,
        encrypted,
        permissions: permissionsUsed,
        watermarked: watermarkApplied,
    };
}

// ============================================================
// MAIN
// ============================================================
async function main() {
    console.log(chalk.bold.cyan('\n' + '='.repeat(60)));
    console.log(chalk.bold.cyan('📱 WHATSAPP & EMAIL JOB APPLICATION BOT'));
    console.log(chalk.bold.cyan('   PDF urutan: Surat → CV → Ijazah (FIXED)'));
    console.log(chalk.bold.cyan(`   Flatten: Surat=${BOT_CONFIG.FLATTEN_LETTER ? 'IMG' : 'TEXT'} ` +
        `CV=${BOT_CONFIG.FLATTEN_CV ? 'IMG' : 'TEXT'} ` +
        `Ijazah=${BOT_CONFIG.FLATTEN_IJAZAH ? 'IMG' : 'TEXT'}`));
    console.log(chalk.bold.cyan(`   DPI: ${BOT_CONFIG.FLATTEN_DPI} | Format: ${BOT_CONFIG.FLATTEN_FORMAT}`));
    console.log(chalk.bold.cyan(`   Watermark: ${BOT_CONFIG.WATERMARK_ENABLED ? 'AKTIF' : 'off'}`));
    console.log(chalk.bold.cyan(`   Print: ${BOT_CONFIG.BLOCK_PRINT ? 'BLOCKED' : 'OK'} | ` +
        `Copy: ${BOT_CONFIG.ALLOW_COPY ? 'OK' : 'BLOCKED'} | ` +
        `Annot: ${BOT_CONFIG.ALLOW_ANNOTATION ? 'OK' : 'BLOCKED'}`));
    console.log(chalk.bold.cyan(`   Nama file: {Nama}_{Posisi}.pdf`));

    const gsVer = await checkGhostscript();
    if (gsVer) console.log(chalk.bold.green(`   ✅ Ghostscript v${gsVer}`));
    else {
        console.log(chalk.bold.yellow(`   ⚠️  Ghostscript TIDAK ada — rasterize & encrypt dilewati`));
        console.log(chalk.white(`      Install: apt-get install -y ghostscript`));
        if (BOT_CONFIG.REQUIRE_GHOSTSCRIPT) { console.log(chalk.bold.red(`   ❌ REQUIRE_GHOSTSCRIPT=true`)); process.exit(1); }
    }
    if (!PDFDocument) console.log(chalk.bold.yellow('   ⚠️  pdf-lib TIDAK ada!'));
    console.log(chalk.bold.cyan('='.repeat(60) + '\n'));

    initializeFolders();

    if (!fs.existsSync(AI_CONFIG.TEMPLATE_PATH)) { console.log(chalk.red(`❌ Template: ${AI_CONFIG.TEMPLATE_PATH}`)); process.exit(1); }
    if (!fs.existsSync(BOT_CONFIG.CV_PATH)) { console.log(chalk.red(`❌ CV: ${BOT_CONFIG.CV_PATH}`)); process.exit(1); }
    if (!SPREADSHEET_ID) { console.log(chalk.red('❌ SPREADSHEET_ID belum diatur')); process.exit(1); }

    const targets = await loadTargetsFromGoogleSheets();
    if (!targets) { console.log(chalk.red('❌ Gagal load Sheets.')); process.exit(1); }
    if (!targets.length) { console.log(chalk.green('\n✅ Tidak ada target.')); process.exit(0); }

    console.log(chalk.yellow('\n📋 Target:'));
    targets.forEach((t, i) => {
        const channel = isEmail(t.contact) ? '📧 Email' : '📱 WhatsApp';
        const subjInfo = channel.includes('Email') ? ` — subjek: "${(t.subject && t.subject.trim()) || buildDefaultSubject(t.position)}"` : '';
        console.log(`  ${i+1}. ${t.company} - ${t.position} (${channel})`);
        console.log(chalk.gray(`      📄 ${buildOutputFileName(t.position)}${subjInfo}`));
    });

    const hasWhatsApp = targets.some(t => !isEmail(t.contact));
    let waState = null;
    if (hasWhatsApp) {
        console.log(chalk.cyan('\n📱 Menghubungkan WhatsApp...'));
        waState = createWaState();
        await connectToWhatsApp(waState);
        console.log(chalk.green(`✅ WhatsApp siap, tunggu ${BOT_CONFIG.WA_SYNC_DELAY_MS / 1000}s...`));
        await delay(BOT_CONFIG.WA_SYNC_DELAY_MS);
    }

    let successCount = 0, failCount = 0, cacheHitCount = 0, aiCallCount = 0;
    let mergedCount = 0, encryptedCount = 0, watermarkedCount = 0;
    const startTime = Date.now();

    for (let i = 0; i < targets.length; i++) {
        const target = targets[i];
        console.log(chalk.cyan(`\n[${i+1}/${targets.length}] ${target.company} — ${target.position}`));
        const safeName = sanitizeName(target.company);

        const letterInfo = await getOrCreateLetterPdf(target, safeName);
        if (!letterInfo) { await updateGoogleSheetStatus(target.company, target.position, 'failed'); failCount++; continue; }
        let letterPath = letterInfo.pdfPath;
        if (letterInfo.fromCache) { cacheHitCount++; console.log(chalk.gray(`   💾 Cache hit`)); }
        else { aiCallCount++; }

        const mergedInfo = await getOrCreateMergedPdf(letterPath, target, safeName);
        letterPath = mergedInfo.path;
        const letterFileName = mergedInfo.fileName || path.basename(letterPath);

        if (mergedInfo.merged) {
            mergedCount++;
            if (mergedInfo.encrypted) encryptedCount++;
            if (mergedInfo.watermarked) watermarkedCount++;
        }

        let sendSuccess = false;
        if (isEmail(target.contact)) {
            const subjUsed = (target.subject && target.subject.trim()) ? target.subject.trim() : buildDefaultSubject(target.position);
            console.log(chalk.gray(`   ✉️  Subjek: "${subjUsed}"${(target.subject && target.subject.trim()) ? ' (custom)' : ' (default)'}`));
            console.log(chalk.gray(`   📄 File  : ${letterFileName}`));
            sendSuccess = await sendEmail(target.contact, target.company, target.position, BOT_CONFIG.CV_PATH, letterPath, target.subject, letterFileName);
        } else if (waState) {
            console.log(chalk.gray(`   📄 File  : ${letterFileName}`));
            sendSuccess = await sendWhatsApp(waState, target.contact, target.company, target.position, BOT_CONFIG.CV_PATH, letterPath, letterFileName);
        }

        if (sendSuccess) { await updateGoogleSheetStatus(target.company, target.position, 'sent', letterPath); successCount++; }
        else { await updateGoogleSheetStatus(target.company, target.position, 'failed', letterPath); failCount++; }

        if (i < targets.length - 1) {
            const dt = isEmail(target.contact)
                ? BOT_CONFIG.DELAY_BETWEEN_TARGETS_EMAIL_MIN + Math.floor(Math.random() * (BOT_CONFIG.DELAY_BETWEEN_TARGETS_EMAIL_MAX - BOT_CONFIG.DELAY_BETWEEN_TARGETS_EMAIL_MIN))
                : BOT_CONFIG.DELAY_BETWEEN_TARGETS_WA_MIN + Math.floor(Math.random() * (BOT_CONFIG.DELAY_BETWEEN_TARGETS_WA_MAX - BOT_CONFIG.DELAY_BETWEEN_TARGETS_WA_MIN));
            console.log(chalk.gray(`⏳ Delay ${dt/1000}s...`));
            await delay(dt);
        }
    }

    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(chalk.bold.green('\n📊 SUMMARY'));
    console.log(chalk.green(`✅ Berhasil: ${successCount}`));
    console.log(chalk.red(`❌ Gagal: ${failCount}`));
    console.log(chalk.cyan(`💾 Cache hit: ${cacheHitCount}`));
    console.log(chalk.cyan(`🤖 AI calls: ${aiCallCount}`));
    console.log(chalk.cyan(`🔗 Merged: ${mergedCount}`));
    console.log(chalk.cyan(`💧 Watermarked: ${watermarkedCount}`));
    console.log(chalk.cyan(`🔒 Encrypted: ${encryptedCount}`));
    console.log(chalk.cyan(`⏱️  ${elapsed} detik`));

    process.exit(0);
}

main().catch(error => {
    console.error(chalk.red('\n❌ Error:'), error.message);
    process.exit(1);
});
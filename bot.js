// bot.js - WhatsApp & Email Job Application Bot (Baileys + MongoDB)
// Satu file: WhatsApp + Google Sheets + Email
// Session dari MongoDB (collection: sessionschemas)
//
// Alur:
//   1. Ambil target dari Sheet
//   2. Generate surat via lamaran-ai.js (cache kalau sudah ada)
//   3. Gabung Surat → CV → Ijazah (fixed order) jadi 1 PDF plain
//   4. Merge + stamp + footer via test-merge.js
//   5. FLATTEN ke gambar + proteksi (no print, no copy) via Ghostscript
//   6. Kirim dengan nama file: {Nama}_{Posisi}.pdf
//   7. Update Sheet: STATUS + FILE_PATH
//
// Fitur:
//   - Custom SUBJECT email dari kolom SUBJECT di spreadsheet
//   - Gabung CV + Surat + Ijazah jadi 1 file (pdf-lib)
//   - Nama file output: {APPLICANT_NAME}_{POSITION}.pdf
//   - Urutan: Surat → CV → Ijazah (fixed)
//   - Flatten ke gambar + blokir print & copy (Ghostscript)
//   - Self-test message ke diri sendiri dihapus

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

// pdf-lib untuk gabung dokumen
let PDFDocument = null;
try {
    PDFDocument = require('pdf-lib').PDFDocument;
} catch {
    console.log('⚠️ pdf-lib tidak terinstall — fitur gabung dokumen akan dilewati.');
    console.log('   Install dengan: npm install pdf-lib');
}

const execFileAsync = promisify(execFile);

// ============================================================
// FORCE IPv4-FIRST
// ============================================================
try {
    dns.setDefaultResultOrder('ipv4first');
} catch { /* abaikan */ }

dotenv.config();

// ============================================================
// FILTER LOG NOISE
// ============================================================
(function installConsoleNoiseFilter() {
    const NOISE_SUBSTRINGS = [
        'Closing session', 'Opening session', 'Session already closed',
        'SessionEntry', '"class":"baileys"', 'failed to decrypt message',
        'sent retry receipt', 'Bad MAC', 'Closing open session',
    ];
    const isNoiseText = (s) => {
        if (typeof s !== 'string' || s.length === 0) return false;
        for (const p of NOISE_SUBSTRINGS) if (s.includes(p)) return true;
        return false;
    };
    const isNoiseArgs = (args) => {
        if (!args || args.length === 0) return false;
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
            } catch { /* lanjut */ }
            if (isNoiseText(s)) {
                if (typeof encoding === 'function') { encoding(); return true; }
                if (typeof cb === 'function') { cb(); return true; }
                return true;
            }
            return orig(chunk, encoding, cb);
        };
    };
    try { wrapStream(process.stdout); } catch { /* abaikan */ }
    try { wrapStream(process.stderr); } catch { /* abaikan */ }
})();

// ============================================================
// MODE DEBUG
// ============================================================
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
// IMPORT lamaran-ai.js
// ============================================================
const {
    generateApplicationLetter,
    CONFIG: AI_CONFIG
} = require('./lamaran-ai.js');

// ============================================================
// IMPORT test-merge.js
// ============================================================
const {
    mergeAndStamp,
    sanitizeName,
    isUsableFile: _tmIsUsableFile,
    FINAL_DIR,
    IJAZAH_PATH,
    STAMP: TM_STAMP,
} = require('./test-merge.js');

// ============================================================
// KONFIGURASI BOT
// ============================================================
const BOT_CONFIG = {
    // Nama pelamar (dipakai untuk nama file: {Nama}_{Posisi}.pdf)
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

    // ---- Integrasi merge ----
    MERGE_WITH_IJAZAH: true,
    FINAL_DIR: FINAL_DIR,
    IJAZAH_PATH: IJAZAH_PATH,

    // ---- PROTEKSI (BARU) ----
    // Semua halaman jadi GAMBAR → tidak bisa copy teks sama sekali
    FLATTEN_TO_IMAGE: true,

    // Blokir print
    BLOCK_PRINT: true,

    // Blokir copy
    BLOCK_COPY: true,

    // Resolusi rasterize (DPI). 150 = kualitas bagus, ukuran wajar.
    // 200+ = lebih tajam tapi file lebih besar. 100 = lebih kecil.
    FLATTEN_DPI: 150,

    // Owner password untuk PDF (bisa dikosongkan)
    // Kalau diisi, perubahan permission harus pakai password ini
    GS_OWNER_PASSWORD: process.env.GS_OWNER_PASSWORD || '',

    // Kalau Ghostscript tidak ada, tetap kirim PDF normal (tanpa flatten)
    // Set true untuk force error kalau Ghostscript tidak ada
    REQUIRE_GHOSTSCRIPT: false,
};

const TEMP_DIR = path.join(__dirname, 'temp');

// ============================================================
// DEFAULT SUBJECT EMAIL
// ============================================================
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
        const ver = (stdout || '').trim();
        HAS_GHOSTSCRIPT = true;
        return ver;
    } catch {
        HAS_GHOSTSCRIPT = false;
        return null;
    }
}

// ============================================================
// FUNGSI MONGODB
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

// ============================================================
// IMPORT BAILEYS
// ============================================================
let _baileysModule = null;
async function getBaileys() {
    if (!_baileysModule) {
        _baileysModule = await import('@whiskeysockets/baileys');
    }
    return _baileysModule;
}

// ============================================================
// PERBAIKAN BUFFER
// ============================================================
function looksLikeByteArrayObject(val) {
    if (!val || typeof val !== 'object' || Array.isArray(val) || Buffer.isBuffer(val)) return false;
    const keys = Object.keys(val);
    if (keys.length === 0) return false;
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
        if (!doc) {
            console.log(chalk.yellow('⚠️ Session tidak ditemukan di MongoDB'));
            return null;
        }
        console.log(chalk.green('✅ Dokumen session ditemukan'));

        if (doc.session) {
            console.log(chalk.green('✅ Menggunakan field session (string)'));
            try {
                const { BufferJSON } = await getBaileys();
                const parsed = JSON.parse(doc.session, BufferJSON.reviver);
                if (parsed.creds && parsed.keys) {
                    const repairedCount = { n: 0 };
                    const repaired = deepRepairBuffers(parsed, repairedCount);
                    if (repairedCount.n > 0) {
                        console.log(chalk.yellow(`🔧 Memperbaiki ${repairedCount.n} field Buffer dari data lama...`));
                        await saveSessionToMongo(repaired.creds, repaired.keys);
                    }
                    return repaired;
                }
            } catch (e) {
                console.error(chalk.red('❌ Gagal parse session:'), e.message);
            }
        }
        if (doc.state && doc.state.creds && doc.state.keys) {
            console.log(chalk.yellow('⚠️ Fallback ke field state (berisiko).'));
            const repairedCount = { n: 0 };
            const repaired = deepRepairBuffers(doc.state.toObject ? doc.state.toObject() : doc.state, repairedCount);
            if (repairedCount.n > 0) {
                console.log(chalk.yellow(`🔧 Memperbaiki ${repairedCount.n} field Buffer dari field state...`));
            }
            return repaired;
        }
        return null;
    } catch (err) {
        console.error(chalk.red('❌ Gagal load session:'), err.message);
        return null;
    }
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
            {
                sessionId: SESSION_ID,
                state: state,
                session: JSON.stringify(state, BufferJSON.replacer),
                updatedAt: new Date()
            },
            { upsert: true }
        );
        console.log(chalk.gray('💾 Session disimpan ke MongoDB'));
    } catch (err) {
        console.error(chalk.yellow('⚠️ Gagal simpan session:'), err.message);
    }
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
        console.log(chalk.gray('🗑️ Session dihapus dari MongoDB'));
    } catch (err) {
        console.error(chalk.yellow('⚠️ Gagal hapus session:'), err.message);
    }
}

// ============================================================
// STATE KONEKSI WA
// ============================================================
function createWaState() {
    return {
        sock: null,
        isOpen: false,
        pendingAcks: new Map(),
        recentMessages: new Map(),
        reconnectLog: [],
        sentTargets: new Set(),
    };
}

function makeGetMessage(waState) {
    return async (key) => {
        const entry = waState.recentMessages.get(key.id);
        return entry ? entry.message : undefined;
    };
}

function makeSimpleCacheStore() {
    const store = new Map();
    return {
        get: (k) => store.get(k),
        set: (k, v) => { store.set(k, v); return true; },
        del: (k) => store.delete(k),
        flushAll: () => store.clear(),
    };
}

// ============================================================
// ACK TRACKING
// ============================================================
function attachAckTracking(sock, waState) {
    sock.ev.on('messages.update', (updates) => {
        for (const { key, update } of updates) {
            if (key?.fromMe && typeof update?.status === 'number') {
                console.log(chalk.gray(`   [ACK] id=${key.id} status=${update.status}`));
            }
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
            if (waState.isOpen && waState.sock) {
                clearInterval(interval);
                resolve(true);
            } else if (Date.now() - start > timeoutMs) {
                clearInterval(interval);
                resolve(false);
            }
        }, 1000);
    });
}

// ============================================================
// KIRIM PESAN
// ============================================================
async function sendMessageConfirmed(waState, jid, content, options = {}, ackLogTimeoutMs = BOT_CONFIG.ACK_LOG_TIMEOUT_MS) {
    const sock = waState.sock;
    if (!sock) return { success: false, id: null, reason: 'no-socket' };

    let sendResult;
    try {
        sendResult = await sock.sendMessage(jid, content, options);
    } catch (err) {
        return { success: false, id: null, reason: 'send-error', error: err.message };
    }

    const id = sendResult?.key?.id;
    if (!id) return { success: false, id: null, reason: 'no-id' };

    waState.recentMessages.set(id, { key: sendResult.key, message: sendResult.message || content });
    if (waState.recentMessages.size > 200) {
        const oldestKey = waState.recentMessages.keys().next().value;
        waState.recentMessages.delete(oldestKey);
    }

    let ackStatus = null;
    try {
        ackStatus = await new Promise((resolve) => {
            const timer = setTimeout(() => {
                waState.pendingAcks.delete(id);
                resolve(null);
            }, ackLogTimeoutMs);
            waState.pendingAcks.set(id, (status) => {
                clearTimeout(timer);
                resolve(status);
            });
        });
    } catch { /* abaikan */ }

    return { success: true, id, ackStatus };
}

function noteReconnectAndMaybeCooldown(waState) {
    const now = Date.now();
    waState.reconnectLog.push(now);
    waState.reconnectLog = waState.reconnectLog.filter(t => now - t < 10 * 60 * 1000);
    if (waState.reconnectLog.length >= 4) {
        const cooldownMs = 10 * 60 * 1000;
        console.log(chalk.red(`🛑 Reconnect ${waState.reconnectLog.length}x dalam 10 menit. Jeda ${cooldownMs / 60000} menit...`));
        return cooldownMs;
    }
    return 0;
}

// ============================================================
// KONEKSI WHATSAPP
// ============================================================
async function connectToWhatsApp(waState) {
    await connectMongo();

    let state = await loadSessionFromMongo();

    if (state && state.creds && state.keys) {
        console.log(chalk.green('✅ Session valid, langsung terhubung...'));
    } else {
        console.log(chalk.yellow('ℹ️ Tidak ada session valid, minta pairing code/QR...'));
        state = { creds: {}, keys: {} };
    }

    const {
        default: makeWASocket,
        DisconnectReason,
        fetchLatestBaileysVersion,
        initAuthCreds,
        makeCacheableSignalKeyStore
    } = await getBaileys();

    if (!state.creds || !state.creds.noiseKey) {
        state = { creds: initAuthCreds(), keys: state.keys || {} };
    }

    const { version } = await fetchLatestBaileysVersion();
    console.log(chalk.gray(`ℹ️ Baileys v${version.join('.')}`));

    const logger = pino({ level: DEBUG_WA ? 'debug' : 'silent' });
    const keysData = state.keys || {};

    let latestCreds = state.creds;
    const debouncedSaveKeys = makeDebouncedSaver(() => latestCreds, keysData);
    let pendingCredsSave = Promise.resolve();

    const sock = makeWASocket({
        logger,
        auth: {
            creds: state.creds,
            keys: makeCacheableSignalKeyStore(makeMongoKeyStore(keysData, debouncedSaveKeys), logger)
        },
        version,
        browser: ['Ubuntu', 'Chrome', '20.0.00'],
        connectTimeoutMs: 60000,
        qrTimeout: 60000,
        markOnlineOnConnect: true,
        emitOwnEvents: true,
        fireInitQueries: true,
        generateHighQualityLinkPreview: true,
        syncFullHistory: false,
        retryRequestDelayMs: 1000,
        getMessage: makeGetMessage(waState),
        msgRetryCounterCache: makeSimpleCacheStore(),
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
                console.log(chalk.green('\n✅ Pairing code:'));
                console.log(chalk.bold.yellow(`   ${code}\n`));
                console.log(chalk.white('Masukkan di HP: Pengaturan → Perangkat Tertaut → Tautkan dengan nomor telepon'));
            } catch (err) {
                console.log(chalk.yellow(`⚠️ Gagal pairing: ${err.message}`));
                pairingRequested = false;
            }
        } else if (qr && !PHONE_NUMBER) {
            console.log(chalk.yellow('\n📱 SCAN QR CODE:'));
            qrcode.generate(qr, { small: true });
            console.log(chalk.cyan('Scan QR di atas dengan WhatsApp Anda.'));
        }
    });

    sock.ev.on('connection.update', async (update) => {
        if (waState.sock !== sock) return;

        const { connection, lastDisconnect } = update;

        if (connection === 'close') {
            waState.isOpen = false;
            const statusCode = lastDisconnect?.error?.output?.statusCode;

            if (statusCode === DisconnectReason.loggedOut) {
                console.log(chalk.red('❌ Logged out. Hapus session...'));
                await deleteSessionFromMongo();
                process.exit(1);
            } else if (statusCode === DisconnectReason.badSession) {
                console.log(chalk.red('❌ Session tidak valid. Hapus...'));
                await deleteSessionFromMongo();
                const cooldown = noteReconnectAndMaybeCooldown(waState);
                await pendingCredsSave.catch(() => {});
                setTimeout(() => connectToWhatsApp(waState).catch(console.error), cooldown || 3000);
            } else {
                reconnectAttempts++;
                if (reconnectAttempts > 10) {
                    console.log(chalk.red('❌ Terlalu banyak reconnect. Keluar.'));
                    process.exit(1);
                }
                const reason = statusCode === DisconnectReason.restartRequired
                    ? 'restart diperlukan (normal setelah pairing/QR)'
                    : `status code ${statusCode || 'tidak diketahui'}`;
                const cooldown = noteReconnectAndMaybeCooldown(waState);
                const waitMs = cooldown || 3000;
                console.log(chalk.yellow(`🔄 Koneksi tertutup (${reason}). Reconnect (${reconnectAttempts}/10) dalam ${Math.round(waitMs / 1000)} detik...`));

                await pendingCredsSave.catch(() => {});
                setTimeout(() => connectToWhatsApp(waState).catch(console.error), waitMs);
            }
        } else if (connection === 'open') {
            reconnectAttempts = 0;
            waState.isOpen = true;
            console.log(chalk.green('✅ WhatsApp Baileys terhubung!'));

            const userInfo = sock.user;
            if (userInfo && userInfo.id) {
                console.log(chalk.cyan(`   👤 Login sebagai: ${userInfo.id} (${userInfo.name || 'tanpa nama'})`));
            } else {
                console.log(chalk.red('   ⚠️ PERINGATAN: sock.user kosong! Sesi kemungkinan tidak valid.'));
            }
        }
    });

    const opened = await waitForOpenConnection(waState, 180000);
    if (!opened) throw new Error('Timeout menunggu koneksi (180 detik)');

    return sock;
}

// ============================================================
// INISIALISASI FOLDER
// ============================================================
function initializeFolders() {
    const folders = ['output', 'output/final', 'output/preview', 'templates', 'cv', 'temp'];
    folders.forEach(dir => {
        const dirPath = path.join(__dirname, dir);
        if (!fs.existsSync(dirPath)) fs.mkdirSync(dirPath, { recursive: true });
    });
}

// ============================================================
// DOKUMEN & KONVERSI
// ============================================================
function resolveDocMeta(filePath) {
    const ext = (path.extname(filePath) || '').toLowerCase();
    switch (ext) {
        case '.pdf':
            return { ext: '.pdf', mimetype: 'application/pdf', label: 'PDF', isPdf: true };
        case '.docx':
            return {
                ext: '.docx',
                mimetype: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                label: 'DOCX', isPdf: false
            };
        case '.doc':
            return { ext: '.doc', mimetype: 'application/msword', label: 'DOC', isPdf: false };
        default:
            return {
                ext: ext || '.pdf',
                mimetype: 'application/pdf',
                label: (ext.replace('.', '') || 'PDF').toUpperCase(),
                isPdf: ext === '.pdf'
            };
    }
}

function isUsableFile(filePath, minBytes = BOT_CONFIG.MIN_PDF_BYTES) {
    try {
        return !!filePath && fs.existsSync(filePath) && fs.statSync(filePath).size > minBytes;
    } catch {
        return false;
    }
}

function moveToTemp(srcPath) {
    if (!isUsableFile(srcPath)) return null;
    const srcDir = path.resolve(path.dirname(srcPath));
    if (srcDir === path.resolve(TEMP_DIR)) return srcPath;

    try {
        if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
    } catch { /* abaikan */ }

    const destPath = path.join(TEMP_DIR, path.basename(srcPath));
    try {
        fs.renameSync(srcPath, destPath);
        return destPath;
    } catch (e) {
        try {
            fs.copyFileSync(srcPath, destPath);
            fs.unlinkSync(srcPath);
            return destPath;
        } catch (e2) {
            console.log(chalk.yellow(`⚠️ Gagal pindah ke temp: ${path.basename(srcPath)} — ${e2.message}`));
            return null;
        }
    }
}

function deleteTempFile(filePath) {
    try {
        if (!filePath) return;
        const dir = path.resolve(path.dirname(filePath));
        if (dir !== path.resolve(TEMP_DIR)) return;
        if (fs.existsSync(filePath)) {
            fs.unlinkSync(filePath);
            console.log(chalk.gray(`🗑️ Temp dihapus: ${path.basename(filePath)}`));
        }
    } catch (e) {
        console.log(chalk.gray(`   (Gagal hapus temp: ${e.message})`));
    }
}

// ============================================================
// DETEKSI KONVERTER
// ============================================================
function buildConverterCandidates() {
    const isWin = process.platform === 'win32';
    const isMac = process.platform === 'darwin';
    const candidates = [];

    if (isWin) {
        candidates.push({ cmd: 'soffice.exe', args: ['--version'], kind: 'libreoffice' });
        candidates.push({ cmd: 'soffice',     args: ['--version'], kind: 'libreoffice' });
        candidates.push({ cmd: 'libreoffice', args: ['--version'], kind: 'libreoffice' });
    } else {
        candidates.push({ cmd: 'libreoffice', args: ['--version'], kind: 'libreoffice' });
        candidates.push({ cmd: 'soffice',     args: ['--version'], kind: 'libreoffice' });
    }

    if (!isWin && !isMac) {
        candidates.push({ cmd: '/usr/bin/libreoffice', args: ['--version'], kind: 'libreoffice' });
        candidates.push({ cmd: '/usr/bin/soffice',     args: ['--version'], kind: 'libreoffice' });
        candidates.push({ cmd: '/usr/local/bin/libreoffice', args: ['--version'], kind: 'libreoffice' });
        candidates.push({ cmd: '/usr/local/bin/soffice',     args: ['--version'], kind: 'libreoffice' });
        candidates.push({ cmd: '/snap/bin/libreoffice', args: ['--version'], kind: 'libreoffice' });

        try {
            const optDir = '/opt';
            if (fs.existsSync(optDir)) {
                for (const name of fs.readdirSync(optDir)) {
                    if (name.toLowerCase().startsWith('libreoffice')) {
                        candidates.push({
                            cmd: path.join(optDir, name, 'program', 'soffice'),
                            args: ['--version'], kind: 'libreoffice'
                        });
                    }
                }
            }
        } catch { /* abaikan */ }
    }

    if (isWin) {
        const roots = [
            process.env['PROGRAMFILES'], process.env['PROGRAMFILES(X86)'],
            process.env['ProgramW6432'], 'C:\\Program Files', 'C:\\Program Files (x86)',
        ].filter(Boolean);
        const subdirs = [
            'LibreOffice', 'LibreOffice 7', 'LibreOffice 6',
            'LibreOffice 24', 'LibreOffice 25', 'LibreOffice\\program',
        ];
        for (const root of roots) {
            for (const sub of subdirs) {
                const exePath = sub.toLowerCase().endsWith('program')
                    ? path.join(root, sub, 'soffice.exe')
                    : path.join(root, sub, 'program', 'soffice.exe');
                candidates.push({ cmd: exePath, args: ['--version'], kind: 'libreoffice' });
            }
            try {
                if (fs.existsSync(root)) {
                    for (const name of fs.readdirSync(root)) {
                        if (name.toLowerCase().startsWith('libreoffice')) {
                            const exe = path.join(root, name, 'program', 'soffice.exe');
                            candidates.push({ cmd: exe, args: ['--version'], kind: 'libreoffice' });
                        }
                    }
                }
            } catch { /* abaikan */ }
        }
    }

    if (isMac) {
        candidates.push({
            cmd: '/Applications/LibreOffice.app/Contents/MacOS/soffice',
            args: ['--version'], kind: 'libreoffice'
        });
        candidates.push({
            cmd: path.join(process.env.HOME || '', 'Applications/LibreOffice.app/Contents/MacOS/soffice'),
            args: ['--version'], kind: 'libreoffice'
        });
    }

    const seen = new Set();
    const unique = [];
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
    const candidates = buildConverterCandidates();
    const tried = [];

    for (const c of candidates) {
        tried.push(c.cmd);
        try {
            await execFileAsync(c.cmd, c.args, {
                timeout: 15000, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'],
            });
            console.log(chalk.gray(`ℹ️ Konverter terdeteksi: ${c.cmd}`));
            _converterCache = c;
            return c;
        } catch { /* coba berikutnya */ }
    }

    console.log(chalk.red('\n❌ TIDAK ADA KONVERTER DOCX→PDF DI SISTEM INI.'));
    console.log(chalk.yellow(`   Sudah dicoba ${tried.length} lokasi:`));
    for (const t of tried.slice(0, 12)) console.log(chalk.gray(`     • ${t}`));
    if (tried.length > 12) console.log(chalk.gray(`     • ...(${tried.length - 12} lokasi lain)`));
    console.log(chalk.yellow('\n   Install LibreOffice:'));
    console.log(chalk.white('   • Docker: apt-get install -y libreoffice-writer fonts-liberation'));
    console.log(chalk.white('   • Windows: https://www.libreoffice.org/download/download-libreoffice/'));
    console.log(chalk.yellow('\n   Tanpa konverter, DOCX akan dikirim apa adanya.\n'));

    _converterCache = null;
    return null;
}

// ============================================================
// KONVERSI DOCX → PDF
// ============================================================
async function convertDocxToPdf(srcPath, pdfOutputDir) {
    if (!isUsableFile(srcPath)) return null;
    const srcExt = path.extname(srcPath).toLowerCase();
    if (srcExt === '.pdf') return srcPath;

    const defaultOutDir = path.dirname(srcPath);
    const outDir = pdfOutputDir || defaultOutDir;
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    const wasInTempBefore = path.resolve(path.dirname(srcPath)) === path.resolve(TEMP_DIR);
    let workingPath = srcPath;
    if (!wasInTempBefore) {
        const moved = moveToTemp(srcPath);
        if (!moved) {
            console.log(chalk.yellow(`⚠️ Gagal memindah DOCX ke temp, dilewati.`));
            return null;
        }
        workingPath = moved;
        console.log(chalk.gray(`📦 DOCX → temp: ${path.basename(workingPath)}`));
    }

    const pdfBasename = path.basename(workingPath, path.extname(workingPath)) + '.pdf';
    const targetPdfPath = path.join(outDir, pdfBasename);

    if (isUsableFile(targetPdfPath)) {
        console.log(chalk.green(`✅ PDF sudah ada, lewati konversi: ${path.basename(targetPdfPath)}`));
        deleteTempFile(workingPath);
        return targetPdfPath;
    }

    const converter = await detectConverter();
    if (!converter) {
        console.log(chalk.yellow(`   ℹ️ Konversi dilewati, DOCX tetap di temp untuk dikirim apa adanya.`));
        return workingPath;
    }

    const profileDir = path.join(TEMP_DIR, '.lo_profile');
    if (!fs.existsSync(profileDir)) fs.mkdirSync(profileDir, { recursive: true });
    const profileUri = 'file:///' + profileDir.replace(/\\/g, '/').replace(/^\//, '');

    const args = [
        '--headless', '--invisible', '--nologo', '--nofirststartwizard',
        '--norestore', '--nolockcheck', '--nodefault',
        `-env:UserInstallation=${profileUri}`,
        '--convert-to', 'pdf', '--outdir', outDir, workingPath
    ];

    try {
        await execFileAsync(converter.cmd, args, {
            timeout: 90000, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'],
        });
    } catch (e) {
        const detail = (e.stderr || e.stdout || e.message || '').toString().trim();
        console.log(chalk.yellow(`⚠️ Konversi gagal via ${converter.cmd}: ${detail || e.message}`));
        return workingPath;
    }

    if (isUsableFile(targetPdfPath)) {
        console.log(chalk.green(`✅ Konversi DOCX→PDF berhasil: ${path.basename(targetPdfPath)}`));
        deleteTempFile(workingPath);
        return targetPdfPath;
    }

    if (process.platform === 'win32') {
        const fallback = path.join(process.cwd(), pdfBasename);
        if (isUsableFile(fallback)) {
            console.log(chalk.yellow(`ℹ️ PDF muncul di CWD, dipindah: ${pdfBasename}`));
            try {
                fs.renameSync(fallback, targetPdfPath);
                deleteTempFile(workingPath);
                return targetPdfPath;
            } catch {
                return fallback;
            }
        }
    }

    console.log(chalk.yellow(`⚠️ Konversi selesai tapi PDF tidak ditemukan di ${outDir}.`));
    return workingPath;
}

// ============================================================
// GABUNG DOKUMEN (Surat + CV + Ijazah) — URUTAN FIXED
// ------------------------------------------------------------
// Urutan TETAP: Surat → CV → Ijazah (tidak bisa diubah)
// ============================================================
async function mergeDocsPlain(paths, outputPath) {
    if (!PDFDocument) return { success: false, reason: 'pdf-lib-not-installed' };

    const { letterPath, cvPath, ijazahPath } = paths;

    // URUTAN FIXED: Surat dulu, lalu CV, terakhir Ijazah
    const FIXED_ORDER = ['letter', 'cv', 'ijazah'];

    const sourceMap = {
        letter: letterPath,
        cv: cvPath,
        ijazah: ijazahPath,
    };

    try {
        const merged = await PDFDocument.create();
        const pageCounts = {};
        const addedOrder = [];

        for (const key of FIXED_ORDER) {
            const srcPath = sourceMap[key];
            if (!isUsableFile(srcPath)) {
                console.log(chalk.gray(`   ℹ️ [${key}] tidak tersedia, dilewati`));
                continue;
            }
            const bytes = fs.readFileSync(srcPath);
            const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
            const pages = await merged.copyPages(doc, doc.getPageIndices());
            pages.forEach(p => merged.addPage(p));
            pageCounts[key] = doc.getPageCount();
            addedOrder.push(key);
        }

        if (addedOrder.length === 0) {
            return { success: false, reason: 'tidak-ada-dokumen-untuk-digabung' };
        }

        const mergedBytes = await merged.save();
        fs.writeFileSync(outputPath, mergedBytes);

        return {
            success: true,
            path: outputPath,
            pageCounts,
            addedOrder,
            totalPages: merged.getPageCount(),
        };
    } catch (e) {
        return { success: false, reason: e.message };
    }
}

// ============================================================
// 🔒 FLATTEN KE GAMBAR + PROTEKSI (BARU)
// ------------------------------------------------------------
// 1. Rasterize PDF → PNG (Ghostscript)
// 2. Rebuild PDF image-only (pdf-lib) → tidak bisa copy teks
// 3. Encrypt dengan permission: no print, no copy (Ghostscript)
// ============================================================
async function flattenAndProtect(inputPath, outputPath, options = {}) {
    const {
        dpi = BOT_CONFIG.FLATTEN_DPI,
        blockPrint = BOT_CONFIG.BLOCK_PRINT,
        blockCopy = BOT_CONFIG.BLOCK_COPY,
        ownerPassword = BOT_CONFIG.GS_OWNER_PASSWORD || ('galeh-protect-' + Date.now()),
        tmpDir = TEMP_DIR,
    } = options;

    if (!HAS_GHOSTSCRIPT) {
        return { success: false, error: 'ghostscript-tidak-tersedia' };
    }
    if (!PDFDocument) {
        return { success: false, error: 'pdf-lib-tidak-tersedia' };
    }
    if (!isUsableFile(inputPath)) {
        return { success: false, error: 'input-tidak-ada' };
    }

    // Hitung nilai permissions
    // Base: -3904 = deny all (reserved bits 1-2 = 0, bit 7-8 = 1, bits 13-32 = 1)
    // Bit value: print=4, modify=8, copy=16, annotate=32, forms=256, access=512, assemble=1024, print-hi=2048
    let permissions = -3904;
    permissions += 512;            // accessibility (biasanya selalu diizinkan)
    if (!blockPrint) permissions += 4 + 2048;  // print + print high quality
    if (!blockCopy)  permissions += 16;        // copy

    if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });
    const workDir = path.join(tmpDir, `flat_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`);
    fs.mkdirSync(workDir, { recursive: true });

    try {
        // ---- STEP 1: Rasterize PDF → PNG ----
        await execFileAsync('gs', [
            '-q', '-dNOPAUSE', '-dBATCH', '-dSAFER',
            '-sDEVICE=png16m',
            `-r${dpi}`,
            '-dTextAlphaBits=4',
            '-dGraphicsAlphaBits=4',
            `-sOutputFile=${path.join(workDir, 'page-%04d.png')}`,
            inputPath,
        ], { timeout: 180000, windowsHide: true });

        const pngs = fs.readdirSync(workDir)
            .filter(f => f.endsWith('.png'))
            .sort();

        if (pngs.length === 0) {
            throw new Error('Ghostscript tidak menghasilkan halaman');
        }

        // ---- STEP 2: Rebuild jadi PDF image-only ----
        const newPdf = await PDFDocument.create();
        for (const png of pngs) {
            const imgBytes = fs.readFileSync(path.join(workDir, png));
            const img = await newPdf.embedPng(imgBytes);
            const page = newPdf.addPage([img.width, img.height]);
            page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
        }
        newPdf.setTitle('');
        newPdf.setAuthor('');
        newPdf.setSubject('');
        newPdf.setKeywords([]);
        newPdf.setProducer('');
        newPdf.setCreator('');

        const plainBytes = await newPdf.save();
        const plainPath = path.join(workDir, '_plain.pdf');
        fs.writeFileSync(plainPath, plainBytes);

        // ---- STEP 3: Encrypt dengan permission ----
        const gsArgs = [
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
            plainPath,
        ];
        await execFileAsync('gs', gsArgs, { timeout: 90000, windowsHide: true });

        // ---- Cleanup ----
        try {
            for (const f of fs.readdirSync(workDir)) {
                fs.unlinkSync(path.join(workDir, f));
            }
            fs.rmdirSync(workDir);
        } catch { /* abaikan */ }

        return {
            success: true,
            path: outputPath,
            pageCount: pngs.length,
            permissions,
            blockedPrint: blockPrint,
            blockedCopy: blockCopy,
        };
    } catch (e) {
        try {
            if (fs.existsSync(workDir)) {
                for (const f of fs.readdirSync(workDir)) {
                    fs.unlinkSync(path.join(workDir, f));
                }
                fs.rmdirSync(workDir);
            }
        } catch { /* abaikan */ }
        return { success: false, error: e.message };
    }
}

// ============================================================
// RESOLVE PDF DARI CACHE
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
        console.log(chalk.yellow(`ℹ️ FILE_PATH ${meta.label} → konversi ke PDF (via temp)...`));
        const pdf = await convertDocxToPdf(target.filePath, path.dirname(target.filePath));
        if (pdf) return pdf;
    }

    if (isUsableFile(outputPdfPath)) {
        console.log(chalk.green(`✅ Pakai PDF di output: ${path.basename(outputPdfPath)}`));
        return outputPdfPath;
    }

    if (isUsableFile(outputDocxPath)) {
        console.log(chalk.yellow(`ℹ️ DOCX di output → konversi ke PDF via temp: ${path.basename(outputDocxPath)}`));
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
        const auth = new google.auth.GoogleAuth({
            credentials,
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });
        const sheets = google.sheets({ version: 'v4', auth });
        const SHEET_NAME = 'Sheet1';

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${SHEET_NAME}!A:G`,
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

        if (companyIdx === -1 || positionIdx === -1 || contactIdx === -1 || statusIdx === -1) {
            console.error(chalk.red('❌ Header tidak lengkap. Perlu: COMPANY, POSITION, CONTACT, STATUS'));
            return null;
        }

        const allRows = [];
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i] || [];
            const company = row[companyIdx]?.trim() || '';
            const position = row[positionIdx]?.trim() || '';
            const contact = row[contactIdx]?.trim() || '';
            const requirements = requirementsIdx !== -1 ? (row[requirementsIdx]?.trim() || '') : '';
            const status = row[statusIdx]?.trim().toLowerCase() || '';
            const filePath = filePathIdx !== -1 ? (row[filePathIdx]?.trim() || '') : '';
            const subject = subjectIdx !== -1 ? (row[subjectIdx]?.trim() || '') : '';
            allRows.push({
                sheetRow: i + 1,
                company, position, contact, requirements, status, filePath, subject,
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
                console.log(chalk.yellow(
                    `🗑️  [baris ${r.sheetRow}] ${r.company} — ${r.position} → SUDAH PERNAH DIKIRIM. Baris duplikat dihapus.`
                ));
                continue;
            }

            if (seenInQueue.has(key)) {
                rowsToDelete.push(r.sheetRow);
                console.log(chalk.yellow(
                    `🗑️  [baris ${r.sheetRow}] ${r.company} — ${r.position} → DUPLIKAT. Baris duplikat dihapus.`
                ));
                continue;
            }

            seenInQueue.add(key);
            targets.push({
                company: r.company,
                position: r.position,
                contact: r.contact,
                requirements: r.requirements,
                filePath: r.filePath,
                subject: r.subject,
            });
        }

        if (rowsToDelete.length > 0) {
            try {
                await deleteSheetRowsByIndex(sheets, SHEET_NAME, rowsToDelete);
                console.log(chalk.green(`✅ ${rowsToDelete.length} baris duplikat berhasil dihapus dari spreadsheet.`));
            } catch (delErr) {
                console.log(chalk.yellow(`⚠️ Gagal hapus baris duplikat di Sheet: ${delErr.message}`));
            }
        }

        return targets;
    } catch (error) {
        console.error(chalk.red('❌ Gagal membaca Google Sheets:'), error.message);
        console.log(chalk.yellow('\n💡 Cek dari terminal:'));
        console.log(chalk.white('   nslookup oauth2.googleapis.com'));
        console.log(chalk.white('   curl -v https://oauth2.googleapis.com/token'));
        return null;
    }
}

async function deleteSheetRowsByIndex(sheets, sheetTitle, rowNumbers) {
    const meta = await sheets.spreadsheets.get({
        spreadsheetId: SPREADSHEET_ID,
        fields: 'sheets.properties',
    });
    const sheet = (meta.data.sheets || []).find(s => s.properties?.title === sheetTitle);
    if (!sheet) throw new Error(`Sheet "${sheetTitle}" tidak ditemukan`);

    const sheetId = sheet.properties.sheetId;
    const sortedDesc = [...new Set(rowNumbers)].sort((a, b) => b - a);

    const requests = sortedDesc.map(rowNum => ({
        deleteDimension: {
            range: {
                sheetId, dimension: 'ROWS',
                startIndex: rowNum - 1, endIndex: rowNum,
            },
        },
    }));

    await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        resource: { requests },
    });
}

async function updateGoogleSheetStatus(company, position, status = 'sent', filePath = '') {
    try {
        const auth = new google.auth.GoogleAuth({
            credentials,
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });
        const sheets = google.sheets({ version: 'v4', auth });

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Sheet1!A:G',
        });
        const rows = response.data.values;
        if (!rows) return false;

        const headers = rows[0].map(h => h.trim().toUpperCase());
        const companyIdx = headers.indexOf('COMPANY');
        const positionIdx = headers.indexOf('POSITION');
        const statusIdx = headers.indexOf('STATUS');
        const filePathIdx = headers.indexOf('FILE_PATH');

        let targetRow = -1;
        for (let i = 1; i < rows.length; i++) {
            const rowCompany = rows[i]?.[companyIdx]?.trim() || '';
            const rowPosition = rows[i]?.[positionIdx]?.trim() || '';
            if (rowCompany === company && rowPosition === position) {
                targetRow = i + 1;
                break;
            }
        }

        if (targetRow === -1) return false;

        const statusColumnLetter = String.fromCharCode(65 + statusIdx);
        await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `Sheet1!${statusColumnLetter}${targetRow}`,
            valueInputOption: 'RAW',
            resource: { values: [[status]] },
        });

        if (filePath && filePathIdx !== -1) {
            const filePathColumnLetter = String.fromCharCode(65 + filePathIdx);
            await sheets.spreadsheets.values.update({
                spreadsheetId: SPREADSHEET_ID,
                range: `Sheet1!${filePathColumnLetter}${targetRow}`,
                valueInputOption: 'RAW',
                resource: { values: [[filePath]] },
            });
        }

        return true;
    } catch (error) {
        console.error(chalk.red('❌ Gagal update Google Sheets:'), error.message);
        return false;
    }
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
    if (!process.env.GOOGLE_REFRESH_TOKEN) {
        const err = new Error('REFRESH_TOKEN tidak ditemukan di .env');
        err.isConfigError = true;
        throw err;
    }
    oAuth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
    try { await oAuth2Client.getAccessToken(); } catch (e) {
        console.log(chalk.yellow('⚠️ Refresh token tidak valid, coba perbarui di .env'));
        const err = new Error('Invalid refresh token');
        err.isConfigError = true;
        throw err;
    }
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

    const subject = (customSubject && customSubject.trim())
        ? customSubject.trim()
        : buildDefaultSubject(position);

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
    // CV tidak dilampirkan terpisah — sudah digabung ke surat
    if (isUsableFile(letterPath)) {
        const fname = letterFileName && letterFileName.trim()
            ? letterFileName.trim()
            : path.basename(letterPath);
        attachments.push({ filename: fname, path: letterPath });
    }

    const mailOptions = {
        from: process.env.GOOGLE_USER || process.env.EMAIL_USER,
        to: contact,
        subject,
        text,
        attachments,
    };
    await transporter.sendMail(mailOptions);
}

async function sendEmail(contact, company, position, cvPath, letterPath, customSubject = '', letterFileName = '') {
    let attempt = 0;
    while (true) {
        attempt++;
        try {
            console.log(chalk.gray(`   ⏳ Mengirim email ke ${contact}...`));
            await attemptSendEmail(contact, company, position, cvPath, letterPath, customSubject, letterFileName);
            console.log(chalk.green(`📧 Email terkirim ke ${contact}${attempt > 1 ? ` (percobaan ke-${attempt})` : ''}`));
            return true;
        } catch (error) {
            console.error(chalk.red(`❌ Gagal kirim email ke ${contact} (percobaan ke-${attempt}):`), error.message);

            if (error.isConfigError) {
                console.log(chalk.yellow('💡 Perbarui GOOGLE_REFRESH_TOKEN di .env.'));
                return false;
            }

            const waitMs = Math.min(5000 * attempt, 5 * 60 * 1000);
            console.log(chalk.yellow(`🔄 Reconnect & coba lagi dalam ${Math.round(waitMs / 1000)} detik...`));
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
    if (waState.sentTargets.has(targetKey)) {
        console.log(chalk.gray(`   ℹ️ ${targetKey} sudah terkirim di run ini, dilewati.`));
        return true;
    }

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
        if (!connected) {
            console.log(chalk.red(`❌ Koneksi WhatsApp tidak pulih, lewati percobaan ${attempt}/${maxRetries}.`));
            if (attempt < maxRetries) await delay(5000 * attempt);
            continue;
        }

        const sock = waState.sock;
        try {
            let targetJid = jid;
            try {
                const [result] = await sock.onWhatsApp(jid);
                if (!result || !result.exists) {
                    console.log(chalk.yellow(`⚠️ Nomor ${contact} tidak terdaftar di WhatsApp. Lewati.`));
                    return false;
                }
                if (result.jid && result.jid !== jid) {
                    console.log(chalk.gray(`   ℹ️ JID resmi: ${result.jid}`));
                    targetJid = result.jid;
                }
            } catch (checkErr) {
                console.log(chalk.gray(`   (Tidak bisa verifikasi nomor: ${checkErr.message})`));
            }

            console.log(chalk.gray('   ⏳ Mengirim pesan teks...'));
            const textResult = await sendMessageConfirmed(waState, targetJid, { text: messageText });
            if (!textResult.success) {
                throw new Error(`Pesan teks gagal (${textResult.reason}${textResult.error ? ': ' + textResult.error : ''})`);
            }
            const textAck = textResult.ackStatus !== null ? `ack ${textResult.ackStatus}` : '(ack belum masuk, tetap terkirim)';
            console.log(chalk.gray(`   ↳ pesan teks id: ${textResult.id} ${textAck}`));
            await delay(jitterDelay(BOT_CONFIG.DELAY_BETWEEN_SENDS));

            // CV tidak dikirim terpisah — sudah digabung di file final
            if (isUsableFile(letterPath)) {
                const letterBuffer = fs.readFileSync(letterPath);
                const letterMeta = resolveDocMeta(letterPath);
                const fname = letterFileName && letterFileName.trim()
                    ? letterFileName.trim()
                    : `Surat-${company}${letterMeta.ext}`;

                if (!letterMeta.isPdf) {
                    console.log(chalk.yellow(`   ⚠️ File surat bukan PDF (${letterMeta.label}): ${path.basename(letterPath)}`));
                }
                console.log(chalk.gray(`   ⏳ Mengirim dokumen lamaran (${letterMeta.label}) — ${fname}...`));
                const letterResult = await sendMessageConfirmed(waState, targetJid, {
                    document: letterBuffer,
                    mimetype: letterMeta.mimetype,
                    fileName: fname,
                    caption: `Lamaran - ${company} (${position})`
                });
                if (!letterResult.success) {
                    throw new Error(`Surat gagal (${letterResult.reason}${letterResult.error ? ': ' + letterResult.error : ''})`);
                }
                const letterAck = letterResult.ackStatus !== null ? `ack ${letterResult.ackStatus}` : '(ack belum masuk)';
                console.log(chalk.gray(`   ↳ Dokumen id: ${letterResult.id} ${letterAck}`));
                await delay(jitterDelay(BOT_CONFIG.DELAY_BETWEEN_SENDS));
            }

            console.log(chalk.green(`📱 WhatsApp terkirim ke ${contact}`));
            waState.sentTargets.add(targetKey);
            return true;
        } catch (error) {
            console.error(chalk.red(`❌ Gagal kirim WhatsApp (${attempt}/${maxRetries}):`), error.message);
            if (attempt < maxRetries) await delay(5000 * attempt);
        }
    }
    return false;
}

// ============================================================
// HELPER
// ============================================================
function getGreetingByTime() {
    const wibHour = parseInt(
        new Date().toLocaleString('en-US', {
            timeZone: 'Asia/Jakarta', hour: '2-digit', hour12: false
        }),
        10
    );
    if (wibHour >= 5 && wibHour < 11) return 'Selamat pagi';
    if (wibHour >= 11 && wibHour < 15) return 'Selamat siang';
    if (wibHour >= 15 && wibHour < 18) return 'Selamat sore';
    return 'Selamat malam';
}

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function jitterDelay(baseMs) {
    const variance = baseMs * 0.3;
    return Math.round(baseMs - variance + Math.random() * (variance * 2));
}

function isEmail(contact) {
    return contact && contact.includes('@');
}

// ============================================================
// NAMA FILE OUTPUT: {Nama}_{Posisi}.pdf
// ============================================================
function buildOutputFileName(position) {
    const safeApplicant = sanitizeName(BOT_CONFIG.APPLICANT_NAME || 'Pelamar');
    const safePosition = sanitizeName(position || 'Posisi');
    return `${safeApplicant}_${safePosition}.pdf`;
}

// ============================================================
// GENERATE ATAU AMBIL PDF SURAT
// ============================================================
async function getOrCreateLetterPdf(target, safeName) {
    const existing = await resolveExistingLetterPdf(target, safeName);
    if (existing) {
        return { pdfPath: existing, fromCache: true };
    }

    const outputPdfPath = path.join(AI_CONFIG.OUTPUT_DIR, `${target.position} - ${safeName}.pdf`);
    const spinner = ora('Membuat surat via AI...').start();

    let result;
    try {
        result = await generateApplicationLetter(
            target.company,
            target.position,
            outputPdfPath,
            target.requirements || ''
        );
    } catch (error) {
        spinner.fail('Error: ' + error.message);
        return null;
    }

    if (!result || !result.success || !result.actualPath) {
        spinner.fail('Gagal generate');
        return null;
    }

    let genPath = result.actualPath;

    if (!isUsableFile(genPath)) {
        spinner.fail(`File hasil generate tidak ditemukan/kosong: ${genPath}`);
        return null;
    }

    const genMeta = resolveDocMeta(genPath);
    if (genMeta.isPdf) {
        spinner.succeed(`✅ Surat PDF siap: ${path.basename(genPath)}`);
        return { pdfPath: genPath, fromCache: false };
    }

    spinner.text = `Konversi ${genMeta.label} → PDF (via temp)...`;
    const pdfPath = await convertDocxToPdf(genPath, AI_CONFIG.OUTPUT_DIR);

    if (!pdfPath) {
        spinner.fail(`Gagal konversi ${genMeta.label} ke PDF`);
        return null;
    }

    const finalMeta = resolveDocMeta(pdfPath);
    if (finalMeta.isPdf) {
        spinner.succeed(`✅ Surat PDF siap (dari konversi): ${path.basename(pdfPath)}`);
    } else {
        spinner.warn(`⚠️ Konversi gagal, kirim sebagai ${finalMeta.label}: ${path.basename(pdfPath)}`);
    }
    return { pdfPath, fromCache: false };
}

// ============================================================
// MERGE FINAL: Surat → CV → Ijazah, lalu flatten + protect
// ------------------------------------------------------------
// Alur:
//   1. Gabung Surat → CV → Ijazah (fixed order) → combined temp PDF
//   2. test-merge.js: stamp + footer (tanpa encrypt)
//   3. flattenAndProtect: rasterize + encrypt (no print, no copy)
//   4. Output: output/final/{Nama}_{Posisi}.pdf
// ============================================================
async function getOrCreateMergedPdf(letterPath, target, safeName) {
    const letterDir = path.resolve(path.dirname(letterPath));
    const officialOutputDir = path.resolve(path.join(__dirname, 'output'));
    const cvPath = BOT_CONFIG.CV_PATH;

    // Kalau surat sudah di final/, tidak perlu merge
    if (letterDir === path.resolve(FINAL_DIR)) {
        console.log(chalk.gray(`   💾 Surat sudah dalam bentuk merged (final/)`));
        const existingName = path.basename(letterPath);
        return { path: letterPath, merged: false, fromFinal: true, fileName: existingName };
    }

    if (!BOT_CONFIG.MERGE_WITH_IJAZAH) {
        return { path: letterPath, merged: false, fromFinal: false, fileName: path.basename(letterPath) };
    }

    // Nama file output
    const finalFileName = buildOutputFileName(target.position);
    const finalPath = path.join(FINAL_DIR, finalFileName);

    const hasIjazah = isUsableFile(IJAZAH_PATH);
    const hasCv = isUsableFile(cvPath);

    // ---- STEP 1: gabung dokumen (URUTAN FIXED: Surat → CV → Ijazah) ----
    let letterForMerge = letterPath;
    let tempCombined = null;

    const willCombine = PDFDocument && (hasCv || hasIjazah);

    if (willCombine) {
        if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });
        tempCombined = path.join(
            TEMP_DIR,
            `combined_${sanitizeName(target.position)}_${Date.now()}.pdf`
        );

        const mergeRes = await mergeDocsPlain(
            {
                letterPath,
                cvPath: hasCv ? cvPath : null,
                ijazahPath: hasIjazah ? IJAZAH_PATH : null,
            },
            tempCombined
        );

        if (mergeRes.success) {
            letterForMerge = tempCombined;
            const parts = mergeRes.addedOrder
                .map(k => `${k}:${mergeRes.pageCounts[k]}h`)
                .join(' + ');
            console.log(chalk.cyan(`   📎 Gabung ${parts} → ${mergeRes.totalPages}h`));
        } else {
            console.log(chalk.yellow(`   ⚠️ Gagal gabung dokumen: ${mergeRes.reason}`));
            if (tempCombined && fs.existsSync(tempCombined)) {
                try { fs.unlinkSync(tempCombined); } catch {}
            }
            tempCombined = null;
        }
    } else if (!PDFDocument) {
        console.log(chalk.yellow(`   ℹ️ pdf-lib tidak tersedia — dokumen tidak digabung`));
    } else {
        if (!hasCv) console.log(chalk.yellow(`   ℹ️ CV tidak ditemukan — lewati`));
        if (!hasIjazah) console.log(chalk.yellow(`   ℹ️ Ijazah tidak ditemukan — lewati`));
    }

    // ---- STEP 2: stamp + footer (TANPA encrypt — biar bisa di-flatten) ----
    const stampedPath = path.join(TEMP_DIR, `stamped_${safeName}_${Date.now()}.pdf`);

    console.log(chalk.cyan(
        `   🔗 Stamp + footer${hasIjazah ? ' (ijazah sudah digabung)' : ''}...`
    ));

    let stampResult;
    try {
        stampResult = await mergeAndStamp({
            letterPath: letterForMerge,
            ijazahPath: null,  // ijazah sudah digabung di STEP 1
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
            noEncrypt: true,   // ← PENTING: jangan encrypt dulu
        });
    } catch (e) {
        if (tempCombined) deleteTempFile(tempCombined);
        console.log(chalk.yellow(`   ⚠️ Merge error: ${e.message}`));
        return { path: letterPath, merged: false, fromFinal: false, fileName: path.basename(letterPath) };
    }

    if (tempCombined) deleteTempFile(tempCombined);

    if (!stampResult.success) {
        console.log(chalk.yellow(`   ⚠️ Stamp gagal: ${stampResult.error}`));
        return { path: letterPath, merged: false, fromFinal: false, error: stampResult.error, fileName: path.basename(letterPath) };
    }

    console.log(chalk.gray(`   🆔 Tracking ID: ${stampResult.trackingId} (${stampResult.dateStr})`));

    // ---- STEP 3: FLATTEN ke gambar + proteksi ----
    let flattenResult = { success: false, error: 'disabled' };
    let finalUsedPath = stampedPath;
    let wasFlattened = false;

    if (BOT_CONFIG.FLATTEN_TO_IMAGE && HAS_GHOSTSCRIPT) {
        console.log(chalk.cyan(
            `   🖼️  Flatten ke gambar (DPI ${BOT_CONFIG.FLATTEN_DPI}) ` +
            `+ ${BOT_CONFIG.BLOCK_PRINT ? 'blokir print' : 'print OK'} ` +
            `+ ${BOT_CONFIG.BLOCK_COPY ? 'blokir copy' : 'copy OK'}...`
        ));

        flattenResult = await flattenAndProtect(stampedPath, finalPath, {
            dpi: BOT_CONFIG.FLATTEN_DPI,
            blockPrint: BOT_CONFIG.BLOCK_PRINT,
            blockCopy: BOT_CONFIG.BLOCK_COPY,
        });

        if (flattenResult.success) {
            finalUsedPath = finalPath;
            wasFlattened = true;
            console.log(chalk.green(
                `   🔒 Terproteksi: ${flattenResult.pageCount} halaman jadi gambar, ` +
                `print=${BOT_CONFIG.BLOCK_PRINT ? 'NO' : 'YES'}, ` +
                `copy=${BOT_CONFIG.BLOCK_COPY ? 'NO' : 'YES'}`
            ));
            // Hapus file stamped (intermediate)
            try { fs.unlinkSync(stampedPath); } catch {}
        } else {
            console.log(chalk.yellow(`   ⚠️ Flatten gagal: ${flattenResult.error}`));
            console.log(chalk.yellow(`   → Fallback: pakai PDF stamped tanpa proteksi`));
        }
    } else if (!HAS_GHOSTSCRIPT && BOT_CONFIG.FLATTEN_TO_IMAGE) {
        console.log(chalk.yellow(`   ⚠️ Ghostscript tidak ada — flatten dilewati`));
        if (BOT_CONFIG.REQUIRE_GHOSTSCRIPT) {
            if (tempCombined) deleteTempFile(tempCombined);
            return { path: letterPath, merged: false, fromFinal: false, error: 'ghostscript-required', fileName: path.basename(letterPath) };
        }
    }

    // Kalau flatten gagal / disabled, move stamped → final
    if (!wasFlattened) {
        try {
            if (fs.existsSync(finalPath)) fs.unlinkSync(finalPath);
            fs.renameSync(stampedPath, finalPath);
            finalUsedPath = finalPath;
        } catch (e) {
            console.log(chalk.yellow(`   ⚠️ Gagal pindah ke final: ${e.message}`));
            finalUsedPath = stampedPath;
        }
    }

    const tagList = [];
    if (hasCv) tagList.push('CV ✓');
    if (hasIjazah) tagList.push('Ijazah ✓');
    if (wasFlattened) tagList.push('🖼️ GAMBAR');

    console.log(chalk.green(
        `   ✅ Final: ${finalFileName} ` +
        `(${stampResult.totalPages}h${tagList.length ? ', ' + tagList.join(', ') : ''})`
    ));

    // Hapus surat asli HANYA kalau di folder output/ resmi
    if (letterDir === officialOutputDir) {
        try {
            fs.unlinkSync(letterPath);
            console.log(chalk.gray(`   🗑️  Surat asli dihapus: ${path.basename(letterPath)}`));
        } catch (e) {
            console.log(chalk.yellow(`   ⚠️ Gagal hapus surat asli: ${e.message}`));
        }
    } else {
        console.log(chalk.gray(`   ℹ️ Surat asli di luar output/ — tidak dihapus`));
    }

    return {
        path: finalUsedPath,
        merged: true,
        fromFinal: false,
        fileName: finalFileName,
        trackingId: stampResult.trackingId,
        totalPages: stampResult.totalPages,
        flattened: wasFlattened,
        blockedPrint: flattenResult.blockedPrint || false,
        blockedCopy: flattenResult.blockedCopy || false,
    };
}

// ============================================================
// MAIN
// ============================================================
async function main() {
    console.log(chalk.bold.cyan('\n' + '='.repeat(60)));
    console.log(chalk.bold.cyan('📱 WHATSAPP & EMAIL JOB APPLICATION BOT (Baileys)'));
    console.log(chalk.bold.cyan('   Sumber: Google Spreadsheet'));
    console.log(chalk.bold.cyan('   Session: MongoDB (sessionschemas)'));
    console.log(chalk.bold.cyan('   PDF urutan: Surat → CV → Ijazah (FIXED)'));
    console.log(chalk.bold.cyan(`   PDF: stamp + footer + flatten gambar`));
    console.log(chalk.bold.cyan(`   Proteksi: print=${BOT_CONFIG.BLOCK_PRINT ? 'BLOCKED' : 'allowed'}, copy=${BOT_CONFIG.BLOCK_COPY ? 'BLOCKED' : 'allowed'}`));
    console.log(chalk.bold.cyan(`   Nama file: {Nama}_{Posisi}.pdf`));
    console.log(chalk.bold.cyan('   Email: custom SUBJECT dari sheet (fallback ke default)'));

    // Cek Ghostscript
    const gsVer = await checkGhostscript();
    if (gsVer) {
        console.log(chalk.bold.green(`   ✅ Ghostscript v${gsVer} terdeteksi`));
    } else if (BOT_CONFIG.FLATTEN_TO_IMAGE) {
        console.log(chalk.bold.yellow(`   ⚠️  Ghostscript TIDAK terdeteksi — flatten akan dilewati`));
        console.log(chalk.bold.white(`      Install: apt-get install ghostscript`));
        if (BOT_CONFIG.REQUIRE_GHOSTSCRIPT) {
            console.log(chalk.bold.red(`   ❌ REQUIRE_GHOSTSCRIPT=true, keluar.`));
            process.exit(1);
        }
    }

    if (!PDFDocument) {
        console.log(chalk.bold.yellow('   ⚠️  pdf-lib TIDAK ADA — dokumen tidak akan digabung!'));
    }
    if (DEBUG_WA) {
        console.log(chalk.bold.yellow('   ⚠️  MODE DEBUG_WA AKTIF'));
    }
    console.log(chalk.bold.cyan('='.repeat(60) + '\n'));

    initializeFolders();

    if (!fs.existsSync(AI_CONFIG.TEMPLATE_PATH)) {
        console.log(chalk.red(`❌ Template DOCX tidak ditemukan: ${AI_CONFIG.TEMPLATE_PATH}`));
        process.exit(1);
    }
    if (!fs.existsSync(BOT_CONFIG.CV_PATH)) {
        console.log(chalk.red(`❌ CV tidak ditemukan: ${BOT_CONFIG.CV_PATH}`));
        process.exit(1);
    }
    if (!SPREADSHEET_ID) {
        console.log(chalk.red('❌ SPREADSHEET_ID belum diatur di .env'));
        process.exit(1);
    }

    const targets = await loadTargetsFromGoogleSheets();
    if (!targets) {
        console.log(chalk.red('❌ Gagal memuat target dari Google Sheets.'));
        process.exit(1);
    }

    if (targets.length === 0) {
        console.log(chalk.green('\n✅ Tidak ada target yang perlu diproses.'));
        process.exit(0);
    }

    console.log(chalk.yellow('\n📋 Target yang akan diproses:'));
    targets.forEach((t, i) => {
        const channel = isEmail(t.contact) ? '📧 Email' : '📱 WhatsApp';
        const subjInfo = channel.includes('Email')
            ? ` — subjek: "${(t.subject && t.subject.trim()) || buildDefaultSubject(t.position)}"`
            : '';
        const fileName = buildOutputFileName(t.position);
        console.log(`  ${i+1}. ${t.company} - ${t.position} (${channel})`);
        console.log(chalk.gray(`      📄 file: ${fileName}${subjInfo}`));
    });

    const hasWhatsApp = targets.some(t => !isEmail(t.contact));
    let waState = null;
    if (hasWhatsApp) {
        console.log(chalk.cyan('\n📱 Menghubungkan WhatsApp...'));
        waState = createWaState();
        await connectToWhatsApp(waState);
        console.log(chalk.green(`✅ WhatsApp siap, tunggu sinkronisasi ${BOT_CONFIG.WA_SYNC_DELAY_MS / 1000}s...`));
        await delay(BOT_CONFIG.WA_SYNC_DELAY_MS);
    }

    let successCount = 0;
    let failCount = 0;
    let cacheHitCount = 0;
    let aiCallCount = 0;
    let mergeCount = 0;
    let flattenedCount = 0;
    let blockedPrintCount = 0;
    let blockedCopyCount = 0;
    const startTime = Date.now();

    for (let i = 0; i < targets.length; i++) {
        const target = targets[i];
        console.log(chalk.cyan(`\n[${i+1}/${targets.length}] ${target.company} — ${target.position}`));

        const safeName = sanitizeName(target.company);

        // ---- 1. Dapatkan surat ----
        const letterInfo = await getOrCreateLetterPdf(target, safeName);
        if (!letterInfo) {
            await updateGoogleSheetStatus(target.company, target.position, 'failed');
            failCount++;
            continue;
        }
        let letterPath = letterInfo.pdfPath;
        if (letterInfo.fromCache) {
            cacheHitCount++;
            console.log(chalk.gray(`   💾 Cache hit — AI tidak dipanggil (hemat token).`));
        } else {
            aiCallCount++;
            console.log(chalk.gray(`   🤖 AI dipanggil sekali untuk generate surat.`));
        }

        // ---- 2. Merge final + flatten + protect ----
        const mergedInfo = await getOrCreateMergedPdf(letterPath, target, safeName);
        letterPath = mergedInfo.path;
        const letterFileName = mergedInfo.fileName || path.basename(letterPath);

        if (mergedInfo.merged) {
            mergeCount++;
            if (mergedInfo.flattened) flattenedCount++;
            if (mergedInfo.blockedPrint) blockedPrintCount++;
            if (mergedInfo.blockedCopy) blockedCopyCount++;
        }

        // ---- 3. Kirim ----
        let sendSuccess = false;
        if (isEmail(target.contact)) {
            const subjUsed = (target.subject && target.subject.trim())
                ? target.subject.trim()
                : buildDefaultSubject(target.position);
            console.log(chalk.gray(`   ✉️  Subjek: "${subjUsed}"${(target.subject && target.subject.trim()) ? ' (custom)' : ' (default)'}`));
            console.log(chalk.gray(`   📄 File  : ${letterFileName}`));
            sendSuccess = await sendEmail(
                target.contact,
                target.company,
                target.position,
                BOT_CONFIG.CV_PATH,
                letterPath,
                target.subject,
                letterFileName
            );
        } else if (waState) {
            console.log(chalk.gray(`   📄 File  : ${letterFileName}`));
            sendSuccess = await sendWhatsApp(
                waState,
                target.contact,
                target.company,
                target.position,
                BOT_CONFIG.CV_PATH,
                letterPath,
                letterFileName
            );
        }

        // ---- 4. Update Sheet ----
        if (sendSuccess) {
            await updateGoogleSheetStatus(target.company, target.position, 'sent', letterPath);
            successCount++;
        } else {
            console.log(chalk.yellow(`⚠️ Gagal kirim ke ${target.company}, status 'failed'.`));
            await updateGoogleSheetStatus(target.company, target.position, 'failed', letterPath);
            failCount++;
        }

        if (i < targets.length - 1) {
            const delayTime = isEmail(target.contact)
                ? BOT_CONFIG.DELAY_BETWEEN_TARGETS_EMAIL_MIN + Math.floor(Math.random() * (BOT_CONFIG.DELAY_BETWEEN_TARGETS_EMAIL_MAX - BOT_CONFIG.DELAY_BETWEEN_TARGETS_EMAIL_MIN))
                : BOT_CONFIG.DELAY_BETWEEN_TARGETS_WA_MIN + Math.floor(Math.random() * (BOT_CONFIG.DELAY_BETWEEN_TARGETS_WA_MAX - BOT_CONFIG.DELAY_BETWEEN_TARGETS_WA_MIN));
            console.log(chalk.gray(`⏳ Delay ${delayTime/1000}s...`));
            await delay(delayTime);
        }
    }

    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(chalk.bold.green('\n📊 SUMMARY'));
    console.log(chalk.green(`✅ Berhasil: ${successCount}`));
    console.log(chalk.red(`❌ Gagal: ${failCount}`));
    console.log(chalk.cyan(`💾 Cache hit (AI tidak dipanggil): ${cacheHitCount}`));
    console.log(chalk.cyan(`🤖 Panggilan AI: ${aiCallCount}`));
    console.log(chalk.cyan(`🔗 Merge PDF: ${mergeCount}`));
    console.log(chalk.cyan(`🖼️  Flatten ke gambar: ${flattenedCount}`));
    console.log(chalk.cyan(`🚫 Print diblokir: ${blockedPrintCount}`));
    console.log(chalk.cyan(`🚫 Copy diblokir: ${blockedCopyCount}`));
    console.log(chalk.cyan(`⏱️  ${elapsed} detik`));

    process.exit(0);
}

main().catch(error => {
    console.error(chalk.red('\n❌ Error:'), error.message);
    console.log(chalk.yellow('💡 Periksa konfigurasi .env, credentials.json, dan koneksi internet.'));
    process.exit(1);
});

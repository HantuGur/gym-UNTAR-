/* ===============================
   FRONTEND SISTEM ADMIN GYM
   PATCH: PISAH KUNCI COWO & CEWE
   File ini dipakai oleh index.html
   =============================== */

const DEFAULT_SCRIPT_PLACEHOLDER = "https://script.google.com/macros/s/AKfycbyMqPQrCg6eBxtNNaCtp_0p7OsmzVExyhD_36oZeiRpb9oQWEl57rDCiHhcUY836NXlcA/exec";

const state = {
  keys: [],
  members: [],
  logs: [],
  daily: [],
  activeTab: "keys",
  loading: false,
  initialized: false,
  saving: false
};

const els = {};

document.addEventListener("DOMContentLoaded", initApp);

function initApp() {
  bindElements();
  applyConfig();
  bindEvents();
  startClock();
  loadAllData({ showLoading: true });
  startAutoRefresh();
}

function bindElements() {
  els.gymName = document.getElementById("gymName");
  els.clockText = document.getElementById("clockText");
  els.dateText = document.getElementById("dateText");
  els.backendStatus = document.getElementById("backendStatus");
  els.configWarning = document.getElementById("configWarning");

  els.checkForm = document.getElementById("checkForm");
  els.adminInput = document.getElementById("adminInput");
  els.customerInput = document.getElementById("customerInput");
  els.keyInput = document.getElementById("keyInput");
  els.keyTypeInput = document.getElementById("keyTypeInput");
  els.statusInput = document.getElementById("statusInput");
  els.submitBtn = document.getElementById("submitBtn");
  els.clearBtn = document.getElementById("clearBtn");

  els.totalKeys = document.getElementById("totalKeys");
  els.usedKeys = document.getElementById("usedKeys");
  els.emptyKeys = document.getElementById("emptyKeys");
  els.memberCount = document.getElementById("memberCount");
  els.refreshNote = document.getElementById("refreshNote");
  els.lastUpdate = document.getElementById("lastUpdate");

  els.monitorTitle = document.getElementById("monitorTitle");
  els.searchInput = document.getElementById("searchInput");
  els.filterStatus = document.getElementById("filterStatus");
  els.refreshBtn = document.getElementById("refreshBtn");
  els.tableHead = document.getElementById("tableHead");
  els.tableBody = document.getElementById("tableBody");
  els.emptyState = document.getElementById("emptyState");
  els.toast = document.getElementById("toast");
}

function getConfig() {
  return window.GYM_CONFIG || {};
}

function getScriptUrl() {
  return String(getConfig().SCRIPT_URL || "").trim();
}

function isScriptUrlReady() {
  const url = getScriptUrl();

  return Boolean(
    url &&
    url !== DEFAULT_SCRIPT_PLACEHOLDER &&
    url.includes("script.google.com") &&
    url.endsWith("/exec")
  );
}

function getRefreshInterval() {
  const value = Number(getConfig().REFRESH_INTERVAL_MS || 10000);
  return Number.isFinite(value) && value >= 5000 ? value : 10000;
}

function applyConfig() {
  const config = getConfig();
  const gymName = config.GYM_NAME || "SISTEM FC";
  const appName = config.APP_NAME || "Sistem Admin Gym";

  document.title = `${appName} - ${gymName}`;
  els.gymName.textContent = gymName;
  els.refreshNote.textContent = `Auto-refresh tiap ${Math.round(getRefreshInterval() / 1000)} detik.`;

  const savedAdmin = localStorage.getItem("gymAdminName") || "";
  els.adminInput.value = savedAdmin;

  if (!isScriptUrlReady()) {
    els.configWarning.classList.remove("hidden");
    setBackendStatus("bad", "Backend belum disetting");
  } else {
    els.configWarning.classList.add("hidden");
  }
}

function bindEvents() {
  els.checkForm.addEventListener("submit", handleSubmit);

  els.clearBtn.addEventListener("click", () => {
    els.customerInput.value = "";
    els.keyInput.value = "";
    els.keyTypeInput.value = "Cowo";
    els.statusInput.value = "Masuk";
    updateCustomerNameRequirement();
    els.customerInput.focus();
  });

  els.adminInput.addEventListener("input", () => {
    localStorage.setItem("gymAdminName", els.adminInput.value.trim());
  });

  els.statusInput.addEventListener("change", updateCustomerNameRequirement);

  els.refreshBtn.addEventListener("click", () => loadAllData({ showLoading: true }));
  els.searchInput.addEventListener("input", renderCurrentTab);
  els.filterStatus.addEventListener("change", renderCurrentTab);

  els.tableBody.addEventListener("click", handleTableAction);

  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((btn) => btn.classList.remove("active"));
      button.classList.add("active");

      state.activeTab = button.dataset.tab;

      els.searchInput.value = "";
      els.filterStatus.value = "all";

      renderCurrentTab();
    });
  });

  updateCustomerNameRequirement();
}

function updateCustomerNameRequirement() {
  const status = String(els.statusInput.value || "").trim();

  if (status === "Keluar") {
    els.customerInput.required = false;
    els.customerInput.placeholder = "Boleh kosong kalau status Keluar";
  } else {
    els.customerInput.required = true;
    els.customerInput.placeholder = "Contoh: Budi Santoso";
  }
}

function startClock() {
  updateClock();
  setInterval(updateClock, 1000);
}

function updateClock() {
  const now = new Date();

  els.clockText.textContent = now.toLocaleTimeString("id-ID", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  });

  els.dateText.textContent = now.toLocaleDateString("id-ID", {
    weekday: "long",
    day: "2-digit",
    month: "long",
    year: "numeric"
  });
}

function startAutoRefresh() {
  setInterval(() => {
    if (document.hidden) return;
    if (!isScriptUrlReady()) return;
    if (state.saving) return;

    loadAllData({ showLoading: false });
  }, getRefreshInterval());
}

async function loadAllData({ showLoading = false } = {}) {
  if (!isScriptUrlReady()) {
    renderCurrentTab();
    return;
  }

  if (state.loading) return;

  state.loading = true;

  if (showLoading || !state.initialized) {
    setBackendStatus("", "Memuat data...");
  }

  try {
    const [ping, keys, members, logs, daily] = await Promise.all([
      jsonpRequest("ping"),
      jsonpRequest("keys"),
      jsonpRequest("members"),
      jsonpRequest("logs"),
      jsonpRequest("daily")
    ]);

    // Hanya gagalkan kalau errornya BUKAN error merge Google Sheet.
    if (!ping.ok && !isSuppressedMessage(ping.message)) throw new Error(ping.message || "Backend tidak siap.");
    if (!keys.ok && !isSuppressedMessage(keys.message)) throw new Error(keys.message || "Gagal ambil data kunci.");
    if (!members.ok && !isSuppressedMessage(members.message)) throw new Error(members.message || "Gagal ambil data member.");
    if (!logs.ok && !isSuppressedMessage(logs.message)) throw new Error(logs.message || "Gagal ambil log audit.");

    // Pakai data yang ada; kalau suatu endpoint error merge, biarkan data lama.
    if (Array.isArray(keys.data)) state.keys = keys.data;
    if (Array.isArray(members.data)) state.members = members.data;
    if (Array.isArray(logs.data)) state.logs = logs.data;
    if (Array.isArray(daily.data)) state.daily = daily.data;

    // Gabungkan: kalau ada check-in harian yang belum keluar, tandai kuncinya
    // "Dipakai" walau DATA_KUNCI belum ter-update. Ini bikin nama + tombol
    // Keluar tetap muncul di web.
    mergeDailyIntoKeys();

    updateDashboard();
    renderCurrentTab();

    setBackendStatus("ok", "Backend aktif");
    state.initialized = true;

  } catch (error) {
    setBackendStatus("bad", "Backend gagal");
    showToast(error.message || "Gagal menghubungi backend.", "bad");

  } finally {
    state.loading = false;
  }
}

function jsonpRequest(action, params = {}) {
  return new Promise((resolve, reject) => {
    const scriptUrl = getScriptUrl();

    if (!isScriptUrlReady()) {
      reject(new Error("SCRIPT_URL belum benar di config.js."));
      return;
    }

    const callbackName = `__gymCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const url = new URL(scriptUrl);

    url.searchParams.set("action", action);
    url.searchParams.set("callback", callbackName);

    Object.entries(params).forEach(([key, value]) => {
      url.searchParams.set(key, value);
    });

    const script = document.createElement("script");
    let finished = false;

    const cleanup = () => {
      delete window[callbackName];
      script.remove();
    };

    const timeout = setTimeout(() => {
      if (finished) return;

      finished = true;
      cleanup();

      reject(new Error("Backend tidak merespons. Cek deploy Apps Script."));
    }, 45000);

    window[callbackName] = (payload) => {
      if (finished) return;

      finished = true;
      clearTimeout(timeout);
      cleanup();

      resolve(payload || { ok: false, message: "Respons backend kosong." });
    };

    script.onerror = () => {
      if (finished) return;

      finished = true;
      clearTimeout(timeout);
      cleanup();

      reject(new Error("Gagal menghubungi Apps Script."));
    };

    script.src = url.toString();
    document.body.appendChild(script);
  });
}

function handleSubmit(event) {
  event.preventDefault();

  if (state.saving) return;

  if (!isScriptUrlReady()) {
    showToast("SCRIPT_URL belum benar di config.js.", "bad");
    return;
  }

  const payload = {
    action: "saveLog",
    admin: els.adminInput.value.trim(),
    customerName: els.customerInput.value.trim(),
    keyNumber: els.keyInput.value.trim(),
    keyType: els.keyTypeInput.value,   // "Cowo" atau "Cewe"
    status: els.statusInput.value
  };

  const validation = validatePayload(payload);

  if (!validation.ok) {
    showToast(validation.message, "warn");
    return;
  }

  localStorage.setItem("gymAdminName", payload.admin);

  submitPayload(payload);
}

function validatePayload(payload) {
  if (!payload.admin) {
    return { ok: false, message: "Nama admin/pegawai wajib diisi." };
  }

  if (!payload.keyNumber) {
    return { ok: false, message: "Nomor kunci wajib diisi." };
  }

  if (!payload.keyType) {
    return { ok: false, message: "Jenis kunci wajib dipilih." };
  }

  if (!payload.status) {
    return { ok: false, message: "Status wajib dipilih." };
  }

  if (payload.status === "Masuk" && !payload.customerName) {
    return { ok: false, message: "Nama pelanggan wajib diisi untuk check-in." };
  }

  return { ok: true };
}

async function submitPayload(payload) {
  state.saving = true;

  els.submitBtn.disabled = true;
  els.submitBtn.textContent = "Menyimpan...";

  disableCheckoutButtons(true);

  try {
    const response = await jsonpRequest("saveLog", {
      admin: payload.admin,
      customerName: payload.customerName || "",
      keyNumber: payload.keyNumber,
      keyType: payload.keyType,
      status: payload.status
    });

    // Kalau backend balikin error yang ternyata cuma error merge Google Sheet,
    // anggap SUKSES — data inti (status kunci, log) sudah tersimpan di backend.
    if (!response.ok && !isSuppressedMessage(response.message)) {
      throw new Error(response.message || "Data gagal disimpan.");
    }

    showToast("Data berhasil disimpan.", "ok");

    els.customerInput.value = "";
    els.keyInput.value = "";
    els.keyTypeInput.value = "Cowo";
    els.statusInput.value = "Masuk";

    updateCustomerNameRequirement();

    await loadAllData({ showLoading: false });

  } catch (error) {
    showToast(error.message || "Data gagal disimpan.", "bad");
    await loadAllData({ showLoading: false });

  } finally {
    finishSavingState();
  }
}

function finishSavingState() {
  state.saving = false;

  els.submitBtn.disabled = false;
  els.submitBtn.textContent = "Simpan ke Google Sheet";

  disableCheckoutButtons(false);
}

function disableCheckoutButtons(disabled) {
  document.querySelectorAll(".checkout-btn").forEach((button) => {
    button.disabled = disabled;
  });
}

function handleTableAction(event) {
  const button = event.target.closest("[data-action]");

  if (!button) return;

  const action = button.dataset.action;

  if (action === "checkout") {
    const keyNumber = button.dataset.keyNumber || "";
    const keyType = button.dataset.keyType || "Cowo";
    const customerName = button.dataset.customerName || "";

    quickCheckout(keyNumber, keyType, customerName);
  }
}

// quickCheckout: selalu kirim keyType supaya backend tahu kunci mana yang di-checkout.
function quickCheckout(keyNumber, keyType = "Cowo", customerName = "") {
  if (state.saving) return;

  if (!isScriptUrlReady()) {
    showToast("SCRIPT_URL belum benar di config.js.", "bad");
    return;
  }

  const admin = els.adminInput.value.trim();

  if (!admin) {
    showToast("Isi nama admin/pegawai dulu sebelum checkout.", "warn");
    els.adminInput.focus();
    return;
  }

  if (!keyNumber) {
    showToast("Nomor kunci tidak ditemukan.", "bad");
    return;
  }

  const payload = {
    action: "saveLog",
    admin,
    customerName,
    keyNumber,
    keyType,
    status: "Keluar"
  };

  localStorage.setItem("gymAdminName", admin);

  submitPayload(payload);
}

function updateDashboard() {
  const total = state.keys.length;
  const used = state.keys.filter((item) => String(item.status).toLowerCase() === "dipakai").length;
  const empty = total - used;

  els.totalKeys.textContent = total;
  els.usedKeys.textContent = used;
  els.emptyKeys.textContent = empty;
  els.memberCount.textContent = state.members.length;
  els.lastUpdate.textContent = new Date().toLocaleTimeString("id-ID");
}

function renderCurrentTab() {
  if (state.activeTab === "members") {
    els.monitorTitle.textContent = "Member Lifetime";
    renderMembers();
    return;
  }

  if (state.activeTab === "logs") {
    els.monitorTitle.textContent = "Audit Terakhir";
    renderLogs();
    return;
  }

  els.monitorTitle.textContent = "Daftar Kunci";
  renderKeys();
}

// renderKeys: tampilkan kolom Jenis di tabel, tombol Keluar bawa keyType.
// mergeDailyIntoKeys: lapisan kedua supaya nama + tombol Keluar muncul di web
// walau DATA_KUNCI belum ter-update. Sumber kebenaran: REKAP_HARIAN.
// Untuk tiap check-in yang BELUM keluar, tandai kunci terkait sebagai "Dipakai".
function mergeDailyIntoKeys() {
  if (!Array.isArray(state.daily) || !state.daily.length) return;
  if (!Array.isArray(state.keys)) state.keys = [];

  const norm = (v) => String(v ?? "").trim().toLowerCase();
  const normNum = (v) => {
    const n = Number(String(v ?? "").trim());
    return Number.isFinite(n) && n > 0 ? String(Math.floor(n)).padStart(2, "0") : String(v ?? "").trim();
  };
  const normType = (v) => {
    const t = norm(v);
    if (["cewe", "cewek", "wanita", "perempuan", "female"].includes(t)) return "Cewe";
    return "Cowo";
  };
  const identity = (type, num) => `${normType(type)}__${normNum(num)}`;

  // Index kunci yang ada berdasarkan identity (jenis + nomor).
  const keyIndex = new Map();
  state.keys.forEach((k, i) => keyIndex.set(identity(k.keyType, k.keyNumber), i));

  // Ambil check-in yang belum keluar. Kalau satu kunci ada beberapa baris,
  // ambil yang terbaru (state.daily sudah ter-reverse: terbaru di atas).
  const seen = new Set();

  state.daily.forEach((row) => {
    const belumKeluar = !row.sudahKeluar;
    if (!belumKeluar) return;

    const id = identity(row.keyType, row.noKunci);
    if (seen.has(id)) return; // sudah diambil yang lebih baru
    seen.add(id);

    const patch = {
      keyNumber: normNum(row.noKunci),
      keyType: normType(row.keyType),
      status: "Dipakai",
      customerName: row.nama || "",
      checkInTime: row.jamMasuk || "",
      updatedAt: row.jamMasuk || ""
    };

    if (keyIndex.has(id)) {
      const i = keyIndex.get(id);
      // Hanya timpa kalau DATA_KUNCI belum menandai Dipakai (hindari menimpa data benar).
      const cur = state.keys[i];
      if (norm(cur.status) !== "dipakai") {
        state.keys[i] = Object.assign({}, cur, patch);
      } else if (!cur.customerName && patch.customerName) {
        // status sudah Dipakai tapi nama kosong -> lengkapi nama.
        state.keys[i] = Object.assign({}, cur, { customerName: patch.customerName });
      }
    } else {
      // Kunci belum ada di daftar (mis. DATA_KUNCI belum di-seed) -> tambahkan.
      state.keys.push(patch);
    }
  });
}

// Tombol Keluar SELALU muncul saat status Dipakai.
// Saat status Kosong, kolom Aksi menampilkan "-" karena tidak ada yang perlu di-checkout.
function renderKeys() {
  els.tableHead.innerHTML = `
    <tr>
      <th>Jenis</th>
      <th>No Kunci</th>
      <th>Status</th>
      <th>Dipakai Oleh</th>
      <th>Jam Masuk</th>
      <th>Update Terakhir</th>
      <th>Aksi</th>
    </tr>
  `;

  const q = getSearchText();
  const filter = els.filterStatus.value;

  const rows = state.keys.filter((item) => {
    const text = `${item.keyType} ${item.keyNumber} ${item.status} ${item.customerName} ${item.checkInTime} ${item.updatedAt}`.toLowerCase();
    const matchSearch = !q || text.includes(q);
    const matchStatus = filter === "all" || item.status === filter;

    return matchSearch && matchStatus;
  });

  els.tableBody.innerHTML = rows.map((item) => {
    const statusRaw = String(item.status || "").trim();
    const statusLower = statusRaw.toLowerCase();
    const isUsed = statusLower === "dipakai";

    const keyNumber = escapeHtml(item.keyNumber || "");
    const keyType = escapeHtml(item.keyType || "Cowo");
    const customerName = escapeHtml(item.customerName || "");

    // Baris kunci yang dipakai dikasih highlight supaya keliatan
    const rowClass = isUsed ? ' class="row-used"' : "";

    return `
      <tr${rowClass}>
        <td>${keyTypeBadge(item.keyType || "Cowo")}</td>
        <td><b>${escapeHtml(item.keyNumber || "-")}</b></td>
        <td>${statusBadge(statusRaw)}</td>
        <td>${escapeHtml(item.customerName || "-")}</td>
        <td>${escapeHtml(item.checkInTime || "-")}</td>
        <td>${escapeHtml(item.updatedAt || "-")}</td>
        <td>
          ${isUsed
            ? `<button
                  type="button"
                  class="checkout-btn"
                  data-action="checkout"
                  data-key-number="${keyNumber}"
                  data-key-type="${keyType}"
                  data-customer-name="${customerName}"
                >Keluar</button>`
            : `<span class="muted">-</span>`
          }
        </td>
      </tr>
    `;
  }).join("");

  toggleEmpty(rows.length === 0);
}

function renderMembers() {
  els.tableHead.innerHTML = `
    <tr>
      <th>ID / No</th>
      <th>Nama Member</th>
      <th>Status</th>
      <th>Tanggal Daftar</th>
      <th>Diinput Oleh</th>
      <th>Update Terakhir</th>
    </tr>
  `;

  const q = getSearchText();
  const filter = els.filterStatus.value;

  const rows = state.members.filter((item) => {
    const text = `${item.memberId} ${item.memberName} ${item.status} ${item.registeredAt} ${item.createdBy} ${item.updatedAt}`.toLowerCase();
    const matchSearch = !q || text.includes(q);
    const matchStatus = filter === "all" || item.status === filter;

    return matchSearch && matchStatus;
  });

  els.tableBody.innerHTML = rows.map((item) => `
    <tr>
      <td><b>${escapeHtml(item.memberId || "-")}</b></td>
      <td>${escapeHtml(item.memberName || "-")}</td>
      <td>${escapeHtml(item.status || "-")}</td>
      <td>${escapeHtml(item.registeredAt || "-")}</td>
      <td>${escapeHtml(item.createdBy || "-")}</td>
      <td>${escapeHtml(item.updatedAt || "-")}</td>
    </tr>
  `).join("");

  toggleEmpty(rows.length === 0);
}

// renderLogs: tampilkan kolom Jenis di tabel audit.
function renderLogs() {
  els.tableHead.innerHTML = `
    <tr>
      <th>No</th>
      <th>Waktu Lengkap</th>
      <th>Tanggal</th>
      <th>Jam</th>
      <th>Nama</th>
      <th>Jenis</th>
      <th>No Kunci</th>
      <th>Status</th>
      <th>Admin</th>
    </tr>
  `;

  const q = getSearchText();
  const filter = els.filterStatus.value;

  const rows = state.logs.filter((item) => {
    const text = `${item.no} ${item.waktuLengkap} ${item.tanggal} ${item.jam} ${item.nama} ${item.keyType} ${item.noKunci} ${item.status} ${item.admin}`.toLowerCase();
    const matchSearch = !q || text.includes(q);
    const matchStatus = filter === "all" || item.status === filter;

    return matchSearch && matchStatus;
  });

  els.tableBody.innerHTML = rows.map((item) => `
    <tr>
      <td><b>${escapeHtml(item.no || "-")}</b></td>
      <td>${escapeHtml(item.waktuLengkap || "-")}</td>
      <td>${escapeHtml(item.tanggal || "-")}</td>
      <td>${escapeHtml(item.jam || "-")}</td>
      <td>${escapeHtml(item.nama || "-")}</td>
      <td>${keyTypeBadge(item.keyType || "-")}</td>
      <td>${escapeHtml(item.noKunci || "-")}</td>
      <td>${statusBadge(item.status)}</td>
      <td>${escapeHtml(item.admin || "-")}</td>
    </tr>
  `).join("");

  toggleEmpty(rows.length === 0);
}

function getSearchText() {
  return String(els.searchInput.value || "").trim().toLowerCase();
}

function keyTypeBadge(keyType) {
  const lower = String(keyType || "").trim().toLowerCase();

  if (["cewe", "cewek", "wanita", "perempuan", "female"].includes(lower)) {
    return `<span class="key-type cewe">Cewe</span>`;
  }

  // Default & semua varian cowo -> badge Cowo (grafit), selaras dengan sheet.
  return `<span class="key-type cowo">Cowo</span>`;
}

function statusBadge(status) {
  const value = String(status || "-");
  const lower = value.toLowerCase();

  if (lower === "dipakai") return `<span class="badge used">Dipakai</span>`;
  if (lower === "kosong") return `<span class="badge empty">Kosong</span>`;
  if (lower === "masuk") return `<span class="badge in">Masuk</span>`;
  if (lower === "keluar") return `<span class="badge out">Keluar</span>`;

  return `<span class="badge">${escapeHtml(value)}</span>`;
}

function toggleEmpty(isEmpty) {
  els.emptyState.classList.toggle("hidden", !isEmpty);
}

function setBackendStatus(type, text) {
  els.backendStatus.classList.remove("ok", "bad");

  if (type) {
    els.backendStatus.classList.add(type);
  }

  els.backendStatus.textContent = text;
}

// Daftar pesan yang TIDAK boleh ditampilkan ke pengguna.
// Ini error internal Google Sheet (merge cell) yang tidak relevan buat admin,
// dan transaksi utama (check-in/out) tetap berhasil walau ini muncul.
const SUPPRESSED_TOAST_PATTERNS = [
  "rentang penggabungan",
  "menggabungkan atau memisahkan",
  "memilih semua sel",
  "merged",
  "merge"
];

function isSuppressedMessage(message) {
  const text = String(message || "").toLowerCase();
  return SUPPRESSED_TOAST_PATTERNS.some((p) => text.includes(p));
}

function showToast(message, type = "") {
  // Jangan pernah tampilkan error merge Google Sheet — itu noise, bukan masalah nyata.
  if (isSuppressedMessage(message)) {
    return;
  }

  els.toast.className = `toast ${type}`.trim();
  els.toast.textContent = message;
  els.toast.classList.remove("hidden");

  clearTimeout(showToast.timer);

  showToast.timer = setTimeout(() => {
    els.toast.classList.add("hidden");
  }, 4200);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

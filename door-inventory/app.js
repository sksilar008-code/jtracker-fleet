/* ============================================================
   Door Inventory App — completely free, offline-first PWA
   Data lives only on your phone (IndexedDB). No server, no cost.
   ============================================================ */

const DB_NAME = "DoorInventoryDB";
const DB_VERSION = 1;
const STORE_NAME = "doors";

let db = null;
let doors = [];
let currentEditId = null;

// ---------- UI elements ----------
const els = {
  installBanner: document.getElementById("installBanner"),
  installBtn: document.getElementById("installBtn"),
  statusDot: document.getElementById("statusDot"),
  statusText: document.getElementById("statusText"),
  statTotalItems: document.getElementById("statTotalItems"),
  statTotalQty: document.getElementById("statTotalQty"),
  statTotalValue: document.getElementById("statTotalValue"),
  statLowStock: document.getElementById("statLowStock"),
  tbody: document.getElementById("tbody"),
  emptyState: document.getElementById("emptyState"),
  searchInput: document.getElementById("searchInput"),
  vendorFilter: document.getElementById("vendorFilter"),
  addBtn: document.getElementById("addBtn"),
  receiveBtn: document.getElementById("receiveBtn"),
  exportBtn: document.getElementById("exportBtn"),
  importFile: document.getElementById("importFile"),
  clearBtn: document.getElementById("clearBtn"),
  addModal: document.getElementById("addModal"),
  receiveModal: document.getElementById("receiveModal"),
  exportModal: document.getElementById("exportModal"),
  modalCloseButtons: document.querySelectorAll("[data-close-modal]"),
  addForm: document.getElementById("addForm"),
  receiveForm: document.getElementById("receiveForm"),
  csvOutput: document.getElementById("csvOutput"),
  downloadCsv: document.getElementById("downloadCsv"),
  closeExport: document.getElementById("closeExport"),
  toast: document.getElementById("toast"),
  deleteConfirmModal: document.getElementById("deleteConfirmModal"),
  confirmDeleteBtn: document.getElementById("confirmDeleteBtn"),
  cancelDeleteBtn: document.getElementById("cancelDeleteBtn"),
};

let itemToDelete = null;

// ---------- IndexedDB ----------
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const database = e.target.result;
      if (!database.objectStoreNames.contains(STORE_NAME)) {
        database.createObjectStore(STORE_NAME, { keyPath: "id", autoIncrement: true });
      }
    };
    req.onsuccess = (e) => {
      db = e.target.result;
      resolve(db);
    };
    req.onerror = (e) => reject(e.target.error);
  });
}

function getAllDoors() {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readonly");
    const store = tx.objectStore(STORE_NAME);
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.target.error);
  });
}

function addDoor(door) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    const store = tx.objectStore(STORE_NAME);
    const req = store.add(door);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.target.error);
  });
}

function updateDoor(id, changes) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    const store = tx.objectStore(STORE_NAME);
    const getReq = store.get(id);
    getReq.onsuccess = () => {
      const record = getReq.result;
      if (!record) return reject(new Error("Not found"));
      const updated = { ...record, ...changes, id };
      const putReq = store.put(updated);
      putReq.onsuccess = () => resolve(updated);
      putReq.onerror = () => reject(putReq.target.error);
    };
    getReq.onerror = () => reject(getReq.target.error);
  });
}

function deleteDoor(id) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    const store = tx.objectStore(STORE_NAME);
    const req = store.delete(id);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.target.error);
  });
}

function clearAllDoors() {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    const store = tx.objectStore(STORE_NAME);
    const req = store.clear();
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.target.error);
  });
}

// ---------- Helpers ----------
function money(n) {
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(n || 0);
}

function fmtDate(iso) {
  if (!iso) return "-";
  const d = new Date(iso);
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
}

function showToast(message, type = "info") {
  els.toast.textContent = message;
  els.toast.className = "toast show " + type;
  setTimeout(() => {
    els.toast.classList.remove("show");
  }, 2200);
}

function closeAllModals() {
  [els.addModal, els.receiveModal, els.exportModal, els.deleteConfirmModal].forEach((m) => {
    if (m) m.style.display = "none";
  });
  currentEditId = null;
  itemToDelete = null;
}

// ---------- Render ----------
async function loadDoors() {
  doors = await getAllDoors();
  render();
}

function render() {
  const search = els.searchInput.value.toLowerCase().trim();
  const vendor = els.vendorFilter.value.toLowerCase().trim();

  let filtered = doors.filter((d) => {
    const matchSearch =
      !search ||
      (d.modelNumber || "").toLowerCase().includes(search) ||
      (d.vendor || "").toLowerCase().includes(search) ||
      (d.size || "").toLowerCase().includes(search) ||
      (d.doorNumber || "").toLowerCase().includes(search) ||
      (d.location || "").toLowerCase().includes(search);
    const matchVendor = !vendor || (d.vendor || "").toLowerCase().includes(vendor);
    return matchSearch && matchVendor;
  });

  // Sort: newest first
  filtered.sort((a, b) => new Date(b.updatedAt || b.createdAt) - new Date(a.createdAt || b.updatedAt));

  // Stats
  const totalItems = doors.length;
  const totalQty = doors.reduce((s, d) => s + (Number(d.quantity) || 0), 0);
  const totalValue = doors.reduce(
    (s, d) => s + (Number(d.quantity) || 0) * (Number(d.pricePerUnit) || 0),
    0
  );
  const lowStock = doors.filter((d) => (Number(d.quantity) || 0) <= 5).length;

  els.statTotalItems.textContent = totalItems;
  els.statTotalQty.textContent = totalQty;
  els.statTotalValue.textContent = money(totalValue);
  els.statLowStock.textContent = lowStock;
  els.statLowStock.className = lowStock > 0 ? "value low" : "value";

  // Vendor filter options
  renderVendorOptions();

  // Table
  els.tbody.innerHTML = "";
  if (filtered.length === 0) {
    els.emptyState.style.display = "block";
    return;
  }
  els.emptyState.style.display = "none";

  filtered.forEach((door) => {
    const qty = Number(door.quantity) || 0;
    const price = Number(door.pricePerUnit) || 0;
    const value = qty * price;
    const isLow = qty <= 5;

    const tr = document.createElement("tr");
    if (isLow) tr.classList.add("low-stock-row");

    tr.innerHTML = `
      <td>
        <div class="model">${escapeHtml(door.modelNumber || "-")}</div>
        <div class="vendor">${escapeHtml(door.vendor || "-")}</div>
        ${door.doorNumber ? `<div class="hint" style="margin-top:2px;">Door #${escapeHtml(door.doorNumber)}</div>` : ""}
      </td>
      <td class="vendor">${escapeHtml(door.size || "-")}</td>
      <td class="qty">${qty}</td>
      <td class="price">${money(price)}</td>
      <td class="total-value">${money(value)}</td>
      <td>${escapeHtml(door.location || "-")}</td>
      <td class="actions">
        <button class="btn btn-icon btn-ghost" title="Edit" onclick="app.editDoor(${door.id})" aria-label="Edit ${escapeHtml(door.modelNumber || "")}">✏️</button>
        <button class="btn btn-icon btn-ghost" title="Receive more stock" onclick="app.receiveStock(${door.id})" aria-label="Add stock ${escapeHtml(door.modelNumber || "")}">➕</button>
        <button class="btn btn-icon btn-danger" title="Delete" onclick="app.requestDelete(${door.id})" aria-label="Delete ${escapeHtml(door.modelNumber || "")}">🗑️</button>
      </td>
    `;
    els.tbody.appendChild(tr);
  });
}

function renderVendorOptions() {
  const vendors = [...new Set(doors.map((d) => d.vendor).filter(Boolean))].sort();
  els.vendorFilter.innerHTML = '<option value="">All vendors</option>';
  vendors.forEach((v) => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    els.vendorFilter.appendChild(opt);
  });
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

// ---------- Actions ----------
window.app = {
  openAdd: () => {
    currentEditId = null;
    document.getElementById("modalTitle").textContent = "Add Door Stock";
    document.getElementById("doorId").value = "";
    els.addForm.reset();
    document.getElementById("createdAtField").style.display = "none";
    els.addModal.style.display = "flex";
  },

  openReceive: (id = null) => {
    if (id) {
      const door = doors.find((d) => d.id === id);
      if (!door) return;
      document.getElementById("receiveModalTitle").textContent = "Receive Stock - " + (door.modelNumber || "");
      document.getElementById("receiveDoorId").value = door.id;
      document.getElementById("receiveModel").value = door.modelNumber || "";
      document.getElementById("receiveVendor").value = door.vendor || "";
      document.getElementById("receiveSize").value = door.size || "";
      document.getElementById("receiveDoorNumber").value = door.doorNumber || "";
      document.getElementById("receiveLocation").value = door.location || "";
      document.getElementById("receiveQty").value = "";
      document.getElementById("receivePrice").value = door.pricePerUnit || "";
      document.getElementById("receiveNote").value = "";
      document.getElementById("receiveCreatedAtField").style.display = "none";
    } else {
      document.getElementById("receiveModalTitle").textContent = "New Vendor Bill";
      document.getElementById("receiveDoorId").value = "";
      els.receiveForm.reset();
      document.getElementById("receiveCreatedAtField").style.display = "none";
    }
    els.receiveModal.style.display = "flex";
  },

  editDoor: (id) => {
    const door = doors.find((d) => d.id === id);
    if (!door) return;
    currentEditId = id;
    document.getElementById("modalTitle").textContent = "Edit Door";
    document.getElementById("doorId").value = door.id;
    document.getElementById("addModel").value = door.modelNumber || "";
    document.getElementById("addVendor").value = door.vendor || "";
    document.getElementById("addSize").value = door.size || "";
    document.getElementById("addQty").value = door.quantity || "";
    document.getElementById("addPrice").value = door.pricePerUnit || "";
    document.getElementById("addLocation").value = door.location || "";
    document.getElementById("addDoorNumber").value = door.doorNumber || "";
    document.getElementById("createdAtValue").textContent = fmtDate(door.createdAt);
    document.getElementById("createdAtField").style.display = "block";
    els.addModal.style.display = "flex";
  },

  receiveStock: (id) => {
    window.app.openReceive(id);
  },

  requestDelete: (id) => {
    itemToDelete = id;
    els.deleteConfirmModal.style.display = "flex";
  },

  confirmDelete: async () => {
    if (itemToDelete == null) return;
    try {
      await deleteDoor(itemToDelete);
      await loadDoors();
      showToast("Deleted");
    } catch (e) {
      showToast("Delete failed: " + e.message, "error");
    }
    closeAllModals();
  },
};

// ---------- Form handlers ----------
els.addForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const id = document.getElementById("doorId").value;
  const door = {
    modelNumber: document.getElementById("addModel").value.trim(),
    vendor: document.getElementById("addVendor").value.trim(),
    size: document.getElementById("addSize").value.trim(),
    quantity: Number(document.getElementById("addQty").value) || 0,
    pricePerUnit: Number(document.getElementById("addPrice").value) || 0,
    location: document.getElementById("addLocation").value.trim(),
    doorNumber: document.getElementById("addDoorNumber").value.trim(),
    updatedAt: new Date().toISOString(),
  };
  if (!id) {
    door.createdAt = new Date().toISOString();
    try {
      await addDoor(door);
      showToast("Added to inventory");
    } catch (err) {
      showToast("Save failed: " + err.message, "error");
      return;
    }
  } else {
    try {
      await updateDoor(Number(id), door);
      showToast("Updated");
    } catch (err) {
      showToast("Update failed: " + err.message, "error");
      return;
    }
  }
  closeAllModals();
  await loadDoors();
});

els.receiveForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const id = document.getElementById("receiveDoorId").value;
  const qty = Number(document.getElementById("receiveQty").value) || 0;
  const price = Number(document.getElementById("receivePrice").value) || 0;
  const note = document.getElementById("receiveNote").value.trim();

  if (!id) {
    // Creating new from a vendor bill
    const door = {
      modelNumber: document.getElementById("receiveModel").value.trim(),
      vendor: document.getElementById("receiveVendor").value.trim(),
      size: document.getElementById("receiveSize").value.trim(),
      quantity: qty,
      pricePerUnit: price,
      location: document.getElementById("receiveLocation").value.trim(),
      doorNumber: document.getElementById("receiveDoorNumber").value.trim(),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    if (!door.modelNumber || !door.size) {
      showToast("Model number and size are required", "error");
      return;
    }
    try {
      await addDoor(door);
      showToast("New stock received");
    } catch (err) {
      showToast("Save failed: " + err.message, "error");
      return;
    }
  } else {
    // Adding to existing item (vendor bill received)
    const door = doors.find((d) => d.id === Number(id));
    if (!door) {
      showToast("Item not found", "error");
      return;
    }
    try {
      await updateDoor(Number(id), {
        quantity: (Number(door.quantity) || 0) + qty,
        pricePerUnit: price || door.pricePerUnit,
        vendor: document.getElementById("receiveVendor").value.trim() || door.vendor,
        size: document.getElementById("receiveSize").value.trim() || door.size,
        location: document.getElementById("receiveLocation").value.trim() || door.location,
        doorNumber: document.getElementById("receiveDoorNumber").value.trim() || door.doorNumber,
        updatedAt: new Date().toISOString(),
      });
      showToast(`+${qty} received`);
    } catch (err) {
      showToast("Update failed: " + err.message, "error");
      return;
    }
  }
  closeAllModals();
  await loadDoors();
});

// ---------- Export / Import ----------
els.exportBtn.addEventListener("click", () => {
  const csv = doorsToCSV();
  els.csvOutput.textContent = csv;
  els.downloadCsv.href = "data:text/csv;charset=utf-8," + encodeURIComponent(csv);
  els.downloadCsv.download = `door-inventory-${new Date().toISOString().slice(0,10)}.csv`;
  els.exportModal.style.display = "flex";
});

els.closeExport.addEventListener("click", () => {
  els.exportModal.style.display = "none";
});

function doorsToCSV() {
  const headers = ["Model Number","Vendor","Size","Quantity","Price Per Unit","Location","Door Number","Created","Updated"];
  const rows = doors.map((d) => [
    d.modelNumber || "",
    d.vendor || "",
    d.size || "",
    d.quantity || "",
    d.pricePerUnit || "",
    d.location || "",
    d.doorNumber || "",
    d.createdAt || "",
    d.updatedAt || "",
  ]);
  const escape = (v) => (String(v).includes(",") || String(v).includes('"') ? `"${String(v).replace(/"/g, '""')}"` : v);
  return [headers.map(escape), ...rows.map((r) => r.map(escape))].map((r) => r.join(",")).join("\n");
}

els.importFile.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const text = await file.text();
  try {
    const lines = text.trim().split("\n");
    if (lines.length < 2) throw new Error("Empty CSV");
    const headers = lines[0].split(",");
    const imported = lines.slice(1).map((line) => {
      const vals = line.split(",");
      const obj = {};
      headers.forEach((h, i) => { obj[h.trim()] = vals[i] ? vals[i].trim() : ""; });
      obj.quantity = Number(obj.Quantity || obj.quantity || 0);
      obj.pricePerUnit = Number(obj["Price Per Unit"] || obj.pricePerUnit || 0);
      obj.createdAt = obj.Created || new Date().toISOString();
      obj.updatedAt = obj.Updated || new Date().toISOString();
      return obj;
    });
    for (const d of imported) await addDoor(d);
    els.importFile.value = "";
    await loadDoors();
    showToast(`${imported.length} items imported`);
  } catch (err) {
    showToast("Import failed: " + err.message, "error");
  }
});

els.clearBtn.addEventListener("click", async () => {
  if (!confirm("Delete ALL inventory? This cannot be undone.")) return;
  try {
    await clearAllDoors();
    await loadDoors();
    showToast("Inventory cleared");
  } catch (err) {
    showToast("Clear failed: " + err.message, "error");
  }
});

// ---------- Modal close ----------
els.modalCloseButtons.forEach((btn) => {
  btn.addEventListener("click", closeAllModals);
});

document.addEventListener("click", (e) => {
  if (e.target.classList.contains("modal-overlay")) closeAllModals();
});

els.cancelDeleteBtn.addEventListener("click", closeAllModals);
els.confirmDeleteBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  window.app.confirmDelete();
});

// ---------- Search / filter ----------
els.searchInput.addEventListener("input", render);
els.vendorFilter.addEventListener("change", render);

// ---------- Offline / Online ----------
function updateOnlineStatus() {
  const online = navigator.onLine;
  els.statusDot.className = "status-dot" + (online ? "" : " offline");
  els.statusText.textContent = online ? "Online" : "Offline — data saved on phone";
}
window.addEventListener("online", updateOnlineStatus);
window.addEventListener("offline", updateOnlineStatus);
updateOnlineStatus();

// ---------- PWA install prompt ----------
let deferredInstallPrompt = null;
window.addEventListener("beforeinstallprompt", (e) => {
  e.preventDefault();
  deferredInstallPrompt = e;
  els.installBanner.style.display = "flex";
});

els.installBtn.addEventListener("click", async () => {
  if (!deferredInstallPrompt) return;
  deferredInstallPrompt.prompt();
  const choice = await deferredInstallPrompt.userChoice;
  deferredInstallPrompt = null;
  els.installBanner.style.display = "none";
  if (choice.outcome === "accepted") showToast("Installed on home screen");
});

// ---------- Boot ----------
openDB()
  .then(() => loadDoors())
  .catch((err) => {
    showToast("Database error: " + err.message, "error");
    console.error(err);
  });

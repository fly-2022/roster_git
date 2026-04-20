/* ================= EXCEL MAIN TEMPLATE SYSTEM ================= */

let excelWorkbook = null;
let excelData = {};
let currentMode = "arrival";
let currentShift = "morning";
let currentLane = "car";
let currentColor = "#4CAF50";
let isDragging = false;
let dragMode = "add";
let tableEventsAttached = false;
let otGlobalCounter = 1;
let sosGlobalCounter = 1;
let historyStack = [];
let _saveStateGlobal = null; // set by DOMContentLoaded, used by attachTableEvents

// Cargo variation state
let cargoWorkbooks = { arrival: null, departure: null };
let cargoVariations = { arrival: [], departure: [] }; // list of sheet names
let currentCargoVariation = null; // currently selected sheet name

// Out-of-service counters: Set of "zoneName:counter" strings
const oosCounters = new Set();

function resetDragState() {
    isDragging = false;
    dragMode = "add";
}

// cellStates keyed by "lane_mode_shift"
const cellStates = {};
["car", "bus", "train", "cargo", "owc"].forEach(lane =>
    ["arrival", "departure"].forEach(mode =>
        ["morning", "night"].forEach(shift => {
            cellStates[`${lane}_${mode}_${shift}`] = {};
        })
    )
);

const table = document.getElementById("rosterTable");
const summary = document.getElementById("summary");
const manningSummary = document.getElementById("manningSummary");
const modeHighlight = document.getElementById("modeHighlight");
const shiftHighlight = document.getElementById("shiftHighlight");

const arrivalBtn = document.getElementById("arrivalBtn");
const departureBtn = document.getElementById("departureBtn");
const morningBtn = document.getElementById("morningBtn");
const nightBtn = document.getElementById("nightBtn");

function range(prefix, start, end) {
    const arr = [];
    for (let i = start; i <= end; i++) arr.push(prefix + i);
    return arr;
}

const zones = {
    // ── CAR ──────────────────────────────────────────────────────────────────
    car_arrival: [
        { name: "Zone 1", counters: range("AC", 1, 10) },
        { name: "Zone 2", counters: range("AC", 11, 20) },
        { name: "Zone 3", counters: range("AC", 21, 30) },
        { name: "Zone 4", counters: range("AC", 31, 40) },
        { name: "BIKES", counters: ["AM41", "AM43"] }
    ],
    car_departure: [
        { name: "Zone 1", counters: range("DC", 1, 8) },
        { name: "Zone 2", counters: range("DC", 9, 18) },
        { name: "Zone 3", counters: range("DC", 19, 28) },
        { name: "Zone 4", counters: range("DC", 29, 36) },
        { name: "BIKES", counters: ["DM37A", "DM37C"] }
    ],

    // ── BUS ───────────────────────────────────────────────────────────────────
    // Clusters zone: each cluster (A-D) is one counter
    // Bus Lanes zone: each lane (1-5) is one counter
    bus_arrival: [
        { name: "Arrival Clusters", counters: ["Arrival Cluster A", "Arrival Cluster B", "Arrival Cluster C", "Arrival Cluster D"] },
        { name: "Arrival Bus Lanes", counters: ["ABL 1", "ABL 2", "ABL 3", "ABL 4", "ABL 5"] }
    ],
    bus_departure: [
        { name: "Departure Clusters", counters: ["Departure Cluster A", "Departure Cluster B", "Departure Cluster C", "Departure Cluster D"] },
        { name: "Departure Bus Lanes", counters: ["DBL 1", "DBL 2", "DBL 3", "DBL 4", "DBL 5"] }
    ],

    // ── TRAIN ─────────────────────────────────────────────────────────────────
    // All counters in one zone
    train_arrival: [
        { name: "Train", counters: ["KIOSK 4", "KIOSK 6", "EIACS", "TR6", "TR7", "TR8", "TR9", "TR10"] }
    ],
    train_departure: [
        { name: "Train", counters: ["KIOSK 4", "KIOSK 6", "EIACS", "TR6", "TR7", "TR8", "TR9", "TR10"] }
    ],

    // ── CARGO ─────────────────────────────────────────────────────────────────
    // Single zone, AL/DL lanes and Cargo counters mixed together
    cargo_arrival: [
        { name: "Cargo", counters: ["AL1", "AL2", "AL3", "AL4", "AL5", "AL6", "A-Cargo 1", "A-Cargo 2", "A-Cargo 3", "A-Cargo 4", "A-Cargo 5", "A-Cargo 6"] }
    ],
    cargo_departure: [
        { name: "Cargo", counters: ["DL1", "DL2", "DL3", "DL4", "DL5", "DL6", "D-Cargo 1", "D-Cargo 2", "D-Cargo 3", "D-Cargo 4", "D-Cargo 5", "D-Cargo 6"] }
    ],

    // ── OWC ───────────────────────────────────────────────────────────────────
    owc_arrival: [
        { name: "OWC", counters: range("AL", 7, 10) }
    ],
    owc_departure: [
        { name: "OWC", counters: range("DL", 7, 10) }
    ]
};

/* ---------------- COLOR PICKER ----------------- */
document.querySelectorAll(".color-btn").forEach(btn => {
    btn.addEventListener("click", () => {
        currentColor = btn.dataset.color;
        document.querySelectorAll(".color-btn").forEach(b => b.classList.remove("selected"));
        btn.classList.add("selected");
    });
});

function generateTimeSlots() {
    const slots = [];
    let start, end;

    if (currentShift === "morning") {
        start = 10 * 60;
        end = 22 * 60;
    } else {
        start = 22 * 60;
        end = (24 + 10) * 60;
    }

    for (let time = start; time < end; time += 15) {
        let minutes = time % (24 * 60);
        let hh = Math.floor(minutes / 60);
        let mm = minutes % 60;
        let hhmm = String(hh).padStart(2, "0") + String(mm).padStart(2, "0");
        slots.push(hhmm);
    }

    return slots;
}

/* ==================== renderTableOnce ==================== */
function renderTableOnce() {
    // Cargo lane uses its own dynamic renderer
    if (currentLane === "cargo") {
        const tabBar = document.getElementById("cargoVariationTabs");
        if (tabBar) tabBar.style.display = "flex";
        renderCargoGrid();
        return;
    }
    // Hide cargo tab bar for non-cargo lanes
    const tabBar = document.getElementById("cargoVariationTabs");
    if (tabBar) tabBar.style.display = "none";

    table.innerHTML = "";
    const times = generateTimeSlots();

    zones[currentLane + "_" + currentMode].forEach(zone => {
        let zoneRow = document.createElement("tr");
        let zoneCell = document.createElement("td");
        zoneCell.colSpan = times.length + 1;
        zoneCell.className = "zone-header";
        zoneCell.innerText = zone.name;
        zoneRow.appendChild(zoneCell);
        table.appendChild(zoneRow);

        let timeRow = document.createElement("tr");
        timeRow.className = "time-header";
        timeRow.innerHTML = "<th></th>";
        times.forEach(t => {
            let th = document.createElement("th");
            th.innerText = t;
            timeRow.appendChild(th);
        });
        table.appendChild(timeRow);

        zone.counters.forEach(counter => {
            let row = document.createElement("tr");
            let label = document.createElement("td");
            label.innerText = counter;
            row.appendChild(label);

            times.forEach((t, i) => {
                let cell = document.createElement("td");
                cell.className = "counter-cell";
                cell.dataset.zone = zone.name;
                cell.dataset.time = i;
                cell.dataset.counter = counter;
                row.appendChild(cell);
            });

            table.appendChild(row);
        });

        let subtotalRow = document.createElement("tr");
        subtotalRow.className = "subtotal-row";
        let subtotalLabel = document.createElement("td");
        subtotalLabel.innerText = "Subtotal";
        subtotalRow.appendChild(subtotalLabel);

        times.forEach((t, i) => {
            let td = document.createElement("td");
            td.className = "subtotal-cell";
            td.dataset.zone = zone.name;
            td.dataset.time = i;
            subtotalRow.appendChild(td);
        });

        table.appendChild(subtotalRow);
    });
}

/* ==================== OOS & Counter Swap ==================== */

function oosKey(zone, counter) { return `${zone}:${counter}`; }

// swapHistory: "zone:fromCounter" → "zone:toCounter"
const swapHistory = new Map();

function applyOOSStyling() {
    document.querySelectorAll("#rosterTable .counter-cell").forEach(cell => {
        const key = oosKey(cell.dataset.zone, cell.dataset.counter);
        if (oosCounters.has(key)) {
            if (!cell.classList.contains("active")) cell.style.background = "#ffcccc";
            cell.style.pointerEvents = "none";
        } else {
            if (!cell.classList.contains("active")) cell.style.background = "";
            cell.style.pointerEvents = "";
        }
    });

    // Build a reverse lookup: "toZone||toCounter" → "fromZone:fromCounter"
    const swapDest = new Map();
    swapHistory.forEach((dest, src) => swapDest.set(dest, src));

    document.querySelectorAll("#rosterTable tr").forEach(row => {
        if (row.classList.contains("zone-header") ||
            row.classList.contains("time-header") ||
            row.classList.contains("subtotal-row") ||
            row.classList.contains("grandtotal-row")) return;

        const labelCell = row.cells?.[0];
        const dataCell = row.querySelector(".counter-cell");
        if (!labelCell || !dataCell) return;

        const zone = dataCell.dataset.zone;
        const counter = dataCell.dataset.counter;
        const srcKey = `${zone}:${counter}`;          // key used in swapHistory
        const destKey = `${zone}||${counter}`;          // value stored in swapHistory

        const isOOS = oosCounters.has(oosKey(zone, counter));
        const isSwapSrc = swapHistory.has(srcKey);
        const isSwapDest = swapDest.has(destKey);

        // Strip all previous badges cleanly
        labelCell.textContent = labelCell.textContent
            .replace(/ ⚠/g, "").replace(/ ⇄.*/g, "").replace(/ ←.*/g, "");
        labelCell.removeAttribute("title");

        if (isOOS) {
            labelCell.style.cssText = "background:#e53935;color:white;font-weight:bold;";
            labelCell.textContent += " ⚠";
        } else if (isSwapSrc) {
            // Show where this counter's officers went
            const destRaw = swapHistory.get(srcKey);
            const sepIdx = destRaw.indexOf("||");
            const toZone = destRaw.slice(0, sepIdx);
            const toCounter = destRaw.slice(sepIdx + 2);
            const label = toZone === zone ? toCounter : `${toCounter} (${toZone})`;
            labelCell.style.cssText = "background:#fff3e0;color:#e65100;font-weight:bold;cursor:help;";
            labelCell.textContent += ` ⇄ ${label}`;
            labelCell.title = `Officers swapped to ${toCounter} in ${toZone}. Right-click to swap back.`;
        } else if (isSwapDest) {
            // Show where this counter received officers from
            const srcRaw = swapDest.get(destKey); // "fromZone:fromCounter"
            const lastColon = srcRaw.lastIndexOf(":");
            const fromZone = srcRaw.slice(0, lastColon);
            const fromCounter = srcRaw.slice(lastColon + 1);
            const label = fromZone === zone ? fromCounter : `${fromCounter} (${fromZone})`;
            labelCell.style.cssText = "background:#e8f5e9;color:#2e7d32;font-weight:bold;cursor:help;";
            labelCell.textContent += ` ← ${label}`;
            labelCell.title = `Received officers from ${fromCounter} in ${fromZone}.`;
        } else {
            labelCell.style.cssText = "";
        }
    });
}

// Free candidates sorted by highest counter number first (back counters)
function getFreeCandidates(zoneName, excludeCounter, fromIdx, toIdx) {
    const times = generateTimeSlots();
    const start = (fromIdx !== undefined) ? fromIdx : 0;
    const end = (toIdx !== undefined) ? toIdx : times.length;

    // Search ALL zones (including BIKES) across the current mode
    const results = [];
    zones[currentLane + "_" + currentMode].forEach(zone => {
        zone.counters.forEach(c => {
            if (zone.name === zoneName && c === excludeCounter) return;
            if (oosCounters.has(oosKey(zone.name, c))) return;
            // Must be free for the entire requested window
            for (let i = start; i < end; i++) {
                const cell = document.querySelector(
                    `.counter-cell[data-zone="${zone.name}"][data-time="${i}"][data-counter="${c}"]`);
                if (cell && cell.classList.contains("active")) return;
            }
            results.push({ zone: zone.name, counter: c });
        });
    });

    // Sort: same zone first, then by counter number desc
    results.sort((a, b) => {
        if (a.zone === zoneName && b.zone !== zoneName) return -1;
        if (b.zone === zoneName && a.zone !== zoneName) return 1;
        if (a.zone !== b.zone) return a.zone.localeCompare(b.zone);
        return (parseInt(b.counter.replace(/\D/g, "")) || 0) - (parseInt(a.counter.replace(/\D/g, "")) || 0);
    });

    return results; // array of { zone, counter }
}

function markOOS(zoneName, counter) {
    const key = oosKey(zoneName, counter);
    const activeCells = [...document.querySelectorAll(
        `.counter-cell.active[data-zone="${zoneName}"][data-counter="${counter}"]`
    )];
    const hasOfficers = activeCells.length > 0;
    const labels = [...new Set(activeCells.map(c => c.dataset.officer).filter(Boolean))];

    // Offer all free counters across zones for OOS relocation
    const freeCandidates = getFreeCandidates(zoneName, counter);

    function doMarkOOS(toZone, toCounter) {
        oosCounters.add(key);

        if (toZone && toCounter) {
            const officerSlots = new Map();
            activeCells.forEach(cell => {
                const lbl = cell.dataset.officer;
                if (!officerSlots.has(lbl)) officerSlots.set(lbl, []);
                officerSlots.get(lbl).push({
                    t: parseInt(cell.dataset.time),
                    color: cell.style.background,
                    type: cell.dataset.type
                });
            });
            // Clear OOS counter
            document.querySelectorAll(`.counter-cell[data-zone="${zoneName}"][data-counter="${counter}"]`)
                .forEach(cell => {
                    cell.classList.remove("active");
                    cell.dataset.officer = "";
                    cell.dataset.type = "";
                });
            // Move officers to target
            officerSlots.forEach((slots, label) => {
                const ok = !slots.some(({ t }) => {
                    const c = document.querySelector(
                        `.counter-cell[data-zone="${toZone}"][data-time="${t}"][data-counter="${toCounter}"]`
                    );
                    return !c || c.classList.contains("active");
                });
                if (ok) {
                    slots.forEach(({ t, color, type }) => {
                        const c = document.querySelector(
                            `.counter-cell[data-zone="${toZone}"][data-time="${t}"][data-counter="${toCounter}"]`
                        );
                        if (!c) return;
                        c.classList.add("active");
                        c.style.background = color;
                        c.dataset.officer = label;
                        c.dataset.type = type;
                    });
                }
            });
            swapHistory.set(`${zoneName}:${counter}`, `${toZone}||${toCounter}`);
        }

        applyOOSStyling();
        updateAll();
    }

    if (!hasOfficers) { doMarkOOS(null, null); return; }

    // Build dialog
    const overlay = document.createElement("div");
    overlay.style.cssText = `position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:10000;
        display:flex;align-items:center;justify-content:center;`;
    const box = document.createElement("div");
    box.style.cssText = `background:#fff;border-radius:10px;padding:24px 26px;max-width:380px;width:90%;
        box-shadow:0 8px 32px rgba(0,0,0,.25);font-family:Arial;`;
    const officerList = labels.map(l => `<strong>${l}</strong>`).join(", ");
    const byZone = {};
    freeCandidates.forEach(({ zone, counter: c }) => {
        if (!byZone[zone]) byZone[zone] = [];
        byZone[zone].push(c);
    });
    const candidateOptions = Object.entries(byZone).map(([zone, ctrs]) =>
        `<optgroup label="${zone}">${ctrs.map(c => `<option value="${zone}||${c}">${c}</option>`).join("")
        }</optgroup>`
    ).join("");

    box.innerHTML = `
        <div style="font-size:22px;margin-bottom:10px">⚠️</div>
        <h3 style="margin:0 0 10px;font-size:15px">Mark ${counter} Out of Service?</h3>
        <p style="margin:0 0 14px;font-size:13px;color:#555;line-height:1.5">
            ${labels.length} officer(s) assigned here: ${officerList}</p>
        ${freeCandidates.length ? `
        <label style="font-size:12px;color:#777;display:block;margin-bottom:4px">
            Relocate to (nearest free back counter first):</label>
        <select id="_oosMoveTarget" style="width:100%;padding:7px;border:1px solid #ccc;
            border-radius:6px;font-size:13px;margin-bottom:14px">${candidateOptions}</select>
        <div style="display:flex;flex-direction:column;gap:8px">
            <button id="_oosMove" style="padding:9px;border:none;background:#4CAF50;color:#fff;
                border-radius:6px;cursor:pointer;font-size:13px">✅ Mark OOS &amp; Move Officers</button>
            <button id="_oosKeep" style="padding:9px;border:none;background:#FF9800;color:#fff;
                border-radius:6px;cursor:pointer;font-size:13px">⚠️ Mark OOS Only (reassign manually)</button>
            <button id="_oosCancel" style="padding:9px;border:1px solid #ccc;background:#f5f5f5;
                border-radius:6px;cursor:pointer;font-size:13px">Cancel</button>
        </div>` : `
        <p style="margin:0 0 14px;font-size:13px;color:#c62828">No free counters available.</p>
        <div style="display:flex;flex-direction:column;gap:8px">
            <button id="_oosKeep" style="padding:9px;border:none;background:#FF9800;color:#fff;
                border-radius:6px;cursor:pointer;font-size:13px">⚠️ Mark OOS (reassign manually)</button>
            <button id="_oosCancel" style="padding:9px;border:1px solid #ccc;background:#f5f5f5;
                border-radius:6px;cursor:pointer;font-size:13px">Cancel</button>
        </div>`}`;

    overlay.appendChild(box);
    document.body.appendChild(overlay);
    if (freeCandidates.length) {
        document.getElementById("_oosMove").onclick = () => {
            const val = document.getElementById("_oosMoveTarget").value;
            const sepIdx = val.indexOf("||");
            const tz = val.slice(0, sepIdx);
            const tc = val.slice(sepIdx + 2);
            overlay.remove();
            doMarkOOS(tz, tc);
        };
    }
    document.getElementById("_oosKeep").onclick = () => { overlay.remove(); doMarkOOS(null, null); };
    document.getElementById("_oosCancel").onclick = () => overlay.remove();
    overlay.onclick = e => { if (e.target === overlay) overlay.remove(); };
}

function clearOOS(zoneName, counter) {
    oosCounters.delete(oosKey(zoneName, counter));
    document.querySelectorAll(`.counter-cell[data-zone="${zoneName}"][data-counter="${counter}"]`)
        .forEach(cell => {
            if (!cell.classList.contains("active")) cell.style.background = "";
            cell.style.pointerEvents = "";
        });
    applyOOSStyling();
    updateAll();
}

function swapCounters(fromZone, fromCounter, toZone, toCounter, fromIdx, toIdx) {
    if (oosCounters.has(oosKey(toZone, toCounter))) {
        alert(`Counter ${toCounter} is Out of Service. Cannot swap into it.`); return;
    }
    const times = generateTimeSlots();
    const start = (fromIdx !== undefined) ? fromIdx : 0;
    const end = (toIdx !== undefined) ? toIdx : times.length;

    // Check target is free for the requested window
    for (let i = start; i < end; i++) {
        const cell = document.querySelector(
            `.counter-cell[data-zone="${toZone}"][data-time="${i}"][data-counter="${toCounter}"]`);
        if (cell && cell.classList.contains("active")) {
            alert(`Counter ${toCounter} is occupied in the selected time window. Cannot swap.`); return;
        }
    }

    for (let i = start; i < end; i++) {
        const fromCell = document.querySelector(
            `.counter-cell[data-zone="${fromZone}"][data-time="${i}"][data-counter="${fromCounter}"]`);
        const toCell = document.querySelector(
            `.counter-cell[data-zone="${toZone}"][data-time="${i}"][data-counter="${toCounter}"]`);
        if (!fromCell || !toCell) continue;
        if (fromCell.classList.contains("active")) {
            toCell.classList.add("active");
            toCell.style.background = fromCell.style.background;
            toCell.dataset.officer = fromCell.dataset.officer;
            toCell.dataset.type = fromCell.dataset.type;
            fromCell.classList.remove("active");
            fromCell.style.background = "";
            fromCell.dataset.officer = "";
            fromCell.dataset.type = "";
        }
    }
    swapHistory.set(`${fromZone}:${fromCounter}`, `${toZone}||${toCounter}`);
    applyOOSStyling();
    updateAll();
}

function swapBack(zoneName, fromCounter) {
    const histKey = `${zoneName}:${fromCounter}`;
    const destFull = swapHistory.get(histKey);
    if (!destFull) return;

    // destFull format: "toZone||toCounter"
    const sepIdx = destFull.indexOf("||");
    const toZone = destFull.slice(0, sepIdx);
    const toCounter = destFull.slice(sepIdx + 2);

    if (oosCounters.has(oosKey(zoneName, fromCounter))) {
        alert(`${fromCounter} is still marked Out of Service. Clear OOS first.`); return;
    }

    const times = generateTimeSlots();
    let moved = 0;
    times.forEach((t, i) => {
        const srcCell = document.querySelector(
            `.counter-cell[data-zone="${toZone}"][data-time="${i}"][data-counter="${toCounter}"]`);
        const dstCell = document.querySelector(
            `.counter-cell[data-zone="${zoneName}"][data-time="${i}"][data-counter="${fromCounter}"]`);
        if (!srcCell || !dstCell) return;
        if (srcCell.classList.contains("active") && !dstCell.classList.contains("active")) {
            dstCell.classList.add("active");
            dstCell.style.background = srcCell.style.background;
            dstCell.dataset.officer = srcCell.dataset.officer;
            dstCell.dataset.type = srcCell.dataset.type;
            srcCell.classList.remove("active");
            srcCell.style.background = "";
            srcCell.dataset.officer = "";
            srcCell.dataset.type = "";
            moved++;
        }
    });

    swapHistory.delete(histKey);
    applyOOSStyling();
    updateAll();
}

function buildCounterContextMenu(zoneName, counter) {
    const existing = document.getElementById("_counterCtxMenu");
    if (existing) existing.remove();

    const menu = document.createElement("div");
    menu.id = "_counterCtxMenu";
    menu.style.cssText = `position:fixed;z-index:9999;background:#fff;border:1px solid #ccc;
        border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,.18);padding:4px 0;min-width:200px;font-size:13px;`;

    const isOOS = oosCounters.has(oosKey(zoneName, counter));
    const histKey = `${zoneName}:${counter}`;
    const isSwapSrc = swapHistory.has(histKey);
    const destRaw = isSwapSrc ? swapHistory.get(histKey) : null;
    const destLabel = destRaw
        ? (() => { const s = destRaw.indexOf("||"); const z = destRaw.slice(0, s); const c = destRaw.slice(s + 2); return z === zoneName ? c : `${c} (${z})`; })()
        : null;

    const items = [];
    if (isOOS) {
        items.push({ label: "✅ Clear Out of Service", action: () => clearOOS(zoneName, counter) });
    } else {
        items.push({
            label: "⚠️ Mark Out of Service", action: () => {
                if (_saveStateGlobal) _saveStateGlobal();
                markOOS(zoneName, counter);
            }
        });
    }

    if (isSwapSrc && destLabel) {
        items.push({
            label: `↩️ Swap Back from ${destLabel}`, action: () => {
                if (_saveStateGlobal) _saveStateGlobal();
                swapBack(zoneName, counter);
            }
        });
    } else {
        items.push({
            label: "🔄 Swap Counter →", action: () => {
                if (_saveStateGlobal) _saveStateGlobal();
                showSwapDialog(zoneName, counter);
            }
        });
    }

    items.forEach(({ label, action }) => {
        const item = document.createElement("div");
        item.textContent = label;
        item.style.cssText = "padding:8px 14px;cursor:pointer;white-space:nowrap;";
        item.onmouseenter = () => item.style.background = "#f0f4ff";
        item.onmouseleave = () => item.style.background = "";
        item.onclick = () => { menu.remove(); action(); };
        menu.appendChild(item);
    });

    document.body.appendChild(menu);
    return menu;
}

function showSwapDialog(zoneName, fromCounter) {
    const existing = document.getElementById("_swapDialog");
    if (existing) existing.remove();

    const times = generateTimeSlots();

    // Default time range = first to last occupied slot on fromCounter
    const activeTimes = [...document.querySelectorAll(
        `.counter-cell.active[data-zone="${zoneName}"][data-counter="${fromCounter}"]`
    )].map(c => parseInt(c.dataset.time)).sort((a, b) => a - b);

    const defaultFrom = activeTimes.length ? activeTimes[0] : 0;
    const defaultTo = activeTimes.length ? activeTimes[activeTimes.length - 1] + 1 : times.length;

    const overlay = document.createElement("div");
    overlay.id = "_swapDialog";
    overlay.style.cssText = `position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:10000;
        display:flex;align-items:center;justify-content:center;`;

    const box = document.createElement("div");
    box.style.cssText = `background:#fff;border-radius:10px;padding:24px;min-width:340px;max-width:420px;
        box-shadow:0 8px 32px rgba(0,0,0,.25);font-family:Arial;`;

    const timeOptions = times.map((t, i) => `<option value="${i}">${t}</option>`).join("");

    box.innerHTML = `
        <h3 style="margin:0 0 16px;font-size:15px">🔄 Swap Counter</h3>
        <p style="margin:0 0 12px;font-size:13px;color:#555">
            Moving officers from <strong>${fromCounter}</strong> (${zoneName})</p>

        <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px">
            <div>
                <label style="font-size:12px;color:#777;display:block;margin-bottom:4px">From time</label>
                <select id="_swapFrom" style="width:100%;padding:6px;border:1px solid #ccc;border-radius:6px;font-size:13px">
                    ${timeOptions}
                </select>
            </div>
            <div>
                <label style="font-size:12px;color:#777;display:block;margin-bottom:4px">To time</label>
                <select id="_swapTo" style="width:100%;padding:6px;border:1px solid #ccc;border-radius:6px;font-size:13px">
                    ${timeOptions}
                </select>
            </div>
        </div>

        <label style="font-size:12px;color:#777;display:block;margin-bottom:4px">Swap to counter</label>
        <select id="_swapTarget" style="width:100%;padding:7px;border:1px solid #ccc;border-radius:6px;font-size:13px;margin-bottom:6px">
            <option disabled value="">— select target counter —</option>
        </select>
        <p id="_swapNote" style="margin:0 0 16px;font-size:11px;color:#999;min-height:14px"></p>

        <div style="display:flex;gap:8px;justify-content:flex-end">
            <button id="_swapCancel" style="padding:7px 16px;border:1px solid #ccc;
                background:#f5f5f5;border-radius:6px;cursor:pointer;font-size:13px">Cancel</button>
            <button id="_swapConfirm" style="padding:7px 16px;border:none;background:#2196F3;
                color:#fff;border-radius:6px;cursor:pointer;font-size:13px" disabled>Swap</button>
        </div>`;

    overlay.appendChild(box);
    document.body.appendChild(overlay);

    const selFrom = document.getElementById("_swapFrom");
    const selTo = document.getElementById("_swapTo");
    const selTarget = document.getElementById("_swapTarget");
    const btnConfirm = document.getElementById("_swapConfirm");
    const noteEl = document.getElementById("_swapNote");

    selFrom.value = String(defaultFrom);
    selTo.value = String(defaultTo);

    function refreshCandidates() {
        const fi = parseInt(selFrom.value);
        const ti = parseInt(selTo.value);

        if (fi >= ti) {
            noteEl.textContent = "⚠ 'From' must be before 'To'";
            noteEl.style.color = "#c62828";
            selTarget.innerHTML = `<option disabled value="">— invalid range —</option>`;
            btnConfirm.disabled = true;
            return;
        }

        const candidates = getFreeCandidates(zoneName, fromCounter, fi, ti);
        if (candidates.length === 0) {
            noteEl.textContent = "No counters free for this window";
            noteEl.style.color = "#c62828";
            selTarget.innerHTML = `<option disabled value="">— none available —</option>`;
            btnConfirm.disabled = true;
        } else {
            noteEl.textContent = `${candidates.length} counter(s) available`;
            noteEl.style.color = "#2e7d32";

            // Group by zone using <optgroup>
            const prevVal = selTarget.value;
            const byZone = {};
            candidates.forEach(({ zone, counter }) => {
                if (!byZone[zone]) byZone[zone] = [];
                byZone[zone].push(counter);
            });

            selTarget.innerHTML = Object.entries(byZone).map(([zone, ctrs]) =>
                `<optgroup label="${zone}">${ctrs.map(c => `<option value="${zone}||${c}">${c}</option>`).join("")
                }</optgroup>`
            ).join("");

            // Restore previous selection if still valid
            if ([...selTarget.options].some(o => o.value === prevVal)) selTarget.value = prevVal;
            btnConfirm.disabled = false;
        }
    }

    selFrom.onchange = refreshCandidates;
    selTo.onchange = refreshCandidates;
    refreshCandidates();

    document.getElementById("_swapCancel").onclick = () => overlay.remove();
    btnConfirm.onclick = () => {
        const val = selTarget.value;
        const fi = parseInt(selFrom.value);
        const ti = parseInt(selTo.value);
        if (val && fi < ti) {
            const [toZone, toCounter] = val.split("||");
            swapCounters(zoneName, fromCounter, toZone, toCounter, fi, ti);
        }
        overlay.remove();
    };
    overlay.onclick = e => { if (e.target === overlay) overlay.remove(); };
}

/* ==================== End OOS & Counter Swap ==================== */

function attachTableEvents() {
    if (tableEventsAttached) return;

    const paintedThisDrag = new Set();

    table.addEventListener("pointerdown", e => {
        const cell = e.target.closest(".counter-cell");
        if (!cell) return;
        e.preventDefault();
        paintedThisDrag.clear();
        dragMode = cell.classList.contains("active") ? "remove" : "add";
        isDragging = true;
        table.setPointerCapture(e.pointerId);
        if (_saveStateGlobal) _saveStateGlobal(); // save before paint so drag is one undo step
        toggleCell(cell);
        paintedThisDrag.add(cell);
    });

    table.addEventListener("pointermove", e => {
        if (!isDragging) return;
        const el = document.elementFromPoint(e.clientX, e.clientY);
        const cell = el?.closest(".counter-cell");
        if (!cell || paintedThisDrag.has(cell)) return;
        const isActive = cell.classList.contains("active");
        if (dragMode === "add" && !isActive) { toggleCell(cell); paintedThisDrag.add(cell); }
        if (dragMode === "remove" && isActive) { toggleCell(cell); paintedThisDrag.add(cell); }
    });

    const endDrag = () => { isDragging = false; paintedThisDrag.clear(); };
    table.addEventListener("pointerup", endDrag);
    table.addEventListener("pointercancel", endDrag);
    document.addEventListener("pointerup", endDrag);

    tableEventsAttached = true;
    restoreCellStates();
}

/* ==================== Save / Restore Cell States ==================== */
function saveCellStates() {
    if (!document.querySelector(".counter-cell")) return; // grid not built yet
    const key = currentLane === "cargo"
        ? `${currentLane}_${currentMode}_${currentShift}_${currentCargoVariation}`
        : `${currentLane}_${currentMode}_${currentShift}`;
    cellStates[key] = {};
    document.querySelectorAll(".counter-cell").forEach(cell => {
        const id = `${cell.dataset.zone}_${cell.dataset.counter}_${cell.dataset.time}`;
        cellStates[key][id] = {
            active: cell.classList.contains("active"),
            color: cell.style.background,
            officer: cell.dataset.officer || "",
            type: cell.dataset.type || ""
        };
    });
}

function restoreCellStates() {
    const key = currentLane === "cargo"
        ? `${currentLane}_${currentMode}_${currentShift}_${currentCargoVariation}`
        : `${currentLane}_${currentMode}_${currentShift}`;
    const state = cellStates[key] || {};
    document.querySelectorAll(".counter-cell").forEach(cell => {
        const id = `${cell.dataset.zone}_${cell.dataset.counter}_${cell.dataset.time}`;
        if (state[id]) {
            // Saved state exists — restore it (may override Excel defaults)
            if (state[id].active) {
                cell.classList.add("active");
                // For cargo, don't restore background — renderCargoGrid sets it from CARGO_COLOURS
                if (currentLane !== "cargo") cell.style.background = state[id].color;
                cell.dataset.officer = state[id].officer || "";
                cell.dataset.type = state[id].type || "";
            } else {
                cell.classList.remove("active");
                cell.style.background = "";
                cell.dataset.officer = "";
                cell.dataset.type = "";
            }
        } else if (currentLane === "cargo") {
            // No saved state for cargo cell — keep Excel-derived state as-is
        } else {
            // Non-cargo, no saved state — clear
            cell.classList.remove("active");
            cell.style.background = "";
            cell.dataset.officer = "";
            cell.dataset.type = "";
        }
    });
    updateAll();
}

/* ---------------- Cell Toggle & Update ----------------- */
function toggleCell(cell) {
    if (dragMode === "add") {
        cell.style.background = currentColor;
        cell.classList.add("active");
    } else {
        cell.style.background = "";
        cell.classList.remove("active");
        cell.dataset.officer = "";
    }
    updateAll();
}

function updateAll() {
    updateSubtotals();
    updateGrandTotal();
    updateManningSummary();
    updateMainRoster();
    updateOTRosterTable();
    updateSOSRosterAll();
}

function updateSubtotals() {
    document.querySelectorAll(".subtotal-cell").forEach(td => {
        let zone = td.dataset.zone;
        let time = td.dataset.time;
        let cells = [...document.querySelectorAll(`.counter-cell[data-zone="${zone}"][data-time="${time}"]`)];
        let sum = cells.filter(c => c.classList.contains("active")).length;
        td.innerText = sum;
    });
}

function updateGrandTotal() { }

function updateManningSummary() {
    const times = generateTimeSlots();
    let text = "";

    times.forEach((time, i) => {
        let totalCars = 0;
        let zoneBreakdown = [];

        zones[currentLane + "_" + currentMode].forEach(zone => {
            if (zone.name === "BIKES") return;
            let cells = [...document.querySelectorAll(`.counter-cell[data-zone="${zone.name}"][data-time="${i}"]`)];
            let count = cells.filter(c => c.classList.contains("active")).length;
            totalCars += count;
            zoneBreakdown.push(count);
        });

        let bikeCells = [...document.querySelectorAll(`.counter-cell[data-zone="BIKES"][data-time="${i}"]`)];
        let bikeCount = bikeCells.filter(c => c.classList.contains("active")).length;

        text += `${time}: ${String(totalCars).padStart(2, "0")}/${String(bikeCount).padStart(2, "0")}\n${zoneBreakdown.join("/")}\n\n`;
    });

    manningSummary.textContent = text;
}

/* ---------------- Button Event Listeners ----------------- */
document.getElementById("copySummaryBtn").addEventListener("click", () => {
    navigator.clipboard.writeText(manningSummary.textContent).then(() => {
        let btn = document.getElementById("copySummaryBtn");
        btn.classList.add("copied");
        btn.innerText = "Copied ✓";
        setTimeout(() => {
            btn.classList.remove("copied");
            btn.innerText = "Copy Manning Summary";
        }, 2000);
    });
});

document.getElementById("copyMainRosterBtn")?.addEventListener("click", () => copyMainRoster());
document.getElementById("copySOSRosterBtn")?.addEventListener("click", () => copySOSRoster());
document.getElementById("copyOTRosterBtn")?.addEventListener("click", () => copyOTRoster());

document.getElementById("clearGridBtn").addEventListener("click", () => {
    document.querySelectorAll(".counter-cell").forEach(c => {
        c.style.background = "";
        c.classList.remove("active");
    });
    updateAll();
});

function updateMainRoster() {
    const tbody = document.querySelector("#mainRosterTable tbody");
    if (!tbody) return;

    tbody.innerHTML = "";
    const times = generateTimeSlots();
    const officerMap = {};

    document.querySelectorAll(".counter-cell.active").forEach(cell => {
        const type = cell.dataset.type || "";
        // Only show main officers and manually painted cells (no type)
        if (type && type !== "main") return;
        let officer = cell.dataset.officer;
        // Manually painted cells have no officer — label them by counter
        if (!officer) officer = `[${cell.dataset.counter}]`;
        const time = parseInt(cell.dataset.time);
        const zone = cell.dataset.zone;
        const counter = cell.dataset.counter;
        if (!officerMap[officer]) officerMap[officer] = [];
        officerMap[officer].push({ time, zone, counter });
    });

    Object.keys(officerMap).sort((a, b) => parseInt(a) - parseInt(b)).forEach(officer => {
        const records = officerMap[officer].sort((a, b) => a.time - b.time);
        if (!records.length) return;

        // RA/RO badges for display
        const raro = typeof raroRegistry !== "undefined" ? getRaro(officer) : null;
        const raBadge = raro?.ra
            ? `<span style="font-size:10px;font-weight:700;padding:1px 4px;border-radius:3px;margin-left:4px;
                background:#e3f2fd;color:#1565c0;border:1px solid #90caf9">RA ${formatTime(raro.ra)}</span>`
            : "";
        const roBadge = raro?.ro
            ? `<span style="font-size:10px;font-weight:700;padding:1px 4px;border-radius:3px;margin-left:4px;
                background:#fce4ec;color:#b71c1c;border:1px solid #f48fb1">RO ${formatTime(raro.ro)}</span>`
            : "";
        const raroBadge = raBadge + roBadge;

        let start = records[0].time;
        let prev = records[0].time;
        let currentZone = records[0].zone;
        let currentCounter = records[0].counter;
        let firstRow = true;

        for (let i = 1; i <= records.length; i++) {
            const isBreak =
                i === records.length ||
                records[i].time !== prev + 1 ||
                records[i].zone !== currentZone ||
                records[i].counter !== currentCounter;

            if (isBreak) {
                const counterLabel = currentLane === "cargo"
                    ? currentCounter
                    : `${currentZone} ${currentCounter}`;
                const row = document.createElement("tr");
                row.innerHTML = `
                    <td>${officer}${firstRow ? raroBadge : ""}</td>
                    <td>${counterLabel}</td>
                    <td>${formatTime(times[start])}</td>
                    <td>${formatTime(rosterEndTime(times, prev))}</td>
                `;
                tbody.appendChild(row);
                firstRow = false;

                if (i < records.length) {
                    start = records[i].time;
                    currentZone = records[i].zone;
                    currentCounter = records[i].counter;
                }
            }

            if (i < records.length) prev = records[i].time;
        }
    });
}

function formatTime(hhmm) {
    if (!hhmm) return "";
    return hhmm.slice(0, 2) + ":" + hhmm.slice(2);
}

function rosterEndTime(times, prev) {
    if (times[prev + 1]) return times[prev + 1];
    return currentShift === "night" ? "1000" : "2200";
}

/* ==================== OT Roster Table ==================== */
function updateOTRosterTable() {
    const tbody = document.querySelector("#otRosterTable tbody");
    if (!tbody) return;

    tbody.innerHTML = "";
    const times = generateTimeSlots();
    const officerMap = {};

    document.querySelectorAll('.counter-cell.active[data-type="ot"]').forEach(cell => {
        const officer = cell.dataset.officer;
        const time = parseInt(cell.dataset.time);
        const zone = cell.dataset.zone;
        const counter = cell.dataset.counter;
        if (!officer) return;
        if (!officerMap[officer]) officerMap[officer] = [];
        officerMap[officer].push({ time, zone, counter });
    });

    const officers = Object.keys(officerMap).sort((a, b) =>
        parseInt(a.replace("OT", "")) - parseInt(b.replace("OT", ""))
    );

    officers.forEach(officer => {
        const records = officerMap[officer].sort((a, b) => a.time - b.time);
        if (!records.length) return;

        let start = records[0].time;
        let prev = records[0].time;
        let currentZone = records[0].zone;
        let currentCounter = records[0].counter;

        for (let i = 1; i <= records.length; i++) {
            const isBreak =
                i === records.length ||
                records[i].time !== prev + 1 ||
                records[i].zone !== currentZone ||
                records[i].counter !== currentCounter;

            if (isBreak) {
                const row = document.createElement("tr");
                row.innerHTML = `
                    <td>${officer}</td>
                    <td>${currentZone} ${currentCounter}</td>
                    <td>${formatTime(times[start])}</td>
                    <td>${formatTime(rosterEndTime(times, prev))}</td>
                `;
                tbody.appendChild(row);

                // Insert Break row if gap exists to next record
                if (i < records.length && records[i].time > prev + 1) {
                    const breakRow = document.createElement("tr");
                    breakRow.classList.add("break-row");
                    breakRow.innerHTML = `
                        <td>${officer}</td>
                        <td>Break</td>
                        <td>${formatTime(times[prev + 1])}</td>
                        <td>${formatTime(times[records[i].time])}</td>
                    `;
                    tbody.appendChild(breakRow);
                }

                if (i < records.length) {
                    start = records[i].time;
                    currentZone = records[i].zone;
                    currentCounter = records[i].counter;
                }
            }

            if (i < records.length) prev = records[i].time;
        }
    });
}

function formatHHMM(time) {
    time = parseInt(time);
    const hh = String(Math.floor(time / 100)).padStart(2, "0");
    const mm = String(time % 100).padStart(2, "0");
    return `${hh}:${mm}`;
}

function updateSOSRoster(startTime, endTime) {
    const tbody = document.querySelector("#sosRosterTable tbody");
    if (!tbody) return;
    _renderSOSRoster(tbody, startTime, endTime);
}

function updateSOSRosterAll() {
    const tbody = document.querySelector("#sosRosterTable tbody");
    if (!tbody) return;
    // Derive window from whatever SOS cells exist on the grid
    const times = generateTimeSlots();
    let minT = Infinity, maxT = -Infinity;
    document.querySelectorAll('.counter-cell.active[data-type="sos"]').forEach(cell => {
        const t = parseInt(cell.dataset.time);
        if (t < minT) minT = t;
        if (t > maxT) maxT = t;
    });
    if (minT === Infinity) { tbody.innerHTML = ""; return; }
    _renderSOSRoster(tbody, times[minT], times[maxT + 1] || times[maxT]);
}

function _renderSOSRoster(tbody, startTime, endTime) {
    tbody.innerHTML = "";
    const times = generateTimeSlots();
    const startIndex = times.findIndex(t => t === startTime);
    const endIndex = times.findIndex(t => t === endTime);
    if (startIndex === -1 || endIndex === -1) return;

    const officerMap = {};
    document.querySelectorAll('.counter-cell.active[data-type="sos"]').forEach(cell => {
        const officer = cell.dataset.officer;
        const time = parseInt(cell.dataset.time);
        const zone = cell.dataset.zone;
        const counter = cell.dataset.counter;
        if (!officer) return;
        if (time < startIndex || time >= endIndex) return;
        if (!officerMap[officer]) officerMap[officer] = [];
        officerMap[officer].push({ time, zone, counter });
    });

    Object.keys(officerMap).sort((a, b) => parseInt(a) - parseInt(b)).forEach(officer => {
        const records = officerMap[officer].sort((a, b) => a.time - b.time);
        if (!records.length) return;

        let start = records[0].time;
        let prev = records[0].time;
        let currentZone = records[0].zone;
        let currentCounter = records[0].counter;

        for (let i = 1; i <= records.length; i++) {
            const isBreak =
                i === records.length ||
                records[i].time !== prev + 1 ||
                records[i].zone !== currentZone ||
                records[i].counter !== currentCounter;

            if (isBreak) {
                const row = document.createElement("tr");
                row.innerHTML = `
                    <td>${officer}</td>
                    <td>${currentZone} ${currentCounter}</td>
                    <td>${formatTime(times[start])}</td>
                    <td>${formatTime(rosterEndTime(times, prev))}</td>
                `;
                tbody.appendChild(row);
                if (i < records.length) {
                    start = records[i].time;
                    currentZone = records[i].zone;
                    currentCounter = records[i].counter;
                }
            }
            if (i < records.length) prev = records[i].time;
        }
    });
}

/* ---------------- Mode & Shift Segmented Buttons ----------------- */
const renderedTables = {};
["car", "bus", "train", "cargo", "owc"].forEach(lane =>
    ["arrival", "departure"].forEach(mode =>
        ["morning", "night"].forEach(shift => {
            renderedTables[`${lane}_${mode}_${shift}`] = false;
        })
    )
);

function setLane(lane) {
    resetDragState();
    saveCellStates();
    currentLane = lane;
    document.querySelectorAll(".lane-btn").forEach(b => {
        b.classList.toggle("active", b.dataset.lane === lane);
    });
    oosCounters.clear();
    renderTableOnce();
    if (lane === "cargo") renderCargoVariationTabs();
    restoreCellStates();
    attachCounterContextMenus();
    updateTrainOwcVisibility();
}

function setMode(mode) {
    resetDragState();
    saveCellStates();
    currentMode = mode;

    if (mode === "arrival") {
        currentColor = "#4CAF50";
        modeHighlight.style.transform = "translateX(0%)";
        modeHighlight.style.background = "#4CAF50";
        arrivalBtn.classList.add("active");
        departureBtn.classList.remove("active");
    } else {
        currentColor = "#FF9800";
        modeHighlight.style.transform = "translateX(100%)";
        modeHighlight.style.background = "#FF9800";
        departureBtn.classList.add("active");
        arrivalBtn.classList.remove("active");
    }

    isDragging = false;
    dragMode = "add";
    oosCounters.clear();
    renderTableOnce();
    if (currentLane === "cargo") renderCargoVariationTabs();
    restoreCellStates();
    attachCounterContextMenus();
    updateTrainOwcVisibility();
}

function updateTrainOwcVisibility() {
    const el = document.getElementById("trainOwcFields");
    const label = document.getElementById("trainOwcLabel");
    if (!el) return;
    const isMain = document.querySelector(".mp-type.active")?.dataset.type === "main";
    // Show only on Car lane + night shift + main type (original behaviour)
    const show = currentLane === "car" && currentShift === "night" && isMain;
    el.style.display = show ? "block" : "none";
    if (label) label.textContent = currentMode === "arrival" ? "Train Officers:" : "OWC Officers:";
}

function updateOTSlotDropdown() {
    const sel = document.getElementById("otSlot");
    if (!sel) return;
    if (currentShift === "morning") {
        sel.innerHTML = `
            <option value="1100-1600">1100 – 1600</option>
            <option value="1600-2100">1600 – 2100</option>`;
    } else {
        sel.innerHTML = `
            <option value="0600-1100">0600 – 1100</option>`;
    }
}

function setShift(shift) {
    resetDragState();
    saveCellStates();
    currentShift = shift;

    if (shift === "morning") {
        shiftHighlight.style.transform = "translateX(0%)";
        shiftHighlight.style.background = "#b0bec5";
        morningBtn.classList.add("active");
        nightBtn.classList.remove("active");
    } else {
        shiftHighlight.style.transform = "translateX(100%)";
        shiftHighlight.style.background = "#9e9e9e";
        nightBtn.classList.add("active");
        morningBtn.classList.remove("active");
    }

    if (currentMode === "arrival") {
        currentColor = "#4CAF50";
    } else {
        currentColor = "#FF9800";
    }

    isDragging = false;
    dragMode = "add";
    oosCounters.clear();
    renderTableOnce();
    if (currentLane === "cargo") renderCargoVariationTabs();
    restoreCellStates();
    attachCounterContextMenus();
    updateTrainOwcVisibility();
    updateOTSlotDropdown();
}

arrivalBtn.onclick = () => setMode("arrival");
departureBtn.onclick = () => setMode("departure");
morningBtn.onclick = () => setShift("morning");
nightBtn.onclick = () => setShift("night");

document.querySelectorAll(".lane-btn").forEach(btn => {
    btn.onclick = () => setLane(btn.dataset.lane);
});

function attachCounterContextMenus() {
    if (!document.querySelector("#rosterTable tr")) return; // grid not ready
    document.querySelectorAll("#rosterTable tr").forEach(row => {
        const firstCell = row.cells?.[0];
        const dataCell = row.querySelector(".counter-cell");
        if (!firstCell || !dataCell) return;
        const zoneName = dataCell.dataset.zone;
        const counter = dataCell.dataset.counter;
        if (!zoneName || !counter) return;

        firstCell.style.cursor = "context-menu";
        firstCell.oncontextmenu = (e) => {
            e.preventDefault();
            const menu = buildCounterContextMenu(zoneName, counter);
            menu.style.left = e.clientX + "px";
            menu.style.top = e.clientY + "px";
            // Close on any outside click
            setTimeout(() => {
                document.addEventListener("click", function close() {
                    menu.remove();
                    document.removeEventListener("click", close);
                });
            }, 0);
        };
    });
}


/* ---------------- Excel Template Loading ---------------- */
async function loadExcelTemplate() {
    try {
        const response = await fetch("ROSTER.xlsx");
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const arrayBuffer = await response.arrayBuffer();
        excelWorkbook = XLSX.read(arrayBuffer, { type: "array" });
        excelWorkbook.SheetNames.forEach(sheetName => {
            const sheet = excelWorkbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(sheet);
            excelData[sheetName.toLowerCase()] = json;
        });
        console.log("Excel template loaded:", Object.keys(excelData));
    } catch (err) {
        console.error("Excel loading failed:", err);
        alert("Failed to load Excel. Check filename, location, and local server.");
    }
}

async function loadCargoExcels() {
    const files = [
        { key: "arrival_morning", filename: "CARGO_ARRIVAL_MORNING.xlsx" },
        { key: "arrival_night", filename: "CARGO_ARRIVAL_NIGHT.xlsx" },
        { key: "departure_morning", filename: "CARGO_DEPARTURE_MORNING.xlsx" },
        { key: "departure_night", filename: "CARGO_DEPARTURE_NIGHT.xlsx" }
    ];
    cargoWorkbooks = {};
    cargoVariations = {};
    for (const { key, filename } of files) {
        try {
            const res = await fetch(filename);
            if (!res.ok) { console.warn(`Cargo Excel not found: ${filename}`); continue; }
            const buf = await res.arrayBuffer();
            const wb = XLSX.read(buf, { type: "array" });
            cargoWorkbooks[key] = wb;
            cargoVariations[key] = wb.SheetNames.filter(n => n !== "List" && !n.startsWith("_"));
            console.log(`Cargo ${key} variations:`, cargoVariations[key]);
        } catch (err) {
            console.warn(`Failed to load ${filename}:`, err);
        }
    }
    if (currentLane === "cargo") renderCargoVariationTabs();
}

// ── Parse a cargo sheet into grid data ────────────────────────────────────────
// Returns { counters: [{name, cells: ["DL1"|"#"|""]}, ...], times: ["HHMM",...] }
function parseCargoSheet(sheetName) {
    const key = `${currentMode}_${currentShift}`;
    const wb = cargoWorkbooks[key];
    if (!wb) return null;
    const ws = wb.Sheets[sheetName];
    if (!ws) return null;

    // Structure: Row 1 = shift label (col A) + time slots from col B onwards
    //            Row 2 = "Lorry/Car-go" header (skip)
    //            Rows 3+ = counter rows (col A = name, col B+ = cells)
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
    if (raw.length < 3) return null;

    const timeRow = raw[0] || [];   // Row 1 (index 0)
    const timeColStart = 1;         // col B = index 1

    // Helper: convert various time formats to HHMM
    function toHHMM(v) {
        if (!v && v !== 0) return null;
        // Raw numeric Excel time (fraction of day)
        if (typeof v === 'number') {
            const totalMins = Math.round(v * 24 * 60);
            const hh = Math.floor(totalMins / 60) % 24;
            const mm = totalMins % 60;
            return String(hh).padStart(2, "0") + String(mm).padStart(2, "0");
        }
        const s = String(v).trim();
        // "HH:MM:SS" or "H:MM:SS"
        if (s.includes(":")) {
            const parts = s.split(":");
            return parts[0].padStart(2, "0") + parts[1].padStart(2, "0");
        }
        return null;
    }

    const times = [];
    for (let c = timeColStart; c < timeRow.length; c++) {
        const hhmm = toHHMM(timeRow[c]);
        if (!hhmm) break;
        times.push(hhmm);
    }
    if (!times.length) return null;

    const counters = [];
    const skipNames = new Set([
        "Big", "Small", "Sub-total", "Grand-total",
        "Cover Breaks", "Lorry/Car-go", "Morning", "Night", ""
    ]);
    // Names that belong to the "Cover Breaks" section
    const coverNames = new Set(["Checker", "Checker 1", "Checker 2", "DLSC"]);
    let inCoverSection = false;

    // Rows start at index 2 (row 3) — skip row 2 (Lorry/Car-go header)
    for (let r = 2; r < raw.length; r++) {
        const row = raw[r];
        if (!row || !row[0]) continue;
        const name = String(row[0]).trim();
        // Detect entry into cover break section
        if (name === "Cover Breaks") { inCoverSection = true; continue; }
        if (skipNames.has(name)) continue;

        const isChecker = inCoverSection || coverNames.has(name);
        const cells = [];
        for (let c = timeColStart; c < timeColStart + times.length; c++) {
            const v = row[c];
            // With raw:true, values are native types — convert to string, treat numbers as empty
            const s = (v !== undefined && v !== null && typeof v === 'string') ? v.trim() : "";
            if (s === "#") cells.push("#");
            else if (s !== "") cells.push(s);
            else cells.push("");
        }
        counters.push({ name, cells, isChecker });
    }

    return { times, counters };
}

// ── Render variation tabs above grid when on Cargo lane ────────────────────────
function renderCargoVariationTabs() {
    let tabBar = document.getElementById("cargoVariationTabs");
    if (!tabBar) return;

    const key = `${currentMode}_${currentShift}`;
    const variations = cargoVariations[key] || [];
    if (!variations.length) {
        tabBar.style.display = "none";
        return;
    }
    tabBar.style.display = "flex";
    tabBar.innerHTML = "";

    // If current variation isn't valid for this mode, reset
    if (!variations.includes(currentCargoVariation)) {
        currentCargoVariation = variations[0];
    }

    variations.forEach(name => {
        const btn = document.createElement("button");
        btn.className = "cargo-var-btn" + (name === currentCargoVariation ? " active" : "");
        btn.textContent = name;
        btn.onclick = () => {
            saveCellStates();
            currentCargoVariation = name;
            document.querySelectorAll(".cargo-var-btn").forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            renderCargoGrid();
        };
        tabBar.appendChild(btn);
    });

    renderCargoGrid();
}

// ── Colour map matching Excel conditional formatting ──────────────────────────
const CARGO_COLOURS = {
    "DL1": { bg: "#BBDEFB", text: "#000000" }, "AL1": { bg: "#BBDEFB", text: "#000000" },
    "D-Cargo 1": { bg: "#BBDEFB", text: "#000000" }, "A-Cargo 1": { bg: "#BBDEFB", text: "#000000" },
    "DL2": { bg: "#FFE0B2", text: "#000000" }, "AL2": { bg: "#FFE0B2", text: "#000000" },
    "D-Cargo 2": { bg: "#FFE0B2", text: "#000000" }, "A-Cargo 2": { bg: "#FFE0B2", text: "#000000" },
    "DL3": { bg: "#C8E6C9", text: "#000000" }, "AL3": { bg: "#C8E6C9", text: "#000000" },
    "D-Cargo 3": { bg: "#C8E6C9", text: "#000000" }, "A-Cargo 3": { bg: "#C8E6C9", text: "#000000" },
    "DL4": { bg: "#F8BBD0", text: "#000000" }, "AL4": { bg: "#F8BBD0", text: "#000000" },
    "D-Cargo 4": { bg: "#F8BBD0", text: "#000000" }, "A-Cargo 4": { bg: "#F8BBD0", text: "#000000" },
    "DL5": { bg: "#FFF9C4", text: "#000000" }, "AL5": { bg: "#FFF9C4", text: "#000000" },
    "D-Cargo 5": { bg: "#FFF9C4", text: "#000000" }, "A-Cargo 5": { bg: "#FFF9C4", text: "#000000" },
    "DL6": { bg: "#EEEEEE", text: "#000000" }, "AL6": { bg: "#EEEEEE", text: "#000000" },
    "D-Cargo 6": { bg: "#EEEEEE", text: "#000000" }, "A-Cargo 6": { bg: "#EEEEEE", text: "#000000" },
};

// ── Render the cargo grid from selected variation ──────────────────────────────
function renderCargoGrid() {
    if (currentLane !== "cargo" || !currentCargoVariation) return;
    const data = parseCargoSheet(currentCargoVariation);
    if (!data) { table.innerHTML = "<tr><td>No data for this variation.</td></tr>"; return; }

    table.innerHTML = "";

    // Single zone for cargo
    const zoneName = currentMode === "arrival" ? "Cargo" : "Cargo";

    // Zone header
    const zoneRow = document.createElement("tr");
    const zoneCell = document.createElement("td");
    zoneCell.colSpan = data.times.length + 1;
    zoneCell.className = "zone-header";
    zoneCell.innerText = `Cargo — ${currentCargoVariation}`;
    zoneRow.appendChild(zoneCell);
    table.appendChild(zoneRow);

    // Time header
    const timeRow = document.createElement("tr");
    timeRow.className = "time-header";
    timeRow.innerHTML = "<th></th>";
    data.times.forEach(t => {
        const th = document.createElement("th");
        th.innerText = t;
        timeRow.appendChild(th);
    });
    table.appendChild(timeRow);

    // Counter rows
    let separatorAdded = false;
    data.counters.forEach(({ name, cells, isChecker }) => {
        // Add separator row before first checker/cover-break row
        if (isChecker && !separatorAdded) {
            separatorAdded = true;
            const sepRow = document.createElement("tr");
            const sepCell = document.createElement("td");
            sepCell.colSpan = data.times.length + 1;
            sepCell.className = "cargo-separator";
            sepCell.innerText = "Cover Breaks";
            sepRow.appendChild(sepCell);
            table.appendChild(sepRow);
        }

        const row = document.createElement("tr");
        const label = document.createElement("td");
        label.innerText = name;
        row.appendChild(label);

        cells.forEach((cellVal, i) => {
            const td = document.createElement("td");
            td.dataset.zone = zoneName;
            td.dataset.counter = name;
            td.dataset.time = i;

            if (cellVal !== "#" && cellVal !== "") {
                td.className = "counter-cell active";
                td.dataset.officer = isChecker ? `${name} (${cellVal})` : name;
                td.dataset.type = "main";
                const colKey = isChecker ? cellVal : name;
                // Normalise checker key: if plain number, prefix with AL/DL based on mode
                const prefix = currentMode === "arrival" ? "AL" : "DL";
                const normKey = (isChecker && /^\d+$/.test(colKey)) ? `${prefix}${colKey}` : colKey;
                const colour = CARGO_COLOURS[normKey];
                if (colour) {
                    td.style.background = colour.bg;
                    td.style.color = colour.text;
                } else {
                    td.style.background = currentColor;
                }
            } else {
                td.className = "counter-cell";
            }
            row.appendChild(td);
        });

        table.appendChild(row);
    });

    // Subtotal row
    const subRow = document.createElement("tr");
    subRow.className = "subtotal-row";
    const subLabel = document.createElement("td");
    subLabel.innerText = "Subtotal";
    subRow.appendChild(subLabel);
    data.times.forEach((_, i) => {
        const td = document.createElement("td");
        td.className = "subtotal-cell";
        td.dataset.zone = zoneName;
        td.dataset.time = i;
        subRow.appendChild(td);
    });
    table.appendChild(subRow);

    restoreCellStates();
    attachCounterContextMenus();
    updateAll();
}

// Override restoreCellStates for cargo: if no saved state exists for a cell,
// keep the Excel-derived active state already set by renderCargoGrid.
// The base restoreCellStates handles this correctly because it only overwrites
// cells that have a saved entry — cells without a saved entry are left alone.

/* ================= MANPOWER SYSTEM ================= */
document.addEventListener("DOMContentLoaded", function () {
    /* ---------------- INIT ---------------- */
    setLane("car");
    setMode("arrival");
    setShift("morning");

    attachTableEvents();
    attachCounterContextMenus();
    loadExcelTemplate();
    loadCargoExcels();

    let manpowerType = "main";

    const sosFields = document.getElementById("sosFields");
    const otFields = document.getElementById("otFields");
    const raFields = document.getElementById("raFields");
    const roFields = document.getElementById("roFields");
    const addBtn = document.getElementById("addOfficerBtn");
    const removeBtn = document.getElementById("removeOfficerBtn");
    const undoBtn = document.getElementById("undoBtn");

    if (!addBtn || !removeBtn || !undoBtn) {
        console.error("Manpower buttons not found in HTML.");
        return;
    }

    // Populate officer dropdowns for RA/RO with main officers currently on grid
    function populateRARODropdowns() {
        const officers = [...new Set(
            [...document.querySelectorAll('.counter-cell.active[data-type="main"]')]
                .map(c => c.dataset.officer).filter(Boolean)
        )].sort((a, b) => (parseInt(a) || 0) - (parseInt(b) || 0));

        ["raOfficer", "roOfficer"].forEach(id => {
            const sel = document.getElementById(id);
            if (!sel) return;
            const prev = sel.value;
            sel.innerHTML = officers.map(o => `<option value="${o}">${o}</option>`).join("");
            if (officers.includes(prev)) sel.value = prev;
        });
    }

    /* -------------------- Select Manpower Type -------------------- */
    const raroRow = document.getElementById("raroRow");
    const confirmRaRoBtn = document.getElementById("confirmRaRoBtn");

    function updateMpUI() {
        const isRaRo = manpowerType === "ra" || manpowerType === "ro";
        const isMain = manpowerType === "main";

        sosFields.style.display = manpowerType === "sos" ? "block" : "none";
        otFields.style.display = manpowerType === "ot" ? "block" : "none";
        raFields.style.display = manpowerType === "ra" ? "block" : "none";
        roFields.style.display = manpowerType === "ro" ? "block" : "none";

        // RA/RO sub-row always visible when Main or RA/RO is active
        if (raroRow) raroRow.style.display = (isMain || isRaRo) ? "flex" : "none";

        // Hide count input for RA/RO
        document.getElementById("officerCount").style.display = isRaRo ? "none" : "";
        const countLabel = document.querySelector("label[for='officerCount']");
        if (countLabel) countLabel.style.display = isRaRo ? "none" : "";

        // Swap Add Officers ↔ Confirm button
        addBtn.style.display = isRaRo ? "none" : "";
        if (confirmRaRoBtn) confirmRaRoBtn.style.display = isRaRo ? "" : "none";

        if (isRaRo) populateRARODropdowns();
        updateTrainOwcVisibility();
    }

    document.querySelectorAll(".mp-type").forEach(btn => {
        btn.addEventListener("click", () => {
            document.querySelectorAll(".mp-type").forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            manpowerType = btn.dataset.type;
            updateMpUI();
        });
    });

    // Initialise UI state on page load
    updateMpUI();
    updateOTSlotDropdown();

    /* -------------------- Officer name suffix -------------------- */
    function sosSuffix() {
        const name = (document.getElementById("sosOfficerName")?.value || "").trim();
        return name ? " | " + name : "";
    }
    function otSuffix() {
        const name = (document.getElementById("otOfficerName")?.value || "").trim();
        return name ? " | " + name : "";
    }
    // Main / Train / OWC have no name suffix (identified by serial number)
    function officerSuffix() { return ""; }

    /* -------------------- Save / Restore State -------------------- */
    function saveState() {
        const state = [];
        document.querySelectorAll(".counter-cell").forEach(cell => {
            state.push({
                zone: cell.dataset.zone,
                counter: cell.dataset.counter,
                time: cell.dataset.time,
                active: cell.classList.contains("active"),
                color: cell.style.background,
                officer: cell.dataset.officer || "",
                type: cell.dataset.type || ""
            });
        });
        historyStack.push(state);
        if (historyStack.length > 50) historyStack.shift(); // cap at 50 steps
    }
    _saveStateGlobal = saveState; // expose to attachTableEvents

    function restoreState(state) {
        document.querySelectorAll(".counter-cell").forEach(cell => {
            const found = state.find(s =>
                s.zone === cell.dataset.zone &&
                s.counter === cell.dataset.counter &&
                s.time === cell.dataset.time
            );
            if (found && found.active) {
                cell.classList.add("active");
                cell.style.background = found.color;
                cell.dataset.officer = found.officer;
                cell.dataset.type = found.type;
            } else {
                cell.classList.remove("active");
                cell.style.background = "";
                cell.dataset.officer = "";
                cell.dataset.type = "";
            }
        });
        updateAll();
    }

    /* -------------------- Main Template Assignment -------------------- */
    function applyMainTemplate(officerCount) {
        if (!excelWorkbook) {
            alert("Excel template not loaded.");
            return;
        }

        const sheetName = (currentLane === "car" ? `${currentMode} ${currentShift}` : `${currentLane} ${currentMode} ${currentShift}`).toLowerCase();
        const sheetData = excelData[sheetName];

        if (!sheetData) {
            alert("No sheet found for " + sheetName);
            return;
        }

        const times = generateTimeSlots();

        for (let officer = 1; officer <= officerCount; officer++) {
            const officerRows = sheetData.filter(row => parseInt(row.Officer) === officer);
            const officerLabel = officer + officerSuffix();

            officerRows.forEach(row => {
                const counter = row.Counter;

                function normalizeExcelTime(value) {
                    if (!value) return "";
                    let str = value.toString().trim();
                    if (str.includes(":")) {
                        str = str.substring(0, 5);
                        return str.replace(":", "");
                    }
                    return str.padStart(4, "0");
                }

                const start = normalizeExcelTime(row.Start);
                const end = normalizeExcelTime(row.End);

                let startIndex = times.findIndex(t => t === start);
                let endIndex = times.findIndex(t => t === end);

                if (endIndex === -1) {
                    if ((currentShift === "morning" && end === "2200") ||
                        (currentShift === "night" && end === "1000")) {
                        endIndex = times.length;
                    }
                }

                if (startIndex === -1 || endIndex === -1) return;

                for (let t = startIndex; t < endIndex; t++) {
                    let allCells = [...document.querySelectorAll(`.counter-cell[data-time="${t}"]`)];
                    allCells.forEach(cell => {
                        const rowCounter = cell.parentElement.firstChild.innerText;
                        if (rowCounter === counter) {
                            cell.classList.add("active");
                            cell.style.background = currentColor;
                            cell.dataset.officer = officerLabel;
                            cell.dataset.type = "main";
                        }
                    });
                }
            });
        }

        // Every 4th Officer Special Period (2030 to end of shift)
        if ((currentMode === "arrival" || currentMode === "departure") && currentShift === "morning") {
            const specialStart = "2030";
            const startIndex = times.findIndex(t => t === specialStart);
            const endIndex = times.length; // exclusive — run to end of grid

            if (startIndex !== -1) {
                // Officers 4, 8, 12... that fall within officerCount get the special period
                for (let officer = 4; officer <= officerCount; officer += 4) {
                    const specialLabel = officer + officerSuffix(); // correct label per officer

                    let assigned = false;
                    const candidateZones = zones[currentLane + "_" + currentMode].filter(z => z.name !== "BIKES");
                    const zoneOccupancy = candidateZones.map(zone => {
                        let occupiedCount = 0;
                        for (let t = startIndex; t < endIndex; t++) {
                            occupiedCount += [...document.querySelectorAll(
                                `.counter-cell[data-zone="${zone.name}"][data-time="${t}"]`
                            )].filter(c => c.classList.contains("active")).length;
                        }
                        const totalSlots = zone.counters.length * (endIndex - startIndex);
                        return { zone, occupiedCount, totalSlots, ratio: occupiedCount / totalSlots };
                    });
                    zoneOccupancy.sort((a, b) => a.ratio - b.ratio);

                    for (let z = 0; z < zoneOccupancy.length && !assigned; z++) {
                        const zone = zoneOccupancy[z].zone;
                        const counters = [...zone.counters].reverse(); // back counters first

                        for (let c = 0; c < counters.length && !assigned; c++) {
                            const counter = counters[c];
                            let isFree = true;

                            for (let t = startIndex; t < endIndex; t++) {
                                const cell = document.querySelector(
                                    `.counter-cell[data-zone="${zone.name}"][data-time="${t}"][data-counter="${counter}"]`
                                );
                                if (!cell || cell.classList.contains("active")) { isFree = false; break; }
                            }

                            if (isFree) {
                                for (let t = startIndex; t < endIndex; t++) {
                                    const cell = document.querySelector(
                                        `.counter-cell[data-zone="${zone.name}"][data-time="${t}"][data-counter="${counter}"]`
                                    );
                                    if (cell) {
                                        cell.classList.add("active");
                                        cell.style.background = currentColor;
                                        cell.dataset.officer = specialLabel;
                                        cell.dataset.type = "main";
                                    }
                                }
                                assigned = true;
                            }
                        }
                    }
                }
            }
        }
        updateAll();
    }

    /* -------------------- Train / OWC Template Assignment -------------------- */
    // Arrival night → TRAIN officers; Departure night → OWC officers
    // label prefix comes from the Excel Officer column (e.g. "TRAIN 1", "OWC 2")
    const TRAIN_OWC_COLOR = "#87CEEB"; // sky blue — distinct from main orange

    function applyTrainOwcTemplate(count) {
        if (!excelWorkbook) { alert("Excel template not loaded."); return; }
        if (currentShift !== "night") { alert("Train/OWC officers only apply to night shift."); return; }

        const sheetName = (currentLane === "car" ? `${currentMode} ${currentShift}` : `${currentLane} ${currentMode} ${currentShift}`).toLowerCase();
        const sheetData = excelData[sheetName];
        if (!sheetData) { alert("No sheet found for " + sheetName); return; }

        // Determine prefix based on mode
        const prefix = currentMode === "arrival" ? "TRAIN" : "OWC";

        // Collect all matching officer labels from the sheet, in order
        const allLabels = [];
        sheetData.forEach(row => {
            const o = String(row.Officer || "").trim().toUpperCase();
            if (o.startsWith(prefix) && !allLabels.includes(row.Officer)) {
                allLabels.push(row.Officer);
            }
        });

        // Skip officers already deployed on the grid
        const alreadyOnGrid = new Set(
            [...document.querySelectorAll(".counter-cell.active")]
                .map(c => c.dataset.officer)
                .filter(o => o && o.toUpperCase().startsWith(prefix))
                .map(o => o.split(" | ")[0]) // strip name/serial suffix for matching
        );

        const available = allLabels.filter(l => !alreadyOnGrid.has(l));

        if (available.length === 0) {
            alert(`All ${prefix} officers from the template are already on the grid.`);
            return;
        }

        // Deploy up to `count` of the remaining officers
        const toDeploy = available.slice(0, count);
        if (toDeploy.length < count) {
            alert(`Only ${toDeploy.length} ${prefix} officer(s) remaining in template (${allLabels.length} total). Deploying ${toDeploy.length}.`);
        }

        const times = generateTimeSlots();

        function normalizeExcelTime(value) {
            if (!value) return "";
            let str = value.toString().trim();
            if (str.includes(":")) return str.substring(0, 5).replace(":", "");
            return str.padStart(4, "0");
        }

        toDeploy.forEach(officerLabel => {
            const officerRows = sheetData.filter(row => row.Officer === officerLabel);
            const label = officerLabel + officerSuffix();

            officerRows.forEach(row => {
                const counter = row.Counter;
                const start = normalizeExcelTime(row.Start);
                const end = normalizeExcelTime(row.End);

                let startIndex = times.findIndex(t => t === start);
                let endIndex = times.findIndex(t => t === end);

                if (endIndex === -1 && end === "1000") endIndex = times.length;
                if (startIndex === -1 || endIndex === -1) return;

                for (let t = startIndex; t < endIndex; t++) {
                    [...document.querySelectorAll(`.counter-cell[data-time="${t}"]`)].forEach(cell => {
                        if (cell.parentElement.firstChild.innerText === counter) {
                            cell.classList.add("active");
                            cell.style.background = TRAIN_OWC_COLOR;
                            cell.dataset.officer = label;
                            cell.dataset.type = "main"; // same roster table as main
                        }
                    });
                }
            });
        });

        updateAll();
    }
    function isOTWithinShift(otStart, otEnd) {
        if (currentShift === "morning") {
            return (otStart === "1100" && otEnd === "1600") ||
                (otStart === "1600" && otEnd === "2100");
        } else {
            return (otStart === "0600" && otEnd === "1100");
        }
    }

    /* ================== OT ALLOCATION ==================
     *
     * Chain-handoff groups of 3 [A(+90), B(+135), C(+180)]:
     *   A: X(block1) → Y(block2)    Y = CONTINUOUS (B fills then A)
     *   B: Y(block1) → Z(block2)    Z = CONTINUOUS (C fills then B)
     *   C: Z(block1) → X(block2)    X = gap counter (front, acceptable)
     *
     * Each counter is only checked against its OWN required time windows:
     *   Y must be free: [sIdx→BK1] and [BKE0→end]
     *   Z must be free: [sIdx→BK2] and [BKE1→end]
     *   X must be free: [sIdx→BK0] and [BKE2→end]
     * This lets OT use counters that main officers vacate mid-shift.
     *
     * Selection per group (re-evaluated live after every fill):
     *   Y = highest free back counter in least-manned zone  (50% rule)
     *   Z = highest free back counter overall for Z-windows
     *   X = lowest free front counter overall for X-windows (gap falls here)
     * ================================================== */

    function allocateOTOfficers(count, otStart, otEnd) {
        const times = generateTimeSlots();

        const shiftEndTime = currentShift === "morning" ? "2200" : "1000";
        const otBoundaryEnd = currentShift === "morning" ? "2200" : "1100";

        let startIndex = times.findIndex(t => t === otStart);
        let endIndex = (otEnd === shiftEndTime || otEnd === otBoundaryEnd)
            ? times.length
            : times.findIndex(t => t === otEnd);
        if (startIndex === -1) { alert("OT start time outside current shift."); return; }
        if (endIndex === -1) { alert("OT end time outside current shift."); return; }

        const effectiveEnd = currentShift === "night" ? endIndex : Math.max(startIndex, endIndex - 2);

        // Fixed OT patterns keyed by otStart (HHMM)
        // w2s = null means no second work block
        const OT_PATTERNS = {
            "0600": [
                { w1e: "0730", w2s: "0815" }, // A
                { w1e: "0815", w2s: "0900" }, // B
                { w1e: "0900", w2s: "0945" }  // C — 0945 is last slot, handover to next team
            ],
            "1100": [
                { w1e: "1230", w2s: "1315" }, // A
                { w1e: "1315", w2s: "1400" }, // B
                { w1e: "1400", w2s: "1445" }  // C
            ],
            "1600": [
                { w1e: "1730", w2s: "1815" }, // A
                { w1e: "1815", w2s: "1900" }, // B
                { w1e: "1900", w2s: "1945" }  // C
            ]
        };

        const patterns = OT_PATTERNS[otStart];
        if (!patterns) { alert("No OT pattern defined for start time " + otStart); return; }

        function ti(hhmm) {
            if (!hhmm) return null;
            const idx = times.findIndex(t => t === hhmm);
            return idx === -1 ? effectiveEnd : idx;
        }
        const slotPatterns = patterns.map(p => ({
            w1e: ti(p.w1e),
            w2s: p.w2s !== null ? ti(p.w2s) : null
        }));

        // DOM helpers
        function getCell(zone, counter, t) {
            return document.querySelector(
                `.counter-cell[data-zone="${zone}"][data-time="${t}"][data-counter="${counter}"]`
            );
        }
        function blockFree(zone, counter, from, to) {
            for (let t = from; t < to; t++) {
                const c = getCell(zone, counter, t);
                if (!c || c.classList.contains("active")) return false;
            }
            return true;
        }
        function fillBlock(zone, counter, from, to, label) {
            for (let t = from; t < to; t++) {
                const c = getCell(zone, counter, t);
                if (!c || c.classList.contains("active")) continue;
                c.classList.add("active");
                c.style.background = currentColor;
                c.dataset.officer = label;
                c.dataset.type = "ot";
            }
        }
        function getFreeWindows(zone, counter) {
            const wins = []; let ws = null;
            for (let t = startIndex; t < effectiveEnd; t++) {
                const c = getCell(zone, counter, t);
                const free = c && !c.classList.contains("active");
                if (free && ws === null) ws = t;
                if (!free && ws !== null) { wins.push({ from: ws, to: t }); ws = null; }
            }
            if (ws !== null) wins.push({ from: ws, to: effectiveEnd });
            return wins;
        }

        const nonBikeZones = zones[currentLane + "_" + currentMode].filter(z => z.name !== "BIKES");

        // Build fill queue per zone — fully-free counters first, partials after
        function buildFillQueue(zoneName) {
            const z = zones[currentLane + "_" + currentMode].find(z => z.name === zoneName);
            if (!z) return { fullyFree: [], partials: [] };
            const fullyFree = [], partials = [];
            z.counters.forEach(counter => {
                if (oosCounters.has(oosKey(zoneName, counter))) return;
                let anyFree = false;
                for (let t = startIndex; t < effectiveEnd; t++) {
                    const c = getCell(zoneName, counter, t);
                    if (c && !c.classList.contains("active")) { anyFree = true; break; }
                }
                if (!anyFree) return;
                if (blockFree(zoneName, counter, startIndex, effectiveEnd)) {
                    fullyFree.push(counter);
                } else {
                    getFreeWindows(zoneName, counter).forEach(w => {
                        if (w.to - w.from >= 2) partials.push({ counter, from: w.from, to: w.to });
                    });
                }
            });
            fullyFree.sort((a, b) =>
                (parseInt(b.replace(/\D/g, "")) || 0) - (parseInt(a.replace(/\D/g, "")) || 0)
            );
            partials.sort((a, b) => (b.to - b.from) - (a.to - a.from));
            return { fullyFree, partials };
        }

        // Compute zone quotas
        const n = nonBikeZones.length;
        const base = Math.floor(count / n);
        const rem = count % n;
        const queueMap = {};
        nonBikeZones.forEach(z => { queueMap[z.name] = buildFillQueue(z.name); });

        const zoneQuota = {};
        nonBikeZones.forEach(z => {
            const total = queueMap[z.name].fullyFree.length + queueMap[z.name].partials.length;
            zoneQuota[z.name] = Math.min(base, total);
        });
        const remOrder = [...nonBikeZones].sort((a, b) =>
            (queueMap[b.name].fullyFree.length + queueMap[b.name].partials.length) -
            (queueMap[a.name].fullyFree.length + queueMap[a.name].partials.length) ||
            a.name.localeCompare(b.name)
        );
        let remLeft = rem;
        for (const z of remOrder) {
            if (remLeft <= 0) break;
            const total = queueMap[z.name].fullyFree.length + queueMap[z.name].partials.length;
            if (zoneQuota[z.name] < total) { zoneQuota[z.name]++; remLeft--; }
        }
        let deficit = count - Object.values(zoneQuota).reduce((s, v) => s + v, 0);
        if (deficit > 0) {
            for (const z of remOrder) {
                if (deficit <= 0) break;
                const total = queueMap[z.name].fullyFree.length + queueMap[z.name].partials.length;
                const extra = Math.min(deficit, total - zoneQuota[z.name]);
                if (extra > 0) { zoneQuota[z.name] += extra; deficit -= extra; }
            }
        }

        // Assign OT officers — rotate A→B→C per zone
        const patIdx = {};
        nonBikeZones.forEach(z => { patIdx[z.name] = 0; });

        nonBikeZones.forEach(z => {
            const { fullyFree, partials } = queueMap[z.name];
            const quota = zoneQuota[z.name];
            const countersToUse = [
                ...fullyFree,
                ...partials.map(p => p.counter)
            ].slice(0, quota);

            countersToUse.forEach(counter => {
                const label = "OT" + (otGlobalCounter++) + otSuffix();
                const pat = slotPatterns[patIdx[z.name] % 3];
                patIdx[z.name]++;

                fillBlock(z.name, counter, startIndex, pat.w1e, label);
                if (pat.w2s !== null) {
                    fillBlock(z.name, counter, pat.w2s, effectiveEnd, label);
                }
            });
        });

        updateAll();
        updateOTRosterTable();
    }


    /* ================== SOS ALLOCATION ================== */
    function allocateSOSOfficers(count, sosStart, sosEnd) {
        const times = generateTimeSlots();
        const shiftEndTime = currentShift === "morning" ? "2200" : "1000";

        let startIndex = times.findIndex(t => t === sosStart);
        let endIndex = sosEnd === shiftEndTime ? times.length : times.findIndex(t => t === sosEnd);

        if (startIndex === -1 || endIndex === -1) {
            alert("Invalid SOS timing.");
            return;
        }

        const threeHourSlots = 180 / 15;
        const breakSlots = 45 / 15;

        for (let i = 0; i < count; i++) {
            const officerLabel = "SOS" + (sosGlobalCounter++) + sosSuffix();
            let currentStart = startIndex;

            while (currentStart < endIndex) {
                let workEnd = Math.min(currentStart + threeHourSlots, endIndex);
                deploySOSBlock(officerLabel, currentStart, workEnd);
                currentStart = workEnd;
                if (currentStart < endIndex) {
                    currentStart = Math.min(currentStart + breakSlots, endIndex);
                }
            }
        }

        updateAll();
        updateSOSRoster(sosStart, sosEnd);
    }

    function deploySOSBlock(officerLabel, blockStart, blockEnd) {
        let assigned = false;
        const zoneStats = [];

        zones[currentLane + "_" + currentMode].forEach(zone => {
            if (zone.name === "BIKES") return;
            let activeCount = 0;
            zone.counters.forEach(counter => {
                for (let t = blockStart; t < blockEnd; t++) {
                    const cell = document.querySelector(
                        `.counter-cell[data-zone="${zone.name}"][data-time="${t}"][data-counter="${counter}"]`
                    );
                    if (cell && cell.classList.contains("active")) {
                        activeCount++;
                        break;
                    }
                }
            });
            zoneStats.push({ zone: zone.name, ratio: activeCount / zone.counters.length });
        });

        zoneStats.sort((a, b) => a.ratio - b.ratio);

        for (let z = 0; z < zoneStats.length; z++) {
            const zoneName = zoneStats[z].zone;
            const zone = zones[currentLane + "_" + currentMode].find(z => z.name === zoneName);

            for (let c = zone.counters.length - 1; c >= 0; c--) {
                const counter = zone.counters[c];
                let blockFree = true;

                for (let t = blockStart; t < blockEnd; t++) {
                    const cell = document.querySelector(
                        `.counter-cell[data-zone="${zoneName}"][data-time="${t}"][data-counter="${counter}"]`
                    );
                    if (!cell || cell.classList.contains("active")) {
                        blockFree = false;
                        break;
                    }
                }

                if (blockFree) {
                    // Skip OOS counters
                    if (oosCounters.has(oosKey(zoneName, counter))) {
                        blockFree = false;
                    }
                }
                if (blockFree) {
                    for (let t = blockStart; t < blockEnd; t++) {
                        const cell = document.querySelector(
                            `.counter-cell[data-zone="${zoneName}"][data-time="${t}"][data-counter="${counter}"]`
                        );
                        cell.classList.add("active");
                        cell.style.background = currentColor;
                        cell.dataset.officer = officerLabel;
                        cell.dataset.type = "sos";
                    }
                    assigned = true;
                    break;
                }
            }
            if (assigned) break;
        }
    }

    function isCounterFreeForBlock(zone, counter, start, end) {
        for (let t = start; t < end; t++) {
            const cell = document.querySelector(
                `.counter-cell[data-zone="${zone}"][data-time="${t}"][data-counter="${counter}"]`
            );
            if (!cell || cell.classList.contains("active")) return false;
        }
        return true;
    }

    /* -------------------- Button Clicks -------------------- */
    /* ── Capacity helpers ───────────────────────────────────────────────── */

    // Returns { used, max, remaining, label } for the current context.
    // `windowStart` / `windowEnd` are HHMM strings used for OT/SOS checks.
    function getCapacity(type, windowStart, windowEnd) {
        const times = generateTimeSlots();
        const allZones = zones[currentLane + "_" + currentMode].filter(z => z.name !== "BIKES");

        if (type === "main") {
            // Cap = highest officer number in the template sheet
            const sheetName = (currentLane === "car" ? `${currentMode} ${currentShift}` : `${currentLane} ${currentMode} ${currentShift}`).toLowerCase();
            const sheetData = excelData[sheetName] || [];
            const max = sheetData.reduce((m, row) => {
                const n = parseInt(row.Officer);
                return isNaN(n) ? m : Math.max(m, n);
            }, 0);
            if (max === 0) return null; // template not loaded
            // Used = highest officer number already on grid
            const used = [...new Set(
                [...document.querySelectorAll('.counter-cell.active[data-type="main"]')]
                    .map(c => parseInt(c.dataset.officer))
                    .filter(n => !isNaN(n))
            )].length;
            return { used, max, remaining: max - used, label: "main officers" };
        }

        if (type === "trainowc") {
            const sheetName = (currentLane === "car" ? `${currentMode} ${currentShift}` : `${currentLane} ${currentMode} ${currentShift}`).toLowerCase();
            const sheetData = excelData[sheetName] || [];
            const prefix = currentMode === "arrival" ? "TRAIN" : "OWC";
            const allLabels = [...new Set(
                sheetData
                    .filter(r => String(r.Officer || "").trim().toUpperCase().startsWith(prefix))
                    .map(r => r.Officer)
            )];
            const max = allLabels.length;
            const onGrid = new Set(
                [...document.querySelectorAll(".counter-cell.active")]
                    .map(c => c.dataset.officer?.split(" | ")[0])
                    .filter(o => o && o.toUpperCase().startsWith(prefix))
            );
            const used = onGrid.size;
            return { used, max, remaining: max - used, label: prefix + " officers" };
        }

        if (type === "ot" || type === "sos") {
            if (!windowStart || !windowEnd) return null;
            let si = times.findIndex(t => t === windowStart);
            let ei = times.findIndex(t => t === windowEnd);
            if (ei === -1) ei = times.length;
            if (si === -1 || si >= ei) return null;

            // Count counters that are fully free for the entire window
            let freeCounters = 0;
            allZones.forEach(z => {
                z.counters.forEach(c => {
                    let free = true;
                    for (let t = si; t < ei; t++) {
                        const cell = document.querySelector(
                            `.counter-cell[data-zone="${z.name}"][data-time="${t}"][data-counter="${c}"]`
                        );
                        if (!cell || cell.classList.contains("active")) { free = false; break; }
                    }
                    if (free) freeCounters++;
                });
            });

            // Count already-assigned officers of this type in this window
            const assignedLabels = new Set(
                [...document.querySelectorAll(`.counter-cell.active[data-type="${type}"]`)]
                    .filter(c => {
                        const t = parseInt(c.dataset.time);
                        return t >= si && t < ei;
                    })
                    .map(c => c.dataset.officer)
                    .filter(Boolean)
            );
            const used = assignedLabels.size;
            const max = used + freeCounters; // total capacity = already used + still free
            return { used, max, remaining: freeCounters, label: type.toUpperCase() + " slots" };
        }
        return null;
    }

    // Show a toast/inline warning near the add button
    function showCapacityWarning(message, isError = false) {
        let el = document.getElementById("_capacityWarning");
        if (!el) {
            el = document.createElement("div");
            el.id = "_capacityWarning";
            el.style.cssText = "margin-top:6px;padding:6px 10px;border-radius:5px;font-size:12px;font-weight:600;";
            const actionBtns = document.querySelector(".action-buttons");
            if (actionBtns) actionBtns.after(el);
        }
        el.textContent = message;
        el.style.background = isError ? "#ffebee" : "#fff8e1";
        el.style.color = isError ? "#c62828" : "#e65100";
        el.style.border = isError ? "1px solid #ef9a9a" : "1px solid #ffcc80";
        clearTimeout(el._hideTimer);
        el._hideTimer = setTimeout(() => { if (el.parentNode) el.textContent = ""; el.style.border = "none"; }, 5000);
    }

    // Returns false and shows warning if count exceeds remaining capacity
    function checkCapacityBeforeAdd(type, count, windowStart, windowEnd) {
        const cap = getCapacity(type, windowStart, windowEnd);
        if (!cap) return true; // can't determine → allow

        if (cap.remaining <= 0) {
            showCapacityWarning(
                `⛔ Max ${cap.label} reached (${cap.used}/${cap.max}). No more can be added.`,
                true
            );
            return false;
        }

        if (count > cap.remaining) {
            showCapacityWarning(
                `⚠️ Only ${cap.remaining} ${cap.label} remaining (${cap.used}/${cap.max}). ` +
                `Requested ${count} — only ${cap.remaining} will be added.`
            );
            // Don't block — let the allocation run and it will naturally stop at capacity
            return true;
        }

        return true;
    }

    addBtn.addEventListener("click", () => {
        const count = parseInt(document.getElementById("officerCount").value);
        if (!count || count <= 0) return;

        saveState();

        if (manpowerType === "sos") {
            const start = document.getElementById("sosStart").value.replace(":", "");
            const end = document.getElementById("sosEnd").value.replace(":", "");

            const times = generateTimeSlots();
            const shiftEndTime = currentShift === "morning" ? "2200" : "1000";

            // Allow the shift-end boundary time even though it has no rendered column
            const startIdx = times.findIndex(t => t === start);
            const endIdx = end === shiftEndTime ? times.length : times.findIndex(t => t === end);

            if (startIdx === -1 || endIdx === -1) {
                alert("SOS time range outside current shift grid.");
                return;
            }
            if (!checkCapacityBeforeAdd("sos", count, start, end)) return;
            allocateSOSOfficers(count, start, end);
            updateSOSRoster(start, end);
        }

        if (manpowerType === "ot") {
            const slot = document.getElementById("otSlot").value;
            const [start, end] = slot.split("-").map(s => s.replace(":", ""));
            if (!isOTWithinShift(start, end)) {
                alert(`OT ${start}-${end} is outside current shift (${currentShift}).`);
                return;
            }
            if (!checkCapacityBeforeAdd("ot", count, start, end)) return;
            allocateOTOfficers(count, start, end);
        }

        if (manpowerType === "main") {
            if (!checkCapacityBeforeAdd("main", count)) return;
            applyMainTemplate(count);
        }
    });

    /* ── Confirm RA / RO button ─────────────────────────────────────────── */
    if (confirmRaRoBtn) {
        confirmRaRoBtn.addEventListener("click", () => {
            if (manpowerType === "ra") {
                const officerLabel = document.getElementById("raOfficer").value;
                const raTimeRaw = document.getElementById("raTime").value;
                if (!officerLabel || !raTimeRaw) { alert("Select an officer and enter RA time."); return; }
                applyRA(officerLabel, raTimeRaw.replace(":", ""));
            }
            if (manpowerType === "ro") {
                const officerLabel = document.getElementById("roOfficer").value;
                const roTimeRaw = document.getElementById("roTime").value;
                if (!officerLabel || !roTimeRaw) { alert("Select an officer and enter RO time."); return; }
                applyRO(officerLabel, roTimeRaw.replace(":", ""));
            }
        });
    }

    /* ==================== RA / RO ==================== */

    // raroRegistry: officer label → { ra: HHMM|null, ro: HHMM|null, roRelease: HHMM|null }
    const raroRegistry = {};

    function getRaro(officer) {
        return raroRegistry[officer] || { ra: null, ro: null, roRelease: null };
    }

    function addMinsToHHMM(hhmm, mins) {
        const h = parseInt(hhmm.slice(0, 2)), m = parseInt(hhmm.slice(2));
        const total = h * 60 + m + mins;
        return String(Math.floor(total / 60)).padStart(2, "0") + String(total % 60).padStart(2, "0");
    }

    // RA: clear all active cells for this officer BEFORE raTime
    function applyRA(officerLabel, raTime) {
        const times = generateTimeSlots();
        const raIndex = times.findIndex(t => t === raTime);
        if (raIndex === -1) { alert("RA time is outside the shift grid."); return; }

        let affected = 0;
        document.querySelectorAll(`.counter-cell.active[data-type="main"]`).forEach(cell => {
            if (cell.dataset.officer !== officerLabel) return;
            if (parseInt(cell.dataset.time) < raIndex) {
                cell.classList.remove("active");
                cell.style.background = "";
                cell.dataset.officer = "";
                cell.dataset.type = "";
                affected++;
            }
        });

        if (affected === 0) {
            alert(`No cells found before ${formatTime(raTime)} for officer ${officerLabel}.`);
            return;
        }

        raroRegistry[officerLabel] = { ...getRaro(officerLabel), ra: raTime };
        updateAll();
    }

    // RO: clear all active cells for this officer FROM (roTime − 30min) onwards
    function applyRO(officerLabel, roTime) {
        const times = generateTimeSlots();
        const releaseTime = addMinsToHHMM(roTime, -30);
        const releaseIndex = times.findIndex(t => t === releaseTime);
        if (releaseIndex === -1) { alert("RO release time is outside the shift grid."); return; }

        let affected = 0;
        document.querySelectorAll(`.counter-cell.active[data-type="main"]`).forEach(cell => {
            if (cell.dataset.officer !== officerLabel) return;
            if (parseInt(cell.dataset.time) >= releaseIndex) {
                cell.classList.remove("active");
                cell.style.background = "";
                cell.dataset.officer = "";
                cell.dataset.type = "";
                affected++;
            }
        });

        if (affected === 0) {
            alert(`No cells found from ${formatTime(releaseTime)} onwards for officer ${officerLabel}.`);
            return;
        }

        raroRegistry[officerLabel] = { ...getRaro(officerLabel), ro: roTime, roRelease: releaseTime };
        updateAll();
    }

    /* ==================== End RA / RO ==================== */

    document.getElementById("addTrainOwcBtn")?.addEventListener("click", () => {
        const count = parseInt(document.getElementById("trainOwcCount")?.value || "0");
        if (!count || count <= 0) return;
        if (!checkCapacityBeforeAdd("trainowc", count)) return;
        saveState();
        applyTrainOwcTemplate(count);
    });

    removeBtn.addEventListener("click", () => {
        // Collect unique officer labels currently on the grid
        const labelMap = {}; // label → type
        document.querySelectorAll(".counter-cell.active").forEach(c => {
            if (c.dataset.officer) labelMap[c.dataset.officer] = c.dataset.type || "";
        });

        const labels = Object.keys(labelMap);
        if (labels.length === 0) { alert("No officers on grid to remove."); return; }

        // Sort: numeric (main) first ascending, then OT, then SOS
        labels.sort((a, b) => {
            const aNum = parseInt(a), bNum = parseInt(b);
            const aIsNum = !isNaN(aNum) && String(aNum) === a.split(" ")[0];
            const bIsNum = !isNaN(bNum) && String(bNum) === b.split(" ")[0];
            if (aIsNum && bIsNum) return aNum - bNum;
            if (aIsNum) return -1;
            if (bIsNum) return 1;
            return a.localeCompare(b);
        });

        function displayLabel(l) { return l.replace(" | ", "  "); }

        const times = generateTimeSlots();
        const timeOptions = times.map((t, i) => `<option value="${i}">${t}</option>`).join("");

        // Build modal
        const overlay = document.createElement("div");
        overlay.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center;";

        overlay.innerHTML = `
            <div style="background:#fff;border-radius:10px;padding:22px;min-width:320px;max-width:400px;width:90%;box-shadow:0 6px 24px rgba(0,0,0,.35);font-family:inherit;">
                <h3 style="margin:0 0 6px;font-size:15px;color:#333;">Remove Officer</h3>
                <div style="font-size:11px;color:#888;margin-bottom:10px;">
                    Hold <kbd style="background:#f0f0f0;border:1px solid #ccc;border-radius:3px;padding:0 3px;">Ctrl</kbd> /
                    <kbd style="background:#f0f0f0;border:1px solid #ccc;border-radius:3px;padding:0 3px;">⌘</kbd> to select multiple &nbsp;·&nbsp;
                    ${labels.length} officer${labels.length > 1 ? "s" : ""} on grid
                </div>

                <input id="_removeSearch" type="text" placeholder="Search by serial no. or name..."
                    style="width:100%;box-sizing:border-box;padding:8px 10px;margin-bottom:8px;border:1px solid #ccc;border-radius:6px;font-size:13px;outline:none;"/>
                <select id="_removeSelect" multiple size="8"
                    style="width:100%;box-sizing:border-box;border:1px solid #ccc;border-radius:6px;font-size:13px;padding:4px;line-height:1.6;margin-bottom:12px;">
                    ${labels.map(l => {
            const type = labelMap[l];
            const badge = type === "main"
                ? (l.toUpperCase().startsWith("TRAIN") ? "🚂" : l.toUpperCase().startsWith("OWC") ? "🎓" : "🟠")
                : type === "ot" ? "🟣" : type === "sos" ? "🔵" : "⚪";
            return `<option value="${l}">${badge} ${displayLabel(l)}</option>`;
        }).join("")}
                </select>

                <div style="border-top:1px solid #eee;padding-top:12px;margin-bottom:12px;">
                    <div style="font-size:12px;color:#777;margin-bottom:6px;">Remove for time period (optional):</div>
                    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;align-items:center;">
                        <div>
                            <label style="font-size:11px;color:#999;display:block;margin-bottom:3px;">From</label>
                            <select id="_removeFrom" style="width:100%;padding:6px;border:1px solid #ccc;border-radius:6px;font-size:13px;">
                                <option value="-1">— Start (all) —</option>
                                ${timeOptions}
                            </select>
                        </div>
                        <div>
                            <label style="font-size:11px;color:#999;display:block;margin-bottom:3px;">To</label>
                            <select id="_removeTo" style="width:100%;padding:6px;border:1px solid #ccc;border-radius:6px;font-size:13px;">
                                <option value="-1">— End (all) —</option>
                                ${timeOptions}
                            </select>
                        </div>
                    </div>
                    <div id="_removeRangeNote" style="font-size:11px;color:#2196F3;margin-top:5px;min-height:14px;">
                        Entire shift selected (no time filter)
                    </div>
                </div>

                <div style="display:flex;gap:8px;justify-content:flex-end;">
                    <button id="_removeCancel"
                        style="padding:7px 16px;border-radius:6px;border:1px solid #ccc;background:#f5f5f5;cursor:pointer;font-size:13px;">
                        Cancel
                    </button>
                    <button id="_removeConfirm"
                        style="padding:7px 16px;border-radius:6px;border:none;background:#e53935;color:#fff;cursor:pointer;font-weight:600;font-size:13px;">
                        Remove
                    </button>
                </div>
            </div>
        `;

        document.body.appendChild(overlay);
        const searchEl = overlay.querySelector("#_removeSearch");
        const selectEl = overlay.querySelector("#_removeSelect");
        const fromEl = overlay.querySelector("#_removeFrom");
        const toEl = overlay.querySelector("#_removeTo");
        const noteEl = overlay.querySelector("#_removeRangeNote");
        searchEl.focus();

        function updateNote() {
            const fi = parseInt(fromEl.value);
            const ti = parseInt(toEl.value);
            if (fi === -1 && ti === -1) {
                noteEl.textContent = "Entire shift selected (no time filter)";
                noteEl.style.color = "#2196F3";
            } else if (fi !== -1 && ti !== -1 && fi >= ti) {
                noteEl.textContent = "⚠ 'From' must be before 'To'";
                noteEl.style.color = "#c62828";
            } else {
                const fromLabel = fi === -1 ? "start" : times[fi];
                const toLabel = ti === -1 ? "end" : times[ti];
                noteEl.textContent = `Removing ${fromLabel} → ${toLabel} only`;
                noteEl.style.color = "#2e7d32";
            }
        }

        fromEl.onchange = updateNote;
        toEl.onchange = updateNote;

        // Live filter
        searchEl.addEventListener("input", () => {
            const q = searchEl.value.toLowerCase();
            [...selectEl.options].forEach(opt => {
                opt.hidden = q && !opt.text.toLowerCase().includes(q) && !opt.value.toLowerCase().includes(q);
            });
        });

        const close = () => document.body.removeChild(overlay);
        overlay.querySelector("#_removeCancel").addEventListener("click", close);
        overlay.addEventListener("click", e => { if (e.target === overlay) close(); });

        overlay.querySelector("#_removeConfirm").addEventListener("click", () => {
            const targets = [...selectEl.selectedOptions].map(o => o.value);
            if (targets.length === 0) { close(); return; }

            const fi = parseInt(fromEl.value);
            const ti = parseInt(toEl.value);

            // Validate range if both set
            if (fi !== -1 && ti !== -1 && fi >= ti) {
                noteEl.textContent = "⚠ 'From' must be before 'To'";
                noteEl.style.color = "#c62828";
                return;
            }

            saveState();
            document.querySelectorAll(".counter-cell.active").forEach(cell => {
                if (!targets.includes(cell.dataset.officer)) return;
                const t = parseInt(cell.dataset.time);
                // Apply time filter if set
                if (fi !== -1 && t < fi) return;
                if (ti !== -1 && t >= ti) return;
                cell.classList.remove("active");
                cell.style.background = "";
                cell.dataset.officer = "";
                cell.dataset.type = "";
            });
            updateAll();
            close();
        });
    });

    undoBtn.addEventListener("click", () => {
        if (historyStack.length === 0) return;
        const prev = historyStack.pop();
        restoreState(prev);
    });

    /* ── Edit Names modal ─────────────────────────────────────────────── */
    document.getElementById("editNamesBtn")?.addEventListener("click", () => {
        // Collect all OT and SOS officers (main identified by number, not name)
        const labelMap = {}; // base label → { type, currentLabel }
        document.querySelectorAll(".counter-cell.active").forEach(c => {
            const lbl = c.dataset.officer;
            const type = c.dataset.type;
            if (!lbl || type === "main") return;
            if (!labelMap[lbl]) labelMap[lbl] = { type, currentLabel: lbl };
        });

        const entries = Object.values(labelMap);
        if (entries.length === 0) { alert("No OT or SOS officers on grid to edit."); return; }

        // Sort: SOS first, then OT, each numerically
        entries.sort((a, b) => {
            const aBase = a.currentLabel.split(" | ")[0];
            const bBase = b.currentLabel.split(" | ")[0];
            const aIsOT = aBase.startsWith("OT"), bIsOT = bBase.startsWith("OT");
            const aIsSOS = aBase.startsWith("SOS"), bIsSOS = bBase.startsWith("SOS");
            if (aIsSOS && !bIsSOS) return -1;
            if (!aIsSOS && bIsSOS) return 1;
            return parseInt(aBase.replace(/\D/g, "")) || 0 - parseInt(bBase.replace(/\D/g, "")) || 0;
        });

        const overlay = document.createElement("div");
        overlay.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center;";

        const rows = entries.map(e => {
            const base = e.currentLabel.split(" | ")[0];
            const existing = e.currentLabel.includes(" | ") ? e.currentLabel.split(" | ").slice(1).join(" | ") : "";
            const badge = e.type === "ot" ? "🟣" : "🔵";
            return `
                <tr data-old="${e.currentLabel}">
                    <td style="padding:5px 8px;font-size:13px;white-space:nowrap;">${badge} ${base}</td>
                    <td style="padding:5px 4px;">
                        <input type="text" value="${existing}"
                            placeholder="Name..."
                            style="width:100%;box-sizing:border-box;padding:5px 7px;border:1px solid #ccc;border-radius:5px;font-size:13px;"/>
                    </td>
                </tr>`;
        }).join("");

        overlay.innerHTML = `
            <div style="background:#fff;border-radius:10px;padding:22px;min-width:320px;max-width:420px;width:92%;box-shadow:0 6px 24px rgba(0,0,0,.35);font-family:inherit;display:flex;flex-direction:column;max-height:90vh;">
                <h3 style="margin:0 0 14px;font-size:15px;color:#333;">Edit Officer Names</h3>
                <div style="overflow-y:auto;flex:1;margin-bottom:14px;">
                    <table style="width:100%;border-collapse:collapse;">
                        <thead>
                            <tr style="border-bottom:1px solid #eee;">
                                <th style="text-align:left;padding:4px 8px;font-size:11px;color:#888;">Officer</th>
                                <th style="text-align:left;padding:4px 8px;font-size:11px;color:#888;">Name</th>
                            </tr>
                        </thead>
                        <tbody>${rows}</tbody>
                    </table>
                </div>
                <div style="display:flex;gap:8px;justify-content:flex-end;">
                    <button id="_editNamesCancel" style="padding:7px 16px;border-radius:6px;border:1px solid #ccc;background:#f5f5f5;cursor:pointer;font-size:13px;">Cancel</button>
                    <button id="_editNamesSave"   style="padding:7px 16px;border-radius:6px;border:none;background:#1976d2;color:#fff;cursor:pointer;font-weight:600;font-size:13px;">Save</button>
                </div>
            </div>
        `;

        document.body.appendChild(overlay);

        const close = () => document.body.removeChild(overlay);
        overlay.querySelector("#_editNamesCancel").addEventListener("click", close);
        overlay.addEventListener("click", e => { if (e.target === overlay) close(); });

        overlay.querySelector("#_editNamesSave").addEventListener("click", () => {
            saveState();
            // Build a map of old label → new label
            const renameMap = {};
            overlay.querySelectorAll("tbody tr").forEach(row => {
                const oldLabel = row.dataset.old;
                const base = oldLabel.split(" | ")[0];
                const name = row.querySelector("input").value.trim();
                const newLabel = name ? `${base} | ${name}` : base;
                if (newLabel !== oldLabel) renameMap[oldLabel] = newLabel;
            });

            if (Object.keys(renameMap).length === 0) { close(); return; }

            // Apply renames to all matching cells
            document.querySelectorAll(".counter-cell.active").forEach(cell => {
                const renamed = renameMap[cell.dataset.officer];
                if (renamed) cell.dataset.officer = renamed;
            });

            updateAll();
            close();
        });
    });

}); // end DOMContentLoaded

/* ================== HELPERS ================== */
function getEmptyCellsBackFirst(zoneName, timeIndex) {
    const cells = [...document.querySelectorAll(
        `.counter-cell[data-zone="${zoneName}"][data-time="${timeIndex}"]`
    )];

    let emptyCells = cells.filter(c => !c.classList.contains("active"));

    emptyCells.sort((a, b) =>
        parseInt(b.parentElement.firstChild.innerText.replace(/\D/g, '')) -
        parseInt(a.parentElement.firstChild.innerText.replace(/\D/g, ''))
    );

    return emptyCells;
}

function copyMainRoster() { copyTable("mainRosterTable", "copyMainRosterBtn"); }
function copySOSRoster() { copyTable("sosRosterTable", "copySOSRosterBtn"); }
function copyOTRoster() { copyTable("otRosterTable", "copyOTRosterBtn"); }

function copyTable(tableId, btnId) {
    const table = document.getElementById(tableId);
    if (!table) return;

    let text = "";
    table.querySelectorAll("tr").forEach(row => {
        const cells = row.querySelectorAll("th, td");
        const rowText = [];
        cells.forEach(cell => rowText.push(cell.innerText.trim()));
        text += rowText.join("\t") + "\n";
    });

    navigator.clipboard.writeText(text).then(() => {
        const btn = document.getElementById(btnId);
        if (!btn) return;
        const original = btn.textContent;
        btn.textContent = "Copied!";
        btn.classList.add("copied");
        setTimeout(() => {
            btn.textContent = original;
            btn.classList.remove("copied");
        }, 2000);
    });
}
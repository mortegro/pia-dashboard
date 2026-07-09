(function () {
  "use strict";

  const MONTHS = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember"
  ];
  const WEEKDAYS = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];

  // ---------- state ----------
  const state = {
    entries: [],          // { date: Date, employee: string, leistung: string, menge: number }
    employees: [],        // sorted list of all employee codes
    selectedEmployees: new Set(),
    granularity: "week",
    periods: [],
    periodIndex: 0,
    tarifPeriods: [],      // [{ start: Date, end: Date, label, map }], selected automatically by Leistungsdatum
    tarifLabel: "",
    detail: {
      amountFrom: null,
      amountTo: null,
      search: "",
      sortKey: "date",
      sortDir: "asc",
      page: 0,
      pageSize: 50
    }
  };

  // ---------- date helpers ----------
  function pad2(n) { return String(n).padStart(2, "0"); }
  function fmtDay(d) { return `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.`; }
  function fmtDayFull(d) { return `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${d.getFullYear()}`; }
  function dayOnly(d) {
    const r = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    return r;
  }
  function parseIsoDateLocal(s) {
    const [y, m, d] = s.split("-").map(Number);
    return new Date(y, m - 1, d);
  }
  function startOfWeek(d) {
    const s = dayOnly(d);
    const dow = (s.getDay() + 6) % 7; // 0 = Monday
    s.setDate(s.getDate() - dow);
    return s;
  }
  function isoWeekNumber(d) {
    const target = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    const dayNr = (target.getUTCDay() + 6) % 7;
    target.setUTCDate(target.getUTCDate() - dayNr + 3);
    const firstThursday = new Date(Date.UTC(target.getUTCFullYear(), 0, 4));
    const diff = target - firstThursday;
    return 1 + Math.round(diff / (7 * 86400000));
  }

  const GRAN = {
    day: {
      floor: (d) => dayOnly(d),
      add: (d, n) => { const r = dayOnly(d); r.setDate(r.getDate() + n); return r; },
      end: (start) => new Date(start.getFullYear(), start.getMonth(), start.getDate(), 23, 59, 59, 999),
      label: (start) => `${WEEKDAYS[start.getDay()]}, ${fmtDayFull(start)}`,
      short: (start) => `${WEEKDAYS[start.getDay()]} ${fmtDay(start)}`
    },
    week: {
      floor: (d) => startOfWeek(d),
      add: (d, n) => { const r = new Date(d); r.setDate(r.getDate() + 7 * n); return r; },
      end: (start) => { const e = new Date(start); e.setDate(e.getDate() + 6); e.setHours(23, 59, 59, 999); return e; },
      label: (start, end) => `KW ${isoWeekNumber(start)} · ${fmtDay(start)}–${fmtDay(end)}${end.getFullYear()}`,
      short: (start) => `KW ${isoWeekNumber(start)}`
    },
    month: {
      floor: (d) => new Date(d.getFullYear(), d.getMonth(), 1),
      add: (d, n) => new Date(d.getFullYear(), d.getMonth() + n, 1),
      end: (start) => new Date(start.getFullYear(), start.getMonth() + 1, 0, 23, 59, 59, 999),
      label: (start) => `${MONTHS[start.getMonth()]} ${start.getFullYear()}`,
      short: (start) => MONTHS[start.getMonth()].slice(0, 3)
    },
    quarter: {
      floor: (d) => { const q = Math.floor(d.getMonth() / 3); return new Date(d.getFullYear(), q * 3, 1); },
      add: (d, n) => { const q = Math.floor(d.getMonth() / 3); return new Date(d.getFullYear(), (q + n) * 3, 1); },
      end: (start) => new Date(start.getFullYear(), start.getMonth() + 3, 0, 23, 59, 59, 999),
      label: (start) => `Q${Math.floor(start.getMonth() / 3) + 1} ${start.getFullYear()}`,
      short: (start) => `Q${Math.floor(start.getMonth() / 3) + 1}`
    },
    year: {
      floor: (d) => new Date(d.getFullYear(), 0, 1),
      add: (d, n) => new Date(d.getFullYear() + n, 0, 1),
      end: (start) => new Date(start.getFullYear(), 11, 31, 23, 59, 59, 999),
      label: (start) => `${start.getFullYear()}`,
      short: (start) => `${start.getFullYear()}`
    }
  };

  // maps the main Zeitraum granularity to the bar unit used in the "Verlauf" chart tab
  const CHART_SUBGRAN = { week: "day", month: "week", quarter: "month", year: "month" };
  const CHART_COLORS = [
    "#0d6efd", "#6610f2", "#d63384", "#fd7e14", "#198754",
    "#20c997", "#0dcaf0", "#6f42c1", "#dc3545", "#ffc107"
  ];
  function colorForEmployee(emp) {
    const idx = state.employees.indexOf(emp);
    return CHART_COLORS[(idx < 0 ? 0 : idx) % CHART_COLORS.length];
  }

  function buildPeriods(gran, minDate, maxDate) {
    const g = GRAN[gran];
    const firstStart = g.floor(minDate);
    const lastStart = g.floor(maxDate);
    const periods = [];
    let i = 0;
    let cursor = firstStart;
    while (cursor <= lastStart && i < 3000) {
      const end = g.end(cursor);
      periods.push({ start: cursor, end, label: g.label(cursor, end) });
      i++;
      cursor = g.add(firstStart, i);
    }
    return periods;
  }

  // ---------- excel parsing ----------
  function excelSerialToDate(serial) {
    // Excel/1900 date system, matches SheetJS + xlrd behaviour used to build this export
    const utcDays = Math.floor(serial - 25569);
    const utcMs = utcDays * 86400 * 1000;
    return new Date(utcMs);
  }

  function parseDateCell(v) {
    if (v instanceof Date && !isNaN(v)) return v;
    if (typeof v === "number") return excelSerialToDate(v);
    if (typeof v === "string") {
      const m = v.trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
      if (m) {
        let [, dd, mm, yyyy] = m;
        if (yyyy.length === 2) yyyy = "20" + yyyy;
        return new Date(Number(yyyy), Number(mm) - 1, Number(dd));
      }
    }
    return null;
  }

  function extractEntries(workbook) {
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, cellDates: true });

    let headerRowIndex = -1;
    let colDatum = -1, colArzt = -1, colLeistung = -1, colMenge = -1;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!Array.isArray(row)) continue;
      const norm = row.map((c) => (typeof c === "string" ? c.trim() : c));
      const iDatum = norm.indexOf("Datum");
      const iArzt = norm.indexOf("Arzt");
      const iLeistung = norm.indexOf("Leistung");
      if (iDatum !== -1 && iArzt !== -1 && iLeistung !== -1) {
        headerRowIndex = i;
        colDatum = iDatum;
        colArzt = iArzt;
        colLeistung = iLeistung;
        colMenge = norm.indexOf("Menge");
        break;
      }
    }

    if (headerRowIndex === -1) {
      throw new Error(
        "Unbekanntes Dateiformat: Spalten 'Datum', 'Arzt' und 'Leistung' wurden nicht gefunden."
      );
    }

    const entries = [];
    for (let r = headerRowIndex + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!Array.isArray(row)) continue;
      // rows without a Leistungsnummer (column F) are headers, "Patient: ...",
      // "Summe:" and other non-data lines — this is what distinguishes real
      // service rows regardless of how many extra rows the export contains
      const leistungRaw = row[colLeistung];
      const leistung = leistungRaw == null ? "" : String(leistungRaw).trim();
      if (!leistung) continue;
      const date = parseDateCell(row[colDatum]);
      if (!date) continue;
      const employeeRaw = row[colArzt];
      const employee = employeeRaw == null ? "" : String(employeeRaw).trim();
      if (!employee) continue;
      const mengeRaw = colMenge !== -1 ? row[colMenge] : null;
      const menge = Number(mengeRaw) > 0 ? Number(mengeRaw) : 1;
      entries.push({ date: dayOnly(date), employee, leistung, menge });
    }
    return entries;
  }

  // ---------- tariff (Vergütungstabelle) ----------
  function parseTarifCSV(text) {
    const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
    if (!lines.length) throw new Error("Vergütungstabelle ist leer.");
    const header = lines[0].split(";").map((c) => c.trim());
    const colGebNr = header.indexOf("Geb.Nr.");
    const colVerguetung = header.indexOf("Verguetung_EUR");
    if (colGebNr === -1 || colVerguetung === -1) {
      throw new Error(
        "Unbekanntes Format der Vergütungstabelle: Spalten 'Geb.Nr.' und 'Verguetung_EUR' wurden nicht gefunden."
      );
    }
    const colKategorie = header.indexOf("Kategorie");
    const colMinuten = header.indexOf("Minuten");
    const colBeschreibung = header.indexOf("Leistungsbeschreibung");

    const map = new Map();
    for (let i = 1; i < lines.length; i++) {
      const cols = lines[i].split(";");
      const gebNrRaw = cols[colGebNr];
      if (gebNrRaw == null || gebNrRaw.trim() === "") continue;
      const gebNr = parseInt(gebNrRaw, 10);
      if (isNaN(gebNr)) continue;
      const priceRaw = (cols[colVerguetung] || "").trim().replace(/\./g, "").replace(",", ".");
      const verguetung = parseFloat(priceRaw);
      if (isNaN(verguetung)) continue;
      map.set(gebNr, {
        gebNr,
        verguetung,
        kategorie: colKategorie !== -1 ? cols[colKategorie] : "",
        minuten: colMinuten !== -1 ? cols[colMinuten] : "",
        beschreibung: colBeschreibung !== -1 ? cols.slice(colBeschreibung).join(";") : ""
      });
    }
    if (!map.size) throw new Error("Keine gültigen Zeilen in der Vergütungstabelle gefunden.");
    return map;
  }

  function lookupInMap(map, leistungCode) {
    const digits = String(leistungCode).replace(/^[^0-9]+/, "");
    if (!digits) return null;
    const n = parseInt(digits, 10);
    if (isNaN(n)) return null;
    return map.get(n) || null;
  }

  function findTarifPeriod(date) {
    return state.tarifPeriods.find((p) => date >= p.start && date <= p.end) || null;
  }

  // returns { tarif, missing, periodLabel } where missing is null, "period" (no
  // Vergütungstabelle covers this Leistungsdatum) or "code" (table found but
  // Leistungscode not in it); periodLabel names the table that was checked
  function resolveTarif(leistungCode, date) {
    const period = findTarifPeriod(date);
    if (!period) return { tarif: null, missing: "period", periodLabel: null };
    const tarif = lookupInMap(period.map, leistungCode);
    return { tarif, missing: tarif ? null : "code", periodLabel: period.label };
  }

  function applyTarifPeriods(periods) {
    state.tarifPeriods = periods;
    state.tarifLabel = periods.map((p) => p.label).join(" · ");
    updateTarifStatus();
  }

  function updateTarifStatus() {
    const warnBox = el("tarifWarning");
    const errorBox = el("tarifPeriodError");

    if (!state.entries.length) {
      warnBox.classList.add("hidden");
      warnBox.textContent = "";
      errorBox.classList.add("hidden");
      errorBox.textContent = "";
      return;
    }

    const unmatchedCodes = new Set();
    const missingPeriodDates = new Set();
    state.entries.forEach((e) => {
      const { tarif, missing } = resolveTarif(e.leistung, e.date);
      if (tarif) return;
      if (missing === "period") missingPeriodDates.add(fmtDayFull(e.date));
      else unmatchedCodes.add(e.leistung);
    });

    if (missingPeriodDates.size) {
      const dates = Array.from(missingPeriodDates).sort();
      errorBox.textContent =
        `Fehler: Für ${missingPeriodDates.size} Leistungsdatum/-daten liegt keine Vergütungstabelle vor ` +
        `(Vergütung = 0 €): ${dates.slice(0, 20).join(", ")}${dates.length > 20 ? " …" : ""}`;
      errorBox.classList.remove("hidden");
    } else {
      errorBox.classList.add("hidden");
      errorBox.textContent = "";
    }

    if (unmatchedCodes.size) {
      warnBox.textContent =
        `Achtung: ${unmatchedCodes.size} Leistungscode(s) ohne Zuordnung in der Vergütungstabelle (Vergütung = 0 €): ` +
        Array.from(unmatchedCodes).sort().join(", ");
      warnBox.classList.remove("hidden");
    } else {
      warnBox.classList.add("hidden");
      warnBox.textContent = "";
    }

    renderDetailTable();
  }

  async function loadDefaultTarifPeriods() {
    try {
      const res = await fetch("/api/tarif-periods");
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const raw = await res.json();
      if (!Array.isArray(raw) || !raw.length) {
        throw new Error("Keine Vergütungstabellen im Ordner 'pia-verträge' gefunden.");
      }
      const periods = raw
        .map((p) => {
          const start = parseIsoDateLocal(p.start);
          const end = parseIsoDateLocal(p.end);
          end.setHours(23, 59, 59, 999);
          return {
            start,
            end,
            label: `${fmtDayFull(start)}–${fmtDayFull(end)}`,
            map: parseTarifCSV(p.csvText),
          };
        })
        .sort((a, b) => a.start - b.start);
      applyTarifPeriods(periods);
    } catch (err) {
      console.error("Vergütungstabellen konnten nicht geladen werden:", err);
      applyTarifPeriods([]);
    }
  }

  // ---------- rendering ----------
  const el = (id) => document.getElementById(id);

  // ---------- detail (verification) table ----------
  function getDetailFiltered() {
    const d = state.detail;
    const period = state.periods[state.periodIndex];
    const needle = d.search.trim().toLowerCase();
    return state.entries
      .filter((e) => state.selectedEmployees.has(e.employee))
      .filter((e) => !period || (e.date >= period.start && e.date <= period.end))
      .map((e) => {
        const { tarif, missing, periodLabel } = resolveTarif(e.leistung, e.date);
        const price = tarif ? tarif.verguetung : 0;
        const amount = price * e.menge;
        return {
          ...e,
          price,
          amount,
          tarifFound: !!tarif,
          tarifMissing: missing,
          tarifPeriodLabel: periodLabel,
          beschreibung: tarif ? tarif.beschreibung : ""
        };
      })
      .filter((e) => d.amountFrom == null || e.amount >= d.amountFrom)
      .filter((e) => d.amountTo == null || e.amount <= d.amountTo)
      .filter((e) => {
        if (!needle) return true;
        return (
          e.leistung.toLowerCase().includes(needle) ||
          e.employee.toLowerCase().includes(needle) ||
          e.beschreibung.toLowerCase().includes(needle)
        );
      });
  }

  function sortDetailRows(rows) {
    const { sortKey, sortDir } = state.detail;
    const mult = sortDir === "asc" ? 1 : -1;
    const rowsCopy = rows.slice();
    rowsCopy.sort((a, b) => {
      let av, bv;
      switch (sortKey) {
        case "employee": av = a.employee; bv = b.employee; break;
        case "leistung": av = a.leistung; bv = b.leistung; break;
        case "menge": av = a.menge; bv = b.menge; break;
        case "price": av = a.price; bv = b.price; break;
        case "amount": av = a.amount; bv = b.amount; break;
        case "date":
        default: av = a.date; bv = b.date; break;
      }
      if (av < bv) return -1 * mult;
      if (av > bv) return 1 * mult;
      return 0;
    });
    return rowsCopy;
  }

  function renderDetailTable() {
    const tbody = el("detailTableBody");
    if (!tbody) return;
    const emptyEl = el("detailTableEmpty");

    document.querySelectorAll("#detailTable th.sortable").forEach((th) => {
      th.classList.remove("sort-asc", "sort-desc");
      if (th.dataset.sort === state.detail.sortKey) {
        th.classList.add(state.detail.sortDir === "asc" ? "sort-asc" : "sort-desc");
      }
    });

    const filtered = sortDetailRows(getDetailFiltered());
    const totalAmount = filtered.reduce((s, e) => s + e.amount, 0);

    const pageSize = state.detail.pageSize;
    const pageCount = Math.max(1, Math.ceil(filtered.length / pageSize));
    if (state.detail.page >= pageCount) state.detail.page = pageCount - 1;
    if (state.detail.page < 0) state.detail.page = 0;
    const startIdx = state.detail.page * pageSize;
    const pageRows = filtered.slice(startIdx, startIdx + pageSize);

    tbody.innerHTML = "";
    pageRows.forEach((e) => {
      const tr = document.createElement("tr");
      if (e.tarifFound) {
        tr.title = `Vergütungstabelle: ${e.tarifPeriodLabel}`;
      } else if (e.tarifMissing === "period") {
        tr.classList.add("table-danger");
        tr.title = "Keine Vergütungstabelle für dieses Leistungsdatum vorhanden (Vergütung = 0 €)";
      } else {
        tr.classList.add("table-warning");
        tr.title = `Vergütungstabelle: ${e.tarifPeriodLabel} — kein Tarif für diesen Leistungscode gefunden (Vergütung = 0 €)`;
      }
      const cells = [
        fmtDayFull(e.date),
        e.employee,
        e.leistung,
        e.beschreibung || "–",
        e.menge,
        fmtEuro(e.price),
        fmtEuro(e.amount)
      ];
      cells.forEach((val) => {
        const td = document.createElement("td");
        td.textContent = val;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    emptyEl.classList.toggle("hidden", filtered.length > 0);

    el("detailSummary").textContent =
      `${filtered.length} Eintraege · Summe Vergütung: ${fmtEuro(totalAmount)}`;
    el("detailPageLabel").textContent = `Seite ${state.detail.page + 1} / ${pageCount}`;
    el("detailPrevPage").disabled = state.detail.page <= 0;
    el("detailNextPage").disabled = state.detail.page >= pageCount - 1;
  }

  function renderEmployeeList() {
    const list = el("employeeList");
    list.innerHTML = "";
    state.employees.forEach((emp, i) => {
      const wrap = document.createElement("div");
      wrap.className = "form-check";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.className = "form-check-input";
      cb.id = `emp-cb-${i}`;
      cb.checked = state.selectedEmployees.has(emp);
      cb.addEventListener("change", () => {
        if (cb.checked) state.selectedEmployees.add(emp);
        else state.selectedEmployees.delete(emp);
        renderEmployeeCount();
        state.detail.page = 0;
        renderDashboard();
        renderDetailTable();
        renderChartView();
      });
      const label = document.createElement("label");
      label.className = "form-check-label small";
      label.htmlFor = cb.id;
      label.textContent = emp;
      wrap.appendChild(cb);
      wrap.appendChild(label);
      list.appendChild(wrap);
    });
  }

  function renderEmployeeCount() {
    const total = state.employees.length;
    const selected = state.selectedEmployees.size;
    el("employeeCount").textContent = `(${selected}/${total})`;
  }

  function renderPeriodLabel() {
    const p = state.periods[state.periodIndex];
    el("periodLabel").textContent = p ? p.label : "–";
    el("prevPeriod").disabled = state.periodIndex <= 0;
    el("nextPeriod").disabled = state.periodIndex >= state.periods.length - 1;
  }

  function renderDashboard() {
    renderPeriodLabel();
    const period = state.periods[state.periodIndex];
    const tableEmptyEl = el("tableEmpty");
    const tbody = el("dataTableBody");
    tbody.innerHTML = "";

    if (!period) {
      tableEmptyEl.classList.remove("hidden");
      el("totalCount").textContent = "0";
      el("summaryRow").innerHTML = "";
      return;
    }

    const filtered = state.entries.filter(
      (e) =>
        e.date >= period.start &&
        e.date <= period.end &&
        state.selectedEmployees.has(e.employee)
    );

    const stats = new Map(); // employee -> { count, amount }
    filtered.forEach((e) => {
      const { tarif } = resolveTarif(e.leistung, e.date);
      const price = tarif ? tarif.verguetung : 0;
      const s = stats.get(e.employee) || { count: 0, amount: 0 };
      s.count += 1;
      s.amount += price * e.menge;
      stats.set(e.employee, s);
    });

    const rows = Array.from(stats.entries()).sort((a, b) => b[1].amount - a[1].amount);
    const total = filtered.length;
    const totalAmount = Array.from(stats.values()).reduce((s, v) => s + v.amount, 0);
    const maxAmount = rows.length ? rows[0][1].amount : 0;
    const activeEmployees = rows.length;
    const avg = activeEmployees ? (total / activeEmployees).toFixed(1) : "0";

    const statCard = (value, label) => `
      <div class="col"><div class="card h-100"><div class="card-body">
        <div class="fs-4 fw-bold">${value}</div>
        <div class="small text-secondary">${label}</div>
      </div></div></div>`;
    el("summaryRow").innerHTML =
      statCard(total, "Leistungen im Zeitraum") +
      statCard(fmtEuro(totalAmount), "Vergütung im Zeitraum") +
      statCard(activeEmployees, "Aktive Mitarbeiter") +
      statCard(avg, "Ø Leistungen / Mitarbeiter");

    if (!rows.length) {
      tableEmptyEl.classList.remove("hidden");
      el("totalCount").textContent = "0";
      el("totalAmount").textContent = fmtEuro(0);
      return;
    }
    tableEmptyEl.classList.add("hidden");

    rows.forEach(([emp, s]) => {
      const tr = document.createElement("tr");
      const tdEmp = document.createElement("td");
      tdEmp.textContent = emp;
      const tdCount = document.createElement("td");
      tdCount.textContent = s.count;
      const tdAmount = document.createElement("td");
      tdAmount.textContent = fmtEuro(s.amount);
      const tdBar = document.createElement("td");
      const track = document.createElement("div");
      track.className = "progress";
      track.style.height = "8px";
      const fill = document.createElement("div");
      fill.className = "progress-bar";
      fill.style.width = `${maxAmount ? (s.amount / maxAmount) * 100 : 0}%`;
      track.appendChild(fill);
      tdBar.appendChild(track);
      tr.appendChild(tdEmp);
      tr.appendChild(tdCount);
      tr.appendChild(tdAmount);
      tr.appendChild(tdBar);
      tbody.appendChild(tr);
    });

    el("totalCount").textContent = total;
    el("totalAmount").textContent = fmtEuro(totalAmount);
  }

  function fmtEuro(v) {
    return v.toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €";
  }

  // ---------- "Verlauf" chart (stacked bars per week/month within the selected period) ----------
  const CHART_BAR_HEIGHT = 220;

  function renderChartView() {
    const container = el("chartRow");
    const legend = el("chartLegend");
    const unsupportedEl = el("chartUnsupported");
    const emptyEl = el("chartEmpty");
    container.innerHTML = "";
    legend.innerHTML = "";
    unsupportedEl.classList.add("hidden");
    emptyEl.classList.add("hidden");

    const period = state.periods[state.periodIndex];
    const subGran = CHART_SUBGRAN[state.granularity];
    if (!period || !subGran) {
      unsupportedEl.textContent =
        "Der Verlauf ist für die Wochenansicht nicht verfügbar – bitte Monat, Quartal oder Jahr auswählen.";
      unsupportedEl.classList.remove("hidden");
      return;
    }

    const bars = buildPeriods(subGran, period.start, period.end).map((sp) => {
      const perEmployee = new Map();
      state.entries.forEach((e) => {
        if (!state.selectedEmployees.has(e.employee)) return;
        if (e.date < sp.start || e.date > sp.end) return;
        const { tarif } = resolveTarif(e.leistung, e.date);
        const price = tarif ? tarif.verguetung : 0;
        perEmployee.set(e.employee, (perEmployee.get(e.employee) || 0) + price * e.menge);
      });
      const total = Array.from(perEmployee.values()).reduce((a, b) => a + b, 0);
      return { period: sp, perEmployee, total };
    });

    const maxTotal = bars.reduce((m, b) => Math.max(m, b.total), 0);
    if (!maxTotal) {
      emptyEl.classList.remove("hidden");
      return;
    }

    const presentEmployees = new Set();

    bars.forEach((b) => {
      const col = document.createElement("div");
      col.className = "chart-col";

      const totalLabel = document.createElement("div");
      totalLabel.className = "chart-total small text-secondary";
      totalLabel.textContent = b.total ? fmtEuro(b.total) : "";

      const bar = document.createElement("div");
      bar.className = "chart-bar";
      bar.style.height = `${CHART_BAR_HEIGHT}px`;

      Array.from(b.perEmployee.entries())
        .filter(([, amount]) => amount > 0)
        .sort((x, y) => x[0].localeCompare(y[0]))
        .forEach(([emp, amount]) => {
          presentEmployees.add(emp);
          const seg = document.createElement("div");
          seg.className = "chart-segment";
          seg.style.height = `${(amount / maxTotal) * CHART_BAR_HEIGHT}px`;
          seg.style.background = colorForEmployee(emp);
          seg.title = `${emp}: ${fmtEuro(amount)}`;
          bar.appendChild(seg);
        });

      const axisLabel = document.createElement("div");
      axisLabel.className = "chart-axis-label small text-secondary";
      axisLabel.textContent = GRAN[subGran].short(b.period.start, b.period.end);

      col.appendChild(totalLabel);
      col.appendChild(bar);
      col.appendChild(axisLabel);
      container.appendChild(col);
    });

    Array.from(presentEmployees)
      .sort()
      .forEach((emp) => {
        const item = document.createElement("div");
        item.className = "d-flex align-items-center gap-2 small";
        const swatch = document.createElement("span");
        swatch.className = "chart-swatch";
        swatch.style.background = colorForEmployee(emp);
        item.appendChild(swatch);
        item.appendChild(document.createTextNode(emp));
        legend.appendChild(item);
      });
  }

  function initFromEntries(entries) {
    state.entries = entries;
    state.employees = Array.from(new Set(entries.map((e) => e.employee))).sort();
    state.selectedEmployees = new Set(state.employees);

    const dates = entries.map((e) => e.date);
    const minDate = new Date(Math.min(...dates));
    const maxDate = new Date(Math.max(...dates));
    state.periods = buildPeriods(state.granularity, minDate, maxDate);
    state.periodIndex = state.periods.length - 1;

    state.detail.amountFrom = null;
    state.detail.amountTo = null;
    state.detail.search = "";
    state.detail.page = 0;
    el("detailAmountFrom").value = "";
    el("detailAmountTo").value = "";
    el("detailSearch").value = "";

    updateTarifStatus();
    renderEmployeeList();
    renderEmployeeCount();
    renderDashboard();
    renderChartView();

    el("emptyState").classList.add("hidden");
    el("loadedContent").classList.remove("hidden");
    el("headerFileBtn").classList.remove("hidden");
    el("filterDrawerToggle").classList.remove("hidden");
    el("filterDrawer").classList.remove("hidden");
  }

  function rebuildPeriodsKeepingDate() {
    if (!state.entries.length) return;
    const currentPeriod = state.periods[state.periodIndex];
    const anchorDate = currentPeriod ? currentPeriod.start : state.entries[0].date;

    const dates = state.entries.map((e) => e.date);
    const minDate = new Date(Math.min(...dates));
    const maxDate = new Date(Math.max(...dates));
    state.periods = buildPeriods(state.granularity, minDate, maxDate);

    let idx = state.periods.findIndex((p) => anchorDate >= p.start && anchorDate <= p.end);
    if (idx === -1) idx = state.periods.length - 1;
    state.periodIndex = idx;
    state.detail.page = 0;
    renderDashboard();
    renderDetailTable();
    renderChartView();
  }

  // ---------- file loading ----------
  function loadFile(file) {
    if (!file) return;
    el("fileStatus").textContent = "Lade …";
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const entries = extractEntries(workbook);
        if (!entries.length) {
          throw new Error("Keine Leistungszeilen in der Datei gefunden.");
        }
        initFromEntries(entries);
        el("fileLabel").textContent = "Andere Datei laden";
        el("fileStatus").textContent = `${file.name} · ${entries.length} Leistungen`;
      } catch (err) {
        console.error(err);
        el("fileStatus").textContent = "";
        alert("Fehler beim Einlesen der Datei:\n" + err.message);
      }
    };
    reader.onerror = () => {
      el("fileStatus").textContent = "";
      alert("Datei konnte nicht gelesen werden.");
    };
    reader.readAsArrayBuffer(file);
  }

  // ---------- wiring ----------
  document.addEventListener("DOMContentLoaded", () => {
    loadDefaultTarifPeriods();

    el("fileInput").addEventListener("change", (e) => {
      loadFile(e.target.files[0]);
    });
    el("emptyFileInput").addEventListener("change", (e) => {
      loadFile(e.target.files[0]);
    });

    document.getElementById("granularity").addEventListener("click", (e) => {
      const btn = e.target.closest("button[data-gran]");
      if (!btn) return;
      document.querySelectorAll("#granularity button").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      state.granularity = btn.dataset.gran;
      rebuildPeriodsKeepingDate();
    });

    el("prevPeriod").addEventListener("click", () => {
      if (state.periodIndex > 0) {
        state.periodIndex--;
        state.detail.page = 0;
        renderDashboard();
        renderDetailTable();
        renderChartView();
      }
    });
    el("nextPeriod").addEventListener("click", () => {
      if (state.periodIndex < state.periods.length - 1) {
        state.periodIndex++;
        state.detail.page = 0;
        renderDashboard();
        renderDetailTable();
        renderChartView();
      }
    });

    el("selectAll").addEventListener("click", () => {
      state.selectedEmployees = new Set(state.employees);
      renderEmployeeList();
      renderEmployeeCount();
      state.detail.page = 0;
      renderDashboard();
      renderDetailTable();
      renderChartView();
    });
    el("selectNone").addEventListener("click", () => {
      state.selectedEmployees = new Set();
      renderEmployeeList();
      renderEmployeeCount();
      state.detail.page = 0;
      renderDashboard();
      renderDetailTable();
      renderChartView();
    });

    // right-side filter drawer (Zeitraum & Mitarbeiter), shared by both views
    // hidden until data is loaded; toggle button appears once loaded
    el("filterDrawerToggle").addEventListener("click", () => {
      el("filterDrawer").classList.toggle("hidden");
    });
    el("drawerClose").addEventListener("click", () => {
      el("filterDrawer").classList.add("hidden");
    });

    // view tabs
    el("viewTabs").addEventListener("click", (e) => {
      const btn = e.target.closest("button[data-view]");
      if (!btn) return;
      document.querySelectorAll("#viewTabs button").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      const view = btn.dataset.view;
      el("overviewView").classList.toggle("hidden", view !== "overview");
      el("detailView").classList.toggle("hidden", view !== "detail");
      el("chartView").classList.toggle("hidden", view !== "chart");
      if (view === "detail") renderDetailTable();
      if (view === "chart") renderChartView();
    });

    // detail amount / search filters
    el("detailAmountFrom").addEventListener("input", (e) => {
      state.detail.amountFrom = e.target.value === "" ? null : Number(e.target.value);
      state.detail.page = 0;
      renderDetailTable();
    });
    el("detailAmountTo").addEventListener("input", (e) => {
      state.detail.amountTo = e.target.value === "" ? null : Number(e.target.value);
      state.detail.page = 0;
      renderDetailTable();
    });
    el("detailSearch").addEventListener("input", (e) => {
      state.detail.search = e.target.value;
      state.detail.page = 0;
      renderDetailTable();
    });
    el("detailResetFilters").addEventListener("click", () => {
      state.detail.amountFrom = null;
      state.detail.amountTo = null;
      state.detail.search = "";
      state.detail.page = 0;
      el("detailAmountFrom").value = "";
      el("detailAmountTo").value = "";
      el("detailSearch").value = "";
      renderDetailTable();
    });

    // sortable detail table headers
    document.querySelectorAll("#detailTable th.sortable").forEach((th) => {
      th.addEventListener("click", () => {
        const key = th.dataset.sort;
        if (state.detail.sortKey === key) {
          state.detail.sortDir = state.detail.sortDir === "asc" ? "desc" : "asc";
        } else {
          state.detail.sortKey = key;
          state.detail.sortDir = "asc";
        }
        renderDetailTable();
      });
    });

    // detail pagination
    el("detailPrevPage").addEventListener("click", () => {
      if (state.detail.page > 0) {
        state.detail.page--;
        renderDetailTable();
      }
    });
    el("detailNextPage").addEventListener("click", () => {
      state.detail.page++;
      renderDetailTable();
    });

    // drag & drop anywhere on the page
    const dropHint = el("dropHint");
    let dragCounter = 0;
    document.addEventListener("dragenter", (e) => {
      e.preventDefault();
      dragCounter++;
      dropHint.classList.remove("hidden");
    });
    document.addEventListener("dragleave", (e) => {
      e.preventDefault();
      dragCounter--;
      if (dragCounter <= 0) {
        dragCounter = 0;
        dropHint.classList.add("hidden");
      }
    });
    document.addEventListener("dragover", (e) => e.preventDefault());
    document.addEventListener("drop", (e) => {
      e.preventDefault();
      dragCounter = 0;
      dropHint.classList.add("hidden");
      const file = e.dataTransfer.files[0];
      if (file) loadFile(file);
    });
  });
})();

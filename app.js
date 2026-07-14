(() => {
  "use strict";

  const CONFIG = {
    appName: "bbz CRM",

    graph: {
      tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
      clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
      authority: "https://login.microsoftonline.com/3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
      redirectUri: "https://markusbaechler.github.io/crm-spa/",
      // FIX 3a: Scope auf ReadWrite erweitert — verhindert zweiten Login-Prompt beim Write-Layer
      scopes: ["User.Read", "Sites.ReadWrite.All"]
    },

    sharePoint: {
      siteHostname: "bbzsg.sharepoint.com",
      sitePath: "/sites/CRM"
    },

    lists: {
      firms: "CRMFirms",
      contacts: "CRMContacts",
      history: "CRMHistory",
      tasks: "CRMTasks"
    },

    defaults: {
      route: "firms",
      contactArchiveDefaultHidden: true,
      planningShowOnlyOpen: true,
      // Firma für Privatpersonen ohne Firmenbezug — exakter SP-Titel
      privateFirmTitle: "Privatpersonen"
    }
  };

  const SCHEMA = {
    firms: {
      listTitle: CONFIG.lists.firms,
      fields: {
        title: "Title",
        adresse: "Adresse",
        plz: "PLZ",
        ort: "Ort",
        land: "Land",
        hauptnummer: "Hauptnummer",
        klassifizierung: "Klassifizierung",
        vip: "VIP",
        kategorie: "Kategorie"
      }
    },

    contacts: {
      listTitle: CONFIG.lists.contacts,
      fields: {
        nachname: "Title",
        vorname: "Vorname",
        anrede: "Anrede",
        firma: "Firma",
        firmaLookupId: "FirmaLookupId",
        funktion: "Funktion",
        email1: "Email1",
        email2: "Email2",
        direktwahl: "Direktwahl",
        mobile: "Mobile",
        rolle: "Rolle",
        leadbbz0: "Leadbbz0",
        sgf: "SGF",
        geburtstag: "Geburtstag",
        kommentar: "Kommentar",
        event: "Event",
        eventhistory: "Eventhistory",
        archiviert: "Archiviert"
      }
    },

    history: {
      listTitle: CONFIG.lists.history,
      fields: {
        title: "Title",
        kontakt: "Nachname",
        kontaktLookupId: "NachnameLookupId",
        datum: "Datum",
        // KORREKTUR: SP-Feldname ist "Kontaktart", nicht "Typ"
        typ: "Kontaktart",
        notizen: "Notizen",
        projektbezug: "Projektbezug",
        leadbbz: "Leadbbz"
      }
    },

    tasks: {
      listTitle: CONFIG.lists.tasks,
      fields: {
        title: "Title",
        kontakt: "Name",
        kontaktLookupId: "NameLookupId",
        deadline: "Deadline",
        status: "Status",
        leadbbz: "Leadbbz"
      }
    }
  };

  const state = {
    auth: {
      msal: null,
      account: null,
      token: null,
      isAuthenticated: false,
      isReady: false
    },

    meta: {
      siteId: null,
      loading: false,
      lastError: null,
      // Choice-Werte aus SharePoint — pro Liste, pro SP-Feldname
      // Struktur: { "CRMContacts": { "Anrede": ["Herr", "Frau", ...], ... }, ... }
      choices: {},
      // ID der Firma "Privatpersonen" — wird nach enrich() automatisch gesetzt
      privateFirmId: null
    },

    data: {
      firms: [],
      contacts: [],
      history: [],
      tasks: []
    },

    enriched: {
      firms: [],
      contacts: [],
      history: [],
      tasks: [],
      events: []
    },

    filters: {
      route: CONFIG.defaults.route,
      firms: { kategorie: "Kunde", klassifizierung: "", vip: false, search: "", legendeOffen: false, sortBy: "title", sortDir: "asc" },
      contacts: { search: "", archiviertAusblenden: CONFIG.defaults.contactArchiveDefaultHidden, sortBy: "fullName", sortDir: "asc" },
      planning: { search: "", onlyOpen: CONFIG.defaults.planningShowOnlyOpen, groupBy: "none", sortBy: "deadline", sortDir: "asc", segment: "", leadbbz: "", faelligkeit: "" },
      history: { search: "", kontaktart: "", viewMode: "firms", lens: "", granularitaet: "monat", periode: "", expandedFirms: [] },
      events: { search: "", onlyWithOpenTasks: false, sortBy: "contactName", sortDir: "asc", segment: "", selectedEvent: "" },
      admin: { zeitfenster: "30" }
    },

    selection: {
      firmId: null,
      contactId: null
    },

    // Modal-State fuer Write-Layer
    modal: null
  };

  const helpers = {
    escapeHtml(value) {
      return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
    },

    bool(value) {
      if (typeof value === "boolean") return value;
      if (typeof value === "number") return value === 1;
      if (typeof value === "string") {
        const v = value.trim().toLowerCase();
        return ["true", "1", "ja", "yes"].includes(v);
      }
      return false;
    },

    isEmpty(value) {
      return value === null || value === undefined || value === "";
    },

    toArray(value) {
      if (Array.isArray(value)) return value;
      if (value === null || value === undefined || value === "") return [];
      if (typeof value === "string") {
        if (value.includes(";#")) return value.split(";#").map(v => v.trim()).filter(Boolean);
        if (value.includes(",")) return value.split(",").map(v => v.trim()).filter(Boolean);
        return [value.trim()].filter(Boolean);
      }
      return [value];
    },

    normalizeChoiceList(value) {
      return helpers.toArray(value).filter(Boolean);
    },

    toDate(value) {
      if (!value) return null;
      const s = typeof value === "string" ? value.trim() : null;
      if (!s) return null;
      // Alle ISO-Datumsstrings aus SharePoint (YYYY-MM-DD oder YYYY-MM-DDT...)
      // werden als lokales Datum interpretiert — der Datums-Teil (YYYY-MM-DD) wird
      // direkt verwendet ohne UTC-Konvertierung, da SP Datumsfelder ohne Uhrzeit speichert
      // und der UTC-Shift in CH (UTC+1/+2) sonst den Tag verschiebt.
      const dateOnly = /^(\d{4}-\d{2}-\d{2})/.exec(s);
      if (dateOnly) {
        const [y, m, day] = dateOnly[1].split("-").map(Number);
        return new Date(y, m - 1, day);
      }
      const d = new Date(value);
      return Number.isNaN(d.getTime()) ? null : d;
    },

    formatDate(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      return d.toLocaleDateString("de-CH", { day: "2-digit", month: "2-digit", year: "numeric" });
    },

    formatDateTime(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      return d.toLocaleString("de-CH", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
    },

    // FIX 1: fehlende Hilfsfunktion fuer <input type="date"> — gibt YYYY-MM-DD zurueck
    // Wichtig: Lokale Datum-Komponenten verwenden (nicht toISOString = UTC),
    // sonst verschiebt sich das Datum in Zeitzonen wie CH (UTC+1) um einen Tag
    toDateInput(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");
      return `${y}-${m}-${day}`;
    },

    todayStart() {
      const d = new Date();
      d.setHours(0, 0, 0, 0);
      return d;
    },

    isOpenTask(status) {
      const v = String(status || "").trim().toLowerCase();
      return !["erledigt", "geschlossen", "completed", "done", "closed"].includes(v);
    },

    isOverdue(deadline) {
      const d = helpers.toDate(deadline);
      if (!d) return false;
      return d < helpers.todayStart();
    },

    compareDateAsc(a, b) {
      const ad = helpers.toDate(a), bd = helpers.toDate(b);
      if (!ad && !bd) return 0;
      if (!ad) return 1;
      if (!bd) return -1;
      return ad - bd;
    },

    compareDateDesc(a, b) {
      const ad = helpers.toDate(a), bd = helpers.toDate(b);
      if (!ad && !bd) return 0;
      if (!ad) return 1;
      if (!bd) return -1;
      return bd - ad;
    },

    textIncludes(haystack, needle) {
      return String(haystack || "").toLowerCase().includes(String(needle || "").toLowerCase());
    },

    joinNonEmpty(values, sep = " · ") {
      return values.filter(v => !helpers.isEmpty(v)).join(sep);
    },

    fullName(contact) {
      return helpers.joinNonEmpty([contact.vorname, contact.nachname], " ").trim();
    },

    firmBadgeClass(value) {
      const v = String(value || "").toUpperCase();
      if (v === "A" || v === "A-KUNDE") return "bbz-pill bbz-pill-a";
      if (v === "B" || v === "B-KUNDE") return "bbz-pill bbz-pill-b";
      if (v === "C" || v === "C-KUNDE") return "bbz-pill bbz-pill-c";
      return "bbz-pill";
    },

    // Leadbbz als farbiges Pill
    leadbbzBadgeHtml(value) {
      if (!value) return '<span class="bbz-muted">—</span>';
      return `<span class="bbz-pill bbz-pill-lead">${helpers.escapeHtml(value)}</span>`;
    },

    // Detailband-Klasse je nach Segment und VIP
    detailBandClass(firm) {
      if (!firm) return "bbz-detail-band-default";
      if (firm.vip) return "bbz-detail-band bbz-detail-band-vip";
      const v = String(firm.klassifizierung || "").toUpperCase();
      if (v.includes("A")) return "bbz-detail-band bbz-detail-band-a";
      if (v.includes("B")) return "bbz-detail-band bbz-detail-band-b";
      if (v.includes("C")) return "bbz-detail-band bbz-detail-band-c";
      return "bbz-detail-band bbz-detail-band-default";
    },

    statusClass(status, deadline) {
      if (!helpers.isOpenTask(status)) return "bbz-success";
      if (helpers.isOverdue(deadline)) return "bbz-danger";
      return "bbz-warning";
    },

    multiChoiceHtml(values) {
      const list = helpers.normalizeChoiceList(values);
      if (!list.length) return '<span class="bbz-muted">—</span>';
      return list.map(v => `<span class="bbz-chip">${helpers.escapeHtml(v)}</span>`).join("");
    },

    // Avatar-Initialen: gibt fertiges HTML-Element zurück
    // Farbe wird deterministisch aus dem Namen gehasht (0–5)
    avatarHtml(contact) {
      const first = String(contact.vorname || "").charAt(0).toUpperCase();
      const last  = String(contact.nachname || "").charAt(0).toUpperCase();
      const initials = (first + last) || "?";
      // Einfacher Hash aus Zeichencodes
      const seed = [...initials].reduce((s, c) => s + c.charCodeAt(0), 0);
      const idx  = seed % 6;
      return `<span class="bbz-avatar" data-idx="${idx}">${helpers.escapeHtml(initials)}</span>`;
    },

    // Status-Chip: gibt ein farbiges Pill-HTML zurück
    // status: Taskstatus-String, deadline: ISO-Datum
    statusChipHtml(status, deadline) {
      if (!helpers.isOpenTask(status)) {
        return `<span class="bbz-status-chip bbz-status-done">${helpers.escapeHtml(status || "Erledigt")}</span>`;
      }
      if (helpers.isOverdue(deadline)) {
        return `<span class="bbz-status-chip bbz-status-overdue">${helpers.escapeHtml(status || "Überfällig")}</span>`;
      }
      return `<span class="bbz-status-chip bbz-status-open">${helpers.escapeHtml(status || "Offen")}</span>`;
    },

    // Relatives Datum: "heute", "gestern", "vor 3 Tagen", "vor 2 Wochen"
    // Fällt nach 60 Tagen auf formatDate zurück
    relativeDate(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      const today = helpers.todayStart();
      const diffMs = today - d;
      const diffDays = Math.floor(diffMs / 86400000);
      if (diffDays < 0) {
        const futureDays = Math.abs(diffDays);
        if (futureDays === 1) return "morgen";
        if (futureDays < 7) return `in ${futureDays} Tagen`;
        if (futureDays < 14) return "nächste Woche";
        return helpers.formatDate(value);
      }
      if (diffDays === 0) return "heute";
      if (diffDays === 1) return "gestern";
      if (diffDays < 7) return `vor ${diffDays} Tagen`;
      if (diffDays < 14) return "vor 1 Woche";
      if (diffDays < 30) return `vor ${Math.floor(diffDays / 7)} Wochen`;
      if (diffDays < 60) return `vor ${Math.floor(diffDays / 30)} Monat${Math.floor(diffDays / 30) > 1 ? "en" : ""}`;
      return helpers.formatDate(value);
    },

    // Bare Dauer seit `value`, exakte Differenz, ohne Prefix:
    // "1 Tag" | "3 Tagen" | "2 Wochen" | "5 Monaten" (Monate via getFullYear/getMonth).
    agePhrase(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      const today = helpers.todayStart();
      const diffDays = Math.floor((today - d) / 86400000);
      if (diffDays <= 1) return "1 Tag";
      if (diffDays < 7) return `${diffDays} Tagen`;
      if (diffDays < 30) { const w = Math.floor(diffDays / 7); return `${w} Woche${w !== 1 ? "n" : ""}`; }
      const m = (today.getFullYear() - d.getFullYear()) * 12 + (today.getMonth() - d.getMonth());
      return `${m} Monat${m !== 1 ? "en" : ""}`;
    },

    // Geburtstage: gibt Kontakte mit Geburtstag in den nächsten `days` Tagen zurück
    // Jahresunabhängig — nur Monat und Tag werden verglichen
    // contacts-Parameter optional — wenn nicht gesetzt, alle nicht-archivierten Kontakte
    // Rückgabe: [{ contact, daysUntil, nextBirthday (Date), age (number|null) }], aufsteigend sortiert
    upcomingBirthdays(days = 30, contacts = null) {
      const today = helpers.todayStart();
      const source = contacts || state.enriched.contacts.filter(c => !c.archiviert);
      const result = [];
      for (const c of source) {
        if (!c.geburtstag) continue;
        const bDay = helpers.toDate(c.geburtstag);
        if (!bDay) continue;
        let next = new Date(today.getFullYear(), bDay.getMonth(), bDay.getDate());
        if (next < today) next = new Date(today.getFullYear() + 1, bDay.getMonth(), bDay.getDate());
        const daysUntil = Math.round((next - today) / 86400000);
        if (daysUntil > days) continue;
        const age = next.getFullYear() - bDay.getFullYear();
        result.push({ contact: c, daysUntil, nextBirthday: next, age });
      }
      return result.sort((a, b) => a.daysUntil - b.daysUntil);
    },

    birthdayLabel(daysUntil, nextBirthday) {
      if (daysUntil === 0) return "Heute";
      if (daysUntil === 1) return "Morgen";
      if (daysUntil < 7)  return `In ${daysUntil} Tagen`;
      return nextBirthday.toLocaleDateString("de-CH", { day: "2-digit", month: "2-digit" });
    },

    // Aktivitäts-Signal (nur Kunden): "" | "overdue" | "never" | "cold" | "ok"
    // GATE: nur kategorie === "Kunde"; Lieferant/Übrige/leer -> "" (kein Dot).
    //       VIP ist ein separates Flag und beeinflusst den Gate NICHT.
    // "overdue" — >=1 offene, überfällige Task
    // "never"   — noch kein History-Eintrag (für ALLE Kunden)
    // "cold"    — letzte Aktivität > 12 Monate (exakte Monatsdifferenz)
    // "ok"      — Kunde on track (keine der obigen Bedingungen)
    firmSignal(firm) {
      if (firm.kategorie !== "Kunde") return "";
      if (firm.openTasksCount > 0 && firm.tasks.some(t => t.isOpen && t.isOverdue)) return "overdue";
      if (firm.history.length === 0) return "never";
      const last = helpers.toDate(firm.latestActivity);
      if (!last) return "never";
      const today = helpers.todayStart();
      const months = (today.getFullYear() - last.getFullYear()) * 12 + (today.getMonth() - last.getMonth());
      if (months > 12) return "cold";
      return "ok";
    },

    // Cockpit-Linse: Firma gehört zur Linse, wenn einer ihrer Kontakte das
    // gewählte Lead-BBZ trägt. Leere Linse ("") matcht alle Firmen.
    firmMatchesLens(firm, lens) {
      return !lens || (firm.contacts || []).some(c => c.leadbbz0 === lens);
    },

    // Mini-Sparkline aus einem Zahlen-Array (nur numerisch → kein Escaping nötig).
    sparklineHtml(values, opts = {}) {
      if (!values || !values.length) return "";
      const w = opts.width || 88, h = opts.height || 22, stroke = opts.stroke || "var(--blue)";
      const max = Math.max(1, ...values);
      const step = values.length > 1 ? w / (values.length - 1) : w;
      const pts = values.map((v, i) => {
        const x = Math.round(i * step);
        const y = Math.round(h - 1 - (v / max) * (h - 2));
        return `${x},${y}`;
      }).join(" ");
      return `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" style="display:block;">`
        + `<polyline points="${pts}" fill="none" stroke="${stroke}" stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round" /></svg>`;
    },

    // Momentum-Heatmap: je Woche ein Kästchen, Intensität aus Aktivitäts-Count.
    momentumHeatmapHtml(weekCounts) {
      if (!weekCounts || !weekCounts.length) return "";
      const max = Math.max(1, ...weekCounts);
      const boxes = weekCounts.map(c => {
        const op = c === 0 ? 0.08 : (0.25 + 0.75 * (c / max)).toFixed(2);
        return `<span title="${c} Aktivität${c !== 1 ? "en" : ""}" style="flex:1;height:14px;border-radius:2px;background:var(--blue);opacity:${op};"></span>`;
      }).join("");
      return `<div style="display:flex;gap:3px;align-items:stretch;">${boxes}</div>`;
    },

    // Perioden-Key einer Aktivität je Granularität ("2026-03" / "2026-Q1" / "2026").
    periodKey(value, gran) {
      const dt = value instanceof Date ? value : helpers.toDate(value);
      if (!dt) return "";
      const y = dt.getFullYear();
      if (gran === "jahr")    return `${y}`;
      if (gran === "quartal") return `${y}-Q${Math.floor(dt.getMonth() / 3) + 1}`;
      return `${y}-${String(dt.getMonth() + 1).padStart(2, "0")}`; // monat (Default)
    },

    // Kompaktes Label eines Perioden-Keys.
    periodLabel(key) {
      if (!key) return "alle";
      const q = key.indexOf("-Q");
      if (q !== -1) return "Q" + key.slice(q + 2);
      const m = key.match(/^(\d{4})-(\d{2})$/);
      if (m) return ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"][parseInt(m[2], 10) - 1];
      return key; // Jahr
    },

    // Debounce: verhindert excessive DOM-Rebuilds beim Tippen in Suchfeldern
    debounce(fn, ms = 150) {
      let timer = null;
      return (...args) => {
        clearTimeout(timer);
        timer = setTimeout(() => fn(...args), ms);
      };
    },

    // Rendert ein <select> aus SP-Choices — fällt auf <input> zurück wenn keine Choices geladen
    choiceSelectHtml(name, listTitle, spFieldName, currentValue, required = false) {
      const choices = state.meta.choices?.[listTitle]?.[spFieldName] || [];
      if (!choices.length) {
        // Fallback: Freitext — tritt auf wenn Choices noch nicht geladen oder SP-Feld kein Choice
        return `<input class="bbz-input" name="${name}" value="${helpers.escapeHtml(currentValue || "")}" ${required ? "required" : ""} placeholder="Wird geladen..." />`;
      }
      return `
        <select class="bbz-select" name="${name}" ${required ? "required" : ""}>
          <option value="">— bitte wählen —</option>
          ${choices.map(c => `<option value="${helpers.escapeHtml(c)}" ${currentValue === c ? "selected" : ""}>${helpers.escapeHtml(c)}</option>`).join("")}
        </select>
      `;
    },

    // Rendert Checkboxen für Multi-Choice-Felder aus SP
    // currentValues: string[] der aktuell gesetzten Werte
    choiceMultiHtml(name, listTitle, spFieldName, currentValues) {
      const choices = state.meta.choices?.[listTitle]?.[spFieldName] || [];
      const selected = new Set(Array.isArray(currentValues) ? currentValues : []);
      if (!choices.length) {
        return `<input class="bbz-input" name="${name}" value="${helpers.escapeHtml([...selected].join(", "))}" placeholder="Wird geladen..." />`;
      }
      return `
        <div class="bbz-multi-choice">
          ${choices.map(c => `
            <label class="bbz-multi-choice-item">
              <input type="checkbox" name="${name}" value="${helpers.escapeHtml(c)}" ${selected.has(c) ? "checked" : ""} />
              <span>${helpers.escapeHtml(c)}</span>
            </label>
          `).join("")}
        </div>
      `;
    },

    ensureMsalAvailable() {
      if (!window.msal || !window.msal.PublicClientApplication) {
        throw new Error("MSAL-Bibliothek wurde nicht geladen.");
      }
    },

    validateConfig() {
      const missing = [];
      if (!CONFIG.graph.clientId) missing.push("clientId");
      if (!CONFIG.graph.tenantId) missing.push("tenantId");
      if (!CONFIG.graph.authority) missing.push("authority");
      if (!CONFIG.graph.redirectUri) missing.push("redirectUri");
      if (!CONFIG.sharePoint.siteHostname) missing.push("sharePoint.siteHostname");
      if (!CONFIG.sharePoint.sitePath) missing.push("sharePoint.sitePath");
      if (missing.length) throw new Error(`Konfiguration unvollstaendig: ${missing.join(", ")}`);
    }
  };

  const ui = {
    els: {
      viewRoot: null,
      authStatus: null,
      globalMessage: null,
      btnLogin: null,
      btnRefresh: null,
      navButtons: []
    },

    init() {
      this.els.viewRoot = document.getElementById("view-root");
      this.els.authStatus = document.getElementById("auth-status");
      this.els.globalMessage = document.getElementById("global-message");
      this.els.btnLogin = document.getElementById("btn-login");
      this.els.btnRefresh = document.getElementById("btn-refresh");
      this.els.navButtons = [...document.querySelectorAll(".bbz-nav-btn")];

      if (this.els.btnLogin) this.els.btnLogin.addEventListener("click", () => controller.handleLogin());
      if (this.els.btnRefresh) this.els.btnRefresh.addEventListener("click", () => controller.handleRefresh());

      // Admin-Panel: Doppelklick auf den Auth-Status-Bereich (unsichtbar für normale User)
      if (this.els.authStatus) {
        this.els.authStatus.addEventListener("dblclick", () => {
          controller.navigate("admin");
        });
        this.els.authStatus.title = "";  // kein Tooltip-Hinweis
      }

      this.els.navButtons.forEach(btn => {
        btn.addEventListener("click", () => controller.navigate(btn.dataset.route));
      });

      // Zentraler Click-Handler
      document.addEventListener("click", (event) => {
        const openFirm = event.target.closest("[data-action='open-firm']");
        if (openFirm) { controller.openFirm(openFirm.dataset.id); return; }

        const navPlanning = event.target.closest("[data-action='navigate-planning']");
        if (navPlanning) { event.preventDefault(); controller.navigate("planning"); return; }

        const navPlanningFiltered = event.target.closest("[data-action='navigate-planning-filtered']");
        if (navPlanningFiltered) {
          event.preventDefault();
          state.filters.planning.faelligkeit = navPlanningFiltered.dataset.faelligkeit || "";
          controller.navigate("planning");
          return;
        }

        // KPI-Schnellfilter — setzt Filter und navigiert bei Bedarf
        const kpiFilter = event.target.closest("[data-action='kpi-filter']");
        if (kpiFilter) {
          const scope = kpiFilter.dataset.scope;
          const value = kpiFilter.dataset.value;
          if (scope === "firms-kategorie") {
            state.filters.firms.kategorie = value;
            // Stufe-2-Filter gelten nur für Kunden -> beim Wechsel weg zuruecksetzen
            if (value !== "Kunde") { state.filters.firms.klassifizierung = ""; state.filters.firms.vip = false; }
          } else if (scope === "firms-klassifizierung") {
            state.filters.firms.klassifizierung = value;
          } else if (scope === "firms-vip") {
            // VIP ist additiver Toggle — unabhängig von A/B/C
            state.filters.firms.vip = !state.filters.firms.vip;
          } else if (scope === "contacts-mode") {
            // direkt route setzen — NICHT controller.navigate() da das _kpiMode zurücksetzt
            const newMode = state.filters.contacts._kpiMode === value ? "all" : value;
            state.filters.contacts._kpiMode = newMode;
            state.filters.route = "contacts";
            state.selection.contactId = null;
            state.modal = null;
          } else if (scope === "history-kontaktart") {
            state.filters.history.kontaktart = state.filters.history.kontaktart === value ? "" : value;
          } else if (scope === "planning-faelligkeit") {
            state.filters.planning.faelligkeit = state.filters.planning.faelligkeit === value ? "" : value;
          } else if (scope === "planning-segment") {
            state.filters.planning.segment = state.filters.planning.segment === value ? "" : value;
          } else if (scope === "planning-leadbbz") {
            state.filters.planning.leadbbz = state.filters.planning.leadbbz === value ? "" : value;
          } else if (scope === "events-segment") {
            state.filters.events.segment = state.filters.events.segment === value ? "" : value;
          } else if (scope === "events-selected") {
            state.filters.events.selectedEvent = state.filters.events.selectedEvent === value ? "" : value;
          } else if (scope === "navigate") {
            controller.navigate(value);
            return;
          }
          controller.render();
          return;
        }

        const toggleLegende = event.target.closest("[data-action='toggle-firm-legende']");
        if (toggleLegende) { state.filters.firms.legendeOffen = !state.filters.firms.legendeOffen; controller.render(); return; }

        const openContact = event.target.closest("[data-action='open-contact']");
        if (openContact) { controller.openContact(openContact.dataset.id); return; }

        const backToFirms = event.target.closest("[data-action='back-to-firms']");
        if (backToFirms) { controller.navigate("firms"); return; }

        const backToContacts = event.target.closest("[data-action='back-to-contacts']");
        if (backToContacts) { controller.navigate("contacts"); return; }

        const openForm = event.target.closest("[data-action='open-contact-form']");
        if (openForm) {
          const itemId = openForm.dataset.itemId ? Number(openForm.dataset.itemId) : null;
          const firmId = openForm.dataset.firmId ? Number(openForm.dataset.firmId) : null;
          controller.openContactForm(itemId, firmId);
          return;
        }

        // FIX 2a: Modal schliessen via Button oder Backdrop-Klick
        const closeModal = event.target.closest("[data-close-modal]");
        if (closeModal) { controller.closeModal(); return; }

        const backdrop = event.target.closest(".bbz-modal-backdrop");
        if (backdrop && !event.target.closest(".bbz-modal")) { controller.closeModal(); return; }

        // Kontakt löschen
        const deleteContact = event.target.closest("[data-action='delete-contact']");
        if (deleteContact) {
          controller.handleDeleteContact(deleteContact.dataset.id, deleteContact.dataset.name);
          return;
        }

        // Firma bearbeiten
        const openFirmForm = event.target.closest("[data-action='open-firm-form']");
        if (openFirmForm) {
          controller.openFirmForm(openFirmForm.dataset.id);
          return;
        }

        // History-Formular öffnen
        const openHistoryForm = event.target.closest("[data-action='open-history-form']");
        if (openHistoryForm) {
          const contactId = openHistoryForm.dataset.contactId || null;
          const firmId    = openHistoryForm.dataset.firmId    || null;
          // Guard: Firma ohne Kontakte — kein Modal öffnen
          if (!contactId && firmId) {
            const firm = dataModel.getFirmById(Number(firmId));
            if (firm && firm.contacts.length === 0) {
              ui.setMessage(`"${firm.title}" hat noch keine Kontakte. Bitte zuerst einen Kontakt erfassen.`, "error");
              return;
            }
          }
          controller.openHistoryForm(contactId ? Number(contactId) : null, firmId ? Number(firmId) : null, null, openHistoryForm.dataset.typ || null);
          return;
        }

        // History-Eintrag bearbeiten
        const editHistory = event.target.closest("[data-action='edit-history']");
        if (editHistory) {
          controller.openHistoryForm(null, null, Number(editHistory.dataset.id));
          return;
        }

        // History-Eintrag löschen
        const deleteHistory = event.target.closest("[data-action='delete-history']");
        if (deleteHistory) {
          controller.handleDeleteHistory(deleteHistory.dataset.id, deleteHistory.dataset.title);
          return;
        }

        // Task-Formular öffnen
        const openTaskForm = event.target.closest("[data-action='open-task-form']");
        if (openTaskForm) {
          const contactId = openTaskForm.dataset.contactId || null;
          const firmId    = openTaskForm.dataset.firmId    || null;
          // Guard: Firma ohne Kontakte — kein Modal öffnen
          if (!contactId && firmId) {
            const firm = dataModel.getFirmById(Number(firmId));
            if (firm && firm.contacts.length === 0) {
              ui.setMessage(`"${firm.title}" hat noch keine Kontakte. Bitte zuerst einen Kontakt erfassen.`, "error");
              return;
            }
          }
          controller.openTaskForm(contactId ? Number(contactId) : null, firmId ? Number(firmId) : null, null);
          return;
        }

        // Task bearbeiten
        const editTask = event.target.closest("[data-action='edit-task']");
        if (editTask) {
          controller.openTaskForm(null, null, Number(editTask.dataset.id));
          return;
        }

        // Task löschen
        const deleteTask = event.target.closest("[data-action='delete-task']");
        if (deleteTask) {
          controller.handleDeleteTask(deleteTask.dataset.id, deleteTask.dataset.title);
          return;
        }

        // Spalten-Sortierung (Planung)
        const setSort = event.target.closest("[data-action='set-sort']");
        if (setSort) {
          const col = setSort.dataset.col;
          const scope = setSort.dataset.scope || "planning";
          const f = scope === "firms" ? state.filters.firms : scope === "contacts" ? state.filters.contacts : state.filters.planning;
          if (f.sortBy === col) {
            f.sortDir = f.sortDir === "asc" ? "desc" : "asc";
          } else {
            f.sortBy = col;
            f.sortDir = "asc";
          }
          controller.render();
          return;
        }

        // Firma löschen
        const deleteFirm = event.target.closest("[data-action='delete-firm']");
        if (deleteFirm) {
          if (Number(deleteFirm.dataset.contacts) > 0) {
            ui.setMessage("Diese Firma hat noch Kontakte und kann nicht gelöscht werden.", "error");
            return;
          }
          controller.handleDeleteFirm(deleteFirm.dataset.id, deleteFirm.dataset.name);
          return;
        }

        // History: Notizen aufklappen/zuklappen (kein Re-Render)
        const expandBtn = event.target.closest("[data-action='toggle-expand']");
        if (expandBtn) {
          const card = expandBtn.closest(".bbz-timeline-item");
          if (card) {
            card.classList.toggle("bbz-expanded");
            expandBtn.textContent = card.classList.contains("bbz-expanded") ? "weniger" : "mehr";
          }
          return;
        }

        // History: Firma-Filter setzen
        const histFirmaFilter = event.target.closest("[data-action='history-firma-filter']");
        if (histFirmaFilter) {
          const firmTitle = histFirmaFilter.dataset.firmTitle || "";
          // Toggle: nochmals klicken = Filter aufheben
          state.filters.history.search = state.filters.history.search === firmTitle ? "" : firmTitle;
          controller.render();
          return;
        }

        // Handlungszentrum: Task per 1-Klick als erledigt markieren
        const completeTask = event.target.closest("[data-action='complete-task']");
        if (completeTask) {
          const choices = state.meta.choices?.[CONFIG.lists.tasks]?.["Status"] || [];
          const doneStatus = choices.find(s => !helpers.isOpenTask(s));
          if (doneStatus) controller.handleTaskStatusChange(Number(completeTask.dataset.id), doneStatus);
          return;
        }

        // Firmen-Cockpit: Firma-Zeile auf-/zuklappen
        const toggleFirm = event.target.closest("[data-action='toggle-firm-expand']");
        if (toggleFirm) {
          const fid = Number(toggleFirm.dataset.firmId);
          const arr = state.filters.history.expandedFirms;
          const i = arr.indexOf(fid);
          if (i === -1) arr.push(fid); else arr.splice(i, 1);
          controller.render();
          return;
        }

        // Aktivitäten-Ansicht umschalten: Firmen-Bericht / Chronologisch
        const viewMode = event.target.closest("[data-action='history-view-mode']");
        if (viewMode) {
          state.filters.history.viewMode = viewMode.dataset.mode;
          controller.render();
          return;
        }

        // Aktivitäten-Balkenfilter: Lead BBZ (case-insensitiv togglen)
        const filterLens = event.target.closest("[data-action='filter-lens']");
        if (filterLens) {
          const v = filterLens.dataset.value || "";
          const cur = state.filters.history.lens || "";
          state.filters.history.lens = cur.toLowerCase() === v.toLowerCase() ? "" : v;
          controller.render();
          return;
        }

        // Aktivitäten-Balkenfilter: Periode (Toggle; leerer Wert = alle)
        const filterPeriode = event.target.closest("[data-action='filter-periode']");
        if (filterPeriode) {
          const v = filterPeriode.dataset.value || "";
          state.filters.history.periode = (v === "" || state.filters.history.periode === v) ? "" : v;
          controller.render();
          return;
        }

        // Aktivitäten-Balkenfilter: Granularität wechseln (Periode zurücksetzen)
        const filterGran = event.target.closest("[data-action='filter-granularitaet']");
        if (filterGran) {
          state.filters.history.granularitaet = filterGran.dataset.value || "monat";
          state.filters.history.periode = "";
          controller.render();
          return;
        }

        // Batch-Event-Dialog öffnen
        const openBatchEvent = event.target.closest("[data-action='open-batch-event']");
        if (openBatchEvent) {
          const eventName = openBatchEvent.dataset.eventName || "";
          const mode = openBatchEvent.dataset.mode || "anmelden";
          controller.openBatchEventForm(eventName, mode);
          return;
        }

        // Event Einladungsliste öffnen
        const openEventEinladung = event.target.closest("[data-action='open-event-einladung']");
        if (openEventEinladung) {
          state.modal = {
            type: "event-einladung",
            payload: {
              eventName: openEventEinladung.dataset.eventName || "",
              listLabel: openEventEinladung.dataset.listLabel || "Einladungsliste",
              filterSeg: "", filterSearch: ""
            }
          };
          controller.render();
          return;
        }

        // Event Nachbearbeitung öffnen
        const openEventNachbearbeitung = event.target.closest("[data-action='open-event-nachbearbeitung']");
        if (openEventNachbearbeitung) {
          const evName = openEventNachbearbeitung.dataset.eventName || "";
          // Ersten passenden Eventhistory-Choice vorselektieren
          const histChoices = state.meta.choices?.[CONFIG.lists.contacts]?.["Eventhistory"] || [];
          state.modal = {
            type: "event-nachbearbeitung",
            payload: {
              eventName: evName,
              checkedIds: [],
              selectedVersion: histChoices[0] || "",
              filterSearch: ""
            }
          };
          controller.render();
          return;
        }

        // Event Nachbearbeitung: Checkbox togglen
        const nbToggle = event.target.closest("[data-action='event-nb-toggle']");
        if (nbToggle && state.modal?.payload) {
          const cid = Number(nbToggle.dataset.contactId);
          const ids = state.modal.payload.checkedIds;
          const idx = ids.indexOf(cid);
          if (idx === -1) ids.push(cid); else ids.splice(idx, 1);
          nbToggle.classList.toggle("checked", ids.includes(cid));
          nbToggle.closest("tr").style.background = ids.includes(cid) ? "#f0fdf4" : "";
          // Save-Button + Stats live aktualisieren ohne full re-render
          const saveBtn = document.querySelector("[data-action='event-nb-save']");
          if (saveBtn) {
            saveBtn.textContent = `✓ Teilnahmen speichern (${ids.length})`;
            saveBtn.disabled = ids.length === 0 || !state.modal.payload.selectedVersion;
          }
          const markedStat = document.querySelector("[data-nb-marked]");
          if (markedStat) markedStat.textContent = ids.length;
          return;
        }

        // Event Nachbearbeitung: Speichern
        const nbSave = event.target.closest("[data-action='event-nb-save']");
        if (nbSave) {
          controller.handleEventNachbearbeitungSave();
          return;
        }

        // Event Einladungsliste: Kontakt entfernen
        const evRemove = event.target.closest("[data-action='event-remove-contact']");
        if (evRemove) {
          controller.handleEventRemoveContact(
            evRemove.dataset.eventName,
            Number(evRemove.dataset.contactId)
          );
          return;
        }

        // Event Stats-Bar Filter
        const evStatFilter = event.target.closest("[data-action='event-stat-filter']");
        if (evStatFilter && state.modal?.type === "event-einladung") {
          state.modal.payload.filterSeg = evStatFilter.dataset.seg || "";
          controller.render();
          return;
        }

        // === Event-Matrix: Modal öffnen ===
        const openEventMatrix = event.target.closest("[data-action='open-event-matrix']");
        if (openEventMatrix) {
          state.modal = {
            type: "event-matrix",
            payload: {
              filterSearch: "",
              filterFirmId: "",
              filterLeadbbz: "",
              filterSegment: "",
              sortBy: "firm",
              sortDir: "asc",
              pendingChanges: {}
            }
          };
          controller.render();
          return;
        }

        // === Event-Matrix: Filter zurücksetzen ===
        const matrixClearFilters = event.target.closest("[data-action='matrix-clear-filters']");
        if (matrixClearFilters && state.modal?.type === "event-matrix") {
          state.modal.payload.filterSearch = "";
          state.modal.payload.filterFirmId = "";
          state.modal.payload.filterLeadbbz = "";
          state.modal.payload.filterSegment = "";
          controller.render();
          return;
        }

        // === Event-Matrix: Spalten-Sortierung ===
        const matrixSort = event.target.closest("[data-action='matrix-sort']");
        if (matrixSort && state.modal?.type === "event-matrix") {
          const key = matrixSort.dataset.sortKey;
          const p = state.modal.payload;
          if (p.sortBy === key) {
            p.sortDir = p.sortDir === "asc" ? "desc" : "asc";
          } else {
            p.sortBy = key;
            p.sortDir = "asc";
          }
          controller.render();
          return;
        }

        // === Event-Matrix: einzelne Zelle togglen ===
        const matrixCellToggle = event.target.closest("[data-action='matrix-cell-toggle']");
        if (matrixCellToggle && state.modal?.type === "event-matrix") {
          const cid = Number(matrixCellToggle.dataset.contactId);
          const evName = matrixCellToggle.dataset.eventName;
          const newVal = matrixCellToggle.checked;
          const pc = state.modal.payload.pendingChanges;
          if (!pc[cid]) pc[cid] = {};
          // Wenn neuer Wert = Original, lösche Pending (nicht dirty)
          const contact = state.enriched.contacts.find(c => c.id === cid);
          const original = contact ? helpers.toArray(contact.event).includes(evName) : false;
          if (newVal === original) {
            delete pc[cid][evName];
            if (Object.keys(pc[cid]).length === 0) delete pc[cid];
          } else {
            pc[cid][evName] = newVal;
          }
          controller.render();
          return;
        }

        // === Event-Matrix: Spalten-Toggle (alle gefilterten Kontakte für eine Event-Spalte) ===
        const matrixColToggle = event.target.closest("[data-action='matrix-col-toggle']");
        if (matrixColToggle && state.modal?.type === "event-matrix") {
          const evName = matrixColToggle.dataset.eventName;
          const targetVal = matrixColToggle.checked;
          const payload = state.modal.payload;
          // Aktuell gefilterte Kontakte erneut bestimmen (DRY-Verstoss wäre hier teurer als Duplikation)
          const firmMap = new Map(state.enriched.firms.map(f => [f.id, f]));
          let rows = state.enriched.contacts.filter(c => !c.archiviert);
          if (payload.filterFirmId) rows = rows.filter(c => String(c.firmId) === String(payload.filterFirmId));
          if (payload.filterLeadbbz) rows = rows.filter(c => c.leadbbz0 === payload.filterLeadbbz);
          if (payload.filterSegment) rows = rows.filter(c => String(firmMap.get(c.firmId)?.klassifizierung || "").toUpperCase().startsWith(payload.filterSegment));
          if (payload.filterSearch.trim()) {
            const s = payload.filterSearch.trim().toLowerCase();
            rows = rows.filter(c => [c.fullName, c.firmTitle].some(v => helpers.textIncludes(v, s)));
          }
          const pc = payload.pendingChanges;
          rows.forEach(c => {
            const original = helpers.toArray(c.event).includes(evName);
            if (targetVal === original) {
              if (pc[c.id]) {
                delete pc[c.id][evName];
                if (Object.keys(pc[c.id]).length === 0) delete pc[c.id];
              }
            } else {
              if (!pc[c.id]) pc[c.id] = {};
              pc[c.id][evName] = targetVal;
            }
          });
          controller.render();
          return;
        }

        // === Event-Matrix: Speichern ===
        const matrixSave = event.target.closest("[data-action='matrix-save']");
        if (matrixSave && state.modal?.type === "event-matrix") {
          controller.handleEventMatrixSave();
          return;
        }

        // === Event-Matrix: Änderungen verwerfen ===
        const matrixDiscard = event.target.closest("[data-action='matrix-discard']");
        if (matrixDiscard && state.modal?.type === "event-matrix") {
          if (confirm("Alle ausstehenden Änderungen verwerfen?")) {
            state.modal.payload.pendingChanges = {};
            controller.render();
          }
          return;
        }

        // Event Excel-Export
        const evExport = event.target.closest("[data-action='event-export-excel']");
        if (evExport) {
          controller.handleEventExcelExport(evExport.dataset.eventName);
          return;
        }

        // Admin-Panel: Zeitfilter umschalten
        const adminZf = event.target.closest("[data-action='admin-zeitfilter']");
        if (adminZf) {
          if (!state.filters.admin) state.filters.admin = { zeitfenster: "30" };
          state.filters.admin.zeitfenster = adminZf.dataset.zf || "30";
          controller.render();
          return;
        }

        // Hilfsfunktion: Zähler, Submit-Button und Alle-Checkbox synchronisieren
        const syncBatchUI = (sel) => {
          const payload = state.modal?.payload;
          if (!payload) return;
          const preview = payload.previewContacts || [];
          const max = preview.length;
          const form = document.querySelector("[data-modal-form='batch-event']");
          const isEH = form?.dataset.mode === "eventhistory";
          const cat  = form?.dataset.eventName || "";

          const submitBtn = form?.querySelector("button[type='submit']");
          if (submitBtn) {
            submitBtn.textContent = isEH
              ? `+ ${sel.length} × Eventhistory «${cat}» setzen`
              : `+ ${sel.length} × Event «${cat}» setzen`;
            submitBtn.disabled = sel.length === 0;
          }
          const counter = document.querySelector("[data-batch-counter]");
          if (counter) counter.textContent = `${sel.length} von ${max} ausgewählt${max >= 200 ? " (max. 200 — Filter verfeinern)" : ""}`;
          const allChecked = max > 0 && preview.every(c => sel.includes(c.id));
          const allCb = document.querySelector("input[data-action='batch-toggle-all']");
          if (allCb) allCb.checked = allChecked;
        };

        // Batch-Event-Auswahl: einzelne Checkbox
        const batchToggle = event.target.closest("[data-action='batch-toggle-contact']");
        if (batchToggle) {
          const cid = Number(batchToggle.dataset.contactId);
          if (!state.modal?.payload?.selected) return;
          const sel = state.modal.payload.selected;
          const idx = sel.indexOf(cid);
          if (idx === -1) sel.push(cid); else sel.splice(idx, 1);
          // data-action ist direkt auf dem <input> — batchToggle IST die Checkbox
          batchToggle.checked = sel.includes(cid);
          batchToggle.closest("tr")?.classList.toggle("bbz-row-ok", sel.includes(cid));
          syncBatchUI(sel);
          return;
        }

        // Batch-Event: Alle/Keine togglen
        const batchToggleAll = event.target.closest("[data-action='batch-toggle-all']");
        if (batchToggleAll && state.modal?.payload) {
          const preview = state.modal.payload.previewContacts || [];
          const allIds = preview.map(c => c.id);
          const allSelected = allIds.every(id => state.modal.payload.selected.includes(id));
          state.modal.payload.selected = allSelected ? [] : [...allIds];
          const newSel = state.modal.payload.selected;
          document.querySelectorAll("input[data-action='batch-toggle-contact']").forEach(cb => {
            cb.checked = newSel.includes(Number(cb.dataset.contactId));
            cb.closest("tr")?.classList.toggle("bbz-row-ok", cb.checked);
          });
          syncBatchUI(newSel);
          return;
        }

        // KEIN separater Handler für [data-modal-submit] nötig:
        // Der Button hat type="submit" und löst den nativen Form-Submit aus,
        // der vom submit-Listener unten abgefangen wird.
        // Ein zusätzlicher dispatchEvent hier würde double-submit verursachen.
      });

      // isPrivat-Label: dynamisch aktualisieren wenn Firma im Kontaktformular wechselt
      document.addEventListener("change", (event) => {
        const firmSelect = event.target.closest("[data-modal-form='contact'] select[name='firmaLookupId']");
        if (firmSelect && state.meta.privateFirmId !== null) {
          const isPrivat = String(firmSelect.value) === String(state.meta.privateFirmId);
          const label = firmSelect.closest(".bbz-modal")?.querySelector("label[data-kommentar-label]");
          if (label) label.textContent = isPrivat ? "Adresse / Notizen (Privatperson — Adresse hier erfassen)" : "Kommentar";
        }
      }, true); // capture: true — vor dem bbz change-listener feuern

      // FIX 2c: Zentraler Form-Submit-Handler — Guard gegen Double-Submit
      document.addEventListener("submit", (event) => {
        const form = event.target.closest("[data-modal-form]");
        if (form) {
          event.preventDefault();
          if (state.meta.loading) return;
          const formType = form.dataset.modalForm;
          if (formType === "firm") {
            controller.handleFirmModalSubmit(form, form.dataset.mode, form.dataset.itemId || null);
          } else if (formType === "history") {
            controller.handleHistoryModalSubmit(form);
          } else if (formType === "task") {
            controller.handleTaskModalSubmit(form);
          } else if (formType === "batch-event") {
            controller.handleBatchEventSubmit(form);
          } else {
            controller.handleModalSubmit(form, form.dataset.mode, form.dataset.itemId || null);
          }
        }
      });

      const debouncedRender = helpers.debounce(() => controller.render(), 150);

      // Browser-Back/Forward — State aus history.state wiederherstellen
      window.addEventListener("popstate", (event) => {
        const s = event.state;
        if (!s) {
          // Kein State (z.B. erster Eintrag) — zur Startseite
          state.filters.route = CONFIG.defaults.route;
          state.selection.firmId = null;
          state.selection.contactId = null;
        } else {
          state.filters.route = s.route || CONFIG.defaults.route;
          state.selection.firmId = s.firmId || null;
          state.selection.contactId = s.contactId || null;
        }
        state.modal = null;
        window.scrollTo(0, 0);
        controller.render();
      });

      // Initialen State setzen damit der erste Back-Schritt korrekt funktioniert
      // WICHTIG: Nicht ausführen wenn MSAL gerade einen Auth-Redirect verarbeitet
      // (Hash enthält "code=" oder "error=") — sonst überschreibt replaceState den Auth-Hash
      // und handleRedirectPromise() findet keinen gültigen Hash mehr
      // MSAL v3 PKCE: Auth-Code kommt als Query-Parameter (?code=) ODER im Hash
      // Guard: replaceState nicht ausführen wenn MSAL gerade einen Redirect verarbeitet
      const currentHash = window.location.hash;
      const currentSearch = window.location.search;
      const isMsalRedirect = currentHash.includes("code=") || currentHash.includes("error=") || currentHash.includes("state=")
                          || currentSearch.includes("code=") || currentSearch.includes("error=") || currentSearch.includes("state=");
      // Bekannte App-Routen aus dem Hash lesen — verhindert Ueberschreiben von #admin etc.
      const knownRoutes = ["firms","contacts","planning","history","events","birthdays","admin"];
      const hashRoute = currentHash.replace("#", "").split("-")[0];
      if (!isMsalRedirect && knownRoutes.includes(hashRoute)) {
        state.filters.route = hashRoute;
        history.replaceState({ route: hashRoute, firmId: null, contactId: null }, "", currentHash);
      } else if (!isMsalRedirect) {
        history.replaceState(
          { route: state.filters.route, firmId: null, contactId: null },
          "",
          `#${state.filters.route}`
        );
      }

      document.addEventListener("input", (event) => {
        const el = event.target;
        if (el.matches("[data-filter='firms-search']")) { state.filters.firms.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='contacts-search']")) { state.filters.contacts.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='planning-search']")) { state.filters.planning.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='history-search']")) { state.filters.history.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='events-search']")) { state.filters.events.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='batch-search']") && state.modal?.payload) { state.modal.payload.filterSearch = el.value; state.modal.payload.selected = []; debouncedRender(); }
        if (el.matches("[data-filter='event-einladung-search']") && state.modal?.payload) { state.modal.payload.filterSearch = el.value; debouncedRender(); }
        if (el.matches("[data-filter='event-nb-search']") && state.modal?.payload) { state.modal.payload.filterSearch = el.value; debouncedRender(); }
        if (el.matches("[data-filter='batch-eventhistory-category-text']") && state.modal?.payload) { state.modal.payload.selectedHistoryCategory = el.value; state.modal.payload.selected = []; debouncedRender(); }
        if (el.matches("[data-filter='matrix-search']") && state.modal?.type === "event-matrix") { state.modal.payload.filterSearch = el.value; debouncedRender(); }
      });

      document.addEventListener("change", (event) => {
        const el = event.target;
        if (el.matches("[data-filter='firms-sortdir']")) { state.filters.firms.sortDir = el.value; controller.render(); }
        if (el.matches("[data-filter='contacts-archiviert']")) { state.filters.contacts.archiviertAusblenden = el.checked; controller.render(); }
        if (el.matches("[data-filter='planning-open']")) { state.filters.planning.onlyOpen = el.checked; controller.render(); }
        if (el.matches("[data-filter='planning-groupby']")) { state.filters.planning.groupBy = el.value; controller.render(); }
        if (el.matches("[data-filter='planning-sortdir']")) { state.filters.planning.sortDir = el.value; controller.render(); }
        if (el.matches("[data-filter='history-kontaktart']")) { state.filters.history.kontaktart = el.value; controller.render(); }
        if (el.matches("[data-filter='events-open']")) { state.filters.events.onlyWithOpenTasks = el.checked; controller.render(); }
        if (el.matches("[data-filter='events-sortby']")) { state.filters.events.sortBy = el.value; controller.render(); }
        // Batch-Event-Modal Filter
        if (el.matches("[data-filter='batch-segment']") && state.modal?.payload) { state.modal.payload.filterSegment = el.value; state.modal.payload.selected = []; controller.render(); }
        if (el.matches("[data-filter='batch-leadbbz']") && state.modal?.payload) { state.modal.payload.filterLeadbbz = el.value; state.modal.payload.selected = []; controller.render(); }
        if (el.matches("[data-filter='batch-sgf']") && state.modal?.payload) { state.modal.payload.filterSgf = el.value; state.modal.payload.selected = []; controller.render(); }
        if (el.matches("[data-filter='batch-eventhistory-category']") && state.modal?.payload) { state.modal.payload.selectedHistoryCategory = el.value; state.modal.payload.selected = []; controller.render(); }
        // Event Einladungs-Modal Filter
        if (el.matches("[data-filter='event-einladung-seg']") && state.modal?.payload) { state.modal.payload.filterSeg = el.value; controller.render(); }
        // Event Nachbearbeitungs-Modal: Version wählen
        if (el.matches("[data-filter='event-nb-version']") && state.modal?.payload) {
          state.modal.payload.selectedVersion = el.value;
          const saveBtn = document.querySelector("[data-action='event-nb-save']");
          if (saveBtn) saveBtn.disabled = state.modal.payload.checkedIds.length === 0;
        }
        // Event-Matrix-Modal Filter
        if (el.matches("[data-filter='matrix-firm']") && state.modal?.type === "event-matrix") { state.modal.payload.filterFirmId = el.value; controller.render(); }
        if (el.matches("[data-filter='matrix-leadbbz']") && state.modal?.type === "event-matrix") { state.modal.payload.filterLeadbbz = el.value; controller.render(); }
        if (el.matches("[data-filter='matrix-segment']") && state.modal?.type === "event-matrix") { state.modal.payload.filterSegment = el.value; controller.render(); }
        if (el.matches("[data-action='task-status-change']")) {
          controller.handleTaskStatusChange(Number(el.dataset.taskId), el.value);
        }
      });
    },

    setLoading(isLoading) {
      state.meta.loading = isLoading;
      this.renderShell();
    },

    setMessage(message, type = "info") {
      const el = this.els.globalMessage;
      if (!el) return;
      if (!message) { el.className = "bbz-banner"; el.textContent = ""; return; }
      const cls = { success: "bbz-banner bbz-banner-success show", warning: "bbz-banner bbz-banner-warning show", error: "bbz-banner bbz-banner-error show", info: "bbz-banner bbz-banner-info show" };
      el.className = cls[type] || cls.info;
      el.textContent = message;
    },

    renderShell() {
      // Desktop nav active state
      this.els.navButtons.forEach(btn => {
        btn.classList.toggle("active", btn.dataset.route === state.filters.route);
      });

      // Mobile bottom nav active state — direkt synchronisieren, kein MutationObserver nötig
      document.querySelectorAll(".bbz-bottom-btn").forEach(btn => {
        btn.classList.toggle("active", btn.dataset.route === state.filters.route);
      });

      if (state.auth.isAuthenticated && state.auth.account) {
        this.els.authStatus.innerHTML = (() => {
          const acc = state.auth.account;
          // MSAL v3: username = UPN/E-Mail, idTokenClaims.preferred_username als Fallback
          const email = acc.username
            || acc.idTokenClaims?.preferred_username
            || acc.idTokenClaims?.email
            || acc.idTokenClaims?.upn
            || acc.name
            || "";
          return `<span class="bbz-auth-dot"></span><span>Angemeldet: ${helpers.escapeHtml(email)}</span>`;
        })();
      } else if (state.auth.isReady) {
        this.els.authStatus.innerHTML = `<span class="bbz-auth-dot" style="background:#94a3b8;"></span><span>Nicht angemeldet</span>`;
      } else {
        this.els.authStatus.innerHTML = `<span class="bbz-auth-dot" style="background:#f59e0b;"></span><span>Authentifizierung wird initialisiert ...</span>`;
      }

      if (this.els.btnLogin) {
        this.els.btnLogin.textContent = state.auth.isAuthenticated ? "Angemeldet" : "Anmelden";
        this.els.btnLogin.disabled = state.meta.loading || !state.auth.isReady;
      }
      if (this.els.btnRefresh) {
        this.els.btnRefresh.disabled = state.meta.loading || !state.auth.isReady;
      }
    },

    renderView(html) {
      if (!this.els.viewRoot) return;
      // Fokus + Cursor-Position bei Suchfeldern vor dem Re-Render merken
      const active = document.activeElement;
      const isSearchInput = active && active.matches("[data-filter$='-search']");
      const savedFilter = isSearchInput ? active.dataset.filter : null;
      const savedStart  = isSearchInput ? active.selectionStart : null;
      const savedEnd    = isSearchInput ? active.selectionEnd   : null;

      // Scroll-Position von Modal-Body merken (falls offenes Modal mit scrollbarem Body)
      const modalBody = document.querySelector(".bbz-modal-backdrop.show .bbz-modal-body");
      const savedScrollTop  = modalBody ? modalBody.scrollTop  : null;
      const savedScrollLeft = modalBody ? modalBody.scrollLeft : null;

      this.els.viewRoot.innerHTML = html;

      // Fokus + Cursor wiederherstellen
      if (savedFilter) {
        const restored = this.els.viewRoot.querySelector(`[data-filter="${savedFilter}"]`);
        if (restored) {
          restored.focus();
          try { restored.setSelectionRange(savedStart, savedEnd); } catch (_) {}
        }
      }

      // Scroll-Position wiederherstellen
      if (savedScrollTop !== null) {
        const newModalBody = document.querySelector(".bbz-modal-backdrop.show .bbz-modal-body");
        if (newModalBody) {
          newModalBody.scrollTop  = savedScrollTop;
          newModalBody.scrollLeft = savedScrollLeft;
        }
      }
    },

    loadingBlock(text = "Daten werden geladen ...") {
      return `<section class="bbz-section"><div class="bbz-section-body"><div style="display:flex;align-items:center;gap:10px;"><div class="bbz-loader"></div><div style="font-size:13px;color:var(--muted);">${helpers.escapeHtml(text)}</div></div></div></section>`;
    },

    emptyBlock(text = "Keine Daten vorhanden.", action = null, actionLabel = null) {
      if (action && actionLabel) {
        return `<div class="bbz-empty">${helpers.escapeHtml(text)}<br><button class="bbz-button bbz-button-secondary" style="margin-top:10px;height:32px;font-size:13px;" data-action="${helpers.escapeHtml(action)}">${helpers.escapeHtml(actionLabel)}</button></div>`;
      }
      return `<div class="bbz-empty">${helpers.escapeHtml(text)}</div>`;
    },

    kv(label, value) {
      return `<div class="bbz-kv"><div class="bbz-kv-label">${helpers.escapeHtml(label)}</div><div class="bbz-kv-value">${value || '<span class="bbz-muted">—</span>'}</div></div>`;
    },

    // Wrapper für KV-Gruppen — gibt eine section mit kompakten Rows zurück
    kvSection(title, rows) {
      return `<section class="bbz-section"><div class="bbz-section-header"><div class="bbz-section-title">${helpers.escapeHtml(title)}</div></div><div class="bbz-section-body">${rows.join("")}</div></section>`;
    }
  };

  const api = {
    async initAuth() {
      helpers.ensureMsalAvailable();
      helpers.validateConfig();

      state.auth.isReady = false;
      state.auth.msal = null;

      // MSAL v3: PublicClientApplication mit PKCE — löst Hash-Problem auf Mobile
      const msalInstance = new window.msal.PublicClientApplication({
        auth: {
          clientId: CONFIG.graph.clientId,
          authority: CONFIG.graph.authority,
          redirectUri: CONFIG.graph.redirectUri
        },
        cache: {
          cacheLocation: "localStorage",
          storeAuthStateInCookie: true
        },
        system: {
          allowNativeBroker: false
        }
      });

      // MSAL v3: initialize() verarbeitet Redirect-Response automatisch
      await msalInstance.initialize();
      state.auth.msal = msalInstance;

      // MSAL v3: handleRedirectPromise() nach initialize() aufrufen
      try {
        const redirectResponse = await state.auth.msal.handleRedirectPromise();
        if (redirectResponse?.account) {
          state.auth.account = redirectResponse.account;
          state.auth.isAuthenticated = true;
          // URL nach Redirect bereinigen
          history.replaceState(
            { route: CONFIG.defaults.route, firmId: null, contactId: null },
            "",
            `#${CONFIG.defaults.route}`
          );
        }
      } catch (error) {
        console.warn("handleRedirectPromise Fehler:", error);
        state.meta.lastError = error;
      }

      // Account aus Cache laden
      if (!state.auth.account) {
        const accounts = state.auth.msal.getAllAccounts();
        if (accounts.length > 0) {
          state.auth.account = accounts.find(a => a.tenantId === CONFIG.graph.tenantId) || accounts[0];
          state.auth.isAuthenticated = true;
        }
      }

      state.auth.isReady = true;
    },

    // Zentrale Interaktions-Funktion: versucht Popup, fällt bei popup_window_error
    // automatisch auf Redirect zurück (Popups durch Browser/Policy geblockt).
    // Gibt null zurück bei Redirect — Browser navigiert weg, kein Code läuft weiter.
    async msalInteract(request) {
      const isMobile = window.innerWidth < 768 || /Mobi|Android/i.test(navigator.userAgent);
      if (isMobile) {
        await state.auth.msal.loginRedirect(request);
        return null;
      }
      try {
        return await state.auth.msal.loginPopup(request);
      } catch (popupErr) {
        if (popupErr.errorCode === "popup_window_error" || (popupErr.message || "").includes("popup_window")) {
          console.warn("Popup geblockt — Redirect-Fallback.");
          await state.auth.msal.loginRedirect(request);
          return null;
        }
        throw popupErr;
      }
    },

    async login() {
      if (!state.auth.msal) throw new Error("MSAL ist nicht initialisiert.");
      const loginResponse = await this.msalInteract({
        scopes: CONFIG.graph.scopes,
        prompt: "select_account"
      });
      if (loginResponse === null) return; // Redirect — Browser navigiert weg

      if (!loginResponse?.account) throw new Error("Keine Kontoinformation aus dem Login erhalten.");
      state.auth.account = loginResponse.account;
      state.auth.isAuthenticated = true;
      await this.acquireToken();
    },

    // FIX 3b: robusteres Token-Handling mit Account-Fallback
    async acquireToken() {
      if (!state.auth.msal) throw new Error("MSAL ist nicht initialisiert.");

      // Account aus Cache nachladen falls leer
      if (!state.auth.account) {
        const accounts = state.auth.msal.getAllAccounts();
        if (accounts.length > 0) {
          // Tenant-Match bevorzugen — verhindert falschen Account bei Multi-Tenant-Umgebungen
          state.auth.account = accounts.find(a => a.tenantId === CONFIG.graph.tenantId) || accounts[0];
          state.auth.isAuthenticated = true;
        } else {
          throw new Error("Kein angemeldetes Konto gefunden.");
        }
      }

      try {
        const tokenResponse = await state.auth.msal.acquireTokenSilent({
          account: state.auth.account,
          scopes: CONFIG.graph.scopes,
          forceRefresh: false
        });
        if (!tokenResponse?.accessToken) throw new Error("Kein Token aus acquireTokenSilent erhalten.");
        state.auth.token = tokenResponse.accessToken;
        return state.auth.token;
      } catch (silentError) {
        console.warn("Silent token fehlgeschlagen:", silentError);
        const isMobile = window.innerWidth < 768 || /Mobi|Android/i.test(navigator.userAgent);
        if (isMobile) {
          // Mobile: Token-Refresh via Redirect
          await state.auth.msal.acquireTokenRedirect({
            account: state.auth.account,
            scopes: CONFIG.graph.scopes
          });
          return state.auth.token;
        }
        // Desktop: Token-Refresh via Popup (mit Redirect-Fallback bei geblockten Popups)
        const interactResponse = await this.msalInteract({
          account: state.auth.account,
          scopes: CONFIG.graph.scopes
        });
        if (interactResponse === null) return state.auth.token; // Redirect — Browser navigiert weg
        if (!interactResponse?.accessToken) throw new Error("Kein Token aus acquireTokenPopup erhalten.");
        state.auth.token = interactResponse.accessToken;
        return state.auth.token;
      }
    },

    async graphRequest(path, options = {}) {
      // Token immer frisch via acquireToken — nicht auf gecachten state.auth.token verlassen
      const token = await this.acquireToken();
      const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
        method: options.method || "GET",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(options.headers || {}) },
        body: options.body ? JSON.stringify(options.body) : undefined
      });

      if (!response.ok) {
        let detail = "";
        try {
          detail = await response.text();
          console.error(`Graph ${response.status} auf ${options.method || "GET"} ${path}:`, detail);
        } catch { detail = response.statusText; }
        throw new Error(`Graph ${response.status}: ${detail}`);
      }

      if (response.status === 204) return null;
      return await response.json();
    },

    // Prüft ob Consent für Sites.ReadWrite.All bereits erteilt wurde.
    // Macht einen einzelnen, sequenziellen Probe-Call auf /sites/{ref} —
    // bevor Promise.all() 4 parallele Calls startet die alle gleichzeitig
    // mit 403 scheitern und sich gegenseitig mit interaction_in_progress blockieren.
    async ensureConsent() {
      const siteRef = `${CONFIG.sharePoint.siteHostname}:${CONFIG.sharePoint.sitePath}`;
      const token = await this.acquireToken();
      const probe = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteRef}`, {
        headers: { Authorization: `Bearer ${token}` }
      });

      if (probe.ok) {
        // Consent vorhanden — SiteId gleich cachen
        const data = await probe.json();
        state.meta.siteId = data.id;
        return;
      }

      let detail = "";
      try { detail = await probe.text(); } catch { detail = probe.statusText; }

      if (probe.status === 403 && detail.includes("accessDenied")) {
        console.warn("Consent fehlt für Sites.ReadWrite.All — starte Consent-Flow.");
        const isMobile = window.innerWidth < 768 || /Mobi|Android/i.test(navigator.userAgent);
        if (isMobile) {
          await state.auth.msal.loginRedirect({
            account: state.auth.account,
            scopes: CONFIG.graph.scopes,
            prompt: "consent"
          });
          return; // Browser navigiert weg
        } else {
          const consentResponse = await this.msalInteract({
            account: state.auth.account,
            scopes: CONFIG.graph.scopes,
            prompt: "consent"
          });
          if (consentResponse === null) return; // Redirect — Browser navigiert weg
          if (consentResponse?.account) {
            state.auth.account = consentResponse.account;
            state.auth.token = consentResponse.accessToken;
          }
          // Zweiter Probe nach Consent — wenn immer noch 403: echter Berechtigungsfehler
          const probe2 = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteRef}`, {
            headers: { Authorization: `Bearer ${state.auth.token}` }
          });
          if (!probe2.ok) throw new Error("Zugriff auf SharePoint auch nach Consent verweigert. Bitte Administrator kontaktieren.");
          const data2 = await probe2.json();
          state.meta.siteId = data2.id;
          return;
        }
      }

      // Anderer Fehler (kein Consent-Problem)
      throw new Error(`Graph ${probe.status}: ${detail}`);
    },

    async getSiteId() {
      if (state.meta.siteId) return state.meta.siteId;
      const siteRef = `${CONFIG.sharePoint.siteHostname}:${CONFIG.sharePoint.sitePath}`;
      const data = await this.graphRequest(`/sites/${siteRef}`);
      state.meta.siteId = data.id;
      return state.meta.siteId;
    },

    async getListItems(listTitle) {
      const siteId = await this.getSiteId();
      // Wichtig: expand=fields UND fields=createdBy,lastModifiedBy,createdDateTime,lastModifiedDateTime
      // liefert SP-Metadaten auf Item-Ebene (nicht in fields{}) — nötig für Admin-Auswertungen
      const url = `/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/items`
        + `?expand=fields`
        + `&$select=id,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,fields`
        + `&top=5000`;
      const data = await this.graphRequest(url);
      return data.value || [];
    },

    async loadAll() {
      if (!state.auth.isAuthenticated) throw new Error("Nicht angemeldet — loadAll() abgebrochen.");
      const [firms, contacts, history, tasks] = await Promise.all([
        this.getListItems(SCHEMA.firms.listTitle),
        this.getListItems(SCHEMA.contacts.listTitle),
        this.getListItems(SCHEMA.history.listTitle),
        this.getListItems(SCHEMA.tasks.listTitle)
      ]);

      state.data.firms = firms.map(item => normalizer.firm(item));
      state.data.contacts = contacts.map(item => normalizer.contact(item));
      state.data.history = history.map(item => normalizer.history(item));
      state.data.tasks = tasks.map(item => normalizer.task(item));

      dataModel.enrich();
    },

    // Liest alle Choice-Felder aller relevanten Listen aus SharePoint
    // Schreibt in state.meta.choices[listTitle][spFieldName] = ["Wert1", "Wert2", ...]
    // Wird bei loadAll() und handleRefresh() mitgeladen — SP ist Single Source of Truth
    async loadColumnChoices() {
      const lists = [
        CONFIG.lists.firms,
        CONFIG.lists.contacts,
        CONFIG.lists.history,
        CONFIG.lists.tasks
      ];

      const siteId = await this.getSiteId();

      await Promise.all(lists.map(async (listTitle) => {
        try {
          const data = await this.graphRequest(
            `/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/columns`
          );

          const choicesForList = {};
          for (const col of (data.value || [])) {
            if (col.choice && Array.isArray(col.choice.choices) && col.choice.choices.length > 0) {
              choicesForList[col.name] = col.choice.choices;
            }
          }
          state.meta.choices[listTitle] = choicesForList;
        } catch (err) {
          // Nicht-fatal: Choices bleiben leer, Formular fällt auf Freitext zurück
          console.warn(`loadColumnChoices fehlgeschlagen für ${listTitle}:`, err);
          state.meta.choices[listTitle] = {};
        }
      }));
    },

    // Write-Layer — POST (neues Item anlegen)
    async postItem(listTitle, fields) {
      const siteId = await this.getSiteId();
      return await this.graphRequest(
        `/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/items`,
        { method: "POST", body: { fields } }
      );
    },

    // Write-Layer — PATCH (bestehendes Item aktualisieren)
    async patchItem(listTitle, itemId, fields) {
      const siteId = await this.getSiteId();
      return await this.graphRequest(
        `/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/items/${itemId}/fields`,
        {
          method: "PATCH",
          body: fields,
          // If-Match: * — überschreibt unabhängig vom eTag, verhindert 409 resourceModified
          // bei parallelen Patches auf dieselbe Liste (SharePoint refresht eTags listenweit).
          headers: { "If-Match": "*" }
        }
      );
    },

    // Write-Layer — DELETE
    async deleteItem(listTitle, itemId) {
      const siteId = await this.getSiteId();
      return await this.graphRequest(
        `/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/items/${itemId}`,
        { method: "DELETE" }
      );
    }
  };

  const normalizer = {
    getField(item, fieldName) { return item?.fields?.[fieldName]; },
    itemId(item) { return Number(item?.id) || null; },

    firm(item) {
      const f = SCHEMA.firms.fields;
      return {
        id: this.itemId(item),
        title: this.getField(item, f.title) || "",
        adresse: this.getField(item, f.adresse) || "",
        plz: this.getField(item, f.plz) || "",
        ort: this.getField(item, f.ort) || "",
        land: this.getField(item, f.land) || "",
        hauptnummer: this.getField(item, f.hauptnummer) || "",
        klassifizierung: this.getField(item, f.klassifizierung) || "",
        vip: helpers.bool(this.getField(item, f.vip)),
        kategorie: (this.getField(item, f.kategorie) || "").trim()
      };
    },

    contact(item) {
      const f = SCHEMA.contacts.fields;
      return {
        id: this.itemId(item),
        nachname: this.getField(item, f.nachname) || "",
        vorname: this.getField(item, f.vorname) || "",
        anrede: this.getField(item, f.anrede) || "",
        firmaRaw: this.getField(item, f.firma),
        firmaLookupId: Number(this.getField(item, f.firmaLookupId)) || null,
        funktion: this.getField(item, f.funktion) || "",
        email1: this.getField(item, f.email1) || "",
        email2: this.getField(item, f.email2) || "",
        direktwahl: this.getField(item, f.direktwahl) || "",
        mobile: this.getField(item, f.mobile) || "",
        rolle: this.getField(item, f.rolle) || "",
        leadbbz0: this.getField(item, f.leadbbz0) || "",
        sgf: helpers.normalizeChoiceList(this.getField(item, f.sgf)),
        geburtstag: this.getField(item, f.geburtstag) || "",
        kommentar: this.getField(item, f.kommentar) || "",
        event: helpers.normalizeChoiceList(this.getField(item, f.event)),
        // FIX: eventhistory konsistent als Array normalisieren (wie sgf und event)
        eventhistory: helpers.normalizeChoiceList(this.getField(item, f.eventhistory)),
        archiviert: helpers.bool(this.getField(item, f.archiviert)),
        // SP-Metadaten (Datenqualität) — direkt am Item-Objekt, nicht in fields
        spCreated: item?.createdDateTime || "",
        spCreatedBy: item?.createdBy?.user?.displayName || "",
        spModified: item?.lastModifiedDateTime || "",
        spModifiedBy: item?.lastModifiedBy?.user?.displayName || ""
      };
    },

    history(item) {
      const f = SCHEMA.history.fields;
      return {
        id: this.itemId(item),
        title: this.getField(item, f.title) || "",
        kontaktRaw: this.getField(item, f.kontakt),
        kontaktLookupId: Number(this.getField(item, f.kontaktLookupId)) || null,
        datum: this.getField(item, f.datum) || "",
        typ: this.getField(item, f.typ) || "",
        notizen: this.getField(item, f.notizen) || "",
        projektbezug: this.getField(item, f.projektbezug) || "",
        leadbbz: this.getField(item, f.leadbbz) || "",
        // SP-Metadaten für Admin-Auswertungen
        spCreated: item?.createdDateTime || "",
        spCreatedBy: item?.createdBy?.user?.displayName || "",
        spModified: item?.lastModifiedDateTime || "",
        spModifiedBy: item?.lastModifiedBy?.user?.displayName || ""
      };
    },

    task(item) {
      const f = SCHEMA.tasks.fields;
      return {
        id: this.itemId(item),
        title: this.getField(item, f.title) || "",
        kontaktRaw: this.getField(item, f.kontakt),
        kontaktLookupId: Number(this.getField(item, f.kontaktLookupId)) || null,
        deadline: this.getField(item, f.deadline) || "",
        status: this.getField(item, f.status) || "",
        leadbbz: this.getField(item, f.leadbbz) || "",
        // SP-Metadaten für Admin-Auswertungen
        spCreated: item?.createdDateTime || "",
        spCreatedBy: item?.createdBy?.user?.displayName || "",
        spModified: item?.lastModifiedDateTime || "",
        spModifiedBy: item?.lastModifiedBy?.user?.displayName || ""
      };
    }
  };

  const dataModel = {
    enrich() {
      const firmById = new Map(state.data.firms.map(f => [f.id, f]));
      const contactById = new Map(state.data.contacts.map(c => [c.id, c]));

      const contacts = state.data.contacts.map(contact => {
        const firm = firmById.get(contact.firmaLookupId) || null;
        return { ...contact, fullName: helpers.fullName(contact), firmId: firm?.id || contact.firmaLookupId || null, firmTitle: firm?.title || contact.firmaRaw || "", firm };
      });

      const history = state.data.history.map(entry => {
        const contact = contactById.get(entry.kontaktLookupId) || null;
        const firm = contact ? firmById.get(contact.firmaLookupId) || null : null;
        return { ...entry, contactId: contact?.id || entry.kontaktLookupId || null, contactName: contact ? helpers.fullName(contact) : (entry.kontaktRaw || ""), firmId: firm?.id || null, firmTitle: firm?.title || "", projektbezugBool: helpers.bool(entry.projektbezug) };
      });

      const tasks = state.data.tasks.map(task => {
        const contact = contactById.get(task.kontaktLookupId) || null;
        const firm = contact ? firmById.get(contact.firmaLookupId) || null : null;
        return { ...task, contactId: contact?.id || task.kontaktLookupId || null, contactName: contact ? helpers.fullName(contact) : (task.kontaktRaw || ""), firmId: firm?.id || null, firmTitle: firm?.title || "", isOpen: helpers.isOpenTask(task.status), isOverdue: helpers.isOverdue(task.deadline) };
      });

      const firms = state.data.firms.map(firm => {
        const firmContacts = contacts.filter(c => c.firmId === firm.id);
        const firmContactIds = new Set(firmContacts.map(c => c.id));
        const firmTasks = tasks.filter(t => firmContactIds.has(t.contactId));
        const firmHistory = history.filter(h => firmContactIds.has(h.contactId));
        const openTasks = firmTasks.filter(t => t.isOpen);
        const nextDeadlineTask = openTasks.filter(t => helpers.toDate(t.deadline)).sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline))[0] || null;
        const latestHistory = [...firmHistory].sort((a, b) => helpers.compareDateDesc(a.datum, b.datum))[0] || null;

        return {
          ...firm,
          contactsCount: firmContacts.length,
          contacts: firmContacts.sort((a, b) => a.fullName.localeCompare(b.fullName, "de")),
          tasks: firmTasks.sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline)),
          history: firmHistory.sort((a, b) => helpers.compareDateDesc(a.datum, b.datum)),
          openTasksCount: openTasks.length,
          nextDeadline: nextDeadlineTask?.deadline || "",
          latestActivity: latestHistory?.datum || ""
        };
      });

      const eventMap = new Map();
      contacts.forEach(contact => {
        const contactTasks = tasks.filter(t => t.contactId === contact.id);
        const contactHistory = history.filter(h => h.contactId === contact.id).sort((a, b) => helpers.compareDateDesc(a.datum, b.datum));
        const latestH = contactHistory[0] || null;
        const openTasks = contactTasks.filter(t => t.isOpen);

        contact.event.forEach(eventName => {
          const key = String(eventName || "").trim();
          if (!key) return;
          if (!eventMap.has(key)) eventMap.set(key, { name: key, contacts: [], contactCount: 0, openTasksCount: 0 });
          eventMap.get(key).contacts.push({
            contactId: contact.id,
            contactName: contact.fullName || contact.nachname,
            firmId: contact.firmId,
            firmTitle: contact.firmTitle,
            rolle: contact.rolle,
            funktion: contact.funktion,
            eventhistory: contact.eventhistory,
            segment: contact.firm ? String(contact.firm.klassifizierung || "").toUpperCase() : "",
            leadbbz: contact.leadbbz0 || "",
            sgf: contact.sgf || [],
            latestHistoryDate: latestH?.datum || "",
            latestHistoryType: latestH?.typ || "",
            latestHistoryText: latestH?.notizen || "",
            openTasksCount: openTasks.length,
            email1: contact.email1
          });
        });
      });

      const eventChoicesOrder = state.meta.choices?.[CONFIG.lists.contacts]?.["Event"] || [];
      const eventOrderIndex = (name) => {
        const idx = eventChoicesOrder.indexOf(name);
        return idx === -1 ? 9999 : idx; // nicht in SP-Choices → ans Ende
      };

      const events = [...eventMap.values()]
        .map(group => ({ ...group, contactCount: group.contacts.length, openTasksCount: group.contacts.reduce((sum, c) => sum + c.openTasksCount, 0), contacts: group.contacts.sort((a, b) => String(a.contactName).localeCompare(String(b.contactName), "de")) }))
        .sort((a, b) => {
          const ia = eventOrderIndex(a.name), ib = eventOrderIndex(b.name);
          if (ia !== ib) return ia - ib;
          return a.name.localeCompare(b.name, "de"); // Fallback alphabetisch
        });

      state.enriched.contacts = contacts.sort((a, b) => a.fullName.localeCompare(b.fullName, "de"));
      state.enriched.history = history.sort((a, b) => helpers.compareDateDesc(a.datum, b.datum));
      state.enriched.tasks = tasks.sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline));
      state.enriched.firms = firms.sort((a, b) => a.title.localeCompare(b.title, "de"));
      state.enriched.events = events;

      // privateFirmId nach jedem enrich() neu auflösen — robust gegen SP-ID-Änderungen
      const privateFirm = state.data.firms.find(
        f => String(f.title).trim() === CONFIG.defaults.privateFirmTitle
      );
      state.meta.privateFirmId = privateFirm?.id || null;
    },

    getFirmById(id) { return state.enriched.firms.find(f => String(f.id) === String(id)) || null; },
    getContactById(id) { return state.enriched.contacts.find(c => String(c.id) === String(id)) || null; }
  };

  const views = {
    // metaType: "" | "alert" | "warn" | "ok"
    kpiBlock(label, value, meta = "", metaType = "") {
      const metaClass = metaType === "alert" ? "bbz-kpi-meta-alert"
        : metaType === "warn" ? "bbz-kpi-meta-warn"
        : metaType === "ok"   ? "bbz-kpi-meta-ok"
        : "bbz-kpi-meta";
      return `<div class="bbz-kpi"><div class="bbz-kpi-label">${helpers.escapeHtml(label)}</div><div class="bbz-kpi-value">${helpers.escapeHtml(String(value))}</div>${meta ? `<div class="${metaClass}">${helpers.escapeHtml(meta)}</div>` : ""}</div>`;
    },

    miniItem(title, meta) {
      return `<div class="bbz-mini-item"><div class="bbz-mini-title">${title}</div><div class="bbz-mini-meta">${meta}</div></div>`;
    },

    renderRoute() {
      if (state.meta.loading) return ui.loadingBlock();

      let viewHtml = "";
      switch (state.filters.route) {
        case "firms": viewHtml = state.selection.firmId ? this.firmDetail() : this.firms(); break;
        case "contacts": viewHtml = state.selection.contactId ? this.contactDetail() : this.contacts(); break;
        case "planning": viewHtml = this.planning(); break;
        case "history": viewHtml = this.historyView(); break;
        case "events": viewHtml = this.events(); break;
        case "birthdays": viewHtml = this.birthdayView(); break;
        case "admin": viewHtml = this.adminPanel(); break;
        default: viewHtml = this.firms();
      }

      // Modal wird ueber dem View gerendert
      let modalHtml = "";
      if (state.modal?.type === "contact") modalHtml = views.renderContactForm(state.modal.mode, state.modal.payload);
      if (state.modal?.type === "firm")    modalHtml = views.renderFirmForm(state.modal.mode, state.modal.payload?.firmId);
      if (state.modal?.type === "history") modalHtml = views.renderHistoryForm(state.modal.payload);
      if (state.modal?.type === "task")    modalHtml = views.renderTaskForm(state.modal.payload);
      if (state.modal?.type === "batch-event") modalHtml = views.renderBatchEventForm(state.modal.payload);
      if (state.modal?.type === "event-einladung") modalHtml = views.renderEventEinladungModal(state.modal.payload);
      if (state.modal?.type === "event-nachbearbeitung") modalHtml = views.renderEventNachbearbeitungModal(state.modal.payload);
      if (state.modal?.type === "event-matrix") modalHtml = views.renderEventMatrixModal(state.modal.payload);
      return viewHtml + modalHtml;
    },

    // Kontakt-Formular — FIX 1 (toDateInput) integriert, FIX 2 (Modal-Infrastruktur) verdrahtet
    renderContactForm(mode, payload = {}) {
      const itemId = Number(payload.itemId || 0) || null;
      const contact = mode === "edit" ? dataModel.getContactById(itemId) : null;
      const title = mode === "edit" ? "Kontakt bearbeiten" : "Neuer Kontakt";
      const preselectedFirmId = Number(payload.prefillFirmId || contact?.firmId || 0) || "";
      const L = CONFIG.lists.contacts;
      // Privatpersonen-Modus: wenn Firma "Privatpersonen" vorgewählt oder gesetzt
      const isPrivat = state.meta.privateFirmId !== null &&
        (String(preselectedFirmId) === String(state.meta.privateFirmId) ||
         (contact && contact.firmId === state.meta.privateFirmId));

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${title}</div>
              <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
            <form data-modal-form="contact" data-mode="${mode}" data-item-id="${itemId || ""}" style="display:flex;flex-direction:column;flex:1;min-height:0;">
              <div class="bbz-modal-body" style="flex:1;overflow-y:auto;-webkit-overflow-scrolling:touch;">
                <div class="bbz-form-grid">

                  <div class="bbz-field">
                    <label>Nachname *</label>
                    <input class="bbz-input" name="nachname" required value="${helpers.escapeHtml(contact?.nachname || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Vorname</label>
                    <input class="bbz-input" name="vorname" value="${helpers.escapeHtml(contact?.vorname || "")}" />
                  </div>

                  <div class="bbz-field">
                    <label>Anrede</label>
                    ${helpers.choiceSelectHtml("anrede", L, "Anrede", contact?.anrede || "")}
                  </div>
                  <div class="bbz-field">
                    <label>Firma *</label>
                    <select class="bbz-select" name="firmaLookupId" required>
                      <option value="">— bitte wählen —</option>
                      ${state.enriched.firms.map(f => `<option value="${f.id}" ${String(preselectedFirmId) === String(f.id) ? "selected" : ""}>${helpers.escapeHtml(f.title)}</option>`).join("")}
                    </select>
                  </div>

                  <div class="bbz-field">
                    <label>Funktion</label>
                    <input class="bbz-input" name="funktion" value="${helpers.escapeHtml(contact?.funktion || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Rolle</label>
                    ${helpers.choiceSelectHtml("rolle", L, "Rolle", contact?.rolle || "")}
                  </div>

                  <div class="bbz-field">
                    <label>Email 1</label>
                    <input class="bbz-input" name="email1" value="${helpers.escapeHtml(contact?.email1 || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Email 2</label>
                    <input class="bbz-input" name="email2" value="${helpers.escapeHtml(contact?.email2 || "")}" />
                  </div>

                  <div class="bbz-field">
                    <label>Direktwahl</label>
                    <input class="bbz-input" name="direktwahl" value="${helpers.escapeHtml(contact?.direktwahl || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Mobile</label>
                    <input class="bbz-input" name="mobile" value="${helpers.escapeHtml(contact?.mobile || "")}" />
                  </div>

                  <div class="bbz-field">
                    <label>Geburtstag</label>
                    <input type="date" class="bbz-input" name="geburtstag" value="${helpers.escapeHtml(helpers.toDateInput(contact?.geburtstag || ""))}" />
                  </div>
                  <div class="bbz-field">
                    <label>Leadbbz</label>
                    ${helpers.choiceSelectHtml("leadbbz0", L, "Leadbbz0", contact?.leadbbz0 || "")}
                  </div>

                  <div class="bbz-field bbz-span-2">
                    <label>SGF <span class="bbz-field-hint">(Mehrfachauswahl)</span></label>
                    ${helpers.choiceMultiHtml("sgf", L, "SGF", contact?.sgf || [])}
                  </div>

                  <div class="bbz-field bbz-span-2">
                    <label>Event <span class="bbz-field-hint">(Mehrfachauswahl)</span></label>
                    ${helpers.choiceMultiHtml("event", L, "Event", contact?.event || [])}
                  </div>

                  <div class="bbz-field bbz-span-2">
                    <label>Eventhistory <span class="bbz-field-hint">(Mehrfachauswahl)</span></label>
                    ${helpers.choiceMultiHtml("eventhistory", L, "Eventhistory", contact?.eventhistory || [])}
                  </div>

                  <div class="bbz-field bbz-span-2">
                    <label data-kommentar-label>${isPrivat ? 'Adresse / Notizen (Privatperson — Adresse hier erfassen)' : 'Kommentar'}</label>
                    <textarea class="bbz-textarea" name="kommentar">${helpers.escapeHtml(contact?.kommentar || "")}</textarea>
                  </div>

                  <label class="bbz-checkbox">
                    <input type="checkbox" name="archiviert" ${contact?.archiviert ? "checked" : ""} />
                    Archiviert
                  </label>

                </div>
              </div>
              <div class="bbz-modal-footer">
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary" ${state.meta.loading ? "disabled" : ""}>Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    renderFirmForm(mode, firmId = null) {
      const firm = mode === "edit" ? dataModel.getFirmById(firmId) : null;
      const title = mode === "edit" ? "Firma bearbeiten" : "Neue Firma";
      const LF = CONFIG.lists.firms;

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${title}</div>
              <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
            <form data-modal-form="firm" data-mode="${mode}" data-item-id="${firmId || ""}" style="display:flex;flex-direction:column;flex:1;min-height:0;">
              <div class="bbz-modal-body" style="flex:1;overflow-y:auto;-webkit-overflow-scrolling:touch;">
                <div class="bbz-form-grid">
                  <div class="bbz-field bbz-span-2">
                    <label>Firmenname *</label>
                    <input class="bbz-input" name="title" required value="${helpers.escapeHtml(firm?.title || "")}" />
                  </div>
                  <div class="bbz-field bbz-span-2">
                    <label>Adresse</label>
                    <input class="bbz-input" name="adresse" value="${helpers.escapeHtml(firm?.adresse || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>PLZ</label>
                    <input class="bbz-input" name="plz" value="${helpers.escapeHtml(firm?.plz || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Ort</label>
                    <input class="bbz-input" name="ort" value="${helpers.escapeHtml(firm?.ort || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Land</label>
                    <input class="bbz-input" name="land" value="${helpers.escapeHtml(firm?.land || "Schweiz")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Hauptnummer</label>
                    <input class="bbz-input" name="hauptnummer" value="${helpers.escapeHtml(firm?.hauptnummer || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Klassifizierung</label>
                    ${helpers.choiceSelectHtml("klassifizierung", LF, "Klassifizierung", firm?.klassifizierung || "")}
                  </div>
                  <div class="bbz-field">
                    <label class="bbz-checkbox" style="border:none;padding:0;margin-top:24px;">
                      <input type="checkbox" name="vip" ${firm?.vip ? "checked" : ""} />
                      VIP
                    </label>
                  </div>
                </div>
              </div>
              <div class="bbz-modal-footer">
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary" ${state.meta.loading ? "disabled" : ""}>Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    renderHistoryForm(payload = {}) {
      const mode = payload.mode || "create";
      const itemId = Number(payload.itemId || 0) || null;
      const entry = mode === "edit" ? state.enriched.history.find(h => h.id === itemId) || null : null;
      const prefillContactId = Number(payload.prefillContactId || entry?.contactId || 0) || "";
      const LH = CONFIG.lists.history;
      const title = mode === "edit" ? "Aktivitaet bearbeiten" : "Aktivitaet erfassen";

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${title}</div>
              <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
            <form data-modal-form="history" data-mode="${mode}" data-item-id="${itemId || ""}" style="display:flex;flex-direction:column;flex:1;min-height:0;">
              <div class="bbz-modal-body" style="flex:1;overflow-y:auto;-webkit-overflow-scrolling:touch;">
                <div class="bbz-form-grid">
                  <div class="bbz-field">
                    <label>Kontakt *</label>
                    <select class="bbz-select" name="kontaktLookupId" required ${mode === "edit" ? "disabled" : ""}>
                      <option value="">— bitte waehlen —</option>
                      ${state.enriched.contacts.filter(c => !c.archiviert || (entry && c.id === entry.contactId)).map(c => `<option value="${c.id}" ${String(prefillContactId) === String(c.id) ? "selected" : ""}>${helpers.escapeHtml(c.fullName || c.nachname)}${c.firmTitle ? " — " + helpers.escapeHtml(c.firmTitle) : ""}</option>`).join("")}
                    </select>
                    ${mode === "edit" ? `<input type="hidden" name="kontaktLookupId" value="${prefillContactId}" />` : ""}
                  </div>
                  <div class="bbz-field">
                    <label>Datum *</label>
                    <input type="date" class="bbz-input" name="datum" required value="${helpers.toDateInput(entry?.datum || new Date())}" />
                  </div>
                  <div class="bbz-field">
                    <label>Kontaktart</label>
                    ${helpers.choiceSelectHtml("kontaktart", LH, "Kontaktart", entry?.typ || payload.prefillTyp || "")}
                  </div>
                  <div class="bbz-field">
                    <label>Leadbbz</label>
                    ${helpers.choiceSelectHtml("leadbbz", LH, "Leadbbz", entry?.leadbbz || "")}
                  </div>
                  <div class="bbz-field bbz-span-2">
                    <label>Projektbezug</label>
                    <label class="bbz-checkbox" style="border:none;padding:0;">
                      <input type="checkbox" name="projektbezug" ${entry?.projektbezugBool ? "checked" : ""} />
                      Ja, mit Projektbezug
                    </label>
                  </div>
                  <div class="bbz-field bbz-span-2">
                    <label>Notizen</label>
                    <textarea class="bbz-textarea" name="notizen" rows="4" placeholder="Was wurde besprochen?">${helpers.escapeHtml(entry?.notizen || "")}</textarea>
                  </div>
                </div>
              </div>
              <div class="bbz-modal-footer">
                <div style="flex:1;">
                  ${mode === "edit" ? `<button type="button" class="bbz-button bbz-button-secondary" style="color:var(--red);border-color:var(--red);" data-action="delete-history" data-id="${itemId}" data-title="${helpers.escapeHtml(entry?.typ || entry?.title || 'Eintrag')}">Löschen</button>` : ""}
                </div>
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary" ${state.meta.loading ? "disabled" : ""}>Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    renderTaskForm(payload = {}) {
      const mode = payload.mode || "create";
      const itemId = Number(payload.itemId || 0) || null;
      const task = mode === "edit" ? state.enriched.tasks.find(t => t.id === itemId) || null : null;
      const prefillContactId = Number(payload.prefillContactId || task?.contactId || 0) || "";
      const LT = CONFIG.lists.tasks;
      const title = mode === "edit" ? "Aufgabe bearbeiten" : "Aufgabe erfassen";

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${title}</div>
              <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
            <form data-modal-form="task" data-mode="${mode}" data-item-id="${itemId || ""}" style="display:flex;flex-direction:column;flex:1;min-height:0;">
              <div class="bbz-modal-body" style="flex:1;overflow-y:auto;-webkit-overflow-scrolling:touch;">
                <div class="bbz-form-grid">
                  <div class="bbz-field bbz-span-2">
                    <label>Titel *</label>
                    <input class="bbz-input" name="title" required value="${helpers.escapeHtml(task?.title || "")}" placeholder="Was ist zu tun?" />
                  </div>
                  <div class="bbz-field">
                    <label>Kontakt *</label>
                    <select class="bbz-select" name="kontaktLookupId" required ${mode === "edit" ? "disabled" : ""}>
                      <option value="">— bitte waehlen —</option>
                      ${state.enriched.contacts.filter(c => !c.archiviert || (task && c.id === task.contactId)).map(c => `<option value="${c.id}" ${String(prefillContactId) === String(c.id) ? "selected" : ""}>${helpers.escapeHtml(c.fullName || c.nachname)}${c.firmTitle ? " — " + helpers.escapeHtml(c.firmTitle) : ""}</option>`).join("")}
                    </select>
                    ${mode === "edit" ? `<input type="hidden" name="kontaktLookupId" value="${prefillContactId}" />` : ""}
                  </div>
                  <div class="bbz-field">
                    <label>Deadline</label>
                    <input type="date" class="bbz-input" name="deadline" value="${helpers.toDateInput(task?.deadline || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Status</label>
                    ${helpers.choiceSelectHtml("status", LT, "Status", task?.status || "")}
                  </div>
                  <div class="bbz-field">
                    <label>Leadbbz</label>
                    ${helpers.choiceSelectHtml("leadbbz", LT, "Leadbbz", task?.leadbbz || "")}
                  </div>
                </div>
              </div>
              <div class="bbz-modal-footer">
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary" ${state.meta.loading ? "disabled" : ""}>Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    renderBatchEventForm(payload = {}) {
      const { eventName = "", mode = "anmelden", filterSegment = "", filterLeadbbz = "", filterSgf = "", filterSearch = "", selected = [], selectedHistoryCategory = "" } = payload;
      const LC = CONFIG.lists.contacts;
      const isEventhistory = mode === "eventhistory";

      // SP-Choices für Eventhistory-Feld laden
      const eventhistoryChoices = state.meta.choices?.[CONFIG.lists.contacts]?.["Eventhistory"] || [];
      const eventChoices        = state.meta.choices?.[CONFIG.lists.contacts]?.["Event"] || [];

      // Im Eventhistory-Modus: Kategorie muss erst gewählt werden
      const activeCategory = isEventhistory ? selectedHistoryCategory : eventName;
      const categoryMissing = isEventhistory && !activeCategory;

      const allLeadbbz = [...new Set(state.enriched.contacts.map(c => c.leadbbz0).filter(Boolean))].sort();
      const allSgf     = [...new Set(state.enriched.contacts.flatMap(c => helpers.toArray(c.sgf)))].filter(Boolean).sort();

      // Kandidaten berechnen — nur wenn Kategorie bekannt
      let candidates = [];
      if (!categoryMissing) {
        const existingContactIds = isEventhistory
          ? new Set() // Eventhistory: alle Kontakte wählbar
          : new Set((state.enriched.events.find(g => g.name === activeCategory)?.contacts || []).map(c => c.contactId));

        candidates = isEventhistory
          ? state.enriched.contacts.filter(c => !c.archiviert)
          : state.enriched.contacts.filter(c => !c.archiviert && !existingContactIds.has(c.id));

        if (filterSegment) {
          const firmMap = new Map(state.enriched.firms.map(f => [f.id, f]));
          candidates = candidates.filter(c => String(firmMap.get(c.firmId)?.klassifizierung || "").toUpperCase().startsWith(filterSegment));
        }
        if (filterLeadbbz) candidates = candidates.filter(c => c.leadbbz0 === filterLeadbbz);
        if (filterSgf) candidates = candidates.filter(c => helpers.toArray(c.sgf).includes(filterSgf));
        if (filterSearch.trim()) {
          const s = filterSearch.trim().toLowerCase();
          candidates = candidates.filter(c => [c.fullName, c.firmTitle].some(v => helpers.textIncludes(v, s)));
        }
      }

      const previewContacts = candidates.slice(0, 200);
      const validSelected = categoryMissing ? [] : selected.filter(id => previewContacts.some(c => c.id === id));
      if (state.modal?.payload) {
        state.modal.payload.previewContacts = previewContacts;
        state.modal.payload.selected = validSelected;
      }
      const allChecked = previewContacts.length > 0 && previewContacts.every(c => validSelected.includes(c.id));

      const leadbbzOptions = [`<option value="">— alle Lead BBZ —</option>`, ...allLeadbbz.map(l =>
        `<option value="${helpers.escapeHtml(l)}" ${filterLeadbbz === l ? "selected" : ""}>${helpers.escapeHtml(l)}</option>`)].join("");
      const sgfOptions = [`<option value="">— alle SGF —</option>`, ...allSgf.map(s =>
        `<option value="${helpers.escapeHtml(s)}" ${filterSgf === s ? "selected" : ""}>${helpers.escapeHtml(s)}</option>`)].join("");

      const modeLabel = isEventhistory ? "Eventhistory setzen" : "Event setzen";
      const choicesForDropdown = isEventhistory ? eventhistoryChoices : [];

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal" style="max-width:780px;width:95vw;">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${isEventhistory ? "Eventhistory setzen" : `${helpers.escapeHtml(activeCategory)} — Event setzen`}</div>
              <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
            <form data-modal-form="batch-event" data-event-name="${helpers.escapeHtml(activeCategory)}" data-mode="${mode}" style="display:flex;flex-direction:column;flex:1;min-height:0;">
              <div class="bbz-modal-body" style="flex:1;overflow-y:auto;-webkit-overflow-scrolling:touch;">

                ${isEventhistory ? `
                <!-- Schritt 1: Kategorie wählen -->
                <div class="bbz-field" style="margin-bottom:14px;">
                  <label style="font-size:13px;font-weight:500;display:block;margin-bottom:6px;">Eventhistory-Kategorie *</label>
                  ${eventhistoryChoices.length
                    ? `<select class="bbz-select" data-filter="batch-eventhistory-category" style="max-width:360px;">
                        <option value="">— Kategorie wählen —</option>
                        ${eventhistoryChoices.map(c => `<option value="${helpers.escapeHtml(c)}" ${selectedHistoryCategory === c ? "selected" : ""}>${helpers.escapeHtml(c)}</option>`).join("")}
                       </select>`
                    : `<input class="bbz-input" data-filter="batch-eventhistory-category-text" type="text"
                         placeholder="Kategorie eingeben (Choices nicht geladen)" value="${helpers.escapeHtml(selectedHistoryCategory)}" style="max-width:360px;" />`
                  }
                </div>
                ${categoryMissing ? `<div style="font-size:13px;color:var(--muted);padding:12px 0;">Bitte zuerst eine Kategorie wählen.</div>` : ""}
                ` : ""}

                ${!categoryMissing ? `
                <!-- Filterzeile -->
                <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;margin-bottom:12px;">
                  <input class="bbz-input" data-filter="batch-search" type="text" placeholder="Name / Firma ..." value="${helpers.escapeHtml(filterSearch)}" style="font-size:12px;" />
                  <select class="bbz-select" data-filter="batch-segment" style="font-size:12px;">
                    <option value="">— Segment —</option>
                    ${["A","B","C"].map(v=>`<option value="${v}" ${filterSegment===v?"selected":""}>${v}</option>`).join("")}
                  </select>
                  <select class="bbz-select" data-filter="batch-leadbbz" style="font-size:12px;">${leadbbzOptions}</select>
                  <select class="bbz-select" data-filter="batch-sgf" style="font-size:12px;">${sgfOptions}</select>
                </div>

                <!-- Kontakt-Tabelle -->
                <div class="bbz-table-wrap" style="max-height:340px;overflow-y:auto;">
                  <table class="bbz-table" style="min-width:500px;">
                    <thead><tr>
                      <th style="width:32px;">
                        <input type="checkbox" data-action="batch-toggle-all" ${allChecked ? "checked" : ""} title="Alle/Keine" />
                      </th>
                      <th>Kontakt</th>
                      <th>Firma</th>
                      <th>Segment</th>
                      <th>Lead BBZ</th>
                    </tr></thead>
                    <tbody>
                      ${previewContacts.length ? previewContacts.map(c => {
                        const firmObj = state.enriched.firms.find(f => f.id === c.firmId);
                        const seg = String(firmObj?.klassifizierung || "").toUpperCase();
                        const isChecked = validSelected.includes(c.id);
                        return `<tr style="${isChecked ? "background:var(--blue-light);" : ""}">
                          <td><input type="checkbox" data-action="batch-toggle-contact" data-contact-id="${c.id}" ${isChecked ? "checked" : ""} /></td>
                          <td>${helpers.avatarHtml(c)} <span style="margin-left:6px;">${helpers.escapeHtml(c.fullName || c.nachname)}</span></td>
                          <td><span class="bbz-muted" style="font-size:12px;">${helpers.escapeHtml(c.firmTitle || "—")}</span></td>
                          <td>${seg ? `<span class="${helpers.firmBadgeClass(seg)}">${helpers.escapeHtml(seg)}</span>` : '<span class="bbz-muted">—</span>'}</td>
                          <td>${helpers.leadbbzBadgeHtml(c.leadbbz0)}</td>
                        </tr>`;
                      }).join("") : `<tr><td colspan="5"><div class="bbz-empty" style="padding:16px;">Keine Kontakte für diese Filter.</div></td></tr>`}
                    </tbody>
                  </table>
                </div>
                <div data-batch-counter style="font-size:12px;color:var(--muted);margin-top:8px;">
                  ${validSelected.length} von ${previewContacts.length} ausgewählt
                  ${previewContacts.length === 200 ? " (max. 200 — Filter verfeinern)" : ""}
                </div>
                <input type="hidden" name="selectedIds" value="${helpers.escapeHtml(JSON.stringify(validSelected))}" />
                ` : ""}
              </div>
              <div class="bbz-modal-footer">
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary"
                  ${state.meta.loading || validSelected.length === 0 || categoryMissing ? "disabled" : ""}>
                  ${isEventhistory
                    ? `+ ${validSelected.length} × Eventhistory «${helpers.escapeHtml(activeCategory || "?")}» setzen`
                    : `+ ${validSelected.length} × Event «${helpers.escapeHtml(activeCategory)}» setzen`}
                </button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    birthdayView() {
      const allActive   = state.enriched.contacts.filter(c => !c.archiviert);
      const withBday    = allActive.filter(c => c.geburtstag);
      const total       = allActive.length;
      const covered     = withBday.length;
      const pct         = total > 0 ? Math.round((covered / total) * 100) : 0;
      const pctStyle    = pct >= 80 ? "color:var(--green);" : pct >= 50 ? "color:var(--amber);" : "color:var(--red);";

      // Alle Geburtstage des Jahres sortiert nach Monat/Tag
      const today = helpers.todayStart();
      const allYearBdays = withBday.map(c => {
        const bDay = helpers.toDate(c.geburtstag);
        if (!bDay) return null;
        let next = new Date(today.getFullYear(), bDay.getMonth(), bDay.getDate());
        if (next < today) next = new Date(today.getFullYear() + 1, bDay.getMonth(), bDay.getDate());
        const daysUntil = Math.round((next - today) / 86400000);
        const age = next.getFullYear() - bDay.getFullYear();
        return { contact: c, daysUntil, nextBirthday: next, age };
      }).filter(Boolean).sort((a, b) => a.daysUntil - b.daysUntil);

      // Gruppen
      const gToday = allYearBdays.filter(b => b.daysUntil === 0);
      const gWeek  = allYearBdays.filter(b => b.daysUntil > 0 && b.daysUntil <= 7);
      const g30    = allYearBdays.filter(b => b.daysUntil > 7 && b.daysUntil <= 30);
      const gRest  = allYearBdays.filter(b => b.daysUntil > 30);

      const monthNames = ["Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember"];

      const bdayRow = (b) => {
        const lbl = helpers.birthdayLabel(b.daysUntil, b.nextBirthday);
        const lblStyle = b.daysUntil === 0
          ? "background:var(--blue-light);border-color:#a8c8e0;color:var(--blue);"
          : b.daysUntil <= 7
          ? "background:#fff9eb;border-color:#f4dfab;color:var(--amber);"
          : "";
        const dateStr = b.nextBirthday.toLocaleDateString("de-CH", { day: "2-digit", month: "2-digit" });
        return `
          <div style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid var(--line-2);">
            ${helpers.avatarHtml(b.contact)}
            <div style="flex:1;min-width:0;">
              <a class="bbz-link" data-action="open-contact" data-id="${b.contact.id}" style="font-size:13px;font-weight:600;">${helpers.escapeHtml(b.contact.fullName || b.contact.nachname)}</a>
              <div style="font-size:12px;color:var(--muted);">${helpers.escapeHtml(b.contact.firmTitle || "—")}${b.contact.funktion ? ` · ${helpers.escapeHtml(b.contact.funktion)}` : ""}</div>
            </div>
            <div style="text-align:right;flex-shrink:0;">
              <span class="bbz-chip" style="font-size:11px;${lblStyle}">${helpers.escapeHtml(lbl)}</span>
              <div style="font-size:11px;color:var(--muted);margin-top:3px;">${dateStr} · wird ${b.age}</div>
            </div>
          </div>`;
      };

      const section = (title, items, emptyText = "") => {
        if (!items.length && !emptyText) return "";
        return `
          <section class="bbz-section" style="margin-bottom:12px;">
            <div class="bbz-section-header">
              <div><div class="bbz-section-title">${title}</div></div>
              <span style="font-size:12px;color:var(--muted);">${items.length} ${items.length === 1 ? "Kontakt" : "Kontakte"}</span>
            </div>
            <div class="bbz-section-body">
              ${items.length ? items.map(bdayRow).join("") : `<div class="bbz-empty">${emptyText}</div>`}
            </div>
          </section>`;
      };

      // Restliche Geburtstage nach Monat gruppieren
      const byMonth = {};
      for (const b of gRest) {
        const key = b.nextBirthday.getMonth();
        if (!byMonth[key]) byMonth[key] = [];
        byMonth[key].push(b);
      }
      const restByMonthHtml = Object.entries(byMonth).sort((a, b) => {
        // Monate ab heute aufsteigend
        const todayMonth = today.getMonth();
        const ma = ((Number(a[0]) - todayMonth + 12) % 12);
        const mb = ((Number(b[0]) - todayMonth + 12) % 12);
        return ma - mb;
      }).map(([m, items]) => section(monthNames[Number(m)], items)).join("");

      return `
        <div>
          <div class="bbz-kpis">
            ${this.kpiBlock("Geburtstage erfasst", covered, `von ${total} aktiven Kontakten`)}
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Erfassungsquote</div>
              <div class="bbz-kpi-value" style="${pctStyle}">${pct}%</div>
              <div class="bbz-kpi-meta">${pct >= 80 ? "Sehr gut ✓" : pct >= 50 ? "Ausbaufähig" : "Viele fehlen noch"}</div>
            </div>
            ${this.kpiBlock("Heute", gToday.length, gToday.length > 0 ? gToday.map(b => helpers.escapeHtml(b.contact.fullName || b.contact.nachname)).join(", ") : "—", gToday.length > 0 ? "ok" : "")}
            ${this.kpiBlock("Diese Woche", gWeek.length, gWeek.length > 0 ? "in den nächsten 7 Tagen" : "keine", "")}
          </div>
          <div style="margin-bottom:10px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;">
            <button class="bbz-button bbz-button-secondary" data-action="kpi-filter" data-scope="navigate" data-value="firms">← Zurück zum Cockpit</button>
            <span style="font-size:12px;color:var(--muted);">${allYearBdays.length} Geburtstage total · sortiert nach nächstem Datum</span>
          </div>
          ${gToday.length ? section("Heute 🎂", gToday) : ""}
          ${gWeek.length  ? section("Diese Woche", gWeek) : ""}
          ${g30.length    ? section("Nächste 30 Tage", g30) : ""}
          ${restByMonthHtml}
          ${!allYearBdays.length ? `<section class="bbz-section"><div class="bbz-section-body"><div class="bbz-empty">Keine Geburtstage erfasst. Trage Geburtstage in den Kontakt-Stammdaten ein.</div></div></section>` : ""}
        </div>
      `;
    },

    // ── ADMIN PANEL ─────────────────────────────────────────────────────────
    // Zugang: URL-Hash #admin oder Doppelklick auf Auth-Status oben rechts
    // Nicht in der Navigation sichtbar — nur für Administratoren gedacht
    adminPanel() {
      if (!state.auth.isAuthenticated) {
        return `<section class="bbz-section"><div class="bbz-section-body"><div class="bbz-empty">Bitte zuerst anmelden.</div></div></section>`;
      }

      // ── Zeitfilter aus State lesen (default: 30 Tage) ────────────────────
      const adminFilter = state.filters.admin || { zeitfenster: "30" };
      const zf = adminFilter.zeitfenster || "30";
      const now = new Date();
      let cutoff = null;
      if (zf !== "all") {
        cutoff = new Date(now);
        cutoff.setDate(cutoff.getDate() - Number(zf));
      }
      const inWindow = (ts) => !cutoff || (ts && new Date(ts) >= cutoff);

      const contacts  = state.enriched.contacts;
      const history   = state.enriched.history;
      const tasks     = state.enriched.tasks;

      // ── Hilfsfunktion: User-Statistiken — gefiltert nach Zeitfenster ─────
      // nameField: welches Feld den User-Namen enthält
      // tsField:   welches Feld den Zeitstempel enthält (für Fenster-Filter)
      const userStats = (items, nameField, tsField) => {
        const map = new Map();
        for (const item of items) {
          const ts = item[tsField] || "";
          if (!inWindow(ts)) continue;
          const user = item[nameField] || "Unbekannt";
          if (!map.has(user)) map.set(user, { count: 0, latest: "" });
          const entry = map.get(user);
          entry.count++;
          if (ts && (!entry.latest || ts > entry.latest)) entry.latest = ts;
        }
        return [...map.entries()]
          .map(([name, d]) => ({ name, count: d.count, latest: d.latest }))
          .sort((a, b) => b.count - a.count);
      };

      // ── Alle bekannten User ermitteln (Union aus allen Erfassern) ─────────
      const allKnownUsers = new Set();
      for (const c of contacts)  { if (c.spCreatedBy) allKnownUsers.add(c.spCreatedBy); }
      for (const h of history)   { if (h.spCreatedBy) allKnownUsers.add(h.spCreatedBy); }
      for (const t of tasks)     { if (t.spCreatedBy) allKnownUsers.add(t.spCreatedBy); }

      // ── Statistiken im Zeitfenster ────────────────────────────────────────
      const contactCreations = userStats(contacts, "spCreatedBy", "spCreated");
      const contactMutations = userStats(contacts, "spModifiedBy", "spModified");
      const historyCreations = userStats(history,  "spCreatedBy", "spCreated");
      const taskCreations    = userStats(tasks,    "spCreatedBy", "spCreated");

      // User die im Zeitfenster KEINE Aktivität haben
      const activeUsersInWindow = new Set([
        ...contactCreations.map(u => u.name),
        ...historyCreations.map(u => u.name),
        ...taskCreations.map(u => u.name)
      ]);
      const inactiveUsers = [...allKnownUsers].filter(u => !activeUsersInWindow.has(u)).sort();

      // ── Aktivitäten-Statistik nach Datum-Feld (History-Datum) ─────────────
      // Hier wird das CRM-Datum (nicht SP-Metadatum) gefiltert
      const activityByUser = new Map();
      for (const h of history) {
        if (cutoff) {
          const d = helpers.toDate(h.datum);
          if (!d || d < cutoff) continue;
        }
        const user = h.spCreatedBy || "Unbekannt";
        if (!activityByUser.has(user)) activityByUser.set(user, { count: 0, latest: "" });
        const e = activityByUser.get(user);
        e.count++;
        if (h.datum && (!e.latest || h.datum > e.latest)) e.latest = h.datum;
      }
      const activityStats = [...activityByUser.entries()]
        .map(([name, d]) => ({ name, count: d.count, latest: d.latest }))
        .sort((a, b) => b.count - a.count);

      // ── Datenqualität (immer über alle Daten, kein Zeitfilter) ────────────
      const activeContacts  = contacts.filter(c => !c.archiviert);
      const missingEmail    = activeContacts.filter(c => !c.email1 && !c.email2);
      const missingPhone    = activeContacts.filter(c => !c.direktwahl && !c.mobile);
      const missingFunktion = activeContacts.filter(c => !c.funktion);
      const missingLeadbbz  = activeContacts.filter(c => !c.leadbbz0);
      const missingSgf      = activeContacts.filter(c => !c.sgf || c.sgf.length === 0);

      // ── Detail-Tabellen: letzte N Items im Zeitfenster ────────────────────
      const recentContacts = [...contacts]
        .filter(c => c.spCreated && inWindow(c.spCreated))
        .sort((a, b) => (b.spCreated > a.spCreated ? 1 : -1))
        .slice(0, 50);
      const recentMutations = [...contacts]
        .filter(c => c.spModified && c.spModified !== c.spCreated && inWindow(c.spModified))
        .sort((a, b) => (b.spModified > a.spModified ? 1 : -1))
        .slice(0, 50);
      const recentHistory = [...history]
        .filter(h => h.spCreated && inWindow(h.spCreated))
        .sort((a, b) => (b.spCreated > a.spCreated ? 1 : -1))
        .slice(0, 50);
      const recentTasks = [...tasks]
        .filter(t => t.spCreated && inWindow(t.spCreated))
        .sort((a, b) => (b.spCreated > a.spCreated ? 1 : -1))
        .slice(0, 50);

      // ── Render-Helfer ─────────────────────────────────────────────────────
      const adminTable = (headers, rows, emptyText = "Keine Daten im gewählten Zeitfenster.") => {
        if (!rows.length) return `<div class="bbz-empty">${emptyText}</div>`;
        return `
          <div class="bbz-table-wrap" style="display:block;overflow-x:auto;">
            <table class="bbz-table" style="width:100%;font-size:12px;">
              <thead><tr>${headers.map(h => `<th style="text-align:left;padding:6px 10px;background:var(--panel-2);border-bottom:1px solid var(--line);font-weight:600;white-space:nowrap;">${helpers.escapeHtml(h)}</th>`).join("")}</tr></thead>
              <tbody>${rows.map((cells, i) => `<tr style="background:${i%2===0?"var(--panel)":"var(--panel-2)"};">${cells.map(c => `<td style="padding:5px 10px;border-bottom:1px solid var(--line-2);white-space:nowrap;">${c}</td>`).join("")}</tr>`).join("")}</tbody>
            </table>
          </div>`;
      };

      const dqRow = (label, items, fieldHint) => {
        const pct = activeContacts.length > 0 ? Math.round((items.length / activeContacts.length) * 100) : 0;
        const color = pct === 0 ? "var(--green)" : pct < 20 ? "var(--amber)" : "var(--red)";
        return `
          <div style="display:flex;align-items:center;gap:10px;padding:7px 0;border-bottom:1px solid var(--line-2);">
            <div style="flex:1;font-size:13px;">${helpers.escapeHtml(label)}</div>
            <div style="font-size:11px;color:var(--muted);">${helpers.escapeHtml(fieldHint)}</div>
            <div style="font-weight:700;color:${color};min-width:32px;text-align:right;">${items.length}</div>
            <div style="font-size:11px;color:var(--muted);min-width:36px;text-align:right;">${pct}%</div>
          </div>`;
      };

      const sec = (title, subtitle, body) => `
        <section class="bbz-section" style="margin-bottom:12px;">
          <div class="bbz-section-header">
            <div>
              <div class="bbz-section-title">${helpers.escapeHtml(title)}</div>
              ${subtitle ? `<div class="bbz-section-subtitle">${helpers.escapeHtml(subtitle)}</div>` : ""}
            </div>
          </div>
          <div class="bbz-section-body" style="padding:10px 14px 12px;">${body}</div>
        </section>`;

      // Zeitfenster-Labels
      const zfLabels = { "7": "letzte 7 Tage", "30": "letzte 30 Tage", "90": "letzte 90 Tage", "all": "gesamte Zeit" };
      const zfLabel  = zfLabels[zf] || zf;

      // Zeitfilter-Buttons HTML
      const zfBtn = (val, label) => {
        const active = zf === val;
        return `<button style="padding:4px 12px;border-radius:var(--r-full);border:1px solid ${active?"var(--blue)":"var(--line)"};background:${active?"var(--blue)":"var(--panel)"};color:${active?"#fff":"var(--text)"};font-size:12px;font-weight:${active?"700":"400"};cursor:pointer;font-family:inherit;" data-action="admin-zeitfilter" data-zf="${val}">${helpers.escapeHtml(label)}</button>`;
      };

      const zeitfilterBar = `
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:14px;padding:10px 14px;background:var(--panel);border:1px solid var(--line);border-radius:var(--r-lg);">
          <span style="font-size:12px;font-weight:600;color:var(--muted);margin-right:4px;">Zeitfenster:</span>
          ${zfBtn("7",   "7 Tage")}
          ${zfBtn("30",  "30 Tage")}
          ${zfBtn("90",  "90 Tage")}
          ${zfBtn("all", "Alles")}
          <span style="font-size:11px;color:var(--muted);margin-left:8px;">— Statistiken und Detail-Tabellen gefiltert auf: <strong>${zfLabel}</strong></span>
        </div>`;

      return `
        <div>
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:8px;">
            <div>
              <div style="font-size:16px;font-weight:700;color:var(--blue);">⚙ Admin-Panel</div>
              <div style="font-size:11px;color:var(--muted);">Nutzungsstatistiken · Datenqualität · Aktivitäten — nur für Administratoren</div>
            </div>
            <button class="bbz-button bbz-button-secondary" data-action="kpi-filter" data-scope="navigate" data-value="firms">← Zurück</button>
          </div>

          ${zeitfilterBar}

          <div class="bbz-kpis" style="margin-bottom:12px;">
            ${this.kpiBlock("Firmen", state.enriched.firms.length, "")}
            ${this.kpiBlock("Kontakte (aktiv)", activeContacts.length, `${contacts.filter(c=>c.archiviert).length} archiviert`)}
            ${this.kpiBlock("Aktivitäten", history.length, "")}
            ${this.kpiBlock("Aufgaben", tasks.length, `${tasks.filter(t=>t.isOpen).length} offen`)}
          </div>

          ${sec(`Kontakterfassungen pro User — ${zfLabel}`, `${recentContacts.length} Erfassungen im Zeitfenster`,
            adminTable(
              ["User", "Erfassungen", "Letzte Erfassung"],
              contactCreations.map(u => [
                helpers.escapeHtml(u.name),
                String(u.count),
                u.latest ? helpers.formatDateTime(u.latest) : "—"
              ])
            )
          )}

          ${sec(`Kontaktmutationen pro User — ${zfLabel}`, "Letzte Änderung pro User",
            adminTable(
              ["User", "Mutationen", "Letzte Mutation"],
              contactMutations.map(u => [
                helpers.escapeHtml(u.name),
                String(u.count),
                u.latest ? helpers.formatDateTime(u.latest) : "—"
              ])
            )
          )}

          ${sec(`Aktivitäten pro User — ${zfLabel}`, `${recentHistory.length} Aktivitäten im Zeitfenster (nach Erfassungszeitpunkt)`,
            adminTable(
              ["User", "Aktivitäten erfasst", "Letzte Erfassung"],
              historyCreations.map(u => [
                helpers.escapeHtml(u.name),
                String(u.count),
                u.latest ? helpers.formatDateTime(u.latest) : "—"
              ])
            )
          )}

          ${sec(`Aktivitäten nach Datum-Feld — ${zfLabel}`, "Ausgewertet nach dem CRM-Aktivitätsdatum (nicht Erfassungszeitpunkt)",
            adminTable(
              ["User", "Aktivitäten", "Letztes Datum"],
              activityStats.map(u => [
                helpers.escapeHtml(u.name),
                String(u.count),
                u.latest ? helpers.formatDate(u.latest) : "—"
              ])
            )
          )}

          ${sec(`Aufgaben pro User — ${zfLabel}`, `${recentTasks.length} Tasks im Zeitfenster`,
            adminTable(
              ["User", "Aufgaben erfasst", "Letzte Erfassung"],
              taskCreations.map(u => [
                helpers.escapeHtml(u.name),
                String(u.count),
                u.latest ? helpers.formatDateTime(u.latest) : "—"
              ])
            )
          )}

          ${inactiveUsers.length > 0 ? sec(
            `Inaktiv im Zeitfenster (${zfLabel})`,
            `${inactiveUsers.length} bekannte User ohne Erfassungen oder Aktivitäten`,
            `<div style="display:flex;flex-wrap:wrap;gap:8px;padding:4px 0;">
              ${inactiveUsers.map(u => `<span class="bbz-chip" style="background:var(--red-light);color:var(--red);">${helpers.escapeHtml(u)}</span>`).join("")}
            </div>`
          ) : sec(
            `Inaktiv im Zeitfenster (${zfLabel})`,
            "Alle bekannten User haben im Zeitfenster etwas erfasst",
            `<div style="color:var(--green);font-size:13px;padding:4px 0;">✓ Keine inaktiven User</div>`
          )}

          ${sec("Datenqualität — fehlende Felder", `Aktive Kontakte: ${activeContacts.length} (kein Zeitfilter)`,
            `<div style="margin-bottom:6px;font-size:11px;color:var(--muted);">Zeigt Anzahl aktiver Kontakte mit fehlendem Feld — 0 ist das Ziel.</div>
            ${dqRow("Keine E-Mail-Adresse", missingEmail, "Email1 / Email2")}
            ${dqRow("Keine Telefonnummer", missingPhone, "Direktwahl / Mobile")}
            ${dqRow("Keine Funktion/Rolle", missingFunktion, "Funktion")}
            ${dqRow("Kein Lead BBZ", missingLeadbbz, "Leadbbz0")}
            ${dqRow("Kein SGF", missingSgf, "SGF (Multi-Choice)")}`
          )}

          ${sec(`Letzte Kontakterfassungen — ${zfLabel}`, `${recentContacts.length} Einträge`,
            adminTable(
              ["Kontakt", "Firma", "Erfasst von", "Datum"],
              recentContacts.map(c => [
                `<a class="bbz-link" data-action="open-contact" data-id="${c.id}">${helpers.escapeHtml(c.fullName || c.nachname)}</a>`,
                helpers.escapeHtml(c.firmTitle || "—"),
                helpers.escapeHtml(c.spCreatedBy || "—"),
                c.spCreated ? helpers.formatDateTime(c.spCreated) : "—"
              ])
            )
          )}

          ${sec(`Letzte Kontaktmutationen — ${zfLabel}`, `${recentMutations.length} Einträge — nur Items mit Mutation nach Erfassung`,
            adminTable(
              ["Kontakt", "Firma", "Mutiert von", "Datum"],
              recentMutations.map(c => [
                `<a class="bbz-link" data-action="open-contact" data-id="${c.id}">${helpers.escapeHtml(c.fullName || c.nachname)}</a>`,
                helpers.escapeHtml(c.firmTitle || "—"),
                helpers.escapeHtml(c.spModifiedBy || "—"),
                c.spModified ? helpers.formatDateTime(c.spModified) : "—"
              ])
            )
          )}

          ${sec(`Letzte Aktivitäten (History) — ${zfLabel}`, `${recentHistory.length} Einträge — nach Erfassungszeitpunkt`,
            adminTable(
              ["Titel", "Kontakt", "Typ", "Erfasst von", "Erfasst am"],
              recentHistory.map(h => [
                helpers.escapeHtml(h.title || "—"),
                helpers.escapeHtml(h.contactName || "—"),
                helpers.escapeHtml(h.typ || "—"),
                helpers.escapeHtml(h.spCreatedBy || "—"),
                h.spCreated ? helpers.formatDateTime(h.spCreated) : "—"
              ])
            )
          )}

          ${sec(`Letzte Aufgaben — ${zfLabel}`, `${recentTasks.length} Einträge — nach Erfassungszeitpunkt`,
            adminTable(
              ["Titel", "Kontakt", "Status", "Erfasst von", "Erfasst am"],
              recentTasks.map(t => [
                helpers.escapeHtml(t.title || "—"),
                helpers.escapeHtml(t.contactName || "—"),
                helpers.escapeHtml(t.status || "—"),
                helpers.escapeHtml(t.spCreatedBy || "—"),
                t.spCreated ? helpers.formatDateTime(t.spCreated) : "—"
              ])
            )
          )}

          <div style="font-size:11px;color:var(--muted);text-align:center;padding:12px 0 4px;">
            bbz CRM Admin-Panel · Daten aus SharePoint · Stand: ${new Date().toLocaleString("de-CH")}
          </div>
        </div>
      `;
    },


    firms() {
      const filters = state.filters.firms;
      const isKunde = filters.kategorie === "Kunde";
      const filteredFirms = state.enriched.firms.filter(firm => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch = !search || [firm.title, firm.ort, firm.klassifizierung, firm.hauptnummer, firm.adresse, firm.land, ...firm.contacts.map(c => c.fullName)].some(v => helpers.textIncludes(v, search));
        const kategorieMatch = firm.kategorie === filters.kategorie;
        // Stufe-2-Filter (Klassifizierung/VIP) greifen nur bei Kunden
        const klassMatch = !isKunde || !filters.klassifizierung || String(firm.klassifizierung || "").toUpperCase().startsWith(filters.klassifizierung.toUpperCase());
        const vipMatch = !isKunde || !filters.vip || firm.vip;
        return searchMatch && kategorieMatch && klassMatch && vipMatch;
      });
      const firmSortDir = filters.sortDir === "asc" ? 1 : -1;
      const rows = [...filteredFirms].sort((a, b) => {
        if (filters.sortBy === "title")          return a.title.localeCompare(b.title, "de") * firmSortDir;
        if (filters.sortBy === "klassifizierung") return String(a.klassifizierung||"").localeCompare(String(b.klassifizierung||""), "de") * firmSortDir;
        if (filters.sortBy === "vip")            return ((b.vip ? 1 : 0) - (a.vip ? 1 : 0)) * firmSortDir;
        if (filters.sortBy === "openTasksCount") return (a.openTasksCount - b.openTasksCount) * firmSortDir;
        if (filters.sortBy === "status") {
          const rank = f => f.tasks.some(t => t.isOpen && t.isOverdue) ? 0 : f.openTasksCount > 0 ? 1 : (helpers.toDate(f.latestActivity) ? 2 : 3);
          const ra = rank(a), rb = rank(b);
          if (ra !== rb) return (ra - rb) * firmSortDir;
          if (ra === 0) return helpers.compareDateAsc(a.nextDeadline, b.nextDeadline) * firmSortDir; // ältester überfälliger zuerst
          if (ra === 2) return helpers.compareDateAsc(a.latestActivity, b.latestActivity) * firmSortDir; // ältestes Aktivitätsalter zuerst
          return 0;
        }
        return 0;
      });

      // Für Task-Badge in der Mobile-Card
      const overdueTasks = state.enriched.tasks.filter(t => t.isOpen && t.isOverdue);
      // Kategorie-Zähler für Stufe-1-Chips
      const katCount = k => state.enriched.firms.filter(f => f.kategorie === k).length;

      return `
        <div>
          <section class="bbz-section">
            <div class="bbz-section-header">
              <div>
                <div class="bbz-section-title">Firmen</div>
                <div class="bbz-section-subtitle">${filteredFirms.length} ${filteredFirms.length === 1 ? "Firma" : "Firmen"} in dieser Ansicht</div>
              </div>
              <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">
                <button class="bbz-button bbz-button-primary" data-action="open-firm-form">+ Firma</button>
              </div>
            </div>
            <div class="bbz-section-body">
              <!-- Stufe 1: Kategorie -->
              <div class="bbz-kpi-chips" style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px;">
                ${[["Kunde","Kunden"],["Lieferant","Lieferanten"],["Übrige","Übrige"]].map(([val,label]) =>
                  `<button class="bbz-kpi-chip ${filters.kategorie === val ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="firms-kategorie" data-value="${val}">${label} <span>${katCount(val)}</span></button>`
                ).join("")}
              </div>
              ${isKunde ? `
              <!-- Stufe 2: Klassifizierung + VIP (nur bei Kunden) -->
              <div class="bbz-kpi-chips" style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-bottom:10px;">
                <button class="bbz-kpi-chip ${!filters.klassifizierung ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="firms-klassifizierung" data-value="">Alle</button>
                ${["A","B","C"].map(k => {
                  const cnt = state.enriched.firms.filter(f => f.kategorie === "Kunde" && String(f.klassifizierung||"").toUpperCase().startsWith(k)).length;
                  return `<button class="bbz-kpi-chip ${filters.klassifizierung === k ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="firms-klassifizierung" data-value="${k}">${k} <span>${cnt}</span></button>`;
                }).join("")}
                <span style="width:1px;height:20px;background:var(--line);margin:0 4px;"></span>
                <button class="bbz-kpi-chip ${filters.vip ? "bbz-kpi-chip-active-gold" : ""}" data-action="kpi-filter" data-scope="firms-vip" data-value="yes">♛ VIP <span>${state.enriched.firms.filter(f => f.kategorie === "Kunde" && f.vip).length}</span></button>
              </div>` : ""}
              <div style="margin-bottom:10px;">
                <input class="bbz-input" style="width:100%;height:40px;font-size:14px;" data-filter="firms-search" type="text"
                  placeholder="Suche nach Firma, Ort, Ansprechpartner ..."
                  value="${helpers.escapeHtml(filters.search)}" />
              </div>
              <!-- Signal-Legende (aufklappbar) -->
              <div style="margin-bottom:10px;">
                <button style="background:none;border:none;padding:0;color:var(--blue);font-size:12px;font-weight:600;cursor:pointer;" data-action="toggle-firm-legende">${filters.legendeOffen ? "▾" : "▸"} Was bedeuten die Punkte?</button>
                ${filters.legendeOffen ? `
                <div style="margin-top:8px;padding:12px 14px;border:1px solid var(--line);border-radius:var(--r-md);background:var(--panel-2);font-size:13px;line-height:1.5;">
                  <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;"><span class="bbz-signal bbz-signal-green"></span><span><strong>Aktiv gepflegt</strong> — Kontakt jünger als 12 Monate, keine offene Frist</span></div>
                  <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;"><span class="bbz-signal bbz-signal-amber"></span><span><strong>Aufmerksamkeit</strong> — kein Kontakt seit über 12 Monaten</span></div>
                  <div style="display:flex;align-items:center;gap:8px;"><span class="bbz-signal bbz-signal-red"></span><span><strong>Nicht aktiv gepflegt</strong> — überfällige Task oder nie kontaktiert</span></div>
                  <div style="color:var(--muted);font-size:12px;margin-top:10px;">Lieferanten und Übrige tragen keinen Punkt.</div>
                </div>` : ""}
              </div>
              <div class="bbz-table-wrap">
                <table class="bbz-table">
                  <thead><tr>
                    ${(()=>{
                      const firmSortTh = (label, col) => {
                        const active = filters.sortBy === col;
                        const icon = active ? (filters.sortDir === "asc" ? " ↑" : " ↓") : "";
                        return `<th style="cursor:pointer;user-select:none;${active?"color:var(--blue);":""}" data-action="set-sort" data-col="${col}" data-scope="firms">${label}${icon}</th>`;
                      };
                      return "<th></th>"
                        + firmSortTh("Firma","title")
                        + "<th>Ort</th>"
                        + firmSortTh("Klassifizierung","klassifizierung")
                        + "<th>Kontakte</th>"
                        + firmSortTh("Status/Aktivität","status");
                    })()}
                  </tr></thead>
                  <tbody>
                    ${rows.length ? rows.map(firm => {
                      const signal = helpers.firmSignal(firm);
                      const signalDot = signal === "overdue"
                        ? `<span class="bbz-signal bbz-signal-red" title="Überfällige Task(s)"></span>`
                        : signal === "never"
                        ? `<span class="bbz-signal bbz-signal-red" title="Noch nie kontaktiert"></span>`
                        : signal === "cold"
                        ? `<span class="bbz-signal bbz-signal-amber" title="Kein Kontakt seit über 12 Monaten"></span>`
                        : signal === "ok"
                        ? `<span class="bbz-signal bbz-signal-green" title="Aktiv"></span>`
                        : `<span class="bbz-signal bbz-signal-none"></span>`;
                      const rowClass = signal === "overdue" ? "bbz-row-alert"
                        : (signal === "cold" || signal === "never") ? "bbz-row-cold"
                        : signal === "ok" ? "bbz-row-ok"
                        : "";
                      return `
                      <tr class="${rowClass}">
                        <td style="width:28px;padding-right:4px;">${signalDot}</td>
                        <td><a class="bbz-link" data-action="open-firm" data-id="${firm.id}">${helpers.escapeHtml(firm.title)}</a><div class="bbz-subtext">${helpers.escapeHtml(firm.hauptnummer || "—")}</div></td>
                        <td>${helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ")) || '<span class="bbz-muted">—</span>'}</td>
                        <td>${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : '<span class="bbz-muted">—</span>'}</td>
                        <td>${firm.contactsCount}</td>
                        <td>${(() => {
                          const overdue = firm.tasks.filter(t => t.isOpen && t.isOverdue);
                          if (overdue.length) {
                            const oldest = [...overdue].sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline))[0];
                            return `<span class="bbz-danger">seit ${helpers.agePhrase(oldest.deadline)} fällig</span>`;
                          }
                          if (firm.openTasksCount > 0) return "laufender Task";
                          const last = helpers.toDate(firm.latestActivity);
                          if (last) {
                            const today = helpers.todayStart();
                            const months = (today.getFullYear() - last.getFullYear()) * 12 + (today.getMonth() - last.getMonth());
                            if (months <= 24) return `<span title="${helpers.formatDate(firm.latestActivity)}">vor ${helpers.agePhrase(firm.latestActivity)}</span>`;
                          }
                          return `<span class="bbz-muted">keine aktuellen Aktivitäten</span>`;
                        })()}</td>
                      </tr>`; }).join("") : `<tr><td colspan="6">${ui.emptyBlock("Keine Firmen für die aktuelle Filterung gefunden.")}</td></tr>`}
                  </tbody>
                </table>
              </div>
              <!-- Mobile Card List (nur sichtbar auf kleinen Screens via CSS) -->
              <div class="bbz-card-list bbz-mobile-only">
                ${rows.length ? rows.map(firm => {
                  const signal = helpers.firmSignal(firm);
                  const sigDot = signal === "overdue"
                    ? `<span class="bbz-signal bbz-signal-red" title="Überfällige Task(s)"></span>`
                    : signal === "never"
                    ? `<span class="bbz-signal bbz-signal-red" title="Noch nie kontaktiert"></span>`
                    : signal === "cold"
                    ? `<span class="bbz-signal bbz-signal-amber" title="Kein Kontakt seit über 12 Monaten"></span>`
                    : signal === "ok"
                    ? `<span class="bbz-signal bbz-signal-green" title="Aktiv"></span>`
                    : `<span style="width:8px;flex-shrink:0;display:inline-block;"></span>`;
                  const taskBadge = firm.openTasksCount > 0
                    ? overdueTasks.some(t => t.firmId === firm.id)
                      ? `<span class="bbz-status-chip bbz-status-overdue">${firm.openTasksCount} überfällig</span>`
                      : `<span class="bbz-status-chip bbz-status-open">${firm.openTasksCount} offen</span>`
                    : "";
                  return `<div class="bbz-list-card" data-action="open-firm" data-id="${firm.id}">
                    ${sigDot}
                    <div class="bbz-list-card-body">
                      <div class="bbz-list-card-title">${helpers.escapeHtml(firm.title)}</div>
                      <div class="bbz-list-card-sub">${helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ") || "")}${firm.latestActivity ? " · " + helpers.relativeDate(firm.latestActivity) : ""}</div>
                    </div>
                    <div class="bbz-list-card-right">
                      ${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : ""}
                      ${taskBadge}
                    </div>
                  </div>`;
                }).join("") : ui.emptyBlock("Keine Firmen gefunden.")}
              </div>
            </div>
          </section>
        </div>
      `;
    },

    firmDetail() {
      const firm = dataModel.getFirmById(state.selection.firmId);
      if (!firm) return ui.emptyBlock("Die ausgewaehlte Firma wurde nicht gefunden.");
      const recentHistory = [...firm.history].slice(0, 20);
      const bandClass = helpers.detailBandClass(firm);
      return `
        <div>
          <div class="${bandClass}">
            <div class="bbz-detail-header" style="margin-bottom:0;">
              <div>
                <button class="bbz-button bbz-button-secondary" style="margin-bottom:12px;background:rgba(255,255,255,0.7);" data-action="back-to-firms">← Firmenliste</button>
                <div class="bbz-detail-title">${helpers.escapeHtml(firm.title)}</div>
                <div class="bbz-detail-subtitle">${helpers.escapeHtml(helpers.joinNonEmpty([firm.adresse, helpers.joinNonEmpty([firm.plz, firm.ort], " "), firm.land], " · ")) || "Keine Adresse erfasst"}</div>
                <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:12px;">
                  ${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : ""}
                  ${firm.vip ? `<span class="bbz-pill bbz-pill-vip">♛</span>` : ""}
                  ${(() => {
                    const firmContacts = firm.contacts.filter(c => !c.archiviert);
                    const bdays = helpers.upcomingBirthdays(14, firmContacts);
                    if (!bdays.length) return "";
                    const todayCount = bdays.filter(b => b.daysUntil === 0).length;
                    const label = todayCount > 0
                      ? (bdays.length === 1 ? `Geburtstag heute` : `${bdays.length} Geburtstage, ${todayCount} heute`)
                      : (bdays.length === 1 ? `Geburtstag in ${bdays[0].daysUntil} ${bdays[0].daysUntil === 1 ? "Tag" : "Tagen"}` : `${bdays.length} Geburtstage bald`);
                    const pillStyle = todayCount > 0
                      ? "background:rgba(255,255,255,0.22);border:0.5px solid rgba(255,255,255,0.45);color:#fff;"
                      : "background:rgba(255,200,80,0.18);border:0.5px solid rgba(255,200,80,0.45);color:#ffe08a;";
                    return `<span style="display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:999px;font-size:12px;font-weight:500;${pillStyle}">🎂 ${helpers.escapeHtml(label)}</span>`;
                  })()}
                </div>
              </div>
              <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                <button class="bbz-button bbz-button-secondary" style="${firm.contactsCount > 0 ? "opacity:0.4;cursor:not-allowed;" : "color:var(--red);border-color:var(--red);"}" data-action="delete-firm" data-id="${firm.id}" data-name="${helpers.escapeHtml(firm.title)}" data-contacts="${firm.contactsCount}">Löschen</button>
                <button class="bbz-button bbz-button-secondary" data-action="open-firm-form" data-id="${firm.id}">Bearbeiten</button>
                <button class="bbz-button bbz-button-secondary" data-action="open-task-form" data-firm-id="${firm.id}">+ Task</button>
                <button class="bbz-button bbz-button-secondary" data-action="open-history-form" data-firm-id="${firm.id}">+ Aktivität</button>
                <button class="bbz-button bbz-button-primary" data-action="open-contact-form" data-firm-id="${firm.id}">+ Kontakt</button>
              </div>
            </div>
          </div>
          <div class="bbz-kpis" style="margin-top:10px;">
            ${this.kpiBlock("Kontakte", firm.contactsCount)}
            ${this.kpiBlock("Offene Tasks", firm.openTasksCount, firm.tasks.some(t => t.isOpen && t.isOverdue) ? "überfällig" : firm.openTasksCount > 0 ? "offen" : "keine offen", firm.tasks.some(t => t.isOpen && t.isOverdue) ? "alert" : "")}
            ${this.kpiBlock("Nächste Deadline", firm.nextDeadline ? helpers.relativeDate(firm.nextDeadline) : "—", firm.nextDeadline && helpers.isOverdue(firm.nextDeadline) ? "überfällig" : "", firm.nextDeadline && helpers.isOverdue(firm.nextDeadline) ? "alert" : "")}
            ${this.kpiBlock("Aktivitäten", firm.history.length, firm.latestActivity ? helpers.relativeDate(firm.latestActivity) : "noch keine")}
          </div>
          <div class="bbz-grid bbz-grid-3">
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Stammdaten</div></div>
              <div class="bbz-section-body">
                ${ui.kv("Klassifizierung", firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : '<span class="bbz-muted">—</span>')}
                ${ui.kv("VIP", firm.vip ? '<span class="bbz-pill bbz-pill-vip">♛</span>' : '<span class="bbz-muted">Nein</span>')}
                ${ui.kv("Adresse", helpers.escapeHtml(firm.adresse) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("PLZ / Ort", helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ")) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Land", helpers.escapeHtml(firm.land) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Hauptnummer", helpers.escapeHtml(firm.hauptnummer) || '<span class="bbz-muted">—</span>')}
              </div>
            </section>
            <section class="bbz-section" style="grid-column: span 2;">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Kontakte</div><div class="bbz-section-subtitle">Ansprechpartner dieser Firma</div></div></div>
              <div class="bbz-section-body">
                <!-- Desktop: Tabelle -->
                <div class="bbz-table-wrap bbz-desktop-only">
                  <table class="bbz-table">
                    <thead><tr><th></th><th>Name</th><th>Funktion</th><th>Rolle</th><th>E-Mail</th><th>Telefon</th><th>Geburtstag</th></tr></thead>
                    <tbody>
                      ${firm.contacts.length ? firm.contacts.map(c => {
                        const bdays = helpers.upcomingBirthdays(30, [c]);
                        const b = bdays[0] || null;
                        const bdayHtml = b
                          ? (() => {
                              const lbl = helpers.birthdayLabel(b.daysUntil, b.nextBirthday);
                              const s = b.daysUntil === 0
                                ? "background:var(--blue-light);border-color:#a8c8e0;color:var(--blue);"
                                : b.daysUntil <= 7
                                ? "background:#fff9eb;border-color:#f4dfab;color:var(--amber);"
                                : "";
                              return `<span class="bbz-chip" style="${s}">${helpers.escapeHtml(lbl)}</span>`;
                            })()
                          : (c.geburtstag
                              ? `<span class="bbz-muted" style="font-size:11px;">${helpers.formatDate(c.geburtstag)}</span>`
                              : '<span class="bbz-muted">—</span>');
                        return `
                        <tr>
                          <td style="width:36px;padding-right:0;">${helpers.avatarHtml(c)}</td>
                          <td><a class="bbz-link" data-action="open-contact" data-id="${c.id}">${helpers.escapeHtml(c.fullName || c.nachname)}</a>${c.archiviert ? ' <span class="bbz-muted" style="font-size:11px;">(archiviert)</span>' : ""}</td>
                          <td>${helpers.escapeHtml(c.funktion) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${helpers.escapeHtml(c.rolle) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${c.email1 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(c.email1)}">${helpers.escapeHtml(c.email1)}</a>` : '<span class="bbz-muted">—</span>'}</td>
                          <td>${helpers.escapeHtml(helpers.joinNonEmpty([c.direktwahl, c.mobile], " / ")) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${bdayHtml}</td>
                        </tr>`}).join("")
                      : `<tr><td colspan="7">${ui.emptyBlock("Keine Kontakte vorhanden.", "open-contact-form", "+ Ersten Kontakt hinzufügen")}</td></tr>`}
                    </tbody>
                  </table>
                </div>
                <!-- Mobile: Cards -->
                <div class="bbz-card-list bbz-mobile-only">
                  ${firm.contacts.length ? firm.contacts.map(c => {
                    const bdays = helpers.upcomingBirthdays(30, [c]);
                    const b = bdays[0] || null;
                    const bdayBadge = b
                      ? (() => {
                          const lbl = helpers.birthdayLabel(b.daysUntil, b.nextBirthday);
                          const s = b.daysUntil === 0
                            ? "background:var(--blue-light);border-color:#a8c8e0;color:var(--blue);"
                            : "background:#fff9eb;border-color:#f4dfab;color:var(--amber);";
                          return `<span class="bbz-chip" style="font-size:10px;${s}">🎂 ${helpers.escapeHtml(lbl)}</span>`;
                        })()
                      : "";
                    return `
                    <div class="bbz-list-card" data-action="open-contact" data-id="${c.id}">
                      ${helpers.avatarHtml(c)}
                      <div class="bbz-list-card-body">
                        <div class="bbz-list-card-title">${helpers.escapeHtml(c.fullName || c.nachname)}${c.archiviert ? ' <span class="bbz-muted" style="font-size:10px;">(archiviert)</span>' : ""}</div>
                        <div class="bbz-list-card-sub">${helpers.escapeHtml(helpers.joinNonEmpty([c.funktion, c.rolle], " · ")) || "—"}</div>
                      </div>
                      <div class="bbz-list-card-right">
                        ${bdayBadge || (c.email1 ? `<span style="font-size:10px;color:var(--subtle);">${helpers.escapeHtml(c.email1)}</span>` : "")}
                      </div>
                    </div>`;
                  }).join("") : ui.emptyBlock("Keine Kontakte vorhanden.", "open-contact-form", "+ Ersten Kontakt hinzufügen")}
                </div>
              </div>
            </section>
          </div>
          <div class="bbz-grid bbz-grid-2" style="margin-top:12px;">
            <section class="bbz-section">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Aktivitäten</div><div class="bbz-section-subtitle">Aggregiert über alle Kontakte</div></div>
                <button class="bbz-button bbz-button-secondary" style="height:32px;font-size:13px;" data-action="open-history-form" data-firm-id="${firm.id}">+ Aktivität</button>
              </div>
              <div class="bbz-section-body">
                ${recentHistory.length ? `<div class="bbz-timeline">${recentHistory.map(h => `
                  <div class="bbz-timeline-item">
                    <div class="bbz-timeline-date">${helpers.relativeDate(h.datum) || "—"}<br><span class="bbz-muted" style="font-size:11px;">${helpers.formatDate(h.datum)}</span><br><span class="bbz-muted">${helpers.escapeHtml(h.contactName || "")}</span></div>
                    <div>
                      <div class="bbz-timeline-title">${helpers.escapeHtml(h.typ || h.title || "Eintrag")} ${h.projektbezugBool ? '<span class="bbz-chip" style="background:var(--blue-light);color:var(--blue);border-color:#a8c8e0;">Projektbezug</span>' : '<span class="bbz-chip">Allgemein</span>'}</div>
                      <div class="bbz-timeline-text">${helpers.escapeHtml(h.notizen || "—")}</div>
                      <div style="margin-top:6px;display:flex;gap:6px;">
                        <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;" data-action="edit-history" data-id="${h.id}">Bearbeiten</button>
                        <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;color:var(--red);border-color:var(--red);" data-action="delete-history" data-id="${h.id}" data-title="${helpers.escapeHtml(h.typ || h.title || 'Eintrag')}">Löschen</button>
                      </div>
                    </div>
                  </div>`).join("")}</div>`
                  : `<div class="bbz-empty">Noch keine Aktivitäten erfasst.<br><button class="bbz-button bbz-button-secondary" style="margin-top:10px;height:32px;font-size:13px;" data-action="open-history-form" data-firm-id="${firm.id}">+ Erste Aktivität erfassen</button></div>`}
              </div>
            </section>
            <section class="bbz-section">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Aufgaben</div></div>
                <button class="bbz-button bbz-button-secondary" style="height:32px;font-size:13px;" data-action="open-task-form" data-firm-id="${firm.id}">+ Task</button>
              </div>
              <div class="bbz-section-body">
                <!-- Desktop: Tabelle -->
                <div class="bbz-table-wrap bbz-desktop-only">
                  <table class="bbz-table">
                    <thead><tr><th>Titel</th><th>Deadline</th><th>Status</th><th>Kontakt</th><th>Aktionen</th></tr></thead>
                    <tbody>
                      ${firm.tasks.length ? firm.tasks.map(t => `
                        <tr>
                          <td>${helpers.escapeHtml(t.title) || '<span class="bbz-muted">—</span>'}</td>
                          <td class="${helpers.isOpenTask(t.status) && helpers.isOverdue(t.deadline) ? "bbz-danger" : ""}">${t.deadline ? helpers.relativeDate(t.deadline) : '<span class="bbz-muted">—</span>'}</td>
                          <td>${helpers.statusChipHtml(t.status, t.deadline)}</td>
                          <td>${t.contactId ? `<a class="bbz-link" data-action="open-contact" data-id="${t.contactId}">${helpers.escapeHtml(t.contactName || "Kontakt")}</a>` : helpers.escapeHtml(t.contactName || "—")}</td>
                          <td style="white-space:nowrap;">
                            <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 8px;margin-right:3px;" data-action="edit-task" data-id="${t.id}">Bearbeiten</button>
                            <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 8px;color:var(--red);border-color:var(--red);" data-action="delete-task" data-id="${t.id}" data-title="${helpers.escapeHtml(t.title)}">Löschen</button>
                          </td>
                        </tr>`).join("") : `<tr><td colspan="5">${ui.emptyBlock("Keine Aufgaben vorhanden.")}</td></tr>`}
                    </tbody>
                  </table>
                </div>
                <!-- Mobile: Card-List -->
                <div class="bbz-mobile-only">
                  ${firm.tasks.length ? `<div class="bbz-card-list">${firm.tasks.map(t => `
                    <div class="bbz-task-card" data-action="edit-task" data-id="${t.id}">
                      <div class="bbz-task-card-top">
                        <span class="bbz-task-card-title">${helpers.escapeHtml(t.title)}</span>
                        ${helpers.statusChipHtml(t.status, t.deadline)}
                      </div>
                      <div class="bbz-task-card-meta">
                        ${t.contactName ? `<span>${helpers.escapeHtml(t.contactName)}</span>` : ""}
                        ${t.deadline ? `<span>·</span><span class="${t.isOpen && t.isOverdue ? "bbz-danger" : ""}">${helpers.relativeDate(t.deadline)}</span>` : ""}
                      </div>
                    </div>`).join("")}</div>`
                  : ui.emptyBlock("Keine Aufgaben vorhanden.")}
                </div>
              </div>
            </section>
          </div>
        </div>
      `;
    },

    contacts() {
      const filters = state.filters.contacts;
      const kpiMode = filters._kpiMode || "all";

      // Contacts mit History / offenen Tasks für KPI-Counts
      const contactsWithHistory = state.enriched.contacts.filter(c => !c.archiviert && state.enriched.history.some(h => h.contactId === c.id));
      const contactsWithOpenTasks = state.enriched.contacts.filter(c => !c.archiviert && state.enriched.tasks.some(t => t.contactId === c.id && t.isOpen));

      const filteredContacts = state.enriched.contacts.filter(c => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch = !search || [c.fullName, c.firmTitle, c.funktion, c.rolle, c.email1, c.email2, c.direktwahl, c.mobile, c.kommentar, ...c.sgf, ...c.event].some(v => helpers.textIncludes(v, search));
        const archivMatch = !filters.archiviertAusblenden || !c.archiviert;
        const modeMatch = kpiMode === "history"
          ? state.enriched.history.some(h => h.contactId === c.id)
          : kpiMode === "tasks"
          ? state.enriched.tasks.some(t => t.contactId === c.id && t.isOpen)
          : true;
        return searchMatch && archivMatch && modeMatch;
      });
      const cSortDir = filters.sortDir === "asc" ? 1 : -1;
      const rows = [...filteredContacts].sort((a, b) => {
        if (filters.sortBy === "fullName")  return String(a.fullName||"").localeCompare(String(b.fullName||""), "de") * cSortDir;
        if (filters.sortBy === "firmTitle") return String(a.firmTitle||"").localeCompare(String(b.firmTitle||""), "de") * cSortDir;
        if (filters.sortBy === "rolle")     return String(a.rolle||"").localeCompare(String(b.rolle||""), "de") * cSortDir;
        if (filters.sortBy === "leadbbz0")  return String(a.leadbbz0||"").localeCompare(String(b.leadbbz0||""), "de") * cSortDir;
        return 0;
      });
      const cTh = (label, col) => {
        const active = filters.sortBy === col;
        const icon = active ? (filters.sortDir === "asc" ? " ↑" : " ↓") : "";
        return `<th style="cursor:pointer;user-select:none;${active?"color:var(--blue);":""}" data-action="set-sort" data-col="${col}" data-scope="contacts">${label}${icon}</th>`;
      };

      // Counts für KPI-Chips
      const totalActive   = state.enriched.contacts.filter(c => !c.archiviert).length;
      const withHistory   = state.enriched.contacts.filter(c => !c.archiviert && state.enriched.history.some(h => h.contactId === c.id)).length;
      const withOpenTasks = state.enriched.contacts.filter(c => !c.archiviert && state.enriched.tasks.some(t => t.contactId === c.id && t.isOpen)).length;
      const allOpenTasks  = state.enriched.tasks.filter(t => t.isOpen).length;
      const overdueTasks  = state.enriched.tasks.filter(t => t.isOpen && t.isOverdue).length;

      return `
        <div>
          <div class="bbz-kpis">
            <!-- Kontakte mit Schnellfilter -->
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Kontakte</div>
              <div class="bbz-kpi-value">${totalActive}</div>
              <div class="bbz-kpi-chips" style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                <button class="bbz-kpi-chip ${kpiMode==="history"?"bbz-kpi-chip-active":""}" data-action="kpi-filter" data-scope="contacts-mode" data-value="history">Mit History <span>${withHistory}</span></button>
                <button class="bbz-kpi-chip ${kpiMode==="tasks"?"bbz-kpi-chip-active":""}" data-action="kpi-filter" data-scope="contacts-mode" data-value="tasks">Offene Tasks <span>${withOpenTasks}</span></button>
                <button class="bbz-kpi-chip ${kpiMode==="all"||!kpiMode?"bbz-kpi-chip-active":""}" data-action="kpi-filter" data-scope="contacts-mode" data-value="all">Alle</button>
              </div>
            </div>
            <!-- Sichtbar nach Filter -->
            ${this.kpiBlock("Angezeigt", rows.length, rows.length < totalActive ? `von ${totalActive}` : "alle aktiven")}
            <!-- Offene Tasks — klickbar zur Planung -->
            <div class="bbz-kpi bbz-kpi-clickable" data-action="navigate-planning" style="cursor:pointer;">
              <div class="bbz-kpi-label">Offene Tasks</div>
              <div class="bbz-kpi-value">${allOpenTasks}</div>
              ${overdueTasks > 0
                ? `<div class="bbz-kpi-meta-alert">${overdueTasks} überfällig — zur Planung →</div>`
                : `<div class="bbz-kpi-meta-ok">keine überfällig — zur Planung →</div>`}
            </div>
            <!-- Zurück zum Cockpit -->
            <div class="bbz-kpi bbz-kpi-clickable" data-action="kpi-filter" data-scope="navigate" data-value="firms" style="cursor:pointer;">
              <div class="bbz-kpi-label">Firmen-Cockpit</div>
              <div class="bbz-kpi-value">${state.enriched.firms.length}</div>
              <div class="bbz-kpi-meta">← zurück zum Cockpit</div>
            </div>
          </div>
          <section class="bbz-section">
          <div class="bbz-section-header">
            <div><div class="bbz-section-title">Kontakte</div><div class="bbz-section-subtitle">${kpiMode === "history" ? "Mit History-Einträgen" : kpiMode === "tasks" ? "Mit offenen Tasks" : "Operative Ansprechpartner über alle Firmen"}</div></div>
            <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
              <button class="bbz-dense-toggle bbz-desktop-only" onclick="window.bbzToggleDense && window.bbzToggleDense()" title="Kompakte Ansicht">⇕ Kompakt</button>
              <button class="bbz-button bbz-button-secondary" data-action="open-history-form">+ Aktivität</button>
              <button class="bbz-button bbz-button-primary" data-action="open-contact-form">+ Kontakt</button>
            </div>
          </div>
          <div class="bbz-section-body">
            <div style="display:grid;grid-template-columns:1fr;gap:8px;margin-bottom:10px;">
              <input class="bbz-input" data-filter="contacts-search" type="text" placeholder="Suche nach Name, Firma, Funktion, Rolle, E-Mail ..." value="${helpers.escapeHtml(filters.search)}" />
              <label class="bbz-checkbox"><input type="checkbox" data-filter="contacts-archiviert" ${filters.archiviertAusblenden ? "checked" : ""} /> Archivierte ausblenden</label>
            </div>
            <!-- Desktop: Tabelle -->
            <div class="bbz-table-wrap bbz-desktop-only">
              <table class="bbz-table">
                <thead><tr>${cTh("Name","fullName")}${cTh("Firma","firmTitle")}<th>Funktion</th>${cTh("Rolle","rolle")}${cTh("Lead BBZ","leadbbz0")}<th>E-Mail</th><th>Telefon</th><th>Archiviert</th></tr></thead>
                <tbody>
                  ${rows.length ? rows.map(c => `
                    <tr>
                      <td><span class="bbz-td-name">${helpers.avatarHtml(c)}<a class="bbz-link" data-action="open-contact" data-id="${c.id}">${helpers.escapeHtml(c.fullName || c.nachname)}</a></span></td>
                      <td>${c.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${c.firmId}">${helpers.escapeHtml(c.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
                      <td>${helpers.escapeHtml(c.funktion) || '<span class="bbz-muted">—</span>'}</td>
                      <td>${helpers.escapeHtml(c.rolle) || '<span class="bbz-muted">—</span>'}</td>
                      <td>${helpers.escapeHtml(c.leadbbz0) || '<span class="bbz-muted">—</span>'}</td>
                      <td>${c.email1 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(c.email1)}">${helpers.escapeHtml(c.email1)}</a>` : '<span class="bbz-muted">—</span>'}</td>
                      <td>${helpers.escapeHtml(helpers.joinNonEmpty([c.direktwahl, c.mobile], " / ")) || '<span class="bbz-muted">—</span>'}</td>
                      <td>${c.archiviert ? '<span class="bbz-danger">Ja</span>' : '<span class="bbz-muted">Nein</span>'}</td>
                    </tr>`).join("") : `<tr><td colspan="8">${ui.emptyBlock("Keine Kontakte fuer die aktuelle Filterung gefunden.")}</td></tr>`}
                </tbody>
              </table>
            </div>
            <!-- Mobile: Card-List -->
            <div class="bbz-mobile-only bbz-card-list">
              ${rows.length ? rows.map(c => `
                <div class="bbz-list-card" data-action="open-contact" data-id="${c.id}">
                  ${helpers.avatarHtml(c)}
                  <div class="bbz-list-card-body">
                    <div class="bbz-list-card-title">${helpers.escapeHtml(c.fullName || c.nachname)}${c.archiviert ? ' <span class="bbz-muted" style="font-size:10px;">(archiviert)</span>' : ""}</div>
                    <div class="bbz-list-card-sub">
                      ${c.firmTitle ? helpers.escapeHtml(c.firmTitle) : ""}
                      ${c.funktion ? ` · ${helpers.escapeHtml(c.funktion)}` : ""}
                      ${c.rolle ? ` · ${helpers.escapeHtml(c.rolle)}` : ""}
                    </div>
                  </div>
                  <div class="bbz-list-card-right">
                    ${c.leadbbz0 ? helpers.leadbbzBadgeHtml(c.leadbbz0) : ""}
                    ${c.email1 ? `<span style="font-size:10px;color:var(--subtle);max-width:110px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${helpers.escapeHtml(c.email1)}</span>` : ""}
                  </div>
                </div>`).join("")
              : ui.emptyBlock("Keine Kontakte für die aktuelle Filterung gefunden.")}
            </div>
          </div>
          </section>
        </div>
      `;
    },

    contactDetail() {
      const contact = dataModel.getContactById(state.selection.contactId);
      if (!contact) return ui.emptyBlock("Der ausgewaehlte Kontakt wurde nicht gefunden.");
      const contactHistory = state.enriched.history.filter(h => h.contactId === contact.id).sort((a, b) => helpers.compareDateDesc(a.datum, b.datum));
      const contactTasks = state.enriched.tasks.filter(t => t.contactId === contact.id).sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline));
      const isPrivat = state.meta.privateFirmId !== null && contact.firmId === state.meta.privateFirmId;
      // Band-Farbe von der Firma erben
      const firm = contact.firmId ? dataModel.getFirmById(contact.firmId) : null;
      const bandClass = firm ? helpers.detailBandClass(firm) : "bbz-detail-band bbz-detail-band-default";
      const seed = [...(contact.vorname.charAt(0) + contact.nachname.charAt(0))].reduce((s, c) => s + c.charCodeAt(0), 0);
      const avatarIdx = seed % 6;
      const initials = (contact.vorname.charAt(0) + contact.nachname.charAt(0)).toUpperCase() || "?";

      return `
        <div>
          <div class="${bandClass}" style="margin-bottom:10px;">
            <button class="bbz-button bbz-button-secondary" style="margin-bottom:12px;background:rgba(255,255,255,0.7);" data-action="back-to-contacts">← Kontaktliste</button>
            <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:16px;flex-wrap:wrap;">
              <div style="display:flex;align-items:center;gap:14px;">
                <div class="bbz-avatar-lg" data-idx="${avatarIdx}">${helpers.escapeHtml(initials)}</div>
                <div>
                  <div class="bbz-detail-title">${helpers.escapeHtml(contact.fullName || contact.nachname)}</div>
                  <div class="bbz-detail-subtitle">
                    ${isPrivat
                      ? `<span class="bbz-pill" style="font-size:12px;">Privatperson</span>`
                      : contact.firmId
                        ? `<a class="bbz-link" data-action="open-firm" data-id="${contact.firmId}">${helpers.escapeHtml(contact.firmTitle || "Firma")}</a>`
                        : "Keine Firma verknüpft"
                    }
                    ${contact.funktion ? ` · ${helpers.escapeHtml(contact.funktion)}` : ""}
                    ${contact.rolle ? ` · ${helpers.escapeHtml(contact.rolle)}` : ""}
                  </div>
                  <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:8px;">
                    ${contact.leadbbz0 ? helpers.leadbbzBadgeHtml(contact.leadbbz0) : ""}
                    ${contact.archiviert ? '<span class="bbz-pill" style="background:var(--red-soft);color:var(--red);border-color:#f0b0b2;">Archiviert</span>' : ""}
                  </div>
                </div>
              </div>
              <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                ${contact.email1 ? `<a class="bbz-button bbz-button-secondary" href="mailto:${helpers.escapeHtml(contact.email1)}">✉ Mail</a>` : ""}
                <button class="bbz-button bbz-button-secondary" style="color:var(--red);border-color:var(--red);" data-action="delete-contact" data-id="${contact.id}" data-name="${helpers.escapeHtml(contact.fullName || contact.nachname)}">Löschen</button>
                <button class="bbz-button bbz-button-secondary" data-action="open-contact-form" data-item-id="${contact.id}">Bearbeiten</button>
                <button class="bbz-button bbz-button-secondary" data-action="open-task-form" data-contact-id="${contact.id}">+ Task</button>
                <button class="bbz-button bbz-button-primary" data-action="open-history-form" data-contact-id="${contact.id}">+ Aktivität</button>
              </div>
            </div>
          </div>
          <div class="bbz-kpis">
            ${this.kpiBlock("Tasks", contactTasks.length)}
            ${this.kpiBlock("Offen", contactTasks.filter(t => t.isOpen).length, contactTasks.some(t => t.isOpen && t.isOverdue) ? "überfällig" : contactTasks.filter(t => t.isOpen).length > 0 ? "offen" : "keine offen", contactTasks.some(t => t.isOpen && t.isOverdue) ? "alert" : "")}
            ${this.kpiBlock("Aktivitäten", contactHistory.length)}
            ${this.kpiBlock("Letzter Kontakt", contactHistory[0]?.datum ? helpers.relativeDate(contactHistory[0].datum) : "—", contactHistory.length === 0 ? "noch kein Kontakt" : "", contactHistory.length === 0 ? "warn" : "")}
          </div>
          <div class="bbz-grid bbz-grid-3">
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Stammdaten</div></div>
              <div class="bbz-section-body">
                ${ui.kv("Anrede", helpers.escapeHtml(contact.anrede) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Vorname", helpers.escapeHtml(contact.vorname) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Nachname", helpers.escapeHtml(contact.nachname) || '<span class="bbz-muted">—</span>')}
                ${isPrivat
                  ? ui.kv("Adresse / Notizen", helpers.escapeHtml(contact.kommentar) || '<span class="bbz-muted">—</span>')
                  : ui.kv("Firma", contact.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${contact.firmId}">${helpers.escapeHtml(contact.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>')
                }
                ${ui.kv("Funktion", helpers.escapeHtml(contact.funktion) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Rolle", helpers.escapeHtml(contact.rolle) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Email 1", contact.email1 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(contact.email1)}">${helpers.escapeHtml(contact.email1)}</a>` : '<span class="bbz-muted">—</span>')}
                ${ui.kv("Email 2", contact.email2 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(contact.email2)}">${helpers.escapeHtml(contact.email2)}</a>` : '<span class="bbz-muted">—</span>')}
                ${ui.kv("Direktwahl", helpers.escapeHtml(contact.direktwahl) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Mobile", helpers.escapeHtml(contact.mobile) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Geburtstag", helpers.formatDate(contact.geburtstag) || '<span class="bbz-muted">—</span>')}
                ${contact.spModified ? `
                <div style="margin-top:14px;padding-top:10px;border-top:1px solid var(--line-2);">
                  <div style="font-size:11px;color:var(--subtle);line-height:1.6;"><span style="color:var(--muted);min-width:90px;display:inline-block;">Geändert</span>${helpers.formatDateTime(contact.spModified)}${contact.spModifiedBy ? ` · ${helpers.escapeHtml(contact.spModifiedBy)}` : ""}</div>
                </div>` : ""}
              </div>
            </section>
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">CRM-Kontext</div></div>
              <div class="bbz-section-body">
                ${ui.kv("Lead BBZ", helpers.leadbbzBadgeHtml(contact.leadbbz0))}
                ${ui.kv("SGF", helpers.multiChoiceHtml(contact.sgf))}
                ${ui.kv("Event", helpers.multiChoiceHtml(contact.event))}
                ${ui.kv("Eventhistory", helpers.multiChoiceHtml(contact.eventhistory))}
                ${isPrivat ? "" : ui.kv("Kommentar", helpers.escapeHtml(contact.kommentar) || '<span class="bbz-muted">—</span>')}
              </div>
            </section>
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Aufgaben</div>
                <button class="bbz-button bbz-button-secondary" style="height:28px;font-size:12px;" data-action="open-task-form" data-contact-id="${contact.id}">+ Task</button>
              </div>
              <div class="bbz-section-body">
                ${contactTasks.length ? contactTasks.map(t => `
                  <div style="display:flex;align-items:center;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--line-2);">
                    <div>
                      <div style="font-size:13px;font-weight:600;">${helpers.escapeHtml(t.title)}</div>
                      <div style="font-size:12px;color:var(--muted);margin-top:2px;">${t.deadline ? helpers.relativeDate(t.deadline) : "Keine Deadline"}</div>
                    </div>
                    ${helpers.statusChipHtml(t.status, t.deadline)}
                  </div>`).join("") : `<div class="bbz-empty">Noch kein Task erfasst.<br><button class="bbz-button bbz-button-secondary" style="margin-top:10px;height:32px;font-size:13px;" data-action="open-task-form" data-contact-id="${contact.id}">+ Ersten Task erstellen</button></div>`}
              </div>
            </section>
          </div>
          <div class="bbz-grid bbz-grid-2" style="margin-top:12px;">
            <section class="bbz-section">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Aktivitäten</div></div>
                <button class="bbz-button bbz-button-primary" style="height:32px;font-size:13px;" data-action="open-history-form" data-contact-id="${contact.id}">+ Aktivität</button>
              </div>
              <div class="bbz-section-body">
                ${contactHistory.length ? `<div class="bbz-timeline">${contactHistory.map(h => `
                  <div class="bbz-timeline-item">
                    <div class="bbz-timeline-date">${helpers.relativeDate(h.datum) || "—"}<br><span class="bbz-muted" style="font-size:11px;">${helpers.formatDate(h.datum)}</span></div>
                    <div>
                      <div class="bbz-timeline-title">${helpers.escapeHtml(h.typ || h.title || "Eintrag")} ${h.projektbezugBool ? '<span class="bbz-chip" style="background:var(--blue-light);color:var(--blue);border-color:#a8c8e0;">Projektbezug</span>' : '<span class="bbz-chip">Allgemein</span>'}</div>
                      <div class="bbz-timeline-text">${helpers.escapeHtml(h.notizen || "—")}</div>
                      <div style="margin-top:6px;display:flex;gap:6px;">
                        <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;" data-action="edit-history" data-id="${h.id}">Bearbeiten</button>
                        <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;color:var(--red);border-color:var(--red);" data-action="delete-history" data-id="${h.id}" data-title="${helpers.escapeHtml(h.typ || h.title || 'Eintrag')}">Löschen</button>
                      </div>
                    </div>
                  </div>`).join("")}</div>` : ui.emptyBlock("Noch keine Aktivitäten erfasst.")}
              </div>
            </section>
          </div>
        </div>
      `;
    },

    planning() {
      const filters = state.filters.planning;
      const statusChoices = state.meta.choices?.[CONFIG.lists.tasks]?.["Status"] || [];
      const today = helpers.todayStart();
      const in7   = new Date(today); in7.setDate(in7.getDate() + 7);
      const in30  = new Date(today); in30.setDate(in30.getDate() + 30);
      const in365 = new Date(today); in365.setDate(in365.getDate() + 365);

      const baseRows = state.enriched.tasks.filter(t => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch = !search || [t.title, t.status, t.contactName, t.firmTitle, t.leadbbz].some(v => helpers.textIncludes(v, search));
        if (!searchMatch) return false;
        if (filters.onlyOpen && !t.isOpen) return false;

        // Fälligkeits-Filter
        if (filters.faelligkeit) {
          const d = helpers.toDate(t.deadline);
          if (filters.faelligkeit === "overdue" && !(t.isOpen && t.isOverdue)) return false;
          if (filters.faelligkeit === "week"    && !(d && d > today && d <= in7))  return false;
          if (filters.faelligkeit === "month"   && !(d && d > in7  && d <= in30))  return false;
          if (filters.faelligkeit === "rest"    && !(d && d > in30 && d <= in365)) return false;
        }

        // Segment-Filter
        if (filters.segment) {
          const firm = t.firmId ? state.enriched.firms.find(f => f.id === t.firmId) : null;
          const kl = String(firm?.klassifizierung || "").toUpperCase();
          if (!kl.startsWith(filters.segment.toUpperCase())) return false;
        }

        // Lead BBZ-Filter
        if (filters.leadbbz && t.leadbbz !== filters.leadbbz) return false;

        return true;
      });

      // Zähler für Chips
      const cntOverdue = state.enriched.tasks.filter(t => t.isOpen && t.isOverdue).length;
      const cntWeek    = state.enriched.tasks.filter(t => { const d = helpers.toDate(t.deadline); return t.isOpen && d && d > today && d <= in7; }).length;
      const cntMonth   = state.enriched.tasks.filter(t => { const d = helpers.toDate(t.deadline); return t.isOpen && d && d > in7 && d <= in30; }).length;
      const cntRest    = state.enriched.tasks.filter(t => { const d = helpers.toDate(t.deadline); return t.isOpen && d && d > in30 && d <= in365; }).length;

      const allLeadbbz = [...new Set(state.enriched.tasks.map(t => t.leadbbz).filter(Boolean))].sort();

      const chipF = (label, val, cnt, style = "") => {
        const active = filters.faelligkeit === val;
        return `<button class="bbz-kpi-chip ${active ? "bbz-kpi-chip-active" : ""}" style="${style}" data-action="kpi-filter" data-scope="planning-faelligkeit" data-value="${val}">${label} <span>${cnt}</span></button>`;
      };
      const chipS = (label, val) => {
        const cnt = val === "" ? state.enriched.tasks.length : state.enriched.tasks.filter(t => {
          const firm = t.firmId ? state.enriched.firms.find(f => f.id === t.firmId) : null;
          return String(firm?.klassifizierung || "").toUpperCase().startsWith(val);
        }).length;
        const active = filters.segment === val;
        return `<button class="bbz-kpi-chip ${active ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="planning-segment" data-value="${val}">${label === "Alle" ? "Alle" : label} <span>${cnt}</span></button>`;
      };
      const chipL = (l) => {
        const cnt = state.enriched.tasks.filter(t => t.leadbbz === l).length;
        const active = filters.leadbbz === l;
        return `<button class="bbz-kpi-chip ${active ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="planning-leadbbz" data-value="${helpers.escapeHtml(l)}">${helpers.escapeHtml(l)} <span>${cnt}</span></button>`;
      };

      // Multi-column sort
      const dir = filters.sortDir === "asc" ? 1 : -1;
      const sorted = [...baseRows].sort((a, b) => {
        if (filters.sortBy === "deadline") {
          const ad = helpers.toDate(a.deadline), bd = helpers.toDate(b.deadline);
          if (!ad && !bd) return 0;
          if (!ad) return 1;
          if (!bd) return -1;
          return (ad - bd) * dir;
        }
        if (filters.sortBy === "leadbbz") {
          return String(a.leadbbz || "").localeCompare(String(b.leadbbz || ""), "de") * dir;
        }
        if (filters.sortBy === "firma") {
          return String(a.firmTitle || "").localeCompare(String(b.firmTitle || ""), "de") * dir;
        }
        return 0;
      });

      // Gruppierung
      const groupBy = filters.groupBy;
      let groups = [];
      if (groupBy === "none") {
        groups = [{ key: null, label: null, rows: sorted }];
      } else {
        const map = new Map();
        sorted.forEach(t => {
          const key = (groupBy === "status" ? t.status : t.leadbbz) || "—";
          if (!map.has(key)) map.set(key, []);
          map.get(key).push(t);
        });
        groups = [...map.entries()]
          .sort((a, b) => a[0].localeCompare(b[0], "de"))
          .map(([key, rows]) => ({ key, label: key, rows }));
      }

      // Sort-Header helper
      const th = (label, col) => {
        const active = filters.sortBy === col;
        const icon = active ? (filters.sortDir === "asc" ? " ↑" : " ↓") : "";
        return `<th style="cursor:pointer;user-select:none;${active ? "color:var(--blue);" : ""}" data-action="set-sort" data-col="${col}">${label}${icon}</th>`;
      };

      const renderTaskRow = (t) => {
        const statusCell = statusChoices.length
          ? `<select class="bbz-select" style="height:30px;font-size:12px;" data-action="task-status-change" data-task-id="${t.id}">${statusChoices.map(s => `<option value="${helpers.escapeHtml(s)}" ${t.status === s ? "selected" : ""}>${helpers.escapeHtml(s)}</option>`).join("")}</select>`
          : helpers.statusChipHtml(t.status, t.deadline);
        return `
          <tr>
            <td>${t.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${t.firmId}">${helpers.escapeHtml(t.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
            <td>${t.contactId ? `<a class="bbz-link" data-action="open-contact" data-id="${t.contactId}">${helpers.escapeHtml(t.contactName || "Kontakt")}</a>` : helpers.escapeHtml(t.contactName || "—")}</td>
            <td>${helpers.escapeHtml(t.title) || '<span class="bbz-muted">—</span>'}</td>
            <td class="${helpers.isOpenTask(t.status) && helpers.isOverdue(t.deadline) ? "bbz-danger" : ""}">${t.deadline ? helpers.relativeDate(t.deadline) : '<span class="bbz-muted">—</span>'}</td>
            <td>${statusCell}</td>
            <td>${helpers.escapeHtml(t.leadbbz) || '<span class="bbz-muted">—</span>'}</td>
            <td style="white-space:nowrap;">
              <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;margin-right:4px;" data-action="edit-task" data-id="${t.id}">Bearbeiten</button>
              <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;color:var(--red);border-color:var(--red);" data-action="delete-task" data-id="${t.id}" data-title="${helpers.escapeHtml(t.title)}">Löschen</button>
            </td>
          </tr>`;
      };

      const tableHead = `<thead><tr>
        ${th("Firma", "firma")}
        <th>Kontaktperson</th>
        <th>Titel</th>
        ${th("Deadline", "deadline")}
        <th>Status</th>
        ${th("Leadbbz", "leadbbz")}
        <th>Aktionen</th>
      </tr></thead>`;

      const tableBody = groups.map(g => `
        ${g.label ? `<tr><td colspan="7" style="background:#f1f5fb;font-size:12px;font-weight:700;color:var(--text);padding:7px 12px;">${helpers.escapeHtml(g.label)} <span style="font-weight:400;color:var(--muted);">(${g.rows.length})</span></td></tr>` : ""}
        ${g.rows.length ? g.rows.map(renderTaskRow).join("") : `<tr><td colspan="7">${ui.emptyBlock("Keine Tasks.")}</td></tr>`}
      `).join("");

      return `
        <div>
          <div class="bbz-kpis">
            <!-- Kachel 1: Fälligkeit -->
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Tasks gesamt</div>
              <div class="bbz-kpi-value">${state.enriched.tasks.length}</div>
              <div style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${chipF("Überfällig", "overdue", cntOverdue, cntOverdue > 0 ? "background:var(--red-soft);border-color:#f0b0b2;color:var(--red);" : "")}
                ${chipF("Woche", "week", cntWeek, cntWeek > 0 ? "background:#fff9eb;border-color:#f4dfab;color:var(--amber);" : "")}
                ${chipF("Monat", "month", cntMonth)}
                ${chipF("Übrige", "rest", cntRest)}
                ${filters.faelligkeit ? `<button class="bbz-kpi-chip" data-action="kpi-filter" data-scope="planning-faelligkeit" data-value="">Alle</button>` : ""}
              </div>
            </div>
            <!-- Kachel 2: Kundenklassifizierung -->
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Kundenklassifizierung</div>
              <div class="bbz-kpi-value">${filters.segment || "—"}</div>
              <div style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${["A","B","C"].map(k => chipS(k, k)).join("")}
                <button class="bbz-kpi-chip ${!filters.segment ? "bbz-kpi-chip-active" : ""}" data-action="kpi-filter" data-scope="planning-segment" data-value="">Alle</button>
              </div>
            </div>
            <!-- Kachel 3: Lead BBZ -->
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Lead BBZ</div>
              <div class="bbz-kpi-value">${filters.leadbbz ? helpers.escapeHtml(filters.leadbbz) : "—"}</div>
              <div style="margin-top:8px;display:flex;gap:4px;flex-wrap:wrap;">
                ${allLeadbbz.map(chipL).join("")}
                ${filters.leadbbz ? `<button class="bbz-kpi-chip" data-action="kpi-filter" data-scope="planning-leadbbz" data-value="">Alle</button>` : ""}
              </div>
            </div>
            <!-- Kachel 4: Sichtbar -->
            ${this.kpiBlock("Sichtbar", baseRows.length, baseRows.length < state.enriched.tasks.length ? `von ${state.enriched.tasks.length} gefiltert` : "alle", baseRows.length < state.enriched.tasks.length ? "warn" : "")}
          </div>
          <section class="bbz-section">
            <div class="bbz-section-header">
              <div><div class="bbz-section-title">Planung</div><div class="bbz-section-subtitle">Aufgaben mit Fokus auf offen und überfällig</div></div>
              <button class="bbz-button bbz-button-primary" data-action="open-task-form">+ Task</button>
            </div>
            <div class="bbz-section-body">
              <div class="bbz-filters-3 bbz-planning-filters" style="grid-template-columns:2fr 1fr 1fr;">
                <input class="bbz-input" data-filter="planning-search" type="text" placeholder="Suche nach Titel, Firma, Kontakt, Status ..." value="${helpers.escapeHtml(filters.search)}" />
                <select class="bbz-select" data-filter="planning-groupby">
                  <option value="none"    ${filters.groupBy === "none"    ? "selected" : ""}>Keine Gruppierung</option>
                  <option value="status"  ${filters.groupBy === "status"  ? "selected" : ""}>Gruppe: Status</option>
                  <option value="leadbbz" ${filters.groupBy === "leadbbz" ? "selected" : ""}>Gruppe: Leadbbz</option>
                </select>
                <label class="bbz-checkbox"><input type="checkbox" data-filter="planning-open" ${filters.onlyOpen ? "checked" : ""} /> Nur offene Tasks</label>
              </div>
              <!-- Desktop: Tabelle -->
              <div class="bbz-table-wrap bbz-desktop-only">
                <table class="bbz-table" style="min-width:1060px;">
                  ${tableHead}
                  <tbody>${baseRows.length ? tableBody : `<tr><td colspan="7">${ui.emptyBlock("Keine Tasks fuer die aktuelle Filterung gefunden.")}</td></tr>`}</tbody>
                </table>
              </div>
              <!-- Mobile: Card-List -->
              <div class="bbz-mobile-only">
                ${baseRows.length ? `<div class="bbz-card-list">${groups.map(g => `
                  ${g.label ? `<div style="padding:8px 14px 4px;font-size:11px;font-weight:700;color:var(--subtle);text-transform:uppercase;letter-spacing:0.06em;">${helpers.escapeHtml(g.label)} (${g.rows.length})</div>` : ""}
                  ${g.rows.map(t => `
                    <div class="bbz-task-card" data-action="edit-task" data-id="${t.id}">
                      <div class="bbz-task-card-top">
                        <span class="bbz-task-card-title">${helpers.escapeHtml(t.title)}</span>
                        ${helpers.statusChipHtml(t.status, t.deadline)}
                      </div>
                      <div class="bbz-task-card-meta">
                        ${t.firmTitle ? `<span>${helpers.escapeHtml(t.firmTitle)}</span><span>·</span>` : ""}
                        ${t.contactName ? `<span>${helpers.escapeHtml(t.contactName)}</span>` : ""}
                        ${t.deadline ? `<span>·</span><span class="${t.isOpen && t.isOverdue ? "bbz-danger" : ""}">${helpers.relativeDate(t.deadline)}</span>` : ""}
                        ${t.leadbbz ? `<span>·</span>${helpers.leadbbzBadgeHtml(t.leadbbz)}` : ""}
                      </div>
                    </div>`).join("")}
                `).join("")}</div>`
                : ui.emptyBlock("Keine Tasks für die aktuelle Filterung.")}
              </div>
            </div>
          </section>
        </div>
      `;
    },

    historyView() {
      const filters = state.filters.history;
      const today   = helpers.todayStart();
      const lens    = filters.lens;
      const gran    = filters.granularitaet;
      const periode = filters.periode;

      // ── Datenbasis: Perioden-Maps + rollierende Spanne je Granularität ───────
      const contactById = new Map(state.enriched.contacts.map(c => [c.id, c]));
      const firmById    = new Map(state.enriched.firms.map(f => [f.id, f]));
      const leadOf = (h) => (contactById.get(h.contactId)?.leadbbz0 || "").trim();

      // Rollierende Spanne (älteste zuerst): Monat=12 · Quartal=8 · Jahr=max. 3.
      let spanPeriods = [];
      if (gran === "quartal") {
        for (let i = 7; i >= 0; i--) { const d = new Date(today.getFullYear(), today.getMonth() - i * 3, 1); spanPeriods.push(helpers.periodKey(d, gran)); }
      } else if (gran === "jahr") {
        const years = [...new Set(state.enriched.history.map(h => { const d = helpers.toDate(h.datum); return d ? d.getFullYear() : null; }).filter(Boolean))].sort((a, b) => b - a).slice(0, 3).sort((a, b) => a - b);
        spanPeriods = years.length ? years.map(String) : [String(today.getFullYear())];
      } else {
        for (let i = 11; i >= 0; i--) { const d = new Date(today.getFullYear(), today.getMonth() - i, 1); spanPeriods.push(helpers.periodKey(d, gran)); }
      }
      spanPeriods = [...new Set(spanPeriods)];
      const spanSet   = new Set(spanPeriods);
      const spanLabel = gran === "quartal" ? "letzte 8 Quartale" : gran === "jahr" ? "alle Jahre" : "letzte 12 Monate";
      const spanChip  = gran === "quartal" ? "8 Quartale" : gran === "jahr" ? "Jahre" : "12 Monate";

      // Aktivitäten der Spanne + Count je Periode (für die Zeitraum-Balken).
      const base = state.enriched.history.filter(h => spanSet.has(helpers.periodKey(h.datum, gran)));
      const periodCounts = new Map();
      base.forEach(h => { const k = helpers.periodKey(h.datum, gran); periodCounts.set(k, (periodCounts.get(k) || 0) + 1); });
      const periodBars = spanPeriods.map(p => ({ key: p, label: helpers.periodLabel(p), count: periodCounts.get(p) || 0 }));

      // Aktives Zeitfenster (Periode) — Lead-Chart aggregiert über DASSELBE Fenster.
      const inPeriode    = (h) => !periode || helpers.periodKey(h.datum, gran) === periode;
      const activeWindow = base.filter(inPeriode);

      // Lead-BBZ-Balken über activeWindow, case-insensitiv dedupliziert
      // (kanonische Schreibweise = häufigste Variante).
      const leadAgg = new Map(); // lowerKey -> { variants: Map<variant,count>, total }
      activeWindow.forEach(h => {
        const raw = leadOf(h);
        if (!raw) return;
        const key = raw.toLowerCase();
        if (!leadAgg.has(key)) leadAgg.set(key, { variants: new Map(), total: 0 });
        const e = leadAgg.get(key);
        e.total += 1;
        e.variants.set(raw, (e.variants.get(raw) || 0) + 1);
      });
      const leadBars = [...leadAgg.values()].map(e => {
        let canon = "", max = -1;
        e.variants.forEach((c, v) => { if (c > max) { max = c; canon = v; } });
        return { label: canon, count: e.total };
      }).sort((a, b) => b.count - a.count);

      // Scope für Band/Cockpit/Timeline: aktives Fenster + Linse.
      const inLens = (h) => !lens || leadOf(h).toLowerCase() === lens.toLowerCase();
      const scoped = activeWindow.filter(inLens);

      // ── Band-Kennzahlen + dynamisches Ø pro Woche (fensterTage / 7) ──────────
      const daysInMonth  = (y, m0) => new Date(y, m0 + 1, 0).getDate();
      const daysInPeriod = (key) => {
        if (gran === "jahr") { const y = +key; return ((y % 4 === 0 && y % 100 !== 0) || y % 400 === 0) ? 366 : 365; }
        if (gran === "quartal") { const [ys, qs] = key.split("-Q"); const y = +ys, q = +qs; let d = 0; for (let m = (q - 1) * 3; m < q * 3; m++) d += daysInMonth(y, m); return d; }
        const [ys, ms] = key.split("-"); return daysInMonth(+ys, +ms - 1);
      };
      const fensterTage  = (periode ? [periode] : spanPeriods).reduce((s, k) => s + daysInPeriod(k), 0);
      const aktivTotal   = scoped.length;
      const aktiveFirmen = new Set(scoped.map(h => h.firmId).filter(Boolean)).size;
      const avgProWoche  = (aktivTotal / Math.max(1, fensterTage / 7)).toFixed(1);

      const allKontaktart = [...new Set(base.map(h => h.typ).filter(Boolean))].sort();
      const activeFilterBadges = filters.kontaktart
        ? `<span class="bbz-chip" style="background:var(--blue-light);color:var(--blue);border-color:#a8c8e0;cursor:pointer;" data-action="kpi-filter" data-scope="history-kontaktart" data-value="">${helpers.escapeHtml(filters.kontaktart)} ×</span>`
        : "";

      // ── Timeline-Zeilen (Sekundär-Modus): scoped + Suche/Kontaktart ──────────
      const rows = scoped.filter(h => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch = !search || [h.contactName, h.firmTitle, h.typ, h.notizen].some(v => helpers.textIncludes(v, search));
        const artMatch    = !filters.kontaktart || h.typ === filters.kontaktart;
        return searchMatch && artMatch;
      });

      // ── Timeline-Gruppierung ─────────────────────────────────────────────────
      const renderCard = (h, hideFirm = false) => {
        const hasLongText = (h.notizen || "").length > 120;
        const tagsHtml = `
          <div style="display:flex;gap:4px;flex-wrap:wrap;justify-content:flex-end;flex-shrink:0;margin-left:8px;">
            ${h.projektbezugBool ? '<span class="bbz-chip" style="background:var(--blue-light);color:var(--blue);border-color:#a8c8e0;">Projektbezug</span>' : ""}
            ${h.leadbbz ? helpers.leadbbzBadgeHtml(h.leadbbz) : ""}
          </div>`;
        return `
          <div class="bbz-timeline-item">
            <div style="display:flex;align-items:center;gap:10px;margin-bottom:5px;">
              ${helpers.avatarHtml({ vorname: (h.contactName||"").split(" ")[0]||"", nachname: (h.contactName||"").split(" ").slice(-1)[0]||"" })}
              <div style="flex:1;min-width:0;">
                <div style="font-size:13px;font-weight:700;line-height:1.25;">
                  ${h.contactId ? `<a class="bbz-link" data-action="open-contact" data-id="${h.contactId}">${helpers.escapeHtml(h.contactName || "")}</a>` : helpers.escapeHtml(h.contactName || "—")}
                </div>
                ${!hideFirm ? `<div style="font-size:12px;color:var(--muted);margin-top:1px;">
                  ${h.firmId ? `<a class="bbz-link" style="font-weight:500;" data-action="open-firm" data-id="${h.firmId}">${helpers.escapeHtml(h.firmTitle)}</a>` : (h.firmTitle ? helpers.escapeHtml(h.firmTitle) : "")}
                </div>` : ""}
              </div>
              <div style="text-align:right;flex-shrink:0;">
                <div style="font-size:12px;font-weight:600;color:var(--text);" title="${helpers.formatDate(h.datum)}">${helpers.relativeDate(h.datum) || "—"}</div>
                <div style="font-size:11px;color:var(--muted);">${helpers.formatDate(h.datum)}</div>
              </div>
            </div>
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:5px;">
              <span style="font-size:13px;font-weight:600;color:var(--text);">${helpers.escapeHtml(h.typ || "Eintrag")}</span>
              ${tagsHtml}
            </div>
            <div class="bbz-timeline-text bbz-timeline-clamp">${helpers.escapeHtml(h.notizen || "—")}</div>
            ${hasLongText ? `<button class="bbz-button bbz-button-secondary" style="height:22px;font-size:11px;padding:0 8px;margin-top:4px;" data-action="toggle-expand">mehr</button>` : ""}
            <div style="margin-top:8px;display:flex;gap:6px;align-items:center;">
              <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;" data-action="edit-history" data-id="${h.id}">Bearbeiten</button>
              <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;" data-action="open-task-form" data-contact-id="${h.contactId || ""}" title="Task für ${helpers.escapeHtml(h.contactName || "")} erstellen">+ Task</button>
            </div>
          </div>`;
      };

      // Datum-Gruppen
      const makeDateGroups = (items) => {
        const vor7d  = new Date(today); vor7d.setDate(vor7d.getDate() - 7);
        const vor30d = new Date(today); vor30d.setDate(vor30d.getDate() - 30);
        const groups = [
          { key: "today",  label: "Heute",          items: [] },
          { key: "week",   label: "Diese Woche",     items: [] },
          { key: "month",  label: "Dieser Monat",    items: [] },
          { key: "older",  label: "Älter",           items: [] }
        ];
        items.forEach(h => {
          const d = helpers.toDate(h.datum);
          if (!d) { groups[3].items.push(h); return; }
          if (d >= today)   groups[0].items.push(h);
          else if (d >= vor7d)  groups[1].items.push(h);
          else if (d >= vor30d) groups[2].items.push(h);
          else                  groups[3].items.push(h);
        });
        return groups.filter(g => g.items.length > 0);
      };

      // Chronologische Timeline (Sekundär-Modus) — reine Datums-Gruppierung.
      const groups = makeDateGroups(rows);
      const timelineHtml = groups.length ? groups.map(g => `
        <div class="bbz-history-group">
          <div class="bbz-history-group-header">
            <span class="bbz-history-group-label">${helpers.escapeHtml(g.label)}</span>
            <span class="bbz-history-group-count">${g.items.length} Eintrag${g.items.length !== 1 ? "e" : ""}</span>
          </div>
          <div class="bbz-timeline">${g.items.map(h => renderCard(h, false)).join("")}</div>
        </div>`).join("")
        : ui.emptyBlock("Keine Aktivitäten für die aktuelle Filterung gefunden.", "open-history-form", "+ Erste Aktivität erfassen");

      // ── Firmen-Bericht (Default): Firmen mit ≥1 Aktivität im Scope ────────────
      const cockpitSearch = filters.search.trim().toLowerCase();
      const actsByFirm = new Map();
      scoped.forEach(h => {
        if (!h.firmId) return;
        if (!actsByFirm.has(h.firmId)) actsByFirm.set(h.firmId, []);
        actsByFirm.get(h.firmId).push(h);
      });
      const cockpitReport = [...actsByFirm.entries()]
        .map(([fid, acts]) => ({ f: firmById.get(fid), acts: acts.slice().sort((a, b) => helpers.compareDateDesc(a.datum, b.datum)) }))
        .filter(x => x.f && (!cockpitSearch || helpers.textIncludes(x.f.title, cockpitSearch)))
        .sort((a, b) => helpers.compareDateDesc(a.acts[0]?.datum, b.acts[0]?.datum));

      const firmReportRow = (f, acts) => {
        const expanded = filters.expandedFirms.includes(f.id);
        const latest = acts[0]?.datum;
        return `
          <div class="bbz-mini-item" style="padding:0;overflow:hidden;">
            <div style="display:flex;align-items:center;gap:10px;padding:9px 12px;cursor:pointer;" data-action="toggle-firm-expand" data-firm-id="${f.id}">
              <div style="min-width:0;flex:1;">
                <div style="display:flex;align-items:center;gap:6px;min-width:0;">
                  <span class="bbz-mini-title" style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${helpers.escapeHtml(f.title)}</span>
                  ${f.klassifizierung ? `<span class="${helpers.firmBadgeClass(f.klassifizierung)}" style="flex-shrink:0;">${helpers.escapeHtml(f.klassifizierung)}</span>` : ""}
                  ${f.vip ? `<span class="bbz-pill bbz-pill-vip" style="flex-shrink:0;">♛</span>` : ""}
                </div>
                <div class="bbz-mini-meta" style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${latest ? `zuletzt ${helpers.relativeDate(latest)}` : ""} · ${acts.length} Aktivität${acts.length !== 1 ? "en" : ""}</div>
              </div>
              <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:12px;padding:0 9px;flex-shrink:0;" data-action="open-history-form" data-firm-id="${f.id}">+ Aktivität</button>
              <span data-chevron style="flex-shrink:0;color:var(--muted);font-size:12px;transition:transform .15s;transform:${expanded ? "rotate(90deg)" : "none"};">▶</span>
            </div>
            ${expanded ? `<div style="border-top:1px solid var(--line-2);background:var(--panel-2);padding:8px 10px;"><div class="bbz-timeline">${acts.map(h => renderCard(h, true)).join("")}</div></div>` : ""}
          </div>`;
      };

      const firmReportHtml = cockpitReport.length
        ? `<div class="bbz-mini-list">${cockpitReport.map(x => firmReportRow(x.f, x.acts)).join("")}</div>`
        : ui.emptyBlock(cockpitSearch ? "Keine Firma gefunden." : "Keine Aktivitäten im gewählten Zeitraum.");

      // ── Instrumenten-Band: reine Anzeige, 3 Kacheln (12-Monats-Fenster) ──────
      const bandTile = (accent, label, valueHtml, extraHtml) => `
        <div style="background:var(--panel);border:1px solid var(--line);border-top:3px solid ${accent};border-radius:var(--r-xl);padding:14px 16px;">
          <div class="bbz-kpi-label">${label}</div>
          <div class="bbz-kpi-value">${valueHtml}</div>
          ${extraHtml || ""}
        </div>`;
      const bandScopeLabel = `${periode ? helpers.escapeHtml(helpers.periodLabel(periode)) : spanLabel}${lens ? " · " + helpers.escapeHtml(lens) : ""}`;
      const bandHtml = `
        <div class="bbz-kpis" style="grid-template-columns:repeat(auto-fit,minmax(150px,1fr));">
          ${bandTile("var(--blue)", "Aktivitäten total", `${aktivTotal}`, `<div class="bbz-kpi-meta">${bandScopeLabel}</div>`)}
          ${bandTile("var(--green)", "Aktive Firmen", `${aktiveFirmen}`, `<div class="bbz-kpi-meta">mit ≥1 Aktivität</div>`)}
          ${bandTile("var(--amber)", "Ø pro Woche", `${avgProWoche}`, `<div class="bbz-kpi-meta">Fenster ≈ ${Math.round(fensterTage / 7)} Wo.</div>`)}
        </div>`;

      // ── Schnellerfassung (Buttons je Kontaktart-Choice) ──────────────────────
      const schnellTypen = (state.meta.choices?.[CONFIG.lists.history]?.["Kontaktart"] || allKontaktart).slice(0, 6);
      const schnellerfassungHtml = schnellTypen.length ? `
        <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-bottom:12px;">
          <span style="font-size:12px;color:var(--muted);">Schnell erfassen:</span>
          ${schnellTypen.map(t => `<button class="bbz-button bbz-button-secondary" style="height:30px;font-size:12px;" data-action="open-history-form" data-typ="${helpers.escapeHtml(t)}">+ ${helpers.escapeHtml(t)}</button>`).join("")}
        </div>` : "";

      // ── Steuerung: Slim-Actionbar + zwei klickbare Balken-Filter ─────────────
      const ansichtToggle = `
        <div style="display:flex;border-radius:var(--r-sm);overflow:hidden;flex-shrink:0;">
          ${[["firms", "Firmen"], ["timeline", "Chronologisch"]].map(([m, label]) => {
            const active = filters.viewMode === m;
            return `<button class="bbz-button ${active ? "bbz-button-primary" : "bbz-button-secondary"}" style="border-radius:0;height:34px;font-size:12px;" data-action="history-view-mode" data-mode="${m}">${label}</button>`;
          }).join("")}
        </div>`;
      const slimActionbar = `
        <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:12px;">
          ${ansichtToggle}
          <input class="bbz-input" style="flex:1;min-width:180px;" data-filter="history-search" type="text" placeholder="🔍 Firma oder Aktivität suchen ..." value="${helpers.escapeHtml(filters.search)}" />
          <button class="bbz-button bbz-button-primary" style="height:34px;white-space:nowrap;flex-shrink:0;" data-action="open-history-form">+ Aktivität</button>
        </div>`;

      // Lead-BBZ: horizontale Balken (Länge=Count), Klick=Toggle-Filter.
      const maxLead = Math.max(1, ...leadBars.map(b => b.count));
      const leadChartHtml = leadBars.length ? leadBars.map(b => {
        const active = lens && b.label.toLowerCase() === lens.toLowerCase();
        const pct = Math.round(b.count / maxLead * 100);
        return `
          <div class="bbz-actbar" data-action="filter-lens" data-value="${helpers.escapeHtml(b.label)}" title="${helpers.escapeHtml(b.label)}: ${b.count}" style="display:flex;align-items:center;gap:8px;cursor:pointer;padding:1px 0;">
            <span style="font-size:11px;width:64px;flex-shrink:0;text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:${active ? "var(--amber)" : "var(--muted)"};font-weight:${active ? 700 : 400};">${helpers.escapeHtml(b.label)}</span>
            <div style="flex:1;height:13px;background:var(--line);border-radius:3px;overflow:hidden;">
              <div class="bbz-actbar-fill ${active ? "bbz-on" : ""}" style="height:100%;width:${pct}%;"></div>
            </div>
            <span style="font-size:11px;width:22px;flex-shrink:0;text-align:right;color:var(--muted);">${b.count}</span>
          </div>`;
      }).join("") : `<div style="font-size:12px;color:var(--muted);padding:4px 0;">Keine Lead-BBZ-Aktivitäten.</div>`;

      // Zeitraum: vertikale Perioden-Balken je Granularität, Klick=Toggle-Periode.
      const maxPeriod = Math.max(1, ...periodBars.map(b => b.count));
      const periodChartHtml = `
        <div style="display:flex;align-items:flex-end;gap:4px;height:56px;">
          ${periodBars.map(b => {
            const active = periode === b.key;
            const barH = b.count === 0 ? 2 : Math.round(b.count / maxPeriod * 44) + 4;
            return `<div class="bbz-actbar" data-action="filter-periode" data-value="${b.key}" title="${helpers.escapeHtml(helpers.periodLabel(b.key))}: ${b.count}" style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;cursor:pointer;gap:2px;min-width:0;">
              <span style="font-size:9px;color:var(--muted);line-height:1;">${b.count || ""}</span>
              <div class="bbz-actbar-fill ${active ? "bbz-on" : ""}" style="width:100%;height:${barH}px;border-radius:2px 2px 0 0;"></div>
              <span style="font-size:9px;color:${active ? "var(--amber)" : "var(--muted)"};font-weight:${active ? 700 : 400};line-height:1;white-space:nowrap;overflow:hidden;max-width:100%;">${helpers.escapeHtml(b.label)}</span>
            </div>`;
          }).join("")}
        </div>`;

      const chartLabel = (text, right = "") => `<div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:6px;min-height:26px;"><span style="font-size:10px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:var(--muted);flex-shrink:0;">${text}</span>${right}</div>`;

      // Granularität + Reset-Chip (zeigt aktive Spanne).
      const granSwitcher = `
        <div style="display:flex;border-radius:var(--r-sm);overflow:hidden;flex-shrink:0;">
          ${[["monat", "Monat"], ["quartal", "Quartal"], ["jahr", "Jahr"]].map(([g, l]) => {
            const active = gran === g;
            return `<button class="bbz-button ${active ? "bbz-button-primary" : "bbz-button-secondary"}" style="border-radius:0;height:26px;font-size:11px;padding:0 8px;" data-action="filter-granularitaet" data-value="${g}">${l}</button>`;
          }).join("")}
        </div>`;
      const resetPeriodeChip = `<span class="bbz-chip" style="cursor:pointer;flex-shrink:0;${!periode ? "background:var(--blue-light);color:var(--blue);border-color:#a8c8e0;" : ""}" data-action="filter-periode" data-value="">${spanChip}</span>`;

      const controlPanelHtml = `
        <div style="background:var(--panel);border:1px solid var(--line);border-radius:var(--r-xl);box-shadow:var(--shadow-sm);padding:12px 14px;margin-bottom:14px;">
          ${slimActionbar}
          <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:16px;align-items:start;">
            <div>
              ${chartLabel("Lead BBZ")}
              <div style="display:grid;gap:3px;">${leadChartHtml}</div>
            </div>
            <div>
              ${chartLabel("Zeitraum", `<div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;justify-content:flex-end;">${granSwitcher}${resetPeriodeChip}</div>`)}
              ${periodChartHtml}
            </div>
          </div>
        </div>`;

      return `
        <div>
          <!-- Steuerung: Slim-Actionbar + klickbare Balken-Filter -->
          ${controlPanelHtml}

          <!-- Instrumenten-Band (reine Anzeige, 3 Kacheln) -->
          ${bandHtml}

          <!-- Schnellerfassung (über dem Hauptinhalt) -->
          ${schnellerfassungHtml}

          <!-- Hauptinhalt (einspaltig, je nach Ansicht) -->
          ${filters.viewMode === "timeline" ? `
            <section class="bbz-section">
              <div class="bbz-section-header">
                <div>
                  <div class="bbz-section-title">Aktivitäten</div>
                  <div class="bbz-section-subtitle">Chronologische Timeline</div>
                </div>
                <select class="bbz-select" style="height:32px;font-size:12px;min-width:130px;" data-filter="history-kontaktart">
                  <option value="">Alle Kontaktarten</option>
                  ${allKontaktart.map(k => `<option value="${helpers.escapeHtml(k)}" ${filters.kontaktart === k ? "selected" : ""}>${helpers.escapeHtml(k)}</option>`).join("")}
                </select>
              </div>
              <div class="bbz-section-body">
                ${activeFilterBadges ? `<div style="display:flex;gap:4px;flex-wrap:wrap;margin-bottom:8px;">${activeFilterBadges}</div>` : ""}
                ${timelineHtml}
              </div>
            </section>`
          : `
            <section class="bbz-section">
              <div class="bbz-section-header">
                <div>
                  <div class="bbz-section-title">Firmen-Bericht</div>
                  <div class="bbz-section-subtitle">${cockpitReport.length} Firma${cockpitReport.length !== 1 ? "en" : ""} mit Aktivität · neueste zuerst</div>
                </div>
              </div>
              <div class="bbz-section-body">${firmReportHtml}</div>
            </section>`}
        </div>
      `;
    },

    events() {
      // Events mit Nachbearbeitung — hardcodiert, bei neuem Anlass hier ergänzen
      const EVENTS_MIT_NACHBEARBEITUNG = ["BOL", "SummerConv."];

      const allGroups = state.enriched.events;
      const totalActiveContacts = state.enriched.contacts.filter(c => !c.archiviert).length;

      const cardHtml = allGroups.map(group => {
        const isAnlass = EVENTS_MIT_NACHBEARBEITUNG.includes(group.name);
        const listLabel = isAnlass ? "Einladungsliste" : "Versandliste";
        const typeBadge = isAnlass
          ? `<span class="bbz-pill bbz-pill-a" style="font-size:10px;padding:2px 7px;">Anlass</span>`
          : `<span class="bbz-pill" style="font-size:10px;padding:2px 7px;background:#f3e8ff;color:#6d1fb8;">Versand</span>`;

        const contacts = group.contacts;
        const firmIds = new Set(contacts.map(c => c.firmId).filter(Boolean));
        const cntFirmen = firmIds.size;
        const cntA = contacts.filter(c => c.segment.startsWith("A")).length;
        const cntB = contacts.filter(c => c.segment.startsWith("B")).length;
        const cntC = contacts.filter(c => c.segment.startsWith("C")).length;
        const cntRest = contacts.length - cntA - cntB - cntC;
        const pct = totalActiveContacts > 0 ? Math.round((contacts.length / totalActiveContacts) * 100) : 0;

        return `
          <div class="bbz-event-card">
            <div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:12px;">
              <div style="font-size:15px;font-weight:700;letter-spacing:-0.02em;">${helpers.escapeHtml(group.name)}</div>
              ${typeBadge}
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:12px;">
              <div class="bbz-event-stat">
                <div class="bbz-event-stat-label">Firmen</div>
                <div class="bbz-event-stat-value">${cntFirmen}</div>
              </div>
              <div class="bbz-event-stat">
                <div class="bbz-event-stat-label">Kontakte</div>
                <div class="bbz-event-stat-value">${contacts.length}</div>
              </div>
            </div>
            <div style="display:flex;gap:5px;flex-wrap:wrap;margin-bottom:12px;">
              ${cntA ? `<span class="bbz-pill bbz-pill-a" style="font-size:11px;">A: ${cntA}</span>` : ""}
              ${cntB ? `<span class="bbz-pill bbz-pill-b" style="font-size:11px;">B: ${cntB}</span>` : ""}
              ${cntC ? `<span class="bbz-pill bbz-pill-c" style="font-size:11px;">C: ${cntC}</span>` : ""}
              ${cntRest > 0 ? `<span class="bbz-pill" style="font-size:11px;background:#f3e8ff;color:#6d1fb8;">Übrige: ${cntRest}</span>` : ""}
            </div>
            <div style="padding-top:10px;border-top:1px solid var(--line-2);">
              <div style="font-size:11px;color:var(--muted);margin-bottom:10px;">
                Reichweite: <strong style="color:var(--blue);">${pct}%</strong> aller aktiven Kontakte
              </div>
              <div style="display:flex;gap:6px;flex-wrap:wrap;">
                <button class="bbz-button bbz-button-primary" style="height:30px;font-size:11px;font-weight:700;"
                  data-action="open-event-einladung" data-event-name="${helpers.escapeHtml(group.name)}" data-list-label="${helpers.escapeHtml(listLabel)}">
                  ${helpers.escapeHtml(listLabel)}
                </button>
                ${isAnlass ? `
                <button class="bbz-button" style="height:30px;font-size:11px;font-weight:700;background:#edf7f1;border-color:#a8dbb8;color:var(--green);"
                  data-action="open-event-nachbearbeitung" data-event-name="${helpers.escapeHtml(group.name)}">
                  Nachbearbeitung
                </button>` : ""}
              </div>
            </div>
          </div>`;
      }).join("");

      return `
        <div>
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:8px;">
            <div style="font-size:11px;color:var(--muted);">
              ${allGroups.length} Event${allGroups.length !== 1 ? "s" : ""} · ${totalActiveContacts} aktive Kontakte
            </div>
            <button class="bbz-button bbz-button-primary" data-action="open-event-matrix" style="font-weight:700;">
              🗂 Event-Management
            </button>
          </div>
          <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:14px;">
            ${cardHtml || ui.emptyBlock("Keine Events vorhanden.")}
          </div>
        </div>
      `;
    },

    // Modal: Einladungsliste / Versandliste
    renderEventEinladungModal(payload = {}) {
      const { eventName = "", listLabel = "Einladungsliste" } = payload;
      const group = state.enriched.events.find(g => g.name === eventName);
      if (!group) return "";

      const contacts = group.contacts;
      const firmIds = new Set(contacts.map(c => c.firmId).filter(Boolean));
      const cntA = contacts.filter(c => c.segment.startsWith("A")).length;
      const cntB = contacts.filter(c => c.segment.startsWith("B")).length;
      const cntC = contacts.filter(c => c.segment.startsWith("C")).length;
      const totalActive = state.enriched.contacts.filter(c => !c.archiviert).length;
      const pct = totalActive > 0 ? Math.round((contacts.length / totalActive) * 100) : 0;

      const filterSeg = payload.filterSeg || "";
      const filterSearch = payload.filterSearch || "";

      let filtered = contacts;
      if (filterSeg) filtered = filtered.filter(c => c.segment.startsWith(filterSeg));
      if (filterSearch.trim()) {
        const s = filterSearch.trim().toLowerCase();
        filtered = filtered.filter(c => helpers.textIncludes(c.contactName, s) || helpers.textIncludes(c.firmTitle, s));
      }

      const rowsHtml = filtered.length ? filtered.map(item => {
        const av = helpers.avatarHtml({ vorname: (item.contactName||"").split(" ")[0]||"", nachname: (item.contactName||"").split(" ").slice(-1)[0]||"" });
        return `
          <tr>
            <td style="min-width:180px;max-width:260px;">
              <div style="display:flex;align-items:center;gap:10px;">
                ${av}
                <div style="min-width:0;overflow:hidden;">
                  <div style="font-weight:600;font-size:13px;line-height:1.35;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${helpers.escapeHtml(item.contactName)}</div>
                  <div style="font-size:11px;color:var(--muted);line-height:1.35;margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${helpers.escapeHtml(item.firmTitle||"—")}</div>
                </div>
              </div>
            </td>
            <td class="bbz-desktop-only" style="font-size:12px;color:var(--subtle);max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${helpers.escapeHtml(item.funktion||item.rolle||"—")}</td>
            <td style="white-space:nowrap;">${item.segment ? `<span class="${helpers.firmBadgeClass(item.segment)}">${helpers.escapeHtml(item.segment.charAt(0))}</span>` : '<span class="bbz-muted">—</span>'}</td>
            <td class="bbz-desktop-only" style="white-space:nowrap;">${item.leadbbz ? helpers.leadbbzBadgeHtml(item.leadbbz) : '<span class="bbz-muted">—</span>'}</td>
            <td style="text-align:right;white-space:nowrap;">
              <button class="bbz-button bbz-button-secondary" style="height:26px;font-size:11px;padding:0 8px;color:var(--red);border-color:#f2b8ba;background:var(--red-soft);"
                data-action="event-remove-contact" data-event-name="${helpers.escapeHtml(eventName)}" data-contact-id="${item.contactId}">✕</button>
            </td>
          </tr>`;
      }).join("") : `<tr><td colspan="5">${ui.emptyBlock("Keine Kontakte für diese Filterung.")}</td></tr>`;

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal" style="max-width:860px;width:95vw;">
            <!-- Header -->
            <div class="bbz-modal-header">
              <div style="width:32px;height:32px;border-radius:var(--r-md);background:var(--blue-light);display:flex;align-items:center;justify-content:center;font-size:15px;flex-shrink:0;">📋</div>
              <div style="flex:1;min-width:0;">
                <div class="bbz-modal-title">${helpers.escapeHtml(eventName)} — ${helpers.escapeHtml(listLabel)}</div>
                <div style="font-size:11px;color:var(--muted);margin-top:1px;">Kontakte verwalten · Excel-Export</div>
              </div>
              <button class="bbz-button bbz-button-secondary" style="height:28px;width:28px;padding:0;" data-close-modal>✕</button>
            </div>

            <!-- Stats Bar -->
            <div style="display:flex;background:#0d1f35;flex-shrink:0;overflow-x:auto;">
              ${[
                { label:"Firmen",     value:firmIds.size,  color:"#60a5fa", hint:"alle",    seg:"",  always:true  },
                { label:"Einladungen",value:contacts.length,color:"#fff",  hint:"alle",    seg:"",  always:true  },
                { label:"A-Kunden",   value:cntA,          color:"#60a5fa", hint:"filtern", seg:"A", always:false },
                { label:"B-Kunden",   value:cntB,          color:"#fbbf24", hint:"filtern", seg:"B", always:false },
                { label:"C-Kunden",   value:cntC,          color:"rgba(255,255,255,0.38)", hint:"filtern", seg:"C", always:false },
                { label:"Reichweite", value:pct+"%",       color:"#22d98a", hint:"gesamt",  seg:"",  always:true  }
              ].map(s => `
                <div style="flex:1;min-width:72px;padding:10px 12px;border-right:1px solid rgba(255,255,255,0.07);text-align:center;cursor:pointer;${!s.always ? "display:none;" : ""}"
                  class="${s.always ? "" : "bbz-desktop-only-flex"}"
                  data-action="event-stat-filter" data-event-name="${helpers.escapeHtml(eventName)}" data-seg="${s.seg}">
                  <div style="font-size:9px;font-weight:700;color:rgba(255,255,255,0.25);text-transform:uppercase;letter-spacing:0.09em;margin-bottom:3px;white-space:nowrap;">${s.label}</div>
                  <div style="font-size:17px;font-weight:700;letter-spacing:-0.04em;color:${s.color};">${s.value}</div>
                  <div style="font-size:9px;color:rgba(255,255,255,0.18);margin-top:2px;">${s.hint}</div>
                </div>`).join("")}
            </div>

            <!-- Filter -->
            <div style="display:flex;gap:8px;padding:10px 16px;border-bottom:1px solid var(--line-2);flex-shrink:0;align-items:center;">
              <input class="bbz-input" style="flex:1;min-width:0;height:30px;" data-filter="event-einladung-search"
                type="text" placeholder="Name oder Firma …" value="${helpers.escapeHtml(filterSearch)}" />
              <select class="bbz-select" style="height:30px;flex-shrink:0;width:130px;" data-filter="event-einladung-seg">
                <option value="" ${!filterSeg ? "selected" : ""}>— Segment —</option>
                <option value="A" ${filterSeg==="A" ? "selected" : ""}>A</option>
                <option value="B" ${filterSeg==="B" ? "selected" : ""}>B</option>
                <option value="C" ${filterSeg==="C" ? "selected" : ""}>C</option>
              </select>
            </div>

            <!-- Tabelle -->
            <div class="bbz-modal-body" style="padding:0;flex:1;overflow-y:auto;min-height:0;">
              <div class="bbz-table-wrap" style="border:none;border-radius:0;">
                <table class="bbz-table">
                  <thead><tr>
                    <th>Kontakt</th>
                    <th class="bbz-desktop-only">Funktion</th>
                    <th>Seg.</th>
                    <th class="bbz-desktop-only">Lead BBZ</th>
                    <th></th>
                  </tr></thead>
                  <tbody>${rowsHtml}</tbody>
                </table>
              </div>
            </div>

            <!-- Footer -->
            <div class="bbz-modal-footer" style="display:flex;align-items:center;justify-content:space-between;padding:11px 16px;border-top:1px solid var(--line);background:var(--panel-2);flex-shrink:0;gap:8px;flex-wrap:wrap;">
              <div style="display:flex;gap:8px;flex-wrap:wrap;">
                <button class="bbz-button" data-action="event-export-excel" data-event-name="${helpers.escapeHtml(eventName)}">
                  ↓ Excel
                </button>
              </div>
              <button class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
          </div>
        </div>`;
    },

    // Modal: Nachbearbeitung (Teilnahme markieren)
    renderEventNachbearbeitungModal(payload = {}) {
      const { eventName = "" } = payload;
      const group = state.enriched.events.find(g => g.name === eventName);
      if (!group) return "";

      const contacts = group.contacts;
      const firmIds = new Set(contacts.map(c => c.firmId).filter(Boolean));
      const totalActive = state.enriched.contacts.filter(c => !c.archiviert).length;
      const pct = totalActive > 0 ? Math.round((contacts.length / totalActive) * 100) : 0;

      const checkedIds = payload.checkedIds || [];
      const selectedVersion = payload.selectedVersion || "";

      // Eventhistory Choices aus SP laden
      const histChoices = state.meta.choices?.[CONFIG.lists.contacts]?.["Eventhistory"] || [];

      const filterSearch = payload.filterSearch || "";
      let filtered = contacts;
      if (filterSearch.trim()) {
        const s = filterSearch.trim().toLowerCase();
        filtered = filtered.filter(c => helpers.textIncludes(c.contactName, s) || helpers.textIncludes(c.firmTitle, s));
      }

      const rowsHtml = filtered.map(item => {
        const isChecked = checkedIds.includes(item.contactId);
        const av = helpers.avatarHtml({ vorname: (item.contactName||"").split(" ")[0]||"", nachname: (item.contactName||"").split(" ").slice(-1)[0]||"" });
        const histBadges = helpers.toArray(item.eventhistory)
          .filter(Boolean)
          .map(h => `<span class="bbz-pill" style="font-size:10px;background:#edf7f1;color:var(--green);border:1px solid #a8dbb8;margin-right:2px;">${helpers.escapeHtml(h)}</span>`).join("");
        return `
          <tr style="${isChecked ? "background:#f0fdf4;" : ""}">
            <td style="text-align:center;width:40px;white-space:nowrap;">
              <button class="bbz-event-check-btn ${isChecked ? "checked" : ""}"
                data-action="event-nb-toggle" data-contact-id="${item.contactId}">✓</button>
            </td>
            <td style="min-width:180px;max-width:260px;">
              <div style="display:flex;align-items:center;gap:10px;">
                ${av}
                <div style="min-width:0;overflow:hidden;">
                  <div style="font-weight:600;font-size:13px;line-height:1.35;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${helpers.escapeHtml(item.contactName)}</div>
                  <div style="font-size:11px;color:var(--muted);line-height:1.35;margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${helpers.escapeHtml(item.firmTitle||"—")}</div>
                </div>
              </div>
            </td>
            <td class="bbz-desktop-only" style="font-size:12px;color:var(--subtle);max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${helpers.escapeHtml(item.funktion||item.rolle||"—")}</td>
            <td style="white-space:nowrap;">${item.segment ? `<span class="${helpers.firmBadgeClass(item.segment)}">${helpers.escapeHtml(item.segment)}</span>` : '<span class="bbz-muted">—</span>'}</td>
            <td class="bbz-desktop-only" style="white-space:nowrap;">${histBadges || '<span class="bbz-muted">—</span>'}</td>
          </tr>`;
      }).join("") || `<tr><td colspan="5">${ui.emptyBlock("Keine Kontakte gefunden.")}</td></tr>`;

      const versionOptions = histChoices.length
        ? histChoices.map(c => `<option value="${helpers.escapeHtml(c)}" ${selectedVersion===c?"selected":""}>${helpers.escapeHtml(c)}</option>`).join("")
        : `<option value="${helpers.escapeHtml(selectedVersion)}">${helpers.escapeHtml(selectedVersion||"—")}</option>`;

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal" style="max-width:860px;width:95vw;">
            <!-- Header -->
            <div class="bbz-modal-header">
              <div style="width:32px;height:32px;border-radius:var(--r-md);background:#edf7f1;display:flex;align-items:center;justify-content:center;font-size:15px;flex-shrink:0;">✅</div>
              <div style="flex:1;min-width:0;">
                <div class="bbz-modal-title">${helpers.escapeHtml(eventName)} — Nachbearbeitung</div>
                <div style="font-size:11px;color:var(--muted);margin-top:1px;">Wer war dabei? Teilnahme in Eventhistory speichern.</div>
              </div>
              <button class="bbz-button bbz-button-secondary" style="height:28px;width:28px;padding:0;" data-close-modal>✕</button>
            </div>

            <!-- Stats Bar -->
            <div style="display:flex;background:#0d1f35;flex-shrink:0;overflow-x:auto;">
              <div class="bbz-event-stat-bar">
                <div class="bbz-event-stat-bar-label">Firmen</div>
                <div class="bbz-event-stat-bar-value" style="color:#60a5fa;">${firmIds.size}</div>
                <div class="bbz-event-stat-bar-hint">gesamt</div>
              </div>
              <div class="bbz-event-stat-bar">
                <div class="bbz-event-stat-bar-label">Eingeladen</div>
                <div class="bbz-event-stat-bar-value" style="color:#fff;">${contacts.length}</div>
                <div class="bbz-event-stat-bar-hint">total</div>
              </div>
              <div class="bbz-event-stat-bar">
                <div class="bbz-event-stat-bar-label">Markiert</div>
                <div class="bbz-event-stat-bar-value" style="color:#22d98a;" data-nb-marked>${checkedIds.length}</div>
                <div class="bbz-event-stat-bar-hint">dabei</div>
              </div>
              <div class="bbz-event-stat-bar bbz-desktop-only">
                <div class="bbz-event-stat-bar-label">Ausstehend</div>
                <div class="bbz-event-stat-bar-value" style="color:#fbbf24;">${contacts.length - checkedIds.length}</div>
                <div class="bbz-event-stat-bar-hint">offen</div>
              </div>
              <div class="bbz-event-stat-bar">
                <div class="bbz-event-stat-bar-label">Reichweite</div>
                <div class="bbz-event-stat-bar-value" style="color:#22d98a;">${pct}%</div>
                <div class="bbz-event-stat-bar-hint">gesamt</div>
              </div>
            </div>

            <!-- Version + Filter -->
            <div style="display:flex;gap:8px;padding:10px 16px;border-bottom:1px solid var(--line-2);background:var(--panel-2);flex-shrink:0;flex-wrap:wrap;align-items:center;">
              <span style="font-size:12px;font-weight:600;color:var(--muted);">Eventhistory-Wert:</span>
              <select class="bbz-select" style="height:28px;font-weight:700;color:var(--green);background:#edf7f1;border-color:#a8dbb8;"
                data-filter="event-nb-version">${versionOptions}</select>
              <span style="font-size:11px;color:var(--subtle);">← Einmal wählen, dann Checkboxen setzen</span>
              <input class="bbz-input" style="flex:1;min-width:120px;height:28px;margin-left:auto;" data-filter="event-nb-search"
                type="text" placeholder="Name oder Firma …" value="${helpers.escapeHtml(filterSearch)}" />
            </div>

            <!-- Tabelle -->
            <div class="bbz-modal-body" style="padding:0;flex:1;overflow-y:auto;min-height:0;">
              <div class="bbz-table-wrap" style="border:none;border-radius:0;">
                <table class="bbz-table">
                  <thead><tr>
                    <th style="width:44px;">Dabei</th>
                    <th>Kontakt</th>
                    <th class="bbz-desktop-only">Funktion</th>
                    <th>Seg.</th>
                    <th class="bbz-desktop-only">Bereits dabei</th>
                  </tr></thead>
                  <tbody>${rowsHtml}</tbody>
                </table>
              </div>
            </div>

            <!-- Footer -->
            <div class="bbz-modal-footer" style="display:flex;align-items:center;justify-content:space-between;padding:11px 16px;border-top:1px solid var(--line);background:var(--panel-2);flex-shrink:0;gap:8px;flex-wrap:wrap;">
              <button class="bbz-button bbz-button-primary" ${checkedIds.length === 0 || !selectedVersion ? "disabled" : ""}
                data-action="event-nb-save" data-event-name="${helpers.escapeHtml(eventName)}">
                ✓ Teilnahmen speichern (${checkedIds.length})
              </button>
              <button class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
          </div>
        </div>`;
    },

    // Modal: Event-Matrix — alle Kontakte × alle Events als Checkbox-Grid
    renderEventMatrixModal(payload = {}) {
      const {
        filterSearch = "",
        filterFirmId = "",
        filterLeadbbz = "",
        filterSegment = "",
        sortBy = "firm",   // 'contact' | 'firm' | 'segment' | 'leadbbz'
        sortDir = "asc",   // 'asc' | 'desc'
        pendingChanges = {}  // Struktur: { [contactId]: { [eventName]: boolean } }
      } = payload;

      // Anlass- vs. Versand-Events (gleiche Liste wie in events()-View)
      const EVENTS_MIT_NACHBEARBEITUNG = ["BOL", "SummerConv."];
      const eventChoices = state.meta.choices?.[CONFIG.lists.contacts]?.["Event"] || [];

      // Sortierung: Anlässe zuerst, dann Versand
      const sortedEvents = [
        ...eventChoices.filter(e => EVENTS_MIT_NACHBEARBEITUNG.includes(e)),
        ...eventChoices.filter(e => !EVENTS_MIT_NACHBEARBEITUNG.includes(e))
      ];

      // Filter-Optionen aufbauen
      const firmMap = new Map(state.enriched.firms.map(f => [f.id, f]));
      const allFirms = [...state.enriched.firms].sort((a, b) =>
        (a.title || "").localeCompare(b.title || "", "de"));
      const allLeadbbz = [...new Set(state.enriched.contacts.map(c => c.leadbbz0).filter(Boolean))].sort();

      // Kontakte filtern (alle aktiven — KEIN 200er-Limit)
      let rows = state.enriched.contacts.filter(c => !c.archiviert);
      if (filterFirmId) rows = rows.filter(c => String(c.firmId) === String(filterFirmId));
      if (filterLeadbbz) rows = rows.filter(c => c.leadbbz0 === filterLeadbbz);
      if (filterSegment) {
        rows = rows.filter(c => String(firmMap.get(c.firmId)?.klassifizierung || "").toUpperCase().startsWith(filterSegment));
      }
      if (filterSearch.trim()) {
        const s = filterSearch.trim().toLowerCase();
        rows = rows.filter(c => [c.fullName, c.firmTitle].some(v => helpers.textIncludes(v, s)));
      }

      // Konfigurierbare Sortierung — Sekundärschlüssel = Name, damit innerhalb einer Firma alphabetisch
      const dir = sortDir === "desc" ? -1 : 1;
      const cmpStr = (a, b) => (a || "").localeCompare(b || "", "de");
      const segOf = (c) => String(firmMap.get(c.firmId)?.klassifizierung || "").toUpperCase().charAt(0) || "Z";
      rows.sort((a, b) => {
        let primary = 0;
        switch (sortBy) {
          case "firm":     primary = cmpStr(a.firmTitle, b.firmTitle); break;
          case "segment":  primary = cmpStr(segOf(a), segOf(b)); break;
          case "leadbbz":  primary = cmpStr(a.leadbbz0, b.leadbbz0); break;
          case "contact":
          default:         primary = cmpStr(a.fullName, b.fullName); break;
        }
        if (primary !== 0) return primary * dir;
        // Sekundär: immer Name aufsteigend (lesefreundlich)
        return cmpStr(a.fullName, b.fullName);
      });

      // Hilfsfunktion: aktueller (effektiver) Status einer Zelle
      const isChecked = (contact, evName) => {
        const pending = pendingChanges[contact.id]?.[evName];
        if (pending !== undefined) return pending;
        return helpers.toArray(contact.event).includes(evName);
      };

      // Hilfsfunktion: ist Zelle vom Originalstatus abweichend?
      const isDirty = (contact, evName) => {
        const pending = pendingChanges[contact.id]?.[evName];
        if (pending === undefined) return false;
        const original = helpers.toArray(contact.event).includes(evName);
        return pending !== original;
      };

      // Anzahl ausstehender Änderungen zählen (nur echt dirty)
      let dirtyCount = 0;
      Object.entries(pendingChanges).forEach(([cid, evMap]) => {
        const contact = state.enriched.contacts.find(c => String(c.id) === String(cid));
        if (!contact) return;
        const orig = helpers.toArray(contact.event);
        Object.entries(evMap).forEach(([evName, val]) => {
          if (orig.includes(evName) !== val) dirtyCount++;
        });
      });

      // Anzahl betroffener Kontakte
      const dirtyContactCount = new Set(
        Object.entries(pendingChanges).flatMap(([cid, evMap]) => {
          const contact = state.enriched.contacts.find(c => String(c.id) === String(cid));
          if (!contact) return [];
          const orig = helpers.toArray(contact.event);
          return Object.entries(evMap).some(([evName, val]) => orig.includes(evName) !== val) ? [cid] : [];
        })
      ).size;

      // Event-Spalten-Header
      const eventHeadersHtml = sortedEvents.map(evName => {
        const isAnlass = EVENTS_MIT_NACHBEARBEITUNG.includes(evName);
        const bg = isAnlass ? "var(--blue-light)" : "#f3e8ff";
        const color = isAnlass ? "var(--blue)" : "#6d1fb8";
        return `<th style="text-align:center;background:${bg};color:${color};font-size:11px;padding:8px 6px;white-space:nowrap;min-width:80px;position:sticky;top:0;z-index:2;">
          ${helpers.escapeHtml(evName)}
        </th>`;
      }).join("");

      // Spalten-Toggle-Header (alle aus aktueller Filterung an/abwählen pro Event)
      const colToggleHtml = sortedEvents.map(evName => {
        // Status: alle gefilterten Zeilen haben dieses Event aktiv?
        const allChecked = rows.length > 0 && rows.every(c => isChecked(c, evName));
        const someChecked = rows.some(c => isChecked(c, evName));
        const indeterminate = someChecked && !allChecked;
        return `<th style="text-align:center;padding:4px 6px;background:var(--panel-2);position:sticky;top:32px;z-index:2;border-bottom:1px solid var(--line-2);">
          <input type="checkbox"
            data-action="matrix-col-toggle"
            data-event-name="${helpers.escapeHtml(evName)}"
            ${allChecked ? "checked" : ""}
            ${indeterminate ? `ref-indeterminate="1"` : ""}
            title="Alle gefilterten für ${helpers.escapeHtml(evName)} ${allChecked ? "abwählen" : "anwählen"}" />
        </th>`;
      }).join("");

      // Tabellen-Body
      const rowsHtml = rows.length ? rows.map(c => {
        const seg = String(firmMap.get(c.firmId)?.klassifizierung || "").toUpperCase().charAt(0);
        const av = helpers.avatarHtml({ vorname: c.vorname || "", nachname: c.nachname || "" });
        const cellsHtml = sortedEvents.map(evName => {
          const checked = isChecked(c, evName);
          const dirty = isDirty(c, evName);
          return `<td style="text-align:center;padding:6px;${dirty ? "background:#fff7e0;" : ""}">
            <input type="checkbox"
              data-action="matrix-cell-toggle"
              data-contact-id="${c.id}"
              data-event-name="${helpers.escapeHtml(evName)}"
              ${checked ? "checked" : ""} />
          </td>`;
        }).join("");
        return `<tr>
          <td style="position:sticky;left:0;background:var(--panel);z-index:1;min-width:180px;max-width:220px;border-right:1px solid var(--line-2);">
            <div style="display:flex;align-items:center;gap:8px;">
              ${av}
              <div style="min-width:0;overflow:hidden;font-weight:600;font-size:12px;line-height:1.3;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${helpers.escapeHtml(c.fullName || "—")}</div>
            </div>
          </td>
          <td style="min-width:160px;max-width:220px;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${helpers.escapeHtml(c.firmTitle || "—")}</td>
          <td style="white-space:nowrap;text-align:center;">${seg ? `<span class="${helpers.firmBadgeClass(seg)}" style="font-size:10px;">${helpers.escapeHtml(seg)}</span>` : '<span class="bbz-muted">—</span>'}</td>
          <td style="white-space:nowrap;font-size:11px;">${c.leadbbz0 ? helpers.leadbbzBadgeHtml(c.leadbbz0) : '<span class="bbz-muted">—</span>'}</td>
          ${cellsHtml}
        </tr>`;
      }).join("") : `<tr><td colspan="${4 + sortedEvents.length}">${ui.emptyBlock("Keine Kontakte für diese Filterung.")}</td></tr>`;

      const firmOptions = [`<option value="">— alle Firmen —</option>`,
        ...allFirms.map(f => `<option value="${f.id}" ${String(filterFirmId) === String(f.id) ? "selected" : ""}>${helpers.escapeHtml(f.title || "—")}</option>`)
      ].join("");
      const leadOptions = [`<option value="">— alle Lead BBZ —</option>`,
        ...allLeadbbz.map(l => `<option value="${helpers.escapeHtml(l)}" ${filterLeadbbz === l ? "selected" : ""}>${helpers.escapeHtml(l)}</option>`)
      ].join("");

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal" style="max-width:1280px;width:97vw;max-height:92vh;">
            <!-- Header -->
            <div class="bbz-modal-header">
              <div style="width:32px;height:32px;border-radius:var(--r-md);background:var(--blue-light);display:flex;align-items:center;justify-content:center;font-size:15px;flex-shrink:0;">🗂</div>
              <div style="flex:1;min-width:0;">
                <div class="bbz-modal-title">Event-Management</div>
                <div style="font-size:11px;color:var(--muted);margin-top:1px;">Alle Kontakte × alle Events — Häkchen setzen oder entfernen, dann speichern</div>
              </div>
              <button class="bbz-button bbz-button-secondary" style="height:28px;width:28px;padding:0;" data-close-modal>✕</button>
            </div>

            <!-- Filter-Zeile -->
            <div style="display:grid;grid-template-columns:1.5fr 1.2fr 1fr 0.7fr;gap:8px;padding:10px 16px;border-bottom:1px solid var(--line-2);flex-shrink:0;align-items:center;">
              <input class="bbz-input" style="height:30px;" data-filter="matrix-search"
                type="text" placeholder="Name oder Firma suchen …" value="${helpers.escapeHtml(filterSearch)}" />
              <select class="bbz-select" style="height:30px;" data-filter="matrix-firm">${firmOptions}</select>
              <select class="bbz-select" style="height:30px;" data-filter="matrix-leadbbz">${leadOptions}</select>
              <select class="bbz-select" style="height:30px;" data-filter="matrix-segment">
                <option value="" ${!filterSegment ? "selected" : ""}>— Segment —</option>
                <option value="A" ${filterSegment === "A" ? "selected" : ""}>A</option>
                <option value="B" ${filterSegment === "B" ? "selected" : ""}>B</option>
                <option value="C" ${filterSegment === "C" ? "selected" : ""}>C</option>
              </select>
            </div>

            <!-- Stats -->
            <div style="padding:6px 16px;border-bottom:1px solid var(--line-2);font-size:11px;color:var(--muted);flex-shrink:0;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">
              <div>
                <strong style="color:var(--text);">${rows.length}</strong> Kontakt${rows.length !== 1 ? "e" : ""} angezeigt
                · <strong style="color:var(--text);">${sortedEvents.length}</strong> Event-Spalten
                ${filterSearch || filterFirmId || filterLeadbbz || filterSegment
                  ? `<button class="bbz-button bbz-button-secondary" style="height:22px;font-size:10px;padding:0 8px;margin-left:8px;" data-action="matrix-clear-filters">Filter zurücksetzen</button>`
                  : ""}
              </div>
              ${dirtyCount > 0
                ? `<div style="color:var(--blue);font-weight:600;">${dirtyCount} Änderung${dirtyCount !== 1 ? "en" : ""} ausstehend (${dirtyContactCount} Kontakt${dirtyContactCount !== 1 ? "e" : ""})</div>`
                : `<div>Keine ausstehenden Änderungen</div>`}
            </div>

            <!-- Tabelle -->
            <div class="bbz-modal-body" style="padding:0;flex:1;overflow:auto;min-height:0;">
              <table class="bbz-table" style="border-collapse:separate;border-spacing:0;">
                <thead>
                  <tr>
                    ${[
                      { key: "contact",  label: "Kontakt",  sticky: "left:0;",           extra: "min-width:180px;border-right:1px solid var(--line-2);", z: 3 },
                      { key: "firm",     label: "Firma",    sticky: "",                  extra: "min-width:160px;",                                       z: 2 },
                      { key: "segment",  label: "Seg.",     sticky: "",                  extra: "text-align:center;",                                      z: 2 },
                      { key: "leadbbz",  label: "Lead BBZ", sticky: "",                  extra: "",                                                        z: 2 }
                    ].map(h => {
                      const active = sortBy === h.key;
                      const arrow = active ? (sortDir === "asc" ? " ▲" : " ▼") : "";
                      return `<th style="position:sticky;top:0;${h.sticky}background:var(--panel-2);z-index:${h.z};${h.extra}cursor:pointer;user-select:none;${active ? "color:var(--blue);" : ""}"
                        data-action="matrix-sort" data-sort-key="${h.key}" title="Klick: nach ${helpers.escapeHtml(h.label)} sortieren">
                        ${helpers.escapeHtml(h.label)}${arrow}
                      </th>`;
                    }).join("")}
                    ${eventHeadersHtml}
                  </tr>
                  <tr>
                    <th colspan="4" style="position:sticky;top:32px;left:0;background:var(--panel-2);z-index:3;font-size:10px;text-align:right;padding-right:10px;color:var(--muted);font-weight:400;border-right:1px solid var(--line-2);">↓ Spalte: alle gefilterten an/aus</th>
                    ${colToggleHtml}
                  </tr>
                </thead>
                <tbody>${rowsHtml}</tbody>
              </table>
            </div>

            <!-- Footer -->
            <div class="bbz-modal-footer" style="display:flex;align-items:center;justify-content:space-between;padding:11px 16px;border-top:1px solid var(--line);background:var(--panel-2);flex-shrink:0;gap:8px;flex-wrap:wrap;">
              <div style="display:flex;gap:8px;align-items:center;">
                <button class="bbz-button bbz-button-primary" ${dirtyCount === 0 ? "disabled" : ""}
                  data-action="matrix-save">
                  ✓ ${dirtyCount > 0 ? `${dirtyCount} Änderung${dirtyCount !== 1 ? "en" : ""} speichern` : "Speichern"}
                </button>
                ${dirtyCount > 0
                  ? `<button class="bbz-button bbz-button-secondary" data-action="matrix-discard">Verwerfen</button>`
                  : ""}
              </div>
              <button class="bbz-button bbz-button-secondary" data-close-modal>Schliessen</button>
            </div>
          </div>
        </div>`;
    }
  };

  const controller = {
    async init() {
      ui.init();
      ui.renderShell();
      ui.setMessage("");
      ui.renderView(ui.loadingBlock("Authentifizierung wird vorbereitet ..."));

      try {
        ui.setLoading(true);
        await api.initAuth();

        if (state.auth.isAuthenticated) {
          // acquireToken() wird von graphRequest() intern aufgerufen — kein separater Call nötig
          await Promise.all([api.loadAll(), api.loadColumnChoices()]);
          ui.setMessage("Anmeldung erkannt. Daten wurden geladen.", "success");
        } else {
          ui.setMessage("Bitte anmelden, um die SharePoint-Listen ueber Microsoft Graph zu laden.", "warning");
        }
      } catch (error) {
        console.error(error);
        state.meta.lastError = error;
        ui.setMessage(`Fehler beim Initialisieren: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleLogin() {
      try {
        if (!state.auth.isReady) { ui.setMessage("Authentifizierung ist noch nicht bereit. Bitte Seite neu laden.", "warning"); return; }
        ui.setLoading(true);
        ui.setMessage("");
        await api.login();
        // Consent-Probe sequenziell VOR Promise.all() —
        // verhindert parallele 403-Flut + interaction_in_progress bei fehlendem Consent
        await api.ensureConsent();
        await Promise.all([api.loadAll(), api.loadColumnChoices()]);
        ui.setMessage("Anmeldung erfolgreich. Daten wurden geladen.", "success");
      } catch (error) {
        console.error(error);
        // Lesbare Fehlermeldungen statt roher JSON-Blobs
        let msg = error.message || "Unbekannter Fehler.";
        if (msg.includes("accessDenied") || msg.includes("Access denied")) {
          msg = "Zugriff verweigert (403). Fehlende Berechtigung für SharePoint. Bitte erneut anmelden — ein Consent-Dialog sollte erscheinen.";
        } else if (msg.includes("interaction_in_progress")) {
          msg = "Ein anderer Login-Vorgang läuft noch. Bitte Seite neu laden.";
        } else if (msg.includes("Graph 403")) {
          msg = "Kein Zugriff auf SharePoint. Bitte den Administrator kontaktieren (App-Berechtigung prüfen).";
        }
        ui.setMessage(`Anmeldung fehlgeschlagen: ${msg}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleRefresh() {
      if (!state.auth.isReady) { ui.setMessage("Authentifizierung ist noch nicht bereit.", "warning"); return; }
      if (!state.auth.isAuthenticated) { ui.setMessage("Bitte zuerst anmelden.", "warning"); return; }
      try {
        ui.setLoading(true);
        ui.setMessage("");
        await api.acquireToken();
        // Consent-Probe auch bei Refresh — Browser-Profil könnte gewechselt haben
        await api.ensureConsent();
        // Refresh: Choices ebenfalls neu laden — SP-Schema könnte sich geändert haben
        await Promise.all([api.loadAll(), api.loadColumnChoices()]);
        ui.setMessage("Daten erfolgreich neu geladen.", "success");
      } catch (error) {
        console.error(error);
        ui.setMessage(`Fehler beim Laden: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    // FIX 2d: Modal oeffnen
    openContactForm(itemId = null, prefillFirmId = null) {
      state.modal = {
        type: "contact",
        mode: itemId ? "edit" : "create",
        payload: { itemId, prefillFirmId }
      };
      this.render();
    },

    openFirmForm(firmId = null) {
      state.modal = {
        type: "firm",
        mode: firmId ? "edit" : "create",
        payload: { firmId: firmId ? Number(firmId) : null }
      };
      this.render();
    },

    async handleFirmModalSubmit(form, mode, itemId) {
      const fd = new FormData(form);

      if (!fd.get("title")?.trim()) {
        ui.setMessage("Firmenname ist ein Pflichtfeld.", "error");
        return;
      }

      const fields = {
        Title: fd.get("title").trim(),
        VIP:   form.querySelector("[name='vip']")?.checked ?? false,
        // Immer senden — null löscht den Wert in SP
        Adresse:        fd.get("adresse")?.trim()      || null,
        PLZ:            fd.get("plz")?.trim()           || null,
        Ort:            fd.get("ort")?.trim()            || null,
        Land:           fd.get("land")?.trim()           || null,
        Hauptnummer:    fd.get("hauptnummer")?.trim()    || null,
        Klassifizierung: fd.get("klassifizierung")      || "",
      };


      ui.setLoading(true);
      ui.setMessage("");

      try {
        if (mode === "create") {
          await api.postItem(SCHEMA.firms.listTitle, fields);
          ui.setMessage("Firma wurde erfolgreich angelegt.", "success");
        } else {
          if (!itemId) throw new Error("itemId fehlt für PATCH.");
          await api.patchItem(SCHEMA.firms.listTitle, Number(itemId), fields);
          ui.setMessage("Firma wurde erfolgreich gespeichert.", "success");
        }
        await api.loadAll();
        this.closeModal();
      } catch (error) {
        console.error("handleFirmModalSubmit Fehler:", error);
        let msg = error.message || "Unbekannter Fehler";
        if (msg.includes("400")) msg = "Fehler 400: Ungültige Felddaten.";
        if (msg.includes("403")) msg = "Fehler 403: Keine Schreibberechtigung.";
        ui.setMessage(msg, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleDeleteContact(id, name) {
      if (!confirm(`Kontakt "${name}" wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.`)) return;
      try {
        ui.setLoading(true);
        await api.deleteItem(SCHEMA.contacts.listTitle, Number(id));
        ui.setMessage(`Kontakt "${name}" wurde gelöscht.`, "success");
        state.selection.contactId = null;
        state.filters.route = "contacts";
        await api.loadAll();
      } catch (error) {
        console.error("handleDeleteContact:", error);
        ui.setMessage(`Fehler beim Löschen: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleDeleteFirm(id, name) {
      if (!confirm(`Firma "${name}" wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.`)) return;
      try {
        ui.setLoading(true);
        await api.deleteItem(SCHEMA.firms.listTitle, Number(id));
        ui.setMessage(`Firma "${name}" wurde gelöscht.`, "success");
        state.selection.firmId = null;
        state.filters.route = "firms";
        await api.loadAll();
      } catch (error) {
        console.error("handleDeleteFirm:", error);
        ui.setMessage(`Fehler beim Löschen: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    // FIX 2e: Modal schliessen
    closeModal() {
      state.modal = null;
      this.render();
    },

    // Write-Layer: Kontakt speichern (create oder edit)
    async handleModalSubmit(form, mode, itemId) {
      // FormData.entries() gibt bei gleichnamigen Checkboxen nur den letzten Wert zurück.
      // Deshalb getAll() für Multi-Choice-Felder verwenden.
      const fd = new FormData(form);

      const raw = {
        nachname:      fd.get("nachname") || "",
        vorname:       fd.get("vorname") || "",
        anrede:        fd.get("anrede") || "",
        firmaLookupId: fd.get("firmaLookupId") || "",
        funktion:      fd.get("funktion") || "",
        rolle:         fd.get("rolle") || "",
        email1:        fd.get("email1") || "",
        email2:        fd.get("email2") || "",
        direktwahl:    fd.get("direktwahl") || "",
        mobile:        fd.get("mobile") || "",
        geburtstag:    fd.get("geburtstag") || "",
        leadbbz0:      fd.get("leadbbz0") || "",
        kommentar:     fd.get("kommentar") || "",
        // Multi-Choice: getAll() sammelt alle checked Werte
        sgf:           fd.getAll("sgf"),
        event:         fd.getAll("event"),
        eventhistory:  fd.getAll("eventhistory"),
        // Checkbox Archiviert
        archiviert:    form.querySelector("[name='archiviert']")?.checked ?? false
      };

      // Pflichtfeld-Validierung
      if (!raw.nachname.trim()) {
        ui.setMessage("Nachname ist ein Pflichtfeld.", "error");
        return;
      }
      if (!raw.firmaLookupId) {
        ui.setMessage("Bitte eine Firma zuweisen.", "error");
        return;
      }

      // Pflichtfelder — immer senden
      const fields = {
        Title:         raw.nachname.trim(),
        FirmaLookupId: Number(raw.firmaLookupId),
        // Archiviert immer senden — auch false, sonst kann ein archivierter Kontakt nicht reaktiviert werden
        Archiviert:    raw.archiviert
      };

      // Einzelwahl-Choice-Felder — leer = "" (nicht null, SP Choice-Felder akzeptieren "" zuverlässiger)
      fields.Anrede   = raw.anrede   || "";
      fields.Rolle    = raw.rolle    || "";
      fields.Leadbbz0 = raw.leadbbz0 || "";

      // Optionaler Text — immer senden, leer = null zum Löschen in SP
      fields.Vorname    = raw.vorname.trim()    || null;
      fields.Funktion   = raw.funktion.trim()   || null;
      fields.Kommentar  = raw.kommentar.trim()  || null;
      fields.Email1     = raw.email1.trim()     || null;
      fields.Email2     = raw.email2.trim()     || null;
      fields.Direktwahl = raw.direktwahl.trim() || null;
      fields.Mobile     = raw.mobile.trim()     || null;

      // Datum — leer = null zum Löschen
      fields.Geburtstag = raw.geburtstag.trim() ? raw.geburtstag.trim() + "T12:00:00Z" : null;

      // Multi-Choice — @odata.type + Array (befüllen) oder @odata.type + [] (leeren)
      // BESTÄTIGT: @odata.type + Array mit Werten → ✅
      // OFFEN: @odata.type + [] zum Leeren → zu testen
      fields["SGF@odata.type"]          = "Collection(Edm.String)";
      fields["SGF"]                     = raw.sgf;
      fields["Event@odata.type"]        = "Collection(Edm.String)";
      fields["Event"]                   = raw.event;
      fields["Eventhistory@odata.type"] = "Collection(Edm.String)";
      fields["Eventhistory"]            = raw.eventhistory;



      ui.setLoading(true);
      ui.setMessage("");

      try {
        if (mode === "create") {
          // SharePoint Graph: POST akzeptiert nur Title + Lookup-Felder zuverlässig.
          // Alle weiteren Felder müssen per separatem PATCH auf die neue Item-ID geschrieben werden.
          // BESTÄTIGT: POST mit vollem fields-Objekt speichert nur Title.
          const createFields = {
            Title:         fields.Title,
            FirmaLookupId: fields.FirmaLookupId
          };
          const created = await api.postItem(SCHEMA.contacts.listTitle, createFields);
          const newItemId = created?.id || created?.fields?.id;
          if (!newItemId) throw new Error("Neue Item-ID fehlt im POST-Response.");

          // Restliche Felder per PATCH nachschreiben
          const patchFields = { ...fields };
          delete patchFields.Title;
          delete patchFields.FirmaLookupId;
          if (Object.keys(patchFields).length > 0) {
            await api.patchItem(SCHEMA.contacts.listTitle, Number(newItemId), patchFields);
          }

          ui.setMessage("Kontakt wurde erfolgreich angelegt.", "success");
        } else {
          if (!itemId) throw new Error("itemId fehlt für PATCH.");
          await api.patchItem(SCHEMA.contacts.listTitle, Number(itemId), fields);
          ui.setMessage("Kontakt wurde erfolgreich gespeichert.", "success");
        }

        await api.loadAll();
        this.closeModal();
      } catch (error) {
        console.error("handleModalSubmit Fehler:", error);

        // Vollständigen Graph-Fehlertext extrahieren für sauberes Debugging
        let msg = error.message || "Unbekannter Fehler";
        let detail = "";
        try {
          // Graph-Fehler haben oft JSON im message-String
          const match = msg.match(/\{.*\}/s);
          if (match) {
            const parsed = JSON.parse(match[0]);
            detail = parsed?.error?.message || parsed?.message || "";
          }
        } catch { /* ignore parse error */ }

        if (msg.includes("400")) msg = `Fehler 400: Ungültige Felddaten.${detail ? " " + detail : " Bitte Konsole prüfen."}`;
        if (msg.includes("403")) msg = "Fehler 403: Keine Schreibberechtigung auf diese Liste.";
        if (msg.includes("409")) msg = "Fehler 409: Konflikt — Eintrag wurde zwischenzeitlich geändert.";

        ui.setMessage(msg, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    openBatchEventForm(eventName, mode = "anmelden") {
      state.modal = {
        type: "batch-event",
        payload: {
          eventName,
          mode,
          filterSegment: "",
          filterLeadbbz: "",
          filterSgf: "",
          filterSearch: "",
          selected: [],
          previewContacts: [],
          selectedHistoryCategory: ""
        }
      };
      this.render();
    },

    async handleBatchEventSubmit(form) {
      const mode = form.dataset.mode || "anmelden";
      const isEventhistory = mode === "eventhistory";
      // Für eventhistory: aktive Kategorie aus Payload lesen (Dropdown-Auswahl)
      const eventName = isEventhistory
        ? (state.modal?.payload?.selectedHistoryCategory || "")
        : (form.dataset.eventName || "");

      let selectedIds = [];
      // Direkt aus State lesen — zuverlässiger als hidden Input,
      // da DOM-only Updates beim Checkbox-Klick keinen hidden Input pflegen
      selectedIds = state.modal?.payload?.selected || [];
      if (!selectedIds.length) {
        // Fallback: hidden Input (rückwärtskompatibel)
        try { selectedIds = JSON.parse(form.querySelector("[name='selectedIds']")?.value || "[]"); } catch { /* ignore */ }
      }

      if (!eventName) { ui.setMessage("Bitte eine Kategorie wählen.", "error"); return; }
      if (!selectedIds.length) { ui.setMessage("Keine Kontakte ausgewählt.", "error"); return; }

      ui.setLoading(true);
      ui.setMessage("");

      let ok = 0, fail = 0;
      try {
        const results = await Promise.allSettled(selectedIds.map(async cid => {
          const contact = state.enriched.contacts.find(c => c.id === cid);
          if (!contact) throw new Error(`Kontakt ${cid} nicht gefunden`);

          const currentEvent     = helpers.toArray(contact.event);
          const currentEventHist = helpers.toArray(contact.eventhistory);

          const patchFields = {};
          if (isEventhistory) {
            // Eventhistory-Feld: Kategorie additiv hinzufügen
            if (!currentEventHist.includes(eventName)) {
              patchFields["Eventhistory@odata.type"] = "Collection(Edm.String)";
              patchFields["Eventhistory"] = [...currentEventHist, eventName];
            }
          } else {
            // Event-Feld: Kategorie additiv hinzufügen
            if (!currentEvent.includes(eventName)) {
              patchFields["Event@odata.type"] = "Collection(Edm.String)";
              patchFields["Event"] = [...currentEvent, eventName];
            }
          }
          // Nichts zu tun wenn Flag bereits gesetzt
          if (Object.keys(patchFields).length === 0) return;
          await api.patchItem(SCHEMA.contacts.listTitle, Number(cid), patchFields);
        }));

        results.forEach(r => r.status === "fulfilled" ? ok++ : (fail++, console.error(r.reason)));
        await api.loadAll();
        this.closeModal();

        const fieldLabel = isEventhistory ? "Eventhistory" : "Event";
        const msg = `✓ ${fieldLabel} «${eventName}» für ${ok} Kontakt${ok !== 1 ? "e" : ""} gesetzt${fail > 0 ? ` — ${fail} Fehler (Konsole prüfen)` : ""}.`;
        ui.setMessage(msg, fail > 0 ? "error" : "success");
        if (fail === 0) setTimeout(() => ui.setMessage(""), 3000);

      } catch (error) {
        console.error("handleBatchEventSubmit:", error);
        ui.setMessage(`Fehler: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    // Event-Matrix: alle ausstehenden Änderungen batched speichern
    async handleEventMatrixSave() {
      const pc = state.modal?.payload?.pendingChanges || {};
      const contactIds = Object.keys(pc);
      if (!contactIds.length) { ui.setMessage("Keine Änderungen.", "warning"); return; }

      // Pro Kontakt: finalen Event-Array berechnen (nur Patch wenn echt geändert)
      const patches = [];
      for (const cid of contactIds) {
        const contact = state.enriched.contacts.find(c => String(c.id) === String(cid));
        if (!contact) continue;
        const original = helpers.toArray(contact.event);
        const set = new Set(original);
        let changed = false;
        Object.entries(pc[cid]).forEach(([evName, val]) => {
          const had = set.has(evName);
          if (val && !had) { set.add(evName); changed = true; }
          else if (!val && had) { set.delete(evName); changed = true; }
        });
        if (changed) patches.push({ cid: Number(cid), newArr: [...set] });
      }

      if (!patches.length) {
        ui.setMessage("Keine echten Änderungen zum Speichern.", "warning");
        state.modal.payload.pendingChanges = {};
        this.render();
        return;
      }

      ui.setLoading(true);
      ui.setMessage("");

      // Helper: einzelner Patch mit Retry bei 409/429/503
      const patchWithRetry = async (cid, newArr, attempt = 1) => {
        try {
          await api.patchItem(SCHEMA.contacts.listTitle, cid, {
            "Event@odata.type": "Collection(Edm.String)",
            "Event": newArr
          });
        } catch (err) {
          const msg = String(err?.message || "");
          const retriable = /409|429|503|resourceModified|throttl/i.test(msg);
          if (retriable && attempt < 4) {
            // Exponentielles Backoff: 200ms, 600ms, 1.4s
            const delay = 200 * Math.pow(3, attempt - 1) + Math.random() * 100;
            await new Promise(r => setTimeout(r, delay));
            return patchWithRetry(cid, newArr, attempt + 1);
          }
          throw err;
        }
      };

      // Helper: limitiert parallele Ausführung (max 4 gleichzeitig — SP-freundlich)
      const runPool = async (items, worker, concurrency = 4) => {
        const results = [];
        let idx = 0;
        const workers = Array(Math.min(concurrency, items.length)).fill(0).map(async () => {
          while (idx < items.length) {
            const i = idx++;
            try { results[i] = { status: "fulfilled", value: await worker(items[i]) }; }
            catch (e) { results[i] = { status: "rejected", reason: e }; }
          }
        });
        await Promise.all(workers);
        return results;
      };

      let ok = 0, fail = 0;
      try {
        const results = await runPool(patches, p => patchWithRetry(p.cid, p.newArr), 4);
        results.forEach(r => r.status === "fulfilled" ? ok++ : (fail++, console.error(r.reason)));
        await api.loadAll();
        // Pending-State zurücksetzen, Modal offen lassen (User sieht aktualisierten Stand)
        if (state.modal?.payload) state.modal.payload.pendingChanges = {};
        const msg = `✓ ${ok} Kontakt${ok !== 1 ? "e" : ""} aktualisiert${fail > 0 ? ` — ${fail} Fehler (Konsole prüfen)` : ""}.`;
        ui.setMessage(msg, fail > 0 ? "error" : "success");
        if (fail === 0) setTimeout(() => ui.setMessage(""), 3000);
      } catch (error) {
        console.error("handleEventMatrixSave:", error);
        ui.setMessage(`Fehler: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    openHistoryForm(contactId = null, firmId = null, itemId = null, typ = null) {
      let prefillContactId = contactId;
      if (!prefillContactId && firmId) {
        const firm = dataModel.getFirmById(firmId);
        prefillContactId = firm?.contacts?.[0]?.id || null;
      }
      const mode = itemId ? "edit" : "create";
      state.modal = { type: "history", payload: { prefillContactId, mode, itemId, prefillTyp: typ || "" } };
      this.render();
    },

    openTaskForm(contactId = null, firmId = null, itemId = null) {
      let prefillContactId = contactId;
      if (!prefillContactId && firmId) {
        const firm = dataModel.getFirmById(firmId);
        prefillContactId = firm?.contacts?.[0]?.id || null;
      }
      const mode = itemId ? "edit" : "create";
      state.modal = { type: "task", payload: { prefillContactId, mode, itemId } };
      this.render();
    },

    async handleDeleteHistory(id, title) {
      if (!confirm(`Aktivitaet "${title}" wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.`)) return;
      try {
        ui.setLoading(true);
        await api.deleteItem(SCHEMA.history.listTitle, Number(id));
        ui.setMessage(`Aktivitaet "${title}" wurde gelöscht.`, "success");
        await api.loadAll();
      } catch (error) {
        console.error("handleDeleteHistory:", error);
        ui.setMessage(`Fehler beim Löschen: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleDeleteTask(id, title) {
      if (!confirm(`Aufgabe "${title}" wirklich löschen? Diese Aktion kann nicht rückgängig gemacht werden.`)) return;
      try {
        ui.setLoading(true);
        await api.deleteItem(SCHEMA.tasks.listTitle, Number(id));
        ui.setMessage(`Aufgabe "${title}" wurde gelöscht.`, "success");
        await api.loadAll();
      } catch (error) {
        console.error("handleDeleteTask:", error);
        ui.setMessage(`Fehler beim Löschen: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleHistoryModalSubmit(form) {
      const fd = new FormData(form);
      const mode = form.dataset.mode || "create";
      const itemId = Number(form.dataset.itemId || 0) || null;
      const kontaktLookupId = fd.get("kontaktLookupId") || "";
      const datum = fd.get("datum") || "";

      if (!kontaktLookupId) { ui.setMessage("Bitte einen Kontakt waehlen.", "error"); return; }
      if (!datum) { ui.setMessage("Datum ist ein Pflichtfeld.", "error"); return; }

      const kontaktart = fd.get("kontaktart") || "";
      const leadbbz = fd.get("leadbbz") || "";
      const notizen = fd.get("notizen") || "";
      const projektbezug = form.querySelector("[name='projektbezug']")?.checked ?? false;

      ui.setLoading(true);
      ui.setMessage("");
      try {
        if (mode === "edit") {
          if (!itemId) throw new Error("itemId fehlt fuer PATCH.");
          const patchFields = {
            Datum:      datum + "T12:00:00Z",
            Projektbezug: projektbezug,
            Kontaktart: kontaktart || "",
            Leadbbz:    leadbbz    || "",
            Notizen:    notizen.trim() || null,
          };
          await api.patchItem(SCHEMA.history.listTitle, itemId, patchFields);
          ui.setMessage("Aktivitaet wurde gespeichert.", "success");
        } else {
          // POST: nur Pflichtfelder, dann PATCH mit Rest
          const createFields = {
            Title: `Aktivitaet-${datum}`,
            NachnameLookupId: Number(kontaktLookupId)
          };
          const patchFields = { Datum: datum + "T12:00:00Z", Projektbezug: projektbezug };
          if (kontaktart) patchFields.Kontaktart = kontaktart;
          if (leadbbz) patchFields.Leadbbz = leadbbz;
          if (notizen.trim()) patchFields.Notizen = notizen.trim();
          const created = await api.postItem(SCHEMA.history.listTitle, createFields);
          const newId = created?.id || created?.fields?.id;
          if (!newId) throw new Error("Neue Item-ID fehlt im POST-Response.");
          await api.patchItem(SCHEMA.history.listTitle, Number(newId), patchFields);
          ui.setMessage("Aktivitaet wurde erfasst.", "success");
        }
        await api.loadAll();
        this.closeModal();
      } catch (error) {
        console.error("handleHistoryModalSubmit:", error);
        let msg = error.message || "Unbekannter Fehler";
        if (msg.includes("400")) msg = "Fehler 400: Ungueltige Felddaten. Bitte Konsole pruefen.";
        if (msg.includes("403")) msg = "Fehler 403: Keine Schreibberechtigung.";
        ui.setMessage(msg, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleTaskModalSubmit(form) {
      const fd = new FormData(form);
      const mode = form.dataset.mode || "create";
      const itemId = Number(form.dataset.itemId || 0) || null;
      const title = fd.get("title") || "";
      const kontaktLookupId = fd.get("kontaktLookupId") || "";

      if (!title.trim()) { ui.setMessage("Titel ist ein Pflichtfeld.", "error"); return; }
      if (!kontaktLookupId) { ui.setMessage("Bitte einen Kontakt waehlen.", "error"); return; }

      const deadline = fd.get("deadline") || "";
      const status = fd.get("status") || "";
      const leadbbz = fd.get("leadbbz") || "";

      ui.setLoading(true);
      ui.setMessage("");
      try {
        if (mode === "edit") {
          if (!itemId) throw new Error("itemId fehlt fuer PATCH.");
          const patchFields = {
            Title:    title.trim(),
            Deadline: deadline ? deadline + "T12:00:00Z" : null,
            Status:   status   || "",
            Leadbbz:  leadbbz  || "",
          };
          await api.patchItem(SCHEMA.tasks.listTitle, itemId, patchFields);
          ui.setMessage("Aufgabe wurde gespeichert.", "success");
        } else {
          const createFields = { Title: title.trim(), NameLookupId: Number(kontaktLookupId) };
          const patchFields = {};
          if (deadline) patchFields.Deadline = deadline + "T12:00:00Z";
          if (status) patchFields.Status = status;
          if (leadbbz) patchFields.Leadbbz = leadbbz;
          const created = await api.postItem(SCHEMA.tasks.listTitle, createFields);
          const newId = created?.id || created?.fields?.id;
          if (!newId) throw new Error("Neue Item-ID fehlt im POST-Response.");
          if (Object.keys(patchFields).length > 0) {
            await api.patchItem(SCHEMA.tasks.listTitle, Number(newId), patchFields);
          }
          ui.setMessage("Aufgabe wurde erstellt.", "success");
        }
        await api.loadAll();
        this.closeModal();
      } catch (error) {
        console.error("handleTaskModalSubmit:", error);
        let msg = error.message || "Unbekannter Fehler";
        if (msg.includes("400")) msg = "Fehler 400: Ungueltige Felddaten. Bitte Konsole pruefen.";
        if (msg.includes("403")) msg = "Fehler 403: Keine Schreibberechtigung.";
        ui.setMessage(msg, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    async handleTaskStatusChange(taskId, newStatus) {
      if (!taskId || !newStatus) return;
      try {
        ui.setLoading(true);
        ui.setMessage("");
        await api.patchItem(SCHEMA.tasks.listTitle, taskId, { Status: newStatus });
        const isDone = !helpers.isOpenTask(newStatus);
        if (isDone) {
          ui.setMessage(`✓ Task als „${newStatus}" markiert.`, "success");
          // Auto-dismiss nach 2.5s
          setTimeout(() => { ui.setMessage(""); }, 2500);
        } else {
          ui.setMessage(`Status auf „${newStatus}" gesetzt.`, "success");
        }
        await api.loadAll();
      } catch (error) {
        console.error("handleTaskStatusChange:", error);
        ui.setMessage(`Fehler beim Status-Update: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    navigate(route) {
      state.filters.route = route;
      state.selection.firmId = null;
      state.selection.contactId = null;
      state.modal = null;
      state.filters.events.segment = "";
      state.filters.events.selectedEvent = "";
      history.pushState({ route, firmId: null, contactId: null }, "", `#${route}`);
      window.scrollTo(0, 0);
      this.render();
    },

    openFirm(id) {
      state.selection.firmId = id;
      state.selection.contactId = null;
      state.filters.route = "firms";
      state.modal = null;
      history.pushState({ route: "firms", firmId: id, contactId: null }, "", `#firms-${id}`);
      window.scrollTo(0, 0);
      this.render();
    },

    openContact(id) {
      state.selection.contactId = id;
      state.filters.route = "contacts";
      state.modal = null;
      history.pushState({ route: "contacts", firmId: null, contactId: id }, "", `#contacts-${id}`);
      window.scrollTo(0, 0);
      this.render();
    },

    render() {
      ui.renderShell();
      ui.renderView(views.renderRoute());
    },

    // Event Nachbearbeitung: Teilnahmen speichern
    async handleEventNachbearbeitungSave() {
      const p = state.modal?.payload;
      if (!p) return;
      const { eventName, checkedIds, selectedVersion } = p;
      if (!checkedIds.length || !selectedVersion) return;

      ui.setLoading(true);
      ui.setMessage("");
      let ok = 0, fail = 0;
      try {
        const results = await Promise.allSettled(checkedIds.map(async cid => {
          const contact = state.enriched.contacts.find(c => c.id === cid);
          if (!contact) throw new Error(`Kontakt ${cid} nicht gefunden`);
          const currentHist = helpers.toArray(contact.eventhistory);
          if (currentHist.includes(selectedVersion)) return; // bereits gesetzt
          const patchFields = {
            "Eventhistory@odata.type": "Collection(Edm.String)",
            "Eventhistory": [...currentHist, selectedVersion]
          };
          await api.patchItem(SCHEMA.contacts.listTitle, Number(cid), patchFields);
        }));
        results.forEach(r => r.status === "fulfilled" ? ok++ : fail++);
        const msg = fail === 0
          ? `✓ ${ok} Teilnahme(n) als «${selectedVersion}» gespeichert.`
          : `${ok} gespeichert, ${fail} fehlgeschlagen.`;
        ui.setMessage(msg, fail === 0 ? "success" : "warning");
        await api.loadAll();
        this.closeModal();
      } catch (error) {
        console.error("handleEventNachbearbeitungSave:", error);
        ui.setMessage(`Fehler: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    // Event Einladungsliste: Kontakt entfernen
    async handleEventRemoveContact(eventName, contactId) {
      if (!confirm(`Kontakt aus Event «${eventName}» entfernen?`)) return;
      const contact = state.enriched.contacts.find(c => c.id === contactId);
      if (!contact) return;
      const currentEvent = helpers.toArray(contact.event).filter(e => e !== eventName);
      ui.setLoading(true);
      try {
        await api.patchItem(SCHEMA.contacts.listTitle, contactId, {
          "Event@odata.type": "Collection(Edm.String)",
          "Event": currentEvent
        });
        ui.setMessage(`Kontakt aus «${eventName}» entfernt.`, "success");
        await api.loadAll();
        // Modal-Payload aktualisieren damit der refreshte Stand stimmt
        if (state.modal?.payload) {
          state.modal.payload.filterSearch = state.modal.payload.filterSearch || "";
        }
      } catch (error) {
        console.error("handleEventRemoveContact:", error);
        ui.setMessage(`Fehler: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    // Event Excel-Export
    handleEventExcelExport(eventName) {
      const group = state.enriched.events.find(g => g.name === eventName);
      if (!group) return;
      try {
        const rows = group.contacts.map(c => ({
          "Name":      c.contactName || "",
          "Firma":     c.firmTitle   || "",
          "Funktion":  c.funktion    || c.rolle || "",
          "Segment":   c.segment     || "",
          "Lead BBZ":  c.leadbbz     || "",
          "Email":     c.email1      || ""
        }));
        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, eventName.substring(0, 31));
        XLSX.writeFile(wb, `${eventName}_Liste.xlsx`);
        ui.setMessage(`Excel-Export «${eventName}» erstellt.`, "success");
      } catch (error) {
        console.error("handleEventExcelExport:", error);
        ui.setMessage("Excel-Export fehlgeschlagen. XLSX-Bibliothek verfügbar?", "error");
      }
    }
  };

  window._bbzApp = { state, api, helpers, SCHEMA, CONFIG, dataModel, controller };

  function startApp() { controller.init(); }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startApp, { once: true });
  } else {
    startApp();
  }
})();

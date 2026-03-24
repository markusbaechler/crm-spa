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
        vip: "VIP"
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
      firms: { search: "", klassifizierung: "", vip: "" },
      contacts: { search: "", archiviertAusblenden: CONFIG.defaults.contactArchiveDefaultHidden },
      planning: { search: "", onlyOpen: CONFIG.defaults.planningShowOnlyOpen, onlyOverdue: false },
      events: { search: "", onlyWithOpenTasks: false }
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
    toDateInput(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      return d.toISOString().split("T")[0];
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

      this.els.navButtons.forEach(btn => {
        btn.addEventListener("click", () => controller.navigate(btn.dataset.route));
      });

      // Zentraler Click-Handler
      document.addEventListener("click", (event) => {
        const openFirm = event.target.closest("[data-action='open-firm']");
        if (openFirm) { controller.openFirm(openFirm.dataset.id); return; }

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

        // KEIN separater Handler für [data-modal-submit] nötig:
        // Der Button hat type="submit" und löst den nativen Form-Submit aus,
        // der vom submit-Listener unten abgefangen wird.
        // Ein zusätzlicher dispatchEvent hier würde double-submit verursachen.
      });

      // FIX 2c: Zentraler Form-Submit-Handler — Guard gegen Double-Submit
      document.addEventListener("submit", (event) => {
        const form = event.target.closest("[data-modal-form]");
        if (form) {
          event.preventDefault();
          if (state.meta.loading) return;
          const formType = form.dataset.modalForm;
          if (formType === "firm") {
            controller.handleFirmModalSubmit(form, form.dataset.mode, form.dataset.itemId || null);
          } else {
            controller.handleModalSubmit(form, form.dataset.mode, form.dataset.itemId || null);
          }
        }
      });

      const debouncedRender = helpers.debounce(() => controller.render(), 150);

      document.addEventListener("input", (event) => {
        const el = event.target;
        if (el.matches("[data-filter='firms-search']")) { state.filters.firms.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='contacts-search']")) { state.filters.contacts.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='planning-search']")) { state.filters.planning.search = el.value; debouncedRender(); }
        if (el.matches("[data-filter='events-search']")) { state.filters.events.search = el.value; debouncedRender(); }
      });

      document.addEventListener("change", (event) => {
        const el = event.target;
        if (el.matches("[data-filter='firms-klassifizierung']")) { state.filters.firms.klassifizierung = el.value; controller.render(); }
        if (el.matches("[data-filter='firms-vip']")) { state.filters.firms.vip = el.value; controller.render(); }
        if (el.matches("[data-filter='contacts-archiviert']")) { state.filters.contacts.archiviertAusblenden = el.checked; controller.render(); }
        if (el.matches("[data-filter='planning-open']")) { state.filters.planning.onlyOpen = el.checked; controller.render(); }
        if (el.matches("[data-filter='planning-overdue']")) { state.filters.planning.onlyOverdue = el.checked; controller.render(); }
        if (el.matches("[data-filter='events-open']")) { state.filters.events.onlyWithOpenTasks = el.checked; controller.render(); }
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
      this.els.navButtons.forEach(btn => {
        btn.classList.toggle("active", btn.dataset.route === state.filters.route);
      });

      if (state.auth.isAuthenticated && state.auth.account) {
        this.els.authStatus.innerHTML = `<span class="bbz-auth-dot"></span><span>Angemeldet: ${helpers.escapeHtml(state.auth.account.username || state.auth.account.name || "")}</span>`;
      } else if (state.auth.isReady) {
        this.els.authStatus.innerHTML = `<span class="bbz-auth-dot" style="background:#94a3b8;"></span><span>Nicht angemeldet</span>`;
      } else {
        this.els.authStatus.innerHTML = `<span class="bbz-auth-dot" style="background:#f59e0b;"></span><span>Authentifizierung wird initialisiert ...</span>`;
      }

      if (this.els.btnLogin) {
        this.els.btnLogin.textContent = state.auth.isAuthenticated ? "Erneut anmelden" : "Anmelden";
        this.els.btnLogin.disabled = state.meta.loading || !state.auth.isReady;
      }
      if (this.els.btnRefresh) {
        this.els.btnRefresh.disabled = state.meta.loading || !state.auth.isReady;
      }
    },

    renderView(html) {
      if (this.els.viewRoot) this.els.viewRoot.innerHTML = html;
    },

    loadingBlock(text = "Daten werden geladen ...") {
      return `<section class="bbz-section"><div class="bbz-section-body"><div class="flex items-center gap-3"><div class="bbz-loader"></div><div class="text-sm text-slate-500">${helpers.escapeHtml(text)}</div></div></div></section>`;
    },

    emptyBlock(text = "Keine Daten vorhanden.") {
      return `<div class="bbz-empty">${helpers.escapeHtml(text)}</div>`;
    },

    kv(label, value) {
      return `<div class="bbz-kv"><div class="bbz-kv-label">${helpers.escapeHtml(label)}</div><div class="bbz-kv-value">${value || '<span class="bbz-muted">—</span>'}</div></div>`;
    }
  };

  const api = {
    async initAuth() {
      helpers.ensureMsalAvailable();
      helpers.validateConfig();

      state.auth.isReady = false;
      state.auth.msal = null;

      const msalInstance = new window.msal.PublicClientApplication({
        auth: {
          clientId: CONFIG.graph.clientId,
          authority: CONFIG.graph.authority,
          redirectUri: CONFIG.graph.redirectUri
        },
        cache: { cacheLocation: "localStorage" }
      });

      await msalInstance.initialize();
      state.auth.msal = msalInstance;

      try {
        const redirectResponse = await state.auth.msal.handleRedirectPromise();
        if (redirectResponse?.account) {
          state.auth.account = redirectResponse.account;
          state.auth.isAuthenticated = true;
        }
      } catch (error) {
        console.warn("handleRedirectPromise Fehler", error);
      }

      // Accounts aus Cache nachladen falls kein Redirect-Response
      if (!state.auth.account) {
        const accounts = state.auth.msal.getAllAccounts();
        if (accounts.length > 0) {
          state.auth.account = accounts[0];
          state.auth.isAuthenticated = true;
        }
      }

      state.auth.isReady = true;
    },

    async login() {
      if (!state.auth.msal) throw new Error("MSAL ist nicht initialisiert.");

      const loginResponse = await state.auth.msal.loginPopup({
        scopes: CONFIG.graph.scopes,
        prompt: "select_account"
      });

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
          state.auth.account = accounts[0];
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
        console.warn("Silent token fehlgeschlagen, versuche Popup:", silentError);
        const tokenResponse = await state.auth.msal.acquireTokenPopup({
          account: state.auth.account,
          scopes: CONFIG.graph.scopes
        });
        if (!tokenResponse?.accessToken) throw new Error("Kein Token aus acquireTokenPopup erhalten.");
        state.auth.token = tokenResponse.accessToken;
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
          // Vollständigen Body lesen — gibt bei 400 den exakten SP-Feldnamen
          detail = await response.text();
          console.error(`Graph ${response.status} auf ${options.method || "GET"} ${path}:`, detail);
        } catch { detail = response.statusText; }
        throw new Error(`Graph ${response.status}: ${detail}`);
      }

      if (response.status === 204) return null;
      return await response.json();
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
      const data = await this.graphRequest(`/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/items?expand=fields&top=5000`);
      return data.value || [];
    },

    async loadAll() {
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
              // Vollständiger Debug — zeigt alle relevanten SP-Feldnamen
              console.log(`[${listTitle}] Choice:`, {
                name:        col.name,
                displayName: col.displayName,
                description: col.description,
                multiSelect: col.choice.allowMultipleSelection ?? false,
                choices:     col.choice.choices
              });
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
        { method: "PATCH", body: fields }
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
        vip: helpers.bool(this.getField(item, f.vip))
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
        archiviert: helpers.bool(this.getField(item, f.archiviert))
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
        leadbbz: this.getField(item, f.leadbbz) || ""
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
        leadbbz: this.getField(item, f.leadbbz) || ""
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
            latestHistoryDate: latestH?.datum || "",
            latestHistoryType: latestH?.typ || "",
            latestHistoryText: latestH?.notizen || "",
            openTasksCount: openTasks.length,
            email1: contact.email1
          });
        });
      });

      const events = [...eventMap.values()]
        .map(group => ({ ...group, contactCount: group.contacts.length, openTasksCount: group.contacts.reduce((sum, c) => sum + c.openTasksCount, 0), contacts: group.contacts.sort((a, b) => String(a.contactName).localeCompare(String(b.contactName), "de")) }))
        .sort((a, b) => a.name.localeCompare(b.name, "de"));

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
    kpiBlock(label, value, meta = "") {
      return `<div class="bbz-kpi"><div class="bbz-kpi-label">${helpers.escapeHtml(label)}</div><div class="bbz-kpi-value">${helpers.escapeHtml(String(value))}</div>${meta ? `<div class="bbz-kpi-meta">${helpers.escapeHtml(meta)}</div>` : ""}</div>`;
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
        case "events": viewHtml = this.events(); break;
        default: viewHtml = this.firms();
      }

      // Modal wird ueber dem View gerendert
      let modalHtml = "";
      if (state.modal?.type === "contact") modalHtml = views.renderContactForm(state.modal.mode, state.modal.payload);
      if (state.modal?.type === "firm")    modalHtml = views.renderFirmForm(state.modal.mode, state.modal.payload?.firmId);
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
            <form data-modal-form="contact" data-mode="${mode}" data-item-id="${itemId || ""}">
              <div class="bbz-modal-body">
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
                    <label>${isPrivat ? 'Adresse / Notizen <span class="bbz-field-hint">(Privatperson — Adresse hier erfassen)</span>' : 'Kommentar'}</label>
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
            <form data-modal-form="firm" data-mode="${mode}" data-item-id="${firmId || ""}">
              <div class="bbz-modal-body">
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

    firms() {
      const filters = state.filters.firms;
      const rows = state.enriched.firms.filter(firm => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch = !search || [firm.title, firm.ort, firm.klassifizierung, firm.hauptnummer, firm.adresse, firm.land, ...firm.contacts.map(c => c.fullName)].some(v => helpers.textIncludes(v, search));
        const klassMatch = !filters.klassifizierung || String(firm.klassifizierung || "").toLowerCase() === filters.klassifizierung.toLowerCase();
        const vipMatch = !filters.vip || (filters.vip === "yes" && firm.vip) || (filters.vip === "no" && !firm.vip);
        return searchMatch && klassMatch && vipMatch;
      });

      const aCount = state.enriched.firms.filter(f => String(f.klassifizierung).toUpperCase().includes("A")).length;
      const bCount = state.enriched.firms.filter(f => String(f.klassifizierung).toUpperCase().includes("B")).length;
      const cCount = state.enriched.firms.filter(f => String(f.klassifizierung).toUpperCase().includes("C")).length;
      const overdueTasks = state.enriched.tasks.filter(t => t.isOpen && t.isOverdue).length;
      const urgentFirms = [...state.enriched.firms].filter(f => f.openTasksCount > 0).sort((a, b) => helpers.compareDateAsc(a.nextDeadline, b.nextDeadline)).slice(0, 6);
      const latestFirms = [...state.enriched.firms].filter(f => f.latestActivity).sort((a, b) => helpers.compareDateDesc(a.latestActivity, b.latestActivity)).slice(0, 6);

      return `
        <div>
          <div class="bbz-kpis">
            ${this.kpiBlock("Firmen", state.enriched.firms.length, "gesamt")}
            ${this.kpiBlock("A / B / C", `${aCount} / ${bCount} / ${cCount}`, "Segmente")}
            ${this.kpiBlock("Kontakte", state.enriched.contacts.length, "Ansprechpartner")}
            ${this.kpiBlock("Ueberfaellige Tasks", overdueTasks, "sofort pruefen")}
          </div>
          <div class="bbz-grid bbz-grid-70-30">
            <section class="bbz-section">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Firmen-Cockpit</div><div class="bbz-section-subtitle">Hauptarbeitsliste mit Fokus auf Segment, Tasks und Fristen</div></div></div>
              <div class="bbz-section-body">
                <div class="bbz-filters-3">
                  <input class="bbz-input" data-filter="firms-search" type="text" placeholder="Suche nach Firma, Ort, Ansprechpartner ..." value="${helpers.escapeHtml(filters.search)}" />
                  <select class="bbz-select" data-filter="firms-klassifizierung">
                    <option value="">Alle Klassifizierungen</option>
                    <option value="A-Kunde" ${filters.klassifizierung === "A-Kunde" ? "selected" : ""}>A-Kunde</option>
                    <option value="B-Kunde" ${filters.klassifizierung === "B-Kunde" ? "selected" : ""}>B-Kunde</option>
                    <option value="C-Kunde" ${filters.klassifizierung === "C-Kunde" ? "selected" : ""}>C-Kunde</option>
                    <option value="A" ${filters.klassifizierung === "A" ? "selected" : ""}>A</option>
                    <option value="B" ${filters.klassifizierung === "B" ? "selected" : ""}>B</option>
                    <option value="C" ${filters.klassifizierung === "C" ? "selected" : ""}>C</option>
                  </select>
                  <select class="bbz-select" data-filter="firms-vip">
                    <option value="">VIP egal</option>
                    <option value="yes" ${filters.vip === "yes" ? "selected" : ""}>Nur VIP</option>
                    <option value="no" ${filters.vip === "no" ? "selected" : ""}>Nur nicht VIP</option>
                  </select>
                </div>
                <div class="bbz-table-wrap">
                  <table class="bbz-table">
                    <thead><tr><th>Firma</th><th>Ort</th><th>Klassifizierung</th><th>VIP</th><th>Kontakte</th><th>Offene Tasks</th><th>Naechste Deadline</th></tr></thead>
                    <tbody>
                      ${rows.length ? rows.map(firm => `
                        <tr>
                          <td><a class="bbz-link" data-action="open-firm" data-id="${firm.id}">${helpers.escapeHtml(firm.title)}</a><div class="bbz-subtext">${helpers.escapeHtml(firm.hauptnummer || "—")}</div></td>
                          <td>${helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ")) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : '<span class="bbz-muted">—</span>'}</td>
                          <td>${firm.vip ? '<span class="bbz-pill bbz-pill-vip">VIP</span>' : '<span class="bbz-muted">—</span>'}</td>
                          <td>${firm.contactsCount}</td>
                          <td>${firm.openTasksCount}</td>
                          <td class="${firm.nextDeadline && helpers.isOverdue(firm.nextDeadline) ? "bbz-danger" : ""}">${helpers.formatDate(firm.nextDeadline) || '<span class="bbz-muted">—</span>'}</td>
                        </tr>`).join("") : `<tr><td colspan="7">${ui.emptyBlock("Keine Firmen fuer die aktuelle Filterung gefunden.")}</td></tr>`}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
            <div class="bbz-cockpit-stack">
              <section class="bbz-section">
                <div class="bbz-section-header"><div><div class="bbz-section-title">Dringend</div><div class="bbz-section-subtitle">Firmen mit offenen Tasks</div></div></div>
                <div class="bbz-section-body">${urgentFirms.length ? `<div class="bbz-mini-list">${urgentFirms.map(f => this.miniItem(`<a class="bbz-link" data-action="open-firm" data-id="${f.id}">${helpers.escapeHtml(f.title)}</a>`, `${f.openTasksCount} offene Tasks · naechste Deadline ${helpers.formatDate(f.nextDeadline) || "—"}`)).join("")}</div>` : ui.emptyBlock("Keine dringenden Firmen.")}</div>
              </section>
              <section class="bbz-section">
                <div class="bbz-section-header"><div><div class="bbz-section-title">Zuletzt aktiv</div><div class="bbz-section-subtitle">Firmen mit juengster History</div></div></div>
                <div class="bbz-section-body">${latestFirms.length ? `<div class="bbz-mini-list">${latestFirms.map(f => this.miniItem(`<a class="bbz-link" data-action="open-firm" data-id="${f.id}">${helpers.escapeHtml(f.title)}</a>`, `Letzte Aktivitaet ${helpers.formatDateTime(f.latestActivity) || "—"}`)).join("")}</div>` : ui.emptyBlock("Noch keine Aktivitaeten vorhanden.")}</div>
              </section>
            </div>
          </div>
        </div>
      `;
    },

    firmDetail() {
      const firm = dataModel.getFirmById(state.selection.firmId);
      if (!firm) return ui.emptyBlock("Die ausgewaehlte Firma wurde nicht gefunden.");
      const recentHistory = [...firm.history].slice(0, 20);

      return `
        <div>
          <div class="bbz-detail-header">
            <div>
              <button class="bbz-button bbz-button-secondary mb-3" data-action="back-to-firms">Zurueck zur Firmenliste</button>
              <div class="bbz-detail-title">${helpers.escapeHtml(firm.title)}</div>
              <div class="bbz-detail-subtitle">${helpers.escapeHtml(helpers.joinNonEmpty([firm.adresse, helpers.joinNonEmpty([firm.plz, firm.ort], " "), firm.land], " · ")) || "Keine erweiterten Stammdaten"}</div>
              <div class="flex items-center gap-2 flex-wrap mt-3">
                ${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : ""}
                ${firm.vip ? `<span class="bbz-pill bbz-pill-vip">VIP</span>` : ""}
              </div>
            </div>
            <div class="flex items-center gap-2 flex-wrap">
              <button class="bbz-button bbz-button-secondary" style="${firm.contactsCount > 0 ? "opacity:0.4;cursor:not-allowed;" : "color:var(--red);border-color:var(--red);"}" data-action="delete-firm" data-id="${firm.id}" data-name="${helpers.escapeHtml(firm.title)}" data-contacts="${firm.contactsCount}">Löschen</button>
              <button class="bbz-button bbz-button-secondary" data-action="open-firm-form" data-id="${firm.id}">Bearbeiten</button>
              <button class="bbz-button bbz-button-primary" data-action="open-contact-form" data-firm-id="${firm.id}">+ Kontakt</button>
            </div>
          </div>
          <div class="bbz-kpis">
            ${this.kpiBlock("Kontakte", firm.contactsCount)}
            ${this.kpiBlock("Offene Tasks", firm.openTasksCount)}
            ${this.kpiBlock("Naechste Deadline", helpers.formatDate(firm.nextDeadline) || "—")}
            ${this.kpiBlock("History", firm.history.length)}
          </div>
          <div class="bbz-grid bbz-grid-3">
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Uebersicht</div></div>
              <div class="bbz-section-body"><div class="bbz-meta-grid">
                ${ui.kv("Firma", helpers.escapeHtml(firm.title))}
                ${ui.kv("Klassifizierung", firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : '<span class="bbz-muted">—</span>')}
                ${ui.kv("Adresse", helpers.escapeHtml(firm.adresse) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("PLZ / Ort", helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ")) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Land", helpers.escapeHtml(firm.land) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Hauptnummer", helpers.escapeHtml(firm.hauptnummer) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("VIP", firm.vip ? '<span class="bbz-pill bbz-pill-vip">Ja</span>' : '<span class="bbz-muted">Nein</span>')}
              </div></div>
            </section>
            <section class="bbz-section" style="grid-column: span 2;">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Kontakte</div><div class="bbz-section-subtitle">Alle Ansprechpartner dieser Firma</div></div></div>
              <div class="bbz-section-body">
                <div class="bbz-table-wrap">
                  <table class="bbz-table">
                    <thead><tr><th>Name</th><th>Funktion</th><th>Rolle</th><th>E-Mail</th><th>Telefon</th><th>Archiviert</th></tr></thead>
                    <tbody>
                      ${firm.contacts.length ? firm.contacts.map(c => `
                        <tr>
                          <td><a class="bbz-link" data-action="open-contact" data-id="${c.id}">${helpers.escapeHtml(c.fullName || c.nachname)}</a></td>
                          <td>${helpers.escapeHtml(c.funktion) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${helpers.escapeHtml(c.rolle) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${c.email1 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(c.email1)}">${helpers.escapeHtml(c.email1)}</a>` : '<span class="bbz-muted">—</span>'}</td>
                          <td>${helpers.escapeHtml(helpers.joinNonEmpty([c.direktwahl, c.mobile], " / ")) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${c.archiviert ? '<span class="bbz-danger">Ja</span>' : '<span class="bbz-muted">Nein</span>'}</td>
                        </tr>`).join("") : `<tr><td colspan="6">${ui.emptyBlock("Keine Kontakte vorhanden.")}</td></tr>`}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          </div>
          <div class="bbz-grid bbz-grid-2 mt-4">
            <section class="bbz-section">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Aktivitaeten</div><div class="bbz-section-subtitle">Aggregierte History ueber alle Kontakte</div></div></div>
              <div class="bbz-section-body">
                ${recentHistory.length ? `<div class="bbz-timeline">${recentHistory.map(h => `
                  <div class="bbz-timeline-item">
                    <div class="bbz-timeline-date">${helpers.formatDateTime(h.datum) || "—"}<br><span class="bbz-muted">${helpers.escapeHtml(h.contactName || "")}</span></div>
                    <div><div class="bbz-timeline-title">${helpers.escapeHtml(h.typ || h.title || "Eintrag")} ${h.projektbezugBool ? '<span class="bbz-chip">Projektbezug</span>' : '<span class="bbz-chip">Allgemein</span>'}</div><div class="bbz-timeline-text">${helpers.escapeHtml(h.notizen || "—")}</div></div>
                  </div>`).join("")}</div>` : ui.emptyBlock("Keine History-Eintraege vorhanden.")}
              </div>
            </section>
            <section class="bbz-section">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Aufgaben</div><div class="bbz-section-subtitle">Alle Tasks der Firma</div></div></div>
              <div class="bbz-section-body">
                <div class="bbz-table-wrap">
                  <table class="bbz-table">
                    <thead><tr><th>Titel</th><th>Deadline</th><th>Status</th><th>Kontaktperson</th></tr></thead>
                    <tbody>
                      ${firm.tasks.length ? firm.tasks.map(t => `
                        <tr>
                          <td>${helpers.escapeHtml(t.title) || '<span class="bbz-muted">—</span>'}</td>
                          <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.formatDate(t.deadline) || '<span class="bbz-muted">—</span>'}</td>
                          <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.escapeHtml(t.status) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${t.contactId ? `<a class="bbz-link" data-action="open-contact" data-id="${t.contactId}">${helpers.escapeHtml(t.contactName || "Kontakt")}</a>` : helpers.escapeHtml(t.contactName || "—")}</td>
                        </tr>`).join("") : `<tr><td colspan="4">${ui.emptyBlock("Keine Aufgaben vorhanden.")}</td></tr>`}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          </div>
        </div>
      `;
    },

    contacts() {
      const filters = state.filters.contacts;
      const rows = state.enriched.contacts.filter(c => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch = !search || [c.fullName, c.firmTitle, c.funktion, c.rolle, c.email1, c.email2, c.direktwahl, c.mobile, c.kommentar, ...c.sgf, ...c.event].some(v => helpers.textIncludes(v, search));
        return searchMatch && (!filters.archiviertAusblenden || !c.archiviert);
      });

      return `
        <section class="bbz-section">
          <div class="bbz-section-header">
            <div><div class="bbz-section-title">Kontakte</div><div class="bbz-section-subtitle">Operative Ansprechpartner ueber alle Firmen</div></div>
            <button class="bbz-button bbz-button-primary" data-action="open-contact-form">+ Kontakt</button>
          </div>
          <div class="bbz-section-body">
            <div class="bbz-filters-3">
              <input class="bbz-input" data-filter="contacts-search" type="text" placeholder="Suche nach Name, Firma, Funktion, Rolle, E-Mail ..." value="${helpers.escapeHtml(filters.search)}" />
              <label class="bbz-checkbox"><input type="checkbox" data-filter="contacts-archiviert" ${filters.archiviertAusblenden ? "checked" : ""} /> Archivierte ausblenden</label>
              <div></div>
            </div>
            <div class="bbz-table-wrap">
              <table class="bbz-table">
                <thead><tr><th>Name</th><th>Firma</th><th>Funktion</th><th>Rolle</th><th>E-Mail</th><th>Telefon</th><th>Archiviert</th></tr></thead>
                <tbody>
                  ${rows.length ? rows.map(c => `
                    <tr>
                      <td><a class="bbz-link" data-action="open-contact" data-id="${c.id}">${helpers.escapeHtml(c.fullName || c.nachname)}</a></td>
                      <td>${c.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${c.firmId}">${helpers.escapeHtml(c.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
                      <td>${helpers.escapeHtml(c.funktion) || '<span class="bbz-muted">—</span>'}</td>
                      <td>${helpers.escapeHtml(c.rolle) || '<span class="bbz-muted">—</span>'}</td>
                      <td>${c.email1 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(c.email1)}">${helpers.escapeHtml(c.email1)}</a>` : '<span class="bbz-muted">—</span>'}</td>
                      <td>${helpers.escapeHtml(helpers.joinNonEmpty([c.direktwahl, c.mobile], " / ")) || '<span class="bbz-muted">—</span>'}</td>
                      <td>${c.archiviert ? '<span class="bbz-danger">Ja</span>' : '<span class="bbz-muted">Nein</span>'}</td>
                    </tr>`).join("") : `<tr><td colspan="7">${ui.emptyBlock("Keine Kontakte fuer die aktuelle Filterung gefunden.")}</td></tr>`}
                </tbody>
              </table>
            </div>
          </div>
        </section>
      `;
    },

    contactDetail() {
      const contact = dataModel.getContactById(state.selection.contactId);
      if (!contact) return ui.emptyBlock("Der ausgewaehlte Kontakt wurde nicht gefunden.");
      const contactHistory = state.enriched.history.filter(h => h.contactId === contact.id).sort((a, b) => helpers.compareDateDesc(a.datum, b.datum));
      const contactTasks = state.enriched.tasks.filter(t => t.contactId === contact.id).sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline));
      const isPrivat = state.meta.privateFirmId !== null && contact.firmId === state.meta.privateFirmId;

      return `
        <div>
          <div class="bbz-detail-header">
            <div>
              <button class="bbz-button bbz-button-secondary mb-3" data-action="back-to-contacts">Zurueck zur Kontaktliste</button>
              <div class="bbz-detail-title">${helpers.escapeHtml(contact.fullName || contact.nachname)}</div>
              <div class="bbz-detail-subtitle">
                ${isPrivat
                  ? `<span class="bbz-pill" style="font-size:12px;">Privatperson</span>`
                  : contact.firmId
                    ? `<a class="bbz-link" data-action="open-firm" data-id="${contact.firmId}">${helpers.escapeHtml(contact.firmTitle || "Firma")}</a>`
                    : "Keine Firma verknuepft"
                }
                ${contact.funktion ? ` · ${helpers.escapeHtml(contact.funktion)}` : ""}
                ${contact.rolle ? ` · ${helpers.escapeHtml(contact.rolle)}` : ""}
              </div>
            </div>
            <div class="flex items-center gap-2 flex-wrap">
              ${contact.email1 ? `<a class="bbz-button bbz-button-secondary" href="mailto:${helpers.escapeHtml(contact.email1)}">Mail senden</a>` : ""}
              <button class="bbz-button bbz-button-secondary" style="color:var(--red);border-color:var(--red);" data-action="delete-contact" data-id="${contact.id}" data-name="${helpers.escapeHtml(contact.fullName || contact.nachname)}">Löschen</button>
              <button class="bbz-button bbz-button-primary" data-action="open-contact-form" data-item-id="${contact.id}">Bearbeiten</button>
            </div>
          </div>
          <div class="bbz-kpis">
            ${this.kpiBlock("Tasks", contactTasks.length)}
            ${this.kpiBlock("Offene Tasks", contactTasks.filter(t => t.isOpen).length)}
            ${this.kpiBlock("History", contactHistory.length)}
            ${this.kpiBlock("Letzte Aktivitaet", helpers.formatDate(contactHistory[0]?.datum) || "—")}
          </div>
          <div class="bbz-grid bbz-grid-3">
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Stammdaten</div></div>
              <div class="bbz-section-body"><div class="bbz-meta-grid">
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
                ${ui.kv("Archiviert", contact.archiviert ? '<span class="bbz-danger">Ja</span>' : '<span class="bbz-muted">Nein</span>')}
              </div></div>
            </section>
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">CRM-Kontext</div></div>
              <div class="bbz-section-body"><div class="bbz-meta-grid">
                ${ui.kv("Leadbbz0", helpers.escapeHtml(contact.leadbbz0) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("SGF", helpers.multiChoiceHtml(contact.sgf))}
                ${ui.kv("Event", helpers.multiChoiceHtml(contact.event))}
                ${ui.kv("Eventhistory", helpers.multiChoiceHtml(contact.eventhistory))}
                ${isPrivat ? "" : ui.kv("Kommentar", helpers.escapeHtml(contact.kommentar) || '<span class="bbz-muted">—</span>')}
              </div></div>
            </section>
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Uebersicht</div></div>
              <div class="bbz-section-body"><div class="bbz-meta-grid">
                ${ui.kv("Tasks", String(contactTasks.length))}
                ${ui.kv("Offene Tasks", String(contactTasks.filter(t => t.isOpen).length))}
                ${ui.kv("History", String(contactHistory.length))}
                ${ui.kv("Letzte Aktivitaet", helpers.formatDateTime(contactHistory[0]?.datum) || '<span class="bbz-muted">—</span>')}
              </div></div>
            </section>
          </div>
          <div class="bbz-grid bbz-grid-2 mt-4">
            <section class="bbz-section">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Historie</div><div class="bbz-section-subtitle">Timeline aus CRMHistory</div></div></div>
              <div class="bbz-section-body">
                ${contactHistory.length ? `<div class="bbz-timeline">${contactHistory.map(h => `
                  <div class="bbz-timeline-item">
                    <div class="bbz-timeline-date">${helpers.formatDateTime(h.datum) || "—"}</div>
                    <div><div class="bbz-timeline-title">${helpers.escapeHtml(h.typ || h.title || "Eintrag")} ${h.projektbezugBool ? '<span class="bbz-chip">Projektbezug</span>' : '<span class="bbz-chip">Allgemein</span>'}</div><div class="bbz-timeline-text">${helpers.escapeHtml(h.notizen || "—")}</div></div>
                  </div>`).join("")}</div>` : ui.emptyBlock("Keine Historie vorhanden.")}
              </div>
            </section>
            <section class="bbz-section">
              <div class="bbz-section-header"><div><div class="bbz-section-title">Tasks</div><div class="bbz-section-subtitle">Aufgaben dieser Person</div></div></div>
              <div class="bbz-section-body">
                <div class="bbz-table-wrap">
                  <table class="bbz-table">
                    <thead><tr><th>Titel</th><th>Deadline</th><th>Status</th><th>Firma</th></tr></thead>
                    <tbody>
                      ${contactTasks.length ? contactTasks.map(t => `
                        <tr>
                          <td>${helpers.escapeHtml(t.title) || '<span class="bbz-muted">—</span>'}</td>
                          <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.formatDate(t.deadline) || '<span class="bbz-muted">—</span>'}</td>
                          <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.escapeHtml(t.status) || '<span class="bbz-muted">—</span>'}</td>
                          <td>${t.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${t.firmId}">${helpers.escapeHtml(t.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
                        </tr>`).join("") : `<tr><td colspan="4">${ui.emptyBlock("Keine Tasks vorhanden.")}</td></tr>`}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          </div>
        </div>
      `;
    },

    planning() {
      const filters = state.filters.planning;
      const rows = state.enriched.tasks.filter(t => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch = !search || [t.title, t.status, t.contactName, t.firmTitle, t.leadbbz].some(v => helpers.textIncludes(v, search));
        return searchMatch && (!filters.onlyOpen || t.isOpen) && (!filters.onlyOverdue || t.isOverdue);
      });

      const openTasks = state.enriched.tasks.filter(t => t.isOpen).length;
      const overdueTasks = state.enriched.tasks.filter(t => t.isOpen && t.isOverdue).length;
      const nextWeekTasks = state.enriched.tasks.filter(t => {
        if (!t.isOpen) return false;
        const d = helpers.toDate(t.deadline);
        if (!d) return false;
        const today = helpers.todayStart();
        const in7 = new Date(today);
        in7.setDate(in7.getDate() + 7);
        return d >= today && d <= in7;
      }).length;

      return `
        <div>
          <div class="bbz-kpis">
            ${this.kpiBlock("Tasks gesamt", state.enriched.tasks.length)}
            ${this.kpiBlock("Offen", openTasks)}
            ${this.kpiBlock("Ueberfaellig", overdueTasks)}
            ${this.kpiBlock("Naechste 7 Tage", nextWeekTasks)}
          </div>
          <section class="bbz-section">
            <div class="bbz-section-header"><div><div class="bbz-section-title">Planung</div><div class="bbz-section-subtitle">Aufgabenuebersicht mit Fokus auf offen und ueberfaellig</div></div></div>
            <div class="bbz-section-body">
              <div class="bbz-filters-3">
                <input class="bbz-input" data-filter="planning-search" type="text" placeholder="Suche nach Titel, Firma, Kontakt, Status ..." value="${helpers.escapeHtml(filters.search)}" />
                <label class="bbz-checkbox"><input type="checkbox" data-filter="planning-open" ${filters.onlyOpen ? "checked" : ""} /> Nur offene Tasks</label>
                <label class="bbz-checkbox"><input type="checkbox" data-filter="planning-overdue" ${filters.onlyOverdue ? "checked" : ""} /> Nur ueberfaellige Tasks</label>
              </div>
              <div class="bbz-table-wrap">
                <table class="bbz-table">
                  <thead><tr><th>Titel</th><th>Deadline</th><th>Status</th><th>Kontaktperson</th><th>Firma</th></tr></thead>
                  <tbody>
                    ${rows.length ? rows.map(t => `
                      <tr>
                        <td>${helpers.escapeHtml(t.title) || '<span class="bbz-muted">—</span>'}</td>
                        <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.formatDate(t.deadline) || '<span class="bbz-muted">—</span>'}</td>
                        <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.escapeHtml(t.status) || '<span class="bbz-muted">—</span>'}</td>
                        <td>${t.contactId ? `<a class="bbz-link" data-action="open-contact" data-id="${t.contactId}">${helpers.escapeHtml(t.contactName || "Kontakt")}</a>` : helpers.escapeHtml(t.contactName || "—")}</td>
                        <td>${t.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${t.firmId}">${helpers.escapeHtml(t.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
                      </tr>`).join("") : `<tr><td colspan="5">${ui.emptyBlock("Keine Tasks fuer die aktuelle Filterung gefunden.")}</td></tr>`}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        </div>
      `;
    },

    events() {
      const filters = state.filters.events;
      const groups = state.enriched.events.map(group => ({
        ...group,
        contacts: group.contacts.filter(item => {
          const search = filters.search.trim().toLowerCase();
          const searchMatch = !search || [group.name, item.contactName, item.firmTitle, item.rolle, item.funktion, item.eventhistory, item.latestHistoryText].some(v => helpers.textIncludes(v, search));
          return searchMatch && (!filters.onlyWithOpenTasks || item.openTasksCount > 0);
        })
      })).filter(g => g.contacts.length > 0);

      const totalGroups = state.enriched.events.length;
      const totalContacts = state.enriched.events.reduce((sum, e) => sum + e.contactCount, 0);
      const totalOpenTasks = state.enriched.events.reduce((sum, e) => sum + e.openTasksCount, 0);

      return `
        <div>
          <div class="bbz-kpis">
            ${this.kpiBlock("Event-Kategorien", totalGroups)}
            ${this.kpiBlock("Kontakt-Zuordnungen", totalContacts)}
            ${this.kpiBlock("Offene Tasks", totalOpenTasks)}
            ${this.kpiBlock("Sichtbare Kategorien", groups.length)}
          </div>
          <section class="bbz-section">
            <div class="bbz-section-header"><div><div class="bbz-section-title">Events</div><div class="bbz-section-subtitle">Separate Event-Sicht nach Kategorie</div></div></div>
            <div class="bbz-section-body">
              <div class="bbz-filters-3">
                <input class="bbz-input" data-filter="events-search" type="text" placeholder="Suche nach Kategorie, Kontakt, Firma, Rolle ..." value="${helpers.escapeHtml(filters.search)}" />
                <label class="bbz-checkbox"><input type="checkbox" data-filter="events-open" ${filters.onlyWithOpenTasks ? "checked" : ""} /> Nur mit offenen Tasks</label>
                <div></div>
              </div>
              ${groups.length ? `<div class="bbz-cockpit-stack">${groups.map(group => `
                <section class="bbz-section" style="box-shadow:none;">
                  <div class="bbz-section-header"><div><div class="bbz-section-title">${helpers.escapeHtml(group.name)}</div><div class="bbz-section-subtitle">${group.contacts.length} Kontakte · ${group.contacts.reduce((sum, c) => sum + c.openTasksCount, 0)} offene Tasks</div></div></div>
                  <div class="bbz-section-body">
                    <div class="bbz-table-wrap">
                      <table class="bbz-table">
                        <thead><tr><th>Kontakt</th><th>Firma</th><th>Funktion / Rolle</th><th>Eventhistory</th><th>Letzte Aktivitaet</th><th>Offene Tasks</th></tr></thead>
                        <tbody>
                          ${group.contacts.map(item => `
                            <tr>
                              <td><a class="bbz-link" data-action="open-contact" data-id="${item.contactId}">${helpers.escapeHtml(item.contactName)}</a><div class="bbz-subtext">${item.email1 ? helpers.escapeHtml(item.email1) : "—"}</div></td>
                              <td>${item.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${item.firmId}">${helpers.escapeHtml(item.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
                              <td>${helpers.escapeHtml(helpers.joinNonEmpty([item.funktion, item.rolle], " · ")) || '<span class="bbz-muted">—</span>'}</td>
                              <td>${helpers.escapeHtml(item.eventhistory) || '<span class="bbz-muted">—</span>'}</td>
                              <td>${helpers.formatDateTime(item.latestHistoryDate) || '<span class="bbz-muted">—</span>'}${item.latestHistoryType ? `<div class="bbz-subtext">${helpers.escapeHtml(item.latestHistoryType)}</div>` : ""}</td>
                              <td class="${item.openTasksCount > 0 ? "bbz-warning" : "bbz-muted"}">${item.openTasksCount}</td>
                            </tr>`).join("")}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </section>`).join("")}</div>` : ui.emptyBlock("Keine Event-Daten fuer die aktuelle Filterung gefunden.")}
            </div>
          </section>
        </div>
      `;
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
          await api.acquireToken();
          // Choices und Daten beim ersten Laden parallel — Choices nur einmal nötig
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
        // Choices und Daten beim Login parallel laden
        await Promise.all([api.loadAll(), api.loadColumnChoices()]);
        ui.setMessage("Anmeldung erfolgreich. Daten wurden geladen.", "success");
      } catch (error) {
        console.error(error);
        ui.setMessage(`Anmeldung fehlgeschlagen: ${error.message}`, "error");
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
        VIP:   form.querySelector("[name='vip']")?.checked ?? false
      };

      if (fd.get("adresse")?.trim())      fields.Adresse      = fd.get("adresse").trim();
      if (fd.get("plz")?.trim())          fields.PLZ          = fd.get("plz").trim();
      if (fd.get("ort")?.trim())          fields.Ort          = fd.get("ort").trim();
      if (fd.get("land")?.trim())         fields.Land         = fd.get("land").trim();
      if (fd.get("hauptnummer")?.trim())  fields.Hauptnummer  = fd.get("hauptnummer").trim();
      if (fd.get("klassifizierung"))      fields.Klassifizierung = fd.get("klassifizierung");

      console.log("handleFirmModalSubmit fields →", JSON.stringify(fields, null, 2));
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

      // Einzelwahl — nur wenn Wert vorhanden
      if (raw.anrede)   fields.Anrede   = raw.anrede;
      if (raw.rolle)    fields.Rolle    = raw.rolle;
      if (raw.leadbbz0) fields.Leadbbz0 = raw.leadbbz0;

      // Optionaler Text — nur wenn befüllt
      if (raw.vorname.trim())    fields.Vorname    = raw.vorname.trim();
      if (raw.funktion.trim())   fields.Funktion   = raw.funktion.trim();
      if (raw.kommentar.trim())  fields.Kommentar  = raw.kommentar.trim();
      if (raw.email1.trim())     fields.Email1     = raw.email1.trim();
      if (raw.email2.trim())     fields.Email2     = raw.email2.trim();
      if (raw.direktwahl.trim()) fields.Direktwahl = raw.direktwahl.trim();
      if (raw.mobile.trim())     fields.Mobile     = raw.mobile.trim();

      // Datum — nur wenn befüllt, SP erwartet volles ISO-8601 Datetime (nicht nur YYYY-MM-DD)
      if (raw.geburtstag.trim()) fields.Geburtstag = raw.geburtstag.trim() + "T00:00:00Z";

      // Multi-Choice — @odata.type + Array (befüllen) oder @odata.type + [] (leeren)
      // BESTÄTIGT: @odata.type + Array mit Werten → ✅
      // OFFEN: @odata.type + [] zum Leeren → zu testen
      fields["SGF@odata.type"]          = "Collection(Edm.String)";
      fields["SGF"]                     = raw.sgf;
      fields["Event@odata.type"]        = "Collection(Edm.String)";
      fields["Event"]                   = raw.event;
      fields["Eventhistory@odata.type"] = "Collection(Edm.String)";
      fields["Eventhistory"]            = raw.eventhistory;

      // Debug-Log (kann nach stabilem Betrieb entfernt werden)
      console.log("handleModalSubmit fields →", JSON.stringify(fields, null, 2));

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

    navigate(route) {
      state.filters.route = route;
      if (route !== "firms") state.selection.firmId = null;
      if (route !== "contacts") state.selection.contactId = null;
      state.modal = null;
      window.scrollTo(0, 0);
      this.render();
    },

    openFirm(id) {
      state.selection.firmId = id;
      state.selection.contactId = null;
      state.filters.route = "firms";
      state.modal = null;
      window.scrollTo(0, 0);
      this.render();
    },

    openContact(id) {
      state.selection.contactId = id;
      state.filters.route = "contacts";
      state.modal = null;
      window.scrollTo(0, 0);
      this.render();
    },

    render() {
      ui.renderShell();
      ui.renderView(views.renderRoute());
    }
  };

  function startApp() { controller.init(); }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startApp, { once: true });
  } else {
    startApp();
  }
})();

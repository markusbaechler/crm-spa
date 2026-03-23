(() => {
  "use strict";

  const CONFIG = {
    appName: "bbz CRM",

    graph: {
      tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
      clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
      authority: "https://login.microsoftonline.com/3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
      redirectUri: "https://markusbaechler.github.io/crm-spa/",
      scopes: ["User.Read", "Sites.Read.All"]
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
      planningShowOnlyOpen: true
    }
  };

  const SCHEMA = {
    firms: {
      listTitle: CONFIG.lists.firms,
      fields: {
        id: "id",
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
        id: "id",
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
        id: "id",
        title: "Title",
        kontakt: "Nachname",
        kontaktLookupId: "NachnameLookupId",
        datum: "Datum",
        typ: "Typ",
        notizen: "Notizen",
        projektbezug: "Projektbezug",
        leadbbz: "Leadbbz"
      }
    },

    tasks: {
      listTitle: CONFIG.lists.tasks,
      fields: {
        id: "id",
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
      lastError: null
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
      tasks: []
    },

    filters: {
      route: CONFIG.defaults.route,

      firms: {
        search: "",
        klassifizierung: "",
        vip: ""
      },

      contacts: {
        search: "",
        archiviertAusblenden: CONFIG.defaults.contactArchiveDefaultHidden
      },

      planning: {
        search: "",
        onlyOpen: CONFIG.defaults.planningShowOnlyOpen,
        onlyOverdue: false
      }
    },

    selection: {
      firmId: null,
      contactId: null
    }
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
        if (value.includes(";#")) {
          return value.split(";#").map(v => v.trim()).filter(Boolean);
        }
        if (value.includes(",")) {
          return value.split(",").map(v => v.trim()).filter(Boolean);
        }
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
      return d.toLocaleDateString("de-CH", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric"
      });
    },

    formatDateTime(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      return d.toLocaleString("de-CH", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit"
      });
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
      const ad = helpers.toDate(a);
      const bd = helpers.toDate(b);
      if (!ad && !bd) return 0;
      if (!ad) return 1;
      if (!bd) return -1;
      return ad - bd;
    },

    compareDateDesc(a, b) {
      const ad = helpers.toDate(a);
      const bd = helpers.toDate(b);
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
      if (v === "A") return "bbz-badge bbz-badge-a";
      if (v === "B") return "bbz-badge bbz-badge-b";
      if (v === "C") return "bbz-badge bbz-badge-c";
      return "bbz-badge";
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

    ensureMsalAvailable() {
      if (!window.msal || !window.msal.PublicClientApplication) {
        throw new Error("MSAL-Bibliothek wurde nicht geladen. Prüfe index.html und die Script-Referenz.");
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

      if (missing.length) {
        throw new Error(`Konfiguration unvollständig: ${missing.join(", ")}`);
      }
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

      this.els.btnLogin.addEventListener("click", () => controller.handleLogin());
      this.els.btnRefresh.addEventListener("click", () => controller.handleRefresh());

      this.els.navButtons.forEach(btn => {
        btn.addEventListener("click", () => {
          controller.navigate(btn.dataset.route);
        });
      });

      document.addEventListener("click", (event) => {
        const openFirm = event.target.closest("[data-action='open-firm']");
        if (openFirm) {
          controller.openFirm(openFirm.dataset.id);
          return;
        }

        const openContact = event.target.closest("[data-action='open-contact']");
        if (openContact) {
          controller.openContact(openContact.dataset.id);
          return;
        }

        const backToFirms = event.target.closest("[data-action='back-to-firms']");
        if (backToFirms) {
          controller.navigate("firms");
          return;
        }

        const backToContacts = event.target.closest("[data-action='back-to-contacts']");
        if (backToContacts) {
          controller.navigate("contacts");
          return;
        }
      });

      document.addEventListener("input", (event) => {
        const el = event.target;

        if (el.matches("[data-filter='firms-search']")) {
          state.filters.firms.search = el.value;
          controller.render();
        }
        if (el.matches("[data-filter='contacts-search']")) {
          state.filters.contacts.search = el.value;
          controller.render();
        }
        if (el.matches("[data-filter='planning-search']")) {
          state.filters.planning.search = el.value;
          controller.render();
        }
      });

      document.addEventListener("change", (event) => {
        const el = event.target;

        if (el.matches("[data-filter='firms-klassifizierung']")) {
          state.filters.firms.klassifizierung = el.value;
          controller.render();
        }
        if (el.matches("[data-filter='firms-vip']")) {
          state.filters.firms.vip = el.value;
          controller.render();
        }
        if (el.matches("[data-filter='contacts-archiviert']")) {
          state.filters.contacts.archiviertAusblenden = el.checked;
          controller.render();
        }
        if (el.matches("[data-filter='planning-open']")) {
          state.filters.planning.onlyOpen = el.checked;
          controller.render();
        }
        if (el.matches("[data-filter='planning-overdue']")) {
          state.filters.planning.onlyOverdue = el.checked;
          controller.render();
        }
      });
    },

    setLoading(isLoading) {
      state.meta.loading = isLoading;
      this.renderShell();
    },

    setMessage(message, type = "info") {
      const el = this.els.globalMessage;
      if (!message) {
        el.className = "hidden mb-4";
        el.innerHTML = "";
        return;
      }

      let colorClasses = "bg-slate-50 border-slate-200 text-slate-700";
      if (type === "error") colorClasses = "bg-red-50 border-red-200 text-red-700";
      if (type === "success") colorClasses = "bg-green-50 border-green-200 text-green-700";
      if (type === "warning") colorClasses = "bg-amber-50 border-amber-200 text-amber-700";

      el.className = `mb-4 border rounded-xl px-4 py-3 ${colorClasses}`;
      el.textContent = message;
    },

    renderShell() {
      this.els.navButtons.forEach(btn => {
        btn.classList.toggle("active", btn.dataset.route === state.filters.route);
      });

      if (state.auth.isAuthenticated && state.auth.account) {
        this.els.authStatus.textContent = `Angemeldet: ${state.auth.account.username || state.auth.account.name || ""}`;
      } else if (state.auth.isReady) {
        this.els.authStatus.textContent = "Nicht angemeldet";
      } else {
        this.els.authStatus.textContent = "Authentifizierung wird initialisiert ...";
      }

      this.els.btnLogin.textContent = state.auth.isAuthenticated ? "Erneut anmelden" : "Anmelden";
      this.els.btnLogin.disabled = state.meta.loading || !state.auth.isReady;
      this.els.btnRefresh.disabled = state.meta.loading || !state.auth.isReady;
    },

    renderView(html) {
      this.els.viewRoot.innerHTML = html;
    },

    loadingBlock(text = "Daten werden geladen ...") {
      return `
        <div class="bbz-card">
          <div class="bbz-card-body flex items-center gap-3">
            <div class="bbz-loader"></div>
            <div class="text-sm text-slate-600">${helpers.escapeHtml(text)}</div>
          </div>
        </div>
      `;
    },

    emptyBlock(text = "Keine Daten vorhanden.") {
      return `<div class="bbz-empty">${helpers.escapeHtml(text)}</div>`;
    },

    kv(label, value) {
      return `
        <div>
          <div class="bbz-label">${helpers.escapeHtml(label)}</div>
          <div class="bbz-value">${value || '<span class="bbz-muted">—</span>'}</div>
        </div>
      `;
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
        cache: {
          cacheLocation: "localStorage"
        }
      });

      await msalInstance.initialize();

      state.auth.msal = msalInstance;

      try {
        const redirectResponse = await state.auth.msal.handleRedirectPromise();
        if (redirectResponse && redirectResponse.account) {
          state.auth.account = redirectResponse.account;
          state.auth.isAuthenticated = true;
        }
      } catch (error) {
        console.warn("handleRedirectPromise Fehler", error);
      }

      const accounts = state.auth.msal.getAllAccounts();
      if (accounts.length > 0 && !state.auth.account) {
        state.auth.account = accounts[0];
        state.auth.isAuthenticated = true;
      }

      state.auth.isReady = true;
    },

    async login() {
      if (!state.auth.msal) {
        throw new Error("MSAL ist nicht initialisiert.");
      }

      const loginResponse = await state.auth.msal.loginPopup({
        scopes: CONFIG.graph.scopes,
        prompt: "select_account"
      });

      if (!loginResponse || !loginResponse.account) {
        throw new Error("Keine Kontoinformation aus dem Login erhalten.");
      }

      state.auth.account = loginResponse.account;
      state.auth.isAuthenticated = true;

      await this.acquireToken();
    },

    async acquireToken() {
      if (!state.auth.msal) {
        throw new Error("MSAL ist nicht initialisiert.");
      }

      if (!state.auth.account) {
        throw new Error("Kein angemeldetes Konto gefunden.");
      }

      try {
        const tokenResponse = await state.auth.msal.acquireTokenSilent({
          account: state.auth.account,
          scopes: CONFIG.graph.scopes
        });

        if (!tokenResponse || !tokenResponse.accessToken) {
          throw new Error("Kein Access Token aus acquireTokenSilent erhalten.");
        }

        state.auth.token = tokenResponse.accessToken;
        return state.auth.token;
      } catch (silentError) {
        const tokenResponse = await state.auth.msal.acquireTokenPopup({
          account: state.auth.account,
          scopes: CONFIG.graph.scopes
        });

        if (!tokenResponse || !tokenResponse.accessToken) {
          throw new Error("Kein Access Token aus acquireTokenPopup erhalten.");
        }

        state.auth.token = tokenResponse.accessToken;
        return state.auth.token;
      }
    },

    async graphRequest(path, options = {}) {
      const token = state.auth.token || await this.acquireToken();

      const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
        method: options.method || "GET",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          ...(options.headers || {})
        },
        body: options.body ? JSON.stringify(options.body) : undefined
      });

      if (!response.ok) {
        let detail = "";
        try {
          detail = await response.text();
        } catch {
          detail = "";
        }
        throw new Error(`Graph ${response.status}: ${detail || response.statusText}`);
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
    }
  };

  const normalizer = {
    getField(item, fieldName) {
      return item?.fields?.[fieldName];
    },

    itemId(item) {
      return Number(item?.id) || null;
    },

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
        eventhistory: this.getField(item, f.eventhistory) || "",
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
        return {
          ...contact,
          fullName: helpers.fullName(contact),
          firmId: firm?.id || contact.firmaLookupId || null,
          firmTitle: firm?.title || contact.firmaRaw || "",
          firm: firm
        };
      });

      const history = state.data.history.map(entry => {
        const contact = contactById.get(entry.kontaktLookupId) || null;
        const firm = contact ? firmById.get(contact.firmaLookupId) || null : null;

        return {
          ...entry,
          contactId: contact?.id || entry.kontaktLookupId || null,
          contactName: contact ? helpers.fullName(contact) : (entry.kontaktRaw || ""),
          firmId: firm?.id || null,
          firmTitle: firm?.title || "",
          projektbezugBool: helpers.bool(entry.projektbezug)
        };
      });

      const tasks = state.data.tasks.map(task => {
        const contact = contactById.get(task.kontaktLookupId) || null;
        const firm = contact ? firmById.get(contact.firmaLookupId) || null : null;

        return {
          ...task,
          contactId: contact?.id || task.kontaktLookupId || null,
          contactName: contact ? helpers.fullName(contact) : (task.kontaktRaw || ""),
          firmId: firm?.id || null,
          firmTitle: firm?.title || "",
          isOpen: helpers.isOpenTask(task.status),
          isOverdue: helpers.isOverdue(task.deadline)
        };
      });

      const firms = state.data.firms.map(firm => {
        const firmContacts = contacts.filter(c => c.firmId === firm.id);
        const firmContactIds = new Set(firmContacts.map(c => c.id));
        const firmTasks = tasks.filter(t => firmContactIds.has(t.contactId));
        const firmHistory = history.filter(h => firmContactIds.has(h.contactId));

        const openTasks = firmTasks.filter(t => t.isOpen);
        const nextDeadlineTask = openTasks
          .filter(t => helpers.toDate(t.deadline))
          .sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline))[0] || null;

        return {
          ...firm,
          contactsCount: firmContacts.length,
          contacts: firmContacts.sort((a, b) => a.fullName.localeCompare(b.fullName, "de")),
          tasks: firmTasks.sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline)),
          history: firmHistory.sort((a, b) => helpers.compareDateDesc(a.datum, b.datum)),
          openTasksCount: openTasks.length,
          nextDeadline: nextDeadlineTask?.deadline || ""
        };
      });

      state.enriched.contacts = contacts.sort((a, b) => a.fullName.localeCompare(b.fullName, "de"));
      state.enriched.history = history.sort((a, b) => helpers.compareDateDesc(a.datum, b.datum));
      state.enriched.tasks = tasks.sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline));
      state.enriched.firms = firms.sort((a, b) => a.title.localeCompare(b.title, "de"));
    },

    getFirmById(id) {
      return state.enriched.firms.find(f => String(f.id) === String(id)) || null;
    },

    getContactById(id) {
      return state.enriched.contacts.find(c => String(c.id) === String(id)) || null;
    }
  };

  const views = {
    renderRoute() {
      if (state.meta.loading) {
        return ui.loadingBlock();
      }

      switch (state.filters.route) {
        case "firms":
          if (state.selection.firmId) return this.firmDetail();
          return this.firms();
        case "contacts":
          if (state.selection.contactId) return this.contactDetail();
          return this.contacts();
        case "planning":
          return this.planning();
        default:
          return this.firms();
      }
    },

    firms() {
      const filters = state.filters.firms;
      const rows = state.enriched.firms.filter(firm => {
        const search = filters.search.trim().toLowerCase();

        const searchMatch =
          !search ||
          [
            firm.title,
            firm.ort,
            firm.klassifizierung,
            firm.hauptnummer,
            firm.adresse,
            firm.land,
            ...firm.contacts.map(c => c.fullName)
          ].some(v => helpers.textIncludes(v, search));

        const klassifizierungMatch =
          !filters.klassifizierung || firm.klassifizierung === filters.klassifizierung;

        const vipMatch =
          !filters.vip ||
          (filters.vip === "yes" && firm.vip) ||
          (filters.vip === "no" && !firm.vip);

        return searchMatch && klassifizierungMatch && vipMatch;
      });

      const aCount = state.enriched.firms.filter(f => f.klassifizierung === "A").length;
      const bCount = state.enriched.firms.filter(f => f.klassifizierung === "B").length;
      const cCount = state.enriched.firms.filter(f => f.klassifizierung === "C").length;
      const overdueTasks = state.enriched.tasks.filter(t => t.isOpen && t.isOverdue).length;

      return `
        <div class="space-y-4">
          <section class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4">
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Firmen gesamt</div>
              <div class="bbz-kpi-value">${state.enriched.firms.length}</div>
            </div>
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">A / B / C</div>
              <div class="bbz-kpi-value">${aCount} / ${bCount} / ${cCount}</div>
            </div>
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Kontakte gesamt</div>
              <div class="bbz-kpi-value">${state.enriched.contacts.length}</div>
            </div>
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Überfällige Tasks</div>
              <div class="bbz-kpi-value">${overdueTasks}</div>
            </div>
          </section>

          <section class="bbz-card">
            <div class="bbz-card-header">Firmen</div>
            <div class="bbz-card-body space-y-4">
              <div class="grid grid-cols-1 lg:grid-cols-4 gap-3">
                <input
                  class="bbz-input lg:col-span-2"
                  data-filter="firms-search"
                  type="text"
                  placeholder="Suche nach Firma, Ort, Ansprechpartner, Telefon ..."
                  value="${helpers.escapeHtml(filters.search)}"
                />
                <select class="bbz-select" data-filter="firms-klassifizierung">
                  <option value="">Alle Klassifizierungen</option>
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

              <div class="bbz-scroll">
                <table class="bbz-table">
                  <thead>
                    <tr>
                      <th>Firma</th>
                      <th>Ort</th>
                      <th>Klassifizierung</th>
                      <th>VIP</th>
                      <th>Anzahl Kontakte</th>
                      <th>Offene Tasks</th>
                      <th>Nächste Deadline</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${
                      rows.length
                        ? rows.map(firm => `
                          <tr>
                            <td>
                              <a class="bbz-link" data-action="open-firm" data-id="${firm.id}">
                                ${helpers.escapeHtml(firm.title)}
                              </a>
                            </td>
                            <td>${helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ")) || '<span class="bbz-muted">—</span>'}</td>
                            <td><span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung || "—")}</span></td>
                            <td>${firm.vip ? '<span class="bbz-badge bbz-badge-vip">VIP</span>' : '<span class="bbz-muted">—</span>'}</td>
                            <td>${firm.contactsCount}</td>
                            <td>${firm.openTasksCount}</td>
                            <td>${helpers.formatDate(firm.nextDeadline) || '<span class="bbz-muted">—</span>'}</td>
                          </tr>
                        `).join("")
                        : `<tr><td colspan="7">${ui.emptyBlock("Keine Firmen für die aktuelle Filterung gefunden.")}</td></tr>`
                    }
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        </div>
      `;
    },

    firmDetail() {
      const firm = dataModel.getFirmById(state.selection.firmId);
      if (!firm) {
        return ui.emptyBlock("Die ausgewählte Firma wurde nicht gefunden.");
      }

      const recentHistory = [...firm.history].slice(0, 20);
      const firmTasks = [...firm.tasks];
      const contacts = [...firm.contacts];

      return `
        <div class="space-y-4">
          <section class="flex items-center justify-between gap-3 flex-wrap">
            <div>
              <button class="bbz-button bbz-button-secondary mb-3" data-action="back-to-firms">Zurück zur Firmenliste</button>
              <div class="flex items-center gap-2 flex-wrap">
                <div class="text-2xl text-slate-900 font-semibold">${helpers.escapeHtml(firm.title)}</div>
                ${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : ""}
                ${firm.vip ? `<span class="bbz-badge bbz-badge-vip">VIP</span>` : ""}
              </div>
              <div class="text-sm text-slate-500 mt-1">
                ${helpers.escapeHtml(helpers.joinNonEmpty([firm.adresse, helpers.joinNonEmpty([firm.plz, firm.ort], " "), firm.land], " · ")) || "Keine erweiterten Stammdaten"}
              </div>
            </div>
          </section>

          <section class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4">
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Kontakte</div>
              <div class="bbz-kpi-value">${firm.contactsCount}</div>
            </div>
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Offene Tasks</div>
              <div class="bbz-kpi-value">${firm.openTasksCount}</div>
            </div>
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Nächste Deadline</div>
              <div class="bbz-kpi-value text-[20px]">${helpers.formatDate(firm.nextDeadline) || "—"}</div>
            </div>
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">History-Einträge</div>
              <div class="bbz-kpi-value">${firm.history.length}</div>
            </div>
          </section>

          <section class="grid grid-cols-1 xl:grid-cols-3 gap-4">
            <div class="bbz-card xl:col-span-1">
              <div class="bbz-card-header">Übersicht</div>
              <div class="bbz-card-body grid grid-cols-1 gap-4">
                ${ui.kv("Firma", helpers.escapeHtml(firm.title))}
                ${ui.kv("Adresse", helpers.escapeHtml(firm.adresse) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("PLZ / Ort", helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ")) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Land", helpers.escapeHtml(firm.land) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Hauptnummer", helpers.escapeHtml(firm.hauptnummer) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Klassifizierung", firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : '<span class="bbz-muted">—</span>')}
                ${ui.kv("VIP", firm.vip ? '<span class="bbz-badge bbz-badge-vip">Ja</span>' : '<span class="bbz-muted">Nein</span>')}
              </div>
            </div>

            <div class="bbz-card xl:col-span-2">
              <div class="bbz-card-header">Kontakte</div>
              <div class="bbz-card-body">
                <div class="bbz-scroll">
                  <table class="bbz-table">
                    <thead>
                      <tr>
                        <th>Name</th>
                        <th>Funktion</th>
                        <th>Rolle</th>
                        <th>E-Mail</th>
                        <th>Telefon</th>
                        <th>Archiviert</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${
                        contacts.length
                          ? contacts.map(c => `
                            <tr>
                              <td>
                                <a class="bbz-link" data-action="open-contact" data-id="${c.id}">
                                  ${helpers.escapeHtml(c.fullName || c.nachname)}
                                </a>
                              </td>
                              <td>${helpers.escapeHtml(c.funktion) || '<span class="bbz-muted">—</span>'}</td>
                              <td>${helpers.escapeHtml(c.rolle) || '<span class="bbz-muted">—</span>'}</td>
                              <td>${c.email1 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(c.email1)}">${helpers.escapeHtml(c.email1)}</a>` : '<span class="bbz-muted">—</span>'}</td>
                              <td>${helpers.escapeHtml(helpers.joinNonEmpty([c.direktwahl, c.mobile], " / ")) || '<span class="bbz-muted">—</span>'}</td>
                              <td>${c.archiviert ? '<span class="bbz-danger">Ja</span>' : '<span class="bbz-muted">Nein</span>'}</td>
                            </tr>
                          `).join("")
                          : `<tr><td colspan="6">${ui.emptyBlock("Keine Kontakte vorhanden.")}</td></tr>`
                      }
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </section>

          <section class="grid grid-cols-1 xl:grid-cols-2 gap-4">
            <div class="bbz-card">
              <div class="bbz-card-header">Aktivitäten</div>
              <div class="bbz-card-body">
                ${
                  recentHistory.length
                    ? `
                      <div class="bbz-timeline">
                        ${recentHistory.map(h => `
                          <div class="bbz-timeline-item">
                            <div class="bbz-timeline-date">
                              ${helpers.formatDateTime(h.datum) || "—"}<br>
                              <span class="bbz-muted">${helpers.escapeHtml(h.contactName || "")}</span>
                            </div>
                            <div class="bbz-timeline-body">
                              <div class="bbz-timeline-title">
                                ${helpers.escapeHtml(h.typ || h.title || "Eintrag")}
                                ${h.projektbezugBool ? '<span class="bbz-chip">Projektbezug</span>' : '<span class="bbz-chip">Allgemein</span>'}
                              </div>
                              <div class="bbz-timeline-text">${helpers.escapeHtml(h.notizen || "—")}</div>
                            </div>
                          </div>
                        `).join("")}
                      </div>
                    `
                    : ui.emptyBlock("Keine History-Einträge vorhanden.")
                }
              </div>
            </div>

            <div class="bbz-card">
              <div class="bbz-card-header">Aufgaben</div>
              <div class="bbz-card-body">
                <div class="bbz-scroll">
                  <table class="bbz-table">
                    <thead>
                      <tr>
                        <th>Titel</th>
                        <th>Deadline</th>
                        <th>Status</th>
                        <th>Kontaktperson</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${
                        firmTasks.length
                          ? firmTasks.map(t => `
                            <tr>
                              <td>${helpers.escapeHtml(t.title) || '<span class="bbz-muted">—</span>'}</td>
                              <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.formatDate(t.deadline) || '<span class="bbz-muted">—</span>'}</td>
                              <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.escapeHtml(t.status) || '<span class="bbz-muted">—</span>'}</td>
                              <td>
                                ${
                                  t.contactId
                                    ? `<a class="bbz-link" data-action="open-contact" data-id="${t.contactId}">${helpers.escapeHtml(t.contactName || "Kontakt")}</a>`
                                    : helpers.escapeHtml(t.contactName || "—")
                                }
                              </td>
                            </tr>
                          `).join("")
                          : `<tr><td colspan="4">${ui.emptyBlock("Keine Aufgaben vorhanden.")}</td></tr>`
                      }
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </section>
        </div>
      `;
    },

    contacts() {
      const filters = state.filters.contacts;
      const rows = state.enriched.contacts.filter(contact => {
        const search = filters.search.trim().toLowerCase();

        const searchMatch =
          !search ||
          [
            contact.fullName,
            contact.firmTitle,
            contact.funktion,
            contact.rolle,
            contact.email1,
            contact.email2,
            contact.direktwahl,
            contact.mobile,
            contact.kommentar,
            ...contact.sgf,
            ...contact.event
          ].some(v => helpers.textIncludes(v, search));

        const archiveMatch = !filters.archiviertAusblenden || !contact.archiviert;

        return searchMatch && archiveMatch;
      });

      return `
        <div class="space-y-4">
          <section class="bbz-card">
            <div class="bbz-card-header">Kontakte</div>
            <div class="bbz-card-body space-y-4">
              <div class="grid grid-cols-1 lg:grid-cols-3 gap-3">
                <input
                  class="bbz-input lg:col-span-2"
                  data-filter="contacts-search"
                  type="text"
                  placeholder="Suche nach Name, Firma, Funktion, Rolle, E-Mail, SGF, Event ..."
                  value="${helpers.escapeHtml(filters.search)}"
                />
                <label class="flex items-center gap-2 text-sm text-slate-700 h-[38px]">
                  <input
                    type="checkbox"
                    data-filter="contacts-archiviert"
                    ${filters.archiviertAusblenden ? "checked" : ""}
                  />
                  Archivierte Kontakte ausblenden
                </label>
              </div>

              <div class="bbz-scroll">
                <table class="bbz-table">
                  <thead>
                    <tr>
                      <th>Name</th>
                      <th>Firma</th>
                      <th>Funktion</th>
                      <th>Rolle</th>
                      <th>E-Mail</th>
                      <th>Telefon</th>
                      <th>Archiviert</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${
                      rows.length
                        ? rows.map(c => `
                          <tr>
                            <td>
                              <a class="bbz-link" data-action="open-contact" data-id="${c.id}">
                                ${helpers.escapeHtml(c.fullName || c.nachname)}
                              </a>
                            </td>
                            <td>
                              ${
                                c.firmId
                                  ? `<a class="bbz-link" data-action="open-firm" data-id="${c.firmId}">${helpers.escapeHtml(c.firmTitle || "Firma")}</a>`
                                  : `<span class="bbz-muted">—</span>`
                              }
                            </td>
                            <td>${helpers.escapeHtml(c.funktion) || '<span class="bbz-muted">—</span>'}</td>
                            <td>${helpers.escapeHtml(c.rolle) || '<span class="bbz-muted">—</span>'}</td>
                            <td>${c.email1 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(c.email1)}">${helpers.escapeHtml(c.email1)}</a>` : '<span class="bbz-muted">—</span>'}</td>
                            <td>${helpers.escapeHtml(helpers.joinNonEmpty([c.direktwahl, c.mobile], " / ")) || '<span class="bbz-muted">—</span>'}</td>
                            <td>${c.archiviert ? '<span class="bbz-danger">Ja</span>' : '<span class="bbz-muted">Nein</span>'}</td>
                          </tr>
                        `).join("")
                        : `<tr><td colspan="7">${ui.emptyBlock("Keine Kontakte für die aktuelle Filterung gefunden.")}</td></tr>`
                    }
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        </div>
      `;
    },

    contactDetail() {
      const contact = dataModel.getContactById(state.selection.contactId);
      if (!contact) {
        return ui.emptyBlock("Der ausgewählte Kontakt wurde nicht gefunden.");
      }

      const contactHistory = state.enriched.history
        .filter(h => h.contactId === contact.id)
        .sort((a, b) => helpers.compareDateDesc(a.datum, b.datum));

      const contactTasks = state.enriched.tasks
        .filter(t => t.contactId === contact.id)
        .sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline));

      return `
        <div class="space-y-4">
          <section class="flex items-center justify-between gap-3 flex-wrap">
            <div>
              <button class="bbz-button bbz-button-secondary mb-3" data-action="back-to-contacts">Zurück zur Kontaktliste</button>
              <div class="text-2xl text-slate-900 font-semibold">${helpers.escapeHtml(contact.fullName || contact.nachname)}</div>
              <div class="text-sm text-slate-500 mt-1">
                ${
                  contact.firmId
                    ? `<a class="bbz-link" data-action="open-firm" data-id="${contact.firmId}">${helpers.escapeHtml(contact.firmTitle || "Firma")}</a>`
                    : "Keine Firma verknüpft"
                }
                ${contact.funktion ? ` · ${helpers.escapeHtml(contact.funktion)}` : ""}
                ${contact.rolle ? ` · ${helpers.escapeHtml(contact.rolle)}` : ""}
              </div>
            </div>

            <div class="flex items-center gap-2 flex-wrap">
              ${contact.email1 ? `<a class="bbz-button bbz-button-secondary" href="mailto:${helpers.escapeHtml(contact.email1)}">Mail senden</a>` : ""}
              <button class="bbz-button bbz-button-secondary" type="button" disabled>Neue Aufgabe (später)</button>
              <button class="bbz-button bbz-button-secondary" type="button" disabled>Neue History (später)</button>
            </div>
          </section>

          <section class="grid grid-cols-1 xl:grid-cols-3 gap-4">
            <div class="bbz-card xl:col-span-1">
              <div class="bbz-card-header">Stammdaten</div>
              <div class="bbz-card-body grid grid-cols-1 gap-4">
                ${ui.kv("Anrede", helpers.escapeHtml(contact.anrede) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Vorname", helpers.escapeHtml(contact.vorname) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Nachname", helpers.escapeHtml(contact.nachname) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Firma", contact.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${contact.firmId}">${helpers.escapeHtml(contact.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>')}
                ${ui.kv("Funktion", helpers.escapeHtml(contact.funktion) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Rolle", helpers.escapeHtml(contact.rolle) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Email 1", contact.email1 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(contact.email1)}">${helpers.escapeHtml(contact.email1)}</a>` : '<span class="bbz-muted">—</span>')}
                ${ui.kv("Email 2", contact.email2 ? `<a class="bbz-link" href="mailto:${helpers.escapeHtml(contact.email2)}">${helpers.escapeHtml(contact.email2)}</a>` : '<span class="bbz-muted">—</span>')}
                ${ui.kv("Direktwahl", helpers.escapeHtml(contact.direktwahl) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Mobile", helpers.escapeHtml(contact.mobile) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Geburtstag", helpers.formatDate(contact.geburtstag) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Archiviert", contact.archiviert ? '<span class="bbz-danger">Ja</span>' : '<span class="bbz-muted">Nein</span>')}
              </div>
            </div>

            <div class="bbz-card xl:col-span-1">
              <div class="bbz-card-header">CRM-Kontext</div>
              <div class="bbz-card-body grid grid-cols-1 gap-4">
                ${ui.kv("Leadbbz0", helpers.escapeHtml(contact.leadbbz0) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("SGF", helpers.multiChoiceHtml(contact.sgf))}
                ${ui.kv("Event", helpers.multiChoiceHtml(contact.event))}
                ${ui.kv("Eventhistory", helpers.escapeHtml(contact.eventhistory) || '<span class="bbz-muted">—</span>')}
                ${ui.kv("Kommentar", helpers.escapeHtml(contact.kommentar) || '<span class="bbz-muted">—</span>')}
              </div>
            </div>

            <div class="bbz-card xl:col-span-1">
              <div class="bbz-card-header">Übersicht</div>
              <div class="bbz-card-body grid grid-cols-1 gap-4">
                ${ui.kv("Tasks", String(contactTasks.length))}
                ${ui.kv("Offene Tasks", String(contactTasks.filter(t => t.isOpen).length))}
                ${ui.kv("History-Einträge", String(contactHistory.length))}
                ${ui.kv("Letzte Aktivität", helpers.formatDateTime(contactHistory[0]?.datum) || '<span class="bbz-muted">—</span>')}
              </div>
            </div>
          </section>

          <section class="grid grid-cols-1 xl:grid-cols-2 gap-4">
            <div class="bbz-card">
              <div class="bbz-card-header">Historie</div>
              <div class="bbz-card-body">
                ${
                  contactHistory.length
                    ? `
                      <div class="bbz-timeline">
                        ${contactHistory.map(h => `
                          <div class="bbz-timeline-item">
                            <div class="bbz-timeline-date">
                              ${helpers.formatDateTime(h.datum) || "—"}
                            </div>
                            <div class="bbz-timeline-body">
                              <div class="bbz-timeline-title">
                                ${helpers.escapeHtml(h.typ || h.title || "Eintrag")}
                                ${h.projektbezugBool ? '<span class="bbz-chip">Projektbezug</span>' : '<span class="bbz-chip">Allgemein</span>'}
                              </div>
                              <div class="bbz-timeline-text">${helpers.escapeHtml(h.notizen || "—")}</div>
                            </div>
                          </div>
                        `).join("")}
                      </div>
                    `
                    : ui.emptyBlock("Keine Historie vorhanden.")
                }
              </div>
            </div>

            <div class="bbz-card">
              <div class="bbz-card-header">Tasks</div>
              <div class="bbz-card-body">
                <div class="bbz-scroll">
                  <table class="bbz-table">
                    <thead>
                      <tr>
                        <th>Titel</th>
                        <th>Deadline</th>
                        <th>Status</th>
                        <th>Firma</th>
                      </tr>
                    </thead>
                    <tbody>
                      ${
                        contactTasks.length
                          ? contactTasks.map(t => `
                            <tr>
                              <td>${helpers.escapeHtml(t.title) || '<span class="bbz-muted">—</span>'}</td>
                              <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.formatDate(t.deadline) || '<span class="bbz-muted">—</span>'}</td>
                              <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.escapeHtml(t.status) || '<span class="bbz-muted">—</span>'}</td>
                              <td>
                                ${
                                  t.firmId
                                    ? `<a class="bbz-link" data-action="open-firm" data-id="${t.firmId}">${helpers.escapeHtml(t.firmTitle || "Firma")}</a>`
                                    : '<span class="bbz-muted">—</span>'
                                }
                              </td>
                            </tr>
                          `).join("")
                          : `<tr><td colspan="4">${ui.emptyBlock("Keine Tasks vorhanden.")}</td></tr>`
                      }
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </section>
        </div>
      `;
    },

    planning() {
      const filters = state.filters.planning;

      const rows = state.enriched.tasks.filter(task => {
        const search = filters.search.trim().toLowerCase();

        const searchMatch =
          !search ||
          [
            task.title,
            task.status,
            task.contactName,
            task.firmTitle,
            task.leadbbz
          ].some(v => helpers.textIncludes(v, search));

        const openMatch = !filters.onlyOpen || task.isOpen;
        const overdueMatch = !filters.onlyOverdue || task.isOverdue;

        return searchMatch && openMatch && overdueMatch;
      });

      const openTasks = state.enriched.tasks.filter(t => t.isOpen).length;
      const overdueTasks = state.enriched.tasks.filter(t => t.isOpen && t.isOverdue).length;

      return `
        <div class="space-y-4">
          <section class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4">
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Tasks gesamt</div>
              <div class="bbz-kpi-value">${state.enriched.tasks.length}</div>
            </div>
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Offen</div>
              <div class="bbz-kpi-value">${openTasks}</div>
            </div>
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Überfällig</div>
              <div class="bbz-kpi-value">${overdueTasks}</div>
            </div>
            <div class="bbz-kpi">
              <div class="bbz-kpi-label">Erledigt / geschlossen</div>
              <div class="bbz-kpi-value">${state.enriched.tasks.length - openTasks}</div>
            </div>
          </section>

          <section class="bbz-card">
            <div class="bbz-card-header">Planung</div>
            <div class="bbz-card-body space-y-4">
              <div class="grid grid-cols-1 lg:grid-cols-4 gap-3">
                <input
                  class="bbz-input lg:col-span-2"
                  data-filter="planning-search"
                  type="text"
                  placeholder="Suche nach Titel, Firma, Kontakt, Status ..."
                  value="${helpers.escapeHtml(filters.search)}"
                />
                <label class="flex items-center gap-2 text-sm text-slate-700 h-[38px]">
                  <input
                    type="checkbox"
                    data-filter="planning-open"
                    ${filters.onlyOpen ? "checked" : ""}
                  />
                  Nur offene Tasks
                </label>
                <label class="flex items-center gap-2 text-sm text-slate-700 h-[38px]">
                  <input
                    type="checkbox"
                    data-filter="planning-overdue"
                    ${filters.onlyOverdue ? "checked" : ""}
                  />
                  Nur überfällige Tasks
                </label>
              </div>

              <div class="bbz-scroll">
                <table class="bbz-table">
                  <thead>
                    <tr>
                      <th>Titel</th>
                      <th>Deadline</th>
                      <th>Status</th>
                      <th>Kontaktperson</th>
                      <th>Firma</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${
                      rows.length
                        ? rows.map(t => `
                          <tr>
                            <td>${helpers.escapeHtml(t.title) || '<span class="bbz-muted">—</span>'}</td>
                            <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.formatDate(t.deadline) || '<span class="bbz-muted">—</span>'}</td>
                            <td class="${helpers.statusClass(t.status, t.deadline)}">${helpers.escapeHtml(t.status) || '<span class="bbz-muted">—</span>'}</td>
                            <td>
                              ${
                                t.contactId
                                  ? `<a class="bbz-link" data-action="open-contact" data-id="${t.contactId}">${helpers.escapeHtml(t.contactName || "Kontakt")}</a>`
                                  : helpers.escapeHtml(t.contactName || "—")
                              }
                            </td>
                            <td>
                              ${
                                t.firmId
                                  ? `<a class="bbz-link" data-action="open-firm" data-id="${t.firmId}">${helpers.escapeHtml(t.firmTitle || "Firma")}</a>`
                                  : '<span class="bbz-muted">—</span>'
                              }
                            </td>
                          </tr>
                        `).join("")
                        : `<tr><td colspan="5">${ui.emptyBlock("Keine Tasks für die aktuelle Filterung gefunden.")}</td></tr>`
                    }
                  </tbody>
                </table>
              </div>
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
          await api.loadAll();
          ui.setMessage("Anmeldung erkannt. Daten wurden geladen.", "success");
        } else {
          ui.setMessage("Bitte anmelden, um die SharePoint-Listen über Microsoft Graph zu laden.", "warning");
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
        if (!state.auth.isReady) {
          ui.setMessage("Authentifizierung ist noch nicht bereit. Bitte Seite einmal neu laden.", "warning");
          return;
        }

        ui.setLoading(true);
        ui.setMessage("");

        await api.login();
        await api.loadAll();

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
      if (!state.auth.isReady) {
        ui.setMessage("Authentifizierung ist noch nicht bereit.", "warning");
        return;
      }

      if (!state.auth.isAuthenticated) {
        ui.setMessage("Bitte zuerst anmelden.", "warning");
        return;
      }

      try {
        ui.setLoading(true);
        ui.setMessage("");

        await api.acquireToken();
        await api.loadAll();

        ui.setMessage("Daten erfolgreich neu geladen.", "success");
      } catch (error) {
        console.error(error);
        ui.setMessage(`Fehler beim Laden: ${error.message}`, "error");
      } finally {
        ui.setLoading(false);
        this.render();
      }
    },

    navigate(route) {
      state.filters.route = route;
      if (route !== "firms") state.selection.firmId = null;
      if (route !== "contacts") state.selection.contactId = null;
      this.render();
    },

    openFirm(id) {
      state.selection.firmId = id;
      state.selection.contactId = null;
      state.filters.route = "firms";
      this.render();
    },

    openContact(id) {
      state.selection.contactId = id;
      state.filters.route = "contacts";
      this.render();
    },

    render() {
      ui.renderShell();
      ui.renderView(views.renderRoute());
    }
  };

  document.addEventListener("DOMContentLoaded", () => {
    controller.init();
  });
})();

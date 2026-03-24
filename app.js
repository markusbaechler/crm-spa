(() => {
  "use strict";

  const CONFIG = {
    appName: "bbz CRM",

    graph: {
      tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
      clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
      authority: "https://login.microsoftonline.com/3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
      redirectUri: "https://markusbaechler.github.io/crm-spa/",
      scopes: ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"]
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
      lastError: null,
      saving: false
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
      },

      events: {
        search: "",
        onlyWithOpenTasks: false
      }
    },

    selection: {
      firmId: null,
      contactId: null
    },

    modal: {
      type: null,
      payload: null
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
        if (value.includes(";#")) return value.split(";#").map(v => v.trim()).filter(Boolean);
        if (value.includes(",")) return value.split(",").map(v => v.trim()).filter(Boolean);
        return [value.trim()].filter(Boolean);
      }
      return [value];
    },

    normalizeChoiceList(value) {
      return helpers.toArray(value).filter(Boolean);
    },

    splitCsv(value) {
      return String(value || "")
        .split(",")
        .map(v => v.trim())
        .filter(Boolean);
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

    toDateInput(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const dd = String(d.getDate()).padStart(2, "0");
      return `${yyyy}-${mm}-${dd}`;
    },

    toDateTimeLocalInput(value) {
      const d = helpers.toDate(value);
      if (!d) return "";
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const dd = String(d.getDate()).padStart(2, "0");
      const hh = String(d.getHours()).padStart(2, "0");
      const mi = String(d.getMinutes()).padStart(2, "0");
      return `${yyyy}-${mm}-${dd}T${hh}:${mi}`;
    },

    fromDateTimeLocalInput(value) {
      if (!value) return null;
      const d = new Date(value);
      if (Number.isNaN(d.getTime())) return null;
      return d.toISOString();
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
      if (missing.length) throw new Error(`Konfiguration unvollständig: ${missing.join(", ")}`);
    }
  };

  const ui = {
    els: {
      viewRoot: null,
      authStatus: null,
      globalMessage: null,
      btnLogin: null,
      btnRefresh: null,
      navButtons: [],
      modalRoot: null
    },

    init() {
      this.els.viewRoot = document.getElementById("view-root");
      this.els.authStatus = document.getElementById("auth-status");
      this.els.globalMessage = document.getElementById("global-message");
      this.els.btnLogin = document.getElementById("btn-login");
      this.els.btnRefresh = document.getElementById("btn-refresh");
      this.els.navButtons = [...document.querySelectorAll(".bbz-nav-btn")];
      this.els.modalRoot = document.getElementById("modal-root");

      if (this.els.btnLogin) this.els.btnLogin.addEventListener("click", () => controller.handleLogin());
      if (this.els.btnRefresh) this.els.btnRefresh.addEventListener("click", () => controller.handleRefresh());

      this.els.navButtons.forEach(btn => {
        btn.addEventListener("click", () => controller.navigate(btn.dataset.route));
      });

      document.addEventListener("click", (event) => {
        const target = event.target;

        const openFirm = target.closest("[data-action='open-firm']");
        if (openFirm) return controller.openFirm(openFirm.dataset.id);

        const openContact = target.closest("[data-action='open-contact']");
        if (openContact) return controller.openContact(openContact.dataset.id);

        const backToFirms = target.closest("[data-action='back-to-firms']");
        if (backToFirms) return controller.navigate("firms");

        const backToContacts = target.closest("[data-action='back-to-contacts']");
        if (backToContacts) return controller.navigate("contacts");

        const modalOpen = target.closest("[data-open-modal]");
        if (modalOpen) {
          modal.open(modalOpen.dataset.openModal, modalOpen.dataset);
          return;
        }

        const modalClose = target.closest("[data-close-modal]");
        if (modalClose) {
          modal.close();
        }
      });

      document.addEventListener("input", (event) => {
        const el = event.target;
        if (el.matches("[data-filter='firms-search']")) state.filters.firms.search = el.value;
        if (el.matches("[data-filter='contacts-search']")) state.filters.contacts.search = el.value;
        if (el.matches("[data-filter='planning-search']")) state.filters.planning.search = el.value;
        if (el.matches("[data-filter='events-search']")) state.filters.events.search = el.value;
        controller.render();
      });

      document.addEventListener("change", (event) => {
        const el = event.target;
        if (el.matches("[data-filter='firms-klassifizierung']")) state.filters.firms.klassifizierung = el.value;
        if (el.matches("[data-filter='firms-vip']")) state.filters.firms.vip = el.value;
        if (el.matches("[data-filter='contacts-archiviert']")) state.filters.contacts.archiviertAusblenden = el.checked;
        if (el.matches("[data-filter='planning-open']")) state.filters.planning.onlyOpen = el.checked;
        if (el.matches("[data-filter='planning-overdue']")) state.filters.planning.onlyOverdue = el.checked;
        if (el.matches("[data-filter='events-open']")) state.filters.events.onlyWithOpenTasks = el.checked;
        controller.render();
      });

      document.addEventListener("submit", async (event) => {
        const form = event.target;
        if (!form.matches("[data-modal-form]")) return;

        event.preventDefault();

        try {
          state.meta.saving = true;
          modal.setSavingState(true);

          switch (form.dataset.modalForm) {
            case "firm":
              await actions.saveFirm(new FormData(form), form.dataset.mode, form.dataset.itemId);
              break;
            case "contact":
              await actions.saveContact(new FormData(form), form.dataset.mode, form.dataset.itemId);
              break;
            case "task":
              await actions.saveTask(new FormData(form));
              break;
            case "history":
              await actions.saveHistory(new FormData(form));
              break;
            default:
              throw new Error("Unbekannter Formular-Typ.");
          }

          modal.close();
          await controller.reloadData();
        } catch (error) {
          console.error(error);
          ui.setMessage(`Speichern fehlgeschlagen: ${error.message}`, "error");
        } finally {
          state.meta.saving = false;
          modal.setSavingState(false);
        }
      });

      document.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && state.modal.type) modal.close();
      });
    },

    setLoading(isLoading) {
      state.meta.loading = isLoading;
      this.renderShell();
    },

    setMessage(message, type = "info") {
      const el = this.els.globalMessage;
      if (!el) return;

      if (!message) {
        el.className = "bbz-banner";
        el.textContent = "";
        return;
      }

      let cls = "bbz-banner bbz-banner-info show";
      if (type === "success") cls = "bbz-banner bbz-banner-success show";
      if (type === "warning") cls = "bbz-banner bbz-banner-warning show";
      if (type === "error") cls = "bbz-banner bbz-banner-error show";

      el.className = cls;
      el.textContent = message;
    },

    renderShell() {
      this.els.navButtons.forEach(btn => {
        btn.classList.toggle("active", btn.dataset.route === state.filters.route);
      });

      if (state.auth.isAuthenticated && state.auth.account) {
        this.els.authStatus.innerHTML = `
          <span class="bbz-auth-dot"></span>
          <span>Angemeldet: ${helpers.escapeHtml(state.auth.account.username || state.auth.account.name || "")}</span>
        `;
      } else if (state.auth.isReady) {
        this.els.authStatus.innerHTML = `
          <span class="bbz-auth-dot" style="background:#94a3b8;"></span>
          <span>Nicht angemeldet</span>
        `;
      } else {
        this.els.authStatus.innerHTML = `
          <span class="bbz-auth-dot" style="background:#f59e0b;"></span>
          <span>Authentifizierung wird initialisiert ...</span>
        `;
      }

      if (this.els.btnLogin) {
        this.els.btnLogin.textContent = state.auth.isAuthenticated ? "Erneut anmelden" : "Anmelden";
        this.els.btnLogin.disabled = state.meta.loading || !state.auth.isReady || state.meta.saving;
      }

      if (this.els.btnRefresh) {
        this.els.btnRefresh.disabled = state.meta.loading || !state.auth.isReady || state.meta.saving;
      }
    },

    renderView(html) {
      if (this.els.viewRoot) this.els.viewRoot.innerHTML = html;
    },

    loadingBlock(text = "Daten werden geladen ...") {
      return `
        <section class="bbz-section">
          <div class="bbz-section-body">
            <div class="flex items-center gap-3">
              <div class="bbz-loader"></div>
              <div class="text-sm text-slate-500">${helpers.escapeHtml(text)}</div>
            </div>
          </div>
        </section>
      `;
    },

    emptyBlock(text = "Keine Daten vorhanden.") {
      return `<div class="bbz-empty">${helpers.escapeHtml(text)}</div>`;
    },

    kv(label, value) {
      return `
        <div class="bbz-kv">
          <div class="bbz-kv-label">${helpers.escapeHtml(label)}</div>
          <div class="bbz-kv-value">${value || '<span class="bbz-muted">—</span>'}</div>
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

      const accounts = state.auth.msal.getAllAccounts();
      if (accounts.length > 0 && !state.auth.account) {
        state.auth.account = accounts[0];
        state.auth.isAuthenticated = true;
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

    async acquireToken() {
      if (!state.auth.msal) throw new Error("MSAL ist nicht initialisiert.");
      if (!state.auth.account) throw new Error("Kein angemeldetes Konto gefunden.");

      try {
        const tokenResponse = await state.auth.msal.acquireTokenSilent({
          account: state.auth.account,
          scopes: CONFIG.graph.scopes
        });
        if (!tokenResponse?.accessToken) throw new Error("Kein Access Token aus acquireTokenSilent erhalten.");
        state.auth.token = tokenResponse.accessToken;
        return state.auth.token;
      } catch {
        const tokenResponse = await state.auth.msal.acquireTokenPopup({
          account: state.auth.account,
          scopes: CONFIG.graph.scopes
        });
        if (!tokenResponse?.accessToken) throw new Error("Kein Access Token aus acquireTokenPopup erhalten.");
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
        try { detail = await response.text(); } catch { detail = ""; }
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

    async createListItem(listTitle, fields) {
      const siteId = await this.getSiteId();
      return this.graphRequest(`/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/items`, {
        method: "POST",
        body: { fields }
      });
    },

    async updateListItemFields(listTitle, itemId, fields) {
      const siteId = await this.getSiteId();
      return this.graphRequest(`/sites/${siteId}/lists/${encodeURIComponent(listTitle)}/items/${itemId}/fields`, {
        method: "PATCH",
        body: fields
      });
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
          firm
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
        const latestHistory = contactHistory[0] || null;
        const openTasks = contactTasks.filter(t => t.isOpen);

        contact.event.forEach(eventName => {
          const key = String(eventName || "").trim();
          if (!key) return;

          if (!eventMap.has(key)) {
            eventMap.set(key, {
              name: key,
              contacts: [],
              contactCount: 0,
              openTasksCount: 0
            });
          }

          eventMap.get(key).contacts.push({
            contactId: contact.id,
            contactName: contact.fullName || contact.nachname,
            firmId: contact.firmId,
            firmTitle: contact.firmTitle,
            rolle: contact.rolle,
            funktion: contact.funktion,
            eventhistory: contact.eventhistory,
            latestHistoryDate: latestHistory?.datum || "",
            latestHistoryType: latestHistory?.typ || "",
            latestHistoryText: latestHistory?.notizen || "",
            openTasksCount: openTasks.length,
            email1: contact.email1
          });
        });
      });

      const events = [...eventMap.values()].map(group => ({
        ...group,
        contactCount: group.contacts.length,
        openTasksCount: group.contacts.reduce((sum, c) => sum + c.openTasksCount, 0),
        contacts: group.contacts.sort((a, b) => String(a.contactName).localeCompare(String(b.contactName), "de"))
      })).sort((a, b) => a.name.localeCompare(b.name, "de"));

      state.enriched.contacts = contacts.sort((a, b) => a.fullName.localeCompare(b.fullName, "de"));
      state.enriched.history = history.sort((a, b) => helpers.compareDateDesc(a.datum, b.datum));
      state.enriched.tasks = tasks.sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline));
      state.enriched.firms = firms.sort((a, b) => a.title.localeCompare(b.title, "de"));
      state.enriched.events = events;
    },

    getFirmById(id) {
      return state.enriched.firms.find(f => String(f.id) === String(id)) || null;
    },

    getContactById(id) {
      return state.enriched.contacts.find(c => String(c.id) === String(id)) || null;
    }
  };

  const modal = {
    open(type, payload = {}) {
      state.modal.type = type;
      state.modal.payload = payload;
      this.render();
    },

    close() {
      state.modal.type = null;
      state.modal.payload = null;
      this.render();
    },

    setSavingState(isSaving) {
      const submit = document.querySelector("[data-modal-submit]");
      const closeButtons = [...document.querySelectorAll("[data-close-modal]")];
      if (submit) {
        submit.disabled = isSaving;
        submit.textContent = isSaving ? "Speichern ..." : (submit.dataset.defaultLabel || "Speichern");
      }
      closeButtons.forEach(btn => { btn.disabled = isSaving; });
    },

    render() {
      const root = ui.els.modalRoot;
      if (!root) return;

      if (!state.modal.type) {
        root.innerHTML = "";
        return;
      }

      let html = "";
      switch (state.modal.type) {
        case "firm-create":
          html = this.renderFirmForm("create");
          break;
        case "firm-edit":
          html = this.renderFirmForm("edit", Number(state.modal.payload.itemId));
          break;
        case "contact-create":
          html = this.renderContactForm("create", state.modal.payload);
          break;
        case "contact-edit":
          html = this.renderContactForm("edit", state.modal.payload);
          break;
        case "task-create":
          html = this.renderTaskForm(state.modal.payload);
          break;
        case "history-create":
          html = this.renderHistoryForm(state.modal.payload);
          break;
        default:
          html = "";
      }

      root.innerHTML = html;
    },

    renderFirmForm(mode, itemId = null) {
      const firm = mode === "edit" ? dataModel.getFirmById(itemId) : null;
      const title = mode === "edit" ? "Firma bearbeiten" : "Neue Firma";

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${title}</div>
              <button class="bbz-button bbz-button-secondary" data-close-modal>Schließen</button>
            </div>

            <form data-modal-form="firm" data-mode="${mode}" data-item-id="${itemId || ""}">
              <div class="bbz-modal-body">
                <div class="bbz-form-grid">
                  <div class="bbz-field">
                    <label>Firma</label>
                    <input class="bbz-input" name="title" required value="${helpers.escapeHtml(firm?.title || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Klassifizierung</label>
                    <select class="bbz-select" name="klassifizierung">
                      <option value="">Bitte wählen</option>
                      ${["A-Kunde", "B-Kunde", "C-Kunde", "A", "B", "C"].map(v => `<option value="${helpers.escapeHtml(v)}" ${(firm?.klassifizierung || "") === v ? "selected" : ""}>${helpers.escapeHtml(v)}</option>`).join("")}
                    </select>
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
                    <input class="bbz-input" name="land" value="${helpers.escapeHtml(firm?.land || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Hauptnummer</label>
                    <input class="bbz-input" name="hauptnummer" value="${helpers.escapeHtml(firm?.hauptnummer || "")}" />
                  </div>

                  <label class="bbz-checkbox">
                    <input type="checkbox" name="vip" ${firm?.vip ? "checked" : ""} />
                    VIP
                  </label>
                </div>
              </div>

              <div class="bbz-modal-footer">
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary" data-modal-submit data-default-label="Speichern">Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    renderContactForm(mode, payload = {}) {
      const itemId = Number(payload.itemId || 0) || null;
      const contact = mode === "edit" ? dataModel.getContactById(itemId) : null;
      const title = mode === "edit" ? "Kontakt bearbeiten" : "Neuer Kontakt";
      const preselectedFirmId = Number(payload.prefillFirmId || contact?.firmId || 0) || "";

      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">${title}</div>
              <button class="bbz-button bbz-button-secondary" data-close-modal>Schließen</button>
            </div>

            <form data-modal-form="contact" data-mode="${mode}" data-item-id="${itemId || ""}">
              <div class="bbz-modal-body">
                <div class="bbz-form-grid">
                  <div class="bbz-field">
                    <label>Nachname</label>
                    <input class="bbz-input" name="nachname" required value="${helpers.escapeHtml(contact?.nachname || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Vorname</label>
                    <input class="bbz-input" name="vorname" value="${helpers.escapeHtml(contact?.vorname || "")}" />
                  </div>

                  <div class="bbz-field">
                    <label>Anrede</label>
                    <input class="bbz-input" name="anrede" value="${helpers.escapeHtml(contact?.anrede || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Firma</label>
                    <select class="bbz-select" name="firmaLookupId" required>
                      <option value="">Bitte wählen</option>
                      ${state.enriched.firms.map(f => `<option value="${f.id}" ${String(preselectedFirmId) === String(f.id) ? "selected" : ""}>${helpers.escapeHtml(f.title)}</option>`).join("")}
                    </select>
                  </div>

                  <div class="bbz-field">
                    <label>Funktion</label>
                    <input class="bbz-input" name="funktion" value="${helpers.escapeHtml(contact?.funktion || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Rolle</label>
                    <input class="bbz-input" name="rolle" value="${helpers.escapeHtml(contact?.rolle || "")}" />
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
                    <label>Leadbbz0</label>
                    <input class="bbz-input" name="leadbbz0" value="${helpers.escapeHtml(contact?.leadbbz0 || "")}" />
                  </div>
                  <div class="bbz-field">
                    <label>Geburtstag</label>
                    <input type="date" class="bbz-input" name="geburtstag" value="${helpers.escapeHtml(helpers.toDateInput(contact?.geburtstag || ""))}" />
                  </div>

                  <div class="bbz-field">
                    <label>SGF (Komma getrennt)</label>
                    <input class="bbz-input" name="sgf" value="${helpers.escapeHtml((contact?.sgf || []).join(", "))}" />
                  </div>
                  <div class="bbz-field">
                    <label>Event (Komma getrennt)</label>
                    <input class="bbz-input" name="event" value="${helpers.escapeHtml((contact?.event || []).join(", "))}" />
                  </div>

                  <div class="bbz-field bbz-span-2">
                    <label>Eventhistory</label>
                    <textarea class="bbz-textarea" name="eventhistory">${helpers.escapeHtml(contact?.eventhistory || "")}</textarea>
                  </div>

                  <div class="bbz-field bbz-span-2">
                    <label>Kommentar</label>
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
                <button type="submit" class="bbz-button bbz-button-primary" data-modal-submit data-default-label="Speichern">Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    renderTaskForm(payload = {}) {
      const prefillContactId = Number(payload.prefillContactId || 0) || "";
      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">Neue Aufgabe</div>
              <button class="bbz-button bbz-button-secondary" data-close-modal>Schließen</button>
            </div>

            <form data-modal-form="task">
              <div class="bbz-modal-body">
                <div class="bbz-form-grid">
                  <div class="bbz-field bbz-span-2">
                    <label>Titel</label>
                    <input class="bbz-input" name="title" required />
                  </div>

                  <div class="bbz-field">
                    <label>Kontakt</label>
                    <select class="bbz-select" name="kontaktLookupId" required>
                      <option value="">Bitte wählen</option>
                      ${state.enriched.contacts.map(c => `<option value="${c.id}" ${String(prefillContactId) === String(c.id) ? "selected" : ""}>${helpers.escapeHtml(c.fullName)}${c.firmTitle ? ` · ${helpers.escapeHtml(c.firmTitle)}` : ""}</option>`).join("")}
                    </select>
                  </div>

                  <div class="bbz-field">
                    <label>Deadline</label>
                    <input type="date" class="bbz-input" name="deadline" required />
                  </div>

                  <div class="bbz-field">
                    <label>Status</label>
                    <select class="bbz-select" name="status">
                      ${["Offen", "In Arbeit", "Wartend", "Erledigt"].map(v => `<option value="${helpers.escapeHtml(v)}">${helpers.escapeHtml(v)}</option>`).join("")}
                    </select>
                  </div>

                  <div class="bbz-field">
                    <label>Leadbbz</label>
                    <input class="bbz-input" name="leadbbz" />
                  </div>
                </div>
              </div>

              <div class="bbz-modal-footer">
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary" data-modal-submit data-default-label="Speichern">Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    },

    renderHistoryForm(payload = {}) {
      const prefillContactId = Number(payload.prefillContactId || 0) || "";
      return `
        <div class="bbz-modal-backdrop show">
          <div class="bbz-modal">
            <div class="bbz-modal-header">
              <div class="bbz-modal-title">Neue History</div>
              <button class="bbz-button bbz-button-secondary" data-close-modal>Schließen</button>
            </div>

            <form data-modal-form="history">
              <div class="bbz-modal-body">
                <div class="bbz-form-grid">
                  <div class="bbz-field bbz-span-2">
                    <label>Titel</label>
                    <input class="bbz-input" name="title" required />
                  </div>

                  <div class="bbz-field">
                    <label>Kontakt</label>
                    <select class="bbz-select" name="kontaktLookupId" required>
                      <option value="">Bitte wählen</option>
                      ${state.enriched.contacts.map(c => `<option value="${c.id}" ${String(prefillContactId) === String(c.id) ? "selected" : ""}>${helpers.escapeHtml(c.fullName)}${c.firmTitle ? ` · ${helpers.escapeHtml(c.firmTitle)}` : ""}</option>`).join("")}
                    </select>
                  </div>

                  <div class="bbz-field">
                    <label>Datum / Zeit</label>
                    <input type="datetime-local" class="bbz-input" name="datum" value="${helpers.toDateTimeLocalInput(new Date().toISOString())}" required />
                  </div>

                  <div class="bbz-field">
                    <label>Typ</label>
                    <select class="bbz-select" name="typ">
                      ${["Meeting", "Call", "Mail", "Notiz"].map(v => `<option value="${helpers.escapeHtml(v)}">${helpers.escapeHtml(v)}</option>`).join("")}
                    </select>
                  </div>

                  <div class="bbz-field">
                    <label>Leadbbz</label>
                    <input class="bbz-input" name="leadbbz" />
                  </div>

                  <label class="bbz-checkbox">
                    <input type="checkbox" name="projektbezug" />
                    Projektbezug
                  </label>

                  <div class="bbz-field bbz-span-2">
                    <label>Notizen</label>
                    <textarea class="bbz-textarea" name="notizen"></textarea>
                  </div>
                </div>
              </div>

              <div class="bbz-modal-footer">
                <button type="button" class="bbz-button bbz-button-secondary" data-close-modal>Abbrechen</button>
                <button type="submit" class="bbz-button bbz-button-primary" data-modal-submit data-default-label="Speichern">Speichern</button>
              </div>
            </form>
          </div>
        </div>
      `;
    }
  };

  const actions = {
    async saveFirm(formData, mode, itemId) {
      const fields = {
        [SCHEMA.firms.fields.title]: String(formData.get("title") || "").trim(),
        [SCHEMA.firms.fields.adresse]: String(formData.get("adresse") || "").trim(),
        [SCHEMA.firms.fields.plz]: String(formData.get("plz") || "").trim(),
        [SCHEMA.firms.fields.ort]: String(formData.get("ort") || "").trim(),
        [SCHEMA.firms.fields.land]: String(formData.get("land") || "").trim(),
        [SCHEMA.firms.fields.hauptnummer]: String(formData.get("hauptnummer") || "").trim(),
        [SCHEMA.firms.fields.klassifizierung]: String(formData.get("klassifizierung") || "").trim(),
        [SCHEMA.firms.fields.vip]: formData.get("vip") === "on"
      };

      if (!fields[SCHEMA.firms.fields.title]) {
        throw new Error("Firmenname fehlt.");
      }

      if (mode === "edit" && itemId) {
        await api.updateListItemFields(SCHEMA.firms.listTitle, itemId, fields);
        ui.setMessage("Firma gespeichert.", "success");
      } else {
        await api.createListItem(SCHEMA.firms.listTitle, fields);
        ui.setMessage("Firma angelegt.", "success");
      }
    },

    async saveContact(formData, mode, itemId) {
      const firmaLookupId = Number(formData.get("firmaLookupId") || 0);
      if (!firmaLookupId) throw new Error("Bitte eine Firma auswählen.");

      const fields = {
        [SCHEMA.contacts.fields.nachname]: String(formData.get("nachname") || "").trim(),
        [SCHEMA.contacts.fields.vorname]: String(formData.get("vorname") || "").trim(),
        [SCHEMA.contacts.fields.anrede]: String(formData.get("anrede") || "").trim(),
        [SCHEMA.contacts.fields.firmaLookupId]: firmaLookupId,
        [SCHEMA.contacts.fields.funktion]: String(formData.get("funktion") || "").trim(),
        [SCHEMA.contacts.fields.email1]: String(formData.get("email1") || "").trim(),
        [SCHEMA.contacts.fields.email2]: String(formData.get("email2") || "").trim(),
        [SCHEMA.contacts.fields.direktwahl]: String(formData.get("direktwahl") || "").trim(),
        [SCHEMA.contacts.fields.mobile]: String(formData.get("mobile") || "").trim(),
        [SCHEMA.contacts.fields.rolle]: String(formData.get("rolle") || "").trim(),
        [SCHEMA.contacts.fields.leadbbz0]: String(formData.get("leadbbz0") || "").trim(),
        [SCHEMA.contacts.fields.sgf]: helpers.splitCsv(formData.get("sgf")),
        [SCHEMA.contacts.fields.geburtstag]: formData.get("geburtstag") ? new Date(`${formData.get("geburtstag")}T00:00:00`).toISOString() : null,
        [SCHEMA.contacts.fields.kommentar]: String(formData.get("kommentar") || "").trim(),
        [SCHEMA.contacts.fields.event]: helpers.splitCsv(formData.get("event")),
        [SCHEMA.contacts.fields.eventhistory]: String(formData.get("eventhistory") || "").trim(),
        [SCHEMA.contacts.fields.archiviert]: formData.get("archiviert") === "on"
      };

      if (!fields[SCHEMA.contacts.fields.nachname]) throw new Error("Nachname fehlt.");

      if (mode === "edit" && itemId) {
        await api.updateListItemFields(SCHEMA.contacts.listTitle, itemId, fields);
        ui.setMessage("Kontakt gespeichert.", "success");
      } else {
        await api.createListItem(SCHEMA.contacts.listTitle, fields);
        ui.setMessage("Kontakt angelegt.", "success");
      }
    },

    async saveTask(formData) {
      const kontaktLookupId = Number(formData.get("kontaktLookupId") || 0);
      if (!kontaktLookupId) throw new Error("Bitte einen Kontakt auswählen.");

      const deadline = formData.get("deadline")
        ? new Date(`${formData.get("deadline")}T00:00:00`).toISOString()
        : null;

      const fields = {
        [SCHEMA.tasks.fields.title]: String(formData.get("title") || "").trim(),
        [SCHEMA.tasks.fields.kontaktLookupId]: kontaktLookupId,
        [SCHEMA.tasks.fields.deadline]: deadline,
        [SCHEMA.tasks.fields.status]: String(formData.get("status") || "").trim(),
        [SCHEMA.tasks.fields.leadbbz]: String(formData.get("leadbbz") || "").trim()
      };

      if (!fields[SCHEMA.tasks.fields.title]) throw new Error("Titel fehlt.");
      await api.createListItem(SCHEMA.tasks.listTitle, fields);
      ui.setMessage("Task angelegt.", "success");
    },

    async saveHistory(formData) {
      const kontaktLookupId = Number(formData.get("kontaktLookupId") || 0);
      if (!kontaktLookupId) throw new Error("Bitte einen Kontakt auswählen.");

      const datum = helpers.fromDateTimeLocalInput(formData.get("datum"));
      if (!datum) throw new Error("Datum/Zeit fehlt oder ist ungültig.");

      const fields = {
        [SCHEMA.history.fields.title]: String(formData.get("title") || "").trim(),
        [SCHEMA.history.fields.kontaktLookupId]: kontaktLookupId,
        [SCHEMA.history.fields.datum]: datum,
        [SCHEMA.history.fields.typ]: String(formData.get("typ") || "").trim(),
        [SCHEMA.history.fields.notizen]: String(formData.get("notizen") || "").trim(),
        [SCHEMA.history.fields.projektbezug]: formData.get("projektbezug") === "on",
        [SCHEMA.history.fields.leadbbz]: String(formData.get("leadbbz") || "").trim()
      };

      if (!fields[SCHEMA.history.fields.title]) throw new Error("Titel fehlt.");
      await api.createListItem(SCHEMA.history.listTitle, fields);
      ui.setMessage("History-Eintrag angelegt.", "success");
    }
  };

  const views = {
    kpiBlock(label, value, meta = "") {
      return `
        <div class="bbz-kpi">
          <div class="bbz-kpi-label">${helpers.escapeHtml(label)}</div>
          <div class="bbz-kpi-value">${helpers.escapeHtml(String(value))}</div>
          ${meta ? `<div class="bbz-kpi-meta">${helpers.escapeHtml(meta)}</div>` : ""}
        </div>
      `;
    },

    miniItem(title, meta) {
      return `
        <div class="bbz-mini-item">
          <div class="bbz-mini-title">${title}</div>
          <div class="bbz-mini-meta">${meta}</div>
        </div>
      `;
    },

    actionBar(items) {
      return `<div class="flex items-center gap-2 flex-wrap">${items.join("")}</div>`;
    },

    renderRoute() {
      if (state.meta.loading) return ui.loadingBlock();

      switch (state.filters.route) {
        case "firms":
          return state.selection.firmId ? this.firmDetail() : this.firms();
        case "contacts":
          return state.selection.contactId ? this.contactDetail() : this.contacts();
        case "planning":
          return this.planning();
        case "events":
          return this.events();
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
          [firm.title, firm.ort, firm.klassifizierung, firm.hauptnummer, firm.adresse, firm.land, ...firm.contacts.map(c => c.fullName)].some(v => helpers.textIncludes(v, search));
        const klassifizierungMatch =
          !filters.klassifizierung || String(firm.klassifizierung || "").toLowerCase() === String(filters.klassifizierung || "").toLowerCase();
        const vipMatch =
          !filters.vip || (filters.vip === "yes" && firm.vip) || (filters.vip === "no" && !firm.vip);
        return searchMatch && klassifizierungMatch && vipMatch;
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
            ${this.kpiBlock("Überfällige Tasks", overdueTasks, "sofort prüfen")}
          </div>

          <div class="bbz-grid bbz-grid-70-30">
            <section class="bbz-section">
              <div class="bbz-section-header">
                <div>
                  <div class="bbz-section-title">Firmen-Cockpit</div>
                  <div class="bbz-section-subtitle">Hauptarbeitsliste mit Fokus auf Segment, Tasks und Fristen</div>
                </div>
                ${this.actionBar([
                  `<button class="bbz-button bbz-button-primary" data-open-modal="firm-create">Neue Firma</button>`
                ])}
              </div>

              <div class="bbz-section-body">
                <div class="bbz-filters-3">
                  <input class="bbz-input" data-filter="firms-search" type="text" placeholder="Suche nach Firma, Ort, Ansprechpartner, Telefon ..." value="${helpers.escapeHtml(filters.search)}" />
                  <select class="bbz-select" data-filter="firms-klassifizierung">
                    <option value="">Alle Klassifizierungen</option>
                    ${["A-Kunde", "B-Kunde", "C-Kunde", "A", "B", "C"].map(v => `<option value="${helpers.escapeHtml(v)}" ${filters.klassifizierung === v ? "selected" : ""}>${helpers.escapeHtml(v)}</option>`).join("")}
                  </select>
                  <select class="bbz-select" data-filter="firms-vip">
                    <option value="">VIP egal</option>
                    <option value="yes" ${filters.vip === "yes" ? "selected" : ""}>Nur VIP</option>
                    <option value="no" ${filters.vip === "no" ? "selected" : ""}>Nur nicht VIP</option>
                  </select>
                </div>

                <div class="bbz-table-wrap">
                  <table class="bbz-table">
                    <thead>
                      <tr>
                        <th>Firma</th>
                        <th>Ort</th>
                        <th>Klassifizierung</th>
                        <th>VIP</th>
                        <th>Kontakte</th>
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
                                <a class="bbz-link" data-action="open-firm" data-id="${firm.id}">${helpers.escapeHtml(firm.title)}</a>
                                <div class="bbz-subtext">${helpers.escapeHtml(firm.hauptnummer || "—")}</div>
                              </td>
                              <td>${helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ")) || '<span class="bbz-muted">—</span>'}</td>
                              <td>${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : '<span class="bbz-muted">—</span>'}</td>
                              <td>${firm.vip ? '<span class="bbz-pill bbz-pill-vip">VIP</span>' : '<span class="bbz-muted">—</span>'}</td>
                              <td>${firm.contactsCount}</td>
                              <td>${firm.openTasksCount}</td>
                              <td class="${firm.nextDeadline && helpers.isOverdue(firm.nextDeadline) ? 'bbz-danger' : ''}">${helpers.formatDate(firm.nextDeadline) || '<span class="bbz-muted">—</span>'}</td>
                            </tr>
                          `).join("")
                          : `<tr><td colspan="7">${ui.emptyBlock("Keine Firmen für die aktuelle Filterung gefunden.")}</td></tr>`
                      }
                    </tbody>
                  </table>
                </div>
              </div>
            </section>

            <div class="bbz-cockpit-stack">
              <section class="bbz-section">
                <div class="bbz-section-header">
                  <div>
                    <div class="bbz-section-title">Dringend</div>
                    <div class="bbz-section-subtitle">Firmen mit offenen Tasks</div>
                  </div>
                </div>
                <div class="bbz-section-body">
                  ${
                    urgentFirms.length
                      ? `<div class="bbz-mini-list">
                          ${urgentFirms.map(f => this.miniItem(
                            `<a class="bbz-link" data-action="open-firm" data-id="${f.id}">${helpers.escapeHtml(f.title)}</a>`,
                            `${f.openTasksCount} offene Tasks · nächste Deadline ${helpers.formatDate(f.nextDeadline) || "—"}`
                          )).join("")}
                        </div>`
                      : ui.emptyBlock("Keine dringenden Firmen.")
                  }
                </div>
              </section>

              <section class="bbz-section">
                <div class="bbz-section-header">
                  <div>
                    <div class="bbz-section-title">Zuletzt aktiv</div>
                    <div class="bbz-section-subtitle">Firmen mit jüngster History</div>
                  </div>
                </div>
                <div class="bbz-section-body">
                  ${
                    latestFirms.length
                      ? `<div class="bbz-mini-list">
                          ${latestFirms.map(f => this.miniItem(
                            `<a class="bbz-link" data-action="open-firm" data-id="${f.id}">${helpers.escapeHtml(f.title)}</a>`,
                            `Letzte Aktivität ${helpers.formatDateTime(f.latestActivity) || "—"}`
                          )).join("")}
                        </div>`
                      : ui.emptyBlock("Noch keine Aktivitäten vorhanden.")
                  }
                </div>
              </section>
            </div>
          </div>
        </div>
      `;
    },

    firmDetail() {
      const firm = dataModel.getFirmById(state.selection.firmId);
      if (!firm) return ui.emptyBlock("Die ausgewählte Firma wurde nicht gefunden.");

      const recentHistory = [...firm.history].slice(0, 20);
      const firmTasks = [...firm.tasks];
      const contacts = [...firm.contacts];

      return `
        <div>
          <div class="bbz-detail-header">
            <div>
              <button class="bbz-button bbz-button-secondary mb-3" data-action="back-to-firms">Zurück zur Firmenliste</button>
              <div class="bbz-detail-title">${helpers.escapeHtml(firm.title)}</div>
              <div class="bbz-detail-subtitle">${helpers.escapeHtml(helpers.joinNonEmpty([firm.adresse, helpers.joinNonEmpty([firm.plz, firm.ort], " "), firm.land], " · ")) || "Keine erweiterten Stammdaten"}</div>
              <div class="flex items-center gap-2 flex-wrap mt-3">
                ${firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : ""}
                ${firm.vip ? `<span class="bbz-pill bbz-pill-vip">VIP</span>` : ""}
              </div>
            </div>

            ${this.actionBar([
              `<button class="bbz-button bbz-button-secondary" data-open-modal="firm-edit" data-item-id="${firm.id}">Firma bearbeiten</button>`,
              `<button class="bbz-button bbz-button-primary" data-open-modal="contact-create" data-prefill-firm-id="${firm.id}">Kontakt anlegen</button>`
            ])}
          </div>

          <div class="bbz-kpis">
            ${this.kpiBlock("Kontakte", firm.contactsCount)}
            ${this.kpiBlock("Offene Tasks", firm.openTasksCount)}
            ${this.kpiBlock("Nächste Deadline", helpers.formatDate(firm.nextDeadline) || "—")}
            ${this.kpiBlock("History", firm.history.length)}
          </div>

          <div class="bbz-grid bbz-grid-3">
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Übersicht</div></div>
              <div class="bbz-section-body">
                <div class="bbz-meta-grid">
                  ${ui.kv("Firma", helpers.escapeHtml(firm.title))}
                  ${ui.kv("Klassifizierung", firm.klassifizierung ? `<span class="${helpers.firmBadgeClass(firm.klassifizierung)}">${helpers.escapeHtml(firm.klassifizierung)}</span>` : '<span class="bbz-muted">—</span>')}
                  ${ui.kv("Adresse", helpers.escapeHtml(firm.adresse) || '<span class="bbz-muted">—</span>')}
                  ${ui.kv("PLZ / Ort", helpers.escapeHtml(helpers.joinNonEmpty([firm.plz, firm.ort], " ")) || '<span class="bbz-muted">—</span>')}
                  ${ui.kv("Land", helpers.escapeHtml(firm.land) || '<span class="bbz-muted">—</span>')}
                  ${ui.kv("Hauptnummer", helpers.escapeHtml(firm.hauptnummer) || '<span class="bbz-muted">—</span>')}
                  ${ui.kv("VIP", firm.vip ? '<span class="bbz-pill bbz-pill-vip">Ja</span>' : '<span class="bbz-muted">Nein</span>')}
                </div>
              </div>
            </section>

            <section class="bbz-section" style="grid-column: span 2;">
              <div class="bbz-section-header">
                <div>
                  <div class="bbz-section-title">Kontakte</div>
                  <div class="bbz-section-subtitle">Alle Ansprechpartner dieser Firma</div>
                </div>
              </div>
              <div class="bbz-section-body">
                <div class="bbz-table-wrap">
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
                              <td><a class="bbz-link" data-action="open-contact" data-id="${c.id}">${helpers.escapeHtml(c.fullName || c.nachname)}</a></td>
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
            </section>
          </div>

          <div class="bbz-grid bbz-grid-2 mt-4">
            <section class="bbz-section">
              <div class="bbz-section-header">
                <div>
                  <div class="bbz-section-title">Aktivitäten</div>
                  <div class="bbz-section-subtitle">Aggregierte History über alle Kontakte</div>
                </div>
                ${this.actionBar([
                  contacts.length ? `<button class="bbz-button bbz-button-primary" data-open-modal="history-create">History anlegen</button>` : ""
                ])}
              </div>
              <div class="bbz-section-body">
                ${
                  recentHistory.length
                    ? `<div class="bbz-timeline">
                        ${recentHistory.map(h => `
                          <div class="bbz-timeline-item">
                            <div class="bbz-timeline-date">${helpers.formatDateTime(h.datum) || "—"}<br><span class="bbz-muted">${helpers.escapeHtml(h.contactName || "")}</span></div>
                            <div>
                              <div class="bbz-timeline-title">${helpers.escapeHtml(h.typ || h.title || "Eintrag")} ${h.projektbezugBool ? '<span class="bbz-chip">Projektbezug</span>' : '<span class="bbz-chip">Allgemein</span>'}</div>
                              <div class="bbz-timeline-text">${helpers.escapeHtml(h.notizen || "—")}</div>
                            </div>
                          </div>
                        `).join("")}
                      </div>`
                    : ui.emptyBlock("Keine History-Einträge vorhanden.")
                }
              </div>
            </section>

            <section class="bbz-section">
              <div class="bbz-section-header">
                <div>
                  <div class="bbz-section-title">Aufgaben</div>
                  <div class="bbz-section-subtitle">Alle Tasks der Firma</div>
                </div>
                ${this.actionBar([
                  contacts.length ? `<button class="bbz-button bbz-button-primary" data-open-modal="task-create">Task anlegen</button>` : ""
                ])}
              </div>
              <div class="bbz-section-body">
                <div class="bbz-table-wrap">
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
                              <td>${t.contactId ? `<a class="bbz-link" data-action="open-contact" data-id="${t.contactId}">${helpers.escapeHtml(t.contactName || "Kontakt")}</a>` : helpers.escapeHtml(t.contactName || "—")}</td>
                            </tr>
                          `).join("")
                          : `<tr><td colspan="4">${ui.emptyBlock("Keine Aufgaben vorhanden.")}</td></tr>`
                      }
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

      const rows = state.enriched.contacts.filter(contact => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch =
          !search ||
          [contact.fullName, contact.firmTitle, contact.funktion, contact.rolle, contact.email1, contact.email2, contact.direktwahl, contact.mobile, contact.kommentar, ...contact.sgf, ...contact.event]
            .some(v => helpers.textIncludes(v, search));

        const archiveMatch = !filters.archiviertAusblenden || !contact.archiviert;
        return searchMatch && archiveMatch;
      });

      return `
        <section class="bbz-section">
          <div class="bbz-section-header">
            <div>
              <div class="bbz-section-title">Kontakte</div>
              <div class="bbz-section-subtitle">Operative Ansprechpartner über alle Firmen</div>
            </div>
            ${this.actionBar([
              `<button class="bbz-button bbz-button-primary" data-open-modal="contact-create">Kontakt anlegen</button>`
            ])}
          </div>

          <div class="bbz-section-body">
            <div class="bbz-filters-3">
              <input class="bbz-input" data-filter="contacts-search" type="text" placeholder="Suche nach Name, Firma, Funktion, Rolle, E-Mail, SGF, Event ..." value="${helpers.escapeHtml(filters.search)}" />
              <label class="bbz-checkbox">
                <input type="checkbox" data-filter="contacts-archiviert" ${filters.archiviertAusblenden ? "checked" : ""} />
                Archivierte ausblenden
              </label>
              <div></div>
            </div>

            <div class="bbz-table-wrap">
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
                          <td><a class="bbz-link" data-action="open-contact" data-id="${c.id}">${helpers.escapeHtml(c.fullName || c.nachname)}</a></td>
                          <td>${c.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${c.firmId}">${helpers.escapeHtml(c.firmTitle || "Firma")}</a>` : `<span class="bbz-muted">—</span>`}</td>
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
      `;
    },

    contactDetail() {
      const contact = dataModel.getContactById(state.selection.contactId);
      if (!contact) return ui.emptyBlock("Der ausgewählte Kontakt wurde nicht gefunden.");

      const contactHistory = state.enriched.history.filter(h => h.contactId === contact.id).sort((a, b) => helpers.compareDateDesc(a.datum, b.datum));
      const contactTasks = state.enriched.tasks.filter(t => t.contactId === contact.id).sort((a, b) => helpers.compareDateAsc(a.deadline, b.deadline));

      return `
        <div>
          <div class="bbz-detail-header">
            <div>
              <button class="bbz-button bbz-button-secondary mb-3" data-action="back-to-contacts">Zurück zur Kontaktliste</button>
              <div class="bbz-detail-title">${helpers.escapeHtml(contact.fullName || contact.nachname)}</div>
              <div class="bbz-detail-subtitle">
                ${contact.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${contact.firmId}">${helpers.escapeHtml(contact.firmTitle || "Firma")}</a>` : "Keine Firma verknüpft"}
                ${contact.funktion ? ` · ${helpers.escapeHtml(contact.funktion)}` : ""}
                ${contact.rolle ? ` · ${helpers.escapeHtml(contact.rolle)}` : ""}
              </div>
            </div>

            ${this.actionBar([
              contact.email1 ? `<a class="bbz-button bbz-button-secondary" href="mailto:${helpers.escapeHtml(contact.email1)}">Mail senden</a>` : "",
              `<button class="bbz-button bbz-button-secondary" data-open-modal="contact-edit" data-item-id="${contact.id}">Kontakt bearbeiten</button>`,
              `<button class="bbz-button bbz-button-primary" data-open-modal="task-create" data-prefill-contact-id="${contact.id}">Neue Aufgabe</button>`,
              `<button class="bbz-button bbz-button-primary" data-open-modal="history-create" data-prefill-contact-id="${contact.id}">Neue History</button>`
            ])}
          </div>

          <div class="bbz-kpis">
            ${this.kpiBlock("Tasks", contactTasks.length)}
            ${this.kpiBlock("Offene Tasks", contactTasks.filter(t => t.isOpen).length)}
            ${this.kpiBlock("History", contactHistory.length)}
            ${this.kpiBlock("Letzte Aktivität", helpers.formatDate(contactHistory[0]?.datum) || "—")}
          </div>

          <div class="bbz-grid bbz-grid-3">
            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Stammdaten</div></div>
              <div class="bbz-section-body">
                <div class="bbz-meta-grid">
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
            </section>

            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">CRM-Kontext</div></div>
              <div class="bbz-section-body">
                <div class="bbz-meta-grid">
                  ${ui.kv("Leadbbz0", helpers.escapeHtml(contact.leadbbz0) || '<span class="bbz-muted">—</span>')}
                  ${ui.kv("SGF", helpers.multiChoiceHtml(contact.sgf))}
                  ${ui.kv("Event", helpers.multiChoiceHtml(contact.event))}
                  ${ui.kv("Eventhistory", helpers.escapeHtml(contact.eventhistory) || '<span class="bbz-muted">—</span>')}
                  ${ui.kv("Kommentar", helpers.escapeHtml(contact.kommentar) || '<span class="bbz-muted">—</span>')}
                </div>
              </div>
            </section>

            <section class="bbz-section">
              <div class="bbz-section-header"><div class="bbz-section-title">Übersicht</div></div>
              <div class="bbz-section-body">
                <div class="bbz-meta-grid">
                  ${ui.kv("Tasks", String(contactTasks.length))}
                  ${ui.kv("Offene Tasks", String(contactTasks.filter(t => t.isOpen).length))}
                  ${ui.kv("History", String(contactHistory.length))}
                  ${ui.kv("Letzte Aktivität", helpers.formatDateTime(contactHistory[0]?.datum) || '<span class="bbz-muted">—</span>')}
                </div>
              </div>
            </section>
          </div>

          <div class="bbz-grid bbz-grid-2 mt-4">
            <section class="bbz-section">
              <div class="bbz-section-header">
                <div>
                  <div class="bbz-section-title">Historie</div>
                  <div class="bbz-section-subtitle">Timeline aus CRMHistory</div>
                </div>
              </div>
              <div class="bbz-section-body">
                ${
                  contactHistory.length
                    ? `<div class="bbz-timeline">
                        ${contactHistory.map(h => `
                          <div class="bbz-timeline-item">
                            <div class="bbz-timeline-date">${helpers.formatDateTime(h.datum) || "—"}</div>
                            <div>
                              <div class="bbz-timeline-title">${helpers.escapeHtml(h.typ || h.title || "Eintrag")} ${h.projektbezugBool ? '<span class="bbz-chip">Projektbezug</span>' : '<span class="bbz-chip">Allgemein</span>'}</div>
                              <div class="bbz-timeline-text">${helpers.escapeHtml(h.notizen || "—")}</div>
                            </div>
                          </div>
                        `).join("")}
                      </div>`
                    : ui.emptyBlock("Keine Historie vorhanden.")
                }
              </div>
            </section>

            <section class="bbz-section">
              <div class="bbz-section-header">
                <div>
                  <div class="bbz-section-title">Tasks</div>
                  <div class="bbz-section-subtitle">Aufgaben dieser Person</div>
                </div>
              </div>
              <div class="bbz-section-body">
                <div class="bbz-table-wrap">
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
                              <td>${t.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${t.firmId}">${helpers.escapeHtml(t.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
                            </tr>
                          `).join("")
                          : `<tr><td colspan="4">${ui.emptyBlock("Keine Tasks vorhanden.")}</td></tr>`
                      }
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

      const rows = state.enriched.tasks.filter(task => {
        const search = filters.search.trim().toLowerCase();
        const searchMatch =
          !search || [task.title, task.status, task.contactName, task.firmTitle, task.leadbbz].some(v => helpers.textIncludes(v, search));
        const openMatch = !filters.onlyOpen || task.isOpen;
        const overdueMatch = !filters.onlyOverdue || task.isOverdue;
        return searchMatch && openMatch && overdueMatch;
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
            ${this.kpiBlock("Überfällig", overdueTasks)}
            ${this.kpiBlock("Nächste 7 Tage", nextWeekTasks)}
          </div>

          <section class="bbz-section">
            <div class="bbz-section-header">
              <div>
                <div class="bbz-section-title">Planung</div>
                <div class="bbz-section-subtitle">Aufgabenübersicht mit Fokus auf offen und überfällig</div>
              </div>
              ${this.actionBar([
                `<button class="bbz-button bbz-button-primary" data-open-modal="task-create">Task anlegen</button>`
              ])}
            </div>

            <div class="bbz-section-body">
              <div class="bbz-filters-3">
                <input class="bbz-input" data-filter="planning-search" type="text" placeholder="Suche nach Titel, Firma, Kontakt, Status ..." value="${helpers.escapeHtml(filters.search)}" />
                <label class="bbz-checkbox">
                  <input type="checkbox" data-filter="planning-open" ${filters.onlyOpen ? "checked" : ""} />
                  Nur offene Tasks
                </label>
                <label class="bbz-checkbox">
                  <input type="checkbox" data-filter="planning-overdue" ${filters.onlyOverdue ? "checked" : ""} />
                  Nur überfällige Tasks
                </label>
              </div>

              <div class="bbz-table-wrap">
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
                            <td>${t.contactId ? `<a class="bbz-link" data-action="open-contact" data-id="${t.contactId}">${helpers.escapeHtml(t.contactName || "Kontakt")}</a>` : helpers.escapeHtml(t.contactName || "—")}</td>
                            <td>${t.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${t.firmId}">${helpers.escapeHtml(t.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
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
    },

    events() {
      const filters = state.filters.events;

      const groups = state.enriched.events
        .map(group => {
          const contacts = group.contacts.filter(item => {
            const search = filters.search.trim().toLowerCase();
            const searchMatch =
              !search ||
              [group.name, item.contactName, item.firmTitle, item.rolle, item.funktion, item.eventhistory, item.latestHistoryText]
                .some(v => helpers.textIncludes(v, search));
            const openMatch = !filters.onlyWithOpenTasks || item.openTasksCount > 0;
            return searchMatch && openMatch;
          });

          return { ...group, contacts };
        })
        .filter(group => group.contacts.length > 0);

      const totalEventGroups = state.enriched.events.length;
      const totalEventContacts = state.enriched.events.reduce((sum, e) => sum + e.contactCount, 0);
      const totalEventOpenTasks = state.enriched.events.reduce((sum, e) => sum + e.openTasksCount, 0);

      return `
        <div>
          <div class="bbz-kpis">
            ${this.kpiBlock("Event-Kategorien", totalEventGroups)}
            ${this.kpiBlock("Kontakt-Zuordnungen", totalEventContacts)}
            ${this.kpiBlock("Offene Tasks", totalEventOpenTasks)}
            ${this.kpiBlock("Sichtbare Kategorien", groups.length)}
          </div>

          <section class="bbz-section">
            <div class="bbz-section-header">
              <div>
                <div class="bbz-section-title">Events</div>
                <div class="bbz-section-subtitle">Separate Event-Sicht nach Kategorie mit Firmen- und Kontaktbezug</div>
              </div>
            </div>

            <div class="bbz-section-body">
              <div class="bbz-filters-3">
                <input class="bbz-input" data-filter="events-search" type="text" placeholder="Suche nach Kategorie, Kontakt, Firma, Rolle, Eventhistory ..." value="${helpers.escapeHtml(filters.search)}" />
                <label class="bbz-checkbox">
                  <input type="checkbox" data-filter="events-open" ${filters.onlyWithOpenTasks ? "checked" : ""} />
                  Nur mit offenen Tasks
                </label>
                <div></div>
              </div>

              ${
                groups.length
                  ? `<div class="bbz-cockpit-stack">
                      ${groups.map(group => `
                        <section class="bbz-section" style="box-shadow:none;">
                          <div class="bbz-section-header">
                            <div>
                              <div class="bbz-section-title">${helpers.escapeHtml(group.name)}</div>
                              <div class="bbz-section-subtitle">${group.contacts.length} Kontakte · ${group.contacts.reduce((sum, c) => sum + c.openTasksCount, 0)} offene Tasks</div>
                            </div>
                          </div>

                          <div class="bbz-section-body">
                            <div class="bbz-table-wrap">
                              <table class="bbz-table">
                                <thead>
                                  <tr>
                                    <th>Kontakt</th>
                                    <th>Firma</th>
                                    <th>Funktion / Rolle</th>
                                    <th>Eventhistory</th>
                                    <th>Letzte Aktivität</th>
                                    <th>Offene Tasks</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  ${group.contacts.map(item => `
                                    <tr>
                                      <td>
                                        <a class="bbz-link" data-action="open-contact" data-id="${item.contactId}">${helpers.escapeHtml(item.contactName)}</a>
                                        <div class="bbz-subtext">${item.email1 ? helpers.escapeHtml(item.email1) : "—"}</div>
                                      </td>
                                      <td>${item.firmId ? `<a class="bbz-link" data-action="open-firm" data-id="${item.firmId}">${helpers.escapeHtml(item.firmTitle || "Firma")}</a>` : '<span class="bbz-muted">—</span>'}</td>
                                      <td>${helpers.escapeHtml(helpers.joinNonEmpty([item.funktion, item.rolle], " · ")) || '<span class="bbz-muted">—</span>'}</td>
                                      <td>${helpers.escapeHtml(item.eventhistory) || '<span class="bbz-muted">—</span>'}</td>
                                      <td>${helpers.formatDateTime(item.latestHistoryDate) || '<span class="bbz-muted">—</span>'}${item.latestHistoryType ? `<div class="bbz-subtext">${helpers.escapeHtml(item.latestHistoryType)}</div>` : ""}</td>
                                      <td class="${item.openTasksCount > 0 ? 'bbz-warning' : 'bbz-muted'}">${item.openTasksCount}</td>
                                    </tr>
                                  `).join("")}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        </section>
                      `).join("")}
                    </div>`
                  : ui.emptyBlock("Keine Event-Daten für die aktuelle Filterung gefunden.")
              }
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
          ui.setMessage("Authentifizierung ist noch nicht bereit. Bitte Seite neu laden.", "warning");
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

    async reloadData() {
      try {
        ui.setLoading(true);
        await api.acquireToken();
        await api.loadAll();
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
      modal.render();
    }
  };

  function startApp() {
    controller.init();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startApp, { once: true });
  } else {
    startApp();
  }
})();

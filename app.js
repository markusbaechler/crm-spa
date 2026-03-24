(function () {
  'use strict';

  const CONFIG = {
    site: {
      hostname: window.location.hostname,
      sitePath: '/sites/CRM'
    },
    lists: {
      firms: 'CRMFirms',
      contacts: 'CRMContacts',
      history: 'CRMHistory',
      tasks: 'CRMTasks'
    },
    fields: {
      firms: {
        id: 'id',
        title: 'Title',
        abc: 'ABCSegment',
        email: 'Email',
        phone: 'Phone',
        website: 'Website',
        notes: 'Notes'
      },
      contacts: {
        id: 'id',
        title: 'Title',
        firstName: 'FirstName',
        lastName: 'LastName',
        email: 'Email',
        phone: 'Phone',
        mobile: 'Mobile',
        role: 'Role',
        firmLookupId: 'FirmLookupId',
        notes: 'Notes',
        leadDisplay: 'Leadbbz0'
      },
      history: {
        id: 'id',
        contactLookupId: 'ContactLookupId',
        date: 'Date',
        type: 'Type',
        notes: 'Notes',
        projectRelated: 'ProjectRelated'
      },
      tasks: {
        id: 'id',
        title: 'Title',
        contactLookupId: 'ContactLookupId',
        dueDate: 'DueDate',
        status: 'Status',
        notes: 'Notes'
      }
    },
    pageSize: 500,
    debug: true
  };

  const state = {
    token: null,
    siteId: null,
    listIds: {},
    firms: [],
    contacts: [],
    history: [],
    tasks: [],
    selectedFirmId: null,
    selectedContactId: null,
    filters: {
      firmSearch: '',
      contactSearch: ''
    },
    initialized: false
  };

  const dom = {};

  function log(...args) {
    if (CONFIG.debug) {
      console.log('[CRM]', ...args);
    }
  }

  function $(selector) {
    return document.querySelector(selector);
  }

  function byId(id) {
    return document.getElementById(id);
  }

  function qsAny(selectors) {
    for (const selector of selectors) {
      const el = selector.startsWith('#') ? byId(selector.slice(1)) : $(selector);
      if (el) return el;
    }
    return null;
  }

  function collectDom() {
    dom.appStatus = qsAny(['#appStatus', '#statusMessage']);
    dom.reloadButton = qsAny(['#btnReloadAll', '#reloadApp']);

    dom.firmList = qsAny(['#firmList', '#firmsList', '#companyList']);
    dom.firmSearch = qsAny(['#firmSearch', '#searchFirms', '#companySearch']);
    dom.firmCount = qsAny(['#firmCount', '#firmsCount']);
    dom.firmDetail = qsAny(['#firmDetail', '#firmDetails', '#companyDetail']);
    dom.firmDetailTitle = qsAny(['#firmDetailTitle', '#companyDetailTitle']);
    dom.firmDetailMeta = qsAny(['#firmDetailMeta', '#companyDetailMeta']);

    dom.contactList = qsAny(['#contactList', '#contactsList']);
    dom.contactSearch = qsAny(['#contactSearch', '#searchContacts']);
    dom.contactCount = qsAny(['#contactCount', '#contactsCount']);
    dom.contactDetail = qsAny(['#contactDetail', '#contactDetails']);
    dom.contactDetailTitle = qsAny(['#contactDetailTitle']);
    dom.contactDetailMeta = qsAny(['#contactDetailMeta']);

    dom.contactForm = qsAny(['#contactForm', '#formContact']);
    dom.contactFirmId = qsAny(['#contactFirmId', '#contactCompanyId']);
    dom.contactTitle = qsAny(['#contactTitle']);
    dom.contactFirstName = qsAny(['#contactFirstName']);
    dom.contactLastName = qsAny(['#contactLastName']);
    dom.contactEmail = qsAny(['#contactEmail']);
    dom.contactPhone = qsAny(['#contactPhone']);
    dom.contactMobile = qsAny(['#contactMobile']);
    dom.contactRole = qsAny(['#contactRole']);
    dom.contactNotes = qsAny(['#contactNotes']);
    dom.contactLeadDisplay = qsAny(['#contactLeadDisplay', '#contactLeadbbz0']);

    dom.taskForm = qsAny(['#taskForm', '#formTask']);
    dom.taskContactId = qsAny(['#taskContactId']);
    dom.taskTitle = qsAny(['#taskTitle']);
    dom.taskDueDate = qsAny(['#taskDueDate']);
    dom.taskStatus = qsAny(['#taskStatus']);
    dom.taskNotes = qsAny(['#taskNotes']);

    dom.historyForm = qsAny(['#historyForm', '#formHistory']);
    dom.historyContactId = qsAny(['#historyContactId']);
    dom.historyDate = qsAny(['#historyDate']);
    dom.historyType = qsAny(['#historyType']);
    dom.historyNotes = qsAny(['#historyNotes']);
    dom.historyProjectRelated = qsAny(['#historyProjectRelated']);
  }

  function setStatus(message, isError) {
    if (dom.appStatus) {
      dom.appStatus.textContent = message || '';
      dom.appStatus.classList.toggle('text-red-600', Boolean(isError));
      dom.appStatus.classList.toggle('text-slate-600', !isError);
    }
    if (message) {
      (isError ? console.error : console.log)('[CRM STATUS]', message);
    }
  }

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function formatDate(value) {
    if (!value) return '—';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    return new Intl.DateTimeFormat('de-CH', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    }).format(date);
  }

  function formatDateTimeLocalValue(value) {
    if (!value) return '';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return '';
    const pad = (n) => String(n).padStart(2, '0');
    return [
      date.getFullYear(),
      '-',
      pad(date.getMonth() + 1),
      '-',
      pad(date.getDate()),
      'T',
      pad(date.getHours()),
      ':',
      pad(date.getMinutes())
    ].join('');
  }

  function normalizeText(value) {
    return String(value || '').trim().toLowerCase();
  }

  function valueOrEmpty(value) {
    return value == null ? '' : String(value);
  }

  function boolFromField(value) {
    return value === true || value === 'true' || value === 1 || value === '1' || value === 'Ja';
  }

  function getField(item, internalName, fallback) {
    if (!item || !item.fields) return fallback;
    const value = item.fields[internalName];
    return value == null ? fallback : value;
  }

  function requireAccessTokenResult(token) {
    if (!token) {
      throw new Error('Kein Access Token verfügbar. Login ist zwar aktiv, aber app.js findet keinen Token-Provider.');
    }
    return token;
  }

  async function getAccessToken() {
    if (state.token) return state.token;

    if (window.authManager && typeof window.authManager.getAccessToken === 'function') {
      state.token = await window.authManager.getAccessToken();
      return requireAccessTokenResult(state.token);
    }

    if (typeof window.getAccessToken === 'function') {
      state.token = await window.getAccessToken();
      return requireAccessTokenResult(state.token);
    }

    if (window.CRM_AUTH && typeof window.CRM_AUTH.getAccessToken === 'function') {
      state.token = await window.CRM_AUTH.getAccessToken();
      return requireAccessTokenResult(state.token);
    }

    if (window.msalInstance && typeof window.msalInstance.acquireTokenSilent === 'function') {
      const accounts = window.msalInstance.getAllAccounts();
      if (accounts && accounts.length) {
        const response = await window.msalInstance.acquireTokenSilent({
          account: accounts[0],
          scopes: ['Sites.ReadWrite.All']
        });
        state.token = response && response.accessToken;
        return requireAccessTokenResult(state.token);
      }
    }

    throw new Error('Kein kompatibler Token-Provider gefunden.');
  }

  async function graphFetch(path, options) {
    const token = await getAccessToken();
    const headers = Object.assign(
      {
        Authorization: 'Bearer ' + token,
        Accept: 'application/json'
      },
      options && options.body ? { 'Content-Type': 'application/json' } : {},
      (options && options.headers) || {}
    );

    const response = await fetch('https://graph.microsoft.com/v1.0' + path, Object.assign({}, options, { headers }));
    const text = await response.text();
    const payload = text ? JSON.parse(text) : {};

    if (!response.ok) {
      throw new Error('Graph ' + response.status + ': ' + JSON.stringify(payload));
    }

    return payload;
  }

  async function ensureSiteAndLists() {
    if (!state.siteId) {
      const site = await graphFetch('/sites/' + CONFIG.site.hostname + ':' + CONFIG.site.sitePath);
      state.siteId = site.id;
      log('siteId', state.siteId);
    }

    const listNames = Object.values(CONFIG.lists);
    for (const listName of listNames) {
      if (!state.listIds[listName]) {
        const list = await graphFetch('/sites/' + state.siteId + '/lists/' + encodeURIComponent(listName));
        state.listIds[listName] = list.id;
        log('listId', listName, list.id);
      }
    }
  }

  async function getAllItems(listName) {
    await ensureSiteAndLists();
    const listId = state.listIds[listName];
    const items = [];
    let path = '/sites/' + state.siteId + '/lists/' + listId + '/items?expand=fields&top=' + CONFIG.pageSize;

    while (path) {
      const data = await graphFetch(path);
      items.push(...(data.value || []));
      const nextLink = data['@odata.nextLink'];
      if (nextLink) {
        path = nextLink.replace('https://graph.microsoft.com/v1.0', '');
      } else {
        path = null;
      }
    }

    return items;
  }

  async function createItem(listName, fields) {
    await ensureSiteAndLists();
    const listId = state.listIds[listName];
    return graphFetch('/sites/' + state.siteId + '/lists/' + listId + '/items', {
      method: 'POST',
      body: JSON.stringify({ fields })
    });
  }

  function sortByTitle(items, fieldName) {
    return items.slice().sort((a, b) => {
      const left = normalizeText(getField(a, fieldName, ''));
      const right = normalizeText(getField(b, fieldName, ''));
      return left.localeCompare(right, 'de');
    });
  }

  function getFirmName(firm) {
    const f = CONFIG.fields.firms;
    return getField(firm, f.title, 'Ohne Firmenname');
  }

  function getContactName(contact) {
    const f = CONFIG.fields.contacts;
    const firstName = valueOrEmpty(getField(contact, f.firstName, '')).trim();
    const lastName = valueOrEmpty(getField(contact, f.lastName, '')).trim();
    const title = valueOrEmpty(getField(contact, f.title, '')).trim();
    const combined = [firstName, lastName].filter(Boolean).join(' ').trim();
    return combined || title || 'Ohne Name';
  }

  function getFirmById(firmId) {
    return state.firms.find((firm) => String(firm.id) === String(firmId)) || null;
  }

  function getContactById(contactId) {
    return state.contacts.find((contact) => String(contact.id) === String(contactId)) || null;
  }

  function getContactsForFirm(firmId) {
    const field = CONFIG.fields.contacts.firmLookupId;
    return state.contacts.filter((contact) => String(getField(contact, field, '')) === String(firmId));
  }

  function getTasksForContact(contactId) {
    const field = CONFIG.fields.tasks.contactLookupId;
    return state.tasks.filter((task) => String(getField(task, field, '')) === String(contactId));
  }

  function getHistoryForContact(contactId) {
    const field = CONFIG.fields.history.contactLookupId;
    return state.history.filter((entry) => String(getField(entry, field, '')) === String(contactId));
  }

  function getTasksForFirm(firmId) {
    const contactIds = new Set(getContactsForFirm(firmId).map((contact) => String(contact.id)));
    return state.tasks.filter((task) => contactIds.has(String(getField(task, CONFIG.fields.tasks.contactLookupId, ''))));
  }

  function getHistoryForFirm(firmId) {
    const contactIds = new Set(getContactsForFirm(firmId).map((contact) => String(contact.id)));
    return state.history.filter((entry) => contactIds.has(String(getField(entry, CONFIG.fields.history.contactLookupId, ''))));
  }

  function matchesSearch(item, fields, search) {
    if (!search) return true;
    const q = normalizeText(search);
    return fields.some((field) => normalizeText(getField(item, field, '')).includes(q));
  }

  function filteredFirms() {
    const f = CONFIG.fields.firms;
    return sortByTitle(
      state.firms.filter((firm) => matchesSearch(firm, [f.title, f.abc, f.email, f.phone, f.website, f.notes], state.filters.firmSearch)),
      f.title
    );
  }

  function filteredContacts() {
    const f = CONFIG.fields.contacts;
    return state.contacts
      .filter((contact) => {
        const matches = matchesSearch(
          contact,
          [f.title, f.firstName, f.lastName, f.email, f.phone, f.mobile, f.role, f.notes, f.leadDisplay],
          state.filters.contactSearch
        );
        if (!matches) return false;
        if (!state.selectedFirmId) return true;
        return String(getField(contact, f.firmLookupId, '')) === String(state.selectedFirmId);
      })
      .sort((a, b) => getContactName(a).localeCompare(getContactName(b), 'de'));
  }

  function renderFirmList() {
    if (!dom.firmList) return;
    const firms = filteredFirms();
    if (dom.firmCount) dom.firmCount.textContent = String(firms.length);

    dom.firmList.innerHTML = firms.length
      ? firms
          .map((firm) => {
            const firmId = String(firm.id);
            const selected = String(state.selectedFirmId) === firmId;
            const abc = valueOrEmpty(getField(firm, CONFIG.fields.firms.abc, '—'));
            const contactCount = getContactsForFirm(firm.id).length;
            return (
              '<button type="button" class="crm-firm-row w-full text-left border rounded-lg px-3 py-2 mb-2 ' +
              (selected ? 'border-blue-500 bg-blue-50' : 'border-slate-200 bg-white') +
              '" data-firm-id="' +
              escapeHtml(firmId) +
              '">' +
              '<div class="font-semibold">' +
              escapeHtml(getFirmName(firm)) +
              '</div>' +
              '<div class="text-sm text-slate-600">ABC: ' +
              escapeHtml(abc) +
              ' · Kontakte: ' +
              contactCount +
              '</div>' +
              '</button>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine Firmen gefunden.</div>';
  }

  function renderContactList() {
    if (!dom.contactList) return;
    const contacts = filteredContacts();
    if (dom.contactCount) dom.contactCount.textContent = String(contacts.length);

    dom.contactList.innerHTML = contacts.length
      ? contacts
          .map((contact) => {
            const contactId = String(contact.id);
            const selected = String(state.selectedContactId) === contactId;
            const firm = getFirmById(getField(contact, CONFIG.fields.contacts.firmLookupId, ''));
            return (
              '<button type="button" class="crm-contact-row w-full text-left border rounded-lg px-3 py-2 mb-2 ' +
              (selected ? 'border-blue-500 bg-blue-50' : 'border-slate-200 bg-white') +
              '" data-contact-id="' +
              escapeHtml(contactId) +
              '">' +
              '<div class="font-semibold">' +
              escapeHtml(getContactName(contact)) +
              '</div>' +
              '<div class="text-sm text-slate-600">' +
              escapeHtml(getField(contact, CONFIG.fields.contacts.email, '')) +
              '</div>' +
              '<div class="text-xs text-slate-500">' +
              escapeHtml(firm ? getFirmName(firm) : 'Ohne Firma') +
              '</div>' +
              '</button>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine Kontakte gefunden.</div>';
  }

  function renderFirmDetail() {
    if (!dom.firmDetail) return;

    if (!state.selectedFirmId) {
      dom.firmDetail.innerHTML = '<div class="text-slate-500">Bitte eine Firma auswählen.</div>';
      return;
    }

    const firm = getFirmById(state.selectedFirmId);
    if (!firm) {
      dom.firmDetail.innerHTML = '<div class="text-red-600">Ausgewählte Firma nicht gefunden.</div>';
      return;
    }

    const contacts = getContactsForFirm(firm.id).sort((a, b) => getContactName(a).localeCompare(getContactName(b), 'de'));
    const tasks = getTasksForFirm(firm.id).sort((a, b) => {
      const left = new Date(getField(a, CONFIG.fields.tasks.dueDate, '2999-12-31')).getTime();
      const right = new Date(getField(b, CONFIG.fields.tasks.dueDate, '2999-12-31')).getTime();
      return left - right;
    });
    const history = getHistoryForFirm(firm.id).sort((a, b) => {
      const left = new Date(getField(a, CONFIG.fields.history.date, '1900-01-01')).getTime();
      const right = new Date(getField(b, CONFIG.fields.history.date, '1900-01-01')).getTime();
      return right - left;
    });

    if (dom.firmDetailTitle) dom.firmDetailTitle.textContent = getFirmName(firm);
    if (dom.firmDetailMeta) {
      dom.firmDetailMeta.textContent = 'ABC: ' + valueOrEmpty(getField(firm, CONFIG.fields.firms.abc, '—'));
    }

    const contactHtml = contacts.length
      ? contacts
          .map((contact) => {
            return (
              '<button type="button" class="crm-open-contact block w-full text-left border rounded-md px-3 py-2 mb-2 border-slate-200" data-contact-id="' +
              escapeHtml(String(contact.id)) +
              '">' +
              '<div class="font-medium">' +
              escapeHtml(getContactName(contact)) +
              '</div>' +
              '<div class="text-sm text-slate-600">' +
              escapeHtml(getField(contact, CONFIG.fields.contacts.email, '')) +
              '</div>' +
              '</button>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine Kontakte zugeordnet.</div>';

    const taskHtml = tasks.length
      ? tasks
          .map((task) => {
            const contact = getContactById(getField(task, CONFIG.fields.tasks.contactLookupId, ''));
            return (
              '<div class="border rounded-md px-3 py-2 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(getField(task, CONFIG.fields.tasks.title, 'Ohne Titel')) +
              '</div>' +
              '<div class="text-sm text-slate-600">Fällig: ' +
              escapeHtml(formatDate(getField(task, CONFIG.fields.tasks.dueDate, ''))) +
              ' · Status: ' +
              escapeHtml(getField(task, CONFIG.fields.tasks.status, '—')) +
              '</div>' +
              '<div class="text-xs text-slate-500">Kontakt: ' +
              escapeHtml(contact ? getContactName(contact) : '—') +
              '</div>' +
              '</div>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine Tasks vorhanden.</div>';

    const historyHtml = history.length
      ? history
          .map((entry) => {
            const contact = getContactById(getField(entry, CONFIG.fields.history.contactLookupId, ''));
            return (
              '<div class="border rounded-md px-3 py-2 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(formatDate(getField(entry, CONFIG.fields.history.date, ''))) +
              ' · ' +
              escapeHtml(getField(entry, CONFIG.fields.history.type, '')) +
              '</div>' +
              '<div class="text-sm text-slate-700 whitespace-pre-wrap">' +
              escapeHtml(getField(entry, CONFIG.fields.history.notes, '')) +
              '</div>' +
              '<div class="text-xs text-slate-500">Kontakt: ' +
              escapeHtml(contact ? getContactName(contact) : '—') +
              ' · Projektbezug: ' +
              (boolFromField(getField(entry, CONFIG.fields.history.projectRelated, false)) ? 'Ja' : 'Nein') +
              '</div>' +
              '</div>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine History vorhanden.</div>';

    dom.firmDetail.innerHTML =
      '<div class="space-y-6">' +
      '<section><h3 class="font-semibold mb-2">Kontakte</h3>' +
      contactHtml +
      '</section>' +
      '<section><h3 class="font-semibold mb-2">Tasks</h3>' +
      taskHtml +
      '</section>' +
      '<section><h3 class="font-semibold mb-2">History</h3>' +
      historyHtml +
      '</section>' +
      '</div>';
  }

  function renderContactDetail() {
    if (!dom.contactDetail) return;

    if (!state.selectedContactId) {
      dom.contactDetail.innerHTML = '<div class="text-slate-500">Bitte einen Kontakt auswählen.</div>';
      if (dom.contactLeadDisplay) dom.contactLeadDisplay.value = '';
      return;
    }

    const contact = getContactById(state.selectedContactId);
    if (!contact) {
      dom.contactDetail.innerHTML = '<div class="text-red-600">Ausgewählter Kontakt nicht gefunden.</div>';
      return;
    }

    const firm = getFirmById(getField(contact, CONFIG.fields.contacts.firmLookupId, ''));
    const tasks = getTasksForContact(contact.id).sort((a, b) => {
      const left = new Date(getField(a, CONFIG.fields.tasks.dueDate, '2999-12-31')).getTime();
      const right = new Date(getField(b, CONFIG.fields.tasks.dueDate, '2999-12-31')).getTime();
      return left - right;
    });
    const history = getHistoryForContact(contact.id).sort((a, b) => {
      const left = new Date(getField(a, CONFIG.fields.history.date, '1900-01-01')).getTime();
      const right = new Date(getField(b, CONFIG.fields.history.date, '1900-01-01')).getTime();
      return right - left;
    });

    if (dom.contactDetailTitle) dom.contactDetailTitle.textContent = getContactName(contact);
    if (dom.contactDetailMeta) {
      dom.contactDetailMeta.textContent = [
        firm ? getFirmName(firm) : 'Ohne Firma',
        getField(contact, CONFIG.fields.contacts.email, ''),
        getField(contact, CONFIG.fields.contacts.phone, '')
      ]
        .filter(Boolean)
        .join(' · ');
    }

    if (dom.contactLeadDisplay) {
      if ('value' in dom.contactLeadDisplay) {
        dom.contactLeadDisplay.value = valueOrEmpty(getField(contact, CONFIG.fields.contacts.leadDisplay, ''));
      } else {
        dom.contactLeadDisplay.textContent = valueOrEmpty(getField(contact, CONFIG.fields.contacts.leadDisplay, ''));
      }
    }

    const taskHtml = tasks.length
      ? tasks
          .map((task) => {
            return (
              '<div class="border rounded-md px-3 py-2 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(getField(task, CONFIG.fields.tasks.title, 'Ohne Titel')) +
              '</div>' +
              '<div class="text-sm text-slate-600">Fällig: ' +
              escapeHtml(formatDate(getField(task, CONFIG.fields.tasks.dueDate, ''))) +
              ' · Status: ' +
              escapeHtml(getField(task, CONFIG.fields.tasks.status, '—')) +
              '</div>' +
              '<div class="text-sm whitespace-pre-wrap">' +
              escapeHtml(getField(task, CONFIG.fields.tasks.notes, '')) +
              '</div>' +
              '</div>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine Tasks vorhanden.</div>';

    const historyHtml = history.length
      ? history
          .map((entry) => {
            return (
              '<div class="border rounded-md px-3 py-2 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(formatDate(getField(entry, CONFIG.fields.history.date, ''))) +
              ' · ' +
              escapeHtml(getField(entry, CONFIG.fields.history.type, '')) +
              '</div>' +
              '<div class="text-sm whitespace-pre-wrap">' +
              escapeHtml(getField(entry, CONFIG.fields.history.notes, '')) +
              '</div>' +
              '<div class="text-xs text-slate-500">Projektbezug: ' +
              (boolFromField(getField(entry, CONFIG.fields.history.projectRelated, false)) ? 'Ja' : 'Nein') +
              '</div>' +
              '</div>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine History vorhanden.</div>';

    dom.contactDetail.innerHTML =
      '<div class="space-y-6">' +
      '<section><h3 class="font-semibold mb-2">Kontakt</h3>' +
      '<div class="border rounded-md px-3 py-2 border-slate-200">' +
      '<div><strong>Name:</strong> ' +
      escapeHtml(getContactName(contact)) +
      '</div>' +
      '<div><strong>Rolle:</strong> ' +
      escapeHtml(getField(contact, CONFIG.fields.contacts.role, '—')) +
      '</div>' +
      '<div><strong>E-Mail:</strong> ' +
      escapeHtml(getField(contact, CONFIG.fields.contacts.email, '—')) +
      '</div>' +
      '<div><strong>Telefon:</strong> ' +
      escapeHtml(getField(contact, CONFIG.fields.contacts.phone, '—')) +
      '</div>' +
      '<div><strong>Mobile:</strong> ' +
      escapeHtml(getField(contact, CONFIG.fields.contacts.mobile, '—')) +
      '</div>' +
      '</div></section>' +
      '<section><h3 class="font-semibold mb-2">Tasks</h3>' +
      taskHtml +
      '</section>' +
      '<section><h3 class="font-semibold mb-2">History</h3>' +
      historyHtml +
      '</section>' +
      '</div>';
  }

  function fillFirmSelectOptions() {
    if (!dom.contactFirmId) return;
    const current = dom.contactFirmId.value;
    const options = ['<option value="">Firma wählen</option>']
      .concat(
        sortByTitle(state.firms, CONFIG.fields.firms.title).map((firm) => {
          return '<option value="' + escapeHtml(String(firm.id)) + '">' + escapeHtml(getFirmName(firm)) + '</option>';
        })
      )
      .join('');
    dom.contactFirmId.innerHTML = options;
    if (current) dom.contactFirmId.value = current;
  }

  function fillContactSelectOptions() {
    const contactOptions = ['<option value="">Kontakt wählen</option>']
      .concat(
        state.contacts
          .slice()
          .sort((a, b) => getContactName(a).localeCompare(getContactName(b), 'de'))
          .map((contact) => {
            const firm = getFirmById(getField(contact, CONFIG.fields.contacts.firmLookupId, ''));
            const label = getContactName(contact) + (firm ? ' (' + getFirmName(firm) + ')' : '');
            return '<option value="' + escapeHtml(String(contact.id)) + '">' + escapeHtml(label) + '</option>';
          })
      )
      .join('');

    if (dom.taskContactId) {
      const currentTaskContactId = dom.taskContactId.value;
      dom.taskContactId.innerHTML = contactOptions;
      if (currentTaskContactId) dom.taskContactId.value = currentTaskContactId;
    }

    if (dom.historyContactId) {
      const currentHistoryContactId = dom.historyContactId.value;
      dom.historyContactId.innerHTML = contactOptions;
      if (currentHistoryContactId) dom.historyContactId.value = currentHistoryContactId;
    }
  }

  function syncFormsWithSelection() {
    if (dom.contactFirmId && state.selectedFirmId) {
      dom.contactFirmId.value = String(state.selectedFirmId);
    }
    if (dom.taskContactId && state.selectedContactId) {
      dom.taskContactId.value = String(state.selectedContactId);
    }
    if (dom.historyContactId && state.selectedContactId) {
      dom.historyContactId.value = String(state.selectedContactId);
    }
  }

  function renderAll() {
    renderFirmList();
    renderContactList();
    renderFirmDetail();
    renderContactDetail();
    fillFirmSelectOptions();
    fillContactSelectOptions();
    syncFormsWithSelection();
  }

  async function loadAllData() {
    setStatus('Lade CRM-Daten ...', false);
    await ensureSiteAndLists();

    const [firms, contacts, history, tasks] = await Promise.all([
      getAllItems(CONFIG.lists.firms),
      getAllItems(CONFIG.lists.contacts),
      getAllItems(CONFIG.lists.history),
      getAllItems(CONFIG.lists.tasks)
    ]);

    state.firms = firms;
    state.contacts = contacts;
    state.history = history;
    state.tasks = tasks;

    if (!state.selectedFirmId && state.firms.length) {
      state.selectedFirmId = state.firms[0].id;
    }

    if (state.selectedFirmId) {
      const contactsForFirm = getContactsForFirm(state.selectedFirmId);
      if (!state.selectedContactId && contactsForFirm.length) {
        state.selectedContactId = contactsForFirm[0].id;
      }
    }

    renderAll();
    setStatus('CRM-Daten geladen.', false);
  }

  function resetContactForm() {
    if (!dom.contactForm) return;
    dom.contactForm.reset();
    if (dom.contactFirmId && state.selectedFirmId) dom.contactFirmId.value = String(state.selectedFirmId);
  }

  function resetTaskForm() {
    if (!dom.taskForm) return;
    dom.taskForm.reset();
    if (dom.taskContactId && state.selectedContactId) dom.taskContactId.value = String(state.selectedContactId);
  }

  function resetHistoryForm() {
    if (!dom.historyForm) return;
    dom.historyForm.reset();
    if (dom.historyContactId && state.selectedContactId) dom.historyContactId.value = String(state.selectedContactId);
    if (dom.historyDate) dom.historyDate.value = formatDateTimeLocalValue(new Date().toISOString());
  }

  async function onCreateContact(event) {
    event.preventDefault();

    const fields = CONFIG.fields.contacts;
    const firmId = valueOrEmpty(dom.contactFirmId && dom.contactFirmId.value).trim();
    const firstName = valueOrEmpty(dom.contactFirstName && dom.contactFirstName.value).trim();
    const lastName = valueOrEmpty(dom.contactLastName && dom.contactLastName.value).trim();
    const title = valueOrEmpty(dom.contactTitle && dom.contactTitle.value).trim() || [firstName, lastName].filter(Boolean).join(' ').trim();

    if (!firmId) throw new Error('Kontakt kann nicht gespeichert werden: Firma fehlt.');
    if (!title) throw new Error('Kontakt kann nicht gespeichert werden: Name fehlt.');

    const payload = {};
    payload[fields.title] = title;
    payload[fields.firstName] = firstName;
    payload[fields.lastName] = lastName;
    payload[fields.email] = valueOrEmpty(dom.contactEmail && dom.contactEmail.value).trim();
    payload[fields.phone] = valueOrEmpty(dom.contactPhone && dom.contactPhone.value).trim();
    payload[fields.mobile] = valueOrEmpty(dom.contactMobile && dom.contactMobile.value).trim();
    payload[fields.role] = valueOrEmpty(dom.contactRole && dom.contactRole.value).trim();
    payload[fields.notes] = valueOrEmpty(dom.contactNotes && dom.contactNotes.value).trim();
    payload[fields.firmLookupId] = Number(firmId);

    await createItem(CONFIG.lists.contacts, payload);
    await loadAllData();
    resetContactForm();
  }

  async function onCreateTask(event) {
    event.preventDefault();

    const fields = CONFIG.fields.tasks;
    const contactId = valueOrEmpty(dom.taskContactId && dom.taskContactId.value).trim();
    const title = valueOrEmpty(dom.taskTitle && dom.taskTitle.value).trim();
    const dueDate = valueOrEmpty(dom.taskDueDate && dom.taskDueDate.value).trim();

    if (!contactId) throw new Error('Task kann nicht gespeichert werden: Kontakt fehlt.');
    if (!title) throw new Error('Task kann nicht gespeichert werden: Titel fehlt.');

    const payload = {};
    payload[fields.title] = title;
    payload[fields.contactLookupId] = Number(contactId);
    if (dueDate) payload[fields.dueDate] = new Date(dueDate).toISOString();
    payload[fields.status] = valueOrEmpty(dom.taskStatus && dom.taskStatus.value).trim();
    payload[fields.notes] = valueOrEmpty(dom.taskNotes && dom.taskNotes.value).trim();

    await createItem(CONFIG.lists.tasks, payload);
    await loadAllData();
    resetTaskForm();
  }

  async function onCreateHistory(event) {
    event.preventDefault();

    const fields = CONFIG.fields.history;
    const contactId = valueOrEmpty(dom.historyContactId && dom.historyContactId.value).trim();
    const dateValue = valueOrEmpty(dom.historyDate && dom.historyDate.value).trim();
    const type = valueOrEmpty(dom.historyType && dom.historyType.value).trim();
    const notes = valueOrEmpty(dom.historyNotes && dom.historyNotes.value).trim();

    if (!contactId) throw new Error('History kann nicht gespeichert werden: Kontakt fehlt.');
    if (!dateValue) throw new Error('History kann nicht gespeichert werden: Datum fehlt.');
    if (!type) throw new Error('History kann nicht gespeichert werden: Typ fehlt.');
    if (!notes) throw new Error('History kann nicht gespeichert werden: Notizen fehlen.');

    const payload = {};
    payload[fields.contactLookupId] = Number(contactId);
    payload[fields.date] = new Date(dateValue).toISOString();
    payload[fields.type] = type;
    payload[fields.notes] = notes;
    payload[fields.projectRelated] = Boolean(dom.historyProjectRelated && dom.historyProjectRelated.checked);

    await createItem(CONFIG.lists.history, payload);
    await loadAllData();
    resetHistoryForm();
  }

  function selectFirm(firmId) {
    state.selectedFirmId = String(firmId);
    const firmContacts = getContactsForFirm(state.selectedFirmId);
    if (!firmContacts.some((contact) => String(contact.id) === String(state.selectedContactId))) {
      state.selectedContactId = firmContacts.length ? String(firmContacts[0].id) : null;
    }
    renderAll();
  }

  function selectContact(contactId) {
    const contact = getContactById(contactId);
    state.selectedContactId = String(contactId);
    if (contact) {
      state.selectedFirmId = String(getField(contact, CONFIG.fields.contacts.firmLookupId, state.selectedFirmId || ''));
    }
    renderAll();
  }

  function bindEvents() {
    if (dom.reloadButton) {
      dom.reloadButton.addEventListener('click', () => {
        loadAllData().catch(handleError);
      });
    }

    if (dom.firmSearch) {
      dom.firmSearch.addEventListener('input', (event) => {
        state.filters.firmSearch = event.target.value || '';
        renderFirmList();
      });
    }

    if (dom.contactSearch) {
      dom.contactSearch.addEventListener('input', (event) => {
        state.filters.contactSearch = event.target.value || '';
        renderContactList();
      });
    }

    if (dom.firmList) {
      dom.firmList.addEventListener('click', (event) => {
        const button = event.target.closest('[data-firm-id]');
        if (!button) return;
        selectFirm(button.getAttribute('data-firm-id'));
      });
    }

    if (dom.contactList) {
      dom.contactList.addEventListener('click', (event) => {
        const button = event.target.closest('[data-contact-id]');
        if (!button) return;
        selectContact(button.getAttribute('data-contact-id'));
      });
    }

    if (dom.firmDetail) {
      dom.firmDetail.addEventListener('click', (event) => {
        const button = event.target.closest('[data-contact-id]');
        if (!button) return;
        selectContact(button.getAttribute('data-contact-id'));
      });
    }

    if (dom.contactForm) {
      dom.contactForm.addEventListener('submit', (event) => {
        onCreateContact(event).catch(handleError);
      });
    }

    if (dom.taskForm) {
      dom.taskForm.addEventListener('submit', (event) => {
        onCreateTask(event).catch(handleError);
      });
    }

    if (dom.historyForm) {
      dom.historyForm.addEventListener('submit', (event) => {
        onCreateHistory(event).catch(handleError);
      });
    }
  }

  function validateConfig() {
    const missing = [];
    Object.entries(CONFIG.fields).forEach(([entity, map]) => {
      Object.entries(map).forEach(([key, value]) => {
        if (!value) missing.push(entity + '.' + key);
      });
    });
    if (missing.length) {
      throw new Error('Konfiguration unvollständig: ' + missing.join(', '));
    }
  }

  function handleError(error) {
    console.error(error);
    const message = error && error.message ? error.message : String(error);
    setStatus(message, true);
  }

  async function init() {
    if (state.initialized) return;
    state.initialized = true;
    collectDom();
    validateConfig();
    bindEvents();
    resetHistoryForm();
    await loadAllData();
  }

  document.addEventListener('DOMContentLoaded', function () {
    init().catch(handleError);
  });

  window.CRM_APP = {
    init,
    reload: loadAllData,
    state,
    config: CONFIG,
    selectFirm,
    selectContact
  };
})();

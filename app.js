(function () {
  'use strict';

  const CONFIG = {
    sharePoint: {
      hostname: 'bbzsg.sharepoint.com',
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
        title: 'Title',
        abc: 'ABCSegment',
        email: 'Email',
        phone: 'Phone',
        website: 'Website',
        notes: 'Notes'
      },
      contacts: {
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
        contactLookupId: 'ContactLookupId',
        date: 'Date',
        type: 'Type',
        notes: 'Notes',
        projectRelated: 'ProjectRelated'
      },
      tasks: {
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
    initialized: false,
    loading: false,
    authToken: null,
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
    }
  };

  const dom = {};

  function log() {
    if (!CONFIG.debug) return;
    const args = Array.prototype.slice.call(arguments);
    args.unshift('[CRM]');
    console.log.apply(console, args);
  }

  function $(selector) {
    return document.querySelector(selector);
  }

  function byId(id) {
    return document.getElementById(id);
  }

  function qsAny(selectors) {
    for (let i = 0; i < selectors.length; i += 1) {
      const selector = selectors[i];
      const el = selector.charAt(0) === '#' ? byId(selector.slice(1)) : $(selector);
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
    }

    if (!message) return;

    if (isError) {
      console.error('[CRM STATUS]', message);
    } else {
      console.log('[CRM STATUS]', message);
    }
  }

  function handleError(error) {
    console.error(error);
    const message = error && error.message ? error.message : String(error);
    setStatus(message, true);
  }

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function normalizeText(value) {
    return String(value || '').trim().toLowerCase();
  }

  function valueOrEmpty(value) {
    return value == null ? '' : String(value);
  }

  function boolFromField(value) {
    return value === true || value === 1 || value === '1' || value === 'true' || value === 'Ja';
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
    const pad = function (n) {
      return String(n).padStart(2, '0');
    };
    return (
      date.getFullYear() +
      '-' +
      pad(date.getMonth() + 1) +
      '-' +
      pad(date.getDate()) +
      'T' +
      pad(date.getHours()) +
      ':' +
      pad(date.getMinutes())
    );
  }

  function maybeJwt(value) {
    return typeof value === 'string' && value.split('.').length === 3 && value.length > 100;
  }

  function setAccessToken(token) {
    if (!maybeJwt(token)) {
      throw new Error('Ungültiger Access Token übergeben.');
    }
    state.authToken = token;
    window.CRM_APP_AUTH_TOKEN = token;
    log('Access token übernommen');
    setStatus('Login erkannt. CRM kann Daten laden.', false);
  }

  async function resolveAccessToken() {
    if (state.authToken && maybeJwt(state.authToken)) {
      return state.authToken;
    }

    if (window.CRM_APP_AUTH_TOKEN && maybeJwt(window.CRM_APP_AUTH_TOKEN)) {
      state.authToken = window.CRM_APP_AUTH_TOKEN;
      return state.authToken;
    }

    if (window.CRM_AUTH && typeof window.CRM_AUTH.getAccessToken === 'function') {
      const token = await window.CRM_AUTH.getAccessToken();
      if (maybeJwt(token)) {
        state.authToken = token;
        return state.authToken;
      }
    }

    throw new Error('Nicht angemeldet. Bitte zuerst anmelden.');
  }

  async function graphFetch(path, options) {
    const token = await resolveAccessToken();

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
      if (response.status === 401 || response.status === 403) {
        state.authToken = null;
      }
      throw new Error('Graph ' + response.status + ': ' + JSON.stringify(payload));
    }

    return payload;
  }

  async function ensureSite() {
    if (state.siteId) return state.siteId;

    const site = await graphFetch('/sites/' + CONFIG.sharePoint.hostname + ':' + CONFIG.sharePoint.sitePath);

    if (!site || !site.id) {
      throw new Error('Site-ID konnte nicht geladen werden.');
    }

    state.siteId = site.id;
    return state.siteId;
  }

  async function ensureListId(listTitle) {
    await ensureSite();

    if (state.listIds[listTitle]) {
      return state.listIds[listTitle];
    }

    const data = await graphFetch(
      '/sites/' +
        state.siteId +
        '/lists?$filter=displayName eq ' +
        "'" +
        String(listTitle).replace(/'/g, "''") +
        "'"
    );

    const list = data && data.value && data.value.length ? data.value[0] : null;

    if (!list || !list.id) {
      throw new Error('Liste nicht gefunden: ' + listTitle);
    }

    state.listIds[listTitle] = list.id;
    return state.listIds[listTitle];
  }

  async function getAllItems(listTitle) {
    await ensureSite();
    const listId = await ensureListId(listTitle);

    const items = [];
    let path = '/sites/' + state.siteId + '/lists/' + listId + '/items?expand=fields&top=' + CONFIG.pageSize;

    while (path) {
      const data = await graphFetch(path);
      items.push.apply(items, data.value || []);
      const next = data['@odata.nextLink'];
      path = next ? next.replace('https://graph.microsoft.com/v1.0', '') : null;
    }

    return items;
  }

  async function createItem(listTitle, fields) {
    await ensureSite();
    const listId = await ensureListId(listTitle);

    return graphFetch('/sites/' + state.siteId + '/lists/' + listId + '/items', {
      method: 'POST',
      body: JSON.stringify({ fields: fields })
    });
  }

  function getField(item, fieldName, fallback) {
    if (!item || !item.fields) return fallback;
    const value = item.fields[fieldName];
    return value == null ? fallback : value;
  }

  function getItemId(item) {
    if (!item) return null;
    return item.id != null ? item.id : null;
  }

  function sortByField(items, fieldName) {
    return items.slice().sort(function (a, b) {
      const left = normalizeText(getField(a, fieldName, ''));
      const right = normalizeText(getField(b, fieldName, ''));
      return left.localeCompare(right, 'de');
    });
  }

  function matchesSearch(item, fieldNames, search) {
    if (!search) return true;
    const q = normalizeText(search);

    for (let i = 0; i < fieldNames.length; i += 1) {
      const value = normalizeText(getField(item, fieldNames[i], ''));
      if (value.indexOf(q) !== -1) {
        return true;
      }
    }

    return false;
  }

  function getFirmName(firm) {
    return valueOrEmpty(getField(firm, CONFIG.fields.firms.title, '')).trim() || 'Ohne Firmenname';
  }

  function getContactName(contact) {
    if (!contact) return 'Ohne Name';

    const f = CONFIG.fields.contacts;
    const firstName = valueOrEmpty(getField(contact, f.firstName, '')).trim();
    const lastName = valueOrEmpty(getField(contact, f.lastName, '')).trim();
    const title = valueOrEmpty(getField(contact, f.title, '')).trim();
    const combined = [firstName, lastName].filter(Boolean).join(' ').trim();

    return combined || title || 'Ohne Name';
  }

  function getFirmById(firmId) {
    return state.firms.find(function (firm) {
      return String(getItemId(firm)) === String(firmId);
    }) || null;
  }

  function getContactById(contactId) {
    return state.contacts.find(function (contact) {
      return String(getItemId(contact)) === String(contactId);
    }) || null;
  }

  function getContactsForFirm(firmId) {
    const field = CONFIG.fields.contacts.firmLookupId;
    return state.contacts.filter(function (contact) {
      return String(getField(contact, field, '')) === String(firmId);
    });
  }

  function getTasksForContact(contactId) {
    const field = CONFIG.fields.tasks.contactLookupId;
    return state.tasks.filter(function (task) {
      return String(getField(task, field, '')) === String(contactId);
    });
  }

  function getHistoryForContact(contactId) {
    const field = CONFIG.fields.history.contactLookupId;
    return state.history.filter(function (entry) {
      return String(getField(entry, field, '')) === String(contactId);
    });
  }

  function getTasksForFirm(firmId) {
    const contactIds = new Set(
      getContactsForFirm(firmId).map(function (contact) {
        return String(getItemId(contact));
      })
    );

    return state.tasks.filter(function (task) {
      return contactIds.has(String(getField(task, CONFIG.fields.tasks.contactLookupId, '')));
    });
  }

  function getHistoryForFirm(firmId) {
    const contactIds = new Set(
      getContactsForFirm(firmId).map(function (contact) {
        return String(getItemId(contact));
      })
    );

    return state.history.filter(function (entry) {
      return contactIds.has(String(getField(entry, CONFIG.fields.history.contactLookupId, '')));
    });
  }

  function filteredFirms() {
    const f = CONFIG.fields.firms;

    return sortByField(
      state.firms.filter(function (firm) {
        return matchesSearch(
          firm,
          [f.title, f.abc, f.email, f.phone, f.website, f.notes],
          state.filters.firmSearch
        );
      }),
      f.title
    );
  }

  function filteredContacts() {
    const f = CONFIG.fields.contacts;

    return state.contacts
      .filter(function (contact) {
        const match = matchesSearch(
          contact,
          [f.title, f.firstName, f.lastName, f.email, f.phone, f.mobile, f.role, f.notes, f.leadDisplay],
          state.filters.contactSearch
        );

        if (!match) return false;
        if (!state.selectedFirmId) return true;

        return String(getField(contact, f.firmLookupId, '')) === String(state.selectedFirmId);
      })
      .sort(function (a, b) {
        return getContactName(a).localeCompare(getContactName(b), 'de');
      });
  }

  function renderFirmList() {
    if (!dom.firmList) return;

    const firms = filteredFirms();

    if (dom.firmCount) {
      dom.firmCount.textContent = String(firms.length);
    }

    if (!firms.length) {
      dom.firmList.innerHTML = '<div class="text-slate-500">Keine Firmen gefunden.</div>';
      return;
    }

    dom.firmList.innerHTML = firms
      .map(function (firm) {
        const id = String(getItemId(firm));
        const selected = String(state.selectedFirmId) === id;
        const abc = valueOrEmpty(getField(firm, CONFIG.fields.firms.abc, '—'));
        const contactCount = getContactsForFirm(id).length;

        return (
          '<button type="button" class="w-full text-left border rounded-2xl px-3 py-3 mb-2 ' +
          (selected ? 'border-blue-500 bg-blue-50' : 'border-slate-200 bg-white') +
          '" data-firm-id="' +
          escapeHtml(id) +
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
      .join('');
  }

  function renderContactList() {
    if (!dom.contactList) return;

    const contacts = filteredContacts();

    if (dom.contactCount) {
      dom.contactCount.textContent = String(contacts.length);
    }

    if (!contacts.length) {
      dom.contactList.innerHTML = '<div class="text-slate-500">Keine Kontakte gefunden.</div>';
      return;
    }

    dom.contactList.innerHTML = contacts
      .map(function (contact) {
        const id = String(getItemId(contact));
        const selected = String(state.selectedContactId) === id;
        const firm = getFirmById(getField(contact, CONFIG.fields.contacts.firmLookupId, ''));

        return (
          '<button type="button" class="w-full text-left border rounded-2xl px-3 py-3 mb-2 ' +
          (selected ? 'border-blue-500 bg-blue-50' : 'border-slate-200 bg-white') +
          '" data-contact-id="' +
          escapeHtml(id) +
          '">' +
          '<div class="font-semibold">' +
          escapeHtml(getContactName(contact)) +
          '</div>' +
          '<div class="text-sm text-slate-600">' +
          escapeHtml(valueOrEmpty(getField(contact, CONFIG.fields.contacts.email, ''))) +
          '</div>' +
          '<div class="text-xs text-slate-500">' +
          escapeHtml(firm ? getFirmName(firm) : 'Ohne Firma') +
          '</div>' +
          '</button>'
        );
      })
      .join('');
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

    if (dom.firmDetailTitle) {
      dom.firmDetailTitle.textContent = getFirmName(firm);
    }

    if (dom.firmDetailMeta) {
      dom.firmDetailMeta.textContent = 'ABC: ' + (valueOrEmpty(getField(firm, CONFIG.fields.firms.abc, '')) || '—');
    }

    const contacts = getContactsForFirm(getItemId(firm)).sort(function (a, b) {
      return getContactName(a).localeCompare(getContactName(b), 'de');
    });

    const tasks = getTasksForFirm(getItemId(firm)).sort(function (a, b) {
      const left = new Date(getField(a, CONFIG.fields.tasks.dueDate, '2999-12-31')).getTime();
      const right = new Date(getField(b, CONFIG.fields.tasks.dueDate, '2999-12-31')).getTime();
      return left - right;
    });

    const history = getHistoryForFirm(getItemId(firm)).sort(function (a, b) {
      const left = new Date(getField(a, CONFIG.fields.history.date, '1900-01-01')).getTime();
      const right = new Date(getField(b, CONFIG.fields.history.date, '1900-01-01')).getTime();
      return right - left;
    });

    const contactsHtml = contacts.length
      ? contacts
          .map(function (contact) {
            return (
              '<button type="button" class="block w-full text-left border rounded-xl px-3 py-2 mb-2 border-slate-200" data-contact-id="' +
              escapeHtml(String(getItemId(contact))) +
              '">' +
              '<div class="font-medium">' +
              escapeHtml(getContactName(contact)) +
              '</div>' +
              '<div class="text-sm text-slate-600">' +
              escapeHtml(valueOrEmpty(getField(contact, CONFIG.fields.contacts.email, ''))) +
              '</div>' +
              '</button>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine Kontakte zugeordnet.</div>';

    const tasksHtml = tasks.length
      ? tasks
          .map(function (task) {
            const contact = getContactById(getField(task, CONFIG.fields.tasks.contactLookupId, ''));

            return (
              '<div class="border rounded-xl px-3 py-3 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(valueOrEmpty(getField(task, CONFIG.fields.tasks.title, 'Ohne Titel'))) +
              '</div>' +
              '<div class="text-sm text-slate-600">Fällig: ' +
              escapeHtml(formatDate(getField(task, CONFIG.fields.tasks.dueDate, ''))) +
              ' · Status: ' +
              escapeHtml(valueOrEmpty(getField(task, CONFIG.fields.tasks.status, '—'))) +
              '</div>' +
              '<div class="text-xs text-slate-500">Kontakt: ' +
              escapeHtml(contact ? getContactName(contact) : '—') +
              '</div>' +
              '<div class="text-sm whitespace-pre-wrap mt-1">' +
              escapeHtml(valueOrEmpty(getField(task, CONFIG.fields.tasks.notes, ''))) +
              '</div>' +
              '</div>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine Tasks vorhanden.</div>';

    const historyHtml = history.length
      ? history
          .map(function (entry) {
            const contact = getContactById(getField(entry, CONFIG.fields.history.contactLookupId, ''));

            return (
              '<div class="border rounded-xl px-3 py-3 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(formatDate(getField(entry, CONFIG.fields.history.date, ''))) +
              ' · ' +
              escapeHtml(valueOrEmpty(getField(entry, CONFIG.fields.history.type, ''))) +
              '</div>' +
              '<div class="text-sm whitespace-pre-wrap mt-1">' +
              escapeHtml(valueOrEmpty(getField(entry, CONFIG.fields.history.notes, ''))) +
              '</div>' +
              '<div class="text-xs text-slate-500 mt-1">Kontakt: ' +
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
      contactsHtml +
      '</section>' +
      '<section><h3 class="font-semibold mb-2">Tasks</h3>' +
      tasksHtml +
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
      if (dom.contactLeadDisplay) {
        dom.contactLeadDisplay.value = '';
      }
      return;
    }

    const contact = getContactById(state.selectedContactId);

    if (!contact) {
      dom.contactDetail.innerHTML = '<div class="text-red-600">Ausgewählter Kontakt nicht gefunden.</div>';
      return;
    }

    const firm = getFirmById(getField(contact, CONFIG.fields.contacts.firmLookupId, ''));

    const tasks = getTasksForContact(getItemId(contact)).sort(function (a, b) {
      const left = new Date(getField(a, CONFIG.fields.tasks.dueDate, '2999-12-31')).getTime();
      const right = new Date(getField(b, CONFIG.fields.tasks.dueDate, '2999-12-31')).getTime();
      return left - right;
    });

    const history = getHistoryForContact(getItemId(contact)).sort(function (a, b) {
      const left = new Date(getField(a, CONFIG.fields.history.date, '1900-01-01')).getTime();
      const right = new Date(getField(b, CONFIG.fields.history.date, '1900-01-01')).getTime();
      return right - left;
    });

    if (dom.contactDetailTitle) {
      dom.contactDetailTitle.textContent = getContactName(contact);
    }

    if (dom.contactDetailMeta) {
      dom.contactDetailMeta.textContent = [
        firm ? getFirmName(firm) : 'Ohne Firma',
        valueOrEmpty(getField(contact, CONFIG.fields.contacts.email, '')),
        valueOrEmpty(getField(contact, CONFIG.fields.contacts.phone, ''))
      ]
        .filter(Boolean)
        .join(' · ');
    }

    if (dom.contactLeadDisplay) {
      dom.contactLeadDisplay.value = valueOrEmpty(getField(contact, CONFIG.fields.contacts.leadDisplay, ''));
    }

    const tasksHtml = tasks.length
      ? tasks
          .map(function (task) {
            return (
              '<div class="border rounded-xl px-3 py-3 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(valueOrEmpty(getField(task, CONFIG.fields.tasks.title, 'Ohne Titel'))) +
              '</div>' +
              '<div class="text-sm text-slate-600">Fällig: ' +
              escapeHtml(formatDate(getField(task, CONFIG.fields.tasks.dueDate, ''))) +
              ' · Status: ' +
              escapeHtml(valueOrEmpty(getField(task, CONFIG.fields.tasks.status, '—'))) +
              '</div>' +
              '<div class="text-sm whitespace-pre-wrap mt-1">' +
              escapeHtml(valueOrEmpty(getField(task, CONFIG.fields.tasks.notes, ''))) +
              '</div>' +
              '</div>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine Tasks vorhanden.</div>';

    const historyHtml = history.length
      ? history
          .map(function (entry) {
            return (
              '<div class="border rounded-xl px-3 py-3 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(formatDate(getField(entry, CONFIG.fields.history.date, ''))) +
              ' · ' +
              escapeHtml(valueOrEmpty(getField(entry, CONFIG.fields.history.type, ''))) +
              '</div>' +
              '<div class="text-sm whitespace-pre-wrap mt-1">' +
              escapeHtml(valueOrEmpty(getField(entry, CONFIG.fields.history.notes, ''))) +
              '</div>' +
              '<div class="text-xs text-slate-500 mt-1">Projektbezug: ' +
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
      '<div class="border rounded-xl px-3 py-3 border-slate-200">' +
      '<div><strong>Name:</strong> ' +
      escapeHtml(getContactName(contact)) +
      '</div>' +
      '<div><strong>Firma:</strong> ' +
      escapeHtml(firm ? getFirmName(firm) : '—') +
      '</div>' +
      '<div><strong>Rolle:</strong> ' +
      escapeHtml(valueOrEmpty(getField(contact, CONFIG.fields.contacts.role, '—'))) +
      '</div>' +
      '<div><strong>E-Mail:</strong> ' +
      escapeHtml(valueOrEmpty(getField(contact, CONFIG.fields.contacts.email, '—'))) +
      '</div>' +
      '<div><strong>Telefon:</strong> ' +
      escapeHtml(valueOrEmpty(getField(contact, CONFIG.fields.contacts.phone, '—'))) +
      '</div>' +
      '<div><strong>Mobile:</strong> ' +
      escapeHtml(valueOrEmpty(getField(contact, CONFIG.fields.contacts.mobile, '—'))) +
      '</div>' +
      '<div><strong>Leadbbz0:</strong> ' +
      escapeHtml(valueOrEmpty(getField(contact, CONFIG.fields.contacts.leadDisplay, '—'))) +
      '</div>' +
      '</div></section>' +
      '<section><h3 class="font-semibold mb-2">Tasks</h3>' +
      tasksHtml +
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
        sortByField(state.firms, CONFIG.fields.firms.title).map(function (firm) {
          return (
            '<option value="' +
            escapeHtml(String(getItemId(firm))) +
            '">' +
            escapeHtml(getFirmName(firm)) +
            '</option>'
          );
        })
      )
      .join('');

    dom.contactFirmId.innerHTML = options;

    if (current) {
      dom.contactFirmId.value = current;
    } else if (state.selectedFirmId) {
      dom.contactFirmId.value = String(state.selectedFirmId);
    }
  }

  function fillContactSelectOptions() {
    const options = ['<option value="">Kontakt wählen</option>']
      .concat(
        state.contacts
          .slice()
          .sort(function (a, b) {
            return getContactName(a).localeCompare(getContactName(b), 'de');
          })
          .map(function (contact) {
            const firm = getFirmById(getField(contact, CONFIG.fields.contacts.firmLookupId, ''));
            const label = getContactName(contact) + (firm ? ' (' + getFirmName(firm) + ')' : '');

            return (
              '<option value="' +
              escapeHtml(String(getItemId(contact))) +
              '">' +
              escapeHtml(label) +
              '</option>'
            );
          })
      )
      .join('');

    if (dom.taskContactId) {
      const currentTask = dom.taskContactId.value;
      dom.taskContactId.innerHTML = options;
      if (currentTask) {
        dom.taskContactId.value = currentTask;
      } else if (state.selectedContactId) {
        dom.taskContactId.value = String(state.selectedContactId);
      }
    }

    if (dom.historyContactId) {
      const currentHistory = dom.historyContactId.value;
      dom.historyContactId.innerHTML = options;
      if (currentHistory) {
        dom.historyContactId.value = currentHistory;
      } else if (state.selectedContactId) {
        dom.historyContactId.value = String(state.selectedContactId);
      }
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

  function resetContactForm() {
    if (!dom.contactForm) return;
    dom.contactForm.reset();

    if (dom.contactFirmId && state.selectedFirmId) {
      dom.contactFirmId.value = String(state.selectedFirmId);
    }
  }

  function resetTaskForm() {
    if (!dom.taskForm) return;
    dom.taskForm.reset();

    if (dom.taskContactId && state.selectedContactId) {
      dom.taskContactId.value = String(state.selectedContactId);
    }
  }

  function resetHistoryForm() {
    if (!dom.historyForm) return;
    dom.historyForm.reset();

    if (dom.historyContactId && state.selectedContactId) {
      dom.historyContactId.value = String(state.selectedContactId);
    }

    if (dom.historyDate) {
      dom.historyDate.value = formatDateTimeLocalValue(new Date().toISOString());
    }
  }

  async function loadAllData() {
    if (state.loading) return;

    state.loading = true;
    setStatus('Lade CRM-Daten ...', false);

    try {
      await resolveAccessToken();

      const results = await Promise.all([
        getAllItems(CONFIG.lists.firms),
        getAllItems(CONFIG.lists.contacts),
        getAllItems(CONFIG.lists.history),
        getAllItems(CONFIG.lists.tasks)
      ]);

      state.firms = results[0] || [];
      state.contacts = results[1] || [];
      state.history = results[2] || [];
      state.tasks = results[3] || [];

      if (!state.selectedFirmId && state.firms.length) {
        state.selectedFirmId = String(getItemId(state.firms[0]));
      }

      if (state.selectedFirmId) {
        const firmExists = !!getFirmById(state.selectedFirmId);
        if (!firmExists) {
          state.selectedFirmId = state.firms.length ? String(getItemId(state.firms[0])) : null;
        }
      }

      if (state.selectedFirmId) {
        const contactsForFirm = getContactsForFirm(state.selectedFirmId);
        const currentStillValid = contactsForFirm.some(function (contact) {
          return String(getItemId(contact)) === String(state.selectedContactId);
        });

        if (!currentStillValid) {
          state.selectedContactId = contactsForFirm.length ? String(getItemId(contactsForFirm[0])) : null;
        }
      } else {
        state.selectedContactId = null;
      }

      renderAll();
      setStatus('CRM-Daten geladen.', false);
    } finally {
      state.loading = false;
    }
  }

  function selectFirm(firmId) {
    state.selectedFirmId = String(firmId);

    const contacts = getContactsForFirm(state.selectedFirmId);
    const stillValid = contacts.some(function (contact) {
      return String(getItemId(contact)) === String(state.selectedContactId);
    });

    if (!stillValid) {
      state.selectedContactId = contacts.length ? String(getItemId(contacts[0])) : null;
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

  async function onCreateContact(event) {
    event.preventDefault();

    const f = CONFIG.fields.contacts;
    const firmId = valueOrEmpty(dom.contactFirmId && dom.contactFirmId.value).trim();
    const titleInput = valueOrEmpty(dom.contactTitle && dom.contactTitle.value).trim();
    const firstName = valueOrEmpty(dom.contactFirstName && dom.contactFirstName.value).trim();
    const lastName = valueOrEmpty(dom.contactLastName && dom.contactLastName.value).trim();
    const title = titleInput || [firstName, lastName].filter(Boolean).join(' ').trim();

    if (!firmId) {
      throw new Error('Kontakt kann nicht gespeichert werden: Firma fehlt.');
    }

    if (!title) {
      throw new Error('Kontakt kann nicht gespeichert werden: Name fehlt.');
    }

    const payload = {};
    payload[f.title] = title;
    payload[f.firstName] = firstName;
    payload[f.lastName] = lastName;
    payload[f.email] = valueOrEmpty(dom.contactEmail && dom.contactEmail.value).trim();
    payload[f.phone] = valueOrEmpty(dom.contactPhone && dom.contactPhone.value).trim();
    payload[f.mobile] = valueOrEmpty(dom.contactMobile && dom.contactMobile.value).trim();
    payload[f.role] = valueOrEmpty(dom.contactRole && dom.contactRole.value).trim();
    payload[f.notes] = valueOrEmpty(dom.contactNotes && dom.contactNotes.value).trim();
    payload[f.firmLookupId] = Number(firmId);

    await createItem(CONFIG.lists.contacts, payload);
    await loadAllData();
    resetContactForm();
  }

  async function onCreateTask(event) {
    event.preventDefault();

    const f = CONFIG.fields.tasks;
    const contactId = valueOrEmpty(dom.taskContactId && dom.taskContactId.value).trim();
    const title = valueOrEmpty(dom.taskTitle && dom.taskTitle.value).trim();
    const dueDate = valueOrEmpty(dom.taskDueDate && dom.taskDueDate.value).trim();

    if (!contactId) {
      throw new Error('Task kann nicht gespeichert werden: Kontakt fehlt.');
    }

    if (!title) {
      throw new Error('Task kann nicht gespeichert werden: Titel fehlt.');
    }

    const payload = {};
    payload[f.title] = title;
    payload[f.contactLookupId] = Number(contactId);
    payload[f.status] = valueOrEmpty(dom.taskStatus && dom.taskStatus.value).trim();
    payload[f.notes] = valueOrEmpty(dom.taskNotes && dom.taskNotes.value).trim();

    if (dueDate) {
      payload[f.dueDate] = new Date(dueDate).toISOString();
    }

    await createItem(CONFIG.lists.tasks, payload);
    await loadAllData();
    resetTaskForm();
  }

  async function onCreateHistory(event) {
    event.preventDefault();

    const f = CONFIG.fields.history;
    const contactId = valueOrEmpty(dom.historyContactId && dom.historyContactId.value).trim();
    const dateValue = valueOrEmpty(dom.historyDate && dom.historyDate.value).trim();
    const type = valueOrEmpty(dom.historyType && dom.historyType.value).trim();
    const notes = valueOrEmpty(dom.historyNotes && dom.historyNotes.value).trim();

    if (!contactId) {
      throw new Error('History kann nicht gespeichert werden: Kontakt fehlt.');
    }

    if (!dateValue) {
      throw new Error('History kann nicht gespeichert werden: Datum fehlt.');
    }

    if (!type) {
      throw new Error('History kann nicht gespeichert werden: Typ fehlt.');
    }

    if (!notes) {
      throw new Error('History kann nicht gespeichert werden: Notizen fehlen.');
    }

    const payload = {};
    payload[f.contactLookupId] = Number(contactId);
    payload[f.date] = new Date(dateValue).toISOString();
    payload[f.type] = type;
    payload[f.notes] = notes;
    payload[f.projectRelated] = !!(dom.historyProjectRelated && dom.historyProjectRelated.checked);

    await createItem(CONFIG.lists.history, payload);
    await loadAllData();
    resetHistoryForm();
  }

  function bindEvents() {
    if (dom.reloadButton) {
      dom.reloadButton.addEventListener('click', function () {
        loadAllData().catch(handleError);
      });
    }

    if (dom.firmSearch) {
      dom.firmSearch.addEventListener('input', function (event) {
        state.filters.firmSearch = event.target.value || '';
        renderFirmList();
      });
    }

    if (dom.contactSearch) {
      dom.contactSearch.addEventListener('input', function (event) {
        state.filters.contactSearch = event.target.value || '';
        renderContactList();
      });
    }

    if (dom.firmList) {
      dom.firmList.addEventListener('click', function (event) {
        const button = event.target.closest('[data-firm-id]');
        if (!button) return;
        selectFirm(button.getAttribute('data-firm-id'));
      });
    }

    if (dom.contactList) {
      dom.contactList.addEventListener('click', function (event) {
        const button = event.target.closest('[data-contact-id]');
        if (!button) return;
        selectContact(button.getAttribute('data-contact-id'));
      });
    }

    if (dom.firmDetail) {
      dom.firmDetail.addEventListener('click', function (event) {
        const button = event.target.closest('[data-contact-id]');
        if (!button) return;
        selectContact(button.getAttribute('data-contact-id'));
      });
    }

    if (dom.contactForm) {
      dom.contactForm.addEventListener('submit', function (event) {
        onCreateContact(event).catch(handleError);
      });
    }

    if (dom.taskForm) {
      dom.taskForm.addEventListener('submit', function (event) {
        onCreateTask(event).catch(handleError);
      });
    }

    if (dom.historyForm) {
      dom.historyForm.addEventListener('submit', function (event) {
        onCreateHistory(event).catch(handleError);
      });
    }
  }

  function validateConfig() {
    const missing = [];

    Object.keys(CONFIG.fields).forEach(function (entity) {
      Object.keys(CONFIG.fields[entity]).forEach(function (key) {
        if (!CONFIG.fields[entity][key]) {
          missing.push(entity + '.' + key);
        }
      });
    });

    if (missing.length) {
      throw new Error('Konfiguration unvollständig: ' + missing.join(', '));
    }
  }

  async function init() {
    if (state.initialized) return;
    state.initialized = true;

    collectDom();
    validateConfig();
    bindEvents();
    resetHistoryForm();

    try {
      await resolveAccessToken();
      setStatus('Login erkannt. Neu laden lädt CRM-Daten.', false);
    } catch (error) {
      setStatus('Bitte zuerst anmelden und danach Neu laden.', false);
      log('Initialer Token nicht gefunden:', error.message);
    }
  }

  document.addEventListener('DOMContentLoaded', function () {
    init().catch(handleError);
  });

  window.CRM_APP = {
    init: init,
    reload: loadAllData,
    setAccessToken: setAccessToken,
    selectFirm: selectFirm,
    selectContact: selectContact,
    state: state,
    config: CONFIG
  };
})();

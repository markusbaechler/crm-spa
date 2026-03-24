(function () {
  'use strict';

  const CONFIG = {
    lists: {
      firms: 'CRMFirms',
      contacts: 'CRMContacts',
      history: 'CRMHistory',
      tasks: 'CRMTasks'
    },
    fields: {
      firms: {
        id: 'Id',
        title: 'Title',
        abc: 'ABCSegment',
        email: 'Email',
        phone: 'Phone',
        website: 'Website',
        notes: 'Notes'
      },
      contacts: {
        id: 'Id',
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
        id: 'Id',
        contactLookupId: 'ContactLookupId',
        date: 'Date',
        type: 'Type',
        notes: 'Notes',
        projectRelated: 'ProjectRelated'
      },
      tasks: {
        id: 'Id',
        title: 'Title',
        contactLookupId: 'ContactLookupId',
        dueDate: 'DueDate',
        status: 'Status',
        notes: 'Notes'
      }
    },
    debug: true,
    pageSize: 500
  };

  const state = {
    initialized: false,
    requestDigest: null,
    listEntityTypes: {},
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
    if (CONFIG.debug) {
      console.log.apply(console, ['[CRM]'].concat(Array.from(arguments)));
    }
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
      if (dom.appStatus.classList) {
        dom.appStatus.classList.toggle('text-red-600', !!isError);
        dom.appStatus.classList.toggle('text-slate-600', !isError);
      }
    }
    if (message) {
      if (isError) {
        console.error('[CRM STATUS]', message);
      } else {
        console.log('[CRM STATUS]', message);
      }
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

  function sharePointBaseUrl() {
    if (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl) {
      return window._spPageContextInfo.webAbsoluteUrl;
    }
    return window.location.origin;
  }

  async function spFetch(url, options) {
    const response = await fetch(url, Object.assign({ credentials: 'same-origin' }, options || {}));
    const text = await response.text();
    let data = null;

    if (text) {
      try {
        data = JSON.parse(text);
      } catch (e) {
        data = text;
      }
    }

    if (!response.ok) {
      throw new Error('REST ' + response.status + ': ' + (typeof data === 'string' ? data : JSON.stringify(data)));
    }

    return data;
  }

  async function ensureDigest() {
    if (state.requestDigest) return state.requestDigest;

    const url = sharePointBaseUrl() + '/_api/contextinfo';
    const data = await spFetch(url, {
      method: 'POST',
      headers: {
        Accept: 'application/json;odata=nometadata'
      }
    });

    if (!data || !data.FormDigestValue) {
      throw new Error('FormDigest konnte nicht geladen werden.');
    }

    state.requestDigest = data.FormDigestValue;
    return state.requestDigest;
  }

  async function ensureListEntityType(listTitle) {
    if (state.listEntityTypes[listTitle]) return state.listEntityTypes[listTitle];

    const url =
      sharePointBaseUrl() +
      "/_api/web/lists/getbytitle('" +
      encodeURIComponent(listTitle).replace(/'/g, '%27') +
      "')?$select=ListItemEntityTypeFullName";

    const data = await spFetch(url, {
      headers: {
        Accept: 'application/json;odata=nometadata'
      }
    });

    if (!data || !data.ListItemEntityTypeFullName) {
      throw new Error('ListItemEntityTypeFullName für Liste ' + listTitle + ' konnte nicht geladen werden.');
    }

    state.listEntityTypes[listTitle] = data.ListItemEntityTypeFullName;
    return state.listEntityTypes[listTitle];
  }

  async function getAllItems(listTitle) {
    let url =
      sharePointBaseUrl() +
      "/_api/web/lists/getbytitle('" +
      encodeURIComponent(listTitle).replace(/'/g, '%27') +
      "')/items?$top=" +
      CONFIG.pageSize;

    const all = [];

    while (url) {
      const data = await spFetch(url, {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      });

      const values = data && data.value ? data.value : [];
      all.push.apply(all, values);

      url = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : null;
    }

    return all;
  }

  async function createItem(listTitle, payload) {
    const digest = await ensureDigest();
    const entityType = await ensureListEntityType(listTitle);

    const body = Object.assign(
      {
        __metadata: { type: entityType }
      },
      payload
    );

    const url =
      sharePointBaseUrl() +
      "/_api/web/lists/getbytitle('" +
      encodeURIComponent(listTitle).replace(/'/g, '%27') +
      "')/items";

    return spFetch(url, {
      method: 'POST',
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest
      },
      body: JSON.stringify(body)
    });
  }

  function sortByField(items, fieldName) {
    return items.slice().sort(function (a, b) {
      return normalizeText(a[fieldName]).localeCompare(normalizeText(b[fieldName]), 'de');
    });
  }

  function matchesSearch(item, fieldNames, search) {
    if (!search) return true;
    const q = normalizeText(search);
    for (let i = 0; i < fieldNames.length; i += 1) {
      const value = normalizeText(item[fieldNames[i]]);
      if (value.indexOf(q) !== -1) return true;
    }
    return false;
  }

  function getFirmName(firm) {
    return valueOrEmpty(firm && firm[CONFIG.fields.firms.title]).trim() || 'Ohne Firmenname';
  }

  function getContactName(contact) {
    if (!contact) return 'Ohne Name';
    const f = CONFIG.fields.contacts;
    const firstName = valueOrEmpty(contact[f.firstName]).trim();
    const lastName = valueOrEmpty(contact[f.lastName]).trim();
    const title = valueOrEmpty(contact[f.title]).trim();
    const combined = [firstName, lastName].filter(Boolean).join(' ').trim();
    return combined || title || 'Ohne Name';
  }

  function getFirmById(firmId) {
    return state.firms.find(function (firm) {
      return String(firm.Id) === String(firmId);
    }) || null;
  }

  function getContactById(contactId) {
    return state.contacts.find(function (contact) {
      return String(contact.Id) === String(contactId);
    }) || null;
  }

  function getContactsForFirm(firmId) {
    const field = CONFIG.fields.contacts.firmLookupId;
    return state.contacts.filter(function (contact) {
      return String(contact[field]) === String(firmId);
    });
  }

  function getTasksForContact(contactId) {
    const field = CONFIG.fields.tasks.contactLookupId;
    return state.tasks.filter(function (task) {
      return String(task[field]) === String(contactId);
    });
  }

  function getHistoryForContact(contactId) {
    const field = CONFIG.fields.history.contactLookupId;
    return state.history.filter(function (entry) {
      return String(entry[field]) === String(contactId);
    });
  }

  function getTasksForFirm(firmId) {
    const contactIds = new Set(
      getContactsForFirm(firmId).map(function (contact) {
        return String(contact.Id);
      })
    );
    return state.tasks.filter(function (task) {
      return contactIds.has(String(task[CONFIG.fields.tasks.contactLookupId]));
    });
  }

  function getHistoryForFirm(firmId) {
    const contactIds = new Set(
      getContactsForFirm(firmId).map(function (contact) {
        return String(contact.Id);
      })
    );
    return state.history.filter(function (entry) {
      return contactIds.has(String(entry[CONFIG.fields.history.contactLookupId]));
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
        return String(contact[f.firmLookupId]) === String(state.selectedFirmId);
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
        const selected = String(state.selectedFirmId) === String(firm.Id);
        const abc = valueOrEmpty(firm[CONFIG.fields.firms.abc]) || '—';
        const contactCount = getContactsForFirm(firm.Id).length;

        return (
          '<button type="button" class="crm-firm-row w-full text-left border rounded-lg px-3 py-2 mb-2 ' +
          (selected ? 'border-blue-500 bg-blue-50' : 'border-slate-200 bg-white') +
          '" data-firm-id="' +
          escapeHtml(String(firm.Id)) +
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
        const selected = String(state.selectedContactId) === String(contact.Id);
        const firm = getFirmById(contact[CONFIG.fields.contacts.firmLookupId]);

        return (
          '<button type="button" class="crm-contact-row w-full text-left border rounded-lg px-3 py-2 mb-2 ' +
          (selected ? 'border-blue-500 bg-blue-50' : 'border-slate-200 bg-white') +
          '" data-contact-id="' +
          escapeHtml(String(contact.Id)) +
          '">' +
          '<div class="font-semibold">' +
          escapeHtml(getContactName(contact)) +
          '</div>' +
          '<div class="text-sm text-slate-600">' +
          escapeHtml(valueOrEmpty(contact[CONFIG.fields.contacts.email])) +
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
      dom.firmDetailMeta.textContent = 'ABC: ' + (valueOrEmpty(firm[CONFIG.fields.firms.abc]) || '—');
    }

    const contacts = getContactsForFirm(firm.Id).sort(function (a, b) {
      return getContactName(a).localeCompare(getContactName(b), 'de');
    });

    const tasks = getTasksForFirm(firm.Id).sort(function (a, b) {
      const left = new Date(a[CONFIG.fields.tasks.dueDate] || '2999-12-31').getTime();
      const right = new Date(b[CONFIG.fields.tasks.dueDate] || '2999-12-31').getTime();
      return left - right;
    });

    const history = getHistoryForFirm(firm.Id).sort(function (a, b) {
      const left = new Date(a[CONFIG.fields.history.date] || '1900-01-01').getTime();
      const right = new Date(b[CONFIG.fields.history.date] || '1900-01-01').getTime();
      return right - left;
    });

    const contactsHtml = contacts.length
      ? contacts
          .map(function (contact) {
            return (
              '<button type="button" class="crm-open-contact block w-full text-left border rounded-md px-3 py-2 mb-2 border-slate-200" data-contact-id="' +
              escapeHtml(String(contact.Id)) +
              '">' +
              '<div class="font-medium">' +
              escapeHtml(getContactName(contact)) +
              '</div>' +
              '<div class="text-sm text-slate-600">' +
              escapeHtml(valueOrEmpty(contact[CONFIG.fields.contacts.email])) +
              '</div>' +
              '</button>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine Kontakte zugeordnet.</div>';

    const tasksHtml = tasks.length
      ? tasks
          .map(function (task) {
            const contact = getContactById(task[CONFIG.fields.tasks.contactLookupId]);
            return (
              '<div class="border rounded-md px-3 py-2 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(valueOrEmpty(task[CONFIG.fields.tasks.title]) || 'Ohne Titel') +
              '</div>' +
              '<div class="text-sm text-slate-600">Fällig: ' +
              escapeHtml(formatDate(task[CONFIG.fields.tasks.dueDate])) +
              ' · Status: ' +
              escapeHtml(valueOrEmpty(task[CONFIG.fields.tasks.status]) || '—') +
              '</div>' +
              '<div class="text-xs text-slate-500">Kontakt: ' +
              escapeHtml(contact ? getContactName(contact) : '—') +
              '</div>' +
              '<div class="text-sm whitespace-pre-wrap">' +
              escapeHtml(valueOrEmpty(task[CONFIG.fields.tasks.notes])) +
              '</div>' +
              '</div>'
            );
          })
          .join('')
      : '<div class="text-slate-500">Keine Tasks vorhanden.</div>';

    const historyHtml = history.length
      ? history
          .map(function (entry) {
            const contact = getContactById(entry[CONFIG.fields.history.contactLookupId]);
            return (
              '<div class="border rounded-md px-3 py-2 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(formatDate(entry[CONFIG.fields.history.date])) +
              ' · ' +
              escapeHtml(valueOrEmpty(entry[CONFIG.fields.history.type])) +
              '</div>' +
              '<div class="text-sm text-slate-700 whitespace-pre-wrap">' +
              escapeHtml(valueOrEmpty(entry[CONFIG.fields.history.notes])) +
              '</div>' +
              '<div class="text-xs text-slate-500">Kontakt: ' +
              escapeHtml(contact ? getContactName(contact) : '—') +
              ' · Projektbezug: ' +
              (boolFromField(entry[CONFIG.fields.history.projectRelated]) ? 'Ja' : 'Nein') +
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
        if ('value' in dom.contactLeadDisplay) {
          dom.contactLeadDisplay.value = '';
        } else {
          dom.contactLeadDisplay.textContent = '';
        }
      }
      return;
    }

    const contact = getContactById(state.selectedContactId);
    if (!contact) {
      dom.contactDetail.innerHTML = '<div class="text-red-600">Ausgewählter Kontakt nicht gefunden.</div>';
      return;
    }

    const firm = getFirmById(contact[CONFIG.fields.contacts.firmLookupId]);
    const tasks = getTasksForContact(contact.Id).sort(function (a, b) {
      const left = new Date(a[CONFIG.fields.tasks.dueDate] || '2999-12-31').getTime();
      const right = new Date(b[CONFIG.fields.tasks.dueDate] || '2999-12-31').getTime();
      return left - right;
    });

    const history = getHistoryForContact(contact.Id).sort(function (a, b) {
      const left = new Date(a[CONFIG.fields.history.date] || '1900-01-01').getTime();
      const right = new Date(b[CONFIG.fields.history.date] || '1900-01-01').getTime();
      return right - left;
    });

    if (dom.contactDetailTitle) {
      dom.contactDetailTitle.textContent = getContactName(contact);
    }

    if (dom.contactDetailMeta) {
      dom.contactDetailMeta.textContent = [
        firm ? getFirmName(firm) : 'Ohne Firma',
        valueOrEmpty(contact[CONFIG.fields.contacts.email]),
        valueOrEmpty(contact[CONFIG.fields.contacts.phone])
      ]
        .filter(Boolean)
        .join(' · ');
    }

    if (dom.contactLeadDisplay) {
      const leadValue = valueOrEmpty(contact[CONFIG.fields.contacts.leadDisplay]);
      if ('value' in dom.contactLeadDisplay) {
        dom.contactLeadDisplay.value = leadValue;
      } else {
        dom.contactLeadDisplay.textContent = leadValue;
      }
    }

    const tasksHtml = tasks.length
      ? tasks
          .map(function (task) {
            return (
              '<div class="border rounded-md px-3 py-2 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(valueOrEmpty(task[CONFIG.fields.tasks.title]) || 'Ohne Titel') +
              '</div>' +
              '<div class="text-sm text-slate-600">Fällig: ' +
              escapeHtml(formatDate(task[CONFIG.fields.tasks.dueDate])) +
              ' · Status: ' +
              escapeHtml(valueOrEmpty(task[CONFIG.fields.tasks.status]) || '—') +
              '</div>' +
              '<div class="text-sm whitespace-pre-wrap">' +
              escapeHtml(valueOrEmpty(task[CONFIG.fields.tasks.notes])) +
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
              '<div class="border rounded-md px-3 py-2 mb-2 border-slate-200">' +
              '<div class="font-medium">' +
              escapeHtml(formatDate(entry[CONFIG.fields.history.date])) +
              ' · ' +
              escapeHtml(valueOrEmpty(entry[CONFIG.fields.history.type])) +
              '</div>' +
              '<div class="text-sm whitespace-pre-wrap">' +
              escapeHtml(valueOrEmpty(entry[CONFIG.fields.history.notes])) +
              '</div>' +
              '<div class="text-xs text-slate-500">Projektbezug: ' +
              (boolFromField(entry[CONFIG.fields.history.projectRelated]) ? 'Ja' : 'Nein') +
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
      '<div><strong>Firma:</strong> ' +
      escapeHtml(firm ? getFirmName(firm) : '—') +
      '</div>' +
      '<div><strong>Rolle:</strong> ' +
      escapeHtml(valueOrEmpty(contact[CONFIG.fields.contacts.role]) || '—') +
      '</div>' +
      '<div><strong>E-Mail:</strong> ' +
      escapeHtml(valueOrEmpty(contact[CONFIG.fields.contacts.email]) || '—') +
      '</div>' +
      '<div><strong>Telefon:</strong> ' +
      escapeHtml(valueOrEmpty(contact[CONFIG.fields.contacts.phone]) || '—') +
      '</div>' +
      '<div><strong>Mobile:</strong> ' +
      escapeHtml(valueOrEmpty(contact[CONFIG.fields.contacts.mobile]) || '—') +
      '</div>' +
      '<div><strong>Leadbbz0:</strong> ' +
      escapeHtml(valueOrEmpty(contact[CONFIG.fields.contacts.leadDisplay]) || '—') +
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
          return '<option value="' + escapeHtml(String(firm.Id)) + '">' + escapeHtml(getFirmName(firm)) + '</option>';
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
            const firm = getFirmById(contact[CONFIG.fields.contacts.firmLookupId]);
            const label = getContactName(contact) + (firm ? ' (' + getFirmName(firm) + ')' : '');
            return '<option value="' + escapeHtml(String(contact.Id)) + '">' + escapeHtml(label) + '</option>';
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
    setStatus('Lade CRM-Daten ...', false);

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
      state.selectedFirmId = state.firms[0].Id;
    }

    if (state.selectedFirmId) {
      const firmExists = !!getFirmById(state.selectedFirmId);
      if (!firmExists) {
        state.selectedFirmId = state.firms.length ? state.firms[0].Id : null;
      }
    }

    if (state.selectedFirmId) {
      const contactsForFirm = getContactsForFirm(state.selectedFirmId);
      const currentContactStillValid = contactsForFirm.some(function (c) {
        return String(c.Id) === String(state.selectedContactId);
      });
      if (!currentContactStillValid) {
        state.selectedContactId = contactsForFirm.length ? contactsForFirm[0].Id : null;
      }
    } else {
      state.selectedContactId = null;
    }

    renderAll();
    setStatus('CRM-Daten geladen.', false);
  }

  function selectFirm(firmId) {
    state.selectedFirmId = String(firmId);

    const contacts = getContactsForFirm(state.selectedFirmId);
    const stillValid = contacts.some(function (c) {
      return String(c.Id) === String(state.selectedContactId);
    });

    if (!stillValid) {
      state.selectedContactId = contacts.length ? String(contacts[0].Id) : null;
    }

    renderAll();
  }

  function selectContact(contactId) {
    const contact = getContactById(contactId);
    state.selectedContactId = String(contactId);
    if (contact) {
      state.selectedFirmId = String(contact[CONFIG.fields.contacts.firmLookupId] || state.selectedFirmId || '');
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

    await loadAllData();
  }

  document.addEventListener('DOMContentLoaded', function () {
    init().catch(handleError);
  });

  window.CRM_APP = {
    init: init,
    reload: loadAllData,
    selectFirm: selectFirm,
    selectContact: selectContact,
    state: state,
    config: CONFIG
  };
})();

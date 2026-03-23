// =====================================================
// CRM bbz Light - Sprint 1
// Basis: bestehender funktionierender Login/Graph-Stand
// =====================================================

// --- CONFIG & VERSION ---
const appVersion = "0.60";
const appName = "CRM bbz";
const config = {
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
    siteSearch: "bbzsg.sharepoint.com:/sites/CRM",
    redirectUri: "https://markusbaechler.github.io/crm-spa/"
};

// --- APP STATE ---
const state = {
    initialized: false,
    activeView: "dashboard",
    currentSiteId: "",
    currentFirmListId: "",
    currentContactListId: "",
    account: null,

    allFirms: [],
    allContacts: [],

    meta: {
        klassen: [],
        anreden: [],
        rollen: [],
        leads: [],
        sgf: [],
        events: []
    }
};

// --- MSAL ---
const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: config.redirectUri
    },
    cache: {
        cacheLocation: "localStorage"
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginRequest = {
    scopes: [
        "https://graph.microsoft.com/AllSites.Write",
        "https://graph.microsoft.com/AllSites.Read"
    ]
};

// =====================================================
// INITIALISIERUNG
// =====================================================
window.onload = async () => {
    updateFooter();

    try {
        await msalInstance.handleRedirectPromise();
        checkAuthState();
    } catch (e) {
        console.error("Ladefehler:", e);
        showBanner(`Login-Initialisierung fehlgeschlagen: ${e.message}`, "error");
    }
};

function updateFooter() {
    const ft = document.getElementById("footer-text");
    if (ft) ft.innerHTML = `© 2026 ${appName} | <b>Build ${appVersion}</b>`;
}

function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    const btn = document.getElementById("authBtn");

    if (!btn) return;

    if (accounts.length > 0) {
        state.account = accounts[0];
        btn.innerText = "Logout";
        btn.onclick = () => msalInstance.logoutRedirect({ account: accounts[0] });
        loadData();
    } else {
        state.account = null;
        btn.innerText = "Login";
        btn.onclick = () => msalInstance.loginRedirect(loginRequest);
        renderLoggedOutWelcome();
    }
}

// =====================================================
// UI BASIS
// =====================================================
function renderLoggedOutWelcome() {
    const content = document.getElementById("main-content");
    if (!content) return;

    content.innerHTML = `
        <div class="text-center py-12">
            <h2 class="text-3xl font-bold text-gray-900 mb-4">Willkommen im CRM</h2>
            <p class="text-gray-500 mb-8">
                Bitte loggen Sie sich oben rechts ein. Danach lädt das Dashboard automatisch.
            </p>
            <div class="flex justify-center space-x-4">
                <div class="w-12 h-1 bg-blue-500 rounded"></div>
                <div class="w-12 h-1 bg-slate-300 rounded"></div>
                <div class="w-12 h-1 bg-slate-300 rounded"></div>
            </div>
        </div>
    `;
}

function setLoading(message = "Lade Datenbank-Intelligenz...") {
    const content = document.getElementById("main-content");
    if (!content) return;

    content.innerHTML = `
        <div class="p-20 text-center text-slate-400 font-bold uppercase tracking-widest animate-pulse italic">
            ${escapeHtml(message)}
        </div>
    `;
}

function showBanner(message, type = "info") {
    const el = document.getElementById("status-banner");
    if (!el) return;

    const styles = {
        info: "bg-blue-50 text-blue-700 border border-blue-200",
        success: "bg-emerald-50 text-emerald-700 border border-emerald-200",
        error: "bg-red-50 text-red-700 border border-red-200"
    };

    el.className = `mb-4 rounded-xl px-4 py-3 text-sm font-semibold ${styles[type] || styles.info}`;
    el.textContent = message;
    el.classList.remove("hidden");
}

function hideBanner() {
    const el = document.getElementById("status-banner");
    if (!el) return;
    el.classList.add("hidden");
    el.textContent = "";
}

function setActiveNav(view) {
    const navMap = {
        dashboard: "nav-dashboard",
        firms: "nav-firms",
        contacts: "nav-contacts",
        planning: "nav-planning"
    };

    Object.values(navMap).forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        el.classList.remove("text-blue-400", "font-bold");
        el.classList.add("text-white");
    });

    const activeId = navMap[view];
    const activeEl = document.getElementById(activeId);
    if (activeEl) {
        activeEl.classList.remove("text-white");
        activeEl.classList.add("text-blue-400", "font-bold");
    }
}

// =====================================================
// GRAPH HELPERS
// =====================================================
async function getAccessToken() {
    const accounts = msalInstance.getAllAccounts();
    if (!accounts.length) throw new Error("Kein aktiver Login vorhanden.");

    const response = await msalInstance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0]
    });

    return response.accessToken;
}

async function graphGet(url) {
    const token = await getAccessToken();
    const response = await fetch(url, {
        headers: {
            Authorization: `Bearer ${token}`
        }
    });

    const data = await response.json();

    if (!response.ok) {
        throw new Error(data?.error?.message || `GET fehlgeschlagen: ${response.status}`);
    }

    return data;
}

async function graphPost(url, body) {
    const token = await getAccessToken();
    const response = await fetch(url, {
        method: "POST",
        headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });

    const text = await response.text();
    let data = {};
    try {
        data = text ? JSON.parse(text) : {};
    } catch {
        data = {};
    }

    if (!response.ok) {
        throw new Error(data?.error?.message || `POST fehlgeschlagen: ${response.status}`);
    }

    return data;
}

async function graphPatch(url, body) {
    const token = await getAccessToken();
    const response = await fetch(url, {
        method: "PATCH",
        headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });

    const text = await response.text();
    let data = {};
    try {
        data = text ? JSON.parse(text) : {};
    } catch {
        data = {};
    }

    if (!response.ok) {
        throw new Error(data?.error?.message || `PATCH fehlgeschlagen: ${response.status}`);
    }

    return data;
}

// =====================================================
// DATEN LADEN
// =====================================================
async function loadData() {
    setLoading();
    hideBanner();

    try {
        const site = await graphGet(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`);
        state.currentSiteId = site.id;

        const lists = await graphGet(`https://graph.microsoft.com/v1.0/sites/${state.currentSiteId}/lists`);

        const firmList = lists.value.find(x => x.displayName === "CRMFirms");
        const contactList = lists.value.find(x => x.displayName === "CRMContacts");

        if (!firmList) throw new Error('Liste "CRMFirms" wurde nicht gefunden.');
        if (!contactList) throw new Error('Liste "CRMContacts" wurde nicht gefunden.');

        state.currentFirmListId = firmList.id;
        state.currentContactListId = contactList.id;

        const [
            cKlass,
            cAnrede,
            cRolle,
            cLead,
            cSGF,
            cEvent,
            firms,
            contacts
        ] = await Promise.all([
            safeGetColumn(state.currentFirmListId, "Klassifizierung"),
            safeGetColumn(state.currentContactListId, "Anrede"),
            safeGetColumn(state.currentContactListId, "Rolle"),
            safeGetColumn(state.currentContactListId, "Leadbbz0"),
            safeGetColumn(state.currentContactListId, "SGF"),
            safeGetColumn(state.currentContactListId, "Event"),
            graphGet(`https://graph.microsoft.com/v1.0/sites/${state.currentSiteId}/lists/${state.currentFirmListId}/items?expand=fields&top=999`),
            graphGet(`https://graph.microsoft.com/v1.0/sites/${state.currentSiteId}/lists/${state.currentContactListId}/items?expand=fields&top=999`)
        ]);

        state.meta = {
            klassen: cKlass.choice?.choices || [],
            anreden: cAnrede.choice?.choices || [],
            rollen: cRolle.choice?.choices || [],
            leads: cLead.choice?.choices || [],
            sgf: cSGF.choice?.choices || [],
            events: cEvent.choice?.choices || []
        };

        state.allFirms = firms.value || [];
        state.allContacts = contacts.value || [];
        state.initialized = true;

        navigateTo("dashboard");
        showBanner("Daten erfolgreich geladen.", "success");
    } catch (e) {
        console.error("Sync-Fehler:", e);
        const content = document.getElementById("main-content");
        if (content) {
            content.innerHTML = `
                <div class="p-10 rounded-2xl bg-red-50 border border-red-200 text-red-700">
                    <div class="font-black uppercase text-sm mb-2">Sync-Fehler</div>
                    <div>${escapeHtml(e.message)}</div>
                </div>
            `;
        }
        showBanner(`Daten konnten nicht geladen werden: ${e.message}`, "error");
    }
}

async function safeGetColumn(listId, columnName) {
    try {
        return await graphGet(`https://graph.microsoft.com/v1.0/sites/${state.currentSiteId}/lists/${listId}/columns/${columnName}`);
    } catch (e) {
        console.warn(`Spalte ${columnName} nicht lesbar:`, e.message);
        return {};
    }
}

// =====================================================
// NAVIGATION
// =====================================================
function navigateTo(view) {
    if (!state.initialized && view !== "dashboard") {
        renderLoggedOutWelcome();
        return;
    }

    state.activeView = view;
    setActiveNav(view);

    if (view === "dashboard") return renderDashboard();
    if (view === "firms") return renderFirms();
    if (view === "contacts") return renderAllContacts();
    if (view === "planning") return renderPlanning();
}

function loadFirms() {
    navigateTo("firms");
}

function loadAllContacts() {
    navigateTo("contacts");
}

// =====================================================
// DASHBOARD
// =====================================================
function renderDashboard() {
    const content = document.getElementById("main-content");
    if (!content) return;

    if (!state.initialized) {
        renderLoggedOutWelcome();
        return;
    }

    const firmsA = state.allFirms.filter(f => getField(f, "Klassifizierung") === "A").length;
    const firmsB = state.allFirms.filter(f => getField(f, "Klassifizierung") === "B").length;
    const firmsC = state.allFirms.filter(f => getField(f, "Klassifizierung") === "C").length;

    const contactsWithoutMail = state.allContacts.filter(c => !getField(c, "Email1")).length;
    const contactsWithEvent = state.allContacts.filter(c => !!getField(c, "Event")).length;

    const latestContacts = [...state.allContacts]
        .sort((a, b) => Number(b.id) - Number(a.id))
        .slice(0, 8);

    content.innerHTML = `
        <div class="max-w-[1600px] mx-auto">
            <div class="mb-8">
                <h1 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter">Dashboard</h1>
                <p class="text-slate-500 mt-2">Übersicht über Firmen- und Kontaktbestand.</p>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-5 gap-6 mb-8">
                ${renderKpiCard("Firmen gesamt", state.allFirms.length, "🏢")}
                ${renderKpiCard("Kontakte gesamt", state.allContacts.length, "👥")}
                ${renderKpiCard("A-Firmen", firmsA, "A")}
                ${renderKpiCard("Kontakte ohne E-Mail", contactsWithoutMail, "✉️")}
                ${renderKpiCard("Kontakte mit Event", contactsWithEvent, "🏷️")}
            </div>

            <div class="grid grid-cols-1 xl:grid-cols-3 gap-8">
                <div class="bg-white border rounded-3xl p-8 shadow-sm">
                    <h2 class="text-lg font-black text-slate-800 uppercase tracking-tight mb-6">Firmenverteilung ABC</h2>
                    <div class="space-y-4">
                        ${renderDistributionRow("A", firmsA, "emerald")}
                        ${renderDistributionRow("B", firmsB, "amber")}
                        ${renderDistributionRow("C", firmsC, "slate")}
                    </div>
                </div>

                <div class="bg-white border rounded-3xl p-8 shadow-sm xl:col-span-2">
                    <div class="flex justify-between items-center mb-6">
                        <h2 class="text-lg font-black text-slate-800 uppercase tracking-tight">Neueste Kontakte</h2>
                        <button onclick="navigateTo('contacts')" class="text-blue-600 text-sm font-bold hover:underline">Alle Kontakte</button>
                    </div>

                    <div class="overflow-x-auto">
                        <table class="w-full text-left text-sm">
                            <thead class="border-b text-xs uppercase tracking-widest text-slate-400">
                                <tr>
                                    <th class="py-3 pr-4">Name</th>
                                    <th class="py-3 pr-4">Firma</th>
                                    <th class="py-3 pr-4">Rolle</th>
                                    <th class="py-3 pr-4">E-Mail</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-100">
                                ${latestContacts.length ? latestContacts.map(c => {
                                    const firm = getFirmById(getField(c, "FirmaLookupId"));
                                    return `
                                        <tr class="hover:bg-slate-50">
                                            <td class="py-3 pr-4 font-bold text-slate-800 underline cursor-pointer" onclick="renderContactDetail('${c.id}')">
                                                ${escapeHtml(formatContactName(c))}
                                            </td>
                                            <td class="py-3 pr-4">
                                                ${firm ? `<span class="cursor-pointer hover:underline" onclick="renderDetail('${firm.id}')">${escapeHtml(getField(firm, "Title"))}</span>` : "-"}
                                            </td>
                                            <td class="py-3 pr-4">${escapeHtml(getField(c, "Rolle") || "-")}</td>
                                            <td class="py-3 pr-4">${escapeHtml(getField(c, "Email1") || "-")}</td>
                                        </tr>
                                    `;
                                }).join("") : `
                                    <tr>
                                        <td colspan="4" class="py-6 text-slate-400 italic">Keine Kontakte vorhanden.</td>
                                    </tr>
                                `}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    `;
}

function renderKpiCard(label, value, badge) {
    return `
        <div class="bg-white border rounded-3xl p-6 shadow-sm">
            <div class="flex items-start justify-between mb-4">
                <div class="text-slate-400 text-xs uppercase tracking-widest font-black">${escapeHtml(label)}</div>
                <div class="text-xs bg-slate-100 text-slate-600 px-2 py-1 rounded-xl font-black">${escapeHtml(String(badge))}</div>
            </div>
            <div class="text-4xl font-black text-slate-800 tracking-tighter">${escapeHtml(String(value))}</div>
        </div>
    `;
}

function renderDistributionRow(label, value, tone) {
    const total = state.allFirms.length || 1;
    const pct = Math.round((value / total) * 100);

    const colorMap = {
        emerald: "bg-emerald-500",
        amber: "bg-amber-500",
        slate: "bg-slate-400"
    };

    return `
        <div>
            <div class="flex justify-between mb-2 text-sm font-bold text-slate-700">
                <span>${escapeHtml(label)}</span>
                <span>${escapeHtml(String(value))} · ${escapeHtml(String(pct))}%</span>
            </div>
            <div class="w-full h-3 bg-slate-100 rounded-full overflow-hidden">
                <div class="${colorMap[tone] || "bg-slate-400"} h-3 rounded-full" style="width:${pct}%"></div>
            </div>
        </div>
    `;
}

// =====================================================
// FIRMEN
// =====================================================
function renderFirms() {
    const content = document.getElementById("main-content");
    if (!content) return;

    content.innerHTML = `
        <div class="max-w-[1600px] mx-auto">
            <div class="flex flex-col xl:flex-row xl:justify-between xl:items-center gap-4 mb-8 border-b pb-6 border-slate-100">
                <div class="flex gap-10 items-end">
                    <h2 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter">Firmen</h2>
                </div>

                <div class="flex flex-col md:flex-row gap-4">
                    <input
                        type="text"
                        id="firmSearch"
                        onkeyup="filterFirms(this.value)"
                        placeholder="Firma oder Ort suchen..."
                        class="px-5 py-2.5 bg-slate-50 rounded-2xl text-sm w-full md:w-80 shadow-inner font-bold italic outline-none"
                    />
                    <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-6 py-2.5 rounded-xl text-[10px] font-black uppercase">
                        + Firma
                    </button>
                </div>
            </div>

            <div id="addForm" class="hidden mb-8 p-6 bg-slate-50 border rounded-2xl grid grid-cols-1 md:grid-cols-3 gap-4 shadow-inner">
                <input type="text" id="new_fName" placeholder="Name *" class="p-3 border rounded text-sm" />
                <select id="new_fClass" class="p-3 border rounded text-sm">
                    <option value="">- Klasse -</option>
                    ${state.meta.klassen.map(o => `<option value="${escapeHtmlAttr(o)}">${escapeHtml(o)}</option>`).join("")}
                </select>
                <input type="text" id="new_fCity" placeholder="Ort" class="p-3 border rounded text-sm" />
                <button onclick="saveNewFirm()" class="bg-green-600 text-white font-bold text-xs p-3 rounded uppercase shadow-md">
                    In SharePoint speichern
                </button>
            </div>

            <div id="firmList" class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-8">
                ${renderFirmCards(state.allFirms)}
            </div>
        </div>
    `;
}

function renderFirmCards(firms) {
    if (!firms.length) {
        return `
            <div class="col-span-full bg-slate-50 border border-dashed border-slate-200 rounded-3xl p-10 text-slate-400 italic">
                Keine Firmen gefunden.
            </div>
        `;
    }

    return firms.map(f => {
        const firmClass = getField(f, "Klassifizierung");
        const tone =
            firmClass === "A" ? "border-t-emerald-500" :
            firmClass === "B" ? "border-t-amber-500" :
            "border-t-slate-200";

        const contactCount = state.allContacts.filter(c => String(getField(c, "FirmaLookupId")) === String(f.id)).length;

        return `
            <div onclick="renderDetail('${f.id}')" class="bg-white border p-8 rounded-[2.5rem] hover:shadow-2xl transition-all cursor-pointer border-t-[6px] ${tone}">
                <h3 class="font-black text-slate-800 text-xl uppercase italic tracking-tighter mb-4 underline">
                    ${escapeHtml(getField(f, "Title") || "-")}
                </h3>

                <div class="text-[11px] font-bold text-slate-400 uppercase tracking-widest italic mb-8">
                    📍 ${escapeHtml(getField(f, "Ort") || "-")}
                </div>

                <div class="mt-auto pt-6 flex justify-between items-center border-t border-slate-50">
                    <span class="text-[10px] font-black bg-slate-50 px-3 py-1.5 rounded-xl border italic uppercase">
                        ${escapeHtml(firmClass || "-")}
                    </span>
                    <span class="text-[10px] font-black text-blue-600 bg-blue-50 px-3 py-1.5 rounded-xl uppercase italic">
                        👥 ${contactCount}
                    </span>
                </div>
            </div>
        `;
    }).join("");
}

function renderDetail(id) {
    const firm = getFirmById(id);
    if (!firm) {
        showBanner("Firma nicht gefunden.", "error");
        return;
    }

    const f = firm.fields || {};
    const contacts = state.allContacts.filter(c => String(getField(c, "FirmaLookupId")) === String(id));

    setActiveNav("firms");

    document.getElementById("main-content").innerHTML = `
        <div class="max-w-[1600px] mx-auto">
            <div class="bg-white border rounded-2xl p-6 mb-8 flex justify-between items-center shadow-sm">
                <div class="flex items-center gap-6">
                    <button onclick="renderFirms()" class="bg-slate-50 p-2 rounded-xl text-slate-400">←</button>
                    <h2 class="text-2xl font-black text-slate-800 uppercase italic tracking-tighter">
                        ${escapeHtml(f.Title || "-")}
                    </h2>
                </div>

                <button onclick="toggleContactForm()" class="bg-blue-600 text-white px-5 py-2.5 rounded-xl text-[10px] font-black uppercase">
                    + Kontakt
                </button>
            </div>

            <div class="grid grid-cols-1 xl:grid-cols-3 gap-8 mb-8">
                <div class="bg-white border rounded-3xl p-8 shadow-sm">
                    <div class="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-4">Profil Firma</div>
                    <div class="space-y-4 text-sm">
                        <div><span class="font-black text-slate-500 uppercase text-[10px]">Name</span><div class="mt-1 font-bold text-slate-800">${escapeHtml(f.Title || "-")}</div></div>
                        <div><span class="font-black text-slate-500 uppercase text-[10px]">Klassifizierung</span><div class="mt-1 font-bold text-slate-800">${escapeHtml(f.Klassifizierung || "-")}</div></div>
                        <div><span class="font-black text-slate-500 uppercase text-[10px]">Ort</span><div class="mt-1 font-bold text-slate-800">${escapeHtml(f.Ort || "-")}</div></div>
                        <div><span class="font-black text-slate-500 uppercase text-[10px]">Kontakte</span><div class="mt-1 font-bold text-slate-800">${contacts.length}</div></div>
                    </div>
                </div>

                <div class="xl:col-span-2">
                    <div id="addContactForm" class="hidden bg-white border border-blue-100 p-8 rounded-3xl mb-8 shadow-xl">
                        <div class="grid grid-cols-1 md:grid-cols-4 gap-6 mb-6">
                            <div>
                                <label class="text-[9px] font-black uppercase text-slate-400">Anrede</label>
                                <select id="c_anrede" class="w-full p-3 border rounded text-sm">
                                    <option value="">- leer -</option>
                                    ${state.meta.anreden.map(o => `<option value="${escapeHtmlAttr(o)}">${escapeHtml(o)}</option>`).join("")}
                                </select>
                            </div>

                            <div>
                                <label class="text-[9px] font-black uppercase text-slate-400">Vorname</label>
                                <input type="text" id="c_vn" class="w-full p-3 border rounded text-sm" />
                            </div>

                            <div>
                                <label class="text-[9px] font-black uppercase text-slate-400">Nachname *</label>
                                <input type="text" id="c_nn" class="w-full p-3 border rounded text-sm font-black" />
                            </div>

                            <div>
                                <label class="text-[9px] font-black uppercase text-slate-400">Rolle</label>
                                <select id="c_rolle" class="w-full p-3 border rounded text-sm font-bold">
                                    <option value="">- leer -</option>
                                    ${state.meta.rollen.map(o => `<option value="${escapeHtmlAttr(o)}">${escapeHtml(o)}</option>`).join("")}
                                </select>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 md:grid-cols-4 gap-6 mb-6">
                            <div>
                                <label class="text-[9px] font-black uppercase text-slate-400">E-Mail</label>
                                <input type="text" id="c_email1" class="w-full p-3 border rounded text-sm" />
                            </div>

                            <div>
                                <label class="text-[9px] font-black uppercase text-slate-400">Mobile</label>
                                <input type="text" id="c_mo" class="w-full p-3 border rounded text-sm font-black" />
                            </div>

                            <div>
                                <label class="text-[9px] font-black uppercase text-slate-400">SGF</label>
                                <select id="c_sgf" class="w-full p-3 border rounded text-sm font-bold">
                                    <option value="">- leer -</option>
                                    ${state.meta.sgf.map(o => `<option value="${escapeHtmlAttr(o)}">${escapeHtml(o)}</option>`).join("")}
                                </select>
                            </div>

                            <div>
                                <label class="text-[9px] font-black uppercase text-slate-400">Lead bbz</label>
                                <select id="c_lead" class="w-full p-3 border rounded text-sm font-bold">
                                    <option value="">- leer -</option>
                                    ${state.meta.leads.map(o => `<option value="${escapeHtmlAttr(o)}">${escapeHtml(o)}</option>`).join("")}
                                </select>
                            </div>
                        </div>

                        <div class="flex flex-wrap gap-4">
                            <button onclick="saveContact('${id}')" class="bg-blue-600 text-white font-black text-[10px] uppercase px-8 py-3 rounded-2xl shadow-lg">
                                In SharePoint speichern
                            </button>
                            <button onclick="toggleContactForm()" class="text-slate-400 font-bold uppercase text-[10px] px-4 underline">
                                Abbrechen
                            </button>
                        </div>
                    </div>

                    <div class="bg-white border rounded-[2rem] overflow-hidden shadow-sm">
                        <table class="w-full text-left text-[11px] italic">
                            <thead class="bg-slate-50 border-b text-[9px] font-black text-slate-400 uppercase tracking-widest italic">
                                <tr>
                                    <th class="p-4">Name</th>
                                    <th class="p-4">Rolle / Funktion</th>
                                    <th class="p-4">Events</th>
                                    <th class="p-4">Kontaktinfo</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y italic">
                                ${contacts.length ? contacts.map(c => `
                                    <tr class="hover:bg-slate-50 transition-colors group">
                                        <td onclick="renderContactDetail('${c.id}')" class="p-4 font-bold text-slate-800 text-sm underline cursor-pointer group-hover:text-blue-600 uppercase tracking-tighter italic">
                                            ${escapeHtml(formatContactName(c))}
                                        </td>
                                        <td class="p-4">
                                            <div class="font-bold text-blue-600 uppercase text-[9px] tracking-widest">
                                                ${escapeHtml(getField(c, "Rolle") || "-")}
                                            </div>
                                        </td>
                                        <td class="p-4">
                                            <div class="flex flex-wrap gap-1">
                                                ${getField(c, "Event")
                                                    ? `<span class="px-1.5 py-0.5 bg-blue-50 text-blue-600 rounded text-[8px] font-black uppercase border border-blue-100">${escapeHtml(getField(c, "Event"))}</span>`
                                                    : "-"
                                                }
                                            </div>
                                        </td>
                                        <td class="p-4 text-slate-500 font-bold">
                                            📧 ${escapeHtml(getField(c, "Email1") || "-")}<br />
                                            📱 ${escapeHtml(getField(c, "Mobile") || "-")}
                                        </td>
                                    </tr>
                                `).join("") : `
                                    <tr>
                                        <td colspan="4" class="p-6 text-slate-400 italic">Noch keine Kontakte vorhanden.</td>
                                    </tr>
                                `}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    `;
}

// =====================================================
// KONTAKTE
// =====================================================
function renderAllContacts() {
    const content = document.getElementById("main-content");
    if (!content) return;

    content.innerHTML = `
        <div class="max-w-[1600px] mx-auto">
            <div class="flex flex-col xl:flex-row xl:justify-between xl:items-center gap-4 mb-8 border-b pb-6 border-slate-100">
                <div class="flex gap-10 items-end">
                    <h2 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter">Kontakte</h2>
                </div>

                <input
                    type="text"
                    onkeyup="filterContacts(this.value)"
                    placeholder="Person suchen..."
                    class="px-5 py-2.5 bg-slate-50 rounded-2xl text-sm w-full xl:w-96 shadow-inner font-bold italic outline-none"
                />
            </div>

            <div class="bg-white border rounded-[2rem] overflow-hidden shadow-sm">
                <table class="w-full text-left text-[11px] italic">
                    <thead class="bg-slate-50 border-b text-[9px] uppercase font-black text-slate-400 tracking-widest">
                        <tr>
                            <th class="p-6">Name / Firma</th>
                            <th class="p-6">Rolle / SGF</th>
                            <th class="p-6">Events</th>
                            <th class="p-6">Kontaktinfo</th>
                        </tr>
                    </thead>
                    <tbody id="contactTableBody" class="divide-y divide-slate-50 italic">
                        ${renderContactsTableRows(state.allContacts)}
                    </tbody>
                </table>
            </div>
        </div>
    `;
}

function renderContactsTableRows(contacts) {
    if (!contacts.length) {
        return `
            <tr>
                <td colspan="4" class="p-6 text-slate-400 italic">Keine Kontakte gefunden.</td>
            </tr>
        `;
    }

    return contacts.map(c => {
        const firm = getFirmById(getField(c, "FirmaLookupId"));

        return `
            <tr class="hover:bg-blue-50/30 transition-all group">
                <td class="p-6">
                    <div class="text-base font-black text-slate-800 uppercase group-hover:text-blue-600 cursor-pointer underline" onclick="renderContactDetail('${c.id}')">
                        ${escapeHtml(formatContactName(c))}
                    </div>
                    <div class="text-[11px] font-bold text-slate-400 mt-1 cursor-pointer" onclick="${firm ? `renderDetail('${firm.id}')` : ""}">
                        🏢 ${escapeHtml(firm ? getField(firm, "Title") : "-")}
                    </div>
                </td>

                <td class="p-6">
                    <div class="text-[10px] font-black text-blue-500 uppercase">
                        ${escapeHtml(getField(c, "Rolle") || "-")}
                    </div>
                    <div class="text-[9px] font-black bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded border border-emerald-100 mt-1 inline-block uppercase">
                        ${escapeHtml(getField(c, "SGF") || "-")}
                    </div>
                </td>

                <td class="p-6">
                    <div class="flex flex-wrap gap-1">
                        ${getField(c, "Event")
                            ? `<span class="px-2 py-0.5 bg-blue-50 text-blue-600 border border-blue-100 rounded text-[8px] font-black uppercase">${escapeHtml(getField(c, "Event"))}</span>`
                            : "-"
                        }
                    </div>
                </td>

                <td class="p-6 font-bold text-slate-500">
                    📧 ${escapeHtml(getField(c, "Email1") || "-")}<br />
                    📱 ${escapeHtml(getField(c, "Mobile") || "-")}
                </td>
            </tr>
        `;
    }).join("");
}

function renderContactDetail(id) {
    const contact = state.allContacts.find(x => String(x.id) === String(id));
    if (!contact) {
        showBanner("Kontakt nicht gefunden.", "error");
        return;
    }

    const f = contact.fields || {};
    const firm = getFirmById(f.FirmaLookupId);

    setActiveNav("contacts");

    document.getElementById("main-content").innerHTML = `
        <div class="max-w-5xl mx-auto">
            <div class="bg-white border rounded-3xl p-8 mb-8 shadow-sm flex justify-between items-center gap-6">
                <div class="flex items-center gap-6">
                    <button onclick="${firm ? `renderDetail('${f.FirmaLookupId}')` : `renderAllContacts()`}" class="bg-slate-50 p-3 rounded-2xl text-xl text-slate-400">
                        ←
                    </button>

                    <div>
                        <h2 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter">
                            ${escapeHtml(formatContactName(contact))}
                        </h2>
                        <p class="text-slate-400 text-[10px] font-black uppercase mt-1">
                            Firma:
                            <span class="text-blue-600 ${firm ? "cursor-pointer" : ""}" ${firm ? `onclick="renderDetail('${f.FirmaLookupId}')"` : ""}>
                                ${escapeHtml(firm ? getField(firm, "Title") : "-")}
                            </span>
                        </p>
                    </div>
                </div>

                <button onclick="saveContactEdit('${id}')" class="bg-blue-600 text-white px-8 py-3 rounded-2xl text-[10px] font-black uppercase">
                    Änderungen speichern
                </button>
            </div>

            <div class="grid grid-cols-1 xl:grid-cols-12 gap-8 italic">
                <div class="xl:col-span-4 bg-white border rounded-3xl p-8 shadow-sm space-y-4">
                    <h3 class="text-[10px] font-black text-slate-400 uppercase tracking-widest border-b pb-4 mb-4">
                        Profil Kontakt
                    </h3>

                    <select id="ed_anrede" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic">
                        <option value="">-</option>
                        ${state.meta.anreden.map(o => `<option value="${escapeHtmlAttr(o)}" ${f.Anrede === o ? "selected" : ""}>${escapeHtml(o)}</option>`).join("")}
                    </select>

                    <input type="text" id="ed_vn" value="${escapeHtmlAttr(f.Vorname || "")}" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic" placeholder="Vorname" />
                    <input type="text" id="ed_nn" value="${escapeHtmlAttr(f.Title || "")}" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic" placeholder="Nachname" />

                    <select id="ed_rolle" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic">
                        <option value="">-</option>
                        ${state.meta.rollen.map(o => `<option value="${escapeHtmlAttr(o)}" ${f.Rolle === o ? "selected" : ""}>${escapeHtml(o)}</option>`).join("")}
                    </select>
                </div>

                <div class="xl:col-span-8 space-y-6">
                    <div class="bg-white border rounded-3xl p-8 shadow-sm grid grid-cols-1 md:grid-cols-2 gap-6">
                        <input type="text" id="ed_e1" value="${escapeHtmlAttr(f.Email1 || "")}" placeholder="E-Mail Geschäft" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm" />
                        <input type="text" id="ed_mo" value="${escapeHtmlAttr(f.Mobile || "")}" placeholder="Mobile" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm font-bold" />

                        <select id="ed_sgf" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm font-bold italic">
                            <option value="">SGF</option>
                            ${state.meta.sgf.map(o => `<option value="${escapeHtmlAttr(o)}" ${f.SGF === o ? "selected" : ""}>${escapeHtml(o)}</option>`).join("")}
                        </select>

                        <select id="ed_lead" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm font-bold italic">
                            <option value="">Lead bbz</option>
                            ${state.meta.leads.map(o => `<option value="${escapeHtmlAttr(o)}" ${f.Leadbbz0 === o ? "selected" : ""}>${escapeHtml(o)}</option>`).join("")}
                        </select>

                        <select id="ed_event" class="w-full p-3 bg-blue-50 border-none rounded-xl text-sm font-bold italic text-blue-600 md:col-span-2">
                            <option value="">Aktuelles Event</option>
                            ${state.meta.events.map(o => `<option value="${escapeHtmlAttr(o)}" ${f.Event === o ? "selected" : ""}>${escapeHtml(o)}</option>`).join("")}
                        </select>
                    </div>

                    <div class="bg-slate-50 border border-slate-200 rounded-3xl p-8 italic shadow-inner">
                        <label class="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-4 block">
                            Event-History / Kommentar
                        </label>
                        <textarea id="ed_kom" class="w-full p-4 bg-white border-none rounded-2xl text-sm shadow-sm min-h-[150px] font-medium" placeholder="Historie...">${escapeHtml(f.Kommentar || "")}</textarea>
                    </div>
                </div>
            </div>
        </div>
    `;
}

// =====================================================
// PLANUNG
// =====================================================
function renderPlanning() {
    const content = document.getElementById("main-content");
    if (!content) return;

    content.innerHTML = `
        <div class="max-w-[1200px] mx-auto">
            <div class="mb-8">
                <h1 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter">Planung</h1>
                <p class="text-slate-500 mt-2">Platzhalter für Tasks, Wiedervorlagen und Fälligkeiten.</p>
            </div>

            <div class="bg-white border rounded-3xl p-8 shadow-sm">
                <div class="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-4">Sprint 1</div>
                <h2 class="text-xl font-black text-slate-800 mb-3">Planungsmodul folgt in einem späteren Sprint</h2>
                <p class="text-slate-500 leading-relaxed">
                    Hier bauen wir im nächsten Schritt offene Tasks, fällige Wiedervorlagen und Priorisierung ein.
                </p>
            </div>
        </div>
    `;
}

// =====================================================
// ACTIONS
// =====================================================
async function saveContactEdit(id) {
    try {
        hideBanner();

        const fields = {
            Anrede: getValue("ed_anrede"),
            Vorname: getValue("ed_vn"),
            Title: getValue("ed_nn"),
            Rolle: getValue("ed_rolle"),
            Email1: getValue("ed_e1"),
            Mobile: getValue("ed_mo"),
            SGF: getValue("ed_sgf"),
            Leadbbz0: getValue("ed_lead"),
            Event: getValue("ed_event"),
            Kommentar: getValue("ed_kom")
        };

        await graphPatch(
            `https://graph.microsoft.com/v1.0/sites/${state.currentSiteId}/lists/${state.currentContactListId}/items/${id}/fields`,
            fields
        );

        showBanner("Kontakt erfolgreich gespeichert.", "success");
        await loadData();
        renderContactDetail(id);
    } catch (e) {
        console.error(e);
        showBanner(`Kontakt konnte nicht gespeichert werden: ${e.message}`, "error");
    }
}

async function saveContact(firmId) {
    try {
        hideBanner();

        const lastName = getValue("c_nn");
        if (!lastName) {
            showBanner("Nachname fehlt.", "error");
            return;
        }

        const fields = {
            Title: lastName,
            Vorname: getValue("c_vn"),
            Anrede: getValue("c_anrede"),
            Rolle: getValue("c_rolle"),
            Email1: getValue("c_email1"),
            Mobile: getValue("c_mo"),
            SGF: getValue("c_sgf"),
            Leadbbz0: getValue("c_lead"),
            FirmaLookupId: Number(firmId)
        };

        await graphPost(
            `https://graph.microsoft.com/v1.0/sites/${state.currentSiteId}/lists/${state.currentContactListId}/items`,
            { fields }
        );

        showBanner("Kontakt erfolgreich angelegt.", "success");
        await loadData();
        renderDetail(firmId);
    } catch (e) {
        console.error(e);
        showBanner(`Kontakt konnte nicht angelegt werden: ${e.message}`, "error");
    }
}

async function saveNewFirm() {
    try {
        hideBanner();

        const name = getValue("new_fName");
        if (!name) {
            showBanner("Firmenname fehlt.", "error");
            return;
        }

        const fields = {
            Title: name,
            Klassifizierung: getValue("new_fClass"),
            Ort: getValue("new_fCity")
        };

        await graphPost(
            `https://graph.microsoft.com/v1.0/sites/${state.currentSiteId}/lists/${state.currentFirmListId}/items`,
            { fields }
        );

        showBanner("Firma erfolgreich angelegt.", "success");
        await loadData();
        navigateTo("firms");
    } catch (e) {
        console.error(e);
        showBanner(`Firma konnte nicht angelegt werden: ${e.message}`, "error");
    }
}

// =====================================================
// FILTER / TOGGLES
// =====================================================
function filterFirms(query) {
    const q = (query || "").toLowerCase().trim();

    const filtered = state.allFirms.filter(f => {
        const name = (getField(f, "Title") || "").toLowerCase();
        const city = (getField(f, "Ort") || "").toLowerCase();
        return name.includes(q) || city.includes(q);
    });

    const list = document.getElementById("firmList");
    if (list) list.innerHTML = renderFirmCards(filtered);
}

function filterContacts(query) {
    const q = (query || "").toLowerCase().trim();

    const filtered = state.allContacts.filter(c => {
        const firstName = (getField(c, "Vorname") || "").toLowerCase();
        const lastName = (getField(c, "Title") || "").toLowerCase();
        const role = (getField(c, "Rolle") || "").toLowerCase();

        const firm = getFirmById(getField(c, "FirmaLookupId"));
        const firmName = (firm ? getField(firm, "Title") : "").toLowerCase();

        return firstName.includes(q) || lastName.includes(q) || role.includes(q) || firmName.includes(q);
    });

    const body = document.getElementById("contactTableBody");
    if (body) body.innerHTML = renderContactsTableRows(filtered);
}

function toggleContactForm() {
    const el = document.getElementById("addContactForm");
    if (el) el.classList.toggle("hidden");
}

function toggleAddForm() {
    const el = document.getElementById("addForm");
    if (el) el.classList.toggle("hidden");
}

// =====================================================
// MODAL
// =====================================================
function closeModal() {
    const modal = document.getElementById("detail-modal");
    if (modal) modal.classList.add("hidden");
    const body = document.getElementById("modal-body-content");
    if (body) body.innerHTML = "";
}

// =====================================================
// HELPERS
// =====================================================
function getValue(id) {
    const el = document.getElementById(id);
    return el ? el.value.trim() : "";
}

function getField(item, fieldName) {
    return item?.fields?.[fieldName];
}

function getFirmById(id) {
    return state.allFirms.find(x => String(x.id) === String(id));
}

function formatContactName(contact) {
    const first = getField(contact, "Vorname") || "";
    const last = getField(contact, "Title") || "";
    return `${first} ${last}`.trim() || "-";
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function escapeHtmlAttr(value) {
    return escapeHtml(value);
}

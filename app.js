// --- CONFIG & VERSION ---
const appVersion = "0.39";
console.log(`CRM App ${appVersion} - Full Build (Dashboard Architecture)`);

const config = {
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a",
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7",
    siteSearch: "bbzsg.sharepoint.com:/sites/CRM"
};

const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: "https://markusbaechler.github.io/crm-spa/"
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = {
    scopes: ["https://graph.microsoft.com/AllSites.Write", "https://graph.microsoft.com/AllSites.Read"]
};

// State Management
let allFirms = []; 
let allContacts = [];
let classOptions = []; 
let currentSiteId = "";
let currentListId = "";
let contactListId = "";

// --- INITIALISIERUNG ---
window.onload = async () => {
    updateFooter(); 
    try {
        await msalInstance.handleRedirectPromise();
        checkAuthState();
    } catch (err) {
        console.error("Initialisierungsfehler:", err);
    }
};

function updateFooter() {
    const footerText = document.getElementById('footer-text');
    if (footerText) {
        footerText.innerHTML = `© 2026 bbz CRM | <span class="text-slate-400 font-normal">Etappe E</span> | <span class="font-medium text-slate-600">Build ${appVersion}</span>`;
    }
}

function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    if (accounts.length > 0) {
        authBtn.innerText = "Abmelden";
        authBtn.onclick = () => msalInstance.logoutRedirect({ account: accounts[0] });
        authBtn.className = "text-slate-500 text-sm hover:text-slate-800 transition px-4 py-2";
        loadData(); 
    } else {
        authBtn.innerText = "Anmelden";
        authBtn.onclick = () => msalInstance.loginRedirect(loginRequest);
        authBtn.className = "bg-blue-600 text-white px-5 py-2 rounded-lg text-sm font-medium hover:bg-blue-700 transition shadow-sm";
    }
}

// --- DATA ENGINE ---
async function loadData() {
    const content = document.getElementById('main-content'); 
    if (!content) return;
    
    content.innerHTML = `
        <div class="flex flex-col items-center justify-center p-20">
            <div class="w-6 h-6 border-2 border-slate-200 border-t-blue-600 rounded-full animate-spin mb-3"></div>
            <p class="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Synchronisation läuft...</p>
        </div>`;

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })
            .catch(() => msalInstance.acquireTokenRedirect(loginRequest));
        const headers = { 'Authorization': `Bearer ${tokenRes.accessToken}` };

        // 1. IDs holen
        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers });
        const siteData = await siteRes.json();
        currentSiteId = siteData.id;
        
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers });
        const listsData = await listsRes.json();
        
        currentListId = listsData.value.find(l => l.displayName === "CRMFirms").id;
        contactListId = listsData.value.find(l => l.displayName === "CRMContacts").id;

        // 2. Daten & Metadaten
        const [colRes, firmsRes, contactsRes] = await Promise.all([
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/columns/Klassifizierung`, { headers }),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers }),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items?expand=fields`, { headers })
        ]);
        
        classOptions = (await colRes.json()).choice?.choices || ["-"];
        allFirms = (await firmsRes.json()).value;
        allContacts = (await contactsRes.json()).value;

        renderFirms(allFirms);
    } catch (err) { 
        content.innerHTML = `<div class="p-6 text-red-500 font-bold border border-red-100 bg-red-50 rounded-xl">Ladefehler: ${err.message}</div>`; 
    }
}

// Hintergrund-Reload nach Aktionen
async function loadDataSilent() {
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    const headers = { 'Authorization': `Bearer ${tokenRes.accessToken}` };
    const [f, c] = await Promise.all([
        fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers }),
        fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items?expand=fields`, { headers })
    ]);
    allFirms = (await f.json()).value;
    allContacts = (await c.json()).value;
}

// --- VIEW: FIRMENÜBERSICHT ---
function renderFirms(firms) {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-6xl mx-auto animate-in fade-in duration-500">
            <div class="flex justify-between items-end mb-10 border-b border-slate-100 pb-6">
                <div>
                    <h2 class="text-xl font-semibold text-slate-800 tracking-tight">Firmenstamm</h2>
                    <p class="text-slate-400 text-xs mt-1">Zentrale Verwaltung Ihrer Organisationen</p>
                </div>
                <div class="flex gap-3">
                    <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suchen..." class="px-4 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm outline-none focus:border-blue-400 w-64 transition-colors">
                    <button onclick="toggleAddForm()" class="bg-slate-800 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-slate-700 transition">+ Firma</button>
                </div>
            </div>

            <div id="addForm" class="hidden mb-10 p-6 bg-slate-50 border border-slate-200 rounded-xl">
                 <h3 class="text-[10px] font-bold text-slate-400 uppercase mb-4 tracking-widest">Neue Organisation erfassen</h3>
                 <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <input type="text" id="new_fName" placeholder="Firmenname" class="p-2.5 bg-white border border-slate-200 rounded-lg text-sm outline-none">
                    <select id="new_fClass" class="p-2.5 bg-white border border-slate-200 rounded-lg text-sm outline-none text-slate-500">
                        <option value="">Klassierung</option>
                        ${classOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
                    </select>
                    <input type="text" id="new_fCity" placeholder="Ort" class="p-2.5 bg-white border border-slate-200 rounded-lg text-sm outline-none">
                </div>
                <button onclick="saveNewFirm()" class="mt-4 bg-green-600 text-white px-8 py-2 rounded-lg text-[10px] font-bold uppercase tracking-widest hover:bg-green-700 transition">Speichern</button>
            </div>

            <div id="firmList" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                ${firms.map(item => {
                    const f = item.fields;
                    const count = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(item.id)).length;
                    return `
                    <div onclick="renderFirmDetailPage('${item.id}')" class="bg-white border border-slate-200 p-5 rounded-xl hover:border-blue-400 hover:shadow-md transition-all cursor-pointer group flex flex-col h-full border-t-4 ${f.Klassifizierung === 'A' ? 'border-t-emerald-500' : 'border-t-slate-200'}">
                        <div class="flex justify-between items-start mb-4">
                            <h3 class="font-medium text-slate-800 text-lg group-hover:text-blue-600 transition-colors leading-tight">${f.Title || 'Unbenannt'}</h3>
                            ${(f.VIP === true || f.VIP === "true") ? '<span class="text-amber-400">⭐</span>' : ''}
                        </div>
                        <div class="text-[11px] text-slate-400 mb-4 italic tracking-wide">📍 ${f.Ort || 'Kein Standort'}</div>
                        <div class="mt-auto pt-4 flex justify-between items-center border-t border-slate-50">
                            <span class="text-[10px] font-bold text-slate-400 uppercase tracking-widest bg-slate-50 px-2 py-1 rounded">Klasse ${f.Klassifizierung || '-'}</span>
                            <span class="text-[10px] font-semibold text-slate-500 bg-blue-50 px-2.5 py-1 rounded-full flex items-center gap-1">👥 ${count}</span>
                        </div>
                    </div>`;
                }).join('')}
            </div>
        </div>
    `;
}

// --- VIEW: FIRMEN-DETAILSEITE (DASHBOARD REDESIGN) ---
function renderFirmDetailPage(itemId) {
    const firm = allFirms.find(f => String(f.id) === String(itemId));
    const f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(itemId));
    
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-7xl mx-auto animate-in slide-in-from-right duration-500">
            
            <div class="bg-white border border-slate-200 rounded-2xl p-6 mb-8 shadow-sm flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
                <div class="flex items-center gap-5">
                    <button onclick="renderFirms(allFirms)" class="bg-slate-50 hover:bg-slate-100 text-slate-400 p-3 rounded-xl transition">←</button>
                    <div>
                        <div class="flex items-center gap-3">
                            <h2 class="text-2xl font-semibold text-slate-800 leading-none">${f.Title}</h2>
                            ${(f.VIP === true || f.VIP === "true") ? '<span class="text-amber-400 text-xl">⭐</span>' : ''}
                        </div>
                        <p class="text-slate-400 text-[10px] font-bold uppercase tracking-[0.2em] mt-2">
                            ${f.Ort || 'Ort n.a.'} <span class="mx-2 text-slate-200">|</span> ${f.Land || 'CH'} <span class="mx-2 text-slate-200">|</span> ${f.Hauptnummer || 'Keine Nummer'}
                        </p>
                    </div>
                </div>
                <div class="flex gap-3">
                    <button onclick="updateFirm('${itemId}')" class="bg-slate-800 text-white px-6 py-2.5 rounded-xl text-[10px] font-bold uppercase tracking-widest hover:bg-slate-700 transition shadow-lg">Änderungen speichern</button>
                </div>
            </div>

            <div class="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
                
                <div class="lg:col-span-4 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
                        <h3 class="text-[9px] font-black text-slate-300 uppercase tracking-[0.2em] mb-6 border-b pb-3">Unternehmensdaten</h3>
                        <div class="space-y-5">
                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Offizieller Name</label>
                                <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full mt-1.5 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm focus:border-blue-400 outline-none transition-all">
                            </div>
                            <div class="grid grid-cols-2 gap-4">
                                <div>
                                    <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Klassierung</label>
                                    <select id="edit_Klass" class="w-full mt-1.5 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm outline-none">
                                        ${classOptions.map(opt => `<option value="${opt}" ${f.Klassifizierung === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                                    </select>
                                </div>
                                <div class="flex items-end pb-1.5 px-2">
                                    <label class="flex items-center gap-2 cursor-pointer group">
                                        <input type="checkbox" id="edit_VIP" ${(f.VIP === true || f.VIP === "true") ? 'checked' : ''} class="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-0">
                                        <span class="text-[10px] font-bold text-slate-400 group-hover:text-slate-600 uppercase transition-colors">VIP Kunde</span>
                                    </label>
                                </div>
                            </div>
                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Telefon Zentrale</label>
                                <input type="text" id="edit_Phone" value="${f.Hauptnummer || ''}" class="w-full mt-1.5 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm outline-none">
                            </div>
                            <div class="pt-2">
                                <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Standortadresse</label>
                                <input type="text" id="edit_Street" value="${f.Adresse || ''}" placeholder="Strasse / Nr." class="w-full mt-1.5 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm outline-none mb-3">
                                <div class="flex gap-3">
                                    <input type="text" id="edit_City" value="${f.Ort || ''}" placeholder="Ort" class="flex-1 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm font-medium">
                                    <input type="text" id="edit_Country" value="${f.Land || 'CH'}" placeholder="CH" class="w-16 p-2.5 bg-slate-100 border border-slate-100 rounded-xl text-sm text-center font-bold text-slate-400 uppercase">
                                </div>
                            </div>
                            <div class="pt-6 border-t border-slate-50 flex justify-center">
                                <button onclick="deleteFirm('${itemId}', '${f.Title}')" class="text-red-400 text-[9px] font-bold uppercase hover:underline opacity-50 hover:opacity-100 transition">Organisation löschen</button>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="lg:col-span-8 space-y-8">
                    
                    <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
                        <div class="flex justify-between items-center mb-8 border-b border-slate-50 pb-4">
                            <h3 class="text-[10px] font-black text-slate-300 uppercase tracking-[0.2em]">Ansprechpartner (${contacts.length})</h3>
                            <button onclick="addContact('${itemId}')" class="text-blue-600 text-[10px] font-bold uppercase hover:bg-blue-50 px-4 py-2 rounded-xl transition">+ Neuer Kontakt</button>
                        </div>
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                            ${contacts.length > 0 ? contacts.map(c => `
                                <div class="p-5 bg-slate-50/50 border border-slate-100 rounded-2xl flex justify-between items-start group hover:bg-white hover:border-blue-200 hover:shadow-sm transition-all duration-300">
                                    <div class="flex-1">
                                        <div class="text-[9px] font-bold text-blue-500 uppercase tracking-widest mb-2">${c.fields.Anrede || ''} ${c.fields.Rolle || ''}</div>
                                        <div class="text-base font-semibold text-slate-800 mb-1">${c.fields.Vorname || ''} ${c.fields.Title}</div>
                                        <div class="text-[11px] text-slate-400 font-medium mb-4">${c.fields.Funktion || ''}</div>
                                        <div class="space-y-2 text-[10px] text-slate-500 font-medium border-t border-slate-100 pt-4">
                                            <div class="flex items-center gap-2">📧 ${c.fields.Email1 || '-'}</div>
                                            <div class="flex items-center gap-2">📞 ${c.fields.Direktwahl || '-'}</div>
                                            <div class="flex items-center gap-2">📱 ${c.fields.TelefonMobil || '-'}</div>
                                        </div>
                                    </div>
                                    <button onclick="deleteContact('${c.id}', '${itemId}')" class="text-slate-200 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity p-1">✕</button>
                                </div>
                            `).join('') : '<div class="col-span-2 p-12 text-center text-slate-300 text-[10px] font-bold uppercase tracking-widest border-2 border-dashed border-slate-50 rounded-2xl">Keine Kontakte zugeordnet</div>'}
                        </div>
                    </div>

                    <div class="grid grid-cols-1 md:grid-cols-2 gap-6 opacity-40 grayscale">
                        <div class="bg-slate-50 border border-slate-100 rounded-2xl p-10 flex flex-col items-center justify-center">
                            <span class="text-xl mb-2">📜</span>
                            <span class="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Aktivitätshistorie</span>
                        </div>
                        <div class="bg-slate-50 border border-slate-100 rounded-2xl p-10 flex flex-col items-center justify-center">
                            <span class="text-xl mb-2">📅</span>
                            <span class="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Aufgaben</span>
                        </div>
                    </div>

                </div>
            </div>
        </div>
    `;
}

// --- REST ACTIONS ---

async function updateFirm(itemId) {
    const fields = {
        Title: document.getElementById('edit_Title').value,
        Klassifizierung: document.getElementById('edit_Klass').value,
        Adresse: document.getElementById('edit_Street').value,
        Ort: document.getElementById('edit_City').value,
        Land: document.getElementById('edit_Country').value,
        Hauptnummer: document.getElementById('edit_Phone').value,
        VIP: document.getElementById('edit_VIP').checked
    };
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${itemId}/fields`, {
        method: 'PATCH', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(fields)
    });
    await loadDataSilent();
    renderFirmDetailPage(itemId);
}

async function saveNewFirm() {
    const name = document.getElementById('new_fName').value;
    if(!name) return;
    const fields = { Title: name, Klassifizierung: document.getElementById('new_fClass').value, Ort: document.getElementById('new_fCity').value };
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items`, {
        method: 'POST', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields: fields })
    });
    toggleAddForm(); loadData();
}

async function addContact(firmId) {
    const ln = prompt("Nachname (Pflicht):"); if (!ln) return;
    const vn = prompt("Vorname:");
    const fields = { Title: ln, Vorname: vn, FirmaLookupId: firmId };
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items`, {
        method: 'POST', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields: fields })
    });
    await loadDataSilent();
    renderFirmDetailPage(firmId);
}

async function deleteContact(cId, firmId) {
    if(!confirm("Kontakt entfernen?")) return;
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items/${cId}`, {
        method: 'DELETE', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
    });
    await loadDataSilent();
    renderFirmDetailPage(firmId);
}

async function deleteFirm(itemId, name) {
    if(!confirm(`Löschen von "${name}"?`)) return;
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${itemId}`, {
        method: 'DELETE', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
    });
    loadData();
}

function filterFirms(q) {
    const query = q.toLowerCase();
    const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query) || f.fields.Ort?.toLowerCase().includes(query));
    renderFirms(filtered);
}
function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }
function loadFirms() { renderFirms(allFirms); }
function loadAllContacts() {
    const content = document.getElementById('main-content');
    content.innerHTML = `<div class="p-10 text-center text-slate-300 font-bold uppercase tracking-widest">Funktion wird in V0.40 überarbeitet...</div>`;
}

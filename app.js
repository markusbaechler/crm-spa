// --- CONFIG & VERSION ---
const appVersion = "0.42";
console.log(`CRM App ${appVersion} - Stabilitäts-Check & Field-Fix`);

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

// Globaler State
let allFirms = []; 
let allContacts = [];
let classOptions = []; 
let currentSiteId = "";
let currentListId = "";
let contactListId = "";

window.onload = async () => {
    updateFooter(); 
    await msalInstance.handleRedirectPromise();
    checkAuthState();
};

function updateFooter() {
    const ft = document.getElementById('footer-text');
    if (ft) ft.innerHTML = `© 2026 bbz CRM | <span class="font-bold text-slate-600">Version ${appVersion}</span>`;
}

function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    if (accounts.length > 0) {
        authBtn.innerText = "Logout";
        authBtn.onclick = () => msalInstance.logoutRedirect({ account: accounts[0] });
        loadData(); 
    } else {
        authBtn.innerText = "Login";
        authBtn.onclick = () => msalInstance.loginRedirect(loginRequest);
    }
}

async function loadData() {
    const content = document.getElementById('main-content'); // WICHTIG: Muss 'main-content' sein laut HTML
    if (!content) return;
    content.innerHTML = `<div class="p-20 text-center text-slate-400 text-xs uppercase tracking-widest animate-pulse">Synchronisiere...</div>`;

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
        const headers = { 'Authorization': `Bearer ${tokenRes.accessToken}` };

        // 1. Site & Listen auflösen
        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers });
        const siteData = await siteRes.json();
        currentSiteId = siteData.id;
        
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers });
        const listsData = await listsRes.json();
        
        currentListId = listsData.value.find(l => l.displayName === "CRMFirms").id;
        contactListId = listsData.value.find(l => l.displayName === "CRMContacts").id;

        // 2. Alles parallel laden
        const [colRes, firmsRes, contactsRes] = await Promise.all([
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/columns/Klassifizierung`, { headers }),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers }),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items?expand=fields`, { headers })
        ]);
        
        classOptions = (await colRes.json()).choice?.choices || ["-"];
        allFirms = (await firmsRes.json()).value;
        allContacts = (await contactsRes.json()).value;

        renderFirms(allFirms);
    } catch (err) { content.innerHTML = `<div class="p-6 text-red-500 font-bold">Fehler: ${err.message}</div>`; }
}

// --- FIRMENÜBERSICHT ---
function renderFirms(firms) {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-6xl mx-auto animate-in fade-in duration-500">
            <div class="flex justify-between items-center mb-8 pb-4 border-b border-slate-100">
                <h2 class="text-2xl font-semibold text-slate-800 tracking-tight">Firmenstamm</h2>
                <div class="flex gap-3">
                    <input type="text" id="firmSearchInput" onkeyup="filterFirms(this.value)" placeholder="Suchen..." class="px-4 py-2 bg-white border border-slate-200 rounded-lg text-sm outline-none w-64 focus:border-blue-500">
                    <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-5 py-2 rounded-lg text-sm font-bold shadow-sm">+ Firma</button>
                </div>
            </div>

            <div id="addForm" class="hidden mb-10 p-6 bg-slate-50 border border-slate-200 rounded-2xl">
                 <h3 class="text-[10px] font-bold text-slate-400 uppercase mb-4 tracking-widest">Neue Firma erfassen</h3>
                 <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <input type="text" id="new_fName" placeholder="Name" class="p-2.5 bg-white border border-slate-200 rounded-xl text-sm">
                    <select id="new_fClass" class="p-2.5 bg-white border border-slate-200 rounded-xl text-sm">
                        ${classOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
                    </select>
                    <input type="text" id="new_fCity" placeholder="Ort" class="p-2.5 bg-white border border-slate-200 rounded-xl text-sm">
                </div>
                <button onclick="saveNewFirm()" class="mt-4 bg-green-600 text-white px-8 py-2.5 rounded-xl text-[10px] font-bold uppercase tracking-widest">Speichern</button>
            </div>

            <div id="firmList" class="grid grid-cols-1 md:grid-cols-3 gap-6">
                ${firms.map(item => `
                    <div onclick="renderDetailPage('${item.id}')" class="bg-white border border-slate-200 p-5 rounded-xl hover:border-blue-400 hover:shadow-md transition-all cursor-pointer flex flex-col h-full border-t-4 ${item.fields.Klassifizierung === 'A' ? 'border-t-emerald-500' : 'border-t-slate-200'}">
                        <div class="flex justify-between items-start mb-4">
                            <h3 class="font-medium text-slate-800 text-lg">${item.fields.Title || 'Unbenannt'}</h3>
                            ${(item.fields.VIP === true || item.fields.VIP === "true") ? '<span>⭐</span>' : ''}
                        </div>
                        <div class="text-[11px] text-slate-400 mb-4 italic tracking-wide">📍 ${item.fields.Ort || 'k.A.'}</div>
                        <div class="mt-auto pt-4 flex justify-between items-center border-t border-slate-50">
                            <span class="text-[10px] font-bold text-slate-400 uppercase bg-slate-50 px-2 py-1 rounded border border-slate-100">${item.fields.Klassifizierung || '-'}</span>
                            <span class="text-[10px] font-semibold text-slate-500 bg-blue-50 px-2.5 py-1 rounded-full">👥 ${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(item.id)).length}</span>
                        </div>
                    </div>`).join('')}
            </div>
        </div>
    `;
}

// --- DETAILSEITE (DASHBOARD) ---
function renderDetailPage(itemId) {
    const firm = allFirms.find(f => String(f.id) === String(itemId));
    const f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(itemId));
    
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-7xl mx-auto animate-in slide-in-from-right duration-300">
            <div class="bg-white border border-slate-200 rounded-2xl p-6 mb-8 shadow-sm flex flex-col md:flex-row justify-between items-center gap-6">
                <div class="flex items-center gap-5">
                    <button onclick="renderFirms(allFirms)" class="p-3 bg-slate-50 hover:bg-slate-100 rounded-xl transition text-xl text-slate-400">←</button>
                    <div>
                        <div class="flex items-center gap-3">
                            <h2 class="text-2xl font-semibold text-slate-800">${f.Title}</h2>
                            ${(f.VIP === true || f.VIP === "true") ? '<span>⭐</span>' : ''}
                        </div>
                        <div class="flex items-center gap-3 mt-2 text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                            <span>📍 ${f.Ort || '-'}</span>
                            <span class="bg-slate-100 px-2 py-0.5 rounded text-slate-500 border border-slate-200">${f.Land || 'CH'}</span>
                            <span class="text-slate-200">|</span>
                            <span>📞 ${f.Hauptnummer || '-'}</span>
                        </div>
                    </div>
                </div>
                <button onclick="updateFirm('${itemId}')" class="bg-slate-800 text-white px-8 py-3 rounded-xl text-[10px] font-bold uppercase tracking-widest hover:bg-slate-700 shadow-lg">Speichern</button>
            </div>

            <div class="grid grid-cols-12 gap-8 items-start">
                <div class="col-span-12 lg:col-span-4 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
                        <h3 class="text-[10px] font-black text-slate-300 uppercase tracking-widest mb-6 border-b pb-3">Unternehmensdaten</h3>
                        <div class="space-y-4">
                            <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold">
                            <div class="grid grid-cols-2 gap-4">
                                <select id="edit_Klass" class="p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold">
                                    ${classOptions.map(opt => `<option value="${opt}" ${f.Klassifizierung === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                                </select>
                                <label class="flex items-center gap-2"><input type="checkbox" id="edit_VIP" ${(f.VIP === true || f.VIP === "true") ? 'checked' : ''} class="w-4 h-4 rounded text-blue-600"> <span class="text-[10px] font-bold text-slate-500 uppercase">VIP</span></label>
                            </div>
                            <input type="text" id="edit_Phone" value="${f.Hauptnummer || ''}" placeholder="Hauptnummer" class="w-full p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm">
                            <input type="text" id="edit_Street" value="${f.Adresse || ''}" placeholder="Strasse" class="w-full p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm">
                            <div class="flex gap-2">
                                <input type="text" id="edit_City" value="${f.Ort || ''}" placeholder="Ort" class="flex-1 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold">
                                <input type="text" id="edit_Country" value="${f.Land || 'CH'}" class="w-12 p-2.5 bg-slate-100 border border-slate-100 rounded-xl text-sm text-center font-bold text-slate-400">
                            </div>
                            <button onclick="deleteFirm('${itemId}','${f.Title}')" class="w-full text-red-400 text-[9px] font-bold uppercase hover:underline pt-4">Löschen</button>
                        </div>
                    </div>
                </div>

                <div class="col-span-12 lg:col-span-8 space-y-8">
                    <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
                        <div class="flex justify-between items-center mb-8 border-b border-slate-50 pb-4">
                            <h3 class="text-[10px] font-black text-slate-300 uppercase tracking-widest">Ansprechpartner (${contacts.length})</h3>
                            <button onclick="addContact('${itemId}')" class="text-blue-600 text-[10px] font-bold uppercase hover:bg-blue-50 px-4 py-2 rounded-xl transition">+ Neuer Kontakt</button>
                        </div>
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                            ${contacts.length > 0 ? contacts.map(c => `
                                <div class="p-5 bg-slate-50 border border-slate-100 rounded-2xl relative group hover:bg-white hover:border-blue-200 transition-all shadow-sm">
                                    <div class="text-[9px] font-bold text-blue-500 uppercase tracking-widest mb-1">${c.fields.Anrede || ''} ${c.fields.Rolle || ''}</div>
                                    <div class="text-base font-bold text-slate-800 leading-tight">${c.fields.Vorname || ''} ${c.fields.Title}</div>
                                    <div class="text-[11px] text-slate-400 font-medium mb-4 italic tracking-tight">${c.fields.Funktion || ''}</div>
                                    <div class="space-y-2 text-[10px] text-slate-500 font-medium border-t pt-4">
                                        <div class="flex items-center gap-2">📧 ${c.fields.Email1 || '-'}</div>
                                        <div class="flex items-center gap-2"><span class="text-blue-400">📞</span> ${c.fields.Direktwahl || '-'}</div>
                                        <div class="flex items-center gap-2"><span class="text-green-500">📱</span> ${c.fields.Mobile || '-'}</div>
                                    </div>
                                    <button onclick="deleteContact('${c.id}','${itemId}')" class="absolute top-4 right-4 text-slate-200 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity">✕</button>
                                </div>`).join('') : '<div class="col-span-2 py-10 text-center text-slate-300 text-[10px] font-bold uppercase border-2 border-dashed border-slate-50 rounded-2xl tracking-widest">Keine Kontakte hinterlegt</div>'}
                        </div>
                    </div>
                </div>
            </div>
        </div>`;
}

// --- LOGIK-FUNKTIONEN (EXAKT AUS V2.9) ---

async function updateFirm(id) {
    const fields = { Title: document.getElementById('edit_Title').value, Klassifizierung: document.getElementById('edit_Klass').value, Adresse: document.getElementById('edit_Street').value, Ort: document.getElementById('edit_City').value, Land: document.getElementById('edit_Country').value, Hauptnummer: document.getElementById('edit_Phone').value, VIP: document.getElementById('edit_VIP').checked };
    const token = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${id}/fields`, { method: 'PATCH', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify(fields) });
    loadData();
}

async function saveNewFirm() {
    const name = document.getElementById('new_fName').value; if(!name) return;
    const fields = { Title: name, Klassifizierung: document.getElementById('new_fClass').value, Ort: document.getElementById('new_fCity').value };
    const token = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
    toggleAddForm(); loadData();
}

async function addContact(firmId) {
    const ln = prompt("Nachname (Pflicht):"); if (!ln) return;
    const fields = { Title: ln, Vorname: prompt("Vorname:"), FirmaLookupId: firmId };
    const token = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
    loadData();
}

function filterFirms(q) {
    const query = q.toLowerCase();
    const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query) || f.fields.Ort?.toLowerCase().includes(query));
    const listContainer = document.getElementById('firmList');
    if (listContainer) {
        // Wir nutzen hier direkt den Firmen-Renderer für die Karten
        listContainer.innerHTML = filtered.map(item => {
            const f = item.fields;
            const count = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(item.id)).length;
            return `
            <div onclick="renderDetailPage('${item.id}')" class="bg-white border border-slate-200 p-5 rounded-xl hover:border-blue-400 hover:shadow-md transition-all cursor-pointer flex flex-col h-full border-t-4 ${f.Klassifizierung === 'A' ? 'border-t-emerald-500' : 'border-t-slate-200'}">
                <div class="flex justify-between items-start mb-4">
                    <h3 class="font-medium text-slate-800 text-lg leading-tight">${f.Title || 'Unbenannt'}</h3>
                    ${(f.VIP === true || f.VIP === "true") ? '<span>⭐</span>' : ''}
                </div>
                <div class="text-[11px] text-slate-400 mb-4 italic tracking-wide">📍 ${f.Ort || 'k.A.'}</div>
                <div class="mt-auto pt-4 flex justify-between items-center border-t border-slate-50">
                    <span class="text-[10px] font-bold text-slate-400 uppercase bg-slate-50 px-2 py-1 rounded">${f.Klassifizierung || '-'}</span>
                    <span class="text-[10px] font-semibold text-slate-500 bg-blue-50 px-2.5 py-1 rounded-full flex items-center gap-1">👥 ${count}</span>
                </div>
            </div>`;
        }).join('');
    }
}

async function deleteContact(cid, fid) { if(!confirm("Löschen?")) return; const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken; await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items/${cid}`, { method: 'DELETE', headers: { 'Authorization': `Bearer ${t}` } }); loadData(); }
async function deleteFirm(id, name) { if(!confirm(`Löschen von "${name}"?`)) return; const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken; await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${id}`, { method: 'DELETE', headers: { 'Authorization': `Bearer ${t}` } }); loadData(); }

function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }
function loadFirms() { renderFirms(allFirms); }
function showView(v) { if(v === 'dashboard') location.reload(); if(v === 'contacts') alert("Kommt in V0.43..."); }

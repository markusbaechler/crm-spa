// --- CONFIG & VERSION ---
const appVersion = "0.43";
console.log(`CRM App ${appVersion} - Professional Redesign & UI Fix`);

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
let allFirms = [], allContacts = [], classOptions = []; 
let currentSiteId = "", currentListId = "", contactListId = "";

window.onload = async () => {
    updateFooter(); 
    await msalInstance.handleRedirectPromise();
    checkAuthState();
};

function updateFooter() {
    const ft = document.getElementById('footer-text');
    if (ft) {
        ft.innerHTML = `© 2026 bbz CRM | <span class="text-slate-400 font-normal italic">Etappe E</span> | <span class="font-medium text-slate-600">Build ${appVersion}</span>`;
    }
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
    const content = document.getElementById('main-content'); 
    if (!content) return;
    content.innerHTML = `<div class="p-20 text-center text-slate-400 text-xs uppercase tracking-widest animate-pulse">Synchronisiere Daten...</div>`;

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
        const headers = { 'Authorization': `Bearer ${tokenRes.accessToken}` };

        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers });
        const siteData = await siteRes.json();
        currentSiteId = siteData.id;
        
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers });
        const listsData = await listsRes.json();
        
        currentListId = listsData.value.find(l => l.displayName === "CRMFirms").id;
        contactListId = listsData.value.find(l => l.displayName === "CRMContacts").id;

        const [colRes, firmsRes, contactsRes] = await Promise.all([
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/columns/Klassifizierung`, { headers }),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers }),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items?expand=fields`, { headers })
        ]);
        
        classOptions = (await colRes.json()).choice?.choices || ["-"];
        allFirms = (await firmsRes.json()).value;
        allContacts = (await contactsRes.json()).value;

        renderFirms(allFirms);
    } catch (err) { content.innerHTML = `<div class="p-6 text-red-500 font-bold border rounded-xl bg-red-50 text-center uppercase tracking-widest text-xs italic">Sync-Fehler: ${err.message}</div>`; }
}

// --- VIEW: FIRMENÜBERSICHT ---
function renderFirms(firms) {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-6xl mx-auto animate-in fade-in duration-500">
            <div class="flex justify-between items-end mb-10 border-b border-slate-100 pb-6">
                <div>
                    <h2 class="text-2xl font-semibold text-slate-800 tracking-tight">Firmenstamm</h2>
                    <p class="text-slate-400 text-xs mt-1">Zentrale Verwaltung Ihrer Organisationen</p>
                </div>
                <div class="flex gap-3">
                    <div class="relative">
                        <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suche..." class="pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm outline-none focus:border-blue-400 w-64 transition-colors font-medium">
                        <span class="absolute left-3 top-2.5 text-slate-400">🔍</span>
                    </div>
                    <button onclick="toggleAddForm()" class="bg-slate-800 text-white px-5 py-2 rounded-lg text-sm font-bold shadow-md hover:bg-slate-700 transition">+ FIRMA</button>
                </div>
            </div>

            <div id="addForm" class="hidden mb-10 p-8 bg-white border border-slate-200 rounded-2xl shadow-sm animate-in slide-in-from-top duration-300">
                 <h3 class="text-[10px] font-bold text-slate-400 uppercase mb-6 tracking-widest border-b pb-2">Neue Organisation erfassen</h3>
                 <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div class="space-y-4">
                        <input type="text" id="new_fName" placeholder="Firmenname *" class="w-full p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm outline-none focus:border-blue-400 font-medium italic">
                        <select id="new_fClass" class="w-full p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold italic text-slate-500">
                            <option value="">Klassierung wählen</option>
                            ${classOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
                        </select>
                        <input type="text" id="new_fPhone" placeholder="Hauptnummer" class="w-full p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm">
                    </div>
                    <div class="space-y-4">
                        <input type="text" id="new_fStreet" placeholder="Strasse / Nr." class="w-full p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm italic">
                        <div class="flex gap-3">
                            <input type="text" id="new_fCity" placeholder="Ort" class="flex-1 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold italic">
                            <input type="text" id="new_fCountry" placeholder="CH" value="CH" class="w-16 p-2.5 bg-slate-100 border border-slate-100 rounded-xl text-sm text-center font-bold text-slate-400 uppercase">
                        </div>
                    </div>
                </div>
                <div class="mt-8 flex justify-end gap-3">
                    <button onclick="toggleAddForm()" class="text-xs font-bold text-slate-400 uppercase px-4 hover:underline">Abbrechen</button>
                    <button onclick="saveNewFirm()" class="bg-green-600 text-white px-8 py-2.5 rounded-xl text-[10px] font-bold uppercase tracking-widest shadow-lg hover:bg-green-700">Organisation Speichern</button>
                </div>
            </div>

            <div id="firmList" class="grid grid-cols-1 md:grid-cols-3 gap-6">
                ${firms.map(item => generateFirmCard(item)).join('')}
            </div>
        </div>
    `;
}

function generateFirmCard(item) {
    const f = item.fields;
    const count = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(item.id)).length;
    return `
        <div onclick="renderDetailPage('${item.id}')" class="bg-white border border-slate-200 p-6 rounded-2xl hover:border-blue-400 hover:shadow-xl transition-all cursor-pointer group flex flex-col h-full border-t-4 ${f.Klassifizierung === 'A' ? 'border-t-emerald-500' : 'border-t-slate-200'}">
            <div class="flex justify-between items-start mb-4">
                <h3 class="font-semibold text-slate-800 text-lg group-hover:text-blue-600 transition-colors leading-tight italic">${f.Title || 'Unbenannt'}</h3>
                ${(f.VIP === true || f.VIP === "true") ? '<span class="text-amber-400 text-xl">⭐</span>' : ''}
            </div>
            <div class="text-[11px] text-slate-400 mb-4 italic tracking-wide">📍 ${f.Ort || 'Kein Standort'}</div>
            <div class="mt-auto pt-4 flex justify-between items-center border-t border-slate-50">
                <span class="text-[10px] font-bold text-slate-400 uppercase tracking-widest bg-slate-50 px-2 py-1 rounded border border-slate-100">${f.Klassifizierung || '-'}</span>
                <span class="text-[10px] font-semibold text-slate-500 bg-blue-50 px-2.5 py-1 rounded-full flex items-center gap-1 font-bold italic uppercase tracking-tighter">👥 ${count} Kontakte</span>
            </div>
        </div>`;
}

// --- VIEW: FIRMEN-DETAILSEITE ---
function renderDetailPage(itemId) {
    const firm = allFirms.find(f => String(f.id) === String(itemId));
    const f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(itemId));
    
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-7xl mx-auto animate-in slide-in-from-right duration-500">
            
            <div class="bg-white border border-slate-200 rounded-3xl p-6 mb-8 shadow-sm flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
                <div class="flex items-center gap-5">
                    <button onclick="renderFirms(allFirms)" class="bg-slate-50 hover:bg-slate-100 text-slate-400 p-3 rounded-2xl transition text-xl">←</button>
                    <div>
                        <div class="flex items-center gap-3">
                            <h2 class="text-2xl font-bold text-slate-800 tracking-tight italic uppercase">${f.Title}</h2>
                            ${(f.VIP === true || f.VIP === "true") ? '<span class="text-amber-400 text-xl">⭐</span>' : ''}
                        </div>
                        <div class="flex items-center gap-3 mt-2">
                             <span class="text-slate-400 text-[10px] font-bold uppercase tracking-widest">📍 ${f.Ort || 'Ort n.a.'}</span>
                             <span class="px-2 py-0.5 bg-slate-100 text-slate-500 rounded text-[9px] font-black uppercase border border-slate-200">${f.Land || 'SCHWEIZ'}</span>
                             <span class="mx-1 text-slate-200">|</span>
                             <span class="text-slate-400 text-[10px] font-medium tracking-tight">📞 ${f.Hauptnummer || '071 274 02 40'}</span>
                        </div>
                    </div>
                </div>
                <button onclick="updateFirm('${itemId}')" class="bg-slate-900 text-white px-8 py-3 rounded-2xl text-[10px] font-bold uppercase tracking-widest hover:bg-slate-800 shadow-lg transition-all">Änderungen speichern</button>
            </div>

            <div class="grid grid-cols-12 gap-8 items-start">
                
                <div class="col-span-12 lg:col-span-4 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-3xl p-8 shadow-sm">
                        <h3 class="text-[9px] font-black text-slate-300 uppercase tracking-[0.2em] mb-6 border-b pb-3 text-center italic">Unternehmensdaten</h3>
                        <div class="space-y-5">
                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Offizieller Name</label>
                                <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full mt-1.5 p-3 bg-slate-50 border border-slate-100 rounded-2xl text-sm font-bold italic outline-none focus:border-blue-400">
                            </div>
                            <div class="grid grid-cols-2 gap-4">
                                <div>
                                    <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Klassierung</label>
                                    <select id="edit_Klass" class="w-full mt-1.5 p-3 bg-slate-50 border border-slate-100 rounded-2xl text-sm font-bold italic outline-none">
                                        ${classOptions.map(opt => `<option value="${opt}" ${f.Klassifizierung === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                                    </select>
                                </div>
                                <div class="flex items-end pb-1.5 px-2">
                                    <label class="flex items-center gap-2 cursor-pointer group">
                                        <input type="checkbox" id="edit_VIP" ${(f.VIP === true || f.VIP === "true") ? 'checked' : ''} class="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-0">
                                        <span class="text-[10px] font-bold text-slate-400 uppercase italic">VIP</span>
                                    </label>
                                </div>
                            </div>
                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Telefon Zentrale</label>
                                <input type="text" id="edit_Phone" value="${f.Hauptnummer || '071 274 02 40'}" class="w-full mt-1.5 p-3 bg-slate-50 border border-slate-100 rounded-2xl text-sm outline-none">
                            </div>
                            <div class="pt-2">
                                <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Standortadresse</label>
                                <input type="text" id="edit_Street" value="${f.Adresse || 'Zürcherstrasse 202'}" placeholder="Strasse / Nr." class="w-full mt-1.5 p-3 bg-slate-50 border border-slate-100 rounded-2xl text-sm outline-none mb-3 italic">
                                <div class="flex gap-3">
                                    <input type="text" id="edit_City" value="${f.Ort || 'St. Gallen'}" placeholder="Ort" class="flex-1 p-3 bg-slate-50 border border-slate-100 rounded-2xl text-sm font-bold italic">
                                    <input type="text" id="edit_Country" value="${f.Land || 'SCHWEIZ'}" placeholder="CH" class="w-16 p-3 bg-slate-100 border border-slate-100 rounded-2xl text-xs text-center font-bold text-slate-400 uppercase">
                                </div>
                            </div>
                            <div class="pt-6 border-t border-slate-50 flex justify-center">
                                <button onclick="deleteFirm('${itemId}', '${f.Title}')" class="text-red-400 text-[9px] font-bold uppercase hover:underline opacity-50 hover:opacity-100 transition tracking-widest italic">Organisation löschen</button>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-span-12 lg:col-span-8 space-y-8">
                    
                    <div class="bg-white border border-slate-200 rounded-3xl p-8 shadow-sm">
                        <div class="flex justify-between items-center mb-8 border-b border-slate-50 pb-4">
                            <h3 class="text-[10px] font-black text-slate-300 uppercase tracking-[0.2em] italic">Ansprechpartner (${contacts.length})</h3>
                            <button onclick="addContact('${itemId}')" class="text-blue-600 text-[10px] font-bold uppercase hover:bg-blue-50 px-4 py-2 rounded-xl transition tracking-widest italic">+ NEUER KONTAKT</button>
                        </div>
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                            ${contacts.length > 0 ? contacts.map(c => `
                                <div class="p-6 bg-slate-50/50 border border-slate-100 rounded-3xl relative group hover:bg-white hover:border-blue-200 hover:shadow-lg transition-all duration-300">
                                    <div class="flex-1">
                                        <div class="text-[9px] font-bold text-blue-500 uppercase tracking-widest mb-2 italic">${c.fields.Anrede || ''} ${c.fields.Rolle || ''}</div>
                                        <div class="text-lg font-bold text-slate-800 mb-1 leading-tight italic uppercase">${c.fields.Vorname || ''} ${c.fields.Title}</div>
                                        <div class="text-[11px] text-slate-400 font-medium mb-4 italic tracking-tight">${c.fields.Funktion || ''}</div>
                                        <div class="space-y-2 text-[10px] text-slate-500 font-medium border-t border-slate-100 pt-4">
                                            <div class="flex items-center gap-2"><span class="w-4">📧</span> ${c.fields.Email1 || '-'}</div>
                                            <div class="flex items-center gap-2"><span class="w-4 text-blue-400">📞</span> ${c.fields.Direktwahl || '-'}</div>
                                            <div class="flex items-center gap-2"><span class="w-4 text-green-500">📱</span> ${c.fields.Mobile || '-'}</div>
                                        </div>
                                    </div>
                                    <button onclick="deleteContact('${c.id}','${itemId}')" class="absolute top-4 right-4 text-slate-200 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity p-1">✕</button>
                                </div>
                            `).join('') : '<div class="col-span-2 py-16 text-center text-slate-300 text-[10px] font-bold uppercase border-2 border-dashed border-slate-50 rounded-3xl tracking-widest italic">Keine Kontakte hinterlegt</div>'}
                        </div>
                    </div>

                    <div class="grid grid-cols-2 gap-6 opacity-30 grayscale cursor-not-allowed">
                        <div class="bg-slate-50 border border-slate-100 rounded-3xl p-12 flex flex-col items-center justify-center">
                            <span class="text-2xl mb-2">📜</span>
                            <span class="text-[10px] font-bold text-slate-400 uppercase tracking-widest italic tracking-widest">Aktivitätshistorie</span>
                        </div>
                        <div class="bg-slate-50 border border-slate-100 rounded-3xl p-12 flex flex-col items-center justify-center">
                            <span class="text-2xl mb-2">📅</span>
                            <span class="text-[10px] font-bold text-slate-400 uppercase tracking-widest italic tracking-widest font-black italic">Aufgaben</span>
                        </div>
                    </div>

                </div>
            </div>
        </div>`;
}

// --- LOGIK-FUNKTIONEN (IDENTISCH ZU V0.42) ---
async function updateFirm(id) {
    const fields = { Title: document.getElementById('edit_Title').value, Klassifizierung: document.getElementById('edit_Klass').value, Adresse: document.getElementById('edit_Street').value, Ort: document.getElementById('edit_City').value, Land: document.getElementById('edit_Country').value, Hauptnummer: document.getElementById('edit_Phone').value, VIP: document.getElementById('edit_VIP').checked };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${id}/fields`, { method: 'PATCH', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify(fields) });
    loadData();
}

async function saveNewFirm() {
    const n = document.getElementById('new_fName').value; if(!n) return;
    const f = { Title: n, Klassifizierung: document.getElementById('new_fClass').value, Ort: document.getElementById('new_fCity').value, Hauptnummer: document.getElementById('new_fPhone').value || "", Adresse: document.getElementById('new_fStreet').value || "", Land: document.getElementById('new_fCountry').value || "CH" };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields: f }) });
    toggleAddForm(); loadData();
}

async function addContact(firmId) {
    const ln = prompt("Nachname (Pflicht):"); if (!ln) return;
    const fields = { Title: ln, Vorname: prompt("Vorname:"), FirmaLookupId: firmId };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
    loadData();
}

function filterFirms(q) {
    const query = q.toLowerCase();
    const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query) || f.fields.Ort?.toLowerCase().includes(query));
    document.getElementById('firmList').innerHTML = generateFirmCards(filtered);
}

function generateFirmCards(firms) {
    return firms.map(item => {
        const f = item.fields;
        const count = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(item.id)).length;
        return `
            <div onclick="renderDetailPage('${item.id}')" class="bg-white border border-slate-200 p-6 rounded-3xl hover:border-blue-400 hover:shadow-xl transition-all cursor-pointer group flex flex-col h-full border-t-4 ${f.Klassifizierung === 'A' ? 'border-t-emerald-500' : 'border-t-slate-200'}">
                <div class="flex justify-between items-start mb-4">
                    <h3 class="font-bold text-slate-800 text-lg group-hover:text-blue-600 transition-colors leading-tight italic uppercase">${f.Title || 'Unbenannt'}</h3>
                    ${(f.VIP === true || f.VIP === "true") ? '<span class="text-amber-400 text-xl">⭐</span>' : ''}
                </div>
                <div class="text-[11px] text-slate-400 mb-4 italic tracking-wide">📍 ${f.Ort || 'k.A.'}</div>
                <div class="mt-auto pt-4 flex justify-between items-center border-t border-slate-50">
                    <span class="text-[10px] font-bold text-slate-400 uppercase bg-slate-50 px-2 py-1 rounded border border-slate-100 italic">${f.Klassifizierung || '-'}</span>
                    <span class="text-[10px] font-semibold text-slate-500 bg-blue-50 px-2.5 py-1 rounded-full flex items-center gap-1 font-bold italic uppercase tracking-tighter">👥 ${count} Kontakte</span>
                </div>
            </div>`;
    }).join('');
}

async function deleteContact(cid, fid) { if(!confirm("Löschen?")) return; const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken; await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items/${cid}`, { method: 'DELETE', headers: { 'Authorization': `Bearer ${t}` } }); loadData(); }
async function deleteFirm(id, name) { if(!confirm(`Löschen von "${name}"?`)) return; const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken; await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${id}`, { method: 'DELETE', headers: { 'Authorization': `Bearer ${t}` } }); loadData(); }
function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }
function loadFirms() { renderFirms(allFirms); }

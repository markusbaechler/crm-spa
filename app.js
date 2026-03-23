// --- CONFIG & VERSION ---
const appVersion = "V3.3";
console.log(`CRM App ${appVersion} - Modern Business Redesign`);

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
    const footerText = document.getElementById('footer-text');
    if (footerText) {
        footerText.innerHTML = `© 2026 bbz CRM | <span class="text-slate-500">Status: Etappe E</span> | <span class="font-medium text-blue-600">Build ${appVersion}</span>`;
    }
}

function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    if (accounts.length > 0) {
        authBtn.innerText = "Abmelden";
        authBtn.onclick = () => msalInstance.logoutRedirect({ account: accounts[0] });
        authBtn.className = "bg-slate-100 text-slate-700 px-4 py-2 rounded-lg text-sm font-medium hover:bg-slate-200 transition";
        loadData(); 
    } else {
        authBtn.innerText = "Anmelden";
        authBtn.onclick = () => msalInstance.loginRedirect(loginRequest);
        authBtn.className = "bg-blue-600 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-700 transition shadow-sm";
    }
}

async function loadData() {
    const content = document.getElementById('main-content'); 
    if (!content) return;
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return;
    
    content.innerHTML = `
        <div class="flex flex-col items-center justify-center p-20 text-slate-400">
            <div class="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600 mb-4"></div>
            <p class="text-sm font-medium">Daten werden geladen...</p>
        </div>`;

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: accounts[0] })
            .catch(() => msalInstance.acquireTokenRedirect(loginRequest));
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
    } catch (err) { 
        content.innerHTML = `<div class="p-6 bg-red-50 border border-red-100 text-red-600 rounded-xl text-sm">Fehler beim Datentransfer: ${err.message}</div>`; 
    }
}

function renderFirms(firms) {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-6xl mx-auto">
            <div class="flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4">
                <div>
                    <h2 class="text-2xl font-semibold text-slate-800">Firmenstamm</h2>
                    <p class="text-slate-500 text-sm">Verwalten Sie Ihre Kunden und Partnerorganisationen.</p>
                </div>
                <div class="flex gap-3">
                    <div class="relative">
                        <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suchen..." 
                            class="pl-10 pr-4 py-2 bg-white border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none w-64 transition-all">
                        <span class="absolute left-3 top-2.5 text-slate-400">🔍</span>
                    </div>
                    <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-700 transition flex items-center gap-2">
                        <span>+</span> Neue Firma
                    </button>
                </div>
            </div>

            <div id="detailModal" class="hidden fixed inset-0 bg-slate-900/40 backdrop-blur-[2px] z-50 flex items-center justify-center p-4 transition-all">
                <div class="bg-white rounded-xl shadow-2xl w-full max-w-2xl overflow-hidden relative border border-slate-200">
                    <div id="modalContent"></div>
                </div>
            </div>

            <div id="addForm" class="hidden mb-8 p-6 bg-white border border-blue-100 rounded-xl shadow-sm animate-in slide-in-from-top duration-300">
                <h3 class="text-sm font-semibold text-slate-700 mb-4 flex items-center gap-2">📍 Neue Organisation erfassen</h3>
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
                    <input type="text" id="new_fName" placeholder="Name der Firma" class="p-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-blue-500">
                    <select id="new_fClass" class="p-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-blue-500 text-slate-600">
                        <option value="">Klassifizierung</option>
                        ${classOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
                    </select>
                    <input type="text" id="new_fCity" placeholder="Standort / Ort" class="p-2.5 rounded-lg border border-slate-200 text-sm outline-none focus:border-blue-500">
                </div>
                <div class="flex justify-end gap-2">
                    <button onclick="toggleAddForm()" class="px-4 py-2 text-sm text-slate-500 hover:bg-slate-50 rounded-lg">Abbrechen</button>
                    <button onclick="saveNewFirm()" class="bg-blue-600 text-white px-6 py-2 rounded-lg text-sm font-medium hover:bg-blue-700">Speichern</button>
                </div>
            </div>

            <div id="firmList" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
                ${generateFirmCards(firms)}
            </div>
        </div>
    `;
}

function generateFirmCards(firms) {
    return firms.map(item => {
        const f = item.fields;
        const isVIP = f.VIP === true || f.VIP === "true";
        const count = allContacts.filter(c => c.fields.FirmID === item.id).length;
        const klass = f.Klassifizierung || '-';

        let badgeClass = "bg-slate-100 text-slate-600";
        if (klass.includes('A')) badgeClass = "bg-emerald-50 text-emerald-700 border-emerald-100";
        if (klass.includes('B')) badgeClass = "bg-blue-50 text-blue-700 border-blue-100";

        return `
            <div onclick="openFirmDetails('${item.id}')" class="group bg-white border border-slate-200 p-5 rounded-xl hover:border-blue-300 hover:shadow-md transition-all cursor-pointer flex flex-col justify-between h-full">
                <div>
                    <div class="flex justify-between items-start mb-3">
                        <div class="flex items-center gap-2">
                            <span class="text-xl">🏢</span>
                            <h3 class="font-semibold text-slate-800 group-hover:text-blue-600 transition-colors">${f.Title || 'Unbekannt'}</h3>
                        </div>
                        ${isVIP ? '<span class="text-amber-500" title="Top-Kunde">⭐</span>' : ''}
                    </div>
                    <div class="flex items-center gap-2 text-slate-500 text-xs mb-4">
                        <span>📍 ${f.Ort || 'Kein Standort'}</span>
                    </div>
                </div>
                <div class="flex justify-between items-center pt-4 border-t border-slate-50">
                    <span class="px-2.5 py-1 rounded text-[10px] font-bold uppercase border ${badgeClass}">${klass}</span>
                    <span class="text-[11px] font-medium text-slate-400 bg-slate-50 px-2 py-1 rounded-md">👥 ${count} Kontakte</span>
                </div>
            </div>
        `;
    }).join('');
}

function openFirmDetails(itemId) {
    const firm = allFirms.find(f => f.id === itemId);
    const f = firm.fields;
    const relatedContacts = allContacts.filter(c => c.fields.FirmID === itemId);
    const modal = document.getElementById('detailModal');
    modal.classList.remove('hidden');

    document.getElementById('modalContent').innerHTML = `
        <div class="flex items-center justify-between p-5 border-b border-slate-100 bg-slate-50/50">
            <div class="flex items-center gap-3">
                <span class="text-2xl">🏢</span>
                <h2 class="text-lg font-semibold text-slate-800">${f.Title}</h2>
            </div>
            <button onclick="closeModal()" class="text-slate-400 hover:text-slate-600 text-xl px-2">✕</button>
        </div>
        
        <div class="p-6">
            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div class="space-y-4">
                    <h3 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Stammdaten</h3>
                    <div class="space-y-3">
                        <div>
                            <label class="text-[10px] text-slate-400 font-bold uppercase">Name</label>
                            <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full mt-1 p-2 bg-slate-50 border border-slate-200 rounded text-sm focus:ring-1 focus:ring-blue-500 outline-none">
                        </div>
                        <div class="grid grid-cols-2 gap-3">
                            <div>
                                <label class="text-[10px] text-slate-400 font-bold uppercase">Klassierung</label>
                                <select id="edit_Klass" class="w-full mt-1 p-2 bg-slate-50 border border-slate-200 rounded text-sm outline-none">
                                    ${classOptions.map(opt => `<option value="${opt}" ${f.Klassifizierung === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                                </select>
                            </div>
                            <div class="flex items-end pb-2">
                                <label class="flex items-center gap-2 cursor-pointer">
                                    <input type="checkbox" id="edit_VIP" ${f.VIP === true || f.VIP === "true" ? 'checked' : ''} class="rounded text-blue-600">
                                    <span class="text-xs font-medium text-slate-600">VIP Status</span>
                                </label>
                            </div>
                        </div>
                        <div>
                            <label class="text-[10px] text-slate-400 font-bold uppercase">Adresse & Ort</label>
                            <div class="flex gap-2 mt-1">
                                <input type="text" id="edit_Street" value="${f.Adresse || ''}" placeholder="Strasse" class="flex-1 p-2 bg-slate-50 border border-slate-200 rounded text-sm outline-none">
                                <input type="text" id="edit_City" value="${f.Ort || ''}" placeholder="Ort" class="flex-1 p-2 bg-slate-50 border border-slate-200 rounded text-sm outline-none">
                            </div>
                        </div>
                        <button onclick="updateFirm('${itemId}')" class="w-full bg-blue-600 text-white py-2.5 rounded-lg text-sm font-medium hover:bg-blue-700 transition shadow-sm mt-2">Änderungen speichern</button>
                        <button onclick="deleteFirm('${itemId}', '${f.Title}')" class="w-full text-red-500 text-[10px] font-bold uppercase hover:bg-red-50 py-2 rounded transition">Firma löschen</button>
                    </div>
                </div>

                <div class="border-l border-slate-100 pl-6">
                    <div class="flex justify-between items-center mb-4">
                        <h3 class="text-xs font-bold text-slate-400 uppercase tracking-wider">Ansprechpartner</h3>
                        <button onclick="addContact('${itemId}')" class="text-blue-600 text-[10px] font-bold uppercase hover:underline">+ Hinzufügen</button>
                    </div>
                    <div class="space-y-2 max-h-[300px] overflow-y-auto pr-2">
                        ${relatedContacts.length > 0 ? relatedContacts.map(c => `
                            <div class="p-3 bg-slate-50 rounded-lg border border-slate-100 flex justify-between items-center group">
                                <div>
                                    <div class="text-sm font-semibold text-slate-700">${c.fields.FirstName || ''} ${c.fields.Title}</div>
                                    <div class="text-[10px] text-slate-400">${c.fields.Email || 'Keine E-Mail'}</div>
                                </div>
                                <button onclick="deleteContact('${c.id}', '${itemId}')" class="text-slate-300 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity">✕</button>
                            </div>
                        `).join('') : '<div class="py-8 text-center text-slate-400 text-xs italic">Keine Kontakte gefunden</div>'}
                    </div>
                </div>
            </div>
        </div>
    `;
}

// --- LOGIK & FILTER (IDENTISCH ZU V3.2) ---
async function updateFirm(itemId) {
    const fields = {
        Title: document.getElementById('edit_Title').value,
        Klassifizierung: document.getElementById('edit_Klass').value,
        Adresse: document.getElementById('edit_Street').value,
        Ort: document.getElementById('edit_City').value,
        VIP: document.getElementById('edit_VIP').checked
    };
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${itemId}/fields`, {
        method: 'PATCH', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(fields)
    });
    closeModal(); loadData();
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
    const ln = prompt("Nachname:"); if (!ln) return;
    const fields = { Title: ln, FirstName: prompt("Vorname:"), Email: prompt("E-Mail:"), FirmID: firmId };
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items`, {
        method: 'POST', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields: fields })
    });
    loadData(); closeModal();
}

async function deleteContact(cId, firmId) {
    if(!confirm("Ansprechpartner wirklich entfernen?")) return;
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items/${cId}`, {
        method: 'DELETE', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
    });
    loadData(); closeModal();
}

async function deleteFirm(itemId, name) {
    if(!confirm(`Möchten Sie "${name}" wirklich unwiderruflich löschen?`)) return;
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${itemId}`, {
        method: 'DELETE', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
    });
    closeModal(); loadData();
}

function filterFirms(q) {
    const query = q.toLowerCase();
    const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query) || f.fields.Ort?.toLowerCase().includes(query));
    document.getElementById('firmList').innerHTML = generateFirmCards(filtered);
}
function closeModal() { document.getElementById('detailModal').classList.add('hidden'); }
function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }
function loadFirms() { renderFirms(allFirms); }
function loadAllContacts() {
    const content = document.getElementById('main-content');
    content.innerHTML = `<div class="max-w-4xl mx-auto">
        <h2 class="text-2xl font-semibold text-slate-800 mb-6">Alle Kontakte</h2>
        <div class="bg-white border border-slate-200 rounded-xl overflow-hidden">
            <table class="w-full text-left border-collapse text-sm">
                <thead class="bg-slate-50 border-b border-slate-200">
                    <tr><th class="p-4 font-bold text-slate-500 uppercase text-[10px]">Name</th><th class="p-4 font-bold text-slate-500 uppercase text-[10px]">E-Mail</th><th class="p-4 font-bold text-slate-500 uppercase text-[10px]">Firma-ID</th></tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                    ${allContacts.map(c => `<tr class="hover:bg-slate-50 transition-colors"><td class="p-4 font-medium text-slate-700">${c.fields.FirstName || ''} ${c.fields.Title}</td><td class="p-4 text-slate-500">${c.fields.Email || '-'}</td><td class="p-4 font-mono text-slate-300 text-[10px]">${c.fields.FirmID || '-'}</td></tr>`).join('')}
                </tbody>
            </table>
        </div>
    </div>`;
}

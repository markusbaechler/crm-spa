// --- CONFIG & VERSION ---
const appVersion = "V2.9";
console.log(`CRM App ${appVersion} - Dynamische SharePoint-Metadaten`);

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
let classOptions = []; // Hier speichern wir die dynamischen Werte aus SharePoint
let currentSiteId = "";
let currentListId = "";

window.onload = async () => {
    updateFooter(); 
    await msalInstance.handleRedirectPromise();
    checkAuthState();
};

function updateFooter() {
    const footerText = document.getElementById('footer-text');
    if (footerText) {
        footerText.innerHTML = `© 2026 bbz CRM Light | Status: Etappe D | <span class="font-black text-slate-600">Version: ${appVersion}</span>`;
    }
}

function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    if (accounts.length > 0) {
        authBtn.innerText = "Logout";
        authBtn.onclick = () => msalInstance.logoutRedirect({ account: accounts[0] });
        authBtn.classList.replace('bg-blue-600', 'bg-red-600');
        loadData(); // Lädt jetzt Firmen UND Metadaten
    } else {
        authBtn.innerText = "Login";
        authBtn.onclick = () => msalInstance.loginRedirect(loginRequest);
    }
}

async function loadData() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return;
    content.innerHTML = '<p class="p-10 text-center animate-pulse text-blue-600 font-bold uppercase tracking-widest text-xs">Synchronisiere Metadaten...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: accounts[0] })
            .catch(() => msalInstance.acquireTokenRedirect(loginRequest));
        const headers = { 'Authorization': `Bearer ${tokenRes.accessToken}` };

        // 1. Site & List IDs holen
        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers });
        const siteData = await siteRes.json();
        currentSiteId = siteData.id;
        
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers });
        const listsData = await listsRes.json();
        const targetList = listsData.value.find(l => l.displayName === "CRMFirms");
        currentListId = targetList.id;

        // 2. DYNAMISCH: Klassifizierungs-Optionen direkt aus SharePoint Spalte holen
        const columnRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/columns/Klassifizierung`, { headers });
        const columnData = await columnRes.json();
        // Wir speichern die Auswahlmöglichkeiten (Choice) in unserem Array
        classOptions = columnData.choice ? columnData.choice.choices : ["leer"];

        // 3. Firmen laden
        const itemsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers });
        const itemsData = await itemsRes.json();
        allFirms = itemsData.value;

        renderUI();
    } catch (err) { content.innerHTML = `<div class="p-4 bg-red-50 text-red-700">Fehler: ${err.message}</div>`; }
}

function renderUI() {
    const content = document.getElementById('app-content');
    content.innerHTML = `
        <div class="bg-white p-6 rounded-3xl shadow-xl border border-slate-100 relative">
            <div class="flex justify-between items-center mb-8">
                <h2 class="text-3xl font-black text-slate-800 italic uppercase tracking-tighter italic">🏢 Firmen</h2>
                <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-6 py-2 rounded-full font-bold shadow-lg transform hover:scale-105 transition">+ NEU</button>
            </div>

            <div id="detailModal" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
                <div class="bg-white rounded-3xl shadow-2xl w-full max-w-lg p-8 relative">
                    <button onclick="closeModal()" class="absolute top-4 right-4 text-slate-400 hover:text-slate-600 font-bold">✕</button>
                    <div id="modalContent"></div>
                </div>
            </div>

            <div id="addForm" class="hidden mb-8 p-6 bg-slate-50 rounded-2xl border-2 border-white shadow-inner">
                <div class="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
                    <input type="text" id="new_fName" placeholder="Name" class="p-3 rounded-xl border-none shadow-sm font-bold">
                    <select id="new_fClass" class="p-3 rounded-xl border-none shadow-sm font-bold text-slate-500">
                        ${classOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
                    </select>
                </div>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-3 mb-4 items-center">
                    <input type="text" id="new_fCity" placeholder="Ort" class="p-3 rounded-xl border-none shadow-sm">
                    <label class="flex items-center space-x-3 p-3"><input type="checkbox" id="new_fVIP" class="w-5 h-5 rounded border-none shadow-sm text-blue-600"><span class="text-slate-600 font-bold uppercase text-xs">VIP</span></label>
                </div>
                <button onclick="saveNewFirm()" class="bg-green-600 text-white px-8 py-2 rounded-xl font-bold">SPEICHERN</button>
            </div>

            <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suchen..." class="w-full p-4 mb-6 rounded-2xl bg-slate-50 border-none shadow-inner focus:ring-2 focus:ring-blue-500 text-lg">

            <div id="firmList" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                ${generateFirmCards(allFirms)}
            </div>
        </div>
    `;
}

function generateFirmCards(firms) {
    return firms.map(item => {
        const f = item.fields;
        const isVIP = f.VIP === true || f.VIP === "true";
        const klass = f.Klassifizierung || 'leer';

        let colorStyle = "text-slate-400 bg-slate-100 border-slate-200";
        if (klass.startsWith('A')) colorStyle = "text-emerald-600 bg-emerald-50 border border-emerald-100";
        if (klass.startsWith('B')) colorStyle = "text-blue-600 bg-blue-50 border border-blue-100";
        if (klass.startsWith('C')) colorStyle = "text-orange-600 bg-orange-50 border border-orange-100";

        return `
            <div onclick="openFirmDetails('${item.id}')" class="p-5 bg-slate-50 border border-white rounded-3xl shadow-sm hover:shadow-xl transition-all cursor-pointer group relative">
                <div class="flex justify-between items-start mb-2">
                    <span class="font-bold text-slate-700 text-lg group-hover:text-blue-600 leading-tight">${f.Title || 'Unbenannt'}</span>
                    <div class="flex items-center space-x-1">
                        ${isVIP ? '<span class="text-amber-500">👑</span>' : ''}
                        <span class="px-2 py-1 rounded-lg text-[10px] font-black shadow-sm uppercase border ${colorStyle}">${klass}</span>
                    </div>
                </div>
                <div class="text-[11px] text-slate-400 font-medium">${f.Ort || 'Kein Ort'} · ${f.Land || 'CH'}</div>
            </div>
        `;
    }).join('');
}

function openFirmDetails(itemId) {
    const firm = allFirms.find(f => f.id === itemId);
    const f = firm.fields;
    const isVIP = f.VIP === true || f.VIP === "true";
    const modal = document.getElementById('detailModal');
    const container = document.getElementById('modalContent');
    
    modal.classList.remove('hidden');
    container.innerHTML = `
        <h2 class="text-2xl font-black text-slate-800 mb-6 uppercase tracking-tighter italic">Details bearbeiten</h2>
        <div class="space-y-4">
            <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full p-3 bg-slate-100 rounded-xl border-none font-bold shadow-sm">
            
            <div class="grid grid-cols-2 gap-4">
                <select id="edit_Klass" class="w-full p-3 bg-slate-100 rounded-xl border-none font-bold shadow-sm">
                    ${classOptions.map(opt => `<option value="${opt}" ${f.Klassifizierung === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                </select>
                <label class="flex items-center space-x-2 bg-slate-100 p-3 rounded-xl cursor-pointer shadow-sm">
                    <input type="checkbox" id="edit_VIP" ${isVIP ? 'checked' : ''} class="w-4 h-4 rounded text-blue-600 border-none"><span class="text-xs font-bold text-slate-500">VIP</span>
                </label>
            </div>

            <div class="grid grid-cols-2 gap-4">
                <input type="text" id="edit_Street" value="${f.Adresse || ''}" placeholder="Strasse" class="w-full p-3 bg-slate-100 rounded-xl border-none shadow-sm">
                <input type="text" id="edit_Country" value="${f.Land || 'Schweiz'}" class="w-full p-3 bg-slate-100 rounded-xl border-none shadow-sm">
            </div>

            <div class="grid grid-cols-3 gap-4">
                <input type="text" id="edit_Zip" value="${f.PLZ || ''}" placeholder="PLZ" class="p-3 bg-slate-100 rounded-xl border-none shadow-sm">
                <input type="text" id="edit_City" value="${f.Ort || ''}" placeholder="Ort" class="col-span-2 p-3 bg-slate-100 rounded-xl border-none shadow-sm">
            </div>

            <div class="mt-8 pt-6 border-t flex justify-between items-center text-sm">
                <button onclick="updateFirm('${itemId}')" class="bg-blue-600 text-white px-8 py-3 rounded-2xl font-black shadow-lg hover:bg-blue-700 transition">SPEICHERN</button>
                <button onclick="deleteFirm('${itemId}', '${f.Title}')" class="text-red-400 hover:text-red-600 font-bold p-2">Löschen</button>
            </div>
        </div>
    `;
}

// REST-LOGIK (Update, Save, Delete)
async function updateFirm(itemId) {
    const fields = {
        Title: document.getElementById('edit_Title').value,
        Klassifizierung: document.getElementById('edit_Klass').value,
        Adresse: document.getElementById('edit_Street').value,
        PLZ: document.getElementById('edit_Zip').value,
        Ort: document.getElementById('edit_City').value,
        Land: document.getElementById('edit_Country').value,
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
    if(!name) return alert("Name fehlt!");
    const fields = {
        Title: name,
        Klassifizierung: document.getElementById('new_fClass').value,
        Adresse: document.getElementById('new_fStreet').value,
        PLZ: document.getElementById('new_fZip').value,
        Ort: document.getElementById('new_fCity').value,
        Land: document.getElementById('new_fCountry').value,
        VIP: document.getElementById('new_fVIP').checked
    };
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items`, {
        method: 'POST', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields: fields })
    });
    toggleAddForm(); loadData();
}

async function deleteFirm(itemId, name) {
    if(!confirm(`Firma "${name}" wirklich löschen?`)) return;
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${itemId}`, {
        method: 'DELETE', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
    });
    closeModal(); loadData();
}

function filterFirms(q) {
    const query = q.toLowerCase();
    const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query) || f.fields.Ort?.toLowerCase().includes(query) || f.fields.Klassifizierung?.toLowerCase().includes(query));
    document.getElementById('firmList').innerHTML = generateFirmCards(filtered);
}

function closeModal() { document.getElementById('detailModal').classList.add('hidden'); }
function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }
function showView(v) { if(v === 'dashboard') location.reload(); if(v === 'firms') loadData(); }

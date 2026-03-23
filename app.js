// --- CONFIG & VERSION ---
const appVersion = "V2.6";
console.log(`CRM App ${appVersion} - Detail & Edit Modus`);

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
        loadFirms();
    } else {
        authBtn.innerText = "Login";
        authBtn.onclick = () => msalInstance.loginRedirect(loginRequest);
    }
}

async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return;

    content.innerHTML = '<p class="p-10 text-center animate-pulse text-blue-600 font-bold">Lade Daten...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: accounts[0] })
            .catch(() => msalInstance.acquireTokenRedirect(loginRequest));

        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` } });
        const siteData = await siteRes.json();
        currentSiteId = siteData.id;
        
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` } });
        const listsData = await listsRes.json();
        const targetList = listsData.value.find(l => l.displayName === "CRMFirms");
        currentListId = targetList.id;

        const itemsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` } });
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
                <h2 class="text-3xl font-black text-slate-800 italic uppercase">🏢 Firmen</h2>
                <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-6 py-2 rounded-full font-bold shadow-lg transform hover:scale-105 transition">+ NEU</button>
            </div>

            <div id="detailModal" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
                <div class="bg-white rounded-3xl shadow-2xl w-full max-w-lg p-8 relative">
                    <button onclick="closeModal()" class="absolute top-4 right-4 text-slate-400 hover:text-slate-600 font-bold">✕</button>
                    <div id="modalContent"></div>
                </div>
            </div>

            <div id="addForm" class="hidden mb-8 p-6 bg-slate-50 rounded-2xl border-2 border-white shadow-inner">
                <input type="text" id="new_fName" placeholder="Name" class="w-full p-3 mb-3 rounded-xl border-none shadow-sm">
                <div class="grid grid-cols-2 gap-3 mb-3">
                    <select id="new_fClass" class="p-3 rounded-xl border-none shadow-sm font-bold text-slate-500">
                        <option value="leer">leer</option><option value="A">A-Kunde</option><option value="B">B-Kunde</option><option value="C">C-Kunde</option>
                    </select>
                    <input type="text" id="new_fCity" placeholder="Ort" class="p-3 rounded-xl border-none shadow-sm">
                </div>
                <button onclick="saveNewFirm()" class="bg-green-600 text-white px-6 py-2 rounded-xl font-bold">Anlegen</button>
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
        return `
            <div onclick="openFirmDetails('${item.id}')" class="p-5 bg-slate-50 border border-white rounded-3xl shadow-sm hover:shadow-xl transition-all cursor-pointer group">
                <div class="flex justify-between items-start mb-2">
                    <span class="font-bold text-slate-700 text-lg group-hover:text-blue-600">${f.Title || 'Unbenannt'}</span>
                    <span class="px-2 py-1 rounded-lg text-[10px] font-black shadow-sm uppercase bg-white text-blue-500">${f.Klassifizierung || 'leer'}</span>
                </div>
                <div class="text-[11px] text-slate-400">${f.Ort || 'Kein Ort'}</div>
            </div>
        `;
    }).join('');
}

// --- DETAIL & EDIT LOGIK ---

function openFirmDetails(itemId) {
    const firm = allFirms.find(f => f.id === itemId);
    const f = firm.fields;
    const modal = document.getElementById('detailModal');
    const container = document.getElementById('modalContent');
    
    modal.classList.remove('hidden');
    container.innerHTML = `
        <h2 class="text-2xl font-black text-slate-800 mb-6 uppercase tracking-tighter italic">Firma bearbeiten</h2>
        <div class="space-y-4">
            <div><label class="text-[10px] font-bold text-slate-400 uppercase ml-2">Firmenname</label>
                 <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full p-3 bg-slate-100 rounded-xl border-none focus:ring-2 focus:ring-blue-500 font-bold"></div>
            
            <div class="grid grid-cols-2 gap-4">
                <div><label class="text-[10px] font-bold text-slate-400 uppercase ml-2">Klassifizierung</label>
                <select id="edit_Klass" class="w-full p-3 bg-slate-100 rounded-xl border-none font-bold">
                    <option value="leer" ${f.Klassifizierung === 'leer' ? 'selected' : ''}>leer</option>
                    <option value="A" ${f.Klassifizierung === 'A' ? 'selected' : ''}>A-Kunde</option>
                    <option value="B" ${f.Klassifizierung === 'B' ? 'selected' : ''}>B-Kunde</option>
                    <option value="C" ${f.Klassifizierung === 'C' ? 'selected' : ''}>C-Kunde</option>
                </select></div>
                <div><label class="text-[10px] font-bold text-slate-400 uppercase ml-2">Hauptnummer</label>
                <input type="text" id="edit_Phone" value="${f.Hauptnummer || ''}" class="w-full p-3 bg-slate-100 rounded-xl border-none font-bold"></div>
            </div>

            <div><label class="text-[10px] font-bold text-slate-400 uppercase ml-2">Strasse</label>
                 <input type="text" id="edit_Street" value="${f.Adresse || ''}" class="w-full p-3 bg-slate-100 rounded-xl border-none font-bold"></div>

            <div class="grid grid-cols-3 gap-4">
                <div class="col-span-1"><label class="text-[10px] font-bold text-slate-400 uppercase ml-2">PLZ</label>
                     <input type="text" id="edit_Zip" value="${f.PLZ || ''}" class="w-full p-3 bg-slate-100 rounded-xl border-none font-bold"></div>
                <div class="col-span-2"><label class="text-[10px] font-bold text-slate-400 uppercase ml-2">Ort</label>
                     <input type="text" id="edit_City" value="${f.Ort || ''}" class="w-full p-3 bg-slate-100 rounded-xl border-none font-bold"></div>
            </div>

            <div class="mt-8 pt-6 border-t flex justify-between items-center">
                <button onclick="updateFirm('${itemId}')" class="bg-blue-600 text-white px-8 py-3 rounded-2xl font-black shadow-lg hover:bg-blue-700 transition">SPEICHERN</button>
                <button onclick="deleteFirm('${itemId}', '${f.Title}')" class="text-red-400 hover:text-red-600 transition p-2">
                    <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                    </svg>
                </button>
            </div>
        </div>
    `;
}

async function updateFirm(itemId) {
    const fields = {
        Title: document.getElementById('edit_Title').value,
        Klassifizierung: document.getElementById('edit_Klass').value,
        Adresse: document.getElementById('edit_Street').value,
        PLZ: document.getElementById('edit_Zip').value,
        Ort: document.getElementById('edit_City').value,
        Hauptnummer: document.getElementById('edit_Phone').value
    };

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
        const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${itemId}/fields`, {
            method: 'PATCH',
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
            body: JSON.stringify(fields)
        });

        if(res.ok) { closeModal(); loadFirms(); }
    } catch (err) { alert("Fehler beim Update: " + err.message); }
}

async function deleteFirm(itemId, name) {
    if(!confirm(`Firma "${name}" wirklich löschen?`)) return;
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${itemId}`, {
        method: 'DELETE', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
    });
    closeModal(); loadFirms();
}

function filterFirms(q) {
    const query = q.toLowerCase();
    const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query) || f.fields.Ort?.toLowerCase().includes(query));
    document.getElementById('firmList').innerHTML = generateFirmCards(filtered);
}

function closeModal() { document.getElementById('detailModal').classList.add('hidden'); }
function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }
function showView(v) { if(v === 'dashboard') location.reload(); if(v === 'firms') loadFirms(); }

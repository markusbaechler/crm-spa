// --- CONFIG & VERSION ---
const appVersion = "V2.3";
console.log(`CRM App ${appVersion} wird geladen...`);

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

// --- INITIALISIERUNG ---
window.onload = async () => {
    updateFooter(); 
    await msalInstance.handleRedirectPromise();
    checkAuthState();
};

function updateFooter() {
    const footer = document.querySelector('footer p');
    if (footer) {
        footer.innerHTML = `© 2026 bbz CRM Light | Status: Etappe D | <strong>Version: ${appVersion}</strong>`;
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

// --- DATEN LADEN ---
async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return;

    content.innerHTML = '<p class="p-10 text-center animate-pulse text-blue-600 font-bold">Synchronisiere Daten...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: accounts[0] })
            .catch(() => msalInstance.acquireTokenRedirect(loginRequest));

        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, {
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
        });
        const siteData = await siteRes.json();
        
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists`, {
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
        });
        const listsData = await listsRes.json();
        const targetList = listsData.value.find(l => l.displayName === "CRMFirms");

        const itemsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists/${targetList.id}/items?expand=fields`, {
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
        });
        const itemsData = await itemsRes.json();
        allFirms = itemsData.value;

        renderUI(siteData.id, targetList.id);

    } catch (err) {
        content.innerHTML = `<div class="p-4 bg-red-50 text-red-700">Fehler: ${err.message}</div>`;
    }
}

// --- UI RENDERING ---
function renderUI(siteId, listId) {
    const content = document.getElementById('app-content');
    content.innerHTML = `
        <div class="bg-white p-6 rounded-3xl shadow-xl border border-slate-100">
            <div class="flex justify-between items-center mb-8">
                <div>
                    <h2 class="text-3xl font-black text-slate-800 tracking-tighter italic uppercase">🏢 Firmen</h2>
                    <p class="text-slate-400 text-xs font-bold uppercase tracking-widest mt-1 italic">${allFirms.length} Einträge</p>
                </div>
                <button onclick="toggleForm()" class="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-full font-bold shadow-lg transition-all transform hover:scale-105">
                    + NEUE FIRMA
                </button>
            </div>

            <div id="addForm" class="hidden mb-8 p-6 bg-slate-50 rounded-2xl border-2 border-white shadow-inner">
                <div class="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
                    <input type="text" id="fName" placeholder="Firmenname" class="p-3 rounded-xl border-none shadow-sm focus:ring-2 focus:ring-blue-500">
                    <select id="fClass" class="p-3 rounded-xl border-none shadow-sm font-bold text-slate-600">
                        <option value="leer" selected>leer</option>
                        <option value="A">A-Kunde</option>
                        <option value="B">B-Kunde</option>
                        <option value="C">C-Kunde</option>
                    </select>
                </div>
                <div class="grid grid-cols-1 md:grid-cols-3 gap-3">
                    <input type="text" id="fStreet" placeholder="Strasse" class="p-3 rounded-xl border-none shadow-sm">
                    <input type="text" id="fZip" placeholder="PLZ" class="p-3 rounded-xl border-none shadow-sm">
                    <input type="text" id="fCity" placeholder="Ort" class="p-3 rounded-xl border-none shadow-sm">
                </div>
                <div class="mt-4 flex gap-2">
                    <button onclick="saveFirm('${siteId}', '${listId}')" class="bg-green-600 text-white px-6 py-2 rounded-xl font-bold hover:bg-green-700">SPEICHERN</button>
                    <button onclick="toggleForm()" class="text-slate-400 px-4 font-bold">Abbrechen</button>
                </div>
            </div>

            <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suchen nach Name oder Ort..." 
                class="w-full p-4 mb-6 rounded-2xl bg-slate-50 border-none shadow-inner focus:ring-2 focus:ring-blue-500 text-lg">

            <div id="firmList" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                ${generateFirmCards(allFirms)}
            </div>
        </div>
    `;
}

function generateFirmCards(firms) {
    return firms.map(item => {
        const f = item.fields;
        const rawClass = f.Klassifizierung || 'leer';
        
        let displayClass = rawClass;
        if (rawClass === 'A') displayClass = 'A-Kunde';
        if (rawClass === 'B') displayClass = 'B-Kunde';
        if (rawClass === 'C') displayClass = 'C-Kunde';

        let colorStyle = "text-slate-400 bg-slate-100 border-slate-200";
        if (displayClass.startsWith('A')) colorStyle = "text-emerald-600 bg-emerald-50 border border-emerald-100";
        if (displayClass.startsWith('B')) colorStyle = "text-blue-600 bg-blue-50 border border-blue-100";
        if (displayClass.startsWith('C')) colorStyle = "text-orange-600 bg-orange-50 border border-orange-100";

        return `
            <div class="p-5 bg-slate-50 border border-white rounded-3xl shadow-sm hover:shadow-md transition-all group">
                <div class="flex justify-between items-start mb-2">
                    <span class="font-bold text-slate-700 text-lg group-hover:text-blue-600 transition-colors">${f.Title || 'Unbenannt'}</span>
                    <span class="px-2 py-1 rounded-lg text-[10px] font-black shadow-sm italic uppercase border ${colorStyle}">
                        ${displayClass}
                    </span>
                </div>
                <div class="text-[11px] text-slate-400 font-medium">
                    ${f.Adresse ? f.Adresse + '<br>' : ''}
                    ${f.PLZ || ''} ${f.Ort || ''}
                </div>
            </div>
        `;
    }).join('');
}

function filterFirms(query) {
    const q = query.toLowerCase();
    const filtered = allFirms.filter(f => 
        f.fields.Title?.toLowerCase().includes(q) || 
        f.fields.Ort?.toLowerCase().includes(q)
    );
    document.getElementById('firmList').innerHTML = generateFirmCards(filtered);
}

function toggleForm() { document.getElementById('addForm').classList.toggle('hidden'); }

async function saveFirm(siteId, listId) {
    const name = document.getElementById('fName').value;
    const klasse = document.getElementById('fClass').value;
    const street = document.getElementById('fStreet').value;
    const zip = document.getElementById('fZip').value;
    const city = document.getElementById('fCity').value;

    if(!name) return alert("Name fehlt!");

    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields: { 
            Title: name, 
            Klassifizierung: klasse,
            Adresse: street, PLZ: zip, Ort: city, Land: "Schweiz" 
        } })
    });

    toggleForm(); 
    loadFirms(); 
}

function showView(v) { if(v === 'dashboard') location.reload(); if(v === 'firms') loadFirms(); }

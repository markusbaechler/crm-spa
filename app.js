// VERSIONSMARKER: V2.0 - MODERNE ANSICHT MIT SUCHE
console.log("CRM App V2.0 wird geladen...");

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

window.onload = async () => {
    await msalInstance.handleRedirectPromise();
    checkAuthState();
};

function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    if (accounts.length > 0) {
        authBtn.innerText = "Logout";
        authBtn.onclick = () => msalInstance.logoutRedirect({ account: accounts[0] });
        authBtn.classList.replace('bg-blue-600', 'bg-red-600');
        // Automatisches Laden beim Start, wenn eingeloggt
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

    content.innerHTML = '<p class="p-10 text-center animate-pulse text-blue-600 font-bold">Synchronisiere mit SharePoint...</p>';

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

function renderUI(siteId, listId) {
    const content = document.getElementById('app-content');
    content.innerHTML = `
        <div class="bg-white p-6 rounded-3xl shadow-xl border border-slate-100">
            <div class="flex justify-between items-center mb-8">
                <div>
                    <h2 class="text-3xl font-black text-slate-800 tracking-tighter italic">🏢 FIRMEN</h2>
                    <p class="text-slate-400 text-xs font-bold uppercase tracking-widest mt-1">bbz CRM System</p>
                </div>
                <button onclick="toggleForm()" class="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-full font-bold shadow-lg transition-all transform hover:scale-105">
                    + NEUE FIRMA
                </button>
            </div>

            <div id="addForm" class="hidden mb-8 p-6 bg-slate-50 rounded-2xl border-2 border-white shadow-inner">
                <input type="text" id="fName" placeholder="Name der Firma" class="w-full p-3 mb-3 rounded-xl border-none shadow-sm focus:ring-2 focus:ring-blue-500 text-lg">
                <div class="flex gap-2">
                    <select id="fClass" class="flex-1 p-3 rounded-xl border-none shadow-sm font-bold text-slate-600">
                        <option value="A">Klasse A</option>
                        <option value="B">Klasse B</option>
                        <option value="C">Klasse C</option>
                    </select>
                    <button onclick="saveFirm('${siteId}', '${listId}')" class="bg-green-600 text-white px-6 py-2 rounded-xl font-bold hover:bg-green-700">SPEICHERN</button>
                </div>
            </div>

            <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suchen..." 
                class="w-full p-4 mb-6 rounded-2xl bg-slate-50 border-none shadow-inner focus:ring-2 focus:ring-blue-500 text-lg">

            <div id="firmList" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                ${generateFirmCards(allFirms)}
            </div>
        </div>
    `;
}

function generateFirmCards(firms) {
    return firms.map(item => `
        <div class="p-5 bg-slate-50 border border-white rounded-3xl shadow-sm hover:shadow-md transition-all group">
            <div class="flex justify-between items-start">
                <span class="font-bold text-slate-700 text-lg group-hover:text-blue-600 transition-colors">${item.fields.Title || 'Unbenannt'}</span>
                <span class="px-2 py-1 bg-white rounded-lg text-[10px] font-black shadow-sm text-blue-500 italic uppercase">${item.fields.Klassifizierung || '-'}</span>
            </div>
        </div>
    `).join('');
}

function filterFirms(query) {
    const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query.toLowerCase()));
    document.getElementById('firmList').innerHTML = generateFirmCards(filtered);
}

function toggleForm() { document.getElementById('addForm').classList.toggle('hidden'); }

async function saveFirm(siteId, listId) {
    const name = document.getElementById('fName').value;
    const klasse = document.getElementById('fClass').value;
    if(!name) return;

    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`, {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields: { Title: name, Klassifizierung: klasse } })
    });

    if(response.ok) { toggleForm(); loadFirms(); }
}

function showView(v) { if(v === 'dashboard') location.reload(); if(v === 'firms') loadFirms(); }

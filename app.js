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

let allFirms = []; // Speicher für die Suche

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
    } else {
        authBtn.innerText = "Login";
        authBtn.onclick = () => msalInstance.loginRedirect(loginRequest);
    }
}

async function loadFirms() {
    const content = document.getElementById('app-content');
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
        content.innerHTML = '<div class="p-10 text-center text-slate-400 font-bold">Bitte erst einloggen.</div>';
        return;
    }

    content.innerHTML = '<p class="p-10 text-center animate-pulse text-blue-600 font-bold">Lade CRM-Daten...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: accounts[0] })
            .catch(() => msalInstance.acquireTokenRedirect(loginRequest));

        // Site & Liste finden
        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, {
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
        });
        const siteData = await siteRes.json();
        
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/lists`, {
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
        });
        const listsData = await listsRes.json();
        const targetList = listsData.value.find(l => l.displayName === "CRMFirms");

        // Daten holen
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
        <div class="flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4">
            <h2 class="text-3xl font-black text-slate-800 tracking-tighter italic">🏢 FIRMEN</h2>
            <button onclick="toggleForm()" class="bg-green-600 hover:bg-green-700 text-white px-6 py-2 rounded-full font-bold shadow-lg transition-all transform hover:scale-105">
                + NEU
            </button>
        </div>

        <div id="addForm" class="hidden mb-8 p-6 bg-slate-100 rounded-3xl border-2 border-white shadow-inner">
            <h3 class="font-bold mb-4 text-slate-600 uppercase text-xs tracking-widest">Neue Firma erfassen</h3>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
                <input type="text" id="fName" placeholder="Name der Firma" class="p-3 rounded-xl border-none shadow-sm focus:ring-2 focus:ring-blue-500">
                <select id="fClass" class="p-3 rounded-xl border-none shadow-sm">
                    <option value="A">Klasse A (Key Account)</option>
                    <option value="B">Klasse B (Aktiv)</option>
                    <option value="C" selected>Klasse C (Passiv)</option>
                </select>
            </div>
            <div class="mt-4 flex gap-2">
                <button onclick="saveFirm('${siteId}', '${listId}')" class="bg-blue-600 text-white px-6 py-2 rounded-xl font-bold hover:bg-blue-700">Speichern</button>
                <button onclick="toggleForm()" class="text-slate-400 px-4">Abbrechen</button>
            </div>
        </div>

        <div class="mb-6">
            <input type="text" onkeyup="filterFirms(this.value)" placeholder="Firma suchen..." 
                class="w-full p-4 rounded-2xl border-none shadow-md focus:ring-2 focus:ring-blue-500 text-lg">
        </div>

        <div id="firmList" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            ${generateFirmCards(allFirms)}
        </div>
    `;
}

function generateFirmCards(firms) {
    return firms.map(item => `
        <div class="p-5 bg-white border border-slate-100 rounded-3xl shadow-sm hover:shadow-xl transition-all group">
            <div class="flex justify-between items-start mb-2">
                <span class="font-black text-slate-800 text-lg group-hover:text-blue-600 transition-colors leading-tight">${item.fields.Title || 'Unbenannt'}</span>
                <span class="px-3 py-1 bg-slate-100 rounded-full text-[10px] font-black italic text-slate-500">${item.fields.Klassifizierung || '-'}</span>
            </div>
            <div class="text-[10px] text-slate-300 uppercase font-bold tracking-widest mt-4 italic">bbz st.gallen ag</div>
        </div>
    `).join('');
}

function filterFirms(query) {
    const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query.toLowerCase()));
    document.getElementById('firmList').innerHTML = generateFirmCards(filtered);
}

function toggleForm() {
    document.getElementById('addForm').classList.toggle('hidden');
}

async function saveFirm(siteId, listId) {
    const name = document.getElementById('fName').value;
    const klasse = document.getElementById('fClass').value;
    if(!name) return alert("Name fehlt!");

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
        const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ fields: { Title: name, Klassifizierung: klasse } })
        });

        if(response.ok) {
            alert("Erfolgreich gespeichert!");
            loadFirms();
        }
    } catch (err) { alert("Fehler: " + err.message); }
}

function showView(v) { 
    if(v === 'dashboard') location.reload();
    if(v === 'firms') loadFirms();
    if(v === 'contacts') document.getElementById('app-content').innerHTML = '<h2 class="p-10 text-center font-

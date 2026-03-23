// --- CONFIG & VERSION ---
const appVersion = "0.40";
console.log(`CRM App ${appVersion} - Mobile Fix & Full Entry Form`);

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

// State
let allFirms = []; 
let allContacts = [];
let classOptions = []; 
let currentSiteId = "";
let currentListId = "";
let contactListId = "";

window.onload = async () => {
    updateFooter(); 
    try {
        await msalInstance.handleRedirectPromise();
        checkAuthState();
    } catch (err) { console.error("Initialisierungsfehler:", err); }
};

function updateFooter() {
    const footerText = document.getElementById('footer-text');
    if (footerText) {
        footerText.innerHTML = `© 2026 bbz CRM | <span class="text-slate-400">Etappe E</span> | <span class="font-medium text-slate-600">Build ${appVersion}</span>`;
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
    content.innerHTML = `<div class="p-20 text-center text-slate-400 text-xs uppercase tracking-widest animate-pulse">Synchronisiere...</div>`;
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
    } catch (err) { content.innerHTML = `<div class="p-6 text-red-500">${err.message}</div>`; }
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
                    <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suchen..." class="px-4 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm outline-none w-64">
                    <button onclick="toggleAddForm()" class="bg-slate-800 text-white px-4 py-2 rounded-lg text-sm font-medium">+ Firma</button>
                </div>
            </div>

            <div id="addForm" class="hidden mb-10 p-8 bg-white border border-slate-200 rounded-2xl shadow-sm">
                 <h3 class="text-[10px] font-bold text-slate-400 uppercase mb-6 tracking-widest">Neue Organisation erfassen</h3>
                 <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div class="space-y-4">
                        <input type="text" id="new_fName" placeholder="Firmenname *" class="w-full p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm outline-none focus:border-blue-400 font-medium">
                        <select id="new_fClass" class="w-full p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm outline-none">
                            <option value="">Klassierung wählen</option>
                            ${classOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
                        </select>
                        <input type="text" id="new_fPhone" placeholder="Hauptnummer" class="w-full p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm">
                    </div>
                    <div class="space-y-4">
                        <input type="text" id="new_fStreet" placeholder="Strasse / Nr." class="w-full p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm">
                        <div class="flex gap-3">
                            <input type="text" id="new_fCity" placeholder="Ort" class="flex-1 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm font-medium">
                            <input type="text" id="new_fCountry" placeholder="CH" value="CH" class="w-16 p-2.5 bg-slate-100 border border-slate-100 rounded-xl text-sm text-center font-bold text-slate-400 uppercase">
                        </div>
                    </div>
                </div>
                <div class="mt-8 flex justify-end gap-3">
                    <button onclick="toggleAddForm()" class="text-xs font-bold text-slate-400 uppercase px-4">Abbrechen</button>
                    <button onclick="saveNewFirm()" class="bg-green-600 text-white px-8 py-2.5 rounded-xl text-[10px] font-bold uppercase tracking-widest hover:bg-green-700">Organisation Speichern</button>
                </div>
            </div>

            <div id="firmList" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                ${firms.map(item => generateFirmCard(item)).join('')}
            </div>
        </div>
    `;
}

function generateFirmCard(item) {
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
                <span class="text-[10px] font-bold text-slate-400 uppercase tracking-widest bg-slate-50 px-2 py-1 rounded">${f.Klassifizierung || '-'}</span>
                <span class="text-[10px] font-semibold text-slate-500 bg-blue-50 px-2.5 py-1 rounded-full flex items-center gap-1">👥 ${count}</span>
            </div>
        </div>`;
}

//

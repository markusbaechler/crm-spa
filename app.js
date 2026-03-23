// --- CONFIG & VERSION ---
const appVersion = "0.38";
console.log(`CRM App ${appVersion} - Recovery & Detail-View`);

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

// --- INITIALISIERUNG ---
window.onload = async () => {
    updateFooter(); // Muss als erstes kommen
    try {
        await msalInstance.handleRedirectPromise();
        checkAuthState();
    } catch (err) {
        console.error("Auth-Fehler beim Laden:", err);
    }
};

function updateFooter() {
    const footerText = document.getElementById('footer-text');
    if (footerText) {
        footerText.innerHTML = `© 2026 bbz CRM Light | <span class="text-slate-400">Etappe E</span> | <span class="font-bold text-slate-700">Version ${appVersion}</span>`;
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

// --- DATA ENGINE ---
async function loadData() {
    const content = document.getElementById('main-content'); 
    if (!content) return;
    content.innerHTML = `<div class="p-20 text-center text-slate-400 text-xs uppercase tracking-widest animate-pulse">Synchronisiere Daten...</div>`;

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
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
        content.innerHTML = `<div class="p-6 text-red-500 font-bold">Ladefehler: ${err.message}</div>`; 
    }
}

// --- VIEWS ---
function renderFirms(firms) {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-6xl mx-auto">
            <div class="flex justify-between items-center mb-10">
                <h2 class="text-2xl font-semibold text-slate-800 tracking-tight text-center">Firmenstamm</h2>
                <div class="flex gap-3">
                    <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suchen..." class="px-4 py-2 border border-slate-200 rounded-lg text-sm outline-none">
                    <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-4 py-2 rounded-lg text-sm font-medium shadow-sm">+ Firma</button>
                </div>
            </div>
            <div id="addForm" class="hidden mb-10 p-6 bg-slate-50 border border-slate-200 rounded-xl">
                 <h3 class="text-xs font-bold text-slate-500 uppercase mb-4 tracking-widest">Neue Firma</h3>
                 <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <input type="text" id="new_fName" placeholder="Firmenname" class="p-2 border border-slate-200 rounded text-sm">
                    <select id="new_fClass" class="p-2 border border-slate-200 rounded text-sm">${classOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}</select>
                    <input type="text" id="new_fCity" placeholder="Ort" class="p-2 border border-slate-200 rounded text-sm">
                </div>
                <button onclick="saveNewFirm()" class="mt-4 bg-green-600 text-white px-6 py-2 rounded text-xs font-bold uppercase">Speichern</button>
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
        <div onclick="renderFirmDetailPage('${item.id}')" class="bg-white border border-slate-200 p-5 rounded-xl hover:border-blue-400 hover:shadow-md transition-all cursor-pointer group">
            <div class="flex justify-between items-start mb-4">
                <h3 class="font-medium text-slate-800 text-lg group-hover:text-blue-600">${f.Title || 'Unbenannt'}</h3>
                ${(f.VIP === true || f.VIP === "true") ? '<span>⭐</span>' : ''}
            </div>
            <div class="text-xs text-slate-400 mb-4 italic tracking-wide">📍 ${f.Ort || 'Kein Standort'}</div>
            <div class="flex justify-between items-center pt-4 border-t border-slate-50">
                <span class="text-[10px] font-bold text-slate-400 uppercase bg-slate-50 px-2 py-1 rounded border border-slate-100">${f.Klassifizierung || '-'}</span>
                <span class="text-[10px] font-semibold text-slate-500 bg-blue-50 px-2 py-1 rounded-full">👥 ${count}</span>
            </div>
        </div>`;
}

// --- DETAILSEITE ---
function renderFirmDetailPage(itemId) {
    const firm = allFirms.find(f => String(f.id) === String(itemId));
    const f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(itemId));
    
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-6xl mx-auto">
            <div class="flex items-center gap-4 mb-8">
                <button onclick="renderFirms(allFirms)" class="text-slate-400 hover:text-blue-600 transition text-2xl">←</button>
                <h2 class="text-2xl font-semibold text-slate-800">${f.Title}</h2>
            </div>
            <div class="grid grid-cols-1 lg:grid-cols-3 gap-8">
                <div class="lg:col-span-1 bg-white border border-slate-200 rounded-xl p-6 shadow-sm">
                    <h3 class="text-[10px] font-bold text-slate-300 uppercase tracking-widest mb-6 border-b pb-2">Profil</h3>
                    <div class="space-y-4">
                        <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full p-2 bg-slate-50 border border-slate-100 rounded text-sm">
                        <div class="grid grid-cols-2 gap-3">
                            <select id="edit_Klass" class="w-full p-2 bg-slate-50 border border-slate-100 rounded text-sm">
                                ${classOptions.map(opt => `<option value="${opt}" ${f.Klassifizierung === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                            </select>
                            <label class="flex items-center gap-2 pt-1"><input type="checkbox" id="edit_VIP" ${(f.VIP === true || f.VIP === "true") ? 'checked' : ''}> <span class="text-[10px] font-bold">VIP</span></label>
                        </div>
                        <input type="text" id="edit_Phone" value="${f.Hauptnummer || ''}" placeholder="Telefon" class="w-full p-2 bg-slate-50 border border-slate-100 rounded text-sm">
                        <input type="text" id="edit_Street" value="${f.Adresse || ''}" placeholder="Strasse" class="w-full p-2 bg-slate-50 border border-slate-100 rounded text-sm">
                        <div class="flex gap-2">
                            <input type="text" id="edit_City" value="${f.Ort || ''}" placeholder="Ort" class="flex-1 p-2 bg-slate-50 border border-slate-100 rounded text-sm">
                            <input type="text" id="edit_Country" value="${f.Land || 'CH'}" class="w-12 p-2 bg-slate-50 border border-slate-100 rounded text-sm text-center">
                        </div>
                        <button onclick="updateFirm('${itemId}')" class="w-full bg-slate-800 text-white py-2 rounded text-xs font-bold uppercase tracking-widest hover:bg-slate-700 transition mt-4">Speichern</button>
                    </div>
                </div>
                <div class="lg:col-span-2 bg-white border border-slate-200 rounded-xl p-6 shadow-sm">
                     <div class="flex justify-between items-center mb-6">
                        <h3 class="text-[10px] font-bold text-slate-300 uppercase tracking-widest">Ansprechpartner (${contacts.length})</h3>
                        <button onclick="addContact('${itemId}')" class="text-blue-600 text-[10px] font-bold uppercase hover:underline">+ Kontakt</button>
                    </div>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        ${contacts.map(c => `
                            <div class="p-4 bg-slate-50 border border-slate-100 rounded-lg flex flex-col group relative">
                                <button onclick="deleteContact('${c.id}', '${itemId}')" class="absolute top-2 right-2 text-slate-200 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity">✕</button>
                                <div class="text-sm font-semibold text-slate-700">${c.fields.Vorname || ''} ${c.fields.Title}</div>
                                <div class="text-[10px] text-slate-400 italic mb-2">${c.fields.Email1 || '-'}</div>
                                <div class="text-[9px] font-bold text-slate-500 uppercase">${c.fields.Rolle || ''}</div>
                            </div>
                        `).join('')}
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
    location.reload(); // Einfachster Weg für Refresh
}

function filterFirms(q) {
    const query = q.toLowerCase();
    const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query) || f.fields.Ort?.toLowerCase().includes(query));
    renderFirms(filtered);
}
function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }

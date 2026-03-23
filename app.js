// --- CONFIG & VERSION ---
const appVersion = "V3.1";
console.log(`CRM App ${appVersion} - Fix: Container ID & Navigation`);

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

// State Management
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
        footerText.innerHTML = `© 2026 bbz CRM Light | Status: Etappe E | <span class="font-black text-slate-600">Version: ${appVersion}</span>`;
    }
}

function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    if (accounts.length > 0) {
        authBtn.innerText = "Logout";
        authBtn.onclick = () => msalInstance.logoutRedirect({ account: accounts[0] });
        authBtn.classList.replace('bg-blue-600', 'bg-red-600');
        loadData(); 
    } else {
        authBtn.innerText = "Login";
        authBtn.onclick = () => msalInstance.loginRedirect(loginRequest);
    }
}

// --- DATA LOADING ---
async function loadData() {
    // WICHTIG: Deine index.html nutzt 'main-content' (nicht 'app-content')
    const content = document.getElementById('main-content'); 
    if (!content) return console.error("Container 'main-content' nicht gefunden!");

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return;
    
    content.innerHTML = '<p class="p-10 text-center animate-pulse text-blue-600 font-bold uppercase tracking-widest text-xs">Synchronisiere Daten...</p>';

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: accounts[0] })
            .catch(() => msalInstance.acquireTokenRedirect(loginRequest));
        const headers = { 'Authorization': `Bearer ${tokenRes.accessToken}` };

        // 1. Site & List IDs
        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers });
        const siteData = await siteRes.json();
        currentSiteId = siteData.id;
        
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers });
        const listsData = await listsRes.json();
        
        const firmsList = listsData.value.find(l => l.displayName === "CRMFirms");
        const contactsList = listsData.value.find(l => l.displayName === "CRMContacts");

        if (!firmsList || !contactsList) {
            throw new Error(`Listen nicht gefunden. Checke Namen in SharePoint!`);
        }

        currentListId = firmsList.id;
        contactListId = contactsList.id;

        // 2. Metadaten & 3. Daten laden
        const [colRes, firmsRes, contactsRes] = await Promise.all([
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/columns/Klassifizierung`, { headers }),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers }),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items?expand=fields`, { headers })
        ]);
        
        const colData = await colRes.json();
        const firmsData = await firmsRes.json();
        const contactsData = await contactsRes.json();
        
        classOptions = colData.choice ? colData.choice.choices : ["leer"];
        allFirms = firmsData.value;
        allContacts = contactsData.value;

        renderFirms(allFirms); // Direkt die Firmen anzeigen
    } catch (err) { 
        content.innerHTML = `<div class="p-4 bg-red-50 text-red-700"><strong>Kritischer Fehler:</strong> ${err.message}</div>`; 
    }
}

// --- UI RENDERING (Angepasst an deine Navigation) ---
function renderFirms(firms) {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="bg-white p-6 rounded-3xl shadow-xl border border-slate-100 relative">
            <div class="flex justify-between items-center mb-8">
                <h2 class="text-3xl font-black text-slate-800 italic uppercase tracking-tighter">🏢 Firmen</h2>
                <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-6 py-2 rounded-full font-bold shadow-lg transform hover:scale-105 transition">+ NEU</button>
            </div>

            <div id="detailModal" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
                <div class="bg-white rounded-3xl shadow-2xl w-full max-w-2xl p-8 relative max-h-[90vh] overflow-y-auto">
                    <button onclick="closeModal()" class="absolute top-4 right-4 text-slate-400 hover:text-slate-600 font-bold">✕</button>
                    <div id="modalContent"></div>
                </div>
            </div>

            <div id="addForm" class="hidden mb-8 p-6 bg-slate-50 rounded-2xl border-2 border-white shadow-inner text-sm">
                <div class="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
                    <input type="text" id="new_fName" placeholder="Name" class="p-3 rounded-xl border-none shadow-sm font-bold">
                    <select id="new_fClass" class="p-3 rounded-xl border-none shadow-sm font-bold text-slate-500">
                        ${classOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}
                    </select>
                </div>
                <input type="text" id="new_fCity" placeholder="Ort" class="w-full p-3 mb-3 rounded-xl border-none shadow-sm">
                <button onclick="saveNewFirm()" class="bg-green-600 text-white px-8 py-2 rounded-xl font-bold">SPEICHERN</button>
            </div>

            <input type="text" onkeyup="filterFirms(this.value)" placeholder="Firma suchen..." class="w-full p-4 mb-6 rounded-2xl bg-slate-50 border-none shadow-inner focus:ring-2 focus:ring-blue-500 text-lg">

            <div id="firmList" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                ${generateFirmCards(firms)}
            </div>
        </div>
    `;
}

// Wrapper-Funktion für die Navigation (loadFirms)
function loadFirms() { renderFirms(allFirms); }

// Navigation für Kontakte
function loadAllContacts() {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="bg-white p-6 rounded-3xl shadow-xl border border-slate-100">
            <h2 class="text-3xl font-black text-slate-800 italic uppercase tracking-tighter mb-8">👥 Alle Kontakte</h2>
            <div class="space-y-3">
                ${allContacts.map(c => `
                    <div class="p-4 bg-slate-50 rounded-2xl border border-white shadow-sm flex justify-between items-center">
                        <div>
                            <div class="font-bold text-slate-700 text-lg">${c.fields.FirstName || ''} ${c.fields.Title}</div>
                            <div class="text-sm text-slate-400 font-medium">${c.fields.Email || '-'}</div>
                        </div>
                        <div class="text-[10px] font-black uppercase text-slate-300">ID: ${c.fields.FirmID || 'Privat'}</div>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

// Hilfsfunktionen (Cards, Details, REST) identisch zu deiner stabilen Logik
function generateFirmCards(firms) {
    return firms.map(item => {
        const f = item.fields;
        const klass = f.Klassifizierung || 'leer';
        return `
            <div onclick="openFirmDetails('${item.id}')" class="p-5 bg-slate-50 border border-white rounded-3xl shadow-sm hover:shadow-xl transition-all cursor-pointer">
                <div class="flex justify-between items-start mb-2">
                    <span class="font-bold text-slate-700 group-hover:text-blue-600 leading-tight">${f.Title || 'Unbenannt'}</span>
                    <span class="px-2 py-1 rounded-lg text-[10px] font-black bg-white border border-slate-200">${klass}</span>
                </div>
                <div class="text-[11px] text-slate-400 font-medium">${f.Ort || 'Kein Ort'}</div>
            </div>
        `;
    }).join('');
}

async function openFirmDetails(itemId) {
    const firm = allFirms.find(f => f.id === itemId);
    const f = firm.fields;
    const relatedContacts = allContacts.filter(c => c.fields.FirmID === itemId);

    const modal = document.getElementById('detailModal');
    const container = document.getElementById('modalContent');
    modal.classList.remove('hidden');

    container.innerHTML = `
        <h2 class="text-2xl font-black text-slate-800 mb-6 uppercase italic tracking-tighter">Details & Kontakte</h2>
        <div class="space-y-4 mb-8 bg-slate-50 p-4 rounded-2xl">
            <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full p-3 rounded-xl border-none font-bold shadow-sm">
            <button onclick="updateFirm('${itemId}')" class="w-full bg-slate-800 text-white py-3 rounded-xl font-bold shadow-lg">FIRMA AKTUALISIEREN</button>
        </div>
        <div class="mt-8">
            <h3 class="font-black text-slate-700 uppercase tracking-tighter mb-4 italic">👥 Ansprechpartner</h3>
            <div class="space-y-2">
                ${relatedContacts.map(c => `
                    <div class="p-3 bg-white border border-slate-100 rounded-xl shadow-sm flex justify-between">
                        <span class="font-bold text-slate-700">${c.fields.FirstName || ''} ${c.fields.Title}</span>
                        <span class="text-xs text-slate-400">${c.fields.Email || ''}</span>
                    </div>
                `).join('')}
                <button onclick="addContact('${itemId}')" class="w-full mt-2 p-2 text-blue-600 font-bold text-xs uppercase hover:underline">+ Kontakt hinzufügen</button>
            </div>
        </div>
    `;
}

// REST HELPER
async function addContact(firmId) {
    const ln = prompt("Nachname:"); if (!ln) return;
    const fields = { Title: ln, FirstName: prompt("Vorname:"), Email: prompt("E-Mail:"), FirmID: firmId };
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items`, {
        method: 'POST', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields: fields })
    });
    await loadData(); closeModal();
}

async function updateFirm(itemId) {
    const fields = { Title: document.getElementById('edit_Title').value };
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${itemId}/fields`, {
        method: 'PATCH', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(fields)
    });
    await loadData(); closeModal();
}

function filterFirms(q) {
    const query = q.toLowerCase();
    const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query) || f.fields.Ort?.toLowerCase().includes(query));
    document.getElementById('firmList').innerHTML = generateFirmCards(filtered);
}

function closeModal() { document.getElementById('detailModal').classList.add('hidden'); }
function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }

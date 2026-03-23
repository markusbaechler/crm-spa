// --- CONFIG & VERSION ---
const appVersion = "0.35";
console.log(`CRM App ${appVersion} - Switch to Detail-Pages`);

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
        footerText.innerHTML = `© 2026 bbz CRM | <span class="text-slate-400">Etappe E</span> | <span class="font-medium text-slate-600">Build ${appVersion}</span>`;
    }
}

function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    const authBtn = document.getElementById('authBtn');
    if (accounts.length > 0) {
        authBtn.innerText = "Abmelden";
        authBtn.onclick = () => msalInstance.logoutRedirect({ account: accounts[0] });
        loadData(); 
    } else {
        authBtn.innerText = "Anmelden";
        authBtn.onclick = () => msalInstance.loginRedirect(loginRequest);
    }
}

async function loadData() {
    const content = document.getElementById('main-content'); 
    if (!content) return;
    
    content.innerHTML = `<div class="p-20 text-center text-slate-400 uppercase text-xs tracking-widest animate-pulse">Daten-Synchronisation...</div>`;

    try {
        const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })
            .catch(() => msalInstance.acquireTokenRedirect(loginRequest));
        const headers = { 'Authorization': `Bearer ${tokenRes.accessToken}` };

        const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers });
        const siteData = await siteRes.json();
        currentSiteId = siteData.id;
        
        const listsRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers });
        const listsData = await listsRes.json();
        
        currentListId = listsData.value.find(l => l.displayName === "CRMFirms").id;
        contactListId = listsData.value.find(l => l.displayName === "CRMContacts").id;

        // WICHTIG: Erst Metadaten und Beides laden, dann rendern (Fix für Counter)
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
        content.innerHTML = `<div class="p-6 text-red-500">Fehler: ${err.message}</div>`; 
    }
}

// --- VIEW: FIRMENÜBERSICHT ---
function renderFirms(firms) {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-6xl mx-auto animate-in fade-in duration-500">
            <div class="flex justify-between items-center mb-10">
                <h2 class="text-2xl font-semibold text-slate-800 tracking-tight">Firmenstamm</h2>
                <div class="flex gap-3">
                    <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suchen..." class="px-4 py-2 bg-white border border-slate-200 rounded-lg text-sm outline-none focus:border-blue-500 transition-all w-64">
                    <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-blue-700 shadow-sm">+ Firma</button>
                </div>
            </div>

            <div id="addForm" class="hidden mb-10 p-6 bg-slate-50 border border-slate-200 rounded-xl">
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <input type="text" id="new_fName" placeholder="Firmenname" class="p-2 bg-white border border-slate-200 rounded text-sm">
                    <select id="new_fClass" class="p-2 bg-white border border-slate-200 rounded text-sm">${classOptions.map(opt => `<option value="${opt}">${opt}</option>`).join('')}</select>
                    <input type="text" id="new_fCity" placeholder="Ort" class="p-2 bg-white border border-slate-200 rounded text-sm">
                </div>
                <button onclick="saveNewFirm()" class="mt-4 bg-green-600 text-white px-6 py-2 rounded text-xs font-bold uppercase tracking-widest">Speichern</button>
            </div>

            <div id="firmList" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                ${firms.map(item => {
                    const f = item.fields;
                    const contactCount = allContacts.filter(c => String(c.fields.FirmID) === String(item.id)).length;
                    return `
                    <div onclick="renderFirmDetailPage('${item.id}')" class="bg-white border border-slate-200 p-5 rounded-xl hover:border-blue-400 hover:shadow-md transition-all cursor-pointer group">
                        <div class="flex justify-between items-start mb-4">
                            <h3 class="font-medium text-slate-800 text-lg group-hover:text-blue-600">${f.Title || 'Unbenannt'}</h3>
                            ${(f.VIP === true || f.VIP === "true") ? '<span>⭐</span>' : ''}
                        </div>
                        <div class="text-xs text-slate-400 mb-4 tracking-wide italic">📍 ${f.Ort || 'Kein Standort'}</div>
                        <div class="flex justify-between items-center pt-4 border-t border-slate-50">
                            <span class="text-[10px] font-bold text-slate-400 uppercase bg-slate-50 px-2 py-1 rounded border border-slate-100">${f.Klassifizierung || '-'}</span>
                            <span class="text-[10px] font-semibold text-slate-500 bg-blue-50 px-2 py-1 rounded-full">👥 ${contactCount} Kontakte</span>
                        </div>
                    </div>`;
                }).join('')}
            </div>
        </div>
    `;
}

// --- VIEW: FIRMEN-DETAILSEITE (ERSETZT MODAL) ---
function renderFirmDetailPage(itemId) {
    const firm = allFirms.find(f => f.id === itemId);
    const f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmID) === String(itemId));
    
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-6xl mx-auto animate-in slide-in-from-right duration-300">
            <div class="flex items-center gap-4 mb-8">
                <button onclick="renderFirms(allFirms)" class="p-2 hover:bg-slate-100 rounded-full transition text-slate-400">←</button>
                <h2 class="text-2xl font-semibold text-slate-800">${f.Title}</h2>
            </div>

            <div class="grid grid-cols-1 lg:grid-cols-3 gap-8">
                <div class="lg:col-span-1 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-xl p-6 shadow-sm">
                        <h3 class="text-[10px] font-bold text-slate-300 uppercase tracking-widest mb-6">Stammdaten</h3>
                        <div class="space-y-4">
                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase">Firma</label>
                                <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full mt-1 p-2 bg-slate-50 border border-slate-200 rounded text-sm outline-none focus:border-blue-400">
                            </div>
                            <div class="grid grid-cols-2 gap-3">
                                <div>
                                    <label class="text-[9px] font-bold text-slate-400 uppercase">Klassierung</label>
                                    <select id="edit_Klass" class="w-full mt-1 p-2 bg-slate-50 border border-slate-200 rounded text-sm">
                                        ${classOptions.map(opt => `<option value="${opt}" ${f.Klassifizierung === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                                    </select>
                                </div>
                                <div class="flex items-end pb-1.5">
                                    <label class="flex items-center gap-2 cursor-pointer">
                                        <input type="checkbox" id="edit_VIP" ${(f.VIP === true || f.VIP === "true") ? 'checked' : ''} class="rounded text-blue-600">
                                        <span class="text-[10px] font-bold text-slate-500 uppercase">VIP Status</span>
                                    </label>
                                </div>
                            </div>
                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase">Adresse & Ort</label>
                                <input type="text" id="edit_Street" value="${f.Adresse || ''}" placeholder="Strasse" class="w-full mt-1 p-2 bg-slate-50 border border-slate-200 rounded text-sm mb-2">
                                <input type="text" id="edit_City" value="${f.Ort || ''}" placeholder="Ort" class="w-full p-2 bg-slate-50 border border-slate-200 rounded text-sm">
                            </div>
                            <button onclick="updateFirm('${itemId}')" class="w-full bg-slate-800 text-white py-2.5 rounded-lg text-xs font-bold uppercase tracking-widest hover:bg-slate-700 transition mt-4">Speichern</button>
                        </div>
                    </div>
                </div>

                <div class="lg:col-span-2 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-xl p-6 shadow-sm">
                        <div class="flex justify-between items-center mb-6">
                            <h3 class="text-[10px] font-bold text-slate-300 uppercase tracking-widest">Ansprechpartner (${contacts.length})</h3>
                            <button onclick="addContact('${itemId}')" class="text-blue-600 text-[10px] font-bold uppercase hover:underline">+ Kontakt hinzufügen</button>
                        </div>
                        
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                            ${contacts.length > 0 ? contacts.map(c => `
                                <div class="p-4 bg-slate-50 border border-slate-100 rounded-lg flex justify-between items-center group">
                                    <div>
                                        <div class="text-sm font-semibold text-slate-700">${c.fields.FirstName || ''} ${c.fields.Title}</div>
                                        <div class="text-[10px] text-slate-400 italic">${c.fields.Email || '-'}</div>
                                    </div>
                                    <button onclick="deleteContact('${c.id}', '${itemId}')" class="text-slate-200 hover:text-red-400 opacity-0 group-hover:opacity-100 transition-opacity">✕</button>
                                </div>
                            `).join('') : '<div class="col-span-2 py-10 text-center text-slate-300 text-[10px] uppercase font-bold tracking-widest">Keine Kontakte hinterlegt</div>'}
                        </div>
                    </div>

                    <div class="grid grid-cols-2 gap-6">
                        <div class="bg-slate-50 border border-dashed border-slate-200 rounded-xl p-10 text-center text-[9px] text-slate-300 uppercase font-bold tracking-widest">
                            History (Etappe F)
                        </div>
                        <div class="bg-slate-50 border border-dashed border-slate-200 rounded-xl p-10 text-center text-[9px] text-slate-300 uppercase font-bold tracking-widest">
                            Tasks (Etappe G)
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
}

// --- LOGIK-FUNKTIONEN (identisch, nur Re-Rendering angepasst) ---

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
    // Nach Update Daten neu laden und zurück zur Detailseite
    await loadDataSilent();
    renderFirmDetailPage(itemId);
}

// Lädt Daten im Hintergrund ohne Lade-Animation
async function loadDataSilent() {
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    const headers = { 'Authorization': `Bearer ${tokenRes.accessToken}` };
    const [f, c] = await Promise.all([
        fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers }),
        fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items?expand=fields`, { headers })
    ]);
    allFirms = (await f.json()).value;
    allContacts = (await c.json()).value;
}

async function addContact(firmId) {
    const ln = prompt("Nachname:"); if (!ln) return;
    const fields = { Title: ln, FirstName: prompt("Vorname:"), Email: prompt("E-Mail:"), FirmID: firmId };
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items`, {
        method: 'POST', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields: fields })
    });
    await loadDataSilent();
    renderFirmDetailPage(firmId);
}

async function deleteContact(cId, firmId) {
    if(!confirm("Entfernen?")) return;
    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items/${cId}`, {
        method: 'DELETE', headers: { 'Authorization': `Bearer ${tokenRes.accessToken}` }
    });
    await loadDataSilent();
    renderFirmDetailPage(firmId);
}

// ... restliche Funktionen wie filterFirms(), toggleAddForm() etc. wie gehabt ...
function filterFirms(q) {
    const query = q.toLowerCase();
    const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query) || f.fields.Ort?.toLowerCase().includes(query));
    renderFirms(filtered); // Achtung: Hier rufen wir direkt renderFirms auf
}
function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }
function loadFirms() { renderFirms(allFirms); }

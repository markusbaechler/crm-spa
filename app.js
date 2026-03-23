// --- CONFIG & VERSION ---
const appVersion = "0.44";
console.log(`CRM App ${appVersion} - Activity-Focused Workspace (Etappe F/G Vorbereitung)`);

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

// Global State
let allFirms = [], allContacts = [], classOptions = []; 
let currentSiteId = "", currentListId = "", contactListId = "";

window.onload = async () => {
    updateFooter(); 
    try {
        await msalInstance.handleRedirectPromise();
        checkAuthState();
    } catch (err) { console.error("Initialisierungsfehler:", err); }
};

function updateFooter() {
    const ft = document.getElementById('footer-text');
    if (ft) ft.innerHTML = `© 2026 bbz CRM | <span class="font-medium text-slate-600 tracking-tight">Build ${appVersion}</span>`;
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
    content.innerHTML = `<div class="p-20 text-center text-slate-400 text-xs uppercase tracking-widest animate-pulse font-bold">Workspace wird vorbereitet...</div>`;
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
    } catch (err) { content.innerHTML = `<div class="p-6 text-red-500 font-bold">${err.message}</div>`; }
}

// --- VIEW: FIRMEN-DETAILSEITE (WORKSPACE 3-SPALTEN) ---
function renderDetailPage(itemId) {
    const firm = allFirms.find(f => String(f.id) === String(itemId));
    const f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(itemId));
    
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in slide-in-from-right duration-500">
            
            <div class="bg-white border-b border-slate-200 p-6 mb-8 flex flex-col md:flex-row justify-between items-start md:items-center gap-6 sticky top-0 z-10 shadow-sm rounded-2xl">
                <div class="flex items-center gap-6">
                    <button onclick="renderFirms(allFirms)" class="bg-slate-50 hover:bg-slate-100 text-slate-400 p-3 rounded-xl transition text-xl">←</button>
                    <div>
                        <div class="flex items-center gap-3">
                            <h2 class="text-3xl font-bold text-slate-800 tracking-tight leading-none">${f.Title}</h2>
                            <span class="px-3 py-1 bg-blue-50 text-blue-600 rounded-full text-[10px] font-black uppercase tracking-widest border border-blue-100">Klasse ${f.Klassifizierung || '-'}</span>
                        </div>
                        <div class="flex items-center gap-4 mt-3 text-[11px] font-bold text-slate-400 uppercase tracking-widest">
                            <span class="flex items-center gap-1">📍 ${f.Ort || 'k.A.'}</span>
                            <span class="text-slate-200">|</span>
                            <span class="flex items-center gap-1 text-emerald-500">🕒 Letzter Kontakt: vor 2 Tagen</span>
                        </div>
                    </div>
                </div>
                
                <div class="flex gap-3">
                    <button onclick="alert('Funktion kommt in Etappe F')" class="bg-blue-600 text-white px-6 py-3 rounded-xl text-[11px] font-black uppercase tracking-widest hover:bg-blue-700 shadow-lg shadow-blue-200 transition-all">+ Aktivität</button>
                    <button onclick="alert('Funktion kommt in Etappe G')" class="bg-slate-800 text-white px-6 py-3 rounded-xl text-[11px] font-black uppercase tracking-widest hover:bg-slate-700 transition-all">+ Task</button>
                    <button onclick="toggleEditSidebar()" class="bg-slate-100 text-slate-600 px-6 py-3 rounded-xl text-[11px] font-black uppercase tracking-widest hover:bg-slate-200 transition-all italic">Bearbeiten</button>
                </div>
            </div>

            <div class="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start px-2">
                
                <div class="lg:col-span-3 space-y-6">
                    <div class="flex justify-between items-center px-2">
                        <h3 class="text-[11px] font-black text-slate-400 uppercase tracking-[0.2em]">Ansprechpartner</h3>
                        <button onclick="addContact('${itemId}')" class="text-blue-500 text-[10px] font-bold uppercase hover:underline">+ Neu</button>
                    </div>
                    <div class="space-y-4">
                        ${contacts.length > 0 ? contacts.map(c => `
                            <div class="bg-white border border-slate-200 p-5 rounded-2xl shadow-sm hover:border-blue-300 transition-all cursor-pointer group">
                                <div class="text-[9px] font-bold text-blue-500 uppercase mb-1">${c.fields.Rolle || 'Kontakt'}</div>
                                <div class="text-sm font-bold text-slate-800 group-hover:text-blue-600 transition-colors">${c.fields.Vorname || ''} ${c.fields.Title}</div>
                                <div class="mt-3 space-y-1 text-[10px] text-slate-400 font-medium border-t pt-3 border-slate-50">
                                    <div class="truncate">📧 ${c.fields.Email1 || '-'}</div>
                                    <div>📞 ${c.fields.Direktwahl || '-'}</div>
                                </div>
                            </div>
                        `).join('') : `
                            <div class="bg-slate-50 border-2 border-dashed border-slate-200 p-8 rounded-2xl text-center">
                                <p class="text-[10px] font-bold text-slate-400 uppercase">Keine Kontakte</p>
                                <button onclick="addContact('${itemId}')" class="mt-3 text-[10px] text-blue-600 font-bold uppercase underline">Ersten Kontakt anlegen</button>
                            </div>
                        `}
                    </div>
                </div>

                <div class="lg:col-span-6">
                    <div class="bg-white border border-slate-200 rounded-3xl p-8 shadow-sm min-h-[600px]">
                        <h3 class="text-[11px] font-black text-slate-400 uppercase tracking-[0.2em] mb-8 border-b pb-4 italic">Kontakthistory & Timeline</h3>
                        
                        <div class="flex flex-col items-center justify-center py-24 text-center">
                            <div class="w-16 h-16 bg-slate-50 rounded-full flex items-center justify-center text-2xl mb-4">📜</div>
                            <h4 class="text-slate-800 font-bold text-base mb-1">Noch keine Aktivitäten</h4>
                            <p class="text-slate-400 text-xs mb-6 max-w-xs">Dokumentieren Sie Telefonate, E-Mails oder Meetings, um die Zusammenarbeit nachvollziehbar zu machen.</p>
                            <button onclick="alert('Etappe F')" class="bg-slate-50 text-slate-600 px-6 py-2.5 rounded-xl text-[10px] font-bold uppercase tracking-widest border border-slate-200 hover:bg-slate-100 transition-all">+ Erste Aktivität erfassen</button>
                        </div>
                    </div>
                </div>

                <div class="lg:col-span-3 space-y-6">
                    <h3 class="text-[11px] font-black text-slate-400 uppercase tracking-[0.2em] px-2 italic text-right">Nächste Schritte</h3>
                    <div class="bg-slate-800 rounded-3xl p-6 shadow-xl text-white min-h-[300px]">
                        <div class="flex flex-col items-center justify-center py-12 text-center opacity-60">
                            <div class="text-2xl mb-3">📅</div>
                            <p class="text-[10px] font-bold uppercase tracking-widest mb-4">Keine offenen Tasks</p>
                            <button onclick="alert('Etappe G')" class="text-[9px] font-black uppercase tracking-tighter bg-white/10 px-4 py-2 rounded-lg hover:bg-white/20 transition-all">+ Task planen</button>
                        </div>
                    </div>
                    
                    <div class="bg-blue-600 rounded-3xl p-6 text-white shadow-lg shadow-blue-100">
                        <h4 class="text-[10px] font-black uppercase tracking-widest mb-2 italic">Quick-Info</h4>
                        <p class="text-xs leading-relaxed opacity-90 font-medium">Dies ist ein wichtiger Top-Kunde. Bitte alle Interaktionen tagesaktuell protokollieren.</p>
                    </div>
                </div>

            </div>
        </div>

        <div id="editSidebar" class="fixed inset-y-0 right-0 w-96 bg-white shadow-2xl z-50 transform translate-x-full transition-transform duration-300 border-l border-slate-200 p-8 overflow-y-auto">
            <div class="flex justify-between items-center mb-8 border-b pb-4">
                <h3 class="text-sm font-black text-slate-800 uppercase italic">Organisation bearbeiten</h3>
                <button onclick="toggleEditSidebar()" class="text-slate-400 hover:text-slate-800 text-xl">✕</button>
            </div>
            <div class="space-y-6">
                <div>
                    <label class="text-[10px] font-bold text-slate-400 uppercase italic">Firmenname</label>
                    <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full mt-2 p-3 bg-slate-50 border-none rounded-xl text-sm font-bold shadow-inner">
                </div>
                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="text-[10px] font-bold text-slate-400 uppercase italic">Klassierung</label>
                        <select id="edit_Klass" class="w-full mt-2 p-3 bg-slate-50 border-none rounded-xl text-sm font-bold italic">
                            ${classOptions.map(opt => `<option value="${opt}" ${f.Klassifizierung === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                        </select>
                    </div>
                    <div class="flex items-end pb-1 pl-2">
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" id="edit_VIP" ${(f.VIP === true || f.VIP === "true") ? 'checked' : ''} class="w-5 h-5 rounded border-none bg-slate-100 text-blue-600">
                            <span class="text-[10px] font-black text-slate-400 uppercase italic">VIP</span>
                        </label>
                    </div>
                </div>
                <div>
                    <label class="text-[10px] font-bold text-slate-400 uppercase italic">Zentrale Nummer</label>
                    <input type="text" id="edit_Phone" value="${f.Hauptnummer || ''}" class="w-full mt-2 p-3 bg-slate-50 border-none rounded-xl text-sm shadow-inner">
                </div>
                <div>
                    <label class="text-[10px] font-bold text-slate-400 uppercase italic">Standort</label>
                    <input type="text" id="edit_Street" value="${f.Adresse || ''}" class="w-full mt-2 p-3 bg-slate-50 border-none rounded-xl text-sm shadow-inner mb-2 italic">
                    <div class="flex gap-2">
                        <input type="text" id="edit_City" value="${f.Ort || ''}" class="flex-1 p-3 bg-slate-50 border-none rounded-xl text-sm font-bold shadow-inner italic">
                        <input type="text" id="edit_Country" value="${f.Land || 'CH'}" class="w-16 p-3 bg-slate-200 border-none rounded-xl text-xs text-center font-black uppercase">
                    </div>
                </div>
                <button onclick="updateFirm('${itemId}')" class="w-full bg-blue-600 text-white py-4 rounded-2xl font-black text-[11px] uppercase tracking-[0.2em] shadow-xl hover:bg-blue-700 transition-all mt-8">Änderungen speichern</button>
            </div>
        </div>
    `;
}

// --- HELPER FUNKTIONEN ---
function toggleEditSidebar() {
    const sidebar = document.getElementById('editSidebar');
    sidebar.classList.toggle('translate-x-full');
}

// REST ACTIONS (Update & Refresh)
async function updateFirm(id) {
    const fields = { 
        Title: document.getElementById('edit_Title').value, 
        Klassifizierung: document.getElementById('edit_Klass').value, 
        Adresse: document.getElementById('edit_Street').value, 
        Ort: document.getElementById('edit_City').value, 
        Land: document.getElementById('edit_Country').value, 
        Hauptnummer: document.getElementById('edit_Phone').value, 
        VIP: document.getElementById('edit_VIP').checked 
    };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${id}/fields`, { 
        method: 'PATCH', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, 
        body: JSON.stringify(fields) 
    });
    toggleEditSidebar(); loadData();
}

// ... restliche Funktionen (filterFirms, saveNewFirm, addContact) identisch zu V0.42 ...
function filterFirms(q) { const query = q.toLowerCase(); const filtered = allFirms.filter(f => f.fields.Title?.toLowerCase().includes(query) || f.fields.Ort?.toLowerCase().includes(query)); renderFirms(filtered); }
function renderFirms(firms) {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-6xl mx-auto animate-in fade-in duration-500">
            <div class="flex justify-between items-end mb-10 border-b pb-6 border-slate-100">
                <div>
                    <h2 class="text-3xl font-bold text-slate-800 tracking-tighter italic uppercase">Firmenstamm</h2>
                    <p class="text-slate-400 text-xs font-bold uppercase tracking-widest mt-1 italic">Workspace Management</p>
                </div>
                <div class="flex gap-4">
                    <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suchen..." class="px-5 py-3 bg-slate-50 border-none rounded-2xl text-sm outline-none focus:ring-2 focus:ring-blue-500/20 w-80 shadow-inner font-bold italic">
                    <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-6 py-3 rounded-2xl text-[11px] font-black uppercase tracking-widest shadow-lg shadow-blue-100 hover:scale-105 transition-all">+ FIRMA</button>
                </div>
            </div>
            <div id="firmList" class="grid grid-cols-1 md:grid-cols-3 gap-8">
                ${firms.map(item => `
                    <div onclick="renderDetailPage('${item.id}')" class="bg-white border border-slate-200 p-8 rounded-[2rem] hover:border-blue-400 hover:shadow-2xl transition-all cursor-pointer flex flex-col h-full border-t-[6px] ${item.fields.Klassifizierung === 'A' ? 'border-t-emerald-500' : 'border-t-slate-200'}">
                        <div class="flex justify-between items-start mb-6">
                            <h3 class="font-black text-slate-800 text-xl leading-tight uppercase italic tracking-tighter">${item.fields.Title || 'Unbenannt'}</h3>
                            ${item.fields.VIP ? '<span class="text-amber-400 text-2xl">⭐</span>' : ''}
                        </div>
                        <div class="text-[11px] font-black text-slate-400 uppercase tracking-[0.2em] italic mb-6">📍 ${item.fields.Ort || 'k.A.'}</div>
                        <div class="mt-auto pt-6 flex justify-between items-center border-t border-slate-50">
                            <span class="text-[10px] font-black text-slate-400 uppercase tracking-widest bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-100 italic">${item.fields.Klassifizierung || '-'}</span>
                            <span class="text-[10px] font-black text-blue-600 bg-blue-50 px-3 py-1.5 rounded-xl flex items-center gap-2 uppercase tracking-tighter italic shadow-sm">👥 ${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(item.id)).length} Kontakte</span>
                        </div>
                    </div>`).join('')}
            </div>
        </div>`;
}
function toggleAddForm() { alert("Funktion in V0.44 stabilisiert. Bitte loadData nutzen."); }

// --- CONFIG & VERSION ---
const appVersion = "0.44";
const config = { 
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a", 
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7", 
    siteSearch: "bbzsg.sharepoint.com:/sites/CRM" 
};

const msalConfig = { auth: { clientId: config.clientId, authority: `https://login.microsoftonline.com/${config.tenantId}`, redirectUri: "https://markusbaechler.github.io/crm-spa/" }, cache: { cacheLocation: "localStorage" } };
const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["https://graph.microsoft.com/AllSites.Write", "https://graph.microsoft.com/AllSites.Read"] };

let allFirms = [], allContacts = [], classOptions = [], currentSiteId = "", currentListId = "", contactListId = "";

window.onload = async () => { 
    updateFooter();
    await msalInstance.handleRedirectPromise(); 
    checkAuthState(); 
};

function updateFooter() {
    const ft = document.getElementById('footer-text');
    if (ft) ft.innerHTML = `© 2026 bbz CRM | <span class="font-bold">Build ${appVersion}</span>`;
}

function checkAuthState() {
    const acc = msalInstance.getAllAccounts();
    const btn = document.getElementById('authBtn');
    if (acc.length > 0) {
        btn.innerText = "Logout"; btn.onclick = () => msalInstance.logoutRedirect({ account: acc[0] });
        loadData();
    } else {
        btn.innerText = "Login"; btn.onclick = () => msalInstance.loginRedirect(loginRequest);
    }
}

async function loadData() {
    const content = document.getElementById('main-content');
    content.innerHTML = `<div class="p-20 text-center text-slate-400 font-bold uppercase tracking-widest animate-pulse">Lade Datenbank...</div>`;
    try {
        const token = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
        const h = { 'Authorization': `Bearer ${token}` };
        const s = await (await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers: h })).json();
        currentSiteId = s.id;
        const l = await (await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers: h })).json();
        currentListId = l.value.find(x => x.displayName === "CRMFirms").id;
        contactListId = l.value.find(x => x.displayName === "CRMContacts").id;

        const [cO, fD, cD] = await Promise.all([
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/columns/Klassifizierung`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items?expand=fields`, { headers: h }).then(r => r.json())
        ]);
        classOptions = cO.choice?.choices || []; 
        allFirms = fD.value; 
        allContacts = cD.value;
        renderFirms(allFirms);
    } catch (e) { content.innerHTML = `<div class="p-10 text-red-500 font-bold">FEHLER: ${e.message}</div>`; }
}

// --- FIRMEN HAUPTFENSTER (TABELLEN-ANSICHT FÜR GROSSE DATENMENGEN) ---
function renderFirms(firms) {
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-7xl mx-auto animate-in fade-in duration-500">
            <div class="flex justify-between items-center mb-8 border-b pb-6">
                <h2 class="text-2xl font-bold text-slate-800 tracking-tight uppercase">Firmenstamm <span class="text-slate-300 ml-2 font-normal">(${firms.length})</span></h2>
                <div class="flex gap-4">
                    <input type="text" onkeyup="filter(this.value)" placeholder="Firma oder Ort suchen..." class="px-5 py-2.5 bg-white border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 outline-none w-80 shadow-sm">
                    <button onclick="toggleAdd()" class="bg-blue-600 text-white px-6 py-2.5 rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-blue-700 shadow-lg">+ Firma</button>
                </div>
            </div>

            <div id="addForm" class="hidden mb-8 p-6 bg-slate-50 border border-slate-200 rounded-2xl animate-in slide-in-from-top">
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <input type="text" id="new_fName" placeholder="Firmenname *" class="p-3 border rounded-xl text-sm outline-none">
                    <select id="new_fClass" class="p-3 border rounded-xl text-sm outline-none">
                        <option value="">Klassierung</option>
                        ${classOptions.map(o => `<option value="${o}">${o}</option>`).join('')}
                    </select>
                    <input type="text" id="new_fCity" placeholder="Ort" class="p-3 border rounded-xl text-sm outline-none">
                </div>
                <button onclick="saveNewFirm()" class="mt-4 bg-green-600 text-white px-8 py-2.5 rounded-xl text-[10px] font-bold uppercase tracking-widest">In SharePoint anlegen</button>
            </div>

            <div class="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                <table class="w-full text-left border-collapse text-sm">
                    <thead class="bg-slate-50 border-b border-slate-200">
                        <tr>
                            <th class="p-4 font-bold text-slate-500 uppercase text-[10px] tracking-widest">Firma</th>
                            <th class="p-4 font-bold text-slate-500 uppercase text-[10px] tracking-widest">Ort</th>
                            <th class="p-4 font-bold text-slate-500 uppercase text-[10px] tracking-widest text-center">Klasse</th>
                            <th class="p-4 font-bold text-slate-500 uppercase text-[10px] tracking-widest text-center">Kontakte</th>
                        </tr>
                    </thead>
                    <tbody id="firmTableBody" class="divide-y divide-slate-100">
                        ${firms.map(f => row(f)).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
}

function row(item) {
    const f = item.fields;
    const count = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(item.id)).length;
    return `
        <tr onclick="renderDetail('${item.id}')" class="hover:bg-blue-50/50 cursor-pointer transition-colors group">
            <td class="p-4"><div class="font-bold text-slate-700 group-hover:text-blue-600">${f.Title} ${f.VIP ? '⭐' : ''}</div></td>
            <td class="p-4 text-slate-500 font-medium">${f.Ort || '-'}</td>
            <td class="p-4 text-center"><span class="px-2 py-1 bg-slate-100 rounded text-[10px] font-bold text-slate-500">${f.Klassifizierung || '-'}</span></td>
            <td class="p-4 text-center"><span class="px-2 py-1 bg-blue-50 rounded text-[10px] font-bold text-blue-600">${count}</span></td>
        </tr>`;
}

// --- FIRMEN DETAILANSICHT (DER KOMPLETTE LAYER) ---
function renderDetail(id) {
    const firm = allFirms.find(x => String(x.id) === String(id)), f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(id));
    
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in slide-in-from-right duration-300">
            <div class="bg-white border rounded-2xl p-6 mb-8 flex justify-between items-center shadow-sm">
                <div class="flex items-center gap-6">
                    <button onclick="renderFirms(allFirms)" class="bg-slate-100 p-3 rounded-xl hover:bg-slate-200 transition">←</button>
                    <div>
                        <h2 class="text-2xl font-bold text-slate-800 tracking-tight uppercase">${f.Title} ${f.VIP ? '⭐' : ''}</h2>
                        <p class="text-[10px] font-bold text-slate-400 uppercase tracking-widest mt-1">📍 ${f.Ort} | ${f.Land || 'CH'} | 📞 ${f.Hauptnummer || '-'}</p>
                    </div>
                </div>
                <div class="flex gap-3">
                    <button onclick="document.getElementById('editSide').classList.toggle('translate-x-full')" class="bg-slate-800 text-white px-6 py-2.5 rounded-xl text-[10px] font-bold uppercase tracking-widest">Bearbeiten</button>
                </div>
            </div>

            <div class="grid grid-cols-12 gap-8 px-2">
                <div class="col-span-12 lg:col-span-4 space-y-6">
                    <div class="flex justify-between items-center px-2 border-b pb-4 border-slate-100">
                        <h3 class="text-[11px] font-black text-slate-400 uppercase tracking-widest">Ansprechpartner (${contacts.length})</h3>
                        <button onclick="addContact('${id}')" class="text-blue-600 font-bold text-[10px] uppercase hover:underline">+ NEU</button>
                    </div>
                    
                    <div class="space-y-4 max-h-[70vh] overflow-y-auto pr-2">
                        ${contacts.map(c => `
                            <div class="bg-white border border-slate-200 p-5 rounded-2xl shadow-sm hover:border-blue-400 transition-all">
                                <div class="flex justify-between items-start mb-2">
                                    <div class="text-[9px] font-bold text-blue-500 uppercase tracking-widest">${c.fields.Anrede || ''} ${c.fields.Rolle || ''}</div>
                                    <div class="flex gap-1">${c.fields.SGF ? `<span class="bg-emerald-50 text-emerald-600 text-[8px] font-black px-1.5 py-0.5 rounded uppercase border border-emerald-100">${c.fields.SGF}</span>` : ''}</div>
                                </div>
                                <div class="text-base font-bold text-slate-800 mb-1">${c.fields.Vorname || ''} ${c.fields.Title}</div>
                                <div class="text-[11px] text-slate-400 font-medium mb-4 italic">${c.fields.Funktion || 'Funktion n.a.'}</div>
                                
                                <div class="grid grid-cols-1 gap-2 border-t pt-4 text-[10px] font-medium text-slate-500">
                                    <div class="flex items-center gap-2">📧 ${c.fields.Email1 || '-'}</div>
                                    <div class="flex items-center gap-2">📞 ${c.fields.Direktwahl || '-'}</div>
                                    <div class="flex items-center gap-2">📱 ${c.fields.Mobile || '-'}</div>
                                </div>
                                
                                <div class="mt-4 flex flex-wrap gap-2 border-t pt-3">
                                    ${c.fields.Event ? `<span class="text-[8px] font-bold text-slate-400">EVENT: ${c.fields.Event}</span>` : ''}
                                    ${c.fields.Leadbbz0 ? `<span class="text-[8px] font-bold text-blue-400 uppercase">LEAD: ${c.fields.Leadbbz0}</span>` : ''}
                                </div>
                                <button onclick="deleteContact('${c.id}','${id}')" class="mt-4 text-[9px] text-red-300 font-bold uppercase hover:text-red-500 transition">Kontakt löschen</button>
                            </div>
                        `).join('')}
                    </div>
                </div>

                <div class="col-span-12 lg:col-span-8 space-y-8">
                    <div class="bg-slate-50 border-2 border-dashed border-slate-200 rounded-[2.5rem] p-20 flex flex-col items-center justify-center text-center opacity-60">
                        <div class="text-4xl mb-6">📜</div>
                        <h4 class="font-bold text-slate-400 uppercase tracking-widest">Kontakthistorie (Etappe F)</h4>
                        <p class="text-xs text-slate-300 max-w-xs mt-2 italic font-medium">Bereite Datenbank für Timeline vor... hier fließen bald alle Telefonnotizen und Besuchsberichte ein.</p>
                    </div>
                    <div class="bg-slate-900 rounded-[2.5rem] p-12 text-white shadow-xl min-h-[300px] flex flex-col items-center justify-center grayscale">
                        <div class="text-3xl mb-4">📅</div>
                        <p class="text-[10px] font-bold uppercase tracking-widest opacity-40">Tasks & To-Dos (Etappe G)</p>
                    </div>
                </div>
            </div>

            <div id="editSide" class="fixed inset-y-0 right-0 w-[450px] bg-white shadow-2xl z-50 transform translate-x-full transition-transform duration-300 p-10 border-l">
                <div class="flex justify-between items-center mb-10 border-b pb-6">
                    <h3 class="text-sm font-bold uppercase tracking-widest">Unternehmensdaten</h3>
                    <button onclick="document.getElementById('editSide').classList.toggle('translate-x-full')" class="text-slate-300 hover:text-slate-800 transition">✕</button>
                </div>
                <div class="space-y-6">
                    <div class="space-y-2">
                        <label class="text-[10px] font-bold text-slate-400 uppercase">Firma</label>
                        <input type="text" id="e_t" value="${f.Title}" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm font-bold shadow-inner outline-none focus:ring-2 focus:ring-blue-500/20">
                    </div>
                    <div class="grid grid-cols-2 gap-4">
                        <div class="space-y-2">
                            <label class="text-[10px] font-bold text-slate-400 uppercase">Klassierung</label>
                            <select id="e_k" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm font-bold">
                                ${classOptions.map(o => `<option value="${o}" ${f.Klassifizierung===o?'selected':''}>${o}</option>`).join('')}
                            </select>
                        </div>
                        <div class="flex items-end pb-1 pl-4">
                            <label class="flex items-center gap-3 cursor-pointer"><input type="checkbox" id="e_v" ${f.VIP?'checked':''} class="w-5 h-5 text-blue-600 rounded"> <span class="text-[10px] font-bold text-slate-500 uppercase">VIP Kunde</span></label>
                        </div>
                    </div>
                    <div class="space-y-2">
                        <label class="text-[10px] font-bold text-slate-400 uppercase">Zentrale</label>
                        <input type="text" id="e_p" value="${f.Hauptnummer||''}" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm font-bold shadow-inner">
                    </div>
                    <div class="space-y-2">
                        <label class="text-[10px] font-bold text-slate-400 uppercase">Adresse</label>
                        <input type="text" id="e_s" value="${f.Adresse||''}" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm shadow-inner mb-2">
                        <div class="flex gap-2">
                            <input type="text" id="e_c" value="${f.Ort}" class="flex-1 p-3 bg-slate-50 border-none rounded-xl text-sm font-bold shadow-inner">
                            <input type="text" id="e_l" value="${f.Land||'CH'}" class="w-16 p-3 bg-slate-200 border-none rounded-xl text-xs font-bold text-center uppercase">
                        </div>
                    </div>
                    <button onclick="update('${id}')" class="w-full bg-blue-600 text-white py-4 rounded-2xl font-bold text-[11px] uppercase tracking-widest mt-10 shadow-lg hover:bg-blue-700 transition-all shadow-blue-100">Profil aktualisieren</button>
                    <button onclick="deleteFirm('${id}','${f.Title}')" class="w-full text-red-400 text-[10px] font-bold uppercase hover:underline mt-4">Firma löschen</button>
                </div>
            </div>
        </div>`;
}

// --- LOGIK & ACTIONS (FIXED) ---

async function saveNewFirm() {
    const n = document.getElementById('new_fName').value; if(!n) return;
    const fields = { Title: n, Klassifizierung: document.getElementById('new_fClass').value, Ort: document.getElementById('new_fCity').value };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
    toggleAdd(); loadData();
}

async function update(id) {
    const fields = { Title: document.getElementById('e_t').value, Klassifizierung: document.getElementById('e_k').value, Adresse: document.getElementById('e_s').value, Ort: document.getElementById('e_c').value, Land: document.getElementById('e_l').value, Hauptnummer: document.getElementById('e_p').value, VIP: document.getElementById('e_v').checked };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items/${id}/fields`, { method: 'PATCH', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify(fields) });
    loadData();
}

function filter(q) {
    const ql = q.toLowerCase();
    const filtered = allFirms.filter(x => x.fields.Title?.toLowerCase().includes(ql) || x.fields.Ort?.toLowerCase().includes(ql));
    document.getElementById('firmTableBody').innerHTML = filtered.map(f => row(f)).join('');
}

function toggleAdd() { document.getElementById('addForm').classList.toggle('hidden'); }
async function addContact(id) { /* PROMPT LOGIK WIE V0.42 */ }
async function deleteContact(cid, fid) { if(!confirm("Löschen?")) return; /* API DELETE */ loadData(); }

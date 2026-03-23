// --- CONFIG & VERSION ---
const appVersion = "0.52";
const appName = "CRM bbz";
const config = { 
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a", 
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7", 
    siteSearch: "bbzsg.sharepoint.com:/sites/CRM" 
};

let allFirms = [], allContacts = [], currentSiteId = "", currentListId = "", contactListId = "";
let meta = { klassen: [], anreden: [], rollen: [], leads: [], sgf: [], events: [] };

const msalConfig = { auth: { clientId: config.clientId, authority: `https://login.microsoftonline.com/${config.tenantId}`, redirectUri: "https://markusbaechler.github.io/crm-spa/" }, cache: { cacheLocation: "localStorage" } };
const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["https://graph.microsoft.com/AllSites.Write", "https://graph.microsoft.com/AllSites.Read"] };

window.onload = async () => { 
    if(document.getElementById('footer-text')) document.getElementById('footer-text').innerHTML = `© 2026 ${appName} | <b>Build ${appVersion}</b>`;
    await msalInstance.handleRedirectPromise(); 
    checkAuthState(); 
};

function checkAuthState() {
    const acc = msalInstance.getAllAccounts();
    if (acc.length > 0) {
        document.getElementById('authBtn').innerText = "Logout";
        document.getElementById('authBtn').onclick = () => msalInstance.logoutRedirect({ account: acc[0] });
        loadData();
    } else {
        document.getElementById('authBtn').innerText = "Login";
        document.getElementById('authBtn').onclick = () => msalInstance.loginRedirect(loginRequest);
    }
}

async function loadData() {
    const content = document.getElementById('main-content');
    content.innerHTML = `<div class="p-20 text-center text-slate-400 font-bold uppercase tracking-widest animate-pulse">Synchronisiere ${appName}...</div>`;
    try {
        const token = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
        const h = { 'Authorization': `Bearer ${token}` };
        const s = await (await fetch(`https://graph.microsoft.com/v1.0/sites/${config.siteSearch}`, { headers: h })).json();
        currentSiteId = s.id;
        const l = await (await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists`, { headers: h })).json();
        currentListId = l.value.find(x => x.displayName === "CRMFirms").id;
        contactListId = l.value.find(x => x.displayName === "CRMContacts").id;

        const [cKlass, cAnrede, cRolle, cLead, cSGF, cEvent, fData, cData] = await Promise.all([
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/columns/Klassifizierung`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/columns/Anrede`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/columns/Rolle`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/columns/Leadbbz0`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/columns/SGF`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/columns/Event`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items?expand=fields`, { headers: h }).then(r => r.json()),
            fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items?expand=fields`, { headers: h }).then(r => r.json())
        ]);

        meta = {
            klassen: cKlass.choice?.choices || [], anreden: cAnrede.choice?.choices || [],
            rollen: cRolle.choice?.choices || [], leads: cLead.choice?.choices || [],
            sgf: cSGF.choice?.choices || [], events: cEvent.choice?.choices || []
        };
        allFirms = fData.value; allContacts = cData.value;
        renderFirms(); 
    } catch (e) { content.innerHTML = `<div class="p-10 text-red-500 font-bold italic">Sync-Fehler: ${e.message}</div>`; }
}

// --- VIEW: KONTAKT-DETAILSEITE (NEU!) ---
function renderContactDetail(id) {
    const c = allContacts.find(x => String(x.id) === String(id));
    const f = c.fields;
    const firm = allFirms.find(x => String(x.id) === String(f.FirmaLookupId));
    
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-5xl mx-auto animate-in slide-in-from-right duration-300">
            <div class="bg-white border rounded-3xl p-8 mb-8 shadow-sm flex justify-between items-center">
                <div class="flex items-center gap-6">
                    <button onclick="renderDetail('${f.FirmaLookupId}')" class="bg-slate-50 p-3 rounded-2xl text-slate-400 hover:bg-slate-100 transition text-xl">←</button>
                    <div>
                        <h2 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter">${f.Vorname || ''} ${f.Title}</h2>
                        <p class="text-slate-400 text-xs font-bold uppercase tracking-widest mt-1">Firma: <span class="text-blue-600 cursor-pointer" onclick="renderDetail('${f.FirmaLookupId}')">${firm ? firm.fields.Title : '-'}</span></p>
                    </div>
                </div>
                <button onclick="saveContactEdit('${id}')" class="bg-blue-600 text-white px-8 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest">Speichern</button>
            </div>

            <div class="grid grid-cols-12 gap-8">
                <div class="col-span-12 lg:col-span-4 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-3xl p-8 shadow-sm">
                        <h3 class="text-[10px] font-black text-slate-400 uppercase tracking-widest border-b pb-4 mb-6">Stammdaten Kontakt</h3>
                        <div class="space-y-4">
                            <div><label class="text-[9px] font-bold uppercase text-slate-400">Anrede</label>
                                <select id="ed_anrede" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic">
                                    <option value="">-</option>${meta.anreden.map(o => `<option value="${o}" ${f.Anrede===o?'selected':''}>${o}</option>`).join('')}
                                </select>
                            </div>
                            <div><label class="text-[9px] font-bold uppercase text-slate-400">Vorname</label><input type="text" id="ed_vn" value="${f.Vorname||''}" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic"></div>
                            <div><label class="text-[9px] font-bold uppercase text-slate-400">Nachname</label><input type="text" id="ed_nn" value="${f.Title||''}" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic"></div>
                            <div><label class="text-[9px] font-bold uppercase text-slate-400">Rolle</label>
                                <select id="ed_rolle" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic">
                                    <option value="">-</option>${meta.rollen.map(o => `<option value="${o}" ${f.Rolle===o?'selected':''}>${o}</option>`).join('')}
                                </select>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="col-span-12 lg:col-span-8 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-3xl p-8 shadow-sm grid grid-cols-2 gap-8">
                        <div class="space-y-4">
                            <h4 class="text-[10px] font-black uppercase text-slate-300 tracking-widest mb-2 italic">Kommunikation</h4>
                            <input type="text" id="ed_e1" value="${f.Email1||''}" placeholder="E-Mail" class="w-full p-3 bg-slate-50 border-none rounded-2xl text-sm">
                            <input type="text" id="ed_mo" value="${f.Mobile||''}" placeholder="Mobile" class="w-full p-3 bg-slate-50 border-none rounded-2xl text-sm font-bold">
                            <input type="text" id="ed_dw" value="${f.Direktwahl||''}" placeholder="Direktwahl" class="w-full p-3 bg-slate-50 border-none rounded-2xl text-sm">
                        </div>
                        <div class="space-y-4">
                            <h4 class="text-[10px] font-black uppercase text-slate-300 tracking-widest mb-2 italic">Merkmale</h4>
                            <select id="ed_sgf" class="w-full p-3 bg-slate-50 border-none rounded-2xl text-sm font-bold">
                                <option value="">SGF</option>${meta.sgf.map(o => `<option value="${o}" ${f.SGF===o?'selected':''}>${o}</option>`).join('')}
                            </select>
                            <select id="ed_lead" class="w-full p-3 bg-slate-50 border-none rounded-2xl text-sm font-bold">
                                <option value="">Lead bbz</option>${meta.leads.map(o => `<option value="${o}" ${f.Leadbbz0===o?'selected':''}>${o}</option>`).join('')}
                            </select>
                            <select id="ed_event" class="w-full p-3 bg-blue-50 border-none rounded-2xl text-sm font-bold text-blue-600">
                                <option value="">Aktuelles Event</option>${meta.events.map(o => `<option value="${o}" ${f.Event===o?'selected':''}>${o}</option>`).join('')}
                            </select>
                        </div>
                    </div>
                    <div class="bg-slate-50 border border-slate-200 rounded-3xl p-8 italic shadow-inner">
                        <label class="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-4 block">Event-History & Notizen</label>
                        <textarea id="ed_kom" class="w-full p-4 bg-white border-none rounded-2xl text-sm shadow-sm min-h-[150px] font-medium" placeholder="Hier alle historischen Events und Notizen festhalten...">${f.Kommentar || ''}</textarea>
                    </div>
                </div>
            </div>
        </div>`;
}

// --- FIRMENÜBERSICHT & DETAIL ---
function renderFirms() {
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in fade-in">
            <div class="flex justify-between items-center mb-8 border-b pb-6 border-slate-100">
                <div class="flex gap-10 items-end">
                    <h2 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter cursor-pointer" onclick="renderFirms()">Firmen</h2>
                    <nav class="flex gap-6 text-xs font-black uppercase tracking-widest pb-1">
                        <button class="text-blue-600 border-b-2 border-blue-600">Übersicht</button>
                        <button onclick="renderAllContacts()" class="text-slate-300 hover:text-slate-600 transition">Alle Kontakte</button>
                    </nav>
                </div>
                <div class="flex gap-4"><input type="text" onkeyup="filterFirms(this.value)" placeholder="Suche..." class="px-5 py-2.5 bg-slate-50 rounded-2xl text-sm w-80 shadow-inner font-bold italic outline-none">
                <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-6 py-2.5 rounded-2xl text-[10px] font-black uppercase">+ Firma</button></div>
            </div>
            <div id="firmList" class="grid grid-cols-1 md:grid-cols-3 gap-8">${allFirms.map(f => `
                <div onclick="renderDetail('${f.id}')" class="bg-white border p-8 rounded-[2.5rem] hover:shadow-2xl transition-all cursor-pointer border-t-[6px] ${f.fields.Klassifizierung === 'A' ? 'border-t-emerald-500' : 'border-t-slate-200'}">
                    <h3 class="font-black text-slate-800 text-xl uppercase italic tracking-tighter mb-4 underline">${f.fields.Title}</h3>
                    <div class="text-[11px] font-bold text-slate-400 uppercase tracking-widest italic mb-8">📍 ${f.fields.Ort || '-'}</div>
                    <div class="mt-auto pt-6 flex justify-between items-center border-t border-slate-50">
                        <span class="text-[10px] font-black bg-slate-50 px-3 py-1.5 rounded-xl border italic uppercase">${f.fields.Klassifizierung || '-'}</span>
                        <span class="text-[10px] font-black text-blue-600 bg-blue-50 px-3 py-1.5 rounded-xl uppercase italic">👥 ${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(f.id)).length}</span>
                    </div>
                </div>`).join('')}</div>
        </div>`;
}

function renderDetail(id) {
    const firm = allFirms.find(x => String(x.id) === String(id)), f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(id));
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in slide-in-from-right duration-300 px-4">
            <div class="bg-white border rounded-2xl p-6 mb-8 flex justify-between items-center shadow-sm">
                <div class="flex items-center gap-6">
                    <button onclick="renderFirms()" class="bg-slate-50 p-2 rounded-xl text-slate-400">←</button>
                    <h2 class="text-2xl font-black text-slate-800 uppercase italic tracking-tighter">${f.Title}</h2>
                </div>
                <button onclick="toggleContactForm()" class="bg-blue-600 text-white px-5 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest shadow-lg shadow-blue-100">+ Kontakt</button>
            </div>

            <div id="addContactForm" class="hidden bg-white border border-blue-200 p-8 rounded-3xl mb-8 shadow-xl animate-in fade-in">
                <div class="grid grid-cols-1 md:grid-cols-4 gap-6 mb-8 border-b pb-8">
                    <div><label class="text-[9px] font-black uppercase text-slate-400 italic">Anrede</label>
                         <select id="c_anrede" class="w-full p-2 border rounded text-sm italic"><option value="">-</option>${meta.anreden.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                    </div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400 italic">Vorname</label><input type="text" id="c_vn" class="w-full p-2 border rounded text-sm outline-none"></div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400 italic">Nachname *</label><input type="text" id="c_nn" class="w-full p-2 border rounded text-sm font-black outline-none"></div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400 italic">Rolle</label>
                         <select id="c_rolle" class="w-full p-2 border rounded text-sm font-bold italic"><option value="">-</option>${meta.rollen.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                    </div>
                </div>
                <div class="grid grid-cols-1 md:grid-cols-4 gap-6 mb-8 border-b pb-8">
                    <div><label class="text-[9px] font-black uppercase text-slate-400 italic">E-Mail</label><input type="text" id="c_email1" class="w-full p-2 border rounded text-sm"></div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400 italic">Mobile</label><input type="text" id="c_mo" class="w-full p-2 border rounded text-sm font-black"></div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400 italic">SGF</label>
                         <select id="c_sgf" class="w-full p-2 border rounded text-sm font-bold italic"><option value="">-</option>${meta.sgf.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                    </div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400 italic">Lead bbz</label>
                         <select id="c_lead" class="w-full p-2 border rounded text-sm font-bold italic"><option value="">-</option>${meta.leads.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                    </div>
                </div>
                <div class="flex gap-4">
                    <button onclick="saveContact('${id}')" class="bg-blue-600 text-white font-black text-[10px] uppercase px-8 py-3 rounded-2xl shadow-lg hover:bg-blue-700 transition-all tracking-widest">In SharePoint Speichern</button>
                    <button onclick="toggleContactForm()" class="text-slate-400 font-bold uppercase text-[10px] px-4 italic underline">Abbrechen</button>
                </div>
            </div>

            <div class="bg-white border rounded-[2rem] overflow-hidden shadow-sm">
                <table class="w-full text-left text-[11px] italic">
                    <thead class="bg-slate-50 border-b text-[9px] uppercase font-black text-slate-400 tracking-widest">
                        <tr><th class="p-4">Name</th><th class="p-4">Rolle / Funktion</th><th class="p-4">Events</th><th class="p-4">Kontaktinfo</th></tr>
                    </thead>
                    <tbody class="divide-y italic">
                        ${contacts.map(c => `
                            <tr class="hover:bg-slate-50 transition-colors group">
                                <td onclick="renderContactDetail('${c.id}')" class="p-4 font-bold text-slate-800 text-sm underline cursor-pointer group-hover:text-blue-600">${c.fields.Vorname || ''} ${c.fields.Title}</td>
                                <td class="p-4"><div class="font-bold text-blue-600 uppercase text-[9px] tracking-tighter italic">${c.fields.Rolle || '-'}</div><div class="text-slate-400 font-medium">${c.fields.Funktion || ''}</div></td>
                                <td class="p-4"><div class="flex flex-wrap gap-1">${c.fields.Event ? `<span class="px-1.5 py-0.5 bg-blue-50 text-blue-600 rounded text-[8px] font-black uppercase border border-blue-100">${c.fields.Event}</span>` : '-'}</div></td>
                                <td class="p-4 text-slate-500 font-bold">📧 ${c.fields.Email1 || '-'} <br> 📱 ${c.fields.Mobile || '-'}</td>
                            </tr>`).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
}

// --- LOGIK ACTIONS ---
async function saveContactEdit(contactId) {
    const fields = {
        Anrede: document.getElementById('ed_anrede

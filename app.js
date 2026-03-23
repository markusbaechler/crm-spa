// --- CONFIG & VERSION ---
const appVersion = "0.52";
const appName = "CRM bbz";
const config = { 
    clientId: "c4143c1e-33ea-4c4d-a410-58110f966d0a", 
    tenantId: "3643e7ab-d166-4e27-bd5f-c5bbfcd282d7", 
    siteSearch: "bbzsg.sharepoint.com:/sites/CRM" 
};

// Globaler State
let allFirms = [], allContacts = [], currentSiteId = "", currentListId = "", contactListId = "";
let meta = { klassen: [], anreden: [], rollen: [], leads: [], sgf: [], events: [] };

const msalConfig = { 
    auth: { clientId: config.clientId, authority: `https://login.microsoftonline.com/${config.tenantId}`, redirectUri: "https://markusbaechler.github.io/crm-spa/" }, 
    cache: { cacheLocation: "localStorage" } 
};
const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["https://graph.microsoft.com/AllSites.Write", "https://graph.microsoft.com/AllSites.Read"] };

// --- INITIALISIERUNG ---
window.onload = async () => { 
    updateFooter();
    try {
        await msalInstance.handleRedirectPromise(); 
        checkAuthState(); 
    } catch (e) { console.error("Ladefehler:", e); }
};

function updateFooter() {
    const ft = document.getElementById('footer-text');
    if (ft) ft.innerHTML = `© 2026 ${appName} | <b>Build ${appVersion}</b>`;
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
    content.innerHTML = `<div class="p-20 text-center text-slate-400 font-bold uppercase tracking-widest animate-pulse italic">Lade Datenbank-Intelligenz...</div>`;
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
            klassen: cKlass.choice?.choices || [],
            anreden: cAnrede.choice?.choices || [],
            rollen: cRolle.choice?.choices || [],
            leads: cLead.choice?.choices || [],
            sgf: cSGF.choice?.choices || [],
            events: cEvent.choice?.choices || []
        };
        allFirms = fData.value; allContacts = cData.value;
        renderFirms(); 
    } catch (e) { content.innerHTML = `<div class="p-10 text-red-500">Sync-Fehler: ${e.message}</div>`; }
}

// --- NAVIGATION ---
function renderFirms() {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in fade-in">
            <div class="flex justify-between items-center mb-8 border-b pb-6 border-slate-100">
                <div class="flex gap-10 items-end">
                    <h2 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter cursor-pointer" onclick="renderFirms()">Firmen</h2>
                    <nav class="flex gap-6 text-xs font-black uppercase tracking-widest pb-1">
                        <button class="text-blue-600 border-b-2 border-blue-600 px-1">Übersicht</button>
                        <button onclick="renderAllContacts()" class="text-slate-300 hover:text-slate-600 transition">Alle Kontakte</button>
                    </nav>
                </div>
                <div class="flex gap-4">
                    <input type="text" onkeyup="filterFirms(this.value)" placeholder="Suche..." class="px-5 py-2.5 bg-slate-50 rounded-2xl text-sm w-80 shadow-inner font-bold italic outline-none">
                    <button onclick="toggleAddForm()" class="bg-blue-600 text-white px-6 py-2.5 rounded-xl text-[10px] font-black uppercase">+ Firma</button>
                </div>
            </div>
            <div id="addForm" class="hidden mb-8 p-6 bg-slate-50 border rounded-2xl grid grid-cols-1 md:grid-cols-3 gap-4 shadow-inner">
                <input type="text" id="new_fName" placeholder="Name *" class="p-2 border rounded text-sm">
                <select id="new_fClass" class="p-2 border rounded text-sm"><option value="">- Klasse -</option>${meta.klassen.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                <input type="text" id="new_fCity" placeholder="Ort" class="p-2 border rounded text-sm">
                <button onclick="saveNewFirm()" class="bg-green-600 text-white font-bold text-xs p-2 rounded uppercase shadow-md">In SharePoint speichern</button>
            </div>
            <div id="firmList" class="grid grid-cols-1 md:grid-cols-3 gap-8">
                ${allFirms.map(f => `
                    <div onclick="renderDetail('${f.id}')" class="bg-white border p-8 rounded-[2.5rem] hover:shadow-2xl transition-all cursor-pointer border-t-[6px] ${f.fields.Klassifizierung === 'A' ? 'border-t-emerald-500' : 'border-t-slate-200'}">
                        <h3 class="font-black text-slate-800 text-xl uppercase italic tracking-tighter mb-4 underline">${f.fields.Title}</h3>
                        <div class="text-[11px] font-bold text-slate-400 uppercase tracking-widest italic mb-8">📍 ${f.fields.Ort || '-'}</div>
                        <div class="mt-auto pt-6 flex justify-between items-center border-t border-slate-50">
                            <span class="text-[10px] font-black bg-slate-50 px-3 py-1.5 rounded-xl border italic uppercase">${f.fields.Klassifizierung || '-'}</span>
                            <span class="text-[10px] font-black text-blue-600 bg-blue-50 px-3 py-1.5 rounded-xl uppercase italic">👥 ${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(f.id)).length}</span>
                        </div>
                    </div>`).join('')}
            </div>
        </div>`;
}

function renderAllContacts() {
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in fade-in">
            <div class="flex justify-between items-center mb-8 border-b pb-6 border-slate-100">
                <div class="flex gap-10 items-end">
                    <h2 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter cursor-pointer" onclick="renderFirms()">Kontakte</h2>
                    <nav class="flex gap-6 text-xs font-black uppercase tracking-widest pb-1">
                        <button onclick="renderFirms()" class="text-slate-300 hover:text-slate-600 transition">Firmen</button>
                        <button class="text-blue-600 border-b-2 border-blue-600 px-1">Alle Kontakte</button>
                    </nav>
                </div>
                <input type="text" onkeyup="filterContacts(this.value)" placeholder="Person suchen..." class="px-5 py-2.5 bg-slate-50 rounded-2xl text-sm w-96 shadow-inner font-bold italic outline-none">
            </div>
            <div class="bg-white border rounded-[2rem] overflow-hidden shadow-sm">
                <table class="w-full text-left text-[11px] italic">
                    <thead class="bg-slate-50 border-b text-[9px] uppercase font-black text-slate-400 tracking-widest">
                        <tr><th class="p-6">Name / Firma</th><th class="p-6">Rolle / SGF</th><th class="p-6">Events (Badges)</th><th class="p-6">Kontaktinfo</th></tr>
                    </thead>
                    <tbody id="contactTableBody" class="divide-y divide-slate-50 italic">
                        ${allContacts.map(c => {
                            const firm = allFirms.find(x => String(x.id) === String(c.fields.FirmaLookupId));
                            return `<tr class="hover:bg-blue-50/30 transition-all group">
                                <td class="p-6"><div class="text-base font-black text-slate-800 uppercase group-hover:text-blue-600 cursor-pointer underline" onclick="renderContactDetail('${c.id}')">${c.fields.Vorname || ''} ${c.fields.Title}</div>
                                <div class="text-[11px] font-bold text-slate-400 mt-1 cursor-pointer" onclick="renderDetail('${c.fields.FirmaLookupId}')">🏢 ${firm ? firm.fields.Title : '-'}</div></td>
                                <td class="p-6"><div class="text-[10px] font-black text-blue-500 uppercase">${c.fields.Rolle || '-'}</div><div class="text-[9px] font-black bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded border border-emerald-100 mt-1 inline-block uppercase">${c.fields.SGF || '-'}</div></td>
                                <td class="p-6"><div class="flex flex-wrap gap-1">${c.fields.Event ? `<span class="px-2 py-0.5 bg-blue-50 text-blue-600 border border-blue-100 rounded text-[8px] font-black uppercase">${c.fields.Event}</span>` : '-'}</div></td>
                                <td class="p-6 font-bold text-slate-500">📧 ${c.fields.Email1 || '-'}<br>📱 ${c.fields.Mobile || '-'}</td>
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
}

function renderDetail(id) {
    const firm = allFirms.find(x => String(x.id) === String(id)), f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(id));
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-[1600px] mx-auto animate-in slide-in-from-right duration-300">
            <div class="bg-white border rounded-2xl p-6 mb-8 flex justify-between items-center shadow-sm">
                <div class="flex items-center gap-6"><button onclick="renderFirms()" class="bg-slate-50 p-2 rounded-xl text-slate-400">←</button>
                <h2 class="text-2xl font-black text-slate-800 uppercase italic tracking-tighter">${f.Title}</h2></div>
                <button onclick="toggleContactForm()" class="bg-blue-600 text-white px-5 py-2.5 rounded-xl text-[10px] font-black uppercase">+ Kontakt</button>
            </div>
            
            <div id="addContactForm" class="hidden bg-white border border-blue-100 p-8 rounded-3xl mb-8 shadow-xl animate-in fade-in">
                <div class="grid grid-cols-1 md:grid-cols-4 gap-6 mb-6">
                    <div><label class="text-[9px] font-black uppercase text-slate-400">Anrede</label>
                        <select id="c_anrede" class="w-full p-2 border rounded text-sm"><option value="">- leer -</option>${meta.anreden.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                    </div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400">Vorname</label><input type="text" id="c_vn" class="w-full p-2 border rounded text-sm"></div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400">Nachname *</label><input type="text" id="c_nn" class="w-full p-2 border rounded text-sm font-black"></div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400">Rolle</label>
                        <select id="c_rolle" class="w-full p-2 border rounded text-sm font-bold"><option value="">- leer -</option>${meta.rollen.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                    </div>
                </div>
                <div class="grid grid-cols-1 md:grid-cols-4 gap-6 mb-6">
                    <div><label class="text-[9px] font-black uppercase text-slate-400">E-Mail</label><input type="text" id="c_email1" class="w-full p-2 border rounded text-sm"></div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400">Mobile</label><input type="text" id="c_mo" class="w-full p-2 border rounded text-sm font-black"></div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400">SGF</label>
                        <select id="c_sgf" class="w-full p-2 border rounded text-sm font-bold"><option value="">- leer -</option>${meta.sgf.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                    </div>
                    <div><label class="text-[9px] font-black uppercase text-slate-400">Lead bbz</label>
                        <select id="c_lead" class="w-full p-2 border rounded text-sm font-bold"><option value="">- leer -</option>${meta.leads.map(o => `<option value="${o}">${o}</option>`).join('')}</select>
                    </div>
                </div>
                <button onclick="saveContact('${id}')" class="bg-blue-600 text-white font-black text-[10px] uppercase px-8 py-3 rounded-2xl shadow-lg">In SharePoint Speichern</button>
                <button onclick="toggleContactForm()" class="text-slate-400 font-bold uppercase text-[10px] px-4 underline">Abbrechen</button>
            </div>

            <div class="bg-white border rounded-[2rem] overflow-hidden shadow-sm">
                <table class="w-full text-left text-[11px] italic">
                    <thead class="bg-slate-50 border-b text-[9px] font-black text-slate-400 uppercase tracking-widest italic">
                        <tr><th class="p-4">Name</th><th class="p-4">Rolle / Funktion</th><th class="p-4">Events</th><th class="p-4">Kontaktinfo</th></tr>
                    </thead>
                    <tbody class="divide-y italic">
                        ${contacts.map(c => `
                            <tr class="hover:bg-slate-50 transition-colors group">
                                <td onclick="renderContactDetail('${c.id}')" class="p-4 font-bold text-slate-800 text-sm underline cursor-pointer group-hover:text-blue-600 uppercase tracking-tighter italic">${c.fields.Vorname || ''} ${c.fields.Title}</td>
                                <td class="p-4"><div class="font-bold text-blue-600 uppercase text-[9px] tracking-widest">${c.fields.Rolle || '-'}</div></td>
                                <td class="p-4"><div class="flex flex-wrap gap-1">${c.fields.Event ? `<span class="px-1.5 py-0.5 bg-blue-50 text-blue-600 rounded text-[8px] font-black uppercase border border-blue-100">${c.fields.Event}</span>` : '-'}</div></td>
                                <td class="p-4 text-slate-500 font-bold">📧 ${c.fields.Email1 || '-'} <br> 📱 ${c.fields.Mobile || '-'}</td>
                            </tr>`).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
}

function renderContactDetail(id) {
    const c = allContacts.find(x => String(x.id) === String(id));
    const f = c.fields, firm = allFirms.find(x => String(x.id) === String(f.FirmaLookupId));
    document.getElementById('main-content').innerHTML = `
        <div class="max-w-5xl mx-auto animate-in slide-in-from-right duration-300">
            <div class="bg-white border rounded-3xl p-8 mb-8 shadow-sm flex justify-between items-center">
                <div class="flex items-center gap-6"><button onclick="renderDetail('${f.FirmaLookupId}')" class="bg-slate-50 p-3 rounded-2xl text-xl text-slate-400">←</button>
                <div><h2 class="text-3xl font-black text-slate-800 uppercase italic tracking-tighter">${f.Vorname || ''} ${f.Title}</h2><p class="text-slate-400 text-[10px] font-black uppercase mt-1">Firma: <span class="text-blue-600" onclick="renderDetail('${f.FirmaLookupId}')">${firm ? firm.fields.Title : '-'}</span></p></div></div>
                <button onclick="saveContactEdit('${id}')" class="bg-blue-600 text-white px-8 py-3 rounded-2xl text-[10px] font-black uppercase">Änderungen speichern</button>
            </div>
            <div class="grid grid-cols-12 gap-8 italic">
                <div class="col-span-4 bg-white border rounded-3xl p-8 shadow-sm space-y-4">
                    <h3 class="text-[10px] font-black text-slate-400 uppercase tracking-widest border-b pb-4 mb-4">Profil Kontakt</h3>
                    <select id="ed_anrede" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic"><option value="">-</option>${meta.anreden.map(o => `<option value="${o}" ${f.Anrede===o?'selected':''}>${o}</option>`).join('')}</select>
                    <input type="text" id="ed_vn" value="${f.Vorname||''}" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic" placeholder="Vorname">
                    <input type="text" id="ed_nn" value="${f.Title||''}" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic" placeholder="Nachname">
                    <select id="ed_rolle" class="w-full p-2.5 bg-slate-50 border-none rounded-xl text-sm font-bold italic"><option value="">-</option>${meta.rollen.map(o => `<option value="${o}" ${f.Rolle===o?'selected':''}>${o}</option>`).join('')}</select>
                </div>
                <div class="col-span-8 space-y-6">
                    <div class="bg-white border rounded-3xl p-8 shadow-sm grid grid-cols-2 gap-6">
                        <input type="text" id="ed_e1" value="${f.Email1||''}" placeholder="E-Mail Geschäft" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm">
                        <input type="text" id="ed_mo" value="${f.Mobile||''}" placeholder="Mobile" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm font-bold">
                        <select id="ed_sgf" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm font-bold italic"><option value="">SGF</option>${meta.sgf.map(o => `<option value="${o}" ${f.SGF===o?'selected':''}>${o}</option>`).join('')}</select>
                        <select id="ed_lead" class="w-full p-3 bg-slate-50 border-none rounded-xl text-sm font-bold italic"><option value="">Lead bbz</option>${meta.leads.map(o => `<option value="${o}" ${f.Leadbbz0===o?'selected':''}>${o}</option>`).join('')}</select>
                        <select id="ed_event" class="w-full p-3 bg-blue-50 border-none rounded-xl text-sm font-bold italic text-blue-600"><option value="">Aktuelles Event</option>${meta.events.map(o => `<option value="${o}" ${f.Event===o?'selected':''}>${o}</option>`).join('')}</select>
                    </div>
                    <div class="bg-slate-50 border border-slate-200 rounded-3xl p-8 italic shadow-inner">
                        <label class="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-4 block">Event-History / Kommentar</label>
                        <textarea id="ed_kom" class="w-full p-4 bg-white border-none rounded-2xl text-sm shadow-sm min-h-[150px] font-medium" placeholder="Historie...">${f.Kommentar || ''}</textarea>
                    </div>
                </div>
            </div>
        </div>`;
}

// --- ACTIONS ---
async function saveContactEdit(id) {
    const fields = { Anrede: document.getElementById('ed_anrede').value, Vorname: document.getElementById('ed_vn').value, Title: document.getElementById('ed_nn').value, Rolle: document.getElementById('ed_rolle').value, Email1: document.getElementById('ed_e1').value, Mobile: document.getElementById('ed_mo').value, SGF: document.getElementById('ed_sgf').value, Leadbbz0: document.getElementById('ed_lead').value, Event: document.getElementById('ed_event').value, Kommentar: document.getElementById('ed_kom').value };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items/${id}/fields`, { method: 'PATCH', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify(fields) });
    loadData();
}

async function saveContact(firmId) {
    const nn = document.getElementById('c_nn').value; if(!nn) return alert("Nachname fehlt!");
    const fields = { Title: nn, Vorname: document.getElementById('c_vn').value, Anrede: document.getElementById('c_anrede').value, Rolle: document.getElementById('c_rolle').value, Email1: document.getElementById('c_email1').value, Mobile: document.getElementById('c_mo').value, SGF: document.getElementById('c_sgf').value, Leadbbz0: document.getElementById('c_lead').value, FirmaLookupId: firmId };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) });
    loadData();
}

async function saveNewFirm() {
    const n = document.getElementById('new_fName').value; if(!n) return;
    const f = { Title: n, Klassifizierung: document.getElementById('new_fClass').value, Ort: document.getElementById('new_fCity').value };
    const t = (await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] })).accessToken;
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${currentListId}/items`, { method: 'POST', headers: { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields: f }) });
    loadData();
}

function filterFirms(q) { const ql = q.toLowerCase(); const f = allFirms.filter(x => x.fields.Title?.toLowerCase().includes(ql) || x.fields.Ort?.toLowerCase().includes(ql)); document.getElementById('firmList').innerHTML = f.map(x => `<div onclick="renderDetail('${x.id}')" class="bg-white border p-8 rounded-[2.5rem] hover:shadow-2xl transition-all cursor-pointer border-t-[6px] ${x.fields.Klassifizierung === 'A' ? 'border-t-emerald-500' : 'border-t-slate-200'}"><h3 class="font-black text-slate-800 text-xl uppercase italic tracking-tighter mb-4 underline">${x.fields.Title}</h3><div class="text-[11px] font-bold text-slate-400 uppercase tracking-widest italic mb-8">📍 ${x.fields.Ort || '-'}</div><div class="mt-auto pt-6 flex justify-between items-center border-t border-slate-50"><span class="text-[10px] font-black bg-slate-50 px-3 py-1.5 rounded-xl border italic uppercase">${x.fields.Klassifizierung || '-'}</span><span class="text-[10px] font-black text-blue-600 bg-blue-50 px-3 py-1.5 rounded-xl uppercase italic">👥 ${allContacts.filter(c => String(c.fields.FirmaLookupId) === String(x.id)).length}</span></div></div>`).join(''); }
function filterContacts(q) { const ql = q.toLowerCase(); const f = allContacts.filter(c => c.fields.Title?.toLowerCase().includes(ql) || c.fields.Vorname?.toLowerCase().includes(ql)); document.getElementById('contactTableBody').innerHTML = f.map(c => { const firm = allFirms.find(x => String(x.id) === String(c.fields.FirmaLookupId)); return `<tr class="hover:bg-blue-50/30 transition-all group"><td class="p-6"><div class="text-base font-black text-slate-800 uppercase group-hover:text-blue-600 cursor-pointer underline" onclick="renderContactDetail('${c.id}')">${c.fields.Vorname || ''} ${c.fields.Title}</div><div class="text-[11px] font-bold text-slate-400 mt-1 cursor-pointer" onclick="renderDetail('${c.fields.FirmaLookupId}')">🏢 ${firm ? firm.fields.Title : '-'}</div></td><td class="p-6"><div class="text-[10px] font-black text-blue-500 uppercase">${c.fields.Rolle || '-'}</div><div class="text-[9px] font-black bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded border border-emerald-100 mt-1 inline-block uppercase">${c.fields.SGF || '-'}</div></td><td class="p-6"><div class="flex flex-wrap gap-1">${c.fields.Event ? `<span class="px-2 py-0.5 bg-blue-50 text-blue-600 border border-blue-100 rounded text-[8px] font-black uppercase">${c.fields.Event}</span>` : '-'}</div></td><td class="p-6 font-bold text-slate-500">📧 ${c.fields.Email1 || '-'}<br>📱 ${c.fields.Mobile || '-'}</td></tr>`; }).join(''); }

function toggleContactForm() { document.getElementById('addContactForm').classList.toggle('hidden'); }
function toggleAddForm() { document.getElementById('addForm').classList.toggle('hidden'); }

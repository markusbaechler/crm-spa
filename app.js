// --- CONFIG & VERSION ---
const appVersion = "0.38";
console.log(`CRM App ${appVersion} - Final Field Mapping`);

// ... (Identische MSAL Konfiguration wie zuvor) ...

function renderFirmDetailPage(itemId) {
    const firm = allFirms.find(f => String(f.id) === String(itemId));
    const f = firm.fields;
    
    // NATIVE LOOKUP FILTER (Nutzt das Feld 'Firma' aus deinem Screenshot)
    const contacts = allContacts.filter(c => 
        String(c.fields.FirmaLookupId) === String(itemId)
    );
    
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-7xl mx-auto animate-in slide-in-from-right duration-300">
            <div class="flex items-center justify-between mb-8">
                <div class="flex items-center gap-4">
                    <button onclick="renderFirms(allFirms)" class="text-slate-400 hover:text-blue-600 transition text-2xl">←</button>
                    <div>
                        <h2 class="text-2xl font-semibold text-slate-800">${f.Title}</h2>
                        <p class="text-xs text-slate-400 uppercase tracking-widest mt-1">${f.Ort || 'Kein Standort'}</p>
                    </div>
                </div>
                <div class="flex gap-2">
                     <button onclick="deleteFirm('${itemId}', '${f.Title}')" class="text-red-400 text-[10px] font-bold uppercase px-4 py-2 hover:bg-red-50 rounded-lg transition">Firma löschen</button>
                </div>
            </div>

            <div class="grid grid-cols-1 lg:grid-cols-12 gap-8">
                
                <div class="lg:col-span-4 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-xl p-6 shadow-sm">
                        <h3 class="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-6 border-b pb-2">Firmendetails</h3>
                        
                        <div class="space-y-4">
                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase">Offizieller Name</label>
                                <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full mt-1 p-2 bg-slate-50 border border-slate-100 rounded text-sm focus:border-blue-400 outline-none font-medium">
                            </div>

                            <div class="grid grid-cols-2 gap-3">
                                <div>
                                    <label class="text-[9px] font-bold text-slate-400 uppercase">Klassierung</label>
                                    <select id="edit_Klass" class="w-full mt-1 p-2 bg-slate-50 border border-slate-100 rounded text-sm">
                                        ${classOptions.map(opt => `<option value="${opt}" ${f.Klassifizierung === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                                    </select>
                                </div>
                                <div class="flex items-end pb-2">
                                    <label class="flex items-center gap-2 cursor-pointer ml-2">
                                        <input type="checkbox" id="edit_VIP" ${(f.VIP === true || f.VIP === "true") ? 'checked' : ''} class="w-4 h-4 rounded text-blue-600">
                                        <span class="text-[10px] font-bold text-slate-500 uppercase">VIP</span>
                                    </label>
                                </div>
                            </div>

                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase">Zentrale Telefonnummer</label>
                                <input type="text" id="edit_Phone" value="${f.Hauptnummer || ''}" placeholder="+41..." class="w-full mt-1 p-2 bg-slate-50 border border-slate-100 rounded text-sm">
                            </div>

                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase">Adresse / Land</label>
                                <input type="text" id="edit_Street" value="${f.Adresse || ''}" placeholder="Strasse" class="w-full mt-1 p-2 bg-slate-50 border border-slate-100 rounded text-sm mb-2">
                                <div class="flex gap-2">
                                    <input type="text" id="edit_City" value="${f.Ort || ''}" placeholder="Ort" class="flex-1 p-2 bg-slate-50 border border-slate-100 rounded text-sm font-medium">
                                    <input type="text" id="edit_Country" value="${f.Land || 'CH'}" class="w-12 p-2 bg-slate-50 border border-slate-100 rounded text-sm text-center">
                                </div>
                            </div>

                            <button onclick="updateFirm('${itemId}')" class="w-full bg-slate-800 text-white py-2.5 rounded-lg text-[11px] font-bold uppercase tracking-wider hover:bg-blue-700 transition shadow-md">Änderungen speichern</button>
                        </div>
                    </div>
                </div>

                <div class="lg:col-span-8 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-xl p-6 shadow-sm min-h-[400px]">
                        <div class="flex justify-between items-center mb-8">
                            <h3 class="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Ansprechpartner (${contacts.length})</h3>
                            <button onclick="addContact('${itemId}')" class="bg-blue-50 text-blue-600 text-[10px] font-bold uppercase px-3 py-1.5 rounded hover:bg-blue-100 transition tracking-wider">+ Kontakt hinzufügen</button>
                        </div>

                        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                            ${contacts.length > 0 ? contacts.map(c => `
                                <div class="bg-slate-50 border border-slate-100 rounded-xl p-4 hover:border-blue-200 transition-colors group relative shadow-sm">
                                    <div class="flex justify-between">
                                        <div class="text-[10px] font-bold text-blue-600 uppercase mb-1 tracking-tight">${c.fields.Anrede || ''}</div>
                                        <button onclick="deleteContact('${c.id}', '${itemId}')" class="text-slate-200 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity">✕</button>
                                    </div>
                                    <div class="text-base font-semibold text-slate-800 leading-tight">${c.fields.Vorname || ''} ${c.fields.Title}</div>
                                    <div class="text-[11px] text-slate-500 font-medium mb-3">${c.fields.Funktion || 'Funktion n.a.'}</div>
                                    
                                    <div class="space-y-1.5 border-t border-slate-200 pt-3">
                                        <div class="flex items-center gap-2 text-[10px] text-slate-600">
                                            <span class="text-slate-300 w-4 italic">📧</span> ${c.fields.Email1 || '-'}
                                        </div>
                                        <div class="flex items-center gap-2 text-[10px] text-slate-600">
                                            <span class="text-slate-300 w-4 italic">📞</span> ${c.fields.Direktwahl || '-'}
                                        </div>
                                        <div class="flex items-center gap-2 text-[10px] text-slate-600">
                                            <span class="text-slate-300 w-4 italic">📱</span> ${c.fields.TelefonMobil || '-'}
                                        </div>
                                    </div>

                                    <div class="mt-3 flex flex-wrap gap-1">
                                        <span class="px-2 py-0.5 bg-white border border-slate-200 rounded text-[9px] font-bold text-slate-400 uppercase">${c.fields.Rolle || 'Keine Rolle'}</span>
                                        ${c.fields.SGF ? `<span class="px-2 py-0.5 bg-blue-100 text-blue-700 rounded text-[9px] font-bold uppercase">${c.fields.SGF}</span>` : ''}
                                    </div>
                                </div>
                            `).join('') : `
                                <div class="col-span-2 flex flex-col items-center justify-center py-20 bg-slate-50 border-2 border-dashed border-slate-100 rounded-xl">
                                    <p class="text-slate-300 text-[10px] font-bold uppercase tracking-widest">Keine Kontakte zugeordnet</p>
                                </div>
                            `}
                        </div>
                    </div>
                </div>

            </div>
        </div>
    `;
}

// REST UPDATE: Angepasst an die exakten Feldnamen aus deinem Screenshot
async function addContact(firmId) {
    const ln = prompt("Nachname (Pflicht):"); if (!ln) return;
    const vn = prompt("Vorname:");
    
    const fields = {
        Title: ln,                // Gemäss Screenshot #Name = Title
        Vorname: vn,              // Gemäss Screenshot #Name = Vorname
        FirmaLookupId: firmId     // SharePoint Lookup Feld
    };

    const tokenRes = await msalInstance.acquireTokenSilent({ ...loginRequest, account: msalInstance.getAllAccounts()[0] });
    await fetch(`https://graph.microsoft.com/v1.0/sites/${currentSiteId}/lists/${contactListId}/items`, {
        method: 'POST', 
        headers: { 'Authorization': `Bearer ${tokenRes.accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields: fields })
    });
    
    await loadDataSilent();
    renderFirmDetailPage(firmId);
}

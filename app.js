// --- CONFIG & VERSION ---
const appVersion = "0.39";
console.log(`CRM App ${appVersion} - Dashboard Redesign für History/Tasks`);

// ... (Konfiguration bleibt identisch zu V0.38) ...

function renderFirmDetailPage(itemId) {
    const firm = allFirms.find(f => String(f.id) === String(itemId));
    const f = firm.fields;
    const contacts = allContacts.filter(c => String(c.fields.FirmaLookupId) === String(itemId));
    
    const content = document.getElementById('main-content');
    content.innerHTML = `
        <div class="max-w-7xl mx-auto animate-in fade-in duration-500">
            
            <div class="bg-white border border-slate-200 rounded-2xl p-6 mb-8 shadow-sm flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
                <div class="flex items-center gap-5">
                    <button onclick="renderFirms(allFirms)" class="bg-slate-50 hover:bg-slate-100 text-slate-400 p-3 rounded-xl transition">←</button>
                    <div>
                        <div class="flex items-center gap-3">
                            <h2 class="text-2xl font-semibold text-slate-800">${f.Title}</h2>
                            ${(f.VIP === true || f.VIP === "true") ? '<span class="text-amber-400 text-xl">⭐</span>' : ''}
                        </div>
                        <p class="text-slate-400 text-xs font-medium uppercase tracking-widest mt-1">
                            ${f.Ort || 'Kein Ort'} <span class="mx-2 text-slate-200">|</span> ${f.Land || 'CH'} <span class="mx-2 text-slate-200">|</span> ${f.Hauptnummer || 'Keine Nummer'}
                        </p>
                    </div>
                </div>
                <div class="flex gap-3">
                    <button onclick="updateFirm('${itemId}')" class="bg-slate-800 text-white px-6 py-2.5 rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-slate-700 transition">Änderungen speichern</button>
                </div>
            </div>

            <div class="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
                
                <div class="lg:col-span-4 space-y-6">
                    <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
                        <h3 class="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-6 border-b pb-3">Stammdaten-Editor</h3>
                        <div class="space-y-5">
                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Firmenname</label>
                                <input type="text" id="edit_Title" value="${f.Title || ''}" class="w-full mt-1.5 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm outline-none focus:ring-2 focus:ring-blue-500/10 focus:border-blue-400 transition-all">
                            </div>
                            <div class="grid grid-cols-2 gap-4">
                                <div>
                                    <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Klassierung</label>
                                    <select id="edit_Klass" class="w-full mt-1.5 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm outline-none">
                                        ${classOptions.map(opt => `<option value="${opt}" ${f.Klassifizierung === opt ? 'selected' : ''}>${opt}</option>`).join('')}
                                    </select>
                                </div>
                                <div class="flex items-end pb-1 px-2">
                                    <label class="flex items-center gap-2 cursor-pointer group">
                                        <input type="checkbox" id="edit_VIP" ${(f.VIP === true || f.VIP === "true") ? 'checked' : ''} class="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-0">
                                        <span class="text-[10px] font-bold text-slate-400 group-hover:text-slate-600 uppercase transition-colors">VIP Kunde</span>
                                    </label>
                                </div>
                            </div>
                            <div>
                                <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Telefon Zentrale</label>
                                <input type="text" id="edit_Phone" value="${f.Hauptnummer || ''}" class="w-full mt-1.5 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm outline-none">
                            </div>
                            <div class="pt-2">
                                <label class="text-[9px] font-bold text-slate-400 uppercase tracking-tight">Standortadresse</label>
                                <input type="text" id="edit_Street" value="${f.Adresse || ''}" placeholder="Strasse / Nr." class="w-full mt-1.5 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm outline-none mb-3">
                                <div class="flex gap-3">
                                    <input type="text" id="edit_City" value="${f.Ort || ''}" placeholder="Ort" class="flex-1 p-2.5 bg-slate-50 border border-slate-100 rounded-xl text-sm font-medium">
                                    <input type="text" id="edit_Country" value="${f.Land || 'CH'}" placeholder="Land" class="w-16 p-2.5 bg-slate-100 border border-slate-100 rounded-xl text-sm text-center font-bold text-slate-400 uppercase">
                                </div>
                            </div>
                            <div class="pt-4 border-t border-slate-50">
                                <button onclick="deleteFirm('${itemId}', '${f.Title}')" class="text-red-400 text-[9px] font-bold uppercase hover:underline">Firma unwiderruflich löschen</button>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="lg:col-span-8 space-y-8">
                    
                    <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
                        <div class="flex justify-between items-center mb-6">
                            <h3 class="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Ansprechpartner (${contacts.length})</h3>
                            <button onclick="addContact('${itemId}')" class="text-blue-600 text-[10px] font-bold uppercase hover:bg-blue-50 px-3 py-1.5 rounded-lg transition">+ Hinzufügen</button>
                        </div>
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                            ${contacts.length > 0 ? contacts.map(c => `
                                <div class="p-4 bg-slate-50 border border-slate-100 rounded-2xl flex justify-between items-start group hover:border-blue-200 transition-all">
                                    <div>
                                        <div class="text-[10px] font-bold text-blue-500 uppercase tracking-tighter mb-1">${c.fields.Anrede || ''} ${c.fields.Rolle || ''}</div>
                                        <div class="text-sm font-semibold text-slate-800">${c.fields.Vorname || ''} ${c.fields.Title}</div>
                                        <div class="text-[10px] text-slate-400 mt-2 flex flex-col gap-1 font-medium">
                                            <span>📧 ${c.fields.Email1 || '-'}</span>
                                            <span>📞 ${c.fields.Direktwahl || '-'}</span>
                                        </div>
                                    </div>
                                    <button onclick="deleteContact('${c.id}', '${itemId}')" class="text-slate-200 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity p-1">✕</button>
                                </div>
                            `).join('') : '<p class="p-8 text-center text-slate-300 text-xs italic border border-dashed rounded-2xl">Keine Kontakte zugeordnet</p>'}
                        </div>
                    </div>

                    <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <div class="bg-slate-50/50 border-2 border-dashed border-slate-100 rounded-2xl p-12 flex flex-col items-center justify-center grayscale opacity-50">
                            <span class="text-2xl mb-2">📜</span>
                            <h4 class="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Aktivitätshistorie</h4>
                            <p class="text-[9px] text-slate-300 mt-1">Geplant für Etappe F</p>
                        </div>
                        <div class="bg-slate-50/50 border-2 border-dashed border-slate-100 rounded-2xl p-12 flex flex-col items-center justify-center grayscale opacity-50">
                            <span class="text-2xl mb-2">📅</span>
                            <h4 class="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Nächste Aufgaben</h4>
                            <p class="text-[9px] text-slate-300 mt-1">Geplant für Etappe G</p>
                        </div>
                    </div>

                </div>
            </div>
        </div>
    `;
}

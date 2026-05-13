const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  collabs:  () => SS.getSheetByName('Collaborateurs'),
  periodes: () => SS.getSheetByName('Periodes'),
  jours:    () => SS.getSheetByName('Jours'),
  admins:   () => SS.getSheetByName('Admins'),
  equipes:  () => SS.getSheetByName('Equipes'),
};
function fmtDate(d) { if (!d) return ''; try { return Utilities.formatDate(new Date(d), 'Europe/Paris', 'dd/MM/yyyy'); } catch(e) { return String(d); } }
function fmtHeure(v) {
  if (!v) return '';
  const s = String(v).trim();
  if (/^\d{2}:\d{2}$/.test(s)) return s;
  try {
    return Utilities.formatDate(new Date(v), 'Europe/Paris', 'HH:mm');
  } catch(e) { return ''; }
}
function genToken() { return Math.random().toString(36).slice(2, 10); }
function sheetToObjects(sheet) { if (!sheet) return []; const data = sheet.getDataRange().getValues(); if (data.length < 2) return []; const headers = data[0].map(h => String(h).trim()); return data.slice(1).map(row => { const obj = {}; headers.forEach((h,i) => { obj[h] = row[i]; }); return obj; }); }
function jsonpResponse(callback, data) { return ContentService.createTextOutput(callback + '(' + JSON.stringify(data) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT); }
function checkAdmin(token) { return sheetToObjects(SHEETS.admins()).find(r => String(r.token).trim() === String(token).trim() && r.actif === true); }
function checkManager(token) { return sheetToObjects(SHEETS.equipes()).find(r => String(r.manager_token).trim() === String(token).trim() && (r.actif === true || r.actif === 'TRUE')); }
function checkCollab(token) { return sheetToObjects(SHEETS.collabs()).find(r => String(r.token).trim() === String(token).trim() && r.actif === true); }
function doGet(e) {
  const p = e.parameter, action = p.action||'', cb = p.callback||'callback', token = p.token||'';
  try {
    switch(action) {
      case 'getJourDuJour': return jsonpResponse(cb, getJourDuJour(token));
      case 'getMesJours': return jsonpResponse(cb, getMesJours(token, p.periodeId));
      case 'saveHoraires': return jsonpResponse(cb, saveHoraires(token, p));
      case 'getJourById': return jsonpResponse(cb, getJourById(token, p.jourId));
      case 'bo_getCollabs': return jsonpResponse(cb, bo_getCollabs(token));
      case 'bo_saveCollab': return jsonpResponse(cb, bo_saveCollab(token, p));
      case 'bo_toggleCollab': return jsonpResponse(cb, bo_toggleCollab(token, p.collabId));
      case 'bo_getPeriodes': return jsonpResponse(cb, bo_getPeriodes(token));
      case 'bo_getRecap': return jsonpResponse(cb, bo_getRecap(token, p.periodeId));
      case 'bo_getEquipes': return jsonpResponse(cb, bo_getEquipes(token));
      case 'bo_getJoursCollab': return jsonpResponse(cb, bo_getJoursCollab(token, p.collabId, p.periodeId));
      case 'bo_saveRemarqueManager': return jsonpResponse(cb, bo_saveRemarqueManager(token, p.jourId, p.remarque));
      case 'bo_saveJour': return jsonpResponse(cb, bo_saveJour(token, p));
      case 'bo_getRecapManager': return jsonpResponse(cb, bo_getRecapManager(token, p.periodeId));
      case 'bo_savePeriode': return jsonpResponse(cb, bo_savePeriode(token, p));
      case 'bo_deletePeriode': return jsonpResponse(cb, bo_deletePeriode(token, p.periodeId));
      case 'bo_initRecapPaie': return jsonpResponse(cb, bo_initRecapPaie(token, p.periodeId));
      case 'bo_getRecapPaie': return jsonpResponse(cb, bo_getRecapPaie(token, p.periodeId));
      case 'bo_validerCollab': return jsonpResponse(cb, bo_validerCollab(token, p.periodeId, p.collabId, p.note));
      case 'bo_saveNoteAdmin': return jsonpResponse(cb, bo_saveNoteAdmin(token, p.periodeId, p.collabId, p.note));
      default: return jsonpResponse(cb, {ok:false, error:'Action inconnue: '+action});
    }
  } catch(err) { return jsonpResponse(cb, {ok:false, error:err.message}); }
}
function getJourDuJour(token) {
  const collab = checkCollab(token); if (!collab) return {ok:false, error:'Token invalide'};
  const today = new Date(), todayStr = Utilities.formatDate(today, 'Europe/Paris', 'yyyy-MM-dd');
  const JS = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
  const jour = sheetToObjects(SHEETS.jours()).find(j => String(j.collab_id)===String(collab.collab_id) && fmtDate(j.date_jour)===fmtDate(todayStr));
  if (!jour) return {ok:false, error:"Aucun jour trouve pour aujourd'hui"};
  return {ok:true, jour:{jourId:jour.jour_id,dateJour:fmtDate(jour.date_jour),jourSemaine:jour.jour_semaine||JS[today.getDay()],periodeId:jour.periode_id,typeJour:jour.type_jour,c1Debut:fmtHeure(jour.c1_debut),c1Fin:fmtHeure(jour.c1_fin),c2Debut:fmtHeure(jour.c2_debut),c2Fin:fmtHeure(jour.c2_fin),c3Debut:fmtHeure(jour.c3_debut),c3Fin:fmtHeure(jour.c3_fin),commentaire:jour.commentaire,totalHeures:jour.total_heures}, collab:{nomAffiche:collab.nom_affiche||(collab.prenom+' '+collab.nom),heuresHebdo:collab.heures_hebdo}};
}
function getMesJours(token, periodeId) {
  const collab = checkCollab(token); if (!collab) return {ok:false, error:'Token invalide'};
  return {ok:true, jours:sheetToObjects(SHEETS.jours()).filter(j=>String(j.collab_id)===String(collab.collab_id)&&(!periodeId||String(j.periode_id)===String(periodeId))).map(j=>({jourId:j.jour_id,dateJour:fmtDate(j.date_jour),jourSemaine:j.jour_semaine,typeJour:j.type_jour,c1Debut:fmtHeure(j.c1_debut),c1Fin:fmtHeure(j.c1_fin),c2Debut:fmtHeure(j.c2_debut),c2Fin:fmtHeure(j.c2_fin),c3Debut:fmtHeure(j.c3_debut),c3Fin:fmtHeure(j.c3_fin),commentaire:j.commentaire,totalHeures:j.total_heures,totalHebdoProg:j.total_hebdo_prog}))};
}
function getJourById(token, jourId) {
  const collab = checkCollab(token); if (!collab) return {ok:false, error:'Token invalide'};
  const jour = sheetToObjects(SHEETS.jours()).find(j=>String(j.jour_id)===String(jourId)&&String(j.collab_id)===String(collab.collab_id));
  if (!jour) return {ok:false, error:'Jour introuvable'};
  return {ok:true, jour:{jourId:jour.jour_id,dateJour:fmtDate(jour.date_jour),jourSemaine:jour.jour_semaine,typeJour:jour.type_jour,c1Debut:fmtHeure(jour.c1_debut),c1Fin:fmtHeure(jour.c1_fin),c2Debut:fmtHeure(jour.c2_debut),c2Fin:fmtHeure(jour.c2_fin),c3Debut:fmtHeure(jour.c3_debut),c3Fin:fmtHeure(jour.c3_fin),commentaire:jour.commentaire,totalHeures:jour.total_heures,nbModifications:jour.nb_modifications||0}};
}
function saveHoraires(token, params) {
  const collab = checkCollab(token); if (!collab) return {ok:false, error:'Token invalide'};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet=SHEETS.jours(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
    const idx=data.findIndex((row,i)=>i>0&&String(row[headers.indexOf('jour_id')])===String(params.jourId)&&String(row[headers.indexOf('collab_id')])===String(collab.collab_id));
    if(idx===-1)return{ok:false,error:'Jour introuvable'};
    const nbModifs=Number(sheet.getRange(idx+1,headers.indexOf('nb_modifications')+1).getValue())||0;
    if(nbModifs>=3)return{ok:false,error:'Nombre maximum de modifications atteint'};
    function diff(d,f){if(!d||!f)return 0;const[dh,dm]=String(d).split(':').map(Number);const[fh,fm]=String(f).split(':').map(Number);const m=(fh*60+fm)-(dh*60+dm);return m>0?m/60:0;}
    const tj=params.typeJour||'',c1d=params.c1Debut||'',c1f=params.c1Fin||'',c2d=params.c2Debut||'',c2f=params.c2Fin||'',c3d=params.c3Debut||'',c3f=params.c3Fin||'';
    const total=(tj==='travail'||tj==='travaillée'||tj==='travaille'||tj==='AT')?diff(c1d,c1f)+diff(c2d,c2f)+diff(c3d,c3f):0;
    const sc=(n,v)=>{const c=headers.indexOf(n);if(c>=0)sheet.getRange(idx+1,c+1).setValue(v);};
    sc('type_jour',tj);sc('c1_debut',c1d);sc('c1_fin',c1f);sc('c2_debut',c2d);sc('c2_fin',c2f);sc('c3_debut',c3d);sc('c3_fin',c3f);sc('commentaire',params.commentaire||'');sc('total_heures',total);sc('date_derniere_modif',Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy HH:mm'));
    const nm=headers.indexOf('nb_modifications');if(nm>=0){const cur=Number(sheet.getRange(idx+1,nm+1).getValue())||0;sheet.getRange(idx+1,nm+1).setValue(cur+1);}
    _recalculerSemaineCollab(collab.collab_id, data[idx][headers.indexOf('date_jour')], sheet, sheet.getDataRange().getValues(), headers);
    return{ok:true};
  } finally {
    lock.releaseLock();
  }
}
function _recalculerSemaineCollab(collabId, dateJour, sheet, data, headers) {
  // Trouver le lundi de la semaine du jour
  const d = new Date(dateJour);
  const jourSem = d.getDay() === 0 ? 6 : d.getDay() - 1;
  const lundi = new Date(d);
  lundi.setDate(d.getDate() - jourSem);
  lundi.setHours(0,0,0,0);
  const dimanche = new Date(lundi);
  dimanche.setDate(lundi.getDate() + 6);
  dimanche.setHours(23,59,59,999);

  // Trouver tous les jours de la semaine pour ce collab
  const idxCollabId = headers.indexOf('collab_id');
  const idxDateJour = headers.indexOf('date_jour');
  const idxTotalH = headers.indexOf('total_heures');
  const idxTotalProg = headers.indexOf('total_hebdo_prog');
  if (idxTotalProg < 0) return;

  // Collecter les lignes de la semaine, triées par date
  const lignesSemaine = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxCollabId]) !== String(collabId)) continue;
    const dj = new Date(data[i][idxDateJour]);
    if (isNaN(dj)) continue;
    dj.setHours(0,0,0,0);
    if (dj >= lundi && dj <= dimanche) {
      lignesSemaine.push({ row: i + 1, date: dj, totalH: Number(data[i][idxTotalH]) || 0 });
    }
  }

  // Trier par date croissante
  lignesSemaine.sort((a, b) => a.date - b.date);

  // Calculer le progressif et écrire en batch
  let cumul = 0;
  lignesSemaine.forEach(ligne => {
    cumul += ligne.totalH;
    sheet.getRange(ligne.row, idxTotalProg + 1).setValue(cumul);
  });
}
function bo_getCollabs(token) {
  if(!checkAdmin(token))return{ok:false,error:'Acces refuse'};
  const equipes=sheetToObjects(SHEETS.equipes());
  return{ok:true,collabs:sheetToObjects(SHEETS.collabs()).map(c=>{const eq=equipes.find(e=>String(e.equipe_id)===String(c.equipe_id));return{collabId:c.collab_id,prenom:c.prenom,nom:c.nom,nomAffiche:c.nom_affiche||(c.prenom+' '+c.nom),email:c.email,structure:c.structure,typeContrat:c.type_contrat,heuresHebdo:c.heures_hebdo,statut:c.statut,dateActivation:fmtDate(c.date_activation),actif:c.actif,token:c.token,equipeId:c.equipe_id||'',equipeNom:eq?eq.nom_equipe:'',lien:c.token?'https://chantsdelaterre.github.io/suivi-temps/?id='+c.token:''};})};
}
function bo_saveCollab(token, params) {
  if(!checkAdmin(token)) return {ok:false, error:'Acces refuse'};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet=SHEETS.collabs(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
    const sv=(row,n,v)=>{const c=headers.indexOf(n);if(c>=0)sheet.getRange(row,c+1).setValue(v);};
    if(params.collabId){
      const idx=data.findIndex((row,i)=>i>0&&String(row[headers.indexOf('collab_id')])===String(params.collabId));
      if(idx===-1)return{ok:false,error:'Collab introuvable'};const r=idx+1;
      sv(r,'prenom',params.prenom||'');sv(r,'nom',params.nom||'');sv(r,'nom_affiche',(params.prenom||'')+' '+(params.nom||''));
      sv(r,'email',params.email||'');sv(r,'structure',params.structure||'');sv(r,'type_contrat',params.typeContrat||'');
      sv(r,'heures_hebdo',params.heuresHebdo?Number(params.heuresHebdo):'');sv(r,'statut',params.statut||'en_attente');
      sv(r,'date_activation',params.dateActivation||'');sv(r,'equipe_id',params.equipeId||'');
      return{ok:true,action:'updated'};
    }else{
      if(!params.prenom||!params.nom)return{ok:false,error:'Prenom et nom requis'};
      const lastId=data.slice(1).reduce((max,row)=>{const id=String(row[headers.indexOf('collab_id')]||'').replace('COLL','');return Math.max(max,parseInt(id)||0);},0);
      const newId='COLL'+String(lastId+1).padStart(3,'0'),newToken=genToken();
      sheet.appendRow(headers.map(h=>{switch(h){case 'collab_id':return newId;case 'prenom':return params.prenom;case 'nom':return params.nom;case 'nom_affiche':return params.prenom+' '+params.nom;case 'email':return params.email||'';case 'structure':return params.structure||'';case 'type_contrat':return params.typeContrat||'';case 'heures_hebdo':return params.heuresHebdo?Number(params.heuresHebdo):'';case 'statut':return params.statut||'en_attente';case 'date_activation':return params.dateActivation||'';case 'actif':return params.statut==='actif';case 'role':return 'collab';case 'token':return newToken;case 'equipe_id':return params.equipeId||'';default:return '';}}));
      return{ok:true,action:'created',collabId:newId,token:newToken,lien:'https://chantsdelaterre.github.io/suivi-temps/?id='+newToken};
    }
  } finally {
    lock.releaseLock();
  }
}
function bo_toggleCollab(token,collabId){
  if(!checkAdmin(token))return{ok:false,error:'Acces refuse'};
  const sheet=SHEETS.collabs(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
  const idx=data.findIndex((row,i)=>i>0&&String(row[headers.indexOf('collab_id')])===String(collabId));
  if(idx===-1)return{ok:false,error:'Collab introuvable'};
  const sc=headers.indexOf('statut'),ac=headers.indexOf('actif'),ns=String(data[idx][sc])==='inactif'?'actif':'inactif';
  sheet.getRange(idx+1,sc+1).setValue(ns);if(ac>=0)sheet.getRange(idx+1,ac+1).setValue(ns==='actif');
  return{ok:true,newStatut:ns};
}
function bo_getPeriodes(token){
  if(!checkAdmin(token)&&!checkManager(token))return{ok:false,error:'Acces refuse'};
  return{ok:true,periodes:sheetToObjects(SHEETS.periodes()).map(p=>({periodeId:p.periode_id,nomPeriode:p.nom_periode,typePeriode:p.type_periode,dateDebut:fmtDate(p.date_debut),dateFin:fmtDate(p.date_fin),dateCloture:fmtDate(p.date_cloture),statut:p.statut}))};
}
function bo_getEquipes(token){
  if(!checkAdmin(token))return{ok:false,error:'Acces refuse'};
  return{ok:true,equipes:sheetToObjects(SHEETS.equipes()).map(e=>({equipeId:e.equipe_id,nomEquipe:e.nom_equipe,nomManager:e.nom_manager,actif:e.actif}))};
}
function bo_getRecap(token,periodeId){if(!checkAdmin(token))return{ok:false,error:'Acces refuse'};return _getRecapForCollabs(periodeId,null);}
function bo_getRecapManager(token,periodeId){const eq=checkManager(token);if(!eq)return{ok:false,error:'Acces refuse'};return _getRecapForCollabs(periodeId,String(eq.equipe_id));}
function bo_getJoursCollab(token,collabId,periodeId){
  const isAdmin=checkAdmin(token),manager=checkManager(token);
  if(!isAdmin&&!manager)return{ok:false,error:'Acces refuse'};
  if(manager){const c=sheetToObjects(SHEETS.collabs()).find(c=>String(c.collab_id)===String(collabId));if(!c||String(c.equipe_id)!==String(manager.equipe_id))return{ok:false,error:'Collab hors equipe'};}
  const jours=sheetToObjects(SHEETS.jours()).filter(j=>String(j.collab_id)===String(collabId)&&(!periodeId||String(j.periode_id)===String(periodeId))).sort((a,b)=>new Date(a.date_jour)-new Date(b.date_jour)).map(j=>({jourId:j.jour_id,dateJour:fmtDate(j.date_jour),jourSemaine:j.jour_semaine,typeJour:j.type_jour,c1Debut:fmtHeure(j.c1_debut),c1Fin:fmtHeure(j.c1_fin),c2Debut:fmtHeure(j.c2_debut),c2Fin:fmtHeure(j.c2_fin),c3Debut:fmtHeure(j.c3_debut),c3Fin:fmtHeure(j.c3_fin),commentaire:j.commentaire,totalHeures:j.total_heures,totalHebdoProg:j.total_hebdo_prog,remarqueManager:j.remarque_manager||'',dateDerniereModif:j.date_derniere_modif||'',nbModifications:j.nb_modifications||0}));
  const collab=sheetToObjects(SHEETS.collabs()).find(c=>String(c.collab_id)===String(collabId));
  return{ok:true,jours,collab:collab?{nomAffiche:collab.nom_affiche||(collab.prenom+' '+collab.nom),structure:collab.structure,typeContrat:collab.type_contrat,heuresHebdo:collab.heures_hebdo}:null};
}
function bo_saveRemarqueManager(token,jourId,remarque){
  const isAdmin=checkAdmin(token),manager=checkManager(token);
  if(!isAdmin&&!manager)return{ok:false,error:'Acces refuse'};
  const sheet=SHEETS.jours(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
  const idx=data.findIndex((row,i)=>i>0&&String(row[headers.indexOf('jour_id')])===String(jourId));
  if(idx===-1)return{ok:false,error:'Jour introuvable'};
  if(manager){const cid=data[idx][headers.indexOf('collab_id')];const c=sheetToObjects(SHEETS.collabs()).find(c=>String(c.collab_id)===String(cid));if(!c||String(c.equipe_id)!==String(manager.equipe_id))return{ok:false,error:'Hors equipe'};}
  const col=headers.indexOf('remarque_manager');if(col>=0)sheet.getRange(idx+1,col+1).setValue(remarque||'');
  return{ok:true};
}
function bo_saveJour(token, params) {
  if(!checkAdmin(token)) return {ok:false, error:'Acces refuse - admin uniquement'};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet=SHEETS.jours(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
    const idx=data.findIndex((row,i)=>i>0&&String(row[headers.indexOf('jour_id')])===String(params.jourId));
    if(idx===-1)return{ok:false,error:'Jour introuvable'};
    const sc=(n,v)=>{const c=headers.indexOf(n);if(c>=0)sheet.getRange(idx+1,c+1).setValue(v);};
    function diff(d,f){if(!d||!f)return 0;const[dh,dm]=String(d).split(':').map(Number);const[fh,fm]=String(f).split(':').map(Number);const m=(fh*60+fm)-(dh*60+dm);return m>0?m/60:0;}
    const tj=params.typeJour||'',c1d=params.c1Debut||'',c1f=params.c1Fin||'',c2d=params.c2Debut||'',c2f=params.c2Fin||'',c3d=params.c3Debut||'',c3f=params.c3Fin||'';
    sc('type_jour',tj);sc('c1_debut',c1d);sc('c1_fin',c1f);sc('c2_debut',c2d);sc('c2_fin',c2f);sc('c3_debut',c3d);sc('c3_fin',c3f);
    sc('commentaire',params.commentaire||'');
    sc('total_heures',(tj==='travaillée'||tj==='travaille'||tj==='AT')?diff(c1d,c1f)+diff(c2d,c2f)+diff(c3d,c3f):0);
    sc('date_derniere_modif',Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy HH:mm')+' (admin)');
    sc('modif_admin',Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy HH:mm'));
    const dataFresh = sheet.getDataRange().getValues();
const headersFresh = dataFresh[0].map(h => String(h).trim());
const collabIdJour = dataFresh[idx][headersFresh.indexOf('collab_id')];
const dateJourJour = dataFresh[idx][headersFresh.indexOf('date_jour')];
_recalculerSemaineCollab(collabIdJour, dateJourJour, sheet, dataFresh, headersFresh);
    return{ok:true};
  } finally {
    lock.releaseLock();
  }
}
function _getRecapForCollabs(periodeId,filtreEquipeId){
  const periodes=sheetToObjects(SHEETS.periodes());
  let periode=periodeId?periodes.find(p=>String(p.periode_id)===String(periodeId)):periodes.find(p=>p.statut==='ouverte')||periodes[periodes.length-1];
  if(!periode)return{ok:false,error:'Aucune periode trouvee'};
  const pId=String(periode.periode_id);
  let collabs=sheetToObjects(SHEETS.collabs()).filter(c=>c.actif===true||c.actif==='TRUE'||c.actif===1);
  if(filtreEquipeId)collabs=collabs.filter(c=>String(c.equipe_id)===filtreEquipeId);
  const jours=sheetToObjects(SHEETS.jours()).filter(j=>String(j.periode_id)===pId);
  const equipes=sheetToObjects(SHEETS.equipes());
  function joursOuvres(debut,fin){let n=0;const d=new Date(debut),f=new Date(fin);while(d<=f){const day=d.getDay();if(day!==0&&day!==6)n++;d.setDate(d.getDate()+1);}return n;}
  const nbO=joursOuvres(periode.date_debut,periode.date_fin);
  const recap=collabs.map(c=>{
    const mj=jours.filter(j=>String(j.collab_id)===String(c.collab_id));
    const sa=mj.filter(j=>j.type_jour&&j.type_jour!=='');
    const th=mj.reduce((s,j)=>s+(Number(j.total_heures)||0),0);
    const parseDate=d=>{if(!d)return 0;if(String(d).includes('T'))return new Date(d).getTime();const m=String(d).match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);return m?new Date(m[3],m[2]-1,m[1],m[4],m[5]).getTime():0;};
    const ds=sa.length>0?sa.sort((a,b)=>parseDate(b.date_derniere_modif)-parseDate(a.date_derniere_modif))[0].date_derniere_modif:'';
    const eq=equipes.find(e=>String(e.equipe_id)===String(c.equipe_id));
    return{collabId:c.collab_id,nomAffiche:c.nom_affiche||(c.prenom+' '+c.nom),structure:c.structure,typeContrat:c.type_contrat,heuresHebdo:c.heures_hebdo,equipeId:c.equipe_id||'',equipeNom:eq?eq.nom_equipe:'',totalHeures:th,nbSaisis:sa.length,nbNonSaisis:Math.max(0,nbO-sa.length),derniereSaisie:ds,alerte:sa.length<nbO};
  });
  return{ok:true,periode:{periodeId:periode.periode_id,nomPeriode:periode.nom_periode,dateDebut:fmtDate(periode.date_debut),dateFin:fmtDate(periode.date_fin),statut:periode.statut},recap};
}
function activerCollabsEnAttente(){
  const sheet=SHEETS.collabs(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
  const today=new Date();today.setHours(0,0,0,0);
  data.forEach((row,i)=>{if(i===0)return;if(String(row[headers.indexOf('statut')])!=='en_attente')return;const da=row[headers.indexOf('date_activation')];if(!da)return;const d=new Date(da);d.setHours(0,0,0,0);if(d<=today){sheet.getRange(i+1,headers.indexOf('statut')+1).setValue('actif');sheet.getRange(i+1,headers.indexOf('actif')+1).setValue(true);}});
}
function triggerJournalier() {
  activerCollabsEnAttente();
  cloturerPeriodesCloturees();
  ouvrirPeriodesPlanifiees();
  genererJourAujourdhui();
  genererPeriodesSuivantes();
  supprimerVieuxJours(65);
}
function genererPeriodesSuivantes() {
  const periodes = sheetToObjects(SHEETS.periodes());
  const sheet = SHEETS.periodes();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const today = new Date();
  const isNovembre15 = today.getDate() === 15 && today.getMonth() === 10;
  const nbCivil   = periodes.filter(p => p.statut === 'planifiee' && String(p.type_periode) === 'civil').length;
  const nbDecalee = periodes.filter(p => p.statut === 'planifiee' && String(p.type_periode) === 'decalee').length;
  const newRows = [];
  if (nbCivil < 2)   newRows.push(..._calculerPeriodesCiviles(periodes, 2 - nbCivil, headers));
  if (nbDecalee < 2) newRows.push(..._calculerPeriodesDecalees(periodes, 2 - nbDecalee, headers));
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow()+1, 1, newRows.length, newRows[0].length).setValues(newRows);
    Logger.log('Périodes générées : ' + newRows.length);
  }
  if (isNovembre15) {
    _genererAnneeComplete(periodes, headers, today.getFullYear() + 1);
  }
}
function _derniereDate(periodes, typePeriode, champ) {
  const filtered = periodes
    .filter(p => String(p.type_periode) === typePeriode &&
      (p.statut === 'planifiee' || p.statut === 'ouverte'))
    .map(p => {
      const v = p[champ];
      if (!v) return null;
      const s = String(v).trim();
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
        const [d, m, y] = s.split('/');
        return new Date(y, m - 1, d);
      }
      const d = new Date(v);
      return isNaN(d) ? null : d;
    })
    .filter(d => d !== null);
  if (!filtered.length) return null;
  return new Date(Math.max(...filtered));
}
function _calculerPeriodesCiviles(periodes, nb, headers) {
  const rows = [];
  let ref = _derniereDate(periodes, 'civil', 'date_fin');
  if (!ref) {
    const today = new Date();
    ref = new Date(today.getFullYear(), today.getMonth(), 0);
  }
  for (let i = 0; i < nb; i++) {
    const debut   = new Date(ref.getFullYear(), ref.getMonth() + 1, 1);
    const fin     = new Date(ref.getFullYear(), ref.getMonth() + 2, 0);
    const cloture = new Date(fin.getFullYear(), fin.getMonth(), fin.getDate() + 1);
    const mm = String(debut.getMonth() + 1).padStart(2, '0');
    const aa = debut.getFullYear();
    rows.push(headers.map(h => {
      switch(h) {
        case 'periode_id':   return 'CIV_' + aa + '_' + mm;
        case 'nom_periode':  return 'CIVIL_' + aa + '_' + mm;
        case 'type_periode': return 'civil';
        case 'date_debut':   return Utilities.formatDate(debut,   'Europe/Paris', 'yyyy-MM-dd');
        case 'date_fin':     return Utilities.formatDate(fin,     'Europe/Paris', 'yyyy-MM-dd');
        case 'date_cloture': return Utilities.formatDate(cloture, 'Europe/Paris', 'yyyy-MM-dd');
        case 'statut':       return 'planifiee';
        default:             return '';
      }
    }));
    ref = fin;
  }
  return rows;
}
function _calculerPeriodesDecalees(periodes, nb, headers) {
  const rows = [];
  const fmt = d => d.getFullYear() + '-' +
    String(d.getMonth()+1).padStart(2,'0') + '-' +
    String(d.getDate()).padStart(2,'0');
  let ref = _derniereDate(periodes, 'decalee', 'date_fin');
  if (!ref) {
    const today = new Date();
    ref = new Date(today.getFullYear(), today.getMonth(), 14);
  }
  for (let i = 0; i < nb; i++) {
    const moisRef = i === 0
      ? new Date(ref.getFullYear(), ref.getMonth() + 1, 1)
      : new Date(ref.getFullYear(), ref.getMonth(), 1);
    const annee = moisRef.getFullYear();
    const mois  = moisRef.getMonth();
    const debut   = new Date(annee, mois, 15);
    const fin     = new Date(annee, mois + 1, 14);
    const cloture = new Date(annee, mois + 1, 15);
    const mm = String(mois + 1).padStart(2, '0');
    rows.push(headers.map(h => {
      switch(h) {
        case 'periode_id':   return 'DEC_' + annee + '_' + mm;
        case 'nom_periode':  return 'DECAL_' + annee + '_' + mm;
        case 'type_periode': return 'decalee';
        case 'date_debut':   return fmt(debut);
        case 'date_fin':     return fmt(fin);
        case 'date_cloture': return fmt(cloture);
        case 'statut':       return 'planifiee';
        default:             return '';
      }
    }));
    ref = new Date(annee, mois + 1, 14);
  }
  return rows;
}
function _genererAnneeComplete(periodes, headers, annee) {
  const sheet = SHEETS.periodes();
  const dejaFait = periodes.some(p => {
    const d = new Date(p.date_debut);
    return !isNaN(d) && d.getFullYear() === annee;
  });
  if (dejaFait) return;
  const newRows = [];
  for (let mois = 0; mois < 12; mois++) {
    const debutC   = new Date(annee, mois, 1);
    const finC     = new Date(annee, mois + 1, 0);
    const clotureC = new Date(annee, mois + 1, 1);
    const mm = String(mois + 1).padStart(2, '0');
    newRows.push(headers.map(h => {
      switch(h) {
        case 'periode_id':   return 'CIV_' + annee + '_' + mm;
        case 'nom_periode':  return 'CIVIL_' + annee + '_' + mm;
        case 'type_periode': return 'civil';
        case 'date_debut':   return Utilities.formatDate(debutC,   'Europe/Paris', 'yyyy-MM-dd');
        case 'date_fin':     return Utilities.formatDate(finC,     'Europe/Paris', 'yyyy-MM-dd');
        case 'date_cloture': return Utilities.formatDate(clotureC, 'Europe/Paris', 'yyyy-MM-dd');
        case 'statut':       return 'planifiee';
        default:             return '';
      }
    }));
    const debutD   = new Date(annee, mois, 15);
    const finD     = new Date(annee, mois + 1, 14);
    const clotureD = new Date(annee, mois + 1, 15);
    newRows.push(headers.map(h => {
      switch(h) {
        case 'periode_id':   return 'DEC_' + annee + '_' + mm;
        case 'nom_periode':  return 'DECAL_' + annee + '_' + mm;
        case 'type_periode': return 'decalee';
        case 'date_debut':   return Utilities.formatDate(debutD,   'Europe/Paris', 'yyyy-MM-dd');
        case 'date_fin':     return Utilities.formatDate(finD,     'Europe/Paris', 'yyyy-MM-dd');
        case 'date_cloture': return Utilities.formatDate(clotureD, 'Europe/Paris', 'yyyy-MM-dd');
        case 'statut':       return 'planifiee';
        default:             return '';
      }
    }));
  }
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow()+1, 1, newRows.length, newRows[0].length).setValues(newRows);
    try {
      const admin = sheetToObjects(SHEETS.admins()).find(a => a.actif === true);
      if (admin && admin.email) {
        MailApp.sendEmail(
          admin.email,
          '🌱 Périodes ' + annee + ' générées — Chants de la Terre',
          'Les 24 périodes de paie ' + annee + ' ont été générées automatiquement.\n\nPensez à vérifier et ajuster les dates si nécessaire dans le backoffice.\n\nBonne journée !'
        );
      }
    } catch(e) { Logger.log('Email non envoyé : ' + e.message); }
  }
}
function genererJourAujourdhui() {
  const aujourd = new Date();
  aujourd.setHours(0,0,0,0);
  const aujourdStr = Utilities.formatDate(aujourd, 'Europe/Paris', 'yyyy-MM-dd');
  const JS = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
  const jourSemaine = JS[aujourd.getDay()];
  if (aujourd.getDay() === 0 || aujourd.getDay() === 6) return;
  const periodes = sheetToObjects(SHEETS.periodes());
  const periodesOuvertes = periodes.filter(p => p.statut === 'ouverte');
  if (!periodesOuvertes.length) return;
  const collabs = sheetToObjects(SHEETS.collabs()).filter(c =>
    c.actif === true || c.actif === 'TRUE' || c.actif === 1
  );
  const js = SHEETS.jours();
  const jd = js.getDataRange().getValues();
  const jh = jd[0].map(h => String(h).trim());
  const existants = new Set(jd.slice(1).map(r => String(r[jh.indexOf('jour_id')])));
  const newRows = [];
  periodesOuvertes.forEach(periode => {
    const typePeriode = String(periode.type_periode || '').trim();
    const collabsFiltres = collabs.filter(c =>
      String(c.type_periode || '').trim() === typePeriode
    );
    collabsFiltres.forEach(c => {
      const jourId = String(c.collab_id) + '_' + aujourdStr;
      if (!existants.has(jourId)) {
        newRows.push(jh.map(h => {
          switch(h) {
            case 'jour_id':      return jourId;
            case 'collab_id':    return c.collab_id;
            case 'date_jour':    return aujourdStr;
            case 'jour_semaine': return jourSemaine;
            case 'periode_id':   return periode.periode_id;
            case 'type_jour':    return 'travaillée';
            default:             return '';
          }
        }));
      }
    });
  });
  if (newRows.length > 0) {
    js.getRange(js.getLastRow()+1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
}
function supprimerVieuxJours(nbJours) {
  const limite = new Date();
  limite.setDate(limite.getDate() - nbJours);
  limite.setHours(0,0,0,0);
  const periodes = sheetToObjects(SHEETS.periodes());
  const periodesProtegees = new Set(
    periodes
      .filter(p => p.statut === 'ouverte' || p.statut === 'planifiee' || p.statut === 'cloturee' || p.statut === 'validee')
      .map(p => String(p.periode_id))
  );
  const sheet = SHEETS.jours();
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idxDate = headers.indexOf('date_jour');
  const idxPeriode = headers.indexOf('periode_id');
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const d = new Date(data[i][idxDate]);
    const periodeId = String(data[i][idxPeriode]);
    if (!isNaN(d) && d < limite && !periodesProtegees.has(periodeId)) {
      rowsToDelete.push(i + 1);
    }
  }
  rowsToDelete.forEach(r => sheet.deleteRow(r));
}
function ouvrirPeriodesPlanifiees() {
  const sheet=SHEETS.periodes(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
  const today=new Date();today.setHours(0,0,0,0);
  data.forEach((row,i)=>{
    if(i===0)return;
    if(String(row[headers.indexOf('statut')])!=='planifiee')return;
    const dd=new Date(row[headers.indexOf('date_debut')]);
    dd.setHours(0,0,0,0);
    if(dd<=today){sheet.getRange(i+1,headers.indexOf('statut')+1).setValue('ouverte');}
  });
}
function cloturerPeriodesCloturees(){
  const sheet=SHEETS.periodes(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
  const today=new Date();today.setHours(0,0,0,0);
  data.forEach((row,i)=>{
    if(i===0)return;
    if(String(row[headers.indexOf('statut')])!=='ouverte')return;
    const dc=new Date(row[headers.indexOf('date_cloture')]);
    dc.setHours(0,0,0,0);
    if(dc<=today){
      sheet.getRange(i+1,headers.indexOf('statut')+1).setValue('cloturee');
      const periodeId=String(row[headers.indexOf('periode_id')]);
      initRecapPaieAuto(periodeId);
    }
  });
}
function initRecapPaieAuto(periodeId){
  const collabs=sheetToObjects(SHEETS.collabs()).filter(c=>c.actif===true||c.actif==='TRUE'||c.actif===1);
  const jours=sheetToObjects(SHEETS.jours()).filter(j=>String(j.periode_id)===String(periodeId));
  const sheet=SS.getSheetByName('Recap_Paie');
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h=>String(h).trim());
  const allData=sheet.getDataRange().getValues();
  for(let i=allData.length-1;i>=1;i--){
    if(String(allData[i][headers.indexOf('periode_id')])===String(periodeId))sheet.deleteRow(i+1);
  }
  const newRows=collabs.map(c=>{
    const jc=jours.filter(j=>String(j.collab_id)===String(c.collab_id));
    const totalH=jc.reduce((s,j)=>s+(Number(j.total_heures)||0),0);
    const nbCp=jc.filter(j=>String(j.type_jour)==='CP').length;
    const nbAt=jc.filter(j=>String(j.type_jour)==='AT').length;
    const nbTrav=jc.filter(j=>(String(j.type_jour)==='travaillée'||String(j.type_jour)==='travaille')&&(Number(j.total_heures)||0)>0).length;
    return headers.map(h=>{
      switch(h){
        case 'periode_id':return periodeId;
        case 'collab_id':return c.collab_id;
        case 'nom_affiche':return c.nom_affiche||(c.prenom+' '+c.nom);
        case 'structure':return c.structure||'';
        case 'type_contrat':return c.type_contrat||'';
        case 'heures_hebdo':return c.heures_hebdo||'';
        case 'total_heures':return totalH;
        case 'nb_cp':return nbCp;
        case 'nb_at':return nbAt;
        case 'nb_jours_travailles':return nbTrav;
        case 'statut_validation':return 'en_attente';
        case 'date_validation':return '';
        case 'statut_signature':return 'en_attente';
        case 'note_admin':return '';
        default:return '';
      }
    });
  });
  if(newRows.length>0)sheet.getRange(sheet.getLastRow()+1,1,newRows.length,newRows[0].length).setValues(newRows);
}
function bo_savePeriode(token, params) {
  if(!checkAdmin(token)) return {ok:false, error:'Acces refuse'};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet=SHEETS.periodes(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
    if(params.periodeId){
      const idx=data.findIndex((row,i)=>i>0&&String(row[headers.indexOf('periode_id')])===String(params.periodeId));
      if(idx===-1)return{ok:false,error:'Periode introuvable'};
      const sv=(n,v)=>{const c=headers.indexOf(n);if(c>=0)sheet.getRange(idx+1,c+1).setValue(v);};
      sv('nom_periode',params.nomPeriode||'');sv('type_periode',params.typePeriode||'');
      sv('date_debut',params.dateDebut||'');sv('date_fin',params.dateFin||'');
      sv('date_cloture',params.dateCloture||'');sv('statut',params.statut||'planifiee');
      return{ok:true,action:'updated'};
    }else{
      if(!params.nomPeriode||!params.dateDebut||!params.dateFin)return{ok:false,error:'Nom, debut et fin requis'};
      const lastId=data.slice(1).reduce((max,row)=>{const id=String(row[headers.indexOf('periode_id')]||'').replace('PER','');return Math.max(max,parseInt(id)||0);},0);
      const newId='PER'+String(lastId+1).padStart(3,'0');
      sheet.appendRow(headers.map(h=>{switch(h){case 'periode_id':return newId;case 'nom_periode':return params.nomPeriode;case 'type_periode':return params.typePeriode||'';case 'date_debut':return params.dateDebut||'';case 'date_fin':return params.dateFin||'';case 'date_cloture':return params.dateCloture||'';case 'statut':return params.statut||'planifiee';default:return '';}}));
      return{ok:true,action:'created',periodeId:newId};
    }
  } finally {
    lock.releaseLock();
  }
}
function bo_deletePeriode(token, periodeId) {
  if(!checkAdmin(token)) return {ok:false, error:'Acces refuse'};
  const sheet=SHEETS.periodes(), data=sheet.getDataRange().getValues(), headers=data[0].map(h=>String(h).trim());
  const idx=data.findIndex((row,i)=>i>0&&String(row[headers.indexOf('periode_id')])===String(periodeId));
  if(idx===-1) return {ok:false, error:'Periode introuvable'};
  if(String(data[idx][headers.indexOf('statut')])!=='planifiee') return {ok:false, error:'Suppression impossible : statut non planifiée'};
  sheet.deleteRow(idx+1);
  return {ok:true};
}
function bo_initRecapPaie(token, periodeId) {
  if(!checkAdmin(token)) return {ok:false, error:'Acces refuse'};
  const periode = sheetToObjects(SHEETS.periodes()).find(p => String(p.periode_id) === String(periodeId));
  if(!periode) return {ok:false, error:'Periode introuvable'};
  const collabs = sheetToObjects(SHEETS.collabs()).filter(c => c.actif===true||c.actif==='TRUE'||c.actif===1);
  const jours = sheetToObjects(SHEETS.jours()).filter(j => String(j.periode_id)===String(periodeId));
  const sheet = SS.getSheetByName('Recap_Paie');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const allData = sheet.getDataRange().getValues();
  for(let i = allData.length - 1; i >= 1; i--) {
    if(String(allData[i][headers.indexOf('periode_id')]) === String(periodeId)) sheet.deleteRow(i+1);
  }
  const newRows = collabs.map(c => {
    const joursCollab = jours.filter(j => String(j.collab_id)===String(c.collab_id));
    const totalH = joursCollab.reduce((s,j) => s+(Number(j.total_heures)||0), 0);
    const nbCp = joursCollab.filter(j => String(j.type_jour)==='CP').length;
    const nbAt = joursCollab.filter(j => String(j.type_jour)==='AT').length;
    const nbTrav = joursCollab.filter(j => (String(j.type_jour)==='travaillée'||String(j.type_jour)==='travaille') && (Number(j.total_heures)||0)>0).length;
    return headers.map(h => {
      switch(h) {
        case 'periode_id': return periodeId;
        case 'collab_id': return c.collab_id;
        case 'nom_affiche': return c.nom_affiche||(c.prenom+' '+c.nom);
        case 'structure': return c.structure||'';
        case 'type_contrat': return c.type_contrat||'';
        case 'heures_hebdo': return c.heures_hebdo||'';
        case 'total_heures': return totalH;
        case 'nb_cp': return nbCp;
        case 'nb_at': return nbAt;
        case 'nb_jours_travailles': return nbTrav;
        case 'statut_validation': return 'en_attente';
        case 'date_validation': return '';
        case 'statut_signature': return 'en_attente';
        case 'note_admin': return '';
        default: return '';
      }
    });
  });
  Logger.log('nb collabs: ' + collabs.length + ' nb rows: ' + newRows.length);
  if(newRows.length > 0) sheet.getRange(sheet.getLastRow()+1, 1, newRows.length, newRows[0].length).setValues(newRows);
  return {ok:true, nb:newRows.length};
}
function bo_getRecapPaie(token, periodeId) {
  if(!checkAdmin(token)) return {ok:false, error:'Acces refuse'};
  const sheet = SS.getSheetByName('Recap_Paie');
  const data = sheetToObjects(sheet);
  const result = data.filter(r => String(r.periode_id)===String(periodeId));
  return {ok:true, recap:result};
}
function bo_validerCollab(token, periodeId, collabId, note) {
  if(!checkAdmin(token)) return {ok:false, error:'Acces refuse'};
  const jours=sheetToObjects(SHEETS.jours()).filter(j=>String(j.periode_id)===String(periodeId)&&String(j.collab_id)===String(collabId));
  const totalH=jours.reduce((s,j)=>s+(Number(j.total_heures)||0),0);
  const nbCp=jours.filter(j=>String(j.type_jour)==='CP').length;
  const nbAt=jours.filter(j=>String(j.type_jour)==='AT').length;
  const nbTrav=jours.filter(j=>(String(j.type_jour)==='travaillée'||String(j.type_jour)==='travaille')&&(Number(j.total_heures)||0)>0).length;
  const sheet=SS.getSheetByName('Recap_Paie');
  const data=sheet.getDataRange().getValues();
  const headers=data[0].map(h=>String(h).trim());
  const idx=data.findIndex((row,i)=>i>0&&String(row[headers.indexOf('periode_id')])===String(periodeId)&&String(row[headers.indexOf('collab_id')])===String(collabId));
  if(idx===-1) return {ok:false, error:'Ligne introuvable'};
  const sv=(n,v)=>{const c=headers.indexOf(n);if(c>=0)sheet.getRange(idx+1,c+1).setValue(v);};
  const cur=String(data[idx][headers.indexOf('statut_validation')]);
  const next=cur==='valide'?'en_attente':'valide';
  sv('statut_validation',next);
  sv('date_validation',next==='valide'?Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy HH:mm'):'');
  sv('total_heures',totalH);
  sv('nb_cp',nbCp);
  sv('nb_at',nbAt);
  sv('nb_jours_travailles',nbTrav);
  sv('note_admin',note||'');
  return {ok:true, newStatut:next, totalH, nbCp, nbAt};
}
function bo_saveNoteAdmin(token, periodeId, collabId, note) {
  if(!checkAdmin(token)) return {ok:false, error:'Acces refuse'};
  const sheet = SS.getSheetByName('Recap_Paie');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idx = data.findIndex((row,i) => i>0 && String(row[headers.indexOf('periode_id')])===String(periodeId) && String(row[headers.indexOf('collab_id')])===String(collabId));
  if(idx===-1) return {ok:false, error:'Ligne introuvable'};
  const nc = headers.indexOf('note_admin');
  sheet.getRange(idx+1, nc+1).setValue(note||'');
  return {ok:true};
}
function testTotalHebdo() {
  const result = saveHoraires('f69kj5pi', {
    jourId: 'COLL006_2026-05-13',
    typeJour: 'travaillée',
    c1Debut: '08:00',
    c1Fin: '12:00'
  });
  Logger.log(JSON.stringify(result));
}

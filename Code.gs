const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  collabs:  () => SS.getSheetByName('Collaborateurs'),
  periodes: () => SS.getSheetByName('Periodes'),
  jours:    () => SS.getSheetByName('Jours'),
  admins:   () => SS.getSheetByName('Admins'),
  equipes:  () => SS.getSheetByName('Equipes'),
};
function fmtDate(d) { if (!d) return ''; try { return Utilities.formatDate(new Date(d), 'Europe/Paris', 'dd/MM/yyyy'); } catch(e) { return String(d); } }
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
  return {ok:true, jour:{jourId:jour.jour_id,dateJour:fmtDate(jour.date_jour),jourSemaine:jour.jour_semaine||JS[today.getDay()],periodeId:jour.periode_id,typeJour:jour.type_jour,c1Debut:jour.c1_debut,c1Fin:jour.c1_fin,c2Debut:jour.c2_debut,c2Fin:jour.c2_fin,c3Debut:jour.c3_debut,c3Fin:jour.c3_fin,commentaire:jour.commentaire,totalHeures:jour.total_heures}, collab:{nomAffiche:collab.nom_affiche||(collab.prenom+' '+collab.nom),heuresHebdo:collab.heures_hebdo}};
}
function getMesJours(token, periodeId) {
  const collab = checkCollab(token); if (!collab) return {ok:false, error:'Token invalide'};
  return {ok:true, jours:sheetToObjects(SHEETS.jours()).filter(j=>String(j.collab_id)===String(collab.collab_id)&&(!periodeId||String(j.periode_id)===String(periodeId))).map(j=>({jourId:j.jour_id,dateJour:fmtDate(j.date_jour),jourSemaine:j.jour_semaine,typeJour:j.type_jour,c1Debut:j.c1_debut,c1Fin:j.c1_fin,c2Debut:j.c2_debut,c2Fin:j.c2_fin,c3Debut:j.c3_debut,c3Fin:j.c3_fin,commentaire:j.commentaire,totalHeures:j.total_heures}))};
}
function getJourById(token, jourId) {
  const collab = checkCollab(token); if (!collab) return {ok:false, error:'Token invalide'};
  const jour = sheetToObjects(SHEETS.jours()).find(j=>String(j.jour_id)===String(jourId)&&String(j.collab_id)===String(collab.collab_id));
  if (!jour) return {ok:false, error:'Jour introuvable'};
  return {ok:true, jour:{jourId:jour.jour_id,dateJour:fmtDate(jour.date_jour),jourSemaine:jour.jour_semaine,typeJour:jour.type_jour,c1Debut:jour.c1_debut,c1Fin:jour.c1_fin,c2Debut:jour.c2_debut,c2Fin:jour.c2_fin,c3Debut:jour.c3_debut,c3Fin:jour.c3_fin,commentaire:jour.commentaire,totalHeures:jour.total_heures}};
}
function saveHoraires(token, params) {
  const collab = checkCollab(token); if (!collab) return {ok:false, error:'Token invalide'};
  const sheet=SHEETS.jours(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
  const idx=data.findIndex((row,i)=>i>0&&String(row[headers.indexOf('jour_id')])===String(params.jourId)&&String(row[headers.indexOf('collab_id')])===String(collab.collab_id));
  if(idx===-1)return{ok:false,error:'Jour introuvable'};
  function diff(d,f){if(!d||!f)return 0;const[dh,dm]=String(d).split(':').map(Number);const[fh,fm]=String(f).split(':').map(Number);const m=(fh*60+fm)-(dh*60+dm);return m>0?m/60:0;}
  const tj=params.typeJour||'',c1d=params.c1Debut||'',c1f=params.c1Fin||'',c2d=params.c2Debut||'',c2f=params.c2Fin||'',c3d=params.c3Debut||'',c3f=params.c3Fin||'';
  const total=tj==='travail'?diff(c1d,c1f)+diff(c2d,c2f)+diff(c3d,c3f):0;
  const sc=(n,v)=>{const c=headers.indexOf(n);if(c>=0)sheet.getRange(idx+1,c+1).setValue(v);};
  sc('type_jour',tj);sc('c1_debut',c1d);sc('c1_fin',c1f);sc('c2_debut',c2d);sc('c2_fin',c2f);sc('c3_debut',c3d);sc('c3_fin',c3f);sc('commentaire',params.commentaire||'');sc('total_heures',total);sc('date_derniere_modif',Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy HH:mm'));
  const nm=headers.indexOf('nb_modifications');if(nm>=0){const cur=Number(sheet.getRange(idx+1,nm+1).getValue())||0;sheet.getRange(idx+1,nm+1).setValue(cur+1);}
  return{ok:true};
}
function bo_getCollabs(token) {
  if(!checkAdmin(token))return{ok:false,error:'Acces refuse'};
  const equipes=sheetToObjects(SHEETS.equipes());
  return{ok:true,collabs:sheetToObjects(SHEETS.collabs()).map(c=>{const eq=equipes.find(e=>String(e.equipe_id)===String(c.equipe_id));return{collabId:c.collab_id,prenom:c.prenom,nom:c.nom,nomAffiche:c.nom_affiche||(c.prenom+' '+c.nom),email:c.email,structure:c.structure,typeContrat:c.type_contrat,heuresHebdo:c.heures_hebdo,statut:c.statut,dateActivation:fmtDate(c.date_activation),actif:c.actif,token:c.token,equipeId:c.equipe_id||'',equipeNom:eq?eq.nom_equipe:'',lien:c.token?'https://chantsdelaterre.github.io/suivi-temps/?id='+c.token:''};})};
}
function bo_saveCollab(token, params) {
  if(!checkAdmin(token))return{ok:false,error:'Acces refuse'};
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
  return{ok:true,periodes:sheetToObjects(SHEETS.periodes()).map(p=>({periodeId:p.periode_id,nomPeriode:p.nom_periode,typeContrat:p.type_contrat,dateDebut:fmtDate(p.date_debut),dateFin:fmtDate(p.date_fin),dateCloture:fmtDate(p.date_cloture),statut:p.statut}))};
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
  const jours=sheetToObjects(SHEETS.jours()).filter(j=>String(j.collab_id)===String(collabId)&&(!periodeId||String(j.periode_id)===String(periodeId))).sort((a,b)=>new Date(a.date_jour)-new Date(b.date_jour)).map(j=>({jourId:j.jour_id,dateJour:fmtDate(j.date_jour),jourSemaine:j.jour_semaine,typeJour:j.type_jour,c1Debut:j.c1_debut,c1Fin:j.c1_fin,c2Debut:j.c2_debut,c2Fin:j.c2_fin,c3Debut:j.c3_debut,c3Fin:j.c3_fin,commentaire:j.commentaire,totalHeures:j.total_heures,remarqueManager:j.remarque_manager||'',dateDerniereModif:j.date_derniere_modif||'',nbModifications:j.nb_modifications||0}));
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
function bo_saveJour(token,params){
  if(!checkAdmin(token))return{ok:false,error:'Acces refuse - admin uniquement'};
  const sheet=SHEETS.jours(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
  const idx=data.findIndex((row,i)=>i>0&&String(row[headers.indexOf('jour_id')])===String(params.jourId));
  if(idx===-1)return{ok:false,error:'Jour introuvable'};
  const sc=(n,v)=>{const c=headers.indexOf(n);if(c>=0)sheet.getRange(idx+1,c+1).setValue(v);};
  function diff(d,f){if(!d||!f)return 0;const[dh,dm]=String(d).split(':').map(Number);const[fh,fm]=String(f).split(':').map(Number);const m=(fh*60+fm)-(dh*60+dm);return m>0?m/60:0;}
  const tj=params.typeJour||'',c1d=params.c1Debut||'',c1f=params.c1Fin||'',c2d=params.c2Debut||'',c2f=params.c2Fin||'',c3d=params.c3Debut||'',c3f=params.c3Fin||'';
  sc('type_jour',tj);sc('c1_debut',c1d);sc('c1_fin',c1f);sc('c2_debut',c2d);sc('c2_fin',c2f);sc('c3_debut',c3d);sc('c3_fin',c3f);
  sc('commentaire',params.commentaire||'');sc('total_heures',tj==='travail'?diff(c1d,c1f)+diff(c2d,c2f)+diff(c3d,c3f):0);
  sc('date_derniere_modif',Utilities.formatDate(new Date(),'Europe/Paris','dd/MM/yyyy HH:mm')+' (admin)');
  return{ok:true};
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
    const ds=sa.length>0?sa.sort((a,b)=>new Date(b.date_derniere_modif)-new Date(a.date_derniere_modif))[0].date_derniere_modif:'';
    const eq=equipes.find(e=>String(e.equipe_id)===String(c.equipe_id));
    return{collabId:c.collab_id,nomAffiche:c.nom_affiche||(c.prenom+' '+c.nom),structure:c.structure,typeContrat:c.type_contrat,heuresHebdo:c.heures_hebdo,equipeId:c.equipe_id||'',equipeNom:eq?eq.nom_equipe:'',totalHeures:th,nbSaisis:sa.length,nbNonSaisis:Math.max(0,nbO-sa.length),derniereSaisie:ds,alerte:sa.length<nbO};
  });
  return{ok:true,periode:{periodeId:periode.periode_id,nomPeriode:periode.nom_periode,dateDebut:fmtDate(periode.date_debut),dateFin:fmtDate(periode.date_fin),statut:periode.statut},recap};
}
function triggerJournalier(){activerCollabsEnAttente();ouvrirPeriodesPlanifiees();cloturerPeriodesCloturees();}
function activerCollabsEnAttente(){
  const sheet=SHEETS.collabs(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
  const today=new Date();today.setHours(0,0,0,0);
  data.forEach((row,i)=>{if(i===0)return;if(String(row[headers.indexOf('statut')])!=='en_attente')return;const da=row[headers.indexOf('date_activation')];if(!da)return;const d=new Date(da);d.setHours(0,0,0,0);if(d<=today){sheet.getRange(i+1,headers.indexOf('statut')+1).setValue('actif');sheet.getRange(i+1,headers.indexOf('actif')+1).setValue(true);}});
}
function ouvrirPeriodesPlanifiees(){
  const sheet=SHEETS.periodes(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
  const today=new Date();today.setHours(0,0,0,0);
  data.forEach((row,i)=>{if(i===0)return;if(String(row[headers.indexOf('statut')])!=='planifiee')return;const dd=new Date(row[headers.indexOf('date_debut')]);dd.setHours(0,0,0,0);if(dd<=today){sheet.getRange(i+1,headers.indexOf('statut')+1).setValue('ouverte');_genererJoursPeriode(row,headers);}});
}
function cloturerPeriodesCloturees(){
  const sheet=SHEETS.periodes(),data=sheet.getDataRange().getValues(),headers=data[0].map(h=>String(h).trim());
  const today=new Date();today.setHours(0,0,0,0);
  data.forEach((row,i)=>{if(i===0)return;if(String(row[headers.indexOf('statut')])!=='ouverte')return;const dc=new Date(row[headers.indexOf('date_cloture')]);dc.setHours(0,0,0,0);if(dc<=today)sheet.getRange(i+1,headers.indexOf('statut')+1).setValue('cloturee');});
}
function _genererJoursPeriode(periodeRow,periodeHeaders){
  const periodeId=String(periodeRow[periodeHeaders.indexOf('periode_id')]);
  const typeContrat=String(periodeRow[periodeHeaders.indexOf('type_contrat')]);
  const dateDebut=new Date(periodeRow[periodeHeaders.indexOf('date_debut')]);
  const dateFin=new Date(periodeRow[periodeHeaders.indexOf('date_fin')]);
  const js=SHEETS.jours(),jd=js.getDataRange().getValues(),jh=jd[0].map(h=>String(h).trim());
  const collabs=sheetToObjects(SHEETS.collabs()).filter(c=>{if(c.actif!==true&&c.actif!=='TRUE'&&c.actif!==1)return false;if(typeContrat&&typeContrat!=='tous'&&c.type_contrat!==typeContrat)return false;return true;});
  const JS=['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
  let lastId=jd.slice(1).reduce((max,row)=>{const id=String(row[jh.indexOf('jour_id')]||'').replace('J','');return Math.max(max,parseInt(id)||0);},0);
  const newRows=[];
  collabs.forEach(c=>{const d=new Date(dateDebut);while(d<=dateFin){const day=d.getDay();if(day!==0&&day!==6){lastId++;newRows.push(jh.map(h=>{switch(h){case 'jour_id':return 'J'+String(lastId).padStart(6,'0');case 'collab_id':return c.collab_id;case 'date_jour':return Utilities.formatDate(d,'Europe/Paris','yyyy-MM-dd');case 'jour_semaine':return JS[day];case 'periode_id':return periodeId;default:return '';}}));}d.setDate(d.getDate()+1);}});
  if(newRows.length>0)js.getRange(js.getLastRow()+1,1,newRows.length,newRows[0].length).setValues(newRows);
}

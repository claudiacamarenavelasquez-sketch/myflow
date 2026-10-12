// MY FLOW: private Google Apps Script web application.
// Deploy: Execute as me; Who has access: Only myself. Never use "Anyone".
const SHEET_ID='18yT5IQiIgddjdoQUz4A0eKQ8-6py0sW8BOXG35uKyug';
const HEADERS={Registros:['id','kind','desc','at','amount','done'],Digestion:['id','at','bloat','bristol','complete','note'],Perfil:['campo','valor','unidad','fecha','nota'],Alimentos:['alimento','grupo','observacion']};
function setupMyFlow(){
 const ss=SpreadsheetApp.openById(SHEET_ID);
 for(const [name,cols] of Object.entries(HEADERS)){
  let s=ss.getSheetByName(name);if(!s)s=ss.insertSheet(name);
  if(s.getLastRow()===0){s.appendRow(cols);s.setFrozenRows(1);s.getRange(1,1,1,cols.length).setBackground('#f9dce8').setFontWeight('bold')}
 }
 const profile=ss.getSheetByName('Perfil');
 if(profile.getLastRow()===1) [
 ['Estatura',1.55,'m','','Confirmado'],['Peso inicial',48,'kg','','Histórico'],
 ['Peso intermedio',50,'kg','','Pesaje no comparable'],['Peso reciente',52,'kg','','Pesaje provisional'],
 ['Busto',85,'cm','2026-10-10','Confección'],['Bajo busto',76,'cm','2026-10-10','Confección'],
 ['Cintura',72,'cm','2026-10-10','Confección'],['Cadera',84,'cm','2026-10-10','Confección']
 ].forEach(r=>profile.appendRow(r));
 const foods=ss.getSheetByName('Alimentos');
 if(foods.getLastRow()===1) ['Piña','Manzana','Uvas','Blueberries','Fresas','Plátano','Kiwi','Naranja','Papaya','Avena','Yogur griego','Chía','Linaza','Salmas','Quinua'].forEach(x=>foods.appendRow([x,'Habitual o en evaluación','Sin cantidades registradas']));
 return 'Tablas preparadas';
}
function doGet(){
 return HtmlService.createHtmlOutputFromFile('App').setTitle('MY FLOW');
}
function listEntries(){
 const ss=SpreadsheetApp.openById(SHEET_ID),items=[];
 for(const name of ['Registros','Digestion']){
  const sh=ss.getSheetByName(name);if(!sh||sh.getLastRow()<2)continue;
  const cols=HEADERS[name];
  sh.getRange(2,1,sh.getLastRow()-1,cols.length).getDisplayValues().forEach(row=>{
   if(!row[0])return;
   const record={};cols.forEach((k,i)=>record[k]=row[i]);
   record.id=Number(record.id)||record.id;
   record.amount=Number(record.amount)||0;
   record.done=record.done==='TRUE'||record.done==='true';
   if(name==='Digestion')record.kind='Digestión';
   items.push(record);
  });
 }
 return items;
}
function saveEntry(item){
 if(!item||!item.id||!['Comida','Entrenamiento','Gasto','Actividad','Digestión'].includes(item.kind))throw Error('Tipo de registro inválido');
 if(JSON.stringify(item).length>5000)throw Error('Registro demasiado largo');
 const name=item.kind==='Digestión'?'Digestion':'Registros',cols=HEADERS[name],sh=SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const last=sh.getLastRow(),ids=last>1?sh.getRange(2,1,last-1,1).getDisplayValues().flat():[];
  const pos=ids.indexOf(String(item.id)),values=cols.map(k=>item[k]??'');
  if(pos>=0)sh.getRange(pos+2,1,1,cols.length).setValues([values]);
  else sh.appendRow(values);
 }finally{lock.releaseLock()}
 return {ok:true};
}

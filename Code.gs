// MY FLOW: Google Sheets backend. Do not put private keys in GitHub.
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
 return 'Tablas creadas';
}
function check(e){
 const key=PropertiesService.getScriptProperties().getProperty('MYFLOW_KEY');
 return key&&key.length>=20&&e&&e.parameter&&e.parameter.key===key;
}
function output(obj,cb){
 const s=JSON.stringify(obj);
 if(cb&&/^[A-Za-z_$][\w$]{0,79}$/.test(cb))return ContentService.createTextOutput(cb+'('+s+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
 return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e){
 const p=e&&e.parameter||{};
 if(!check(e))return output({ok:false,error:'Sin autorización'},p.callback);
 try{
 const ss=SpreadsheetApp.openById(SHEET_ID),items=[];
 for(const name of ['Registros','Digestion']){
 const sh=ss.getSheetByName(name);if(!sh||sh.getLastRow()<2)continue;
 const keys=HEADERS[name];sh.getRange(2,1,sh.getLastRow()-1,keys.length).getValues().forEach(row=>{
 if(!row[0])return;
 const o={};keys.forEach((k,i)=>o[k]=row[i]);
 if(name==='Digestion')o.kind='Digestión';items.push(o);
 });
 }
 return output({ok:true,items},p.callback);
 }catch(err){return output({ok:false,error:String(err)},p.callback)}
}
function doPost(e){
 if(!check(e))return output({ok:false,error:'Sin autorización'});
 try{
 const item=JSON.parse(e.parameter.payload||'null');
 if(!item||!item.id||!item.kind||JSON.stringify(item).length>5000)throw Error('Registro inválido');
 const name=item.kind==='Digestión'?'Digestion':'Registros',cols=HEADERS[name],sh=SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
 const last=sh.getLastRow(),ids=last>1?sh.getRange(2,1,last-1,1).getDisplayValues().flat():[],pos=ids.indexOf(String(item.id)),values=cols.map(k=>item[k]??'');
 if(pos>=0)sh.getRange(pos+2,1,1,cols.length).setValues([values]);else sh.appendRow(values);
 }finally{lock.releaseLock()}
 return output({ok:true,id:item.id});
 }catch(err){return output({ok:false,error:String(err)})}
}
// En Google Sheets: Extensiones > Apps Script; pega este archivo.
// Ejecuta setupMyFlow y autoriza. En Propiedades del script agrega MYFLOW_KEY
// con al menos 20 caracteres aleatorios. Implementa como Aplicación web:
// Ejecutar como tú. Si eliges acceso público, comprende que la clave escrita
// en el navegador es una protección limitada; no es adecuada para datos de salud sensibles.
// La web actual necesita un paso adicional para leer y escribir en este backend.

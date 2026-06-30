/***********************************************************************
 * ÁREA DE MEMBROS — Código da Reconquista Magnética
 * Backend de contas (Google Apps Script + Google Sheets) — "Rota A".
 *
 * Faz: registro (e-mail + senha), login, e salva/carrega os dados que a
 * aluna inserir (diário etc.) — tudo no servidor, por e-mail, NÃO no cache.
 *
 * A senha NUNCA é guardada em texto puro: salvamos sal + hash SHA-256.
 *
 * É STANDALONE: cria/abre a própria planilha no seu Drive na 1ª execução.
 *
 * >>> DEPOIS DE COLAR:
 *   1) Rode a função autorizar() uma vez (consentir acesso ao Drive/Sheets).
 *   2) Implantar > Nova implantação > App da Web
 *      ("Executar como: Eu", "Quem pode acessar: Qualquer pessoa").
 *   3) Copie a URL /exec e cole na constante GAS_AUTH do index.html.
 ***********************************************************************/

var DB_NAME = 'Reconquista — Membros DB';
var SHEET   = 'usuarios';
var HEADERS = ['email','nome','salt','hash','token','criado_em','atualizado_em','dados'];

/* (opcional) Só deixar criar conta quem comprou: cole aqui os e-mails de
   compradores (minúsculos), OU deixe vazio para liberar qualquer e-mail. */
var COMPRADORES = []; // ex.: ['cliente1@gmail.com','cliente2@hotmail.com']

/* ====================== PLANILHA ====================== */
function getSS(){
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id){ try { return SpreadsheetApp.openById(id); } catch(e){} }
  var ss = SpreadsheetApp.create(DB_NAME);
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}
function sheet(){
  var ss = getSS();
  var sh = ss.getSheetByName(SHEET);
  if (!sh){ sh = ss.insertSheet(SHEET); }
  if (sh.getLastRow() === 0){
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ====================== ROTEADOR ====================== */
function doGet(e){
  return json({ ok:true, service:'membros', hint:'use POST' });
}
function doPost(e){
  try{
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (b.action){
      case 'register': return json(register(b));
      case 'login':    return json(login(b));
      case 'save':     return json(saveData(b));
      case 'load':     return json(loadData(b));
      default:         return json({ ok:false, error:'Ação inválida.' });
    }
  }catch(err){
    return json({ ok:false, error:String(err) });
  }
}

/* ====================== AÇÕES ====================== */
function register(b){
  var email = norm(b.email), nome = String(b.nome||'').trim(), senha = String(b.senha||'');
  if (!emailOk(email))  return { ok:false, error:'E-mail inválido.' };
  if (nome.length < 1)  return { ok:false, error:'Informe seu nome.' };
  if (senha.length < 6) return { ok:false, error:'A senha precisa ter ao menos 6 caracteres.' };
  if (COMPRADORES.length && COMPRADORES.indexOf(email) < 0)
    return { ok:false, error:'Esse e-mail não consta como comprador. Use o e-mail da compra.' };

  var lock = LockService.getScriptLock(); lock.waitLock(8000);
  try{
    if (findRow(email).row > 0) return { ok:false, error:'Esse e-mail já tem conta. Faça login.' };
    var salt  = Utilities.getUuid();
    var token = Utilities.getUuid();
    var dados = '{"diario":[]}';
    var now   = Date.now();
    sheet().appendRow([ email, nome, salt, sha256(salt+'|'+senha), token, now, now, dados ]);
    return { ok:true, nome:nome, token:token, dados:JSON.parse(dados) };
  } finally { lock.releaseLock(); }
}

function login(b){
  var email = norm(b.email), senha = String(b.senha||'');
  if (!emailOk(email)) return { ok:false, error:'E-mail inválido.' };
  var f = findRow(email);
  if (f.row < 0) return { ok:false, error:'E-mail não encontrado. Crie sua conta.' };
  if (f.data.hash !== sha256(f.data.salt+'|'+senha)) return { ok:false, error:'Senha incorreta.' };
  // rotaciona o token a cada login
  var token = Utilities.getUuid();
  sheet().getRange(f.row, col('token')).setValue(token);
  return { ok:true, nome:f.data.nome, token:token, dados:parse(f.data.dados) };
}

function saveData(b){
  var email = norm(b.email);
  var f = findRow(email);
  if (f.row < 0) return { ok:false, error:'Conta não encontrada.' };
  if (String(b.token||'') !== String(f.data.token)) return { ok:false, error:'Sessão inválida. Entre de novo.' };
  var sh = sheet();
  sh.getRange(f.row, col('dados')).setValue(JSON.stringify(b.dados||{}));
  sh.getRange(f.row, col('atualizado_em')).setValue(Date.now());
  return { ok:true };
}

function loadData(b){
  var email = norm(b.email);
  var f = findRow(email);
  if (f.row < 0) return { ok:false, error:'Conta não encontrada.' };
  if (String(b.token||'') !== String(f.data.token)) return { ok:false, error:'Sessão inválida. Entre de novo.' };
  return { ok:true, nome:f.data.nome, dados:parse(f.data.dados) };
}

/* ====================== HELPERS ====================== */
function findRow(email){
  var sh = sheet();
  var vals = sh.getDataRange().getValues();
  for (var i=1;i<vals.length;i++){
    if (String(vals[i][0]).toLowerCase() === email){
      var o = {};
      for (var c=0;c<HEADERS.length;c++){ o[HEADERS[c]] = vals[i][c]; }
      return { row:i+1, data:o };
    }
  }
  return { row:-1, data:null };
}
function col(name){ return HEADERS.indexOf(name) + 1; }
function norm(s){ return String(s||'').trim().toLowerCase(); }
function emailOk(m){ return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m); }
function parse(s){ try{ return JSON.parse(s||'{}'); }catch(e){ return {}; } }
function sha256(s){
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return bytes.map(function(b){ return ('0'+(b & 0xFF).toString(16)).slice(-2); }).join('');
}
function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ====================== SETUP / TESTE ====================== */
function autorizar(){
  sheet(); // cria a planilha e pede acesso ao Drive/Sheets
  Logger.log('OK — planilha: ' + getSS().getUrl());
}
function testeRapido(){
  Logger.log(JSON.stringify(register({action:'register', nome:'Teste', email:'teste@x.com', senha:'123456'})));
  Logger.log(JSON.stringify(login({action:'login', email:'teste@x.com', senha:'123456'})));
}

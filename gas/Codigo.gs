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
var HEADERS = ['email','nome','salt','hash','token','criado_em','atualizado_em','dados','reset_hash','reset_exp'];

// URL da área de membros (usada no link de redefinição de senha).
var MEMBROS_URL = 'https://area-membros-reconquista.vercel.app';
// Quanto tempo o link de redefinição fica válido (1 hora).
var RESET_TTL_MS = 60 * 60 * 1000;

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
  } else if (sh.getLastColumn() < HEADERS.length){
    // migração: acrescenta as colunas novas (reset_hash/reset_exp) ao cabeçalho
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
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
      case 'register':      return json(register(b));
      case 'login':         return json(login(b));
      case 'save':          return json(saveData(b));
      case 'load':          return json(loadData(b));
      case 'reset_request': return json(resetRequest(b));
      case 'reset_confirm': return json(resetConfirm(b));
      default:              return json({ ok:false, error:'Ação inválida.' });
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

/* ---- Esqueci minha senha: pedir link ---- */
function resetRequest(b){
  var email = norm(b.email);
  if (!emailOk(email)) return { ok:false, error:'E-mail inválido.' };
  // Sempre respondemos ok (não revelamos se o e-mail existe). Só enviamos se existir.
  var f = findRow(email);
  if (f.row > 0){
    var lock = LockService.getScriptLock(); lock.waitLock(8000);
    try{
      var code = Utilities.getUuid().replace(/-/g,'');   // token cru (vai no link)
      var sh = sheet();
      sh.getRange(f.row, col('reset_hash')).setValue(sha256(code));  // guardamos só o hash
      sh.getRange(f.row, col('reset_exp')).setValue(Date.now() + RESET_TTL_MS);
    } finally { lock.releaseLock(); }
    try { sendResetEmail(email, f.data.nome, code); } catch(err){}
  }
  return { ok:true };
}

/* ---- Esqueci minha senha: definir nova senha (a partir do link) ---- */
function resetConfirm(b){
  var email = norm(b.email), code = String(b.code||''), senha = String(b.senha||'');
  if (senha.length < 6) return { ok:false, error:'A senha precisa ter ao menos 6 caracteres.' };
  var f = findRow(email);
  if (f.row < 0) return { ok:false, error:'Link inválido.' };
  if (!f.data.reset_hash || sha256(code) !== String(f.data.reset_hash))
    return { ok:false, error:'Link inválido ou já utilizado. Peça um novo.' };
  if (!f.data.reset_exp || Date.now() > Number(f.data.reset_exp))
    return { ok:false, error:'O link expirou. Peça um novo.' };

  var lock = LockService.getScriptLock(); lock.waitLock(8000);
  try{
    var salt  = Utilities.getUuid();
    var token = Utilities.getUuid();
    var sh = sheet();
    sh.getRange(f.row, col('salt')).setValue(salt);
    sh.getRange(f.row, col('hash')).setValue(sha256(salt+'|'+senha));
    sh.getRange(f.row, col('token')).setValue(token);         // desloga sessões antigas
    sh.getRange(f.row, col('reset_hash')).setValue('');       // consome o link
    sh.getRange(f.row, col('reset_exp')).setValue('');
    sh.getRange(f.row, col('atualizado_em')).setValue(Date.now());
    return { ok:true, nome:f.data.nome, token:token, dados:parse(f.data.dados) };
  } finally { lock.releaseLock(); }
}

function sendResetEmail(email, nome, code){
  var primeiro = String(nome||'aluna').split(' ')[0];
  var link = MEMBROS_URL + '?reset=1&email=' + encodeURIComponent(email) + '&code=' + code;
  var subject = '🔑 Redefinir sua senha — Código da Reconquista Magnética';
  var html = '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f7f1f4;font-family:Arial,sans-serif">'
    + '<div style="max-width:520px;margin:30px auto;background:#fff;border-radius:20px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08)">'
    + '<div style="background:linear-gradient(135deg,#ec3a8b,#a855f7);padding:36px 28px;text-align:center;color:#fff">'
    + '<div style="font-size:42px;margin-bottom:8px">🔑</div>'
    + '<h1 style="margin:0;font-size:22px">Oi, ' + primeiro + '</h1>'
    + '<p style="margin:8px 0 0;font-size:14px;opacity:.9">Recebemos um pedido para redefinir sua senha</p>'
    + '</div>'
    + '<div style="padding:28px">'
    + '<p style="font-size:15px;color:#333;line-height:1.6">Clique no botão abaixo para criar uma nova senha. Este link vale por <strong>1 hora</strong>.</p>'
    + '<div style="text-align:center;margin:24px 0">'
    + '<a href="' + link + '" style="display:inline-block;background:linear-gradient(135deg,#ec3a8b,#d6246e);color:#fff;text-decoration:none;padding:14px 36px;border-radius:12px;font-size:16px;font-weight:700">Criar nova senha →</a>'
    + '</div>'
    + '<p style="font-size:13px;color:#888;line-height:1.5">Se você não pediu isso, pode ignorar este e-mail — sua senha atual continua valendo.</p>'
    + '<hr style="border:none;border-top:1px solid #eee;margin:20px 0">'
    + '<p style="font-size:12px;color:#aaa;text-align:center">Código da Reconquista Magnética<br>Qualquer dúvida, responda este e-mail.</p>'
    + '</div></div></body></html>';
  MailApp.sendEmail(email, subject, 'Redefinir sua senha: ' + link, { htmlBody: html, name: 'Código da Reconquista Magnética' });
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
  MailApp.getRemainingDailyQuota(); // pede o escopo de envio de e-mail (redefinir senha)
  Logger.log('OK — planilha: ' + getSS().getUrl());
}
function testeRapido(){
  Logger.log(JSON.stringify(register({action:'register', nome:'Teste', email:'teste@x.com', senha:'123456'})));
  Logger.log(JSON.stringify(login({action:'login', email:'teste@x.com', senha:'123456'})));
}

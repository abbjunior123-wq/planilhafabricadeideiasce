// ============================================================
//  Google Apps Script — Leads + Follow-up → Umbler Talk
//  Fábrica de Ideias
//
//  SETUP INICIAL (fazer uma vez):
//  1. Cole este código no Apps Script
//  2. Crie uma planilha Google Sheets e cole o ID abaixo em SHEET_ID
//     (o ID fica na URL: docs.google.com/spreadsheets/d/ID_AQUI/edit)
//  3. Rode a função criarTrigger() UMA VEZ manualmente:
//     Menu → Executar → criarTrigger
//  4. Implante como App da Web normalmente
// ============================================================

const UMBLER_TOKEN   = "APIGRAFICA-2026-04-17-2094-05-05--FD6A29AD4AC8F4EF7354B64F0E86854FCE8549316D29C6D92C7927239E590DB8";
const ORG_ID         = "aBTCVd0NssbZ1aT7";
const FROM_PHONE     = "+5585988461515";
const API_BASE       = "https://app-utalk.umbler.com/api/v1";
const SHEET_ID       = "COLE_O_ID_DA_PLANILHA_AQUI";

// ============================================================
//  RECEBE O LEAD DO FORMULÁRIO
// ============================================================
function doGet(e) {
  try {
    const data = {
      nome:     e.parameter.nome     || "",
      empresa:  e.parameter.empresa  || e.parameter.cnpj || "",
      whatsapp: e.parameter.whatsapp || "",
      servico:  e.parameter.servico  || "",
      mensagem: e.parameter.mensagem || "",
      tipo:     e.parameter.tipo     || "",
      email:    e.parameter.email    || "",
      source:   e.parameter.source   || "Site"
    };
    const chatId = enviarParaUmbler(data);
    if (chatId) salvarLead(data, chatId);
  } catch (err) {
    Logger.log("Erro doGet: " + err);
  }

  return ContentService
    .createTextOutput("ok")
    .setMimeType(ContentService.MimeType.TEXT);
}

// ============================================================
//  ENVIA MENSAGEM + NOTA INTERNA → retorna chatId
// ============================================================
function enviarParaUmbler(data) {
  const digitos = data.whatsapp.replace(/\D/g, "");
  const toPhone = digitos.startsWith("55") && digitos.length >= 12
    ? "+" + digitos
    : "+55" + digitos;

  const headers = { "Authorization": "Bearer " + UMBLER_TOKEN };

  // 1. Mensagem de abordagem ao lead
  let msgLead  = "Olá, *" + data.nome + "*! 😊\n\n";
  msgLead     += "Recebemos sua solicitação para entrar em contato com a *Gráfica Fábrica de Ideias*.\n\n";
  msgLead     += "O que você procura? 🎨";

  const r1 = UrlFetchApp.fetch(API_BASE + "/messages/simplified/", {
    method: "post", contentType: "application/json", headers: headers,
    payload: JSON.stringify({ organizationId: ORG_ID, toPhone: toPhone, fromPhone: FROM_PHONE, message: msgLead }),
    muteHttpExceptions: true
  });

  Logger.log("Mensagem lead — status: " + r1.getResponseCode());
  const resp1 = JSON.parse(r1.getContentText());
  const chatId = resp1.chat ? resp1.chat.id : null;
  if (!chatId) return null;

  // 2. Nota interna com dados completos
  let nota  = "📋 *Dados do lead (" + data.source + ")*\n";
  nota     += "• Nome: "       + data.nome     + "\n";
  if (data.tipo)    nota += "• Tipo: "      + data.tipo     + "\n";
  if (data.empresa) nota += "• Empresa: "  + data.empresa  + "\n";
  if (data.email)   nota += "• E-mail: "   + data.email    + "\n";
  nota     += "• WhatsApp: "   + data.whatsapp + "\n";
  nota     += "• " + (data.source === "Site" ? "Serviço" : "O que procura") + ": " + data.servico;
  if (data.mensagem) nota += "\n• " + (data.source === "Site" ? "Mensagem" : "Campanha") + ": " + data.mensagem;

  UrlFetchApp.fetch(API_BASE + "/messages/", {
    method: "post", contentType: "application/json", headers: headers,
    payload: JSON.stringify({ organizationId: ORG_ID, chatId: chatId, message: nota, isPrivate: true }),
    muteHttpExceptions: true
  });

  return chatId;
}

// ============================================================
//  SALVA LEAD NO GOOGLE SHEETS
// ============================================================
function salvarLead(data, chatId) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

  // Cria cabeçalho se planilha estiver vazia
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["chatId","nome","whatsapp","servico","source","criadoEm","fu1","fu2","fu3","encerrado"]);
  }

  sheet.appendRow([
    chatId,
    data.nome,
    data.whatsapp,
    data.servico,
    data.source,
    new Date().toISOString(),
    false,  // follow-up 1 enviado?
    false,  // follow-up 2 enviado?
    false,  // follow-up 3 enviado?
    false   // encerrado?
  ]);

  Logger.log("Lead salvo na planilha: " + data.nome);
}

// ============================================================
//  VERIFICA SE O LEAD JÁ RESPONDEU
// ============================================================
function leadRespondeu(chatId) {
  const r = UrlFetchApp.fetch(API_BASE + "/chats/" + chatId + "/?organizationId=" + ORG_ID, {
    headers: { "Authorization": "Bearer " + UMBLER_TOKEN },
    muteHttpExceptions: true
  });
  const chat = JSON.parse(r.getContentText());
  // firstContactMessage existe = lead respondeu
  return !!(chat.firstContactMessage && chat.firstContactMessage.eventAtUTC);
}

// ============================================================
//  FOLLOW-UPS AUTOMÁTICOS (roda a cada hora pelo trigger)
// ============================================================
function verificarFollowUps() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const dados = sheet.getDataRange().getValues();
  const agora = new Date();

  for (let i = 1; i < dados.length; i++) {
    const [chatId, nome, whatsapp, servico, source, criadoEm, fu1, fu2, fu3, encerrado] = dados[i];

    if (encerrado) continue;

    // Se lead já respondeu, encerra follow-ups
    if (leadRespondeu(chatId)) {
      sheet.getRange(i + 1, 10).setValue(true); // encerrado = true
      Logger.log("Lead " + nome + " respondeu — follow-ups encerrados.");
      continue;
    }

    const criacao    = new Date(criadoEm);
    const horasPassadas = (agora - criacao) / (1000 * 60 * 60);
    const digitos    = whatsapp.replace(/\D/g, "");
    const toPhone    = digitos.startsWith("55") && digitos.length >= 12 ? "+" + digitos : "+55" + digitos;

    // Follow-up 1 — após 2h
    if (!fu1 && horasPassadas >= 2) {
      enviarFollowUp(chatId, toPhone, nome, 1);
      sheet.getRange(i + 1, 7).setValue(true);
      Logger.log("FU1 enviado para " + nome);
    }

    // Follow-up 2 — após 24h
    if (fu1 && !fu2 && horasPassadas >= 24) {
      enviarFollowUp(chatId, toPhone, nome, 2);
      sheet.getRange(i + 1, 8).setValue(true);
      Logger.log("FU2 enviado para " + nome);
    }

    // Follow-up 3 — após 48h + encerra
    if (fu2 && !fu3 && horasPassadas >= 48) {
      enviarFollowUp(chatId, toPhone, nome, 3);
      sheet.getRange(i + 1, 9).setValue(true);
      sheet.getRange(i + 1, 10).setValue(true); // encerra após último FU
      Logger.log("FU3 enviado para " + nome + " — encerrado.");
    }
  }
}

// ============================================================
//  MENSAGENS DE FOLLOW-UP
// ============================================================
function enviarFollowUp(chatId, toPhone, nome, numero) {
  const mensagens = {
    1: "Oi, *" + nome + "*! 👋 Só passando para ver se ficou alguma dúvida. Pode falar à vontade! 😊",
    2: "Olá, *" + nome + "*! Ainda estamos por aqui caso queira um orçamento ou conhecer nossos serviços. 🎨 Quando quiser é só chamar!",
    3: "*" + nome + "*, última mensagem da nossa parte! 😊 Se precisar de comunicação visual, brindes ou impressão, a *Gráfica Fábrica de Ideias* está sempre à disposição. Até logo!"
  };

  const headers = { "Authorization": "Bearer " + UMBLER_TOKEN };
  const msg     = mensagens[numero];

  // Envia para o WhatsApp do lead
  UrlFetchApp.fetch(API_BASE + "/messages/simplified/", {
    method: "post", contentType: "application/json", headers: headers,
    payload: JSON.stringify({ organizationId: ORG_ID, toPhone: toPhone, fromPhone: FROM_PHONE, message: msg }),
    muteHttpExceptions: true
  });

  // Registra na conversa como nota interna
  UrlFetchApp.fetch(API_BASE + "/messages/", {
    method: "post", contentType: "application/json", headers: headers,
    payload: JSON.stringify({ organizationId: ORG_ID, chatId: chatId, message: "🔁 Follow-up #" + numero + " enviado automaticamente.", isPrivate: true }),
    muteHttpExceptions: true
  });
}

// ============================================================
//  CRIA O TRIGGER AUTOMÁTICO (rodar UMA VEZ manualmente)
// ============================================================
function criarTrigger() {
  // Remove triggers antigos para não duplicar
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "verificarFollowUps") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Cria trigger que roda a cada 1 hora
  ScriptApp.newTrigger("verificarFollowUps")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("Trigger criado com sucesso — verificarFollowUps rodará a cada hora.");
}

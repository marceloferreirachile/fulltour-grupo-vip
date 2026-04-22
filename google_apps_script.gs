// ╔══════════════════════════════════════════════════════════════╗
// ║   FULL TOUR — GRUPO VIP 2026                                 ║
// ║   Google Apps Script — Cole este código no Apps Script       ║
// ║   da sua planilha Google Sheets                              ║
// ║                                                              ║
// ║   COMO INSTALAR:                                             ║
// ║   1. Abra o Google Sheets da planilha                        ║
// ║   2. Menu → Extensões → Apps Script                          ║
// ║   3. Apague o código existente e cole este aqui              ║
// ║   4. Salve (Ctrl+S)                                          ║
// ║   5. Clique em "Implantar" → "Nova implantação"              ║
// ║   6. Tipo: App da Web                                        ║
// ║   7. Executar como: Eu mesmo                                 ║
// ║   8. Quem tem acesso: Qualquer pessoa                        ║
// ║   9. Copie a URL gerada e cole no formulário HTML            ║
// ╚══════════════════════════════════════════════════════════════╝

// ── CONFIGURAÇÕES ──────────────────────────────────────────────
var CONFIG = {
  email_marcelo:   "marcelo@fulltour.com.br",
  nome_remetente:  "Full Tour — Grupo VIP",
  sheet_name:      "Inscrições",       // nome da aba na planilha
  prazo_limite:    "27/06/2026",
  link_formulario: "",                  // preencha após implantar (URL do formulário)
};

// Colunas da planilha (ordem das colunas, começando em 1)
var COLS = {
  id:           1,
  timestamp:    2,
  nome:         3,
  email:        4,
  whatsapp:     5,
  dia1_opcao:   6,
  dia1_adultos: 7,
  dia1_criancas:8,
  dia2_opcao:   9,
  dia2_adultos: 10,
  dia2_criancas:11,
  farellones:   12,
  far_gratuito: 13,
  far_crianca:  14,
  far_adulto_geral: 15,
  far_adulto_ski:   16,
  clp_cotacao:  17,
  alyan:        18,
  total_passeios:  19,
  total_ingressos: 20,
  total_geral:  21,
  pct_entrada:  22,
  valor_entrada:23,
  saldo:        24,
  pagamento:    25,
  status:       26,
  data_confirmacao: 27,
  observacoes:  28,
};

// ── CABEÇALHOS ─────────────────────────────────────────────────
var HEADERS = [
  "ID", "Data/Hora", "Nome", "E-mail", "WhatsApp",
  "Dia 1 — 28/06", "Adultos Dia 1", "Crianças Dia 1",
  "Dia 2 — 29/06", "Adultos Dia 2", "Crianças Dia 2",
  "Farellones", "Far. Gratuito", "Far. Criança", "Far. Adulto Geral", "Far. Adulto Ski",
  "Cotação CLP", "Ticket Alyan",
  "Total Passeios R$", "Total Ingressos R$", "Total Geral R$",
  "% Entrada", "Valor Entrada R$", "Saldo R$",
  "Forma Pagamento", "STATUS", "Data Confirmação", "Observações"
];

// ── RECEBE SUBMISSÃO DO FORMULÁRIO ──────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var ws   = ss.getSheetByName(CONFIG.sheet_name);

    // Cria a aba e cabeçalhos se não existir
    if (!ws) {
      ws = ss.insertSheet(CONFIG.sheet_name);
      configurarPlanilha(ws);
    }

    // Gera ID sequencial baseado na posição da linha (linhas 1-3 são cabeçalhos)
    // Não lê o valor da célula anterior para evitar propagação de erros #NUM!
    var lastRow = ws.getLastRow();
    var newId   = lastRow <= 3 ? 1 : lastRow - 2;

    // Monta a linha de dados
    var row = new Array(HEADERS.length).fill("");
    row[COLS.id - 1]               = newId;
    row[COLS.timestamp - 1]        = new Date();
    row[COLS.nome - 1]             = data.nome || "";
    row[COLS.email - 1]            = data.email || "";
    row[COLS.whatsapp - 1]         = data.whatsapp || "";
    row[COLS.dia1_opcao - 1]       = data.dia1_label || "";
    row[COLS.dia1_adultos - 1]     = parseInt(data.dia1_adultos) || 0;
    row[COLS.dia1_criancas - 1]    = parseInt(data.dia1_criancas) || 0;
    row[COLS.dia2_opcao - 1]       = data.dia2_label || "";
    row[COLS.dia2_adultos - 1]     = parseInt(data.dia2_adultos) || 0;
    row[COLS.dia2_criancas - 1]    = parseInt(data.dia2_criancas) || 0;
    row[COLS.farellones - 1]       = data.farellones || "Não";
    row[COLS.far_gratuito - 1]     = parseInt(data.far_gratuito) || 0;
    row[COLS.far_crianca - 1]      = parseInt(data.far_crianca) || 0;
    row[COLS.far_adulto_geral - 1] = parseInt(data.far_adulto_geral) || 0;
    row[COLS.far_adulto_ski - 1]   = parseInt(data.far_adulto_ski) || 0;
    row[COLS.clp_cotacao - 1]      = parseFloat(data.clp_cotacao) || 0;
    row[COLS.alyan - 1]            = data.alyan || "N/A";
    row[COLS.total_passeios - 1]   = parseFloat(data.total_passeios) || 0;
    row[COLS.total_ingressos - 1]  = parseFloat(data.total_ingressos) || 0;
    row[COLS.total_geral - 1]      = parseFloat(data.total_geral) || 0;
    row[COLS.pct_entrada - 1]      = data.pct_entrada || "";
    row[COLS.valor_entrada - 1]    = parseFloat(data.valor_entrada) || 0;
    row[COLS.saldo - 1]            = parseFloat(data.saldo) || 0;
    row[COLS.pagamento - 1]        = data.pagamento || "";
    row[COLS.status - 1]           = "⏳ Aguardando";
    row[COLS.data_confirmacao - 1] = "";
    row[COLS.observacoes - 1]      = "";

    ws.appendRow(row);
    formatarUltimaLinha(ws);

    // Envia e-mails
    enviarEmailCliente(data);
    enviarNotificacaoMarcelo(data, newId);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, id: newId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── CONFIGURA A PLANILHA (primeira vez) ────────────────────────
function configurarPlanilha(ws) {
  // Cabeçalho de título
  ws.getRange(1, 1, 1, HEADERS.length).merge()
    .setValue("✈  FULL TOUR — GRUPO VIP | Controle de Inscrições | Junho 2026")
    .setBackground("#071a2e")
    .setFontColor("#ffffff")
    .setFontSize(13)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  ws.setRowHeight(1, 30);

  // Aviso de prazo
  ws.getRange(2, 1, 1, HEADERS.length).merge()
    .setValue("⏰  Valores válidos até 27/06/2026 · Sistema encerra automaticamente · Reabertura 28/06 com valores atualizados")
    .setBackground("#fef5ec")
    .setFontColor("#7a4000")
    .setFontSize(9)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  ws.setRowHeight(2, 18);

  // Cabeçalhos das colunas
  var headerRange = ws.getRange(3, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS])
    .setBackground("#0a6f73")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setWrap(true);
  ws.setRowHeight(3, 32);

  // Congela as 3 primeiras linhas
  ws.setFrozenRows(3);

  // Larguras das colunas
  var widths = [50,140,200,220,130,220,90,90,220,90,90,90,80,80,100,100,90,180,
                110,110,110,90,110,110,130,130,130,250];
  for (var i = 0; i < widths.length && i < HEADERS.length; i++) {
    ws.setColumnWidth(i + 1, widths[i]);
  }

  // Validação dropdown STATUS
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["⏳ Aguardando", "✅ Confirmado", "❌ Cancelado"], true)
    .build();
  ws.getRange(4, COLS.status, 500, 1).setDataValidation(statusRule);

  // Formato de moeda nas colunas de valor
  var moneyFmt = 'R$ #,##0.00';
  [COLS.total_passeios, COLS.total_ingressos, COLS.total_geral,
   COLS.valor_entrada, COLS.saldo].forEach(function(col) {
    ws.getRange(4, col, 500, 1).setNumberFormat(moneyFmt);
  });

  // Formato de data/hora
  ws.getRange(4, COLS.timestamp, 500, 1).setNumberFormat("dd/MM/yyyy HH:mm");
  ws.getRange(4, COLS.data_confirmacao, 500, 1).setNumberFormat("dd/MM/yyyy HH:mm");
}

// ── FORMATA A ÚLTIMA LINHA INSERIDA ────────────────────────────
function formatarUltimaLinha(ws) {
  var lastRow = ws.getLastRow();
  if (lastRow < 4) return;
  ws.getRange(lastRow, 1, 1, HEADERS.length)
    .setBackground("#f5f7fa")
    .setVerticalAlignment("middle");
  ws.setRowHeight(lastRow, 22);
}

// ── ONCHANGE: formata STATUS com cor ───────────────────────────
function onEdit(e) {
  var ws    = e.range.getSheet();
  var col   = e.range.getColumn();
  var row   = e.range.getRow();
  var val   = e.range.getValue();

  if (ws.getName() !== CONFIG.sheet_name) return;
  if (col !== COLS.status || row < 4)     return;

  // Remove formatação anterior
  e.range.setBackground(null).setFontColor(null).setFontWeight(null);

  if (val === "✅ Confirmado") {
    e.range.setBackground("#e9f7ef").setFontColor("#27ae60").setFontWeight("bold");
    // Registra data de confirmação
    ws.getRange(row, COLS.data_confirmacao).setValue(new Date());
    // Envia e-mail de confirmação ao cliente
    var emailCell = ws.getRange(row, COLS.email).getValue();
    var nomeCell  = ws.getRange(row, COLS.nome).getValue();
    var totalCell = ws.getRange(row, COLS.total_geral).getValue();
    var saldoCell = ws.getRange(row, COLS.saldo).getValue();
    var pagCell   = ws.getRange(row, COLS.pagamento).getValue();
    if (emailCell) enviarEmailConfirmacao(emailCell, nomeCell, totalCell, saldoCell, pagCell);

  } else if (val === "❌ Cancelado") {
    e.range.setBackground("#fdf0ef").setFontColor("#c0392b").setFontWeight("bold");
    // Envia e-mail de cancelamento ao cliente
    var emailCel  = ws.getRange(row, COLS.email).getValue();
    var nomeCel   = ws.getRange(row, COLS.nome).getValue();
    var dia1Cel   = ws.getRange(row, COLS.dia1_opcao).getValue();
    var dia2Cel   = ws.getRange(row, COLS.dia2_opcao).getValue();
    var farCel    = ws.getRange(row, COLS.farellones).getValue();
    var totalCel  = ws.getRange(row, COLS.total_geral).getValue();
    if (emailCel) enviarEmailCancelamento(emailCel, nomeCel, dia1Cel, dia2Cel, farCel, totalCel);

  } else if (val === "⏳ Aguardando") {
    e.range.setBackground("#fef9e7").setFontColor("#e67e22").setFontWeight("bold");
  }
}

// ── E-MAIL PARA O CLIENTE (documentação da inscrição) ──────────
function enviarEmailCliente(data) {
  var fmt = function(v) {
    return "R$ " + parseFloat(v || 0).toFixed(2).replace(".", ",").replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  };

  // Montar linhas de pessoas
  var d1pessoas = (data.dia1_adultos || 1) + ' adulto(s)' + (parseInt(data.dia1_criancas) > 0 ? ' + ' + data.dia1_criancas + ' criança(s)' : '');
  var d2pessoas = (data.dia2_adultos || 1) + ' adulto(s)' + (parseInt(data.dia2_criancas) > 0 ? ' + ' + data.dia2_criancas + ' criança(s)' : '');

  // Linha Alyan (só se dia1 tiver Alyan)
  var alyanLinha = '';
  if (data.alyan && data.alyan !== 'N/A') {
    var alyanLabel = data.alyan === 'PIX via Full Tour'
      ? '🔗 Link PIX enviado separado pela Full Tour (Nautt — câmbio do dia, sem IOF)'
      : '⚠️ Compra direta no site da Alyan — responsabilidade do cliente';
    alyanLinha = '<p style="font-size:12.5px;color:#444;margin:5px 0;"><strong>🍷 Ticket Alyan:</strong> ' + alyanLabel + '</p>';
  }

  // Linha Farellones detalhes
  var farLinha = '';
  if (data.farellones === 'Sim') {
    var farPagLabel = data.far_pagamento === 'PIX via Full Tour'
      ? '🔗 Link PIX enviado separado pela Full Tour (Nautt — câmbio do dia, sem IOF)'
      : '⚠️ Compra direta no parque — responsabilidade do cliente';
    var farIngressos = [];
    if (parseInt(data.far_gratuito)   > 0) farIngressos.push(data.far_gratuito + ' gratuito(s)');
    if (parseInt(data.far_crianca)    > 0) farIngressos.push(data.far_crianca + ' criança(s)');
    if (parseInt(data.far_adulto_geral) > 0) farIngressos.push(data.far_adulto_geral + ' adulto(s) geral');
    if (parseInt(data.far_adulto_ski)   > 0) farIngressos.push(data.far_adulto_ski + ' adulto(s) ski/snow');
    farLinha = '<p style="font-size:12.5px;color:#444;margin:5px 0;"><strong>☃️ Farellones — 30/07:</strong> Sim</p>'
      + (farIngressos.length ? '<p style="font-size:12px;color:#555;margin:3px 0 3px 16px;">→ Ingressos: ' + farIngressos.join(', ') + '</p>' : '')
      + '<p style="font-size:12px;color:#555;margin:3px 0 5px 16px;">→ Pagamento: ' + farPagLabel + '</p>';
  } else {
    farLinha = '<p style="font-size:12.5px;color:#444;margin:5px 0;"><strong>☃️ Farellones — 30/07:</strong> Não participa</p>';
  }

  var html = '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:560px;margin:0 auto;background:#f0f2f5;">'
    + '<div style="background:#071a2e;padding:20px 28px;border-bottom:3px solid #11A1A7;">'
    + '<span style="font-size:18px;font-weight:900;color:#fff;letter-spacing:2px;">FULL <span style="color:#11A1A7;">TOUR</span></span>'
    + '<span style="float:right;background:rgba(17,161,167,0.2);border:1px solid rgba(17,161,167,0.5);'
    + 'color:#b0f0f2;font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;">GRUPO VIP 2026</span>'
    + '</div>'
    + '<div style="background:#ffffff;padding:28px 28px 20px;border-left:1px solid #dde;border-right:1px solid #dde;">'
    + '<p style="font-size:21px;font-weight:800;color:#071a2e;margin:0 0 6px;">Inscrição recebida! ✓</p>'
    + '<p style="font-size:13px;color:#444;line-height:1.6;margin:0 0 22px;">Olá, <strong>' + data.nome + '</strong>! Sua inscrição no <strong>Grupo VIP Julho 2026</strong> foi registrada. Nossa equipe entrará em contato em breve com o link de pagamento.</p>'
    + '</div>'
    + '<div style="background:#f8f9fb;border:1px solid #dde;border-top:none;padding:22px 28px;">'
    + '<p style="font-size:10.5px;font-weight:800;color:#0a6f73;text-transform:uppercase;letter-spacing:1.5px;margin:0 0 14px;">📋 Resumo da sua inscrição</p>'
    + '<p style="font-size:12.5px;color:#222;margin:6px 0;"><strong>📅 Dia 1 — 28/07:</strong> ' + (data.dia1_label || '—') + ' — ' + d1pessoas + '</p>'
    + alyanLinha
    + '<p style="font-size:12.5px;color:#222;margin:6px 0;"><strong>📅 Dia 2 — 29/07:</strong> ' + (data.dia2_label || '—') + ' — ' + d2pessoas + '</p>'
    + farLinha
    + '<div style="border-top:1px solid #dde;margin:14px 0;"></div>'
    + '<p style="font-size:12.5px;color:#222;margin:5px 0;"><strong>💰 Total passeios:</strong> ' + fmt(data.total_geral) + '</p>'
    + '<p style="font-size:12.5px;color:#222;margin:5px 0;"><strong>💳 Entrada (' + data.pct_entrada + '):</strong> ' + fmt(data.valor_entrada) + '</p>'
    + (parseFloat(data.saldo) > 0 ? '<p style="font-size:12.5px;color:#222;margin:5px 0;"><strong>📆 Saldo restante:</strong> ' + fmt(data.saldo) + ' (até 48h antes do passeio)</p>' : '')
    + '<p style="font-size:12.5px;color:#222;margin:5px 0;"><strong>💳 Forma de pagamento:</strong> ' + (data.pagamento || '—') + '</p>'
    + '</div>'
    + '<div style="background:#fef5ec;border:1px solid #f0c090;border-top:none;padding:14px 28px;">'
    + '<p style="font-size:12px;color:#7a4000;margin:0;">⏰ <strong>Lembre-se:</strong> os valores são válidos para inscrições confirmadas até <strong>27/04/2026</strong>. Inscrições não pagas até esta data serão canceladas automaticamente.</p>'
    + '</div>'
    + '<div style="background:#ffffff;border:1px solid #dde;border-top:none;padding:14px 28px;">'
    + '<p style="font-size:12px;color:#666;margin:0;">Dúvidas? Fale com a Full Tour pelo <a href="https://wa.me/56982050413" style="color:#0a6f73;font-weight:bold;">WhatsApp</a>.</p>'
    + '</div>'
    + '<div style="background:#071a2e;padding:12px;text-align:center;">'
    + '<p style="font-size:10px;color:#8ab;margin:0;">Full Tour Chile · Atendimento 100% em Português · fulltour.com.br</p>'
    + '</div></div>';

  MailApp.sendEmail({
    to:       data.email,
    subject:  "Full Tour Grupo VIP — Inscrição recebida ✓",
    htmlBody: html,
    name:     CONFIG.nome_remetente,
    replyTo:  CONFIG.email_marcelo,
  });
}

// ── NOTIFICAÇÃO PARA MARCELO (simples, só para documentação) ────
function enviarNotificacaoMarcelo(data, id) {
  var fmt = function(v) {
    return "R$ " + parseFloat(v || 0).toFixed(2).replace(".", ",");
  };
  var html = '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:480px;margin:0 auto;background:#f0f2f5;">'
    + '<div style="background:#071a2e;padding:18px 26px;border-bottom:3px solid #11A1A7;">'
    + '<span style="font-size:17px;font-weight:900;color:#fff;letter-spacing:2px;">FULL <span style="color:#11A1A7;">TOUR</span></span>'
    + '<span style="float:right;background:rgba(17,161,167,0.2);border:1px solid rgba(17,161,167,0.5);color:#b0f0f2;font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;">NOVA INSCRIÇÃO — ID #' + id + '</span>'
    + '</div>'
    + '<div style="background:#ffffff;padding:24px 26px 18px;border-left:1px solid #dde;border-right:1px solid #dde;">'
    + '<p style="font-size:18px;font-weight:800;color:#071a2e;margin:0 0 16px;">📋 Nova inscrição recebida</p>'
    + '<p style="font-size:13px;margin:5px 0;"><strong>Nome:</strong> ' + data.nome + '</p>'
    + '<p style="font-size:13px;margin:5px 0;"><strong>E-mail:</strong> ' + data.email + '</p>'
    + '<p style="font-size:13px;margin:5px 0;"><strong>WhatsApp:</strong> ' + data.whatsapp + '</p>'
    + '</div>'
    + '<div style="background:#f8f9fb;border:1px solid #dde;border-top:none;padding:18px 26px;">'
    + '<p style="font-size:10px;font-weight:800;color:#0a6f73;text-transform:uppercase;letter-spacing:1.5px;margin:0 0 12px;">📋 Escolhas</p>'
    + '<p style="font-size:12.5px;color:#333;margin:5px 0;"><strong>Dia 1 — 28/07:</strong> ' + (data.dia1_label || '—') + '</p>'
    + '<p style="font-size:12.5px;color:#333;margin:5px 0;"><strong>Dia 2 — 29/07:</strong> ' + (data.dia2_label || '—') + '</p>'
    + '<p style="font-size:12.5px;color:#333;margin:5px 0;"><strong>Farellones — 30/07:</strong> ' + (data.farellones || 'Não') + '</p>'
    + '<div style="border-top:1px solid #dde;margin:12px 0;"></div>'
    + '<p style="font-size:12.5px;color:#333;margin:5px 0;"><strong>Total passeios:</strong> ' + fmt(data.total_geral) + '</p>'
    + '<p style="font-size:12.5px;color:#333;margin:5px 0;"><strong>Entrada (' + (data.pct_entrada||'') + '):</strong> ' + fmt(data.valor_entrada) + '</p>'
    + '<p style="font-size:12.5px;color:#333;margin:5px 0;"><strong>Pagamento:</strong> ' + (data.pagamento || '—') + '</p>'
    + '</div>'
    + '<div style="background:#071a2e;padding:12px;text-align:center;">'
    + '<p style="font-size:10px;color:#8ab;margin:0;">Acesse o Google Sheets para gerenciar o status desta inscrição.</p>'
    + '</div></div>';

  MailApp.sendEmail({
    to:       CONFIG.email_marcelo,
    subject:  "Nova inscrição #" + id + " — " + data.nome + " | Full Tour Grupo VIP",
    htmlBody: html,
    name:     "Full Tour — Sistema de Inscrições",
  });
}

// ── E-MAIL DE CONFIRMAÇÃO DE PAGAMENTO (disparado ao marcar ✅) ─
function enviarEmailConfirmacao(email, nome, total, saldo, pagamento) {
  var fmt = function(v) { return "R$ " + parseFloat(v || 0).toFixed(2).replace(".", ","); };
  var html = '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:560px;margin:0 auto;">'
    + '<div style="background:#071a2e;padding:20px 28px;border-bottom:3px solid #11A1A7;">'
    + '<span style="font-size:18px;font-weight:900;color:#fff;letter-spacing:2px;">FULL <span style="color:#11A1A7;">TOUR</span></span>'
    + '</div>'
    + '<div style="background:#fff;padding:28px;border:1px solid #eee;">'
    + '<div style="text-align:center;margin-bottom:20px;">'
    + '<div style="width:60px;height:60px;background:linear-gradient(135deg,#27ae60,#1e8449);border-radius:50%;margin:0 auto 12px;display:flex;align-items:center;justify-content:center;font-size:28px;">✓</div>'
    + '<p style="font-size:20px;font-weight:800;color:#1a1a2e;margin:0;">Reserva confirmada!</p>'
    + '</div>'
    + '<p style="font-size:13px;color:#444;line-height:1.6;">Olá, <strong>' + nome + '</strong>! Seu pagamento foi confirmado e sua vaga está <strong>garantida</strong>. Nos vemos em junho! 🇨🇱</p>'
    + (saldo > 0 ? '<div style="background:#fef5ec;border:1px solid #f0c090;border-radius:8px;padding:14px;margin-top:16px;">'
    + '<p style="font-size:12px;color:#7a4000;margin:0;">💰 <strong>Saldo restante:</strong> ' + fmt(saldo) + ' — a ser pago até <strong>48h antes do primeiro passeio</strong> (26/06/2026).</p>'
    + '</div>' : '')
    + '<p style="font-size:12px;color:#888;margin-top:20px;">Qualquer dúvida, fale com a gente pelo WhatsApp!</p>'
    + '</div>'
    + '<div style="background:#f5f7fa;padding:14px;text-align:center;border-top:1px solid #eee;">'
    + '<p style="font-size:10px;color:#aaa;margin:0;">Full Tour Chile · Atendimento 100% em Português</p>'
    + '</div></div>';

  MailApp.sendEmail({
    to:       email,
    subject:  "Full Tour Grupo VIP — Reserva confirmada! 🎉",
    htmlBody: html,
    name:     CONFIG.nome_remetente,
    replyTo:  CONFIG.email_marcelo,
  });
}

// ── CANCELAMENTO AUTOMÁTICO (agendar para 27/06 às 23h) ────────
// Para agendar: Apps Script → Acionadores → + Adicionar acionador
// Função: executarCancelamentosAutomaticos
// Tipo: Por hora/data → Dia específico → 27/06/2026 → 23:00
function executarCancelamentosAutomaticos() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ws  = ss.getSheetByName(CONFIG.sheet_name);
  if (!ws) return;

  var lastRow = ws.getLastRow();
  var count   = 0;

  for (var row = 4; row <= lastRow; row++) {
    var status = ws.getRange(row, COLS.status).getValue();
    if (String(status).trim() !== "⏳ Aguardando") continue;

    var nome  = ws.getRange(row, COLS.nome).getValue();
    var email = ws.getRange(row, COLS.email).getValue();
    var dia1  = ws.getRange(row, COLS.dia1_opcao).getValue();
    var dia2  = ws.getRange(row, COLS.dia2_opcao).getValue();
    var far   = ws.getRange(row, COLS.farellones).getValue();
    var total = ws.getRange(row, COLS.total_geral).getValue();

    if (!email) continue;

    try {
      enviarEmailCancelamento(email, nome, dia1, dia2, far, total);
      ws.getRange(row, COLS.status).setValue("❌ Cancelado")
        .setBackground("#fdf0ef").setFontColor("#c0392b").setFontWeight("bold");
      count++;
      Utilities.sleep(1000); // 1s entre envios para não exceder cota
    } catch(e) {
      Logger.log("Erro ao cancelar " + email + ": " + e);
    }
  }

  // Notifica Marcelo com relatório
  MailApp.sendEmail({
    to:      CONFIG.email_marcelo,
    subject: "Full Tour Grupo VIP — Cancelamentos enviados: " + count + " inscrições",
    body:    "Processo automático de cancelamento concluído.\n"
           + count + " inscrições foram canceladas e notificadas por e-mail.\n"
           + "Acesse o Google Sheets para ver o relatório completo.",
    name:    "Full Tour — Sistema de Inscrições",
  });

  Logger.log("Cancelamentos enviados: " + count);
}

// ── E-MAIL DE CANCELAMENTO ──────────────────────────────────────
function enviarEmailCancelamento(email, nome, dia1, dia2, farellones, total) {
  var fmt = function(v) { return "R$ " + parseFloat(v || 0).toFixed(2).replace(".", ","); };
  var html = '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:560px;margin:0 auto;">'
    + '<div style="background:#071a2e;padding:20px 28px;border-bottom:3px solid #11A1A7;">'
    + '<span style="font-size:18px;font-weight:900;color:#fff;letter-spacing:2px;">FULL <span style="color:#11A1A7;">TOUR</span></span>'
    + '</div>'
    + '<div style="background:#fff;padding:28px;border:1px solid #eee;">'
    + '<p style="font-size:18px;font-weight:800;color:#1a1a2e;margin:0 0 8px;">Olá, ' + nome + '!</p>'
    + '<p style="font-size:13px;color:#666;line-height:1.6;margin:0 0 20px;">Sua inscrição no <strong>Full Tour Grupo VIP — Junho 2026</strong> não foi confirmada pois o pagamento não foi identificado até o prazo de <strong>27/06/2026</strong>.</p>'
    + '<div style="background:#fdf0ef;border:1.5px solid #f0b0a8;border-radius:10px;padding:16px;margin-bottom:20px;">'
    + '<p style="font-size:12px;font-weight:800;color:#c0392b;margin:0 0 10px;">❌ Inscrição não confirmada</p>'
    + '<p style="font-size:12px;color:#666;margin:3px 0;"><strong>Dia 1 (28/06):</strong> ' + (dia1 || '—') + '</p>'
    + '<p style="font-size:12px;color:#666;margin:3px 0;"><strong>Dia 2 (29/06):</strong> ' + (dia2 || '—') + '</p>'
    + '<p style="font-size:12px;color:#666;margin:3px 0;"><strong>Farellones (30/06):</strong> ' + (farellones || '—') + '</p>'
    + '</div>'
    + '<div style="background:#e0f7f8;border:1.5px solid rgba(17,161,167,0.35);border-radius:10px;padding:16px;margin-bottom:24px;">'
    + '<p style="font-size:13px;font-weight:800;color:#0a6f73;margin:0 0 6px;">🔄 Segundo Lote — a partir de 28/06</p>'
    + '<p style="font-size:12px;color:#444;line-height:1.6;margin:0 0 12px;">O formulário de inscrições reabre em <strong>28 de junho</strong> com valores atualizados. Você pode garantir sua vaga no segundo lote!</p>'
    + (CONFIG.link_formulario ? '<a href="' + CONFIG.link_formulario + '" style="display:inline-block;background:linear-gradient(135deg,#11A1A7,#0a6f73);color:#fff;font-size:13px;font-weight:800;padding:11px 22px;border-radius:8px;text-decoration:none;">Inscrever-me no 2º Lote →</a>' : '')
    + '</div>'
    + '<p style="font-size:12px;color:#888;">Dúvidas? Fale conosco pelo WhatsApp.</p>'
    + '</div>'
    + '<div style="background:#f5f7fa;padding:14px;text-align:center;border-top:1px solid #eee;">'
    + '<p style="font-size:10px;color:#aaa;margin:0;">Full Tour Chile · Atendimento 100% em Português · fulltour.com.br</p>'
    + '</div></div>';

  MailApp.sendEmail({
    to:       email,
    subject:  "Full Tour Grupo VIP — Sua inscrição não foi confirmada",
    htmlBody: html,
    name:     CONFIG.nome_remetente,
    replyTo:  CONFIG.email_marcelo,
  });
}

// ── HANDLER GET (teste de conectividade) ───────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "online", app: "Full Tour Grupo VIP" }))
    .setMimeType(ContentService.MimeType.JSON);
}

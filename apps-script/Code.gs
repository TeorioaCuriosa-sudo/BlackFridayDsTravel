/**
 * =========================================================
 *  Black Friday — Think Black Travel Far
 *  Backend em Google Apps Script
 *
 *  Por cada resposta ao formulário:
 *   1. Gera um comprovativo PDF do consentimento (data/hora, e-mail,
 *      versão e texto aceite, IP, user agent, hash de integridade)
 *      e guarda-o no Google Drive.
 *   2. Acrescenta uma linha no Google Sheet (com coluna Sim/Não para
 *      novas campanhas e link para o PDF).
 *   3. Atualiza uma cópia .xlsx (Excel) do Sheet na mesma pasta do Drive.
 *   4. Envia o e-mail automático de agradecimento.
 *
 *  Primeira utilização: correr a função setup() uma vez (ver README).
 * =========================================================
 */

// ---------------- CONFIGURAÇÃO ----------------
const CONFIG = {
  CAMPANHA: 'Black Friday 2026',
  MARCA: 'Think Black Travel Far',

  // Pasta principal no Drive (criada pelo setup()).
  NOME_PASTA: 'Black Friday 2026 — Think Black Travel Far',
  NOME_SUBPASTA_PDF: 'Comprovativos de Consentimento',
  NOME_SHEET: 'Respostas Black Friday 2026',
  NOME_EXCEL: 'Respostas Black Friday 2026.xlsx',

  // E-mail
  // Endereço do domínio configurado no Gmail em "Enviar email como".
  // Se ficar vazio (''), os e-mails saem da conta Google que corre o script.
  EMAIL_FROM: '',
  EMAIL_NOME_REMETENTE: 'Think Black Travel Far',
  EMAIL_REPLY_TO: 'geral@exemplo.pt',        // <-- e-mail monitorizado para respostas
  EMAIL_NO_REPLY: false,                     // true só funciona em contas Google Workspace
  EMAIL_ASSUNTO: 'Obrigado pela sua participação na Black Friday!',
  LOGO_URL: '',                              // URL público do logo (opcional, para o e-mail)

  // Identificação do responsável pelo tratamento (aparece no PDF)
  RESPONSAVEL_TRATAMENTO: '[Nome da empresa, NIF, morada]',
  CONTACTO_PRIVACIDADE: 'privacidade@exemplo.pt',

  // Versão em vigor do texto dos consentimentos (tem de bater com script.js)
  CONSENT_VERSION: 'BF2026-v1',
};

/**
 * Texto OFICIAL dos consentimentos, por versão.
 * É este texto (do servidor) que fica no comprovativo, não o que vem do browser,
 * para que não possa ser adulterado. Se alterar o texto na landing page,
 * crie aqui uma nova versão e mantenha as antigas (histórico).
 */
const CONSENT_TEXTS = {
  'BF2026-v1': {
    privacidade:
      'Li e aceito a Política de Privacidade e autorizo o tratamento dos meus dados pessoais ' +
      'para responder a este pedido no âmbito da campanha Black Friday.',
    marketing:
      'Aceito receber por e-mail comunicações sobre novas campanhas, promoções e novidades. ' +
      'Posso retirar este consentimento a qualquer momento.',
  },
};

const HEADERS = [
  'ID', 'Data/Hora', 'Nome completo', 'E-mail', 'WhatsApp', 'Destino(s)',
  'Pessoas / Faixa etária', 'Orçamento', 'Datas', 'Informação adicional',
  'Aceita Política de Privacidade', 'Aceita novas campanhas', 'Versão do consentimento',
  'IP', 'User Agent', 'Comprovativo PDF', 'Hash SHA-256', 'E-mail enviado',
];

// ---------------- SETUP (correr 1 vez) ----------------
function setup() {
  const props = PropertiesService.getScriptProperties();

  const pasta = DriveApp.createFolder(CONFIG.NOME_PASTA);
  const subpasta = pasta.createFolder(CONFIG.NOME_SUBPASTA_PDF);

  const ss = SpreadsheetApp.create(CONFIG.NOME_SHEET);
  DriveApp.getFileById(ss.getId()).moveTo(pasta);

  const sh = ss.getSheets()[0];
  sh.setName('Respostas');
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#000000').setFontColor('#5ADFFF');
  sh.setFrozenRows(1);

  // Coluna "Aceita novas campanhas" com lista Sim/Não
  const colMkt = HEADERS.indexOf('Aceita novas campanhas') + 1;
  const regra = SpreadsheetApp.newDataValidation().requireValueInList(['Sim', 'Não'], true).build();
  sh.getRange(2, colMkt, sh.getMaxRows() - 1, 1).setDataValidation(regra);

  // Formatação condicional: verde para Sim, vermelho para Não
  const rng = sh.getRange(2, colMkt, sh.getMaxRows() - 1, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Sim')
      .setBackground('#d1fae5').setFontColor('#065f46').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Não')
      .setBackground('#fee2e2').setFontColor('#991b1b').setRanges([rng]).build(),
  ]);

  props.setProperties({
    FOLDER_ID: pasta.getId(),
    PDF_FOLDER_ID: subpasta.getId(),
    SHEET_ID: ss.getId(),
  });

  Logger.log('Setup concluído.');
  Logger.log('Pasta: ' + pasta.getUrl());
  Logger.log('Sheet: ' + ss.getUrl());
}

// ---------------- ENDPOINT ----------------
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    validar_(d);

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    let registo;
    try {
      registo = processar_(d);
    } finally {
      lock.releaseLock();
    }

    // Fora do lock: e-mail e Excel não precisam de bloquear outras submissões
    registo.emailOk = enviarEmail_(registo);
    marcarEmail_(registo.linha, registo.emailOk);
    try { atualizarExcel_(); } catch (err) { console.error('Excel: ' + err); }

    return json_({ ok: true, id: registo.id });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// Permite testar no browser que a Web App está online
function doGet() {
  return json_({ ok: true, status: 'online', campanha: CONFIG.CAMPANHA });
}

// ---------------- LÓGICA ----------------
function validar_(d) {
  const obrig = ['nome', 'email', 'whatsapp', 'destino', 'pessoas', 'orcamento', 'datas'];
  obrig.forEach(k => {
    if (!d[k] || !String(d[k]).trim()) throw new Error('Campo obrigatório em falta: ' + k);
  });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(d.email)) throw new Error('E-mail inválido');
  if (d.privacidade !== true) throw new Error('Política de Privacidade não aceite');
  if (!CONSENT_TEXTS[d.consentVersion]) throw new Error('Versão de consentimento desconhecida');
}

function processar_(d) {
  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();
  const agora = new Date();

  const id = 'BF-' + Utilities.formatDate(agora, tz, 'yyyyMMdd-HHmmss') + '-' +
             Utilities.getUuid().slice(0, 4).toUpperCase();
  const dataHora = Utilities.formatDate(agora, tz, "dd/MM/yyyy HH:mm:ss '(" + tz + ")'");
  const isoUTC = agora.toISOString();
  const textos = CONSENT_TEXTS[d.consentVersion];
  const marketing = d.marketing === true;

  // Hash de integridade do registo (prova de que o PDF não foi alterado depois)
  const base = [id, isoUTC, d.email, d.consentVersion, 'privacidade=true',
                'marketing=' + marketing, d.ip || '', textos.privacidade,
                marketing ? textos.marketing : ''].join('|');
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, base, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');

  const registo = {
    id, dataHora, isoUTC, hash, marketing, textos,
    nome: d.nome, email: d.email, ip: d.ip || 'Não disponível',
    userAgent: d.userAgent || '', versao: d.consentVersion, pagina: d.pagina || '',
  };

  // 1. PDF
  const pdf = gerarPdf_(registo, props.getProperty('PDF_FOLDER_ID'));
  registo.pdfUrl = pdf.getUrl();

  // 2. Sheet
  const sh = SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getSheetByName('Respostas');
  sh.appendRow([
    id, dataHora, d.nome, d.email, d.whatsapp, d.destino, d.pessoas, d.orcamento,
    d.datas, d.info || '', 'Sim', marketing ? 'Sim' : 'Não', d.consentVersion,
    registo.ip, registo.userAgent, registo.pdfUrl, hash, 'A enviar…',
  ].map(seguro_));
  registo.linha = sh.getLastRow();

  return registo;
}

// Evita "formula injection" (valores começados por = + - @ seriam interpretados como fórmulas)
function seguro_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function marcarEmail_(linha, ok) {
  const sh = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'))
    .getSheetByName('Respostas');
  sh.getRange(linha, HEADERS.indexOf('E-mail enviado') + 1).setValue(ok ? 'Sim' : 'Erro');
}

// ---------------- PDF ----------------
function gerarPdf_(r, pastaId) {
  const esc = escHtml_;
  const linha = (k, v) =>
    `<tr><th>${esc(k)}</th><td>${esc(v)}</td></tr>`;

  const html = `
  <html><head><meta charset="utf-8"><style>
    body{font-family:Arial,Helvetica,sans-serif;color:#111;font-size:11px;margin:28px}
    h1{font-size:18px;margin:0 0 4px}
    .sub{color:#555;margin:0 0 18px}
    .bar{height:6px;background:#000;border-bottom:3px solid #5ADFFF;margin-bottom:18px}
    table{width:100%;border-collapse:collapse;margin-bottom:16px}
    th,td{border:1px solid #ccc;padding:7px 9px;text-align:left;vertical-align:top}
    th{background:#f2f2f2;width:32%}
    .box{border:1px solid #ccc;padding:10px 12px;margin-bottom:10px}
    .estado{font-weight:bold}
    .sim{color:#065f46}.nao{color:#991b1b}
    .hash{font-family:monospace;font-size:9px;word-break:break-all}
    .foot{color:#666;font-size:9px;margin-top:22px}
  </style></head><body>
    <div class="bar"></div>
    <h1>Comprovativo de Consentimento</h1>
    <p class="sub">${esc(CONFIG.MARCA)} — ${esc(CONFIG.CAMPANHA)}</p>

    <table>
      ${linha('ID do registo', r.id)}
      ${linha('Data e hora', r.dataHora)}
      ${linha('Data e hora (UTC, ISO 8601)', r.isoUTC)}
      ${linha('Nome', r.nome)}
      ${linha('E-mail do titular', r.email)}
      ${linha('Endereço IP', r.ip)}
      ${linha('Navegador (User Agent)', r.userAgent)}
      ${linha('Página de origem', r.pagina)}
      ${linha('Versão do texto aceite', r.versao)}
    </table>

    <h3>Consentimentos</h3>
    <div class="box">
      <p class="estado sim">✔ ACEITE — Política de Privacidade (obrigatório)</p>
      <p>"${esc(r.textos.privacidade)}"</p>
    </div>
    <div class="box">
      <p class="estado ${r.marketing ? 'sim' : 'nao'}">
        ${r.marketing ? '✔ ACEITE' : '✘ NÃO ACEITE'} — Comunicações de marketing por e-mail (opcional)
      </p>
      <p>"${esc(r.textos.marketing)}"</p>
    </div>

    <h3>Integridade</h3>
    <p>Hash SHA-256 do registo:</p>
    <p class="hash">${esc(r.hash)}</p>

    <p class="foot">
      Responsável pelo tratamento: ${esc(CONFIG.RESPONSAVEL_TRATAMENTO)}.
      Contacto para exercício de direitos (acesso, retificação, apagamento, oposição e retirada
      do consentimento): ${esc(CONFIG.CONTACTO_PRIVACIDADE)}.
      Documento gerado automaticamente no momento da submissão do formulário.
    </p>
  </body></html>`;

  const nomeFicheiro = `${r.id}_${r.email.replace(/[^a-z0-9@._-]/gi, '_')}.pdf`;
  const blob = Utilities.newBlob(html, 'text/html', 'tmp.html').getAs('application/pdf').setName(nomeFicheiro);
  return DriveApp.getFolderById(pastaId).createFile(blob);
}

// ---------------- EXCEL (.xlsx) ----------------
// Exporta o Google Sheet para .xlsx e substitui sempre o mesmo ficheiro na pasta
// (usa o serviço avançado "Drive API" — ativado no appsscript.json).
function atualizarExcel_() {
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID');
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
  const blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
  }).getBlob().setName(CONFIG.NOME_EXCEL);

  const xlsxId = props.getProperty('XLSX_ID');
  if (xlsxId) {
    try {
      Drive.Files.update({}, xlsxId, blob);
      return;
    } catch (err) {
      console.warn('Ficheiro .xlsx não encontrado, a criar novo: ' + err);
    }
  }
  const novo = Drive.Files.create(
    { name: CONFIG.NOME_EXCEL, parents: [props.getProperty('FOLDER_ID')] }, blob);
  props.setProperty('XLSX_ID', novo.id);
}

// ---------------- E-MAIL ----------------
function enviarEmail_(r) {
  try {
    const esc = escHtml_;
    const primeiroNome = String(r.nome).trim().split(/\s+/)[0];

    const blocoMarketing = r.marketing ? `
      <p style="margin:0 0 16px;padding:14px 16px;background:#0f0f10;border-left:3px solid #5ADFFF;color:#e4e4e7">
        Confirmamos que aceitou receber por e-mail as nossas <strong>novas campanhas e promoções</strong>.
        Pode retirar este consentimento a qualquer momento respondendo a este e-mail
        ou escrevendo para <a href="mailto:${esc(CONFIG.CONTACTO_PRIVACIDADE)}" style="color:#5ADFFF">${esc(CONFIG.CONTACTO_PRIVACIDADE)}</a>.
      </p>` : '';

    const logo = CONFIG.LOGO_URL
      ? `<img src="${esc(CONFIG.LOGO_URL)}" alt="${esc(CONFIG.MARCA)}" width="260" style="display:block;margin:0 auto 24px;max-width:100%">`
      : `<p style="text-align:center;font-size:20px;font-weight:bold;color:#fff;margin:0 0 24px">${esc(CONFIG.MARCA)}</p>`;

    const html = `
    <div style="background:#000;padding:32px 16px;font-family:Arial,Helvetica,sans-serif">
      <div style="max-width:560px;margin:0 auto;background:#000;color:#f4f4f5">
        ${logo}
        <h1 style="font-size:22px;color:#5ADFFF;margin:0 0 16px">Obrigado, ${esc(primeiroNome)}!</h1>
        <p style="margin:0 0 16px;line-height:1.6">
          Agradecemos a sua participação na nossa campanha <strong>${esc(CONFIG.CAMPANHA)}</strong>.
          Recebemos o seu pedido e a nossa equipa vai preparar uma proposta à sua medida.
          Entraremos em contacto consigo muito em breve.
        </p>
        ${blocoMarketing}
        <p style="margin:24px 0 0;line-height:1.6">Boas viagens,<br><strong>Equipa ${esc(CONFIG.MARCA)}</strong></p>
        <hr style="border:0;border-top:1px solid #27272a;margin:28px 0 12px">
        <p style="font-size:11px;color:#71717a;line-height:1.5;margin:0">
          Recebeu este e-mail porque submeteu o formulário da campanha ${esc(CONFIG.CAMPANHA)}.
          Referência: ${esc(r.id)}. ${esc(CONFIG.RESPONSAVEL_TRATAMENTO)}.
        </p>
      </div>
    </div>`;

    const texto =
      `Obrigado, ${primeiroNome}!\n\n` +
      `Agradecemos a sua participação na campanha ${CONFIG.CAMPANHA}. ` +
      `A nossa equipa vai entrar em contacto consigo muito em breve.\n\n` +
      (r.marketing
        ? `Confirmamos que aceitou receber por e-mail as nossas novas campanhas e promoções. ` +
          `Pode retirar este consentimento a qualquer momento escrevendo para ${CONFIG.CONTACTO_PRIVACIDADE}.\n\n`
        : '') +
      `Equipa ${CONFIG.MARCA}\nReferência: ${r.id}`;

    const opts = {
      htmlBody: html,
      name: CONFIG.EMAIL_NOME_REMETENTE,
    };
    if (CONFIG.EMAIL_NO_REPLY) opts.noReply = true;
    else opts.replyTo = CONFIG.EMAIL_REPLY_TO;

    if (CONFIG.EMAIL_FROM) {
      // Envia a partir do e-mail do domínio (tem de existir como alias em "Enviar email como")
      const aliases = GmailApp.getAliases();
      if (aliases.indexOf(CONFIG.EMAIL_FROM) === -1) {
        throw new Error('O endereço ' + CONFIG.EMAIL_FROM +
          ' não está configurado em Gmail > Definições > Contas > Enviar email como. Alias disponíveis: ' +
          aliases.join(', '));
      }
      opts.from = CONFIG.EMAIL_FROM;
      GmailApp.sendEmail(r.email, CONFIG.EMAIL_ASSUNTO, texto, opts);
    } else {
      MailApp.sendEmail(Object.assign({ to: r.email, subject: CONFIG.EMAIL_ASSUNTO, body: texto }, opts));
    }
    return true;
  } catch (err) {
    console.error('Erro ao enviar e-mail: ' + err);
    return false;
  }
}

// ---------------- UTILITÁRIOS ----------------
function escHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------- TESTE MANUAL ----------------
// Corra testarSubmissao() no editor para simular uma resposta (use o seu e-mail).
function testarSubmissao() {
  const res = doPost({ postData: { contents: JSON.stringify({
    nome: 'Teste Silva', email: Session.getActiveUser().getEmail(), whatsapp: '+351 912 345 678',
    destino: 'Cabo Verde', pessoas: '2 adultos (30-40)', orcamento: '2.000 €',
    datas: '10 a 17 de março', info: 'Registo de teste',
    privacidade: true, marketing: true, consentVersion: CONFIG.CONSENT_VERSION,
    ip: '0.0.0.0', userAgent: 'Teste manual', pagina: 'editor',
  }) } });
  Logger.log(res.getContent());
}

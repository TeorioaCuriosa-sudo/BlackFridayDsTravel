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
 *  Campos do formulário (script.js): nome, apelido, e-mail, WhatsApp e Nº de pessoas.
 *
 *  Primeira utilização: correr a função setup() uma vez (ver README).
 *  Se o Sheet foi criado com a versão antiga do formulário (com Destino, Orçamento
 *  e Informação adicional), correr migrarSheet() uma vez.
 * =========================================================
 */

// Mude este texto sempre que alterar o código: abra o endereço /exec no browser e veja se mostra a versão nova.
const VERSAO_CODIGO = '2026-11-10-email-dominio-proprio';

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
  EMAIL_REPLY_TO: '',                        // e-mail para respostas. Vazio = as respostas vão para a conta que envia
  EMAIL_NO_REPLY: false,                     // true só funciona em contas Google Workspace
  EMAIL_ASSUNTO: 'Obrigado pela sua inscrição na Black Friday!',
  LOGO_URL: 'https://thinkblacktravelfar.teoriacuriosa.pt/logo.png',                              // URL público do logo (opcional, para o e-mail)
  LOGO_RODAPE_URL: 'https://thinkblacktravelfar.teoriacuriosa.pt/Elementos%20p%20Reels%20.png',
  // Endereço da Web App (o mesmo APPS_SCRIPT_URL do script.js, termina em /exec).
  // Usado no link "Cancelar subscrição" dos e-mails. Se mudar de implementação, atualize aqui.
  URL_WEBAPP: 'https://script.google.com/macros/s/AKfycbwtI1qCOS-kCEtYfICa9cRTXgpL8hRrxzmPF53B9aUY_hEM3aHScyDOs35zd8vckDhUMA/exec',

  // Identificação do responsável pelo tratamento (aparece no PDF)
  RESPONSAVEL_TRATAMENTO: 'TeoriaCuriosa, 515226939, Rua do Rio Novo 395 O, 4495-145, Amorim, Póvoa de Varzim',
  CONTACTO_PRIVACIDADE: 'geral@teoriacuriosa.pt',

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

// Colunas do Sheet. A ordem só conta na criação (setup). Depois, o script encontra
// cada coluna pelo NOME do cabeçalho, por isso mover colunas no Sheet não o estraga.
const HEADERS = [
  'ID', 'Data/Hora', 'Nome completo', 'E-mail', 'WhatsApp', 'Nº Pessoas',
  'Aceita Política de Privacidade', 'Aceita novas campanhas', 'Versão do consentimento',
  'IP', 'User Agent', 'Comprovativo PDF', 'Hash SHA-256', 'E-mail enviado',
  'Cancelou novas campanhas em',
  'Origem', 'Meio', 'Campanha', 'Conteúdo',
];

// Colunas com a origem do lead (vêm dos parâmetros UTM do link). Num Sheet que já existe,
// corra adicionarColunasOrigem() 1 vez para criar estas colunas.
const COLUNAS_ORIGEM = ['Origem', 'Meio', 'Campanha', 'Conteúdo'];

// Campos da versão antiga do formulário (só usados por migrarSheet)
const COLUNAS_REMOVIDAS = ['Destino(s)', 'Orçamento', 'Informação adicional'];
const COLUNA_PESSOAS_ANTIGA = 'Pessoas / Faixa etária';

// ---------------- SETUP (correr 1 vez) ----------------
// Protegido: se já foi executado, não cria pastas/Sheet duplicados (que faziam os dados
// "desaparecer" para outro Sheet). Para criar tudo de novo de propósito, use setupForcar().
function setup() {
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SHEET_ID');
  if (sheetId) {
    Logger.log('O setup já foi executado. Nada foi criado.');
    try { Logger.log('Sheet em uso: ' + SpreadsheetApp.openById(sheetId).getUrl()); } catch (e) { Logger.log('Sheet em uso: ' + sheetId); }
    Logger.log('Para criar tudo de novo (pasta e Sheet novos), corra setupForcar().');
    return;
  }
  setupCriar_();
}

function setupForcar() { setupCriar_(); }

function setupCriar_() {
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

// ---------------- MIGRAÇÃO DO SHEET (correr 1 vez, só se o Sheet já existia) ----------------
// Remove as colunas Destino(s), Orçamento e Informação adicional (o formulário já não as envia)
// e renomeia "Pessoas / Faixa etária" para "Nº Pessoas".
// Antes de apagar, guarda uma cópia do separador "Respostas" no mesmo ficheiro, com todos os dados antigos.
// É seguro correr mais do que uma vez: se já estiver migrado, não faz nada.
function migrarSheet() {
  const sh = abrirSheet_();
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const aRemover = COLUNAS_REMOVIDAS.filter(n => cab.indexOf(n) >= 0);
  const iPessoas = cab.indexOf(COLUNA_PESSOAS_ANTIGA);
  if (!aRemover.length && iPessoas < 0) {
    Logger.log('O Sheet já está atualizado. Nada a fazer.');
    return;
  }

  const tz = Session.getScriptTimeZone();
  const backup = sh.copyTo(sh.getParent());
  backup.setName('Backup antes da migração ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH-mm'));
  Logger.log('Cópia de segurança criada no separador: ' + backup.getName());

  if (iPessoas >= 0) sh.getRange(1, iPessoas + 1).setValue('Nº Pessoas');

  // Apaga da direita para a esquerda, para os números das colunas não se desfasarem
  aRemover.map(n => cab.indexOf(n) + 1).sort((a, b) => b - a).forEach(c => sh.deleteColumn(c));

  Logger.log('Colunas removidas: ' + (aRemover.join(', ') || '(nenhuma)'));
  try { atualizarExcel_(); } catch (err) { console.error('Excel: ' + err); }
  Logger.log('Migração concluída. Quando confirmar que está tudo certo, pode apagar o separador de backup.');
}

// ---------------- COLUNAS DE ORIGEM (correr 1 vez, só se o Sheet já existia) ----------------
// Acrescenta no fim do cabeçalho as colunas Origem, Meio, Campanha e Conteúdo.
// Não mexe em nenhum dado. É seguro correr mais do que uma vez: só cria as que faltam.
function adicionarColunasOrigem() {
  const sh = abrirSheet_();
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const faltam = COLUNAS_ORIGEM.filter(n => cab.indexOf(n) < 0);
  if (!faltam.length) {
    Logger.log('As colunas de origem já existem. Nada a fazer.');
    return;
  }
  sh.getRange(1, sh.getLastColumn() + 1, 1, faltam.length).setValues([faltam])
    .setFontWeight('bold').setBackground('#000000').setFontColor('#5ADFFF');
  Logger.log('Colunas acrescentadas: ' + faltam.join(', '));
  try { atualizarExcel_(); } catch (err) { console.error('Excel: ' + err); }
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
    const envio = enviarEmail_(registo);
    registo.emailOk = envio.ok;
    marcarEmail_(registo.linha, envio.ok ? 'Sim' : 'ERRO: ' + envio.erro);
    try { atualizarExcel_(); } catch (err) { console.error('Excel: ' + err); }

    return json_({ ok: true, id: registo.id });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// GET sem parâmetros: teste de que a Web App está online.
// GET com ?acao=cancelar&id=...&t=...: página de cancelamento de subscrição.
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.acao === 'cancelar') return paginaCancelar_(p.id, p.t);
  return json_({ ok: true, status: 'online', campanha: CONFIG.CAMPANHA, versao: VERSAO_CODIGO,
                replyToDefinido: !!CONFIG.EMAIL_REPLY_TO });
}

// ---------------- LÓGICA ----------------
function validar_(d) {
  const obrig = ['nome', 'email', 'whatsapp', 'pessoas'];
  obrig.forEach(k => {
    if (!d[k] || !String(d[k]).trim()) throw new Error('Campo obrigatório em falta: ' + k);
  });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(d.email)) throw new Error('E-mail inválido');
  const n = Number(String(d.pessoas).trim());
  if (!Number.isInteger(n) || n < 1 || n > 99) throw new Error('Nº de pessoas inválido');
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

  const origem = lerOrigem_(d);   // fonte / meio / campanha / conteúdo + User Agent já sem a etiqueta

  const registo = {
    id, dataHora, isoUTC, hash, marketing, textos,
    nome: String(d.nome).trim(),
    primeiroNome: String(d.primeiroNome || '').trim(),
    pessoas: String(d.pessoas).trim(),
    email: d.email, ip: d.ip || 'Não disponível',
    userAgent: origem.userAgent, versao: d.consentVersion, pagina: d.pagina || '',
  };

  // 1. PDF
  const pdf = gerarPdf_(registo, props.getProperty('PDF_FOLDER_ID'));
  registo.pdfUrl = pdf.getUrl();

  // 2. Sheet — cada valor vai para a coluna com o nome certo (lido do cabeçalho do Sheet)
  const sh = abrirSheet_();
  const valores = {
    'ID': id,
    'Data/Hora': dataHora,
    'Nome completo': registo.nome,
    'E-mail': d.email,
    'WhatsApp': d.whatsapp,
    'Nº Pessoas': registo.pessoas,
    [COLUNA_PESSOAS_ANTIGA]: registo.pessoas,   // Sheet ainda não migrado: continua a funcionar
    'Aceita Política de Privacidade': 'Sim',
    'Aceita novas campanhas': marketing ? 'Sim' : 'Não',
    'Versão do consentimento': d.consentVersion,
    'IP': registo.ip,
    'User Agent': registo.userAgent,
    'Comprovativo PDF': registo.pdfUrl,
    'Hash SHA-256': hash,
    'E-mail enviado': 'A enviar…',
    'Origem': origem.fonte,
    'Meio': origem.meio,
    'Campanha': origem.campanha,
    'Conteúdo': origem.conteudo,
  };
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(cab.map(h => seguro_(valores[h])));
  registo.linha = sh.getLastRow();

  return registo;
}

// Origem do lead. O site envia a etiqueta "[origem: fonte | meio | campanha | conteúdo]"
// no campo "origem" e também à frente do User Agent ("-" = não indicado).
// Só se aceita o formato exato (letras minúsculas, números, ponto, hífen, sublinhado),
// por isso um pedido forjado não consegue meter lixo nas colunas. Sem etiqueta válida, ficam vazias.
// A etiqueta é retirada do User Agent, para essa coluna e o comprovativo PDF ficarem só com o navegador.
const RE_ORIGEM = /^\[origem: ([a-z0-9._-]{1,30}) \| ([a-z0-9._-]{1,30}) \| ([a-z0-9._-]{1,30}) \| ([a-z0-9._-]{1,30})\]/;

function lerOrigem_(d) {
  const ua = String(d.userAgent || '');
  const daUA = RE_ORIGEM.exec(ua.trim());
  const doCampo = RE_ORIGEM.exec(String(d.origem || '').trim());
  const m = doCampo || daUA;
  const campo = i => (m && m[i] !== '-') ? m[i] : '';
  return {
    fonte: campo(1), meio: campo(2), campanha: campo(3), conteudo: campo(4),
    userAgent: daUA ? ua.trim().slice(daUA[0].length).trim() : ua,
  };
}

// Evita "formula injection" (valores começados por = + - @ seriam interpretados como fórmulas)
function seguro_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// Folha "Respostas" do Sheet em uso
function abrirSheet_() {
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'))
    .getSheetByName('Respostas');
}

// Número (a partir de 1) da coluna com este cabeçalho, lido do Sheet real
function coluna_(sh, nome) {
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const i = cab.indexOf(nome);
  if (i < 0) throw new Error('Coluna "' + nome + '" não existe no Sheet. Se o Sheet é anterior à mudança do formulário, corra migrarSheet().');
  return i + 1;
}

// Escreve na coluna "E-mail enviado": "Sim" ou "ERRO: <motivo>" (assim o erro fica visível no Sheet)
function marcarEmail_(linha, texto) {
  const sh = abrirSheet_();
  sh.getRange(linha, coluna_(sh, 'E-mail enviado')).setValue(texto);
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
function enviarEmail_(r, assunto) {
  try {
    const assuntoFinal = assunto || CONFIG.EMAIL_ASSUNTO;   // "assunto" só é usado nos testes
    const esc = escHtml_;
    // Usa o campo "Nome" do formulário tal como foi escrito; se não vier, a primeira palavra do nome completo
    const primeiroNome = r.primeiroNome
      ? String(r.primeiroNome).trim()
      : String(r.nome).trim().split(/\s+/)[0];

    const blocoMarketing = r.marketing ? `
      <p style="margin:0 0 16px;padding:14px 16px;background:#0f0f10;border-left:3px solid #5ADFFF;color:#e4e4e7">
        Confirmamos que aceitou receber por e-mail as nossas <strong>novas campanhas e promoções</strong>.
        Pode retirar este consentimento a qualquer momento:
        <a href="${esc(linkCancelar_(r.id))}" style="color:#5ADFFF;font-weight:bold">Cancelar subscrição</a>,
        ou escreva para <a href="mailto:${esc(CONFIG.CONTACTO_PRIVACIDADE)}" style="color:#5ADFFF">${esc(CONFIG.CONTACTO_PRIVACIDADE)}</a>.
      </p>` : '';

    const logo = CONFIG.LOGO_URL
      ? `<img src="${esc(CONFIG.LOGO_URL)}" alt="${esc(CONFIG.MARCA)}" width="260" style="display:block;margin:0 auto 24px;max-width:100%">`
      : `<p style="text-align:center;font-size:20px;font-weight:bold;color:#fff;margin:0 0 24px">${esc(CONFIG.MARCA)}</p>`;

    // Documento HTML completo (doctype, head, body): os filtros de spam desconfiam de fragmentos soltos
    const html = `<!doctype html>
    <html lang="pt-PT"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(CONFIG.EMAIL_ASSUNTO)}</title></head>
    <body style="margin:0;background:#000">
    <div style="background:#000;padding:32px 16px;font-family:Arial,Helvetica,sans-serif">
      <div style="max-width:560px;margin:0 auto;background:#000;color:#f4f4f5">
        ${logo}
        <h1 style="font-size:22px;color:#5ADFFF;margin:0 0 16px">Obrigado por querer ir mais longe connosco, ${esc(primeiroNome)}!</h1>
        <p style="margin:0 0 16px;line-height:1.6">
          A sua inscrição na campanha THINK BLACK, TRAVEL FAR- Black Friday 2026 foi recebida com sucesso.
        </p>
        <p style="margin:0 0 16px;line-height:1.6">
          A nossa equipa já está a preparar-se para o ajudar a transformar os seus planos de viagem numa oportunidade a não perder. Entraremos em contacto consigo muito em breve, com uma proposta pensada à sua medida.
        </p>
        <p style="margin:0 0 16px;line-height:1.6">
          Até lá, comece a imaginar o seu próximo destino. O mundo está à sua espera!
        </p>
        ${blocoMarketing}
        <p style="margin:24px 0 0;line-height:1.6">Boas viagens,<br><strong>Equipa DS Travel Póvoa de Varzim</strong></p>
        ${CONFIG.LOGO_RODAPE_URL ? `<img src="${esc(CONFIG.LOGO_RODAPE_URL)}" alt="${esc(CONFIG.MARCA)}" width="140" style="display:block;margin:20px 0 0;max-width:100%">` : ''}
        <hr style="border:0;border-top:1px solid #27272a;margin:28px 0 12px">
        <p style="font-size:11px;color:#71717a;line-height:1.5;margin:0">
          Recebeu este e-mail porque submeteu o formulário da campanha ${esc(CONFIG.CAMPANHA)}.
          Referência: ${esc(r.id)}. ${esc(CONFIG.RESPONSAVEL_TRATAMENTO)}.
        </p>
      </div>
    </div>
    </body></html>`;

    const texto =
      `Obrigado por querer ir mais longe connosco, ${primeiroNome}!🌍\n\n` +
      `Agradecemos a sua participação na campanha ${CONFIG.CAMPANHA}. ` +
      `A nossa equipa vai entrar em contacto consigo muito em breve.\n\n` +
      (r.marketing
        ? `Confirmamos que aceitou receber por e-mail as nossas novas campanhas e promoções. ` +
          `Para cancelar a subscrição: ${linkCancelar_(r.id)}\n` +
          `Ou escreva para ${CONFIG.CONTACTO_PRIVACIDADE}.\n\n`
        : '') +
      `Equipa DS Travel Póvoa de Varzim`;

    const opts = {
      htmlBody: html,
      name: CONFIG.EMAIL_NOME_REMETENTE,
    };
    if (CONFIG.EMAIL_NO_REPLY) opts.noReply = true;
    else if (CONFIG.EMAIL_REPLY_TO) opts.replyTo = CONFIG.EMAIL_REPLY_TO;   // só define Reply-To se estiver preenchido

    if (CONFIG.EMAIL_FROM) {
      // Envia a partir do e-mail do domínio (tem de existir como alias em "Enviar email como")
      const aliases = GmailApp.getAliases();
      if (aliases.indexOf(CONFIG.EMAIL_FROM) === -1) {
        throw new Error('O endereço ' + CONFIG.EMAIL_FROM +
          ' não está configurado em Gmail > Definições > Contas > Enviar email como. Alias disponíveis: ' +
          aliases.join(', '));
      }
      opts.from = CONFIG.EMAIL_FROM;
      GmailApp.sendEmail(r.email, assuntoFinal, texto, opts);
    } else {
      MailApp.sendEmail(Object.assign({ to: r.email, subject: assuntoFinal, body: texto }, opts));
    }
    return { ok: true };
  } catch (err) {
    console.error('Erro ao enviar e-mail: ' + err);
    return { ok: false, erro: String(err && err.message || err).slice(0, 200) };
  }
}

// ---------------- CANCELAR SUBSCRIÇÃO ----------------
// Cada e-mail com consentimento de marketing leva um link único e assinado (HMAC).
// O link abre uma página com um botão "Confirmar"; só ao clicar é que o Sheet muda para "Não".
// (A confirmação por botão evita que antivírus/leitores de e-mail cancelem sozinhos ao abrir o link.)

function segredo_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('UNSUB_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('UNSUB_SECRET', s); }
  return s;
}

function tokenCancelar_(id) {
  const sig = Utilities.computeHmacSha256Signature(id, segredo_());
  return sig.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('').slice(0, 32);
}

function idValido_(id) { return /^BF-[0-9A-Z-]{10,40}$/.test(String(id || '')); }

function linkCancelar_(id) {
  const base = CONFIG.URL_WEBAPP || ScriptApp.getService().getUrl();
  return base + '?acao=cancelar&id=' + encodeURIComponent(id) + '&t=' + tokenCancelar_(id);
}

function paginaCancelar_(id, t) {
  const valido = idValido_(id) && /^[0-9a-f]{32}$/.test(String(t || '')) && t === tokenCancelar_(id);
  const estilo = 'body{margin:0;background:#000;color:#f4f4f5;font-family:Arial,Helvetica,sans-serif;' +
    'display:flex;min-height:100vh;align-items:center;justify-content:center;text-align:center}' +
    '.c{max-width:480px;padding:24px}h1{color:#5ADFFF;font-size:22px}' +
    'button{background:#5ADFFF;color:#000;border:0;border-radius:8px;padding:14px 26px;font-size:16px;font-weight:bold;cursor:pointer}' +
    'button:disabled{opacity:.6}p{line-height:1.6;color:#d4d4d8}';
  let corpo;
  if (!valido) {
    corpo = '<h1>Ligação inválida</h1><p>Esta ligação não é válida ou expirou. Para cancelar a subscrição, escreva para ' +
      escHtml_(CONFIG.CONTACTO_PRIVACIDADE) + '.</p>';
  } else {
    corpo = '<h1>Cancelar subscrição</h1>' +
      '<p id="msg">Deixará de receber comunicações sobre novas campanhas e promoções de ' + escHtml_(CONFIG.MARCA) + '.</p>' +
      '<button id="b" onclick="go()">Confirmar cancelamento</button>' +
      '<script>function go(){var b=document.getElementById("b");b.disabled=true;b.textContent="A processar…";' +
      'google.script.run.withSuccessHandler(function(){document.getElementById("msg").textContent=' +
      '"Subscrição cancelada. Não voltará a receber estas comunicações.";b.style.display="none";document.querySelector("h1").textContent="Cancelado";})' +
      '.withFailureHandler(function(){b.disabled=false;b.textContent="Confirmar cancelamento";' +
      'document.getElementById("msg").textContent="Não foi possível cancelar. Tente novamente ou escreva para ' +
      escHtml_(CONFIG.CONTACTO_PRIVACIDADE) + '.";})' +
      '.cancelarSubscricao("' + id + '","' + t + '");}</script>';
  }
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>' + estilo + '</style></head><body><div class="c">' + corpo + '</div></body></html>')
    .setTitle('Cancelar subscrição');
}

// Chamada pelo botão da página (google.script.run). Tem de ser pública (sem "_" no fim).
function cancelarSubscricao(id, t) {
  if (!idValido_(id) || t !== tokenCancelar_(id)) throw new Error('Ligação inválida');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = abrirSheet_();
    const ultima = sh.getLastRow();
    if (ultima < 2) throw new Error('Registo não encontrado');
    const ids = sh.getRange(2, coluna_(sh, 'ID'), ultima - 1, 1).getValues();
    let linha = -1;
    for (let i = 0; i < ids.length; i++) { if (ids[i][0] === id) { linha = i + 2; break; } }
    if (linha < 0) throw new Error('Registo não encontrado');

    const colMkt = coluna_(sh, 'Aceita novas campanhas');
    const colCanc = coluna_(sh, 'Cancelou novas campanhas em');
    if (sh.getRange(linha, colMkt).getValue() === 'Sim') {
      sh.getRange(linha, colMkt).setValue('Não');
      sh.getRange(linha, colCanc).setValue(
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
    }
  } finally {
    lock.releaseLock();
  }
  try { atualizarExcel_(); } catch (err) { console.error('Excel: ' + err); }
  return 'ok';
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

// ---------------- ONDE ESTÃO OS FICHEIROS? ----------------
// Corra mostrarLinks() e abra o registo de execução: mostra o Sheet e as pastas que o script
// está realmente a usar (útil se o setup correu mais de uma vez e há pastas duplicadas).
function mostrarLinks() {
  const p = PropertiesService.getScriptProperties();
  Logger.log('Conta que corre o script: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Sheet: ' + SpreadsheetApp.openById(p.getProperty('SHEET_ID')).getUrl());
  Logger.log('Pasta principal: ' + DriveApp.getFolderById(p.getProperty('FOLDER_ID')).getUrl());
  Logger.log('Pasta dos PDFs: ' + DriveApp.getFolderById(p.getProperty('PDF_FOLDER_ID')).getUrl());
}

// ---------------- TESTE DO E-MAIL ----------------
// 1) Escreva o seu e-mail em DESTINO. 2) Escolha testarEmail no seletor e carregue em Executar.
// 3) Se pedir autorização, aceite. 4) Abra o Registo de execução: mostra a quota e o erro (se houver).
// Ao contrário do envio normal, aqui o erro NÃO é escondido.
function testarEmail() {
  const DESTINO = 'jose.morgado@teoriacuriosa.pt';
  Logger.log('Conta que envia: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Quota de e-mails restante hoje: ' + MailApp.getRemainingDailyQuota());
  MailApp.sendEmail({
    to: DESTINO,
    subject: 'Teste de e-mail — ' + CONFIG.CAMPANHA,
    body: 'Se está a ler isto, o envio de e-mail pelo Apps Script funciona.',
    name: CONFIG.EMAIL_NOME_REMETENTE,
  });
  Logger.log('E-mail enviado para ' + DESTINO + '. Veja também a pasta Spam.');
}

// Envia o e-mail de agradecimento COMPLETO (o mesmo do site) sem passar pelo formulário.
// Mude DESTINO, escolha testarEmailCompleto e carregue em Executar. Mostra o resultado no registo.
// Mude COM_MARKETING para true/false para ver as duas versões do e-mail.
function testarEmailCompleto() {
  const DESTINO = 'morgadojosepedro@gmail.com';
  const COM_MARKETING = true;
  const r = { id: 'BF-TESTE-0000-0000', nome: 'Teste Silva', primeiroNome: 'Teste', email: DESTINO, marketing: COM_MARKETING };
  Logger.log('Quota restante hoje: ' + MailApp.getRemainingDailyQuota());
  Logger.log('Resultado: ' + JSON.stringify(enviarEmail_(r)));
  Logger.log('Reply-To em uso: ' + (CONFIG.EMAIL_NO_REPLY ? '(no-reply)' : CONFIG.EMAIL_REPLY_TO));
}

// Envia 3 e-mails diferentes para descobrir o que está a ser bloqueado:
//   A = texto simples com o assunto real | B = HTML simples | C = e-mail completo (o do site)
// Corra, espere 1 minuto e veja quais chegam. Depois mude DESTINO para o Gmail pessoal e repita.
function testarVariantes() {
  const DESTINO = 'jose.morgado@teoriacuriosa.pt';
  const nome = CONFIG.EMAIL_NOME_REMETENTE;

  MailApp.sendEmail({ to: DESTINO, subject: 'A - ' + CONFIG.EMAIL_ASSUNTO,
    body: 'Teste A: texto simples com o assunto real.', name: nome });

  MailApp.sendEmail({ to: DESTINO, subject: 'B - teste HTML simples',
    body: 'Teste B', htmlBody: '<p>Teste B: HTML simples.</p>', name: nome });

  const r = { id: 'BF-TESTE-0000-0000', nome: 'Teste Silva', primeiroNome: 'Teste', email: DESTINO, marketing: true };
  Logger.log('C (completo): ' + JSON.stringify(enviarEmail_(r)));
  Logger.log('Enviados A, B e C para ' + DESTINO + '. Quota restante: ' + MailApp.getRemainingDailyQuota());
}

// ---------------- DIAGNÓSTICO DO E-MAIL ----------------
// Corra diagnosticoEmail() no editor e abra o Registo de execução. Mostra:
//  1) o que ficou escrito em "E-mail enviado" nas últimas 5 respostas (Sim, ou o motivo do erro);
//  2) a conta que envia e a quota de e-mails que sobra hoje;
//  3) se a mensagem aparece na pasta "Enviados" dessa conta;
//  4) se algum servidor devolveu o e-mail (mensagens "Mail Delivery Subsystem").
// Não envia nada e não precisa de nova implementação: corre direto no editor.
function diagnosticoEmail() {
  Logger.log('Conta que envia: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Quota de e-mails restante hoje: ' + MailApp.getRemainingDailyQuota());
  Logger.log('EMAIL_FROM: ' + (CONFIG.EMAIL_FROM || '(vazio: sai da conta acima)'));

  const sh = abrirSheet_();
  const ultima = sh.getLastRow();
  if (ultima >= 2) {
    const cId = coluna_(sh, 'ID'), cMail = coluna_(sh, 'E-mail'), cEnv = coluna_(sh, 'E-mail enviado');
    for (let i = Math.max(2, ultima - 4); i <= ultima; i++) {
      Logger.log('Linha ' + i + ': ' + sh.getRange(i, cId).getValue() + ' | ' + sh.getRange(i, cMail).getValue() +
                 ' | E-mail enviado = ' + sh.getRange(i, cEnv).getValue());
    }
  }

  const enviados = GmailApp.search('in:sent subject:"Obrigado pela sua inscrição" newer_than:3d', 0, 5);
  Logger.log('Na pasta Enviados (últimos 3 dias): ' + enviados.length + ' conversa(s)');
  enviados.forEach(t => Logger.log('  - para ' + t.getMessages()[0].getTo() + ' em ' + t.getLastMessageDate()));

  const devolvidos = GmailApp.search('(from:mailer-daemon OR from:postmaster) newer_than:3d', 0, 5);
  Logger.log('Mensagens de devolução (últimos 3 dias): ' + devolvidos.length);
  devolvidos.forEach(t => Logger.log('  - ' + t.getFirstMessageSubject() + ' :: ' +
    t.getMessages()[0].getPlainBody().replace(/\s+/g, ' ').slice(0, 300)));
}

// ---------------- TESTE DE ENTREGA (Hotmail/Outlook) ----------------
// Envia 3 e-mails para o mesmo endereço, para descobrir O QUE faz o Outlook mandá-los para o lixo:
//   1 = e-mail completo COM o bloco de marketing (igual ao do site, com o link do script.google.com)
//   2 = e-mail completo SEM marketing (sem esse link)
//   3 = texto simples, sem imagens nem links (o mais "limpo" possível)
// Mude DESTINO se for preciso, corra testarHotmail, espere 1 a 2 minutos e veja em que pasta
// chegou cada um (os assuntos começam por "1 -", "2 -" e "3 -"). Não precisa de nova implementação.
function testarHotmail() {
  const DESTINO = 'jose.marta.morgado99@outlook.com';
  const base = { id: 'BF-TESTE-0000-0000', nome: 'Maria Silva', primeiroNome: 'Maria', email: DESTINO };

  Logger.log('1 (completo, com marketing): ' + JSON.stringify(
    enviarEmail_(Object.assign({}, base, { marketing: true }), '1 - ' + CONFIG.EMAIL_ASSUNTO)));
  Logger.log('2 (completo, sem marketing): ' + JSON.stringify(
    enviarEmail_(Object.assign({}, base, { marketing: false }), '2 - ' + CONFIG.EMAIL_ASSUNTO)));

  MailApp.sendEmail({
    to: DESTINO,
    subject: '3 - ' + CONFIG.EMAIL_ASSUNTO,
    body: 'Obrigado pela sua inscrição na campanha ' + CONFIG.CAMPANHA + '. ' +
          'A nossa equipa vai entrar em contacto consigo muito em breve.\n\nEquipa DS Travel Póvoa de Varzim',
    name: CONFIG.EMAIL_NOME_REMETENTE,
  });
  Logger.log('3 (texto simples): enviado. Quota restante hoje: ' + MailApp.getRemainingDailyQuota());
}

// ---------------- TESTE MANUAL ----------------
// Corra testarSubmissao() no editor para simular uma resposta (use o seu e-mail).
// Envia exatamente os mesmos campos que o script.js da landing page.
function testarSubmissao() {
  const res = doPost({ postData: { contents: JSON.stringify({
    nome: 'Teste Silva', primeiroNome: 'Teste', apelido: 'Silva',
    email: Session.getActiveUser().getEmail(), whatsapp: '+351 912 345 678',
    pessoas: '2',
    privacidade: true, marketing: true, consentVersion: CONFIG.CONSENT_VERSION,
    ip: '0.0.0.0', pagina: 'editor',
    userAgent: '[origem: teste | manual | editor | -] Teste manual',
    origem: '[origem: teste | manual | editor | -]',
  }) } });
  Logger.log(res.getContent());
}

// Mostra, sem criar nada, como a etiqueta de origem é lida. Corra no editor e veja o registo de execução.
function testarLerOrigem() {
  [
    { userAgent: '[origem: instagram | bio | blackfriday2026 | -] Mozilla/5.0', origem: '[origem: instagram | bio | blackfriday2026 | -]' },
    { userAgent: '[origem: direto | - | - | -] Mozilla/5.0' },
    { userAgent: 'Mozilla/5.0 (sem etiqueta)' },
    { userAgent: '[origem: =HYPERLINK("x") | - | - | -] Mozilla/5.0', origem: '[origem: Falso | - | - | -]' },
  ].forEach(d => Logger.log(JSON.stringify(lerOrigem_(d))));
}

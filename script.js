/* =========================================================
   Black Friday — Think Black Travel Far
   Envio do formulário para o Google Apps Script (Code.gs)
   ========================================================= */

// 1) Cole aqui o URL da implementação da Web App do Apps Script (termina em /exec)
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwtI1qCOS-kCEtYfICa9cRTXgpL8hRrxzmPF53B9aUY_hEM3aHScyDOs35zd8vckDhUMA/exec';

// 2) Versão do texto dos consentimentos. Se alterar o texto das checkboxes,
//    suba a versão AQUI e no Code.gs (para o comprovativo refletir o texto certo).
const CONSENT_VERSION = 'BF2026-v1';

const form      = document.getElementById('bf-form');
const btn       = document.getElementById('submit-btn');
const statusEl  = document.getElementById('form-status');
const successEl = document.getElementById('success');

const REQUIRED = ['nome', 'apelido', 'email', 'whatsapp', 'pessoas'];
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
const PHONE_RE = /^\+?[0-9\s().-]{9,20}$/;

// Busca o IP público do utilizador (melhor esforço — se falhar, envia vazio)
async function getPublicIP() {
  try {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 3000);
    const r = await fetch('https://api.ipify.org?format=json', { signal: ctrl.signal });
    clearTimeout(t);
    if (!r.ok) return '';
    const j = await r.json();
    return j.ip || '';
  } catch (_) {
    return '';
  }
}

// ---------------------------------------------------------
// Origem do lead (UTM / referrer)
// É enviada à frente do "User Agent", que o Apps Script já grava no Sheet,
// por isso não é preciso alterar o Code.gs.
// Formato: [origem: fonte | meio | campanha | conteúdo]  ("-" = não indicado)
// ---------------------------------------------------------
const ORIGEM_KEY = 'bf_origem';
const ORIGEM_RE  = /^\[origem: [a-z0-9._ |-]{1,160}\]$/;

// Só letras, números, ponto, hífen e underscore: nada que parta a etiqueta
// ou que o Sheet interprete como fórmula, mesmo que alguém forje o link.
function limparOrigem(v) {
  return String(v || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 30);
}

function etiquetaOrigem(fonte, meio, campanha, conteudo) {
  return `[origem: ${[fonte, meio, campanha, conteudo].map((x) => x || '-').join(' | ')}]`;
}

// Sites que costumam aparecer como referrer → nome curto
const REFERRERS = [
  [/^mail\.google\.com$/, 'gmail'],
  [/(^|\.)instagram\.com$/, 'instagram'],
  [/(^|\.)(facebook\.com|fb\.com|fb\.me)$/, 'facebook'],
  [/(^|\.)(whatsapp\.com|wa\.me)$/, 'whatsapp'],
  [/(^|\.)(t\.co|twitter\.com|x\.com)$/, 'x'],
  [/(^|\.)(linkedin\.com|lnkd\.in)$/, 'linkedin'],
  [/(^|\.)tiktok\.com$/, 'tiktok'],
  [/(^|\.)(youtube\.com|youtu\.be)$/, 'youtube'],
  [/(^|\.)google\./, 'google'],
  [/(^|\.)bing\.com$/, 'bing'],
];

// 1.ª escolha: parâmetros do link (?utm_source=...&utm_medium=...&utm_campaign=...&utm_content=...)
function origemDeUtm() {
  const q = new URLSearchParams(location.search);
  const fonte = limparOrigem(q.get('utm_source') || q.get('ref'));
  if (!fonte) return '';
  return etiquetaOrigem(
    fonte,
    limparOrigem(q.get('utm_medium')),
    limparOrigem(q.get('utm_campaign')),
    limparOrigem(q.get('utm_content'))
  );
}

// 2.ª escolha: o site de onde a pessoa veio (ignora o próprio site)
function origemDeReferrer() {
  let host = '';
  try { host = new URL(document.referrer).hostname.replace(/^www\./, ''); } catch (_) { return ''; }
  if (!host || host === location.hostname.replace(/^www\./, '')) return '';
  const conhecido = REFERRERS.find(([re]) => re.test(host));
  return etiquetaOrigem(conhecido ? conhecido[1] : limparOrigem(host), 'referral');
}

// A origem fica guardada nesta visita (separador), para não se perder se a pessoa
// for ler a Política de Privacidade e voltar ao formulário.
function lerSessao() {
  try {
    const v = sessionStorage.getItem(ORIGEM_KEY) || '';
    return ORIGEM_RE.test(v) ? v : '';
  } catch (_) { return ''; }
}
function guardarSessao(v) {
  try { sessionStorage.setItem(ORIGEM_KEY, v); } catch (_) { /* sem armazenamento: segue sem guardar */ }
}

// Prioridade: link com UTM > origem já guardada nesta visita > referrer > direto
function detetarOrigem() {
  const utm = origemDeUtm();
  if (utm) { guardarSessao(utm); return utm; }
  const guardada = lerSessao();
  if (guardada) return guardada;
  const o = origemDeReferrer() || etiquetaOrigem('direto');
  guardarSessao(o);
  return o;
}

// Calculada logo ao abrir a página
const ORIGEM = detetarOrigem();

function setError(name, msg) {
  const p = form.querySelector(`.error[data-for="${name}"]`);
  const field = form.querySelector(`#${name}`)?.closest('.field');
  if (p) p.textContent = msg || '';
  if (field) field.classList.toggle('has-error', !!msg);
}

function validate() {
  let ok = true;
  let firstInvalid = null;

  REQUIRED.forEach((name) => {
    const el = form.elements[name];
    const v = el.value.trim();
    let msg = '';
    if (!v) msg = 'Campo obrigatório.';
    else if (name === 'email' && !EMAIL_RE.test(v)) msg = 'Introduza um e-mail válido.';
    else if (name === 'whatsapp' && !PHONE_RE.test(v)) msg = 'Introduza um número válido (ex.: +351 912 345 678).';
    else if (name === 'pessoas' && !(Number.isInteger(+v) && +v >= 1 && +v <= 99)) msg = 'Indique um número entre 1 e 99.';
    setError(name, msg);
    if (msg) { ok = false; firstInvalid ??= el; }
  });

  const priv = form.elements.privacidade;
  if (!priv.checked) {
    setError('privacidade', 'Para continuar tem de aceitar a Política de Privacidade.');
    ok = false; firstInvalid ??= priv;
  } else {
    setError('privacidade', '');
  }

  if (firstInvalid) firstInvalid.focus();
  return ok;
}

// Limpa o erro de um campo assim que o utilizador corrige
form.addEventListener('input', (e) => {
  if (e.target.name) setError(e.target.name, '');
});
form.addEventListener('change', (e) => {
  if (e.target.name === 'privacidade' && e.target.checked) setError('privacidade', '');
});

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  statusEl.textContent = '';
  statusEl.classList.remove('is-error');

  // Honeypot preenchido → bot. Finge sucesso e não envia nada.
  if (form.elements.website.value) {
    form.hidden = true; successEl.hidden = false;
    return;
  }

  if (!validate()) return;

  // Aviso claro se o endereço do Apps Script ainda não foi configurado
  if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(APPS_SCRIPT_URL)) {
    statusEl.textContent = 'Configuração em falta: o endereço do Apps Script (APPS_SCRIPT_URL) não está correto no script.js.';
    statusEl.classList.add('is-error');
    console.error('APPS_SCRIPT_URL inválido:', APPS_SCRIPT_URL);
    return;
  }

  btn.disabled = true;
  btn.classList.add('is-loading');
  btn.querySelector('.btn__label').textContent = 'A enviar…';

  const f = form.elements;
  const payload = {
    nome:       `${f.nome.value.trim()} ${f.apelido.value.trim()}`,
    primeiroNome: f.nome.value.trim(),
    apelido:    f.apelido.value.trim(),
    email:      f.email.value.trim().toLowerCase(),
    whatsapp:   f.whatsapp.value.trim(),
    pessoas:    f.pessoas.value.trim(),
    privacidade: f.privacidade.checked,
    marketing:   f.marketing.checked,
    consentVersion:   CONSENT_VERSION,   // o texto oficial de cada versão está no Code.gs
    ip:        await getPublicIP(),
    userAgent:  `${ORIGEM} ${navigator.userAgent}`,   // etiqueta de origem + navegador (campo que o Sheet já guarda)
    origem:     ORIGEM,
    pagina:     location.href,
    timestampCliente: new Date().toISOString(),
  };

  try {
    // text/plain evita o preflight CORS — o Apps Script lê o JSON do body na mesma
    const res = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
    });
    const text = await res.text();
    let data;
    try { data = JSON.parse(text); }
    catch (_) { throw new Error('Resposta inesperada do Apps Script (verifique se o acesso está em "Qualquer pessoa"): ' + text.slice(0, 120)); }
    if (!data.ok) throw new Error(data.error || 'Erro desconhecido');
    // Só há sucesso se o servidor devolver o ID do registo (o doGet de teste não devolve)
    if (!data.id) throw new Error('Resposta sem ID de registo: o pedido não foi processado pelo doPost. Resposta: ' + text.slice(0, 120));

    document.getElementById('success-email').textContent = payload.email;
    form.hidden = true;
    successEl.hidden = false;
    successEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
  } catch (err) {
    console.error(err);
    statusEl.textContent = 'Não foi possível enviar o formulário. Verifique a ligação e tente novamente.';
    statusEl.classList.add('is-error');
  } finally {
    btn.disabled = false;
    btn.classList.remove('is-loading');
    btn.querySelector('.btn__label').textContent = 'Enviar';
  }
});
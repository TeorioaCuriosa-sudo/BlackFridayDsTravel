/* =========================================================
   Black Friday — Think Black Travel Far
   Envio do formulário para o Google Apps Script (Code.gs)
   ========================================================= */

// 1) Cole aqui o URL da implementação da Web App do Apps Script (termina em /exec)
const APPS_SCRIPT_URL = 'https://teorioacuriosa-sudo.github.io/BlackFridayDsTravel/';

// 2) Versão do texto dos consentimentos. Se alterar o texto das checkboxes,
//    suba a versão AQUI e no Code.gs (para o comprovativo refletir o texto certo).
const CONSENT_VERSION = 'BF2026-v1';

const form      = document.getElementById('bf-form');
const btn       = document.getElementById('submit-btn');
const statusEl  = document.getElementById('form-status');
const successEl = document.getElementById('success');

const REQUIRED = ['nome', 'email', 'whatsapp', 'destino', 'pessoas', 'orcamento', 'datas'];
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

  btn.disabled = true;
  btn.classList.add('is-loading');
  btn.querySelector('.btn__label').textContent = 'A enviar…';

  const f = form.elements;
  const payload = {
    nome:       f.nome.value.trim(),
    email:      f.email.value.trim().toLowerCase(),
    whatsapp:   f.whatsapp.value.trim(),
    destino:    f.destino.value.trim(),
    pessoas:    f.pessoas.value.trim(),
    orcamento:  f.orcamento.value.trim(),
    datas:      f.datas.value.trim(),
    info:       f.info.value.trim(),
    privacidade: f.privacidade.checked,
    marketing:   f.marketing.checked,
    consentVersion:   CONSENT_VERSION,   // o texto oficial de cada versão está no Code.gs
    ip:        await getPublicIP(),
    userAgent:  navigator.userAgent,
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
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'Erro desconhecido');

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

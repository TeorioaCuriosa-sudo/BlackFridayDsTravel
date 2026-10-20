# Landing Page Black Friday — Think Black Travel Far

## Ficheiros

| Ficheiro | Para quê |
|---|---|
| `index.html`, `styles.css`, `script.js`, `logo.png` | Landing page (alojar no site / Webflow / Netlify / etc.) |
| `apps-script/Code.gs` | Backend no Google: PDF, Sheet, Excel, e-mail |
| `apps-script/appsscript.json` | Manifesto do Apps Script (fuso horário, permissões, Drive API) |

## Como funciona

```
Formulário → Apps Script (Web App)
               ├─ 1. PDF de consentimento  → Drive / Comprovativos de Consentimento
               ├─ 2. Nova linha             → Google Sheet "Respostas Black Friday 2026"
               ├─ 3. Cópia Excel (.xlsx)    → Drive (atualizada a cada resposta)
               └─ 4. E-mail de agradecimento → cliente (com ou sem frase de novas campanhas)
```

## Configuração (≈10 min, uma vez)

1. Com a conta Google da empresa (a dona do Drive e a que envia os e-mails), abra **script.google.com** e crie um **Novo projeto**.
2. Em **Definições do projeto** (⚙️), ative **"Mostrar o ficheiro de manifesto appsscript.json no editor"**.
3. Cole o conteúdo de `Code.gs` e de `appsscript.json` nos ficheiros com o mesmo nome.
4. No topo de `Code.gs`, preencha o `CONFIG`: `EMAIL_REPLY_TO`, `RESPONSAVEL_TRATAMENTO`, `CONTACTO_PRIVACIDADE` e (opcional) `LOGO_URL`.
5. Selecione a função **`setup`** e clique em **Executar**. Autorize as permissões. Esta função cria a pasta no Drive, a subpasta dos PDFs e o Sheet.
6. (Opcional) Execute **`testarSubmissao`**. Deve receber o e-mail e ver o PDF, a linha no Sheet e o `.xlsx` no Drive.
7. **Implementar → Nova implementação → Tipo: Aplicação Web**
   - Executar como: **Eu**
   - Quem tem acesso: **Qualquer pessoa**
   - Copie o URL (termina em `/exec`).
8. Em `script.js`, cole esse URL em `APPS_SCRIPT_URL`.
9. Em `index.html`, troque `POLITICA_PRIVACIDADE_URL` pelo link real e o texto do disclaimer.
10. Publique os 4 ficheiros da landing page.

> Se alterar o `Code.gs` depois, faça **Implementar → Gerir implementações → Editar → Nova versão**. Assim mantém o mesmo URL.

## Alterar o texto dos consentimentos

O texto que fica no PDF é o do `Code.gs` (`CONSENT_TEXTS`), não o do browser, para não poder ser adulterado. Se mudar o texto das checkboxes:
1. Altere-o em `index.html`.
2. Em `Code.gs`, crie uma nova entrada (`'BF2026-v2': {...}`) e **mantenha a antiga**.
3. Suba `CONSENT_VERSION` para `BF2026-v2` em `Code.gs` **e** `script.js`.

## Notas

- **IP:** é obtido no browser via `api.ipify.org` (o Apps Script não tem acesso ao IP). Se for bloqueado, por exemplo por um adblocker, o PDF fica com "Não disponível". A recolha do IP tem de estar mencionada na Política de Privacidade.
- **Limites de e-mail do Google:** ~100 e-mails/dia numa conta Gmail gratuita e ~1.500/dia em Google Workspace.
- **No-reply:** `EMAIL_NO_REPLY: true` só funciona em contas Workspace. Por defeito usa `Reply-To` para um e-mail monitorizado, que é a opção recomendada.
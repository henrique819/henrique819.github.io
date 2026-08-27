/**
 * Recebe as mensagens do formulário do portfólio e grava numa planilha do Google.
 *
 * A planilha já está apontada abaixo (arquivo "portfolio").
 *
 * COMO INSTALAR
 * 1. Abra a planilha: Extensões > Apps Script.
 * 2. Apague o conteúdo do Code.gs e cole este arquivo inteiro.
 * 3. Salve (ícone de disquete).
 * 4. Clique em Implantar > Nova implantação.
 *    - Tipo: App da Web
 *    - Executar como: Eu (sua conta)
 *    - Quem pode acessar: Qualquer pessoa      <-- precisa ser este
 * 5. Autorize o script quando o Google pedir (a tela de "app não verificado" é
 *    esperada: Avançado > Acessar o projeto).
 * 6. Copie a URL que termina em /exec e cole em ENDPOINT_PLANILHA no index.html.
 *
 * Se depois você alterar este código, precisa clicar em Implantar >
 * Gerenciar implantações > editar (lápis) > Versão: Nova versão > Implantar.
 * Só assim a URL passa a rodar o código novo.
 */

const ID_PLANILHA = '1jXZfYMus1KaQ2zMo-ai2eCGLFheQZ-87DpSqDLI1ym8';
const NOME_ABA    = 'Mensagens';

// Deixe como null para não receber aviso por e-mail a cada mensagem.
const AVISAR_EMAIL = 'henrique819@gmail.com';

function doPost(e) {
  const trava = LockService.getScriptLock();
  trava.waitLock(20000); // evita duas gravações na mesma linha

  try {
    const p = (e && e.parameter) || {};
    const aba = obterAba_();

    aba.appendRow([
      new Date(),
      String(p.nome || '').slice(0, 200),
      String(p.email || '').slice(0, 200),
      String(p.mensagem || '').slice(0, 5000),
      String(p.origem || '')
    ]);

    if (AVISAR_EMAIL) {
      MailApp.sendEmail({
        to: AVISAR_EMAIL,
        subject: 'Nova mensagem no portfólio — ' + (p.nome || 'sem nome'),
        replyTo: p.email || AVISAR_EMAIL,
        body: 'Nome: ' + (p.nome || '-') +
              '\nE-mail: ' + (p.email || '-') +
              '\n\n' + (p.mensagem || '-')
      });
    }

    return resposta_({ ok: true });
  } catch (erro) {
    return resposta_({ ok: false, erro: String(erro) });
  } finally {
    trava.releaseLock();
  }
}

// Abre a página no navegador só para conferir que a implantação está no ar.
function doGet() {
  return resposta_({ ok: true, servico: 'formulario-portfolio' });
}

function obterAba_() {
  const planilha = SpreadsheetApp.openById(ID_PLANILHA);
  let aba = planilha.getSheetByName(NOME_ABA);

  if (!aba) {
    aba = planilha.insertSheet(NOME_ABA);
    aba.appendRow(['Data', 'Nome', 'E-mail', 'Mensagem', 'Origem']);
    aba.getRange('A1:E1').setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

function resposta_(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

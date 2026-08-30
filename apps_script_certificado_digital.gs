// ============================================================================
// APPS SCRIPT — CERTIFICADO DIGITAL (grava CERTIFICADOS/EMAIL/MENSAGEM)
// ============================================================================
// Container-bound: usa SpreadsheetApp.getActiveSpreadsheet(), ou seja, sempre
// escreve na planilha à qual ESTE projeto de script está anexado (Extensões >
// Apps Script daquela planilha específica). Por isso é PORTÁVEL como está —
// não tem SPREADSHEET_ID nem FOLDER_ID pra trocar, só precisa ser colado num
// projeto novo, bound à planilha certa, e implantado.
//
// Recebe um POST com até 3 blocos (certificados/emails/mensagens), cada um
// {cabecalho: [...], linhas: [[...], ...]}, e reescreve a aba correspondente
// (CERTIFICADOS/EMAIL/MENSAGEM) por inteiro (clearContents + setValues) - cria
// a aba se ela ainda não existir. Chamado por _cert_salvar_dados() em
// *_Gestor_Fiscal.py (seção CERTIFICADO DIGITAL), via requests.post(
// APPS_SCRIPT_URL, json=payload).
//
// Este arquivo é só cópia/backup - colar aqui NÃO atualiza o que está
// publicado. Passos pra publicar uma cópia NOVA (uma por escritório - cada
// planilha precisa da sua própria implantação, JAMAIS reaproveitar a mesma
// URL de outro escritório, senão os dados de Certificado Digital de um
// escritório vão parar na planilha do outro):
//   1. Abra a planilha do escritório (a mesma do GOOGLE_SHEET_URL/SHEET_ID
//      daquele *_Gestor_Fiscal.py) > Extensões > Apps Script.
//   2. Apague o conteúdo padrão do editor e cole este código.
//   3. Implantar > Nova implantação > tipo "Aplicativo da Web".
//      Executar como: Eu. Quem pode acessar: Qualquer pessoa.
//   4. Copiar a URL gerada e colar em APPS_SCRIPT_URL, no topo da seção
//      "CERTIFICADO DIGITAL" do *_Gestor_Fiscal.py daquele escritório.
//   5. Garantir que a planilha tenha (ou deixe o script criar sozinho na
//      primeira gravação) as abas CERTIFICADOS, EMAIL e MENSAGEM.
//
// Se um dia mudar a lógica: editar aqui pra manter o histórico, colar em CADA
// um dos projetos publicados (um por escritório - não tem implantação
// compartilhada aqui, diferente do apps_script_pdf_situacao_fiscal.gs), salvar,
// e "Implantar > Gerenciar implantações > editar (lápis) > Nova versão >
// Implantar" em cada um pra manter as URLs já configuradas.
// ============================================================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.certificados) _salvarAba(ss, 'CERTIFICADOS', data.certificados.cabecalho, data.certificados.linhas);
    if (data.emails)       _salvarAba(ss, 'EMAIL',        data.emails.cabecalho,        data.emails.linhas);
    if (data.mensagens)    _salvarAba(ss, 'MENSAGEM',     data.mensagens.cabecalho,     data.mensagens.linhas);

    return ContentService
      .createTextOutput(JSON.stringify({status: 'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function _salvarAba(ss, nome, cabecalho, linhas) {
  var ws = ss.getSheetByName(nome);
  if (!ws) ws = ss.insertSheet(nome);
  ws.clearContents();
  var todas = [cabecalho].concat(linhas);
  if (todas.length > 0 && todas[0].length > 0) {
    ws.getRange(1, 1, todas.length, todas[0].length).setValues(todas);
  }
}

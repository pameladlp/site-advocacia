/**
 * Cadastro de Hóspedes — recebe os cadastros da página /hospede/
 *
 * Para cada cadastro:
 *  1. guarda as fotos do documento numa subpasta da pasta "Cadastro Hóspedes" no Drive
 *  2. acrescenta uma linha na planilha "Cadastro Hóspedes - Planilha"
 *  3. envia um e-mail para EMAIL_DESTINO com os dados e as fotos em anexo
 *
 * Implantação: Implantar > Nova implantação > Tipo "App da Web"
 *   Executar como: Eu  ·  Quem pode acessar: Qualquer pessoa
 */

const EMAIL_DESTINO = 'pameladlp@gmail.com';
const PASTA_ID = '1ymRK5Rme5XtmVYMgY0zaWHE0rxn-JaE3';                    // pasta "Cadastro Hóspedes"
const PLANILHA_ID = '1k1BoELUlf3Gx2LZ18iWl2tZBA4AECo_CnRdqr7OKIn4';      // "Cadastro Hóspedes - Planilha"
const MAX_ARQUIVO_MB = 8;

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d._honey) return resposta({ ok: true });                         // robô: ignora em silêncio
    if (!d.nome || !d.doc || !d.frente) return resposta({ ok: false, erro: 'dados incompletos' });

    const agora = new Date();
    const carimbo = Utilities.formatDate(agora, 'America/Recife', 'yyyy-MM-dd HH:mm');

    // 1) Fotos no Drive
    const pasta = DriveApp.getFolderById(PASTA_ID)
      .createFolder(`${(d.checkin || '').split('/').reverse().join('-')} ${d.nome}`.trim());
    const anexos = [];
    const links = {};
    [['frente', 'Documento - frente'], ['verso', 'Documento - verso']].forEach(([k, rotulo]) => {
      const f = d[k];
      if (!f || !f.data) return;
      const bytes = Utilities.base64Decode(f.data);
      if (bytes.length > MAX_ARQUIVO_MB * 1024 * 1024) return;
      const ext = (f.name.match(/\.\w+$/) || ['.jpg'])[0];
      const blob = Utilities.newBlob(bytes, f.type || 'image/jpeg', `${rotulo}${ext}`);
      const arq = pasta.createFile(blob);
      links[k] = arq.getUrl();
      anexos.push(blob);
    });

    // 2) Planilha
    const acomp = (d.acomp || [])
      .map((a, i) => `${i + 1}) ${a.nome}${a.doc ? ' — ' + a.doc : ''}${a.menor ? ' (menor)' : ''}`)
      .join('\n') || 'Nenhum';
    // O apóstrofo faz a planilha tratar tudo como texto (ex.: "+1 555..." não vira fórmula)
    const txt = v => (v === undefined || v === null || v === '') ? '' : "'" + String(v);
    SpreadsheetApp.openById(PLANILHA_ID).getSheets()[0].appendRow([
      carimbo, d.checkin, d.checkout, d.horario, d.qtd, d.nome, d.doc, d.nasc, d.tel,
      acomp, d.obs, d.idioma, links.frente || '', links.verso || '', d.lgpd ? 'Sim' : 'Não'
    ].map(txt));

    // 3) E-mail
    const linha = (k, v) => v ? `<tr><td style="padding:6px 12px;color:#666">${k}</td><td style="padding:6px 12px"><b>${esc(v)}</b></td></tr>` : '';
    const html = `
      <h2 style="font-family:Arial">Novo cadastro de hóspede</h2>
      <table style="font-family:Arial;font-size:14px;border-collapse:collapse">
        ${linha('Check-in', d.checkin)}${linha('Check-out', d.checkout)}${linha('Chegada prevista', d.horario)}
        ${linha('Nº de hóspedes', d.qtd)}${linha('Nome completo', d.nome)}${linha('CPF / Passaporte', d.doc)}
        ${linha('Nascimento', d.nasc)}${linha('WhatsApp', d.tel)}${linha('Acompanhantes', acomp)}
        ${linha('Observações', d.obs)}${linha('Idioma', d.idioma)}${linha('Recebido em', carimbo)}
      </table>
      <p style="font-family:Arial;font-size:13px">📁 <a href="${pasta.getUrl()}">Abrir a pasta deste hóspede no Drive</a></p>`;
    const flag = d.idioma && d.idioma !== 'Português' ? ` [${d.idioma}]` : '';
    MailApp.sendEmail({
      to: EMAIL_DESTINO,
      subject: `Cadastro de hóspede${flag}: ${d.nome} — check-in ${d.checkin || ''}`,
      htmlBody: html,
      attachments: anexos,
      name: 'Cadastro de Hóspedes'
    });

    return resposta({ ok: true });
  } catch (err) {
    console.error(err);
    return resposta({ ok: false, erro: String(err) });
  }
}

// Abrir o link do app no navegador mostra esta mensagem (serve para testar se está no ar)
function doGet() {
  return resposta({ ok: true, servico: 'Cadastro de Hóspedes' });
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function esc(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/\n/g, '<br>');
}

/** Rode esta função uma vez pelo editor para autorizar o acesso ao Drive, Planilhas e Gmail. */
function autorizar() {
  DriveApp.getFolderById(PASTA_ID).getName();
  SpreadsheetApp.openById(PLANILHA_ID).getName();
  MailApp.getRemainingDailyQuota();
}

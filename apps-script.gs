/* ============================================================
   ART PEL EMBALAGENS — Apps Script
   Recebe POST das LPs (Industrial + Loja), grava na aba certa
   da planilha e encaminha pro CRM Dros quando qualificado.
   ============================================================ */

// ─── CONFIG (edite estes valores antes de publicar) ──────────
const SHEET_ID = '12CwT97kHc6zfuriXeErO6kZ-kPo6BAYpAOy_J6_3SmA'; // ENTRADA DE LEADS ARTPEL EMBALAGENS - DROS AGENCIA
const SHEET_INDUSTRIAL = 'INDUSTRIAS';
const SHEET_LOJA = 'LOJAS E DISTRIBUIDORES';
const CRM_WEBHOOK_URL = 'https://drosagencia.com.br/crm/api/webhooks/sheets/art-pel-embalagens-ltda';
const CRM_WEBHOOK_SECRET = ''; // se o CRM Dros exigir header X-Webhook-Secret, cole aqui
const TAG_INDUSTRIAL = 'LP-INDUSTRIAL';
const TAG_LOJA = 'LP-LOJAS';

// Qualificacao: TODOS os leads vao pro CRM, exceto lojistas/atacadistas
// (redirecionados pra LP2). Nao ha mais desqualificacao por quantidade.
const VALORES_DESQUALIFICADOS = [];

// ─── ENTRADA PRINCIPAL ───────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const origem = String(body.origem || 'LP Industrial Art Pel');
    // Roteia aba+tag POR ORIGEM (mais confiavel que segmento). LP Lojas
    // sempre vai pra aba LOJAS E DISTRIBUIDORES (tag LP-LOJAS), LP Industrial
    // sempre pra INDUSTRIAS (tag LP-INDUSTRIAL). Assim opcoes como
    // 'Supermercado / atacarejo' ou 'Outro' que nao batem no regex de
    // segmento vao pra aba certa da LP que originou.
    const isLoja = /loja|atacadista/i.test(origem);
    const sheetName = isLoja ? SHEET_LOJA : SHEET_INDUSTRIAL;
    const tag = isLoja ? TAG_LOJA : TAG_INDUSTRIAL;

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('Aba nao encontrada: ' + sheetName);

    // Header automatico se aba vazia
    if (sheet.getLastRow() === 0) sheet.appendRow(getHeader());

    // Sem qualificacao: TODOS os leads vao pro CRM (lojista vira tag LP-LOJAS,
    // industria vira LP-INDUSTRIAL — cliente filtra no CRM depois se quiser)
    let statusCRM = 'Não enviado';
    let crmResponse = '';
    const crmResult = enviarParaCRM(body, tag, origem);
    statusCRM = crmResult.ok ? 'Enviado ✓ (' + crmResult.status + ')' : 'Erro (' + crmResult.status + ')';
    crmResponse = String(crmResult.response || '').substring(0, 500);

    // Grava linha
    sheet.appendRow(buildRow(body, statusCRM, crmResponse));

    return json({ ok: true, status: statusCRM });
  } catch (err) {
    console.error('[Apps Script]', err.message);
    return json({ ok: false, error: err.message });
  }
}

// GET simples pra testar que o script ta publicado
function doGet(e) {
  return json({ ok: true, msg: 'Art Pel Apps Script ready', at: new Date().toISOString() });
}

// ─── HELPERS ─────────────────────────────────────────────────

function getHeader() {
  return [
    'Timestamp', 'Nome', 'Empresa', 'WhatsApp', 'Cidade/Estado',
    'Segmento', 'Quantidade de caixas', 'Frequência', 'Personalização',
    'UTM Source', 'UTM Medium', 'UTM Campaign', 'UTM Content', 'UTM Term',
    'GCLID', 'FBCLID', 'FBC', 'FBP', 'Referrer', 'Landing Page', 'User Agent', 'Device',
    'Event ID', 'Qualificado', 'Status CRM', 'CRM Response'
  ];
}

function buildRow(body, statusCRM, crmResponse) {
  return [
    body.timestamp || new Date().toISOString(),
    body.nome || '',
    body.empresa || '',
    body.whatsapp || '',
    body.cidade || '',
    body.segmento || '',
    body.valor_medio || '',  // agora eh "Quantidade de caixas"
    body.frequencia || '',
    body.personalizacao || '',
    body.utm_source || '',
    body.utm_medium || '',
    body.utm_campaign || '',
    body.utm_content || '',
    body.utm_term || '',
    body.gclid || '',
    body.fbclid || '',
    body.fbc || '',
    body.fbp || '',
    body.referrer || '',
    body.landing_page || '',
    body.user_agent || '',
    body.device || '',
    body.fb_event_id || '',
    body.qualificado ? 'Sim' : 'Não',
    statusCRM,
    crmResponse
  ];
}

function enviarParaCRM(body, tag, origem) {
  try {
    const cidadeParts = String(body.cidade || '').split(/[\/,-]/).map(s => s.trim()).filter(Boolean);
    const cidadeName = cidadeParts[0] || '';
    const uf = cidadeParts[1] || '';

    const utms = [body.utm_source, body.utm_medium, body.utm_campaign].filter(Boolean).join(' | ');
    const sourceDetail = utms || body.referrer || 'direto';

    const obs = [
      'Origem: ' + origem,
      'Empresa: ' + (body.empresa || '-'),
      'Segmento: ' + (body.segmento || '-'),
      'Quantidade de caixas por compra: ' + (body.valor_medio || '-'),
      'Frequência de compra: ' + (body.frequencia || '-'),
      'Precisa personalização: ' + (body.personalizacao || '-'),
      'Cidade: ' + (body.cidade || '-'),
      '---',
      'Tracking:',
      '- Referrer: ' + (body.referrer || '-'),
      '- Landing: ' + (body.landing_page || '-'),
      '- UTMs: ' + (utms || '-'),
      body.gclid ? '- GCLID: ' + body.gclid : '',
      body.fbclid ? '- FBCLID: ' + body.fbclid : '',
      body.fbc ? '- FBC cookie: ' + body.fbc : '',
      body.fbp ? '- FBP cookie: ' + body.fbp : '',
      '- Device: ' + (body.device || '-'),
      '- User Agent: ' + (body.user_agent || '-').substring(0, 150),
    ].filter(Boolean).join('\n');

    const payload = {
      name: body.nome || '',
      phone: body.whatsapp || '',
      city: cidadeName,
      state: uf,
      source: origem,
      source_detail: sourceDetail,
      tags: [tag],
      trabalha_anuncio: true,
      observations: obs,
      fbc: body.fbc || null,
      fbp: body.fbp || null,
      ctwa_clid: body.fbclid || null,
      utm_source: body.utm_source || null,
      utm_medium: body.utm_medium || null,
      utm_campaign: body.utm_campaign || null,
      utm_content: body.utm_content || null,
      utm_term: body.utm_term || null,
      fb_event_id: body.fb_event_id || null,
      user_agent: body.user_agent || null,
      landing_page: body.landing_page || null,
      referrer: body.referrer || null
    };

    const headers = {};
    if (CRM_WEBHOOK_SECRET) headers['X-Webhook-Secret'] = CRM_WEBHOOK_SECRET;

    const res = UrlFetchApp.fetch(CRM_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: headers,
      muteHttpExceptions: true,
      followRedirects: true,
    });

    const status = res.getResponseCode();
    return {
      ok: status >= 200 && status < 400,
      status: status,
      response: res.getContentText()
    };
  } catch (err) {
    return { ok: false, status: 0, response: err.message };
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ─── BACKFILL (opcional): reprocessa linhas onde Status CRM = "Não enviado" ─
// Roda manual pelo menu Apps Script: Executar > backfillCRM
function backfillCRM() {
  const sheets = [SHEET_INDUSTRIAL, SHEET_LOJA];
  sheets.forEach(name => {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const data = sheet.getDataRange().getValues();
    const header = data[0];
    const iStatus = header.indexOf('Status CRM');
    const iResp = header.indexOf('CRM Response');
    if (iStatus < 0) return;
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const status = String(row[iStatus] || '');
      if (!/N[aã]o enviado|Erro/i.test(status)) continue;
      const body = rowToBody(row, header);
      const tag = name === SHEET_LOJA ? TAG_LOJA : TAG_INDUSTRIAL;
      const result = enviarParaCRM(body, tag, name === SHEET_LOJA ? 'LP Loja Art Pel' : 'LP Industrial Art Pel');
      sheet.getRange(r + 1, iStatus + 1).setValue(result.ok ? 'Enviado ✓ backfill (' + result.status + ')' : 'Erro backfill (' + result.status + ')');
      if (iResp >= 0) sheet.getRange(r + 1, iResp + 1).setValue(String(result.response || '').substring(0, 500));
      Utilities.sleep(500);
    }
  });
}

function rowToBody(row, header) {
  const map = {};
  const keys = {
    'Timestamp': 'timestamp', 'Nome': 'nome', 'Empresa': 'empresa',
    'WhatsApp': 'whatsapp', 'Cidade/Estado': 'cidade',
    'Segmento': 'segmento', 'Quantidade de caixas': 'valor_medio', 'Frequência': 'frequencia',
    'Personalização': 'personalizacao',
    'UTM Source': 'utm_source', 'UTM Medium': 'utm_medium', 'UTM Campaign': 'utm_campaign',
    'UTM Content': 'utm_content', 'UTM Term': 'utm_term',
    'GCLID': 'gclid', 'FBCLID': 'fbclid', 'FBC': 'fbc', 'FBP': 'fbp',
    'Referrer': 'referrer', 'Landing Page': 'landing_page', 'User Agent': 'user_agent',
    'Device': 'device', 'Event ID': 'fb_event_id'
  };
  header.forEach((h, i) => { if (keys[h]) map[keys[h]] = row[i]; });
  return map;
}

/**
 * Painel "Sua Primeira Publi Paga" — backend (Google Apps Script)
 * Recebe: eventos do site (visitas/cliques/FAQ) + webhook do Kiwify (vendas/reembolsos)
 * Guarda em abas e serve os dados pro painel.
 *
 * Como publicar:
 *  Extensões → Apps Script → cole isto → Implantar → Nova implantação
 *  → tipo "App da Web" → Executar como: Eu → Quem tem acesso: "Qualquer pessoa"
 *  → Implantar → copie a URL que termina em /exec
 */

// ---------- RECEBER DADOS (POST) ----------
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var body = {};
    try { body = JSON.parse(e.postData.contents); } catch (x) { body = e.parameter || {}; }

    if (body.tipo === 'site') {
      registrarTrafego(ss, body);
    } else {
      registrarVenda(ss, body); // qualquer outra coisa = webhook do Kiwify
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, erro: String(err) });
  }
}

function pegarAba(ss, nome, cabecalho) {
  var sh = ss.getSheetByName(nome);
  if (!sh) { sh = ss.insertSheet(nome); sh.appendRow(cabecalho); }
  return sh;
}

// Eventos do site: visita, chegou_oferta, checkout, faq
function registrarTrafego(ss, b) {
  var sh = pegarAba(ss, 'trafego', ['data', 'evento', 'rotulo', 'origem', 'visitante', 'pais']);
  if (sh.getRange(1, 6).getValue() !== 'pais') sh.getRange(1, 6).setValue('pais'); // garante a coluna país
  sh.appendRow([new Date(), b.evento || '', b.rotulo || '', b.origem || '', b.visitante || '', b.pais || '']);
}

// Webhook do Kiwify (guardamos os campos mais comuns + o JSON cru pra não perder nada)
function registrarVenda(ss, b) {
  var sh = pegarAba(ss, 'vendas',
    ['data', 'status', 'nome', 'email', 'metodo', 'valor', 'produto', 'id', 'json_cru']);

  var cust   = b.Customer || b.customer || {};
  var order  = b.Order || b.order || {};
  var status = b.order_status || b.status || order.status || b.webhook_event_type || '';
  var nome   = cust.full_name || cust.name || b.customer_name || '';
  var email  = cust.email || b.customer_email || '';
  var metodo = b.payment_method || order.payment_method || '';
  var produto = (b.Product && b.Product.product_name) || b.product_name || '';
  var id     = b.order_id || order.id || b.id || '';

  var valor = b.charge_amount || b.order_amount || order.charge_amount ||
              (b.Commissions && b.Commissions.charge_amount) || b.commissioned_value || 0;
  valor = Number(valor) || 0;
  if (valor > 1000) valor = valor / 100; // Kiwify às vezes manda em centavos

  sh.appendRow([new Date(), status, nome, email, metodo, valor, produto, id, JSON.stringify(b)]);
}

// ---------- SERVIR DADOS PRO PAINEL (GET) ----------
// Suporta JSONP (?callback=fn) pra o painel ler sem problema de CORS.
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = {
    vendas:  lerAba(ss, 'vendas'),
    trafego: lerAba(ss, 'trafego')
  };
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + JSON.stringify(dados) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json(dados);
}

function lerAba(ss, nome) {
  var sh = ss.getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getDataRange().getValues();
  var head = vals.shift();
  return vals.map(function (row) {
    var o = {};
    head.forEach(function (h, i) { o[h] = row[i]; });
    return o;
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

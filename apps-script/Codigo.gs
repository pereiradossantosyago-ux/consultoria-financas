/**
 * Consultoria PJ & Finanças — agenda e base de dados do site.
 * Roda na conta Google do Yago (Google Apps Script, implantado como app da Web).
 *
 * - GET  ?action=slots   -> horários livres dos próximos dias (lê o Google Agenda)
 * - POST {type:'book'}   -> cria o evento com Google Meet, convida o cliente e grava na planilha
 * - POST {type:'event'}  -> registra um evento anônimo do site (visita, clique, simulador...)
 *
 * Primeira vez: execute a função setup() no editor para criar a planilha.
 */

const CFG = {
  tz: 'America/Sao_Paulo',
  offset: '-03:00',
  duracaoMin: 30,
  diasAFrente: 21,          // janela de agendamento
  antecedenciaHoras: 18,    // não aceita horários muito em cima
  // 0 = domingo ... 6 = sábado
  horarios: {
    1: ['12:00', '12:30'], 2: ['12:00', '12:30'], 3: ['12:00', '12:30'],
    4: ['12:00', '12:30'], 5: ['12:00', '12:30'],
    6: ['09:00', '10:00', '11:00', '14:00', '15:00']
  },
  avisarPorEmail: true
};

const ABA_CONTATOS = 'Contatos';
const ABA_EVENTOS = 'Eventos';
const ABA_PAINEL = 'Painel';
const CAB_CONTATOS = ['Recebido em', 'Nome', 'WhatsApp', 'E-mail', 'Perfil', 'Empresa', 'Assunto', 'Contexto',
  'Data da conversa', 'Horário', 'Formato', 'Link do Meet', 'Origem', 'Status'];
const CAB_EVENTOS = ['Data/hora', 'Sessão', 'Tipo', 'Página', 'Detalhe', 'Origem', 'Dispositivo'];

/* ------------------------------------------------------------------ setup */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const id = props.getProperty('SHEET_ID');
  if (id) { ss = SpreadsheetApp.openById(id); }
  else {
    ss = SpreadsheetApp.create('Consultoria PJ & Finanças — Base de dados do site');
    props.setProperty('SHEET_ID', ss.getId());
  }
  ss.setSpreadsheetTimeZone(CFG.tz);
  const contatos = obterAba_(ss, ABA_CONTATOS, CAB_CONTATOS);
  const eventos = obterAba_(ss, ABA_EVENTOS, CAB_EVENTOS);
  contatos.setColumnWidths(1, CAB_CONTATOS.length, 140);
  eventos.setColumnWidths(1, CAB_EVENTOS.length, 150);
  // fórmulas escritas no padrão en_US e convertidas de volta para o idioma da planilha
  const local = ss.getSpreadsheetLocale();
  ss.setSpreadsheetLocale('en_US');
  montarPainel_(ss);
  SpreadsheetApp.flush();
  ss.setSpreadsheetLocale(local || 'pt_BR');
  const padrao = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1');
  if (padrao && ss.getSheets().length > 1) ss.deleteSheet(padrao);
  ss.setActiveSheet(ss.getSheetByName(ABA_PAINEL));
  ss.moveActiveSheet(1);
  Logger.log('Planilha: ' + ss.getUrl());
  return ss.getUrl();
}

function obterAba_(ss, nome, cab) {
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold')
    .setBackground('#13584A').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  return sh;
}

function montarPainel_(ss) {
  const sh = ss.getSheetByName(ABA_PAINEL) || ss.insertSheet(ABA_PAINEL);
  sh.clear();
  const E = "'" + ABA_EVENTOS + "'", C = "'" + ABA_CONTATOS + "'";
  sh.getRange('A1').setValue('Painel do site — Consultoria PJ & Finanças').setFontSize(14).setFontWeight('bold').setFontColor('#13584A');
  sh.getRange('A2').setValue('Atualiza sozinho a cada visita ou agendamento. Visitas contadas sem cookies e sem identificar quem visitou.').setFontColor('#5B6B6A');
  const kpis = [
    ['Visitas (total)', `=IFERROR(COUNTUNIQUE(FILTER(${E}!B2:B,${E}!C2:C="page_view")),0)`],
    ['Visitas nos últimos 7 dias', `=IFERROR(COUNTUNIQUE(FILTER(${E}!B2:B,${E}!C2:C="page_view",${E}!A2:A>=NOW()-7)),0)`],
    ['Visualizações de página', `=COUNTIF(${E}!C2:C,"page_view")`],
    ['Cliques no WhatsApp', `=COUNTIF(${E}!C2:C,"whatsapp")`],
    ['Diagnósticos concluídos', `=COUNTIF(${E}!C2:C,"diagnostico")`],
    ['Nota média do diagnóstico', `=IFERROR(ROUND(AVERAGE(FILTER(VALUE(${E}!E2:E),${E}!C2:C="diagnostico")),0),"-")`],
    ['Simuladores usados', `=COUNTIF(${E}!C2:C,"simulador")`],
    ['Pedidos da planilha gratuita', `=COUNTIF(${E}!C2:C,"material")`],
    ['Agendamentos', `=COUNTA(${C}!A2:A)`],
    ['Conversão: agendamentos ÷ visitas', `=IFERROR(B12/B4,0)`]
  ];
  sh.getRange(4, 1, kpis.length, 1).setValues(kpis.map(k => [k[0]]));
  sh.getRange(4, 2, kpis.length, 1).setFormulas(kpis.map(k => [k[1]]));
  sh.getRange('B13').setNumberFormat('0.0%');
  sh.getRange(4, 1, kpis.length, 2).setBorder(true, true, true, true, true, true, '#D5DEDC', null);
  sh.getRange(4, 1, kpis.length, 1).setFontWeight('bold');
  const blocos = [
    ['D', 'Páginas mais vistas', `=IFERROR(QUERY(${E}!A:G,"select D, count(B) where C='page_view' group by D order by count(B) desc label D 'Página', count(B) 'Visualizações'",1),"Sem dados ainda")`],
    ['G', 'Origem das visitas', `=IFERROR(QUERY(${E}!A:G,"select F, count(B) where C='page_view' group by F order by count(B) desc label F 'Origem', count(B) 'Visualizações'",1),"Sem dados ainda")`],
    ['J', 'Dispositivo', `=IFERROR(QUERY(${E}!A:G,"select G, count(B) where C='page_view' group by G label G 'Dispositivo', count(B) 'Visualizações'",1),"Sem dados ainda")`],
    ['M', 'Visitas por dia (30 dias)', `=IFERROR(QUERY(${E}!A:G,"select toDate(A), count(B) where C='page_view' group by toDate(A) order by toDate(A) desc limit 30 label toDate(A) 'Dia', count(B) 'Visualizações'",1),"Sem dados ainda")`]
  ];
  blocos.forEach(b => {
    sh.getRange(b[0] + '3').setValue(b[1]).setFontWeight('bold').setFontColor('#13584A');
    sh.getRange(b[0] + '4').setFormula(b[2]);
  });
  sh.getRange('A16').setValue('Agendamentos por perfil').setFontWeight('bold').setFontColor('#13584A');
  sh.getRange('A17').setFormula(`=IFERROR(QUERY(${C}!A:N,"select E, count(A) where A is not null group by E label E 'Perfil', count(A) 'Agendamentos'",1),"Sem dados ainda")`);
  sh.getRange('A24').setValue('Agendamentos por assunto').setFontWeight('bold').setFontColor('#13584A');
  sh.getRange('A25').setFormula(`=IFERROR(QUERY(${C}!A:N,"select G, count(A) where A is not null group by G order by count(A) desc label G 'Assunto', count(A) 'Agendamentos'",1),"Sem dados ainda")`);
  sh.setColumnWidth(1, 260);
  sh.setFrozenRows(2);
}

function planilha_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Execute setup() primeiro.');
  return SpreadsheetApp.openById(id);
}

/* ---------------------------------------------------------------- web app */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'slots') return json_({ ok: true, slots: horariosLivres_() });
  return json_({ ok: true, service: 'consultoria-financas' });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json_({ ok: false, error: 'JSON inválido' }); }
  try {
    if (body.type === 'event') { registrarEvento_(body); return json_({ ok: true }); }
    if (body.type === 'book') return json_(agendar_(body));
    return json_({ ok: false, error: 'Tipo desconhecido' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Erro interno. Tente pelo WhatsApp.' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------ horários */
function horariosLivres_() {
  const agora = new Date();
  const fim = new Date(agora.getTime() + (CFG.diasAFrente + 1) * 864e5);
  const ocupados = CalendarApp.getDefaultCalendar().getEvents(agora, fim)
    .filter(ev => !ev.isAllDayEvent())
    .map(ev => [ev.getStartTime().getTime(), ev.getEndTime().getTime()]);
  const limite = agora.getTime() + CFG.antecedenciaHoras * 36e5;
  const dias = [];
  for (let k = 0; k <= CFG.diasAFrente; k++) {
    const d = new Date(agora.getTime() + k * 864e5);
    const iso = Utilities.formatDate(d, CFG.tz, 'yyyy-MM-dd');
    const dow = Number(Utilities.formatDate(d, CFG.tz, 'u')) % 7; // u: 1=seg..7=dom
    const lista = (CFG.horarios[dow] || []).filter(h => {
      const ini = inicio_(iso, h).getTime(), fimSlot = ini + CFG.duracaoMin * 6e4;
      if (ini < limite) return false;
      return !ocupados.some(o => ini < o[1] && fimSlot > o[0]);
    });
    if (lista.length) dias.push({ d: iso, t: lista });
  }
  return dias;
}

function inicio_(iso, hora) { return new Date(iso + 'T' + hora + ':00' + CFG.offset); }

/* ------------------------------------------------------------ agendar */
function agendar_(b) {
  const req = ['nome', 'whatsapp', 'email', 'perfil', 'assunto', 'data', 'hora', 'formato'];
  for (const k of req) if (!b[k] || String(b[k]).trim() === '') return { ok: false, error: 'Campo obrigatório ausente: ' + k };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(b.email)) return { ok: false, error: 'E-mail inválido.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(b.data) || !/^\d{2}:\d{2}$/.test(b.hora)) return { ok: false, error: 'Data ou horário inválido.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const livres = horariosLivres_();
    const dia = livres.find(x => x.d === b.data);
    if (!dia || dia.t.indexOf(b.hora) < 0) return { ok: false, taken: true, error: 'Esse horário acabou de ser ocupado. Escolha outro.' };
    const ini = inicio_(b.data, b.hora), fim = new Date(ini.getTime() + CFG.duracaoMin * 6e4);
    const fmt = d => Utilities.formatDate(d, CFG.tz, "yyyy-MM-dd'T'HH:mm:ss") + CFG.offset;
    const desc = [
      'Conversa inicial agendada pelo site Consultoria PJ & Finanças.', '',
      'Nome: ' + b.nome, 'WhatsApp: ' + b.whatsapp, 'E-mail: ' + b.email,
      'Perfil: ' + b.perfil + (b.empresa ? ' · ' + b.empresa : ''), 'Assunto: ' + b.assunto,
      b.contexto ? 'Contexto: ' + b.contexto : '', 'Formato preferido: ' + b.formato
    ].filter(Boolean).join('\n');
    const ev = Calendar.Events.insert({
      summary: 'Conversa inicial · ' + b.nome + ' (' + b.assunto + ')',
      description: desc,
      start: { dateTime: fmt(ini), timeZone: CFG.tz },
      end: { dateTime: fmt(fim), timeZone: CFG.tz },
      attendees: [{ email: b.email, displayName: b.nome }],
      conferenceData: { createRequest: { requestId: Utilities.getUuid(), conferenceSolutionKey: { type: 'hangoutsMeet' } } },
      reminders: { useDefault: false, overrides: [{ method: 'email', minutes: 24 * 60 }, { method: 'popup', minutes: 30 }] }
    }, 'primary', { conferenceDataVersion: 1, sendUpdates: 'all' });
    const meet = ev.hangoutLink || '';
    planilha_().getSheetByName(ABA_CONTATOS).appendRow([
      new Date(), b.nome, b.whatsapp, b.email, b.perfil, b.empresa || '', b.assunto, b.contexto || '',
      b.data, b.hora, b.formato, meet, b.origem || '', 'Confirmado'
    ]);
    registrarEvento_({ session: b.session, name: 'agendamento', page: 'contato', detail: b.perfil, ref: b.origem, device: b.device });
    if (CFG.avisarPorEmail) {
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Novo agendamento: ' + b.nome + ' em ' + Utilities.formatDate(ini, CFG.tz, 'dd/MM HH:mm'), desc + '\n\nMeet: ' + meet);
    }
    return { ok: true, meet: meet, quando: Utilities.formatDate(ini, CFG.tz, "dd/MM/yyyy 'às' HH:mm") };
  } finally { lock.releaseLock(); }
}

/* ------------------------------------------------------------ eventos */
function registrarEvento_(b) {
  const permitidos = ['page_view', 'whatsapp', 'simulador', 'diagnostico', 'material', 'agendamento', 'agendamento_inicio'];
  const nome = String(b.name || '');
  if (permitidos.indexOf(nome) < 0) return;
  const corta = (v, n) => String(v || '').slice(0, n);
  planilha_().getSheetByName(ABA_EVENTOS).appendRow([
    new Date(), corta(b.session, 20), nome, corta(b.page, 60), corta(b.detail, 80), corta(b.ref, 60), corta(b.device, 12)
  ]);
}

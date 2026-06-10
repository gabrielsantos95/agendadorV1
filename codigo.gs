// ─────────────────────────────────────────────────────────────
//  Seller Pro · Agendador de Calls
//  Cole este código em script.google.com
//  Deploy → Web App → Execute as: Me | Access: Anyone
// ─────────────────────────────────────────────────────────────

// ── CONFIGURAÇÕES ────────────────────────────────────────────

const CONFIG = {
  // ID do seu Google Calendar (geralmente seu e-mail, ou o ID
  // de um calendário específico — encontre em:
  // Google Calendar → Configurações do calendário → ID do calendário)
  CALENDAR_ID: "primary",

  // Fuso horário
  TIMEZONE: "America/Sao_Paulo",

  // Seu nome (aparece na descrição do evento)
  CONSULTOR_NOME: "GB · Seller Pro",

  // Quantas calls máximo por dia por tipo
  MAX_CALLS_POR_DIA: {
    consultoria: 3,
    mentoria: 3
  },

  // Dias permitidos por tipo (0=Dom, 1=Seg, 2=Ter, 3=Qua, 4=Qui, 5=Sex, 6=Sáb)
  DIAS_PERMITIDOS: {
    consultoria: [1, 3, 4],  // Seg, Qua, Qui
    mentoria:    [2, 4]       // Ter, Qui
  },

  // Slots de horário por tipo
  SLOTS: {
    consultoria: ["09:00", "10:00", "11:00", "13:00", "14:00", "15:00"],
    mentoria:    ["09:30", "11:00", "13:30"]
  },

  // Duração em minutos
  DURACAO: {
    consultoria: 60,
    mentoria:    90
  }
};

// ── ENTRY POINTS ─────────────────────────────────────────────
// Tudo via GET para garantir compatibilidade total com Apps Script Web App.
// doPost mantido como fallback mas o HTML usa GET.

function doGet(e) {
  const action = e.parameter.action || "";

  if (action === "disponibilidade") {
    return handleDisponibilidade(e);
  }

  if (action === "agendar") {
    // Recebe parâmetros via query string
    return handleAgendarGet(e);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: "online", app: "Seller Pro Agendador" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // Fallback — tenta ler postData, mas prefira usar doGet
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw) return jsonResponse({ ok: false, erro: "postData vazio. Use GET." });
    const payload = JSON.parse(raw);
    if (payload.action === "agendar") return handleAgendar(payload);
    return jsonResponse({ ok: false, erro: "Ação desconhecida." });
  } catch (err) {
    return jsonResponse({ ok: false, erro: "Erro no POST: " + err.message });
  }
}

// Wrapper que converte parâmetros GET em payload e chama handleAgendar
function handleAgendarGet(e) {
  const payload = {
    action:        "agendar",
    tipo:          e.parameter.tipo          || "",
    data:          e.parameter.data          || "",
    slot:          e.parameter.slot          || "",
    nomeCliente:   e.parameter.nomeCliente   || "",
    emailCliente:  e.parameter.emailCliente  || ""
  };
  return handleAgendar(payload);
}

// ── DISPONIBILIDADE ──────────────────────────────────────────
// GET ?action=disponibilidade&tipo=consultoria&ano=2026&mes=6

function handleDisponibilidade(e) {
  const tipo = e.parameter.tipo;
  const ano  = parseInt(e.parameter.ano);
  const mes  = parseInt(e.parameter.mes) - 1; // JS month index

  if (!tipo || !CONFIG.DIAS_PERMITIDOS[tipo]) {
    return jsonResponse({ ok: false, erro: "Tipo inválido." });
  }

  const diasPermitidos = CONFIG.DIAS_PERMITIDOS[tipo];
  const maxPorDia      = CONFIG.MAX_CALLS_POR_DIA[tipo];
  const slots          = CONFIG.SLOTS[tipo];
  const duracao        = CONFIG.DURACAO[tipo];
  const cal            = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) ||
                         CalendarApp.getDefaultCalendar();

  const inicioMes = new Date(ano, mes, 1);
  const fimMes    = new Date(ano, mes + 1, 0, 23, 59, 59);
  const eventos   = cal.getEvents(inicioMes, fimMes);

  // Monta mapa: "YYYY-MM-DD HH:MM" → true (horário ocupado)
  const ocupados = {};
  eventos.forEach(ev => {
    const inicio = ev.getStartTime();
    const chave  = formatDateKey(inicio);
    ocupados[chave] = (ocupados[chave] || 0) + 1;
  });

  const hoje   = new Date();
  hoje.setHours(0, 0, 0, 0);
  const resultado = {};

  const totalDias = new Date(ano, mes + 1, 0).getDate();
  for (let d = 1; d <= totalDias; d++) {
    const dt  = new Date(ano, mes, d);
    const dow = dt.getDay();
    if (!diasPermitidos.includes(dow)) continue;
    if (dt < hoje) continue;

    const dateStr = formatDateStr(dt); // "YYYY-MM-DD"
    let callsNoDia = 0;
    const slotsDia = {};

    slots.forEach(slot => {
      const chave = dateStr + " " + slot;
      const livre = !ocupados[chave];
      slotsDia[slot] = livre;
      if (!livre) callsNoDia++;
    });

    resultado[dateStr] = {
      slots: slotsDia,
      cheio: callsNoDia >= maxPorDia
    };
  }

  return jsonResponse({ ok: true, disponibilidade: resultado, tipo, slots, duracao });
}

// ── AGENDAR ──────────────────────────────────────────────────
// POST { action: "agendar", tipo, data, slot, nomeCliente, emailCliente }

function handleAgendar(payload) {
  const { tipo, data, slot, nomeCliente, emailCliente } = payload;

  // Validações
  if (!tipo || !CONFIG.DURACAO[tipo]) {
    return jsonResponse({ ok: false, erro: "Tipo de serviço inválido." });
  }
  if (!data || !slot) {
    return jsonResponse({ ok: false, erro: "Data ou horário ausente." });
  }
  if (!nomeCliente || nomeCliente.trim() === "") {
    return jsonResponse({ ok: false, erro: "Nome do cliente obrigatório." });
  }

  // Verifica se o dia é permitido
  const partes = data.split("-");
  const dtObj  = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));
  const dow    = dtObj.getDay();
  if (!CONFIG.DIAS_PERMITIDOS[tipo].includes(dow)) {
    return jsonResponse({ ok: false, erro: "Dia não disponível para este tipo de serviço." });
  }

  // Monta datetime de início e fim
  const [hh, mm]  = slot.split(":").map(Number);
  const inicio    = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]), hh, mm, 0);
  const durMin    = CONFIG.DURACAO[tipo];
  const fim       = new Date(inicio.getTime() + durMin * 60000);

  // Verifica disponibilidade real no calendário
  const cal    = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) ||
                 CalendarApp.getDefaultCalendar();
  const conflitos = cal.getEvents(inicio, fim);
  if (conflitos.length > 0) {
    return jsonResponse({ ok: false, erro: "Horário já ocupado. Escolha outro slot." });
  }

  // Verifica limite diário
  const inicioDia = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]), 0, 0, 0);
  const fimDia    = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]), 23, 59, 59);
  const eventosNoDia = cal.getEvents(inicioDia, fimDia);
  if (eventosNoDia.length >= CONFIG.MAX_CALLS_POR_DIA[tipo]) {
    return jsonResponse({ ok: false, erro: "Limite de calls neste dia já atingido." });
  }

  // Monta título e descrição
  const tipoLabel = tipo === "consultoria" ? "Consultoria" : "Mentoria ME";
  const titulo    = `${tipoLabel} · ${nomeCliente.trim()}`;
  const descricao = [
    `Tipo: ${tipoLabel}`,
    `Cliente: ${nomeCliente.trim()}`,
    emailCliente ? `E-mail: ${emailCliente.trim()}` : null,
    `Duração: ${durMin} minutos`,
    `Agendado via Seller Pro Scheduler`
  ].filter(Boolean).join("\n");

  // Cria o evento
  const evento = cal.createEvent(titulo, inicio, fim, {
    description: descricao,
    guests: emailCliente ? emailCliente.trim() : undefined,
    sendInvites: !!emailCliente
  });

  return jsonResponse({
    ok: true,
    mensagem: "Agendamento confirmado!",
    evento: {
      id:     evento.getId(),
      titulo: titulo,
      data:   formatDateStr(inicio),
      inicio: slot,
      fim:    formatTime(fim),
      tipo:   tipoLabel
    }
  });
}

// ── HELPERS ──────────────────────────────────────────────────

function jsonResponse(obj) {
  const output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function formatDateStr(d) {
  const y  = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, "0");
  const dy = String(d.getDate()).padStart(2, "0");
  return `${y}-${mo}-${dy}`;
}

function formatDateKey(d) {
  const h  = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  return formatDateStr(d) + " " + h + ":" + mi;
}

function formatTime(d) {
  const h  = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  return h + ":" + mi;
}

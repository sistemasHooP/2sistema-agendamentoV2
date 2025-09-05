// ========================================================
// ARQUIVO: Code.gs (VERSÃO FINAL COM TODAS AS CORREÇÕES)
// ========================================================

const SS_CONFIG = {
  sheets: {
    config: "Configuracoes",
    professionals: "Profissionais",
    appointments: "Agendamentos",
    masterUsers: "Usuarios_Master",
    attendantUsers: "Usuarios_Atendente",
    services: "Servicos",
    importSheet: "Importar_Agendamentos",
    callScreen: "Chamada_Atual",
    clients: "Clientes",
    clientHistory: "HistoricoCliente"
  }
};

// ========================================================
// PONTO DE ENTRADA DA API
// ========================================================

function doGet(e) {
  if (e && e.parameter && e.parameter.noredirect) {
    const template = HtmlService.createTemplateFromFile('index');
    const config = getLoginScreenConfig();
    template.config = JSON.stringify(config);
    
    // Carrega os múltiplos arquivos de script
    template.jsCore = HtmlService.createHtmlOutputFromFile('JavaScript-Core').getContent();
    template.jsMaster = HtmlService.createHtmlOutputFromFile('JavaScript-Master').getContent();
    template.jsAtendente = HtmlService.createHtmlOutputFromFile('JavaScript-Atendente').getContent();
    template.jsProfissional = HtmlService.createHtmlOutputFromFile('JavaScript-Profissional').getContent();
    
    // Remove a linha antiga do 'template.js' se ela existir
    delete template.js;

    return template.evaluate()
      .setTitle('Painel de Agendamento Profissional')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "API de Agendamento está online." }))
    .setMimeType(ContentService.MimeType.JSON);
}


function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const params = body.params || {};

    const functionMap = {
      // Login & Dados Iniciais
      'doLogin': () => doLogin(params.credentials),
      'getInitialData': getInitialData,
      'getProfessionalData': () => getProfessionalData(params.profId),
      'getLatestUpdateTimestamp': getLatestUpdateTimestamp,
      'getLoginScreenConfig': getLoginScreenConfig,
      'getAgendaUpdate': () => getAgendaUpdate(params.profId),
      'getCoreData': getCoreData,
      'getCoreProfessionalData': getCoreProfessionalData,
      'getDashboardStats': getDashboardStats,
      'getRecentAndFutureAppointments': getRecentAndFutureAppointments,

      // Agendamentos
      'getProfessionalWorkdays': () => getProfessionalWorkdays(params.professionalId),
      'getAvailableTimeSlots': () => getAvailableTimeSlots(params.professionalId, params.dateString),
      'scheduleNewAppointment_step1_saveToSheet': () => scheduleNewAppointment_step1_saveToSheet(params.appointmentData),
      'scheduleNewAppointment_step2_backgroundTasks': () => scheduleNewAppointment_step2_backgroundTasks(params.appointmentObject),
      'getPaginatedAppointments': () => getPaginatedAppointments(params.options),
      'getAppointmentDetails': () => getAppointmentDetails(params.appointmentId),
      'updateAppointment': () => updateAppointment(params.appointmentId, params.appointmentData),
      'deleteAppointment': () => deleteAppointment(params.appointmentId),
      
      // Ações do Profissional
      'callClientAndUpdateStatus': () => callClientAndUpdateStatus(params.appointmentId, params.callingProfId),
      'recallClient': () => recallClient(params.appointmentId),
      'updateAppointmentStatus': () => updateAppointmentStatus(params.appointmentId, params.newStatus),
      'toggleAppointmentPriority': () => toggleAppointmentPriority(params.appointmentId),

      // Clientes & Histórico
      'searchClients': () => searchClients(params.searchTerm, params.page, params.pageSize),
      'getClientById': () => getClientById(params.clientId),
      'addOrUpdateClient': () => addOrUpdateClient(params.clientData),
      'getClientNotes': () => getClientNotes(params.clientId),
      'saveClientNote': () => saveClientNote(params.data),
      'updateClientNote': () => updateClientNote(params.noteId, params.newText),
      'deleteClientNote': () => deleteClientNote(params.noteId),

      // Administração (Master)
      'addProfessional': () => addProfessional(params.profData),
      'editProfessional': () => editProfessional(params.profId, params.profData),
      'addAttendant': () => addAttendant(params.attendantData),
      'editAttendant': () => editAttendant(params.originalUsername, params.newData),
      'deleteAttendant': () => deleteAttendant(params.username),
      'addService': () => addService(params.serviceName),
      'editService': () => editService(params.originalName, params.newName),
      'deleteService': () => deleteService(params.serviceName),
      'updateSetting': () => updateSetting(params.key, params.value),
      'importOldAppointments': importOldAppointments,
      'generateAppointmentsPDF': () => generateAppointmentsPDF(params.filters),
      'archiveAndClearAppointments': archiveAndClearAppointments // <-- ADICIONE ESTA LINHA
    };

    if (functionMap[action]) {
      const result = functionMap[action]();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: result }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      throw new Error(`Ação desconhecida: ${action}`);
    }

  } catch (error) {
    Logger.log(`ERRO NA API: ${error.message} \nStack: ${error.stack}`);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: `Erro no servidor: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ========================================================
// FUNÇÕES DE CARGA RÁPIDA (PARA OTIMIZAÇÃO DE LOGIN)
// ========================================================

/**
 * Retorna apenas os dados essenciais para Master/Atendente, SEM agendamentos.
 */
function getCoreData() {
  const professionalsData = getData(SS_CONFIG.sheets.professionals);
  return {
    professionals: professionalsData,
    services: getData(SS_CONFIG.sheets.services),
    config: getData(SS_CONFIG.sheets.config),
    users: {
      master: getData(SS_CONFIG.sheets.masterUsers),
      attendant: getData(SS_CONFIG.sheets.attendantUsers),
      professional: professionalsData,
    },
  };
}

/**
 * Retorna apenas os dados essenciais para Profissional, SEM a agenda.
 */
function getCoreProfessionalData() {
  return {
    professionals: getData(SS_CONFIG.sheets.professionals),
    services: getData(SS_CONFIG.sheets.services),
    config: getData(SS_CONFIG.sheets.config),
  };
}

/**
 * FUNÇÃO DE ATUALIZAÇÃO RÁPIDA: Retorna apenas os dados da agenda do profissional.
 */
function getAgendaUpdate(profId) {
  return getProfessionalAgendaData(profId);
}


// ========================================================
// FUNÇÕES ORIGINAIS DO SISTEMA
// ========================================================

// *** FUNÇÕES DE AGENDAMENTO ***
function getProfessionalWorkdays(professionalId) {
  try {
    const professionals = getData(SS_CONFIG.sheets.professionals);
    const professional = professionals.find(p => p.ID_Profissional === professionalId);
    if (!professional) {
      return { success: true, workdays: { 0: false, 1: false, 2: false, 3: false, 4: false, 5: false, 6: false } };
    }
    const weekDayColumns = ['Domingo', 'Segunda_feira', 'Terca_feira', 'Quarta_feira', 'Quinta_feira', 'Sexta_feira', 'Sabado'];
    const workdays = {};
    weekDayColumns.forEach((day, index) => {
      const worksOnDay = professional[day];
      workdays[index] = (worksOnDay && worksOnDay.toUpperCase() === 'SIM');
    });
    return { success: true, workdays: workdays };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getAvailableTimeSlots(professionalId, dateString) {
    const configData = getData(SS_CONFIG.sheets.config);
    const professional = getData(SS_CONFIG.sheets.professionals).find(p => p.ID_Profissional === professionalId);
    if (!professional || professional.Status !== 'Ativo') return [];

    const slotDurationConfig = configData.find(c => c.Chave === 'DURACAO_PADRAO_SLOT_MINUTOS');
    const slotDuration = slotDurationConfig ? parseInt(slotDurationConfig.Valor, 10) : 30;

    const horarioManha = configData.find(c => c.Chave === 'HORARIO_PADRAO_MANHA')?.Valor || "08:00-12:00";
    const horarioTarde = configData.find(c => c.Chave === 'HORARIO_PADRAO_TARDE')?.Valor || "14:00-18:00";
    
    const date = Utilities.parseDate(dateString, "UTC", "yyyy-MM-dd");
    const dayName = ['Domingo', 'Segunda_feira', 'Terca_feira', 'Quarta_feira', 'Quinta_feira', 'Sexta_feira', 'Sabado'][date.getUTCDay()];
    
    const shift = professional[dayName];
    if (!shift) return [];

    let workingHoursString = "";
    if (shift.includes('M')) { workingHoursString += horarioManha; }
    if (shift.includes('T')) { workingHoursString += (workingHoursString ? "," : "") + horarioTarde; }

    if (!workingHoursString) return [];

    const allAppointments = getAppointments(); 
    const bookedSlots = allAppointments
        .filter(a => a.ID_Profissional === professionalId && a.Data === dateString && ['Confirmado', 'Pendente', 'Chamado'].includes(a.Status))
        .map(a => a.Hora);
    
    const availableSlots = [];
    const timeRanges = workingHoursString.split(',').map(r => r.trim());
    
    timeRanges.forEach(range => {
        const [startStr, endStr] = range.split('-');
        if (!startStr || !endStr) return;
        let currentTime = new Date(`${dateString}T${startStr}:00Z`);
        let endTime = new Date(`${dateString}T${endStr}:00Z`);
        while(currentTime < endTime) {
            const formattedSlot = Utilities.formatDate(currentTime, "UTC", "HH:mm");
            if (!bookedSlots.includes(formattedSlot)) {
                availableSlots.push(formattedSlot);
            }
            currentTime.setUTCMinutes(currentTime.getUTCMinutes() + slotDuration);
        }
    });
    return availableSlots.sort();
}

function generateNextTicketNumber(appointmentDateString) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); 

  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const propKey = `ticket_counter_${appointmentDateString}`;
    const countFromProperties = parseInt(scriptProperties.getProperty(propKey) || 0, 10);
    const appointments = getAppointments(true);
    let maxInSheet = 0;
    appointments.forEach(appt => {
      if (appt.Data === appointmentDateString) {
        const ticketNum = parseInt(appt.Numero_Ficha, 10);
        if (!isNaN(ticketNum) && ticketNum > maxInSheet) {
          maxInSheet = ticketNum;
        }
      }
    });
    const lastNumber = Math.max(countFromProperties, maxInSheet);
    const nextNumber = lastNumber + 1;
    scriptProperties.setProperty(propKey, nextNumber);
    return nextNumber;
  } finally {
    lock.releaseLock();
  }
}

function scheduleNewAppointment_step1_saveToSheet(appointmentData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.appointments);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    const newId = Utilities.getUuid();
    const newTimestamp = new Date();
    const ticketNumber = generateNextTicketNumber(appointmentData.Data);
    appointmentData.Numero_Ficha = ticketNumber;
    appointmentData.Hora = Utilities.formatDate(newTimestamp, Session.getScriptTimeZone(), "HH:mm");
    
    const dataToSave = {...appointmentData};
    dataToSave.ID_Agendamento = newId;
    dataToSave.Data_Agendamento = newTimestamp;
    
    const [year, month, day] = dataToSave.Data.split('-');
    dataToSave.Data = `${day}/${month}/${year}`;
    
    const newRow = headers.map(header => dataToSave[header] || '');
    sheet.appendRow(newRow);
    
    const returnAppointment = {...appointmentData};
    returnAppointment.ID_Agendamento = newId;
    returnAppointment.Data_Agendamento = newTimestamp.toISOString(); 

    const professionals = getData(SS_CONFIG.sheets.professionals);
    const professional = professionals.find(p => p.ID_Profissional === returnAppointment.ID_Profissional);
    const professionalName = professional ? professional.Nome_Completo : 'N/A';

    return { 
      success: true, 
      appointment: returnAppointment,
      ticket: {
        number: ticketNumber,
        clientName: returnAppointment.Nome_Cliente,
        professionalName: professionalName
      }
    };
  } catch(e) {
    Logger.log(`Erro na Etapa 1 do Agendamento: ${e.message} ${e.stack}`);
    return { success: false, message: e.message };
  }
}

function scheduleNewAppointment_step2_backgroundTasks(appointmentObject) {
  try {
    // OTIMIZAÇÃO: Busca os dados aqui, uma única vez.
    const allProfessionals = getData(SS_CONFIG.sheets.professionals);
    const allConfig = getData(SS_CONFIG.sheets.config);
    
    syncAppointmentToCalendar(appointmentObject, allProfessionals, allConfig, false, true); 
    _updateTimestamp();
    return { success: true };
  } catch (e) {
    Logger.log(`Erro na Etapa 2 do Agendamento: ${e.message} ${e.stack}`);
    return { success: false, message: e.message };
  }
}

// *** FUNÇÕES GERAIS E DE UTILIDADE ***
function _updateTimestamp() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.config);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const keyColIndex = headers.indexOf('Chave');
    const valueColIndex = headers.indexOf('Valor');
    const rowIndex = data.findIndex(row => row[keyColIndex] === 'LAST_UPDATE_TIMESTAMP');
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 2, valueColIndex + 1).setValue(new Date().getTime());
    }
  } catch (e) {
    Logger.log("Erro ao atualizar o timestamp: " + e.message);
  }
}

function getLatestUpdateTimestamp() {
  const config = getData(SS_CONFIG.sheets.config, true);
  const timestampConfig = config.find(c => c.Chave === 'LAST_UPDATE_TIMESTAMP');
  return timestampConfig ? timestampConfig.Valor : 0;
}

function getData(sheetName, forceRefresh = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `data_${sheetName}`;
  if (!forceRefresh) {
    const cachedData = cache.get(cacheKey);
    if (cachedData != null) {
      return JSON.parse(cachedData);
    }
  }
  const data = getSheetData(sheetName);
  cache.put(cacheKey, JSON.stringify(data), 7200); 
  return data;
}

function clearAllCache() {
  const cache = CacheService.getScriptCache();
  const keys = [
    `data_${SS_CONFIG.sheets.config}`, `data_${SS_CONFIG.sheets.professionals}`,
    `data_${SS_CONFIG.sheets.masterUsers}`, `data_${SS_CONFIG.sheets.attendantUsers}`,
    `data_${SS_CONFIG.sheets.services}`
  ];
  cache.removeAll(keys);
  Logger.log('Cache do aplicativo limpo com sucesso!');
  try { SpreadsheetApp.getUi().alert('Cache do aplicativo limpo com sucesso!'); } catch (e) {}
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('⚙️ App Agendamento')
      .addItem('Limpar Cache do App', 'clearAllCache')
      .addSeparator() // Adiciona uma linha divisória
      .addItem('Criar Cópia para Novo Cliente', 'createSystemCopy')
      .addToUi();
} 

function findRowIndexById(sheet, idColumnName, idToFind) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idColIndex = headers.indexOf(idColumnName);
  if (idColIndex === -1) return null;

  const idColumnValues = sheet.getRange(2, idColIndex + 1, sheet.getLastRow() - 1, 1).getValues();
  const rowIndexInData = idColumnValues.findIndex(row => row[0] == idToFind);
  return (rowIndexInData !== -1) ? rowIndexInData + 2 : null;
}

function getDashboardStats() {
  const allAppointments = getAppointments();
  const now = new Date();
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const monthStr = todayStr.substring(0, 7);
  const yearStr = todayStr.substring(0, 4);

  const todayCount = allAppointments.filter(a => a.Data === todayStr).length;
  const monthCount = allAppointments.filter(a => a.Data.startsWith(monthStr)).length;
  const yearCount = allAppointments.filter(a => a.Data.startsWith(yearStr)).length;

  const statusCounts = allAppointments.reduce((acc, appt) => {
    const status = appt.Status || "Sem Status";
    acc[status] = (acc[status] || 0) + 1;
    return acc;
  }, {});

  return { today: todayCount, month: monthCount, year: yearCount, status: statusCounts };
}

function getSheetData(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return [];
    const values = sheet.getDataRange().getDisplayValues();
    if (values.length <= 1) return [];
    const headers = values.shift().map(h => h.trim());
    return values.map(row => {
      if (!row[0] || row[0] === "") return null;
      const rowObject = {};
      headers.forEach((header, index) => { rowObject[header] = row[index] || ""; });
      return rowObject;
    }).filter(Boolean);
  } catch (e) {
    Logger.log(`Erro ao ler a aba "${sheetName}": ${e.message}`);
    return [];
  }
}

function getInitialData() {
  const professionalsData = getData(SS_CONFIG.sheets.professionals);
  return {
    professionals: professionalsData,
    services: getData(SS_CONFIG.sheets.services),
    appointments: getRecentAndFutureAppointments(),
    config: getData(SS_CONFIG.sheets.config),
    users: {
      master: getData(SS_CONFIG.sheets.masterUsers),
      attendant: getData(SS_CONFIG.sheets.attendantUsers),
      professional: professionalsData,
    },
    dashboardStats: getDashboardStats(),
    latestUpdate: getLatestUpdateTimestamp()
  };
}

function getProfessionalAgendaData(professionalId) {
    const allProfessionals = getData(SS_CONFIG.sheets.professionals, true);
    const currentProfessional = allProfessionals.find(p => p.ID_Profissional == professionalId);
    if (!currentProfessional) throw new Error("Profissional não encontrado.");

    const agendaGroup = currentProfessional.Grupo_Agenda;
    let professionalIdsInGroup = [professionalId];
    if (agendaGroup && agendaGroup.trim() !== "") {
        professionalIdsInGroup = allProfessionals
            .filter(p => p.Grupo_Agenda === agendaGroup)
            .map(p => p.ID_Profissional);
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = today.toISOString().split('T')[0];
    const allAppointments = getAppointments(true);

    const groupTodaysAppointments = allAppointments.filter(appt =>
        appt.Data === todayStr && professionalIdsInGroup.includes(appt.ID_Profissional)
    );

    const scriptProperties = PropertiesService.getScriptProperties();
    // Chave da propriedade que vai "travar" o próximo da fila
    const nextInLineKey = `nextInLine_group_${agendaGroup || professionalId}`;

    const servingProperties = scriptProperties.getProperties();
    const servingIds = new Set();
    for (const key in servingProperties) {
        if (key.startsWith('serving_')) {
            servingIds.add(key.replace('serving_', ''));
        }
    }

    groupTodaysAppointments.forEach(appt => {
        if (servingIds.has(appt.ID_Agendamento)) {
            appt.Status = 'Chamado';
            appt.ID_Profissional_Chamada = servingProperties['serving_' + appt.ID_Agendamento];
        }
    });

    const currentlyServing = groupTodaysAppointments.find(appt => appt.Status === 'Chamado' && appt.ID_Profissional_Chamada === professionalId) || null;

    const priorityAppointments = [];
    const completedToday = [];
    const normalQueue = [];
    const validStatusForQueue = ['Confirmado', 'Pendente'];

    groupTodaysAppointments.forEach(appt => {
        if (appt.Status === 'Chamado') { return; }
        if (validStatusForQueue.includes(appt.Status)) {
            if (appt.Prioridade === 'SIM') {
                priorityAppointments.push(appt);
            } else {
                normalQueue.push(appt);
            }
        }
        if (appt.ID_Profissional === professionalId && appt.Status === 'Concluído') {
            completedToday.push(appt);
        }
    });

    priorityAppointments.sort((a, b) => parseInt(a.Numero_Ficha, 10) - parseInt(b.Numero_Ficha, 10));
    normalQueue.sort((a, b) => parseInt(a.Numero_Ficha, 10) - parseInt(b.Numero_Ficha, 10));
    completedToday.sort((a, b) => parseInt(a.Numero_Ficha, 10) - parseInt(b.Numero_Ficha, 10));

    // --- LÓGICA DA FILA ATUALIZADA ---
    let nextInLine = null;
    const lockedNextInLineId = scriptProperties.getProperty(nextInLineKey);

    if (lockedNextInLineId) {
        // Se já existe um "próximo" travado, encontre-o na lista de espera.
        nextInLine = [...priorityAppointments, ...normalQueue].find(a => a.ID_Agendamento === lockedNextInLineId);
        
        // Se o cliente travado não for mais encontrado (ex: cancelou), limpa a trava
        if (!nextInLine) {
            scriptProperties.deleteProperty(nextInLineKey);
        }
    }

    // Só executa a lógica de escolha se NÃO houver um "próximo" travado
    if (!nextInLine) {
        const lastCallTypeKey = `lastCallType_prof_${professionalId}`;
        const lastCallType = scriptProperties.getProperty(lastCallTypeKey);
        const hasPriority = priorityAppointments.length > 0;
        const hasNormal = normalQueue.length > 0;

        if (hasPriority && hasNormal) {
            if (lastCallType === 'NORMAL') {
                nextInLine = priorityAppointments[0];
            } else {
                nextInLine = normalQueue[0];
            }
        } else if (hasPriority) {
            nextInLine = priorityAppointments[0];
        } else if (hasNormal) {
            nextInLine = normalQueue[0];
        }

        // Se um novo "próximo" foi escolhido, TRAVA ele na propriedade
        if (nextInLine) {
            scriptProperties.setProperty(nextInLineKey, nextInLine.ID_Agendamento);
        }
    }
    // --- FIM DA LÓGICA DA FILA ---

    const waitingList = [...priorityAppointments, ...normalQueue].sort((a, b) => parseInt(a.Numero_Ficha, 10) - parseInt(b.Numero_Ficha, 10));

    return {
        currentlyServing: currentlyServing,
        nextInLine: nextInLine,
        priorityAppointments: priorityAppointments,
        completedToday: completedToday,
        waitingList: waitingList
    };
}

function getProfessionalData(profId) {
  return {
    professionals: getData(SS_CONFIG.sheets.professionals),
    services: getData(SS_CONFIG.sheets.services),
    config: getData(SS_CONFIG.sheets.config),
    agendaData: getProfessionalAgendaData(profId),
    latestUpdate: getLatestUpdateTimestamp()
  };
}

function getRecentAndFutureAppointments() {
  const allAppointments = getAppointments(true);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(today.getDate() - 30);
  const startDateString = thirtyDaysAgo.toISOString().split('T')[0];
  const filteredAppointments = allAppointments.filter(appt => appt.Data >= startDateString);
  return filteredAppointments;
}

function getAppointments(forceRefresh = false) {
  let allAppointments = getSheetData(SS_CONFIG.sheets.appointments);
  allAppointments = allAppointments.map(appt => {
    const dateStr = appt.Data;
    if (dateStr && /^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) {
      const [day, month, year] = dateStr.split('/');
      appt.Data = `${year}-${month}-${day}`;
    }
    return appt;
  });
  return allAppointments;
}

function callClientAndUpdateStatus(appointmentId, callingProfId) {
    const lock = LockService.getScriptLock();
    const gotLock = lock.tryLock(15000);

    if (!gotLock) {
        return { success: false, message: 'Sistema ocupado. Tente novamente.' };
    }

    try {
        const scriptProperties = PropertiesService.getScriptProperties();

        // --- INÍCIO DA ALTERAÇÃO ---
        // Busca os dados do profissional para identificar o grupo da agenda
        const professional = getData(SS_CONFIG.sheets.professionals, true).find(p => p.ID_Profissional == callingProfId);
        const agendaGroup = professional ? professional.Grupo_Agenda : null;
        // Define a chave da propriedade que "trava" o próximo da fila
        const nextInLineKey = `nextInLine_group_${agendaGroup || callingProfId}`;
        // DESTRAVA o lugar de "próximo da fila", pois este cliente agora será o "em atendimento"
        scriptProperties.deleteProperty(nextInLineKey); 
        // --- FIM DA ALTERAÇÃO ---

        const propKey = `serving_${appointmentId}`;
        if (scriptProperties.getProperty(propKey) != null) {
            return { success: false, message: 'Este cliente já foi chamado. A tela será atualizada.' };
        }

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const apptSheet = ss.getSheetByName(SS_CONFIG.sheets.appointments);
        const headers = apptSheet.getRange(1, 1, 1, apptSheet.getLastColumn()).getValues()[0];
        const rowIndex = findRowIndexById(apptSheet, 'ID_Agendamento', appointmentId);
        if (!rowIndex) return { success: false, message: 'Agendamento não encontrado.' };
        
        const rowData = apptSheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
        const appointmentToCall = headers.reduce((obj, header, i) => {
            obj[header] = rowData[i];
            return obj;
        }, {});

        const statusColIndex = headers.indexOf('Status');
        const statusRange = apptSheet.getRange(rowIndex, statusColIndex + 1);
        const currentStatus = statusRange.getValue();

        if (currentStatus !== 'Pendente' && currentStatus !== 'Confirmado') {
            return { success: false, message: 'Este atendimento não está mais na fila.' };
        }
        
        scriptProperties.setProperty(propKey, callingProfId);
        
        const lastCallTypeKey = `lastCallType_prof_${callingProfId}`;
        const callType = appointmentToCall.Prioridade === 'SIM' ? 'PRIORITY' : 'NORMAL';
        scriptProperties.setProperty(lastCallTypeKey, callType);

        const profCallColIndex = headers.indexOf('ID_Profissional_Chamada');
        statusRange.setValue('Chamado');
        apptSheet.getRange(rowIndex, profCallColIndex + 1).setValue(callingProfId);
        
        SpreadsheetApp.flush();

        const callSheet = ss.getSheetByName(SS_CONFIG.sheets.callScreen);
        if (callSheet) {
            const clientName = appointmentToCall.Nome_Cliente;
            const ticketNumber = appointmentToCall.Numero_Ficha;
            const priorityStatus = appointmentToCall.Prioridade;
            
            const professionals = getData(SS_CONFIG.sheets.professionals, true);
            const callingProfessional = professionals.find(p => p.ID_Profissional == callingProfId);
            const callLocation = callingProfessional ? (callingProfessional.Sala_Atendimento || callingProfessional.Nome_Completo) : 'N/A';
            
            callSheet.insertRowBefore(2);
            callSheet.getRange("A2:E2").setValues([[ clientName, callLocation, new Date(), ticketNumber, priorityStatus ]]); 
            if (callSheet.getLastRow() > 21) {
                callSheet.deleteRow(22);
            }
        }

        _updateTimestamp();
        const newAgendaData = getProfessionalAgendaData(callingProfId);
        return { success: true, newStatus: 'Chamado', newAgendaData: newAgendaData };

    } catch (e) {
        Logger.log(`Erro ao chamar cliente: ${e.message} \nStack: ${e.stack}`);
        return { success: false, message: `Erro interno: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}

function updateAppointmentStatus(appointmentId, newStatus) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.appointments);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowIndex = findRowIndexById(sheet, 'ID_Agendamento', appointmentId);
    if (!rowIndex) {
        Logger.log(`Agendamento ${appointmentId} não encontrado.`);
        return; 
    }
    const statusColIndex = headers.indexOf('Status');
    const profCallColIndex = headers.indexOf('ID_Profissional_Chamada');
    if (statusColIndex === -1 || profCallColIndex === -1) {
        Logger.log('Colunas essenciais não encontradas.');
        return;
    }
    sheet.getRange(rowIndex, statusColIndex + 1).setValue(newStatus);
    if (newStatus === 'Concluído' || newStatus === 'Cancelado') {
      sheet.getRange(rowIndex, profCallColIndex + 1).setValue('');
      const scriptProperties = PropertiesService.getScriptProperties();
      const propKey = `serving_${appointmentId}`;
      scriptProperties.deleteProperty(propKey);
    }
    const rowValues = sheet.getRange(rowIndex, 1, 1, headers.length).getDisplayValues()[0];
    const apptData = headers.reduce((obj, header, index) => {
      obj[header] = rowValues[index];
      return obj;
    }, {});
    apptData.Status = newStatus;

    // OTIMIZAÇÃO: Busca os dados aqui, uma única vez.
    const allProfessionals = getData(SS_CONFIG.sheets.professionals);
    const allConfig = getData(SS_CONFIG.sheets.config);
    const isDeleting = newStatus === 'Cancelado';

    syncAppointmentToCalendar(apptData, allProfessionals, allConfig, isDeleting, false);
    _updateTimestamp();
}

function toggleAppointmentPriority(appointmentId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.appointments);
    const rowIndex = findRowIndexById(sheet, 'ID_Agendamento', appointmentId);
    if (!rowIndex) { return { success: false, message: 'Agendamento não encontrado.' }; }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const priorityColIndex = headers.indexOf('Prioridade');
    if (priorityColIndex === -1) { throw new Error("Coluna 'Prioridade' não encontrada."); }
    const range = sheet.getRange(rowIndex, priorityColIndex + 1);
    const currentValue = range.getValue();
    const newValue = (currentValue === 'SIM') ? '' : 'SIM';
    range.setValue(newValue);
    _updateTimestamp();
    return { success: true, newPriorityStatus: newValue };
  } catch (e) {
    Logger.log(`Erro em toggleAppointmentPriority: ${e.message} ${e.stack}`);
    return { success: false, message: e.message };
  }
}

// *** FUNÇÕES DE ADMINISTRAÇÃO E CRUD ***
function getPaginatedAppointments(options) {
  const page = options.page || 1;
  const pageSize = options.pageSize || 20;
  const filters = options.filters || {};
  try {
    let allAppointments = getAppointments(true); 
    const filtered = allAppointments.filter(appt => {
      if (filters.start && appt.Data < filters.start) return false;
      if (filters.end && appt.Data > filters.end) return false;
      if (filters.prof && appt.ID_Profissional !== filters.prof) return false;
      if (filters.status && appt.Status !== filters.status) return false;
      return true;
    });
    filtered.sort((a, b) => (b.Data + b.Hora).localeCompare(a.Data + a.Hora));
    const totalCount = filtered.length;
    const startIndex = (page - 1) * pageSize;
    const paginatedAppointments = filtered.slice(startIndex, startIndex + pageSize);
    return { success: true, appointments: paginatedAppointments, totalCount: totalCount };
  } catch (e) {
    Logger.log('Erro em getPaginatedAppointments: ' + e.message);
    return { success: false, message: e.message };
  }
}

function generateAppointmentsPDF(filters) {
  try {
    const config = getData(SS_CONFIG.sheets.config); // Reutiliza a função de cache
    const templateIdConfig = config.find(c => c.Chave === 'PDF_TEMPLATE_ID');
    if (!templateIdConfig || !templateIdConfig.Valor) {
      throw new Error("ID do template de PDF não encontrado nas Configurações.");
    }
    const TEMPLATE_ID = templateIdConfig.Valor;
    let allAppointments = getAppointments();
    const professionals = getData(SS_CONFIG.sheets.professionals);
    allAppointments = allAppointments.map(appt => {
      appt.DataYYYYMMDD = appt.Data;
      return appt;
    });
    const filteredAppointments = allAppointments.filter(appt => {
      if (filters.start && appt.DataYYYYMMDD < filters.start) return false;
      if (filters.end && appt.DataYYYYMMDD > filters.end) return false;
      if (filters.prof && appt.ID_Profissional !== filters.prof) return false;
      if (filters.status && appt.Status !== filters.status) return false;
      return true;
    });
    if (filteredAppointments.length === 0) {
      return { success: false, message: "Nenhum agendamento encontrado para os filtros." };
    }
    const totalCount = filteredAppointments.length;
    filteredAppointments.sort((a, b) => (a.DataYYYYMMDD + a.Hora).localeCompare(b.DataYYYYMMDD + b.Hora));
    const templateFile = DriveApp.getFileById(TEMPLATE_ID);
    const newFileName = `Relatório de Agendamentos - ${new Date().toLocaleDateString('pt-BR').replace(/\//g, '-')}.pdf`;
    const tempFolder = DriveApp.getFoldersByName("TempReports").hasNext() ? DriveApp.getFoldersByName("TempReports").next() : DriveApp.createFolder("TempReports");
    const tempFile = templateFile.makeCopy(newFileName + ".copy", tempFolder);
    const tempDoc = DocumentApp.openById(tempFile.getId());
    const body = tempDoc.getBody();
    body.replaceText("{{data_emissao}}", new Date().toLocaleString('pt-BR'));
    body.replaceText("{{usuario_gerador}}", Session.getActiveUser().getEmail());
    let filterDesc = [];
    if (filters.start) filterDesc.push(`De: ${new Date(filters.start+'T00:00:00').toLocaleDateString('pt-BR', {timeZone: 'UTC'})}`);
    if (filters.end) filterDesc.push(`Até: ${new Date(filters.end+'T00:00:00').toLocaleDateString('pt-BR', {timeZone: 'UTC'})}`);
    if (filters.prof) {
      const profName = professionals.find(p => p.ID_Profissional === filters.prof)?.Nome_Completo || 'N/A';
      filterDesc.push(`Profissional: ${profName}`);
    }
    if (filters.status) filterDesc.push(`Status: ${filters.status}`);
    body.replaceText("{{filtros_aplicados}}", filterDesc.length > 0 ? filterDesc.join(' | ') : 'Nenhum');
    body.replaceText("{{total_agendamentos}}", totalCount);
    const table = body.getTables()[0];
    filteredAppointments.forEach(appt => {
      const profName = professionals.find(p => p.ID_Profissional === appt.ID_Profissional)?.Nome_Completo || 'N/A';
      const newRow = table.appendTableRow();
      const [year, month, day] = appt.DataYYYYMMDD.split('-');
      const dataFormatada = `${day}/${month}/${year}`;
      newRow.appendTableCell(dataFormatada);
      newRow.appendTableCell(appt.Hora);
      newRow.appendTableCell(appt.Nome_Cliente);
      newRow.appendTableCell(profName);
      newRow.appendTableCell(appt.Status);
    });
    table.removeRow(1);
    tempDoc.saveAndClose();
    const pdfBlob = tempFile.getAs('application/pdf');
    const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    tempFile.setTrashed(true);
    return { success: true, fileName: newFileName, mimeType: 'application/pdf', data: pdfBase64 };
  } catch (e) {
    Logger.log("Erro ao gerar PDF: " + e.message + " Stack: " + e.stack);
    return { success: false, message: "Ocorreu um erro interno ao gerar o relatório: " + e.message };
  }
}

function doLogin(credentials) {
  const { role, username, password } = credentials;
  const hashedPassword = simpleHash(password);
  try {
    switch (role) {
      case 'master': {
        const users = getData(SS_CONFIG.sheets.masterUsers);
        const userFound = users.find(u => u.Usuario === username && u.Senha_Hash === hashedPassword);
        if (userFound) return { success: true, user: { name: userFound.Usuario, role: 'master' } };
        break;
      }
      case 'atendente': {
        const users = getData(SS_CONFIG.sheets.attendantUsers);
        const userFound = users.find(u => u.Usuario === username && u.Senha_Hash === hashedPassword);
        if (userFound) return { success: true, user: { name: userFound.Usuario, role: 'atendente' } };
        break;
      }
      case 'profissional': {
        const users = getData(SS_CONFIG.sheets.professionals);
        const userFound = users.find(p => p.Contato_Email === username && p.Senha_Hash === hashedPassword && p.Status === 'Ativo');
        if (userFound) return { success: true, user: { id: userFound.ID_Profissional, name: userFound.Nome_Completo, role: 'profissional' } };
        break;
      }
    }
    return { success: false, message: 'Credenciais inválidas ou perfil inativo.' };
  } catch(e) {
    Logger.log(`Erro no Login: ${e.message}`);
    return { success: false, message: 'Ocorreu um erro interno no servidor.'}
  }
}

function getAppointmentDetails(appointmentId) {
  try {
    const appointments = getAppointments(true);
    const appointment = appointments.find(a => a.ID_Agendamento === appointmentId);
    if (!appointment) {
      return { success: false, message: "Agendamento não encontrado." };
    }
    return { success: true, appointment: appointment };
  } catch (e) {
    Logger.log("Erro em getAppointmentDetails: " + e.message);
    return { success: false, message: e.message };
  }
}

function updateAppointment(appointmentId, appointmentData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.appointments);
    const rowIndex = findRowIndexById(sheet, 'ID_Agendamento', appointmentId);
    if (!rowIndex) {
      return { success: false, message: 'Agendamento não encontrado para atualizar.' };
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    appointmentData.ID_Agendamento = appointmentId;
    const originalSchedulingDate = sheet.getRange(rowIndex, headers.indexOf('Data_Agendamento') + 1).getValue();
    appointmentData.Data_Agendamento = originalSchedulingDate;
    const [year, month, day] = appointmentData.Data.split('-');
    const dateForSheet = `${day}/${month}/${year}`;
    const dataToSave = {...appointmentData, Data: dateForSheet};
    const newRow = headers.map(header => dataToSave[header] || '');
    sheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
    
    // OTIMIZAÇÃO: Busca os dados aqui, uma única vez.
    const allProfessionals = getData(SS_CONFIG.sheets.professionals);
    const allConfig = getData(SS_CONFIG.sheets.config);

    syncAppointmentToCalendar(dataToSave, allProfessionals, allConfig, false, false);
    _updateTimestamp();
    const returnData = {...appointmentData, Data: `${year}-${month}-${day}`};
    if (returnData.Data_Agendamento instanceof Date) {
      returnData.Data_Agendamento = returnData.Data_Agendamento.toISOString();
    }
    return { success: true, appointment: returnData };
  } catch (e) {
    Logger.log(`Erro em updateAppointment: ${e.message} ${e.stack}`);
    return { success: false, message: e.message };
  }
}

function deleteAppointment(appointmentId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.appointments);
    const rowIndex = findRowIndexById(sheet, 'ID_Agendamento', appointmentId);
    if (!rowIndex) {
      return { success: false, message: 'Agendamento não encontrado para excluir.' };
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowValues = sheet.getRange(rowIndex, 1, 1, headers.length).getDisplayValues()[0];
    const apptData = headers.reduce((obj, header, index) => {
      obj[header] = rowValues[index];
      return obj;
    }, {});
    sheet.deleteRow(rowIndex);

    // OTIMIZAÇÃO: Busca os dados aqui, uma única vez.
    const allProfessionals = getData(SS_CONFIG.sheets.professionals);
    const allConfig = getData(SS_CONFIG.sheets.config);

    syncAppointmentToCalendar(apptData, allProfessionals, allConfig, true, false);
    _updateTimestamp();
    return { success: true };
  } catch (e) {
    Logger.log(`Erro em deleteAppointment: ${e.message} ${e.stack}`);
    return { success: false, message: e.message };
  }
}

function syncAppointmentToCalendar(appt, allProfessionals, allConfig, deleteEvent = false, isNewAppointment = false) {
  try {
    // OTIMIZAÇÃO: Agora recebe os dados em vez de buscá-los.
    // const professionals = getData(SS_CONFIG.sheets.professionals); <-- LINHA REMOVIDA
    const professional = allProfessionals.find(p => p.ID_Profissional === appt.ID_Profissional);
    if (!professional || !professional.Email_Agenda) return;

    const calendar = CalendarApp.getCalendarById(professional.Email_Agenda);
    if (!calendar) return;

    if (deleteEvent) {
      try {
        const events = calendar.getEvents(new Date(1970), new Date(2100), { search: `ID Agendamento: ${appt.ID_Agendamento}` });
        if (events.length > 0) events[0].deleteEvent();
      } catch(e) { Logger.log(`Evento para deletar não encontrado: ${e.message}`); }
      return;
    }
    
    const eventTitle = `${appt.Nome_Cliente} - ${appt.Servico}`;
    let dateForEvent = appt.Data;
    if (dateForEvent instanceof Date) {
        dateForEvent = Utilities.formatDate(dateForEvent, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof dateForEvent === 'string' && dateForEvent.includes('/')) {
        const [day, month, year] = dateForEvent.split('/');
        dateForEvent = `${year}-${month}-${day}`;
    }

    const startDate = new Date(`${dateForEvent}T${appt.Hora}`);
    // OTIMIZAÇÃO: Agora recebe os dados em vez de buscá-los.
    // const config = getData(SS_CONFIG.sheets.config); <-- LINHA REMOVIDA
    const slotDurationConfig = allConfig.find(c => c.Chave === 'DURACAO_PADRAO_SLOT_MINUTOS');
    const slotDuration = slotDurationConfig ? parseInt(slotDurationConfig.Valor, 10) : 30;
    const endDate = new Date(startDate.getTime() + slotDuration * 60000);
    
    const eventDetails = { description: `Serviço: ${appt.Servico || 'N/A'}\nObs: ${appt.Observacoes || 'Nenhuma'}\n\nID Agendamento: ${appt.ID_Agendamento}` };
    const events = calendar.getEvents(startDate, endDate, { search: `ID Agendamento: ${appt.ID_Agendamento}` });
    let event = events.length > 0 ? events[0] : null;
    
    if (event) {
      event.setTime(startDate, endDate).setTitle(eventTitle).setDescription(eventDetails.description);
    } else {
      calendar.createEvent(eventTitle, startDate, endDate, eventDetails);
    }

    if (isNewAppointment && professional.Receber_Email_Agendamento && professional.Receber_Email_Agendamento.toUpperCase() === 'SIM') {
      const [year, month, day] = dateForEvent.split('-');
      const dataFormatada = `${day}/${month}/${year}`;
      const subject = `Novo Agendamento: ${appt.Nome_Cliente} às ${appt.Hora}`;
      const emailBody = `<h3>Novo Agendamento</h3><p><strong>Cliente:</strong> ${appt.Nome_Cliente}</p><p><strong>Data:</strong> ${dataFormatada}</p><p><strong>Hora:</strong> ${appt.Hora}</p><p><strong>Serviço:</strong> ${appt.Servico}</p>`;
      MailApp.sendEmail({ to: professional.Email_Agenda, subject: subject, htmlBody: emailBody });
    }
  } catch (e) {
    Logger.log(`Erro ao sincronizar com Google Agenda: ${e.message} - ${e.stack}`);
  }
}

function addProfessional(profData) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.professionals);
    const newId = generateNewProfessionalId();
    profData.ID_Profissional = newId;
    if (profData.Senha) { profData.Senha_Hash = simpleHash(profData.Senha); }
    delete profData.Senha;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const newRow = headers.map(header => profData[header] || '');
    sheet.appendRow(newRow);
    getData(SS_CONFIG.sheets.professionals, true);
    _updateTimestamp();
    return { success: true, professional: profData };
}

function editProfessional(profId, profData) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.professionals);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.trim());
    const idColIndex = headers.indexOf('ID_Profissional');
    const rowIndex = data.findIndex(row => row[idColIndex] === profId);
    if (rowIndex === -1) { return { success: false, message: 'Profissional não encontrado.' }; }
    if (profData.Senha && profData.Senha.length > 0) {
      profData.Senha_Hash = simpleHash(profData.Senha);
    } else {
      profData.Senha_Hash = data[rowIndex][headers.indexOf('Senha_Hash')];
    }
    delete profData.Senha;
    const newRow = headers.map((header, index) => profData.hasOwnProperty(header) ? profData[header] : data[rowIndex][index]);
    sheet.getRange(rowIndex + 1, 1, 1, newRow.length).setValues([newRow]);
    getData(SS_CONFIG.sheets.professionals, true);
    _updateTimestamp();
    return { success: true, professional: profData };
}

function updateSetting(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.config);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.trim());
  const keyColIndex = headers.indexOf('Chave');
  const valueColIndex = headers.indexOf('Valor');
  const rowIndex = data.findIndex(row => row[keyColIndex] === key);
  if(rowIndex !== -1) {
    sheet.getRange(rowIndex + 1, valueColIndex + 1).setValue(value);
    getData(SS_CONFIG.sheets.config, true);
    _updateTimestamp();
    return {success: true, key, value};
  }
  return {success: false, message: "Chave de configuração não encontrada."};
}

function addAttendant(attendantData) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.attendantUsers);
    const users = getData(SS_CONFIG.sheets.attendantUsers);
    if (users.some(u => u.Usuario === attendantData.Usuario)) {
        return { success: false, message: 'Este nome de usuário já existe.' };
    }
    const newAttendant = { Usuario: attendantData.Usuario, Senha_Hash: simpleHash(attendantData.Senha) };
    sheet.appendRow([newAttendant.Usuario, newAttendant.Senha_Hash]);
    getData(SS_CONFIG.sheets.attendantUsers, true);
    _updateTimestamp();
    return { success: true, attendant: newAttendant };
}

function editAttendant(originalUsername, newData) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.attendantUsers);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.trim());
    const userColIndex = headers.indexOf('Usuario');
    const passColIndex = headers.indexOf('Senha_Hash');
    const rowIndex = data.findIndex(row => row[userColIndex] === originalUsername);
    if (rowIndex === -1) { return { success: false, message: 'Atendente não encontrado.' }; }
    sheet.getRange(rowIndex + 1, userColIndex + 1).setValue(newData.Usuario);
    if (newData.Senha && newData.Senha.length > 0) {
        sheet.getRange(rowIndex + 1, passColIndex + 1).setValue(simpleHash(newData.Senha));
    }
    getData(SS_CONFIG.sheets.attendantUsers, true);
    _updateTimestamp();
    return { success: true, user: { Usuario: newData.Usuario } };
}

function deleteAttendant(username) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.attendantUsers);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.trim());
    const userColIndex = headers.indexOf('Usuario');
    const rowIndex = data.findIndex(row => row[userColIndex] === username);
    if (rowIndex === -1) { return { success: false, message: 'Atendente não encontrado.' }; }
    sheet.deleteRow(rowIndex + 1);
    getData(SS_CONFIG.sheets.attendantUsers, true);
    _updateTimestamp();
    return { success: true };
}

function addService(serviceName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.services);
    const services = getData(SS_CONFIG.sheets.services);
    if(services.some(s => s.Nome_Servico === serviceName)) { return { success: false, message: 'Este serviço já existe.'}; }
    sheet.appendRow([serviceName]);
    getData(SS_CONFIG.sheets.services, true);
    _updateTimestamp();
    return { success: true, service: { Nome_Servico: serviceName }};
}

function editService(originalName, newName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.services);
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0] === originalName);
    if(rowIndex === -1) { return { success: false, message: 'Serviço não encontrado.' }; }
    sheet.getRange(rowIndex + 1, 1).setValue(newName);
    getData(SS_CONFIG.sheets.services, true);
    _updateTimestamp();
    return { success: true, service: { Nome_Servico: newName }};
}

function deleteService(serviceName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.services);
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0] === serviceName);
    if(rowIndex === -1) { return { success: false, message: 'Serviço não encontrado.' }; }
    sheet.deleteRow(rowIndex + 1);
    getData(SS_CONFIG.sheets.services, true);
    _updateTimestamp();
    return { success: true };
}

function importOldAppointments() {
  const professionals = getData(SS_CONFIG.sheets.professionals);
  const profMap = professionals.reduce((map, prof) => {
    map[prof.Nome_Completo.trim().toLowerCase()] = prof.ID_Profissional;
    return map;
  }, {});
  const oldData = getSheetData(SS_CONFIG.sheets.importSheet);
  const appointmentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.appointments);
  const headers = appointmentsSheet.getRange(1, 1, 1, appointmentsSheet.getLastColumn()).getValues()[0].map(h => h.trim());
  let successCount = 0, errorCount = 0, newRows = [];
  oldData.forEach(row => {
    const profName = row.Profissional ? row.Profissional.trim().toLowerCase() : '';
    const profId = profMap[profName];
    if (profId) {
      const newRow = headers.map(header => {
        switch (header) {
          case 'ID_Agendamento': return Utilities.getUuid();
          case 'Data_Agendamento': return new Date();
          case 'ID_Profissional': return profId;
          case 'Data': return row.Data || '';
          case 'Hora': return row.Hora || '';
          case 'Nome_Cliente': return row.Nome_Cliente || '';
          case 'Telefone_WhatsApp': return row.Telefone_WhatsApp || '';
          case 'Servico': return row.Servico || '';
          case 'Observacoes': return row.Observacoes || '';
          case 'Status': return row.Status || 'Concluído';
          default: return '';
        }
      });
      newRows.push(newRow);
      successCount++;
    } else {
      errorCount++;
      Logger.log(`Profissional não encontrado para: "${row.Profissional}"`);
    }
  });
  if (newRows.length > 0) {
    appointmentsSheet.getRange(appointmentsSheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }
  _updateTimestamp();
  return { success: true, imported: successCount, failed: errorCount };
}

function simpleHash(input) {
  let hash = 0;
  if (!input || input.length === 0) return "0";
  for (let i = 0; i < input.length; i++) {
    const char = input.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash |= 0;
  }
  return hash.toString();
}

function generateNewProfessionalId() {
  const data = getData(SS_CONFIG.sheets.professionals);
  let maxIdNum = 0;
  data.forEach(prof => {
    if (prof.ID_Profissional && String(prof.ID_Profissional).startsWith('P')) {
      const num = parseInt(String(prof.ID_Profissional).substring(1), 10);
      if (!isNaN(num) && num > maxIdNum) {
        maxIdNum = num;
      }
    }
  });
  return 'P' + String(maxIdNum + 1).padStart(3, '0');
}

function generatePasswordHash() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Gerador de Hash', 'Digite a senha:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const plainTextPassword = response.getResponseText();
    if (plainTextPassword) {
      const hashedPassword = simpleHash(plainTextPassword);
      ui.alert(`Hash Gerado`, `O hash para "${plainTextPassword}" é:\n\n${hashedPassword}`, ui.ButtonSet.OK);
    }
  }
}

function getCallScreenData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chamada_Atual");
    if (!sheet) return [];
    const data = sheet.getRange(2, 1, Math.min(sheet.getLastRow() - 1, 20), 3).getDisplayValues();
    return data.map(row => ({ client: row[0], professional: row[1], time: row[2] }));
  } catch (e) {
    Logger.log("Erro ao buscar dados da tela de chamada: " + e.message);
    return [];
  }
}

// ========================================================
// FUNÇÕES DE GERENCIAMENTO DE CLIENTES E HISTÓRICO
// ========================================================

function getClientById(clientId) {
    const clients = getSheetData(SS_CONFIG.sheets.clients);
    return clients.find(c => c.ID_Cliente === clientId);
}

function addOrUpdateClient(clientData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.clients);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  if (clientData.ID_Cliente) { // Atualização
    const rowIndex = findRowIndexById(sheet, 'ID_Cliente', clientData.ID_Cliente);
    if (!rowIndex) return { success: false, message: 'Cliente não encontrado.' };
    
    const row = headers.map(header => clientData[header] || '');
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    return { success: true, client: clientData };

  } else { // Novo Cliente
    const newId = Utilities.getUuid();
    clientData.ID_Cliente = newId;
    clientData.Data_Cadastro = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    
    const newRow = headers.map(header => clientData[header] || '');
    sheet.appendRow(newRow);
    return { success: true, client: clientData };
  }
}

function getClientNotes(clientId) {
  try {
    const allNotes = getSheetData(SS_CONFIG.sheets.clientHistory);
    const clientNotes = allNotes.filter(note => note.ID_Cliente == clientId);
    
    const formattedNotes = clientNotes.map(note => {
        let dataFormatada = note.Data_Anotacao;
        try {
            dataFormatada = Utilities.formatDate(new Date(note.Data_Anotacao), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        } catch(e) {/* ignora */}
        
        return {...note, DataFormatada: dataFormatada };
    }).sort((a,b) => new Date(b.Data_Anotacao) - new Date(a.Data_Anotacao));

    return { success: true, notes: formattedNotes };
  } catch (e) {
    Logger.log(`Erro em getClientNotes: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function saveClientNote(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.clientHistory);
    const newNote = {
      ID_Anotacao: Utilities.getUuid(),
      ID_Cliente: data.clientId,
      Nome_Cliente: data.clientName,
      Anotacao: data.note,
      ID_Profissional: data.professionalId,
      Nome_Profissional: data.professionalName,
      Data_Anotacao: new Date()
    };
    sheet.appendRow([
      newNote.ID_Anotacao, newNote.ID_Cliente, newNote.Nome_Cliente,
      newNote.Anotacao, newNote.ID_Profissional, newNote.Nome_Profissional,
      newNote.Data_Anotacao
    ]);
    _updateTimestamp();
    return { success: true };
  } catch (e) {
    Logger.log(`Erro em saveClientNote: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function updateClientNote(noteId, newText) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.clientHistory);
    const rowIndex = findRowIndexById(sheet, 'ID_Anotacao', noteId);
    if (!rowIndex) throw new Error("Anotação não encontrada.");
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const noteColIndex = headers.indexOf('Anotacao');
    
    sheet.getRange(rowIndex, noteColIndex + 1).setValue(newText);
    _updateTimestamp();
    return { success: true };
  } catch (e) {
    Logger.log(`Erro em updateClientNote: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function deleteClientNote(noteId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_CONFIG.sheets.clientHistory);
    const rowIndex = findRowIndexById(sheet, 'ID_Anotacao', noteId);
    if (!rowIndex) throw new Error("Anotação não encontrada.");
    
    sheet.deleteRow(rowIndex);
    _updateTimestamp();
    return { success: true };
  } catch (e) {
    Logger.log(`Erro em deleteClientNote: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function getLoginScreenConfig() {
  try {
    const config = getData(SS_CONFIG.sheets.config);
    const logoUrl = config.find(c => c.Chave === 'LOGO_URL')?.Valor || "";
    const businessName = config.find(c => c.Chave === 'NOME_NEGOCIO')?.Valor || "Sistema de Agendamento";
    return { logoUrl, businessName };
  } catch (e) {
    return { logoUrl: "", businessName: "Sistema de Agendamento" };
  }
}

function recallClient(appointmentId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const apptSheet = ss.getSheetByName(SS_CONFIG.sheets.appointments);
    const rowIndex = findRowIndexById(apptSheet, 'ID_Agendamento', appointmentId);
    if (!rowIndex) return { success: false, message: 'Agendamento não encontrado.' };
    
    const headers = apptSheet.getRange(1, 1, 1, apptSheet.getLastColumn()).getValues()[0];
    
    const rowData = apptSheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const appointmentToRecall = headers.reduce((obj, header, i) => {
        obj[header] = rowData[i];
        return obj;
    }, {});

    const clientName = appointmentToRecall.Nome_Cliente;
    const ticketNumber = appointmentToRecall.Numero_Ficha;
    const priorityStatus = appointmentToRecall.Prioridade;
    const callingProfId = appointmentToRecall.ID_Profissional_Chamada;
    
    const professionals = getData(SS_CONFIG.sheets.professionals, true);
    const callingProfessional = professionals.find(p => p.ID_Profissional == callingProfId);
    const callLocation = callingProfessional ? (callingProfessional.Sala_Atendimento || callingProfessional.Nome_Completo) : 'N/A';
    
    const callSheet = ss.getSheetByName(SS_CONFIG.sheets.callScreen);
    if (callSheet) {
      callSheet.getRange("A2:E2").setValues([[ clientName, callLocation, new Date(), ticketNumber, priorityStatus ]]);
    }
    
    _updateTimestamp();
    return { success: true };
    
  } catch (e) {
    Logger.log(`Erro ao rechamar cliente: ${e.message}`);
    return { success: false, message: e.message };
  }
}
function searchClients(searchTerm, page = 1, pageSize = 50) {
  try {
    const allClients = getSheetData(SS_CONFIG.sheets.clients).sort((a,b) => a.Nome_Completo.localeCompare(b.Nome_Completo));
    let filteredClients = allClients;

    if (searchTerm && searchTerm.trim() !== '') {
      const lowerCaseSearchTerm = searchTerm.toLowerCase();
      filteredClients = allClients.filter(client => {
        const nameMatch = client.Nome_Completo && client.Nome_Completo.toLowerCase().includes(lowerCaseSearchTerm);
        const phoneMatch = client.Telefone_WhatsApp && client.Telefone_WhatsApp.includes(searchTerm);
        return nameMatch || phoneMatch;
      });
    }

    const totalCount = filteredClients.length;
    const startIndex = (page - 1) * pageSize;
    const paginatedClients = filteredClients.slice(startIndex, startIndex + pageSize);

    return { success: true, clients: paginatedClients, totalCount: totalCount };

  } catch (e) {
    Logger.log(`Erro em searchClients: ${e.message}`);
    return { success: false, message: e.message };
  }
}
// ========================================================
// FUNÇÃO DE BACKUP E LIMPEZA DE AGENDAMENTOS
// ========================================================

function archiveAndClearAppointments() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Espera até 30s para obter o lock

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appointmentsSheet = ss.getSheetByName(SS_CONFIG.sheets.appointments);
    
    if (appointmentsSheet.getLastRow() <= 1) {
      return { success: true, message: "Nenhum agendamento para arquivar." };
    }

    // 1. Ler todos os dados da aba de agendamentos
    const dataToArchive = appointmentsSheet.getDataRange().getValues();

    // 2. Criar um nome único para a nova aba de backup
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
    const backupSheetName = `Backup_Agendamentos_${timestamp}`;
    
    // 3. Criar a nova aba e copiar os dados para ela
    const backupSheet = ss.insertSheet(backupSheetName);
    backupSheet.getRange(1, 1, dataToArchive.length, dataToArchive[0].length).setValues(dataToArchive);
    Logger.log(`${dataToArchive.length - 1} agendamentos arquivados na aba: ${backupSheetName}`);

    // 4. Limpar a aba de agendamentos original (mantendo o cabeçalho)
    appointmentsSheet.getRange(2, 1, appointmentsSheet.getLastRow() - 1, appointmentsSheet.getLastColumn()).clearContent();
    
    // 5. Atualizar o timestamp geral para notificar os clientes de uma grande mudança
    _updateTimestamp();
    SpreadsheetApp.flush(); // Garante que todas as operações sejam concluídas

    return { success: true, message: `Backup concluído! ${dataToArchive.length - 1} registros foram arquivados.` };

  } catch (e) {
    Logger.log(`Erro ao arquivar agendamentos: ${e.message} ${e.stack}`);
    throw new Error(`Ocorreu um erro no servidor durante o arquivamento: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

// ========================================================
// FUNÇÃO PARA CLONAR O SISTEMA PARA UM NOVO CLIENTE
// ========================================================

function createSystemCopy() {
  const ui = SpreadsheetApp.getUi();

  // 1. Pede o nome para a nova cópia
  const response = ui.prompt(
    'Criar Cópia do Sistema',
    'Digite o nome do novo cliente ou da empresa para identificar a cópia:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK || !response.getResponseText()) {
    ui.alert('Operação cancelada.');
    return;
  }
  const clientName = response.getResponseText().trim();
  const newSystemName = `Sistema Agendamento - ${clientName}`;

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Iniciando a cópia para "${clientName}"... Isso pode levar um minuto.`, 'Por favor, aguarde', -1);

    const originalSpreadsheet = SpreadsheetApp.getActive();
    const originalSpreadsheetFile = DriveApp.getFileById(originalSpreadsheet.getId());
    
    // 2. Encontra o template de PDF original usando a configuração
    const configSheet = originalSpreadsheet.getSheetByName(SS_CONFIG.sheets.config);
    const configData = configSheet.getDataRange().getValues();
    const headers = configData.shift();
    const keyColIndex = headers.indexOf('Chave');
    const valueColIndex = headers.indexOf('Valor');
    
    const templateIdRow = configData.find(row => row[keyColIndex] === 'PDF_TEMPLATE_ID');
    if (!templateIdRow) {
      throw new Error("Não foi possível encontrar a chave 'PDF_TEMPLATE_ID' na aba de Configurações.");
    }
    const originalTemplateId = templateIdRow[valueColIndex];
    const originalTemplateFile = DriveApp.getFileById(originalTemplateId);

    // 3. Copia a Planilha (base de dados + script)
    const newSpreadsheetFile = originalSpreadsheetFile.makeCopy(newSystemName);

    // 4. Copia o Template de PDF
    const newTemplateFile = originalTemplateFile.makeCopy(`Template Relatório - ${clientName}`);
    const newTemplateId = newTemplateFile.getId();

    // 5. Abre a NOVA planilha e atualiza o ID do template de PDF
    const newSpreadsheet = SpreadsheetApp.openById(newSpreadsheetFile.getId());
    const newConfigSheet = newSpreadsheet.getSheetByName(SS_CONFIG.sheets.config);
    const newConfigData = newConfigSheet.getDataRange().getValues();
    const newHeaders = newConfigData.shift();
    const newKeyColIdx = newHeaders.indexOf('Chave');
    const newValueColIdx = newHeaders.indexOf('Valor');

    const newTemplateIdRowIndex = newConfigData.findIndex(row => row[newKeyColIdx] === 'PDF_TEMPLATE_ID');
    if (newTemplateIdRowIndex !== -1) {
      // O índice da linha na planilha é o índice do array + 2 (cabeçalho + 1-based index)
      newConfigSheet.getRange(newTemplateIdRowIndex + 2, newValueColIdx + 1).setValue(newTemplateId);
    }

    // 6. (Opcional) Limpa os dados de exemplo da nova planilha
    const sheetsToClear = [
      SS_CONFIG.sheets.appointments, 
      SS_CONFIG.sheets.clients, 
      SS_CONFIG.sheets.clientHistory, 
      SS_CONFIG.sheets.callScreen
    ];
    
    sheetsToClear.forEach(sheetName => {
      const sheet = newSpreadsheet.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
      }
    });
    
    // Limpa o cache e reseta o contador de fichas na nova planilha
    PropertiesService.getScriptProperties().deleteAllProperties();

    SpreadsheetApp.flush(); // Garante que todas as alterações sejam salvas

    // 7. Exibe o resultado para o usuário
    const newSpreadsheetUrl = newSpreadsheetFile.getUrl();
    const htmlOutput = HtmlService.createHtmlOutput(
        `<h3>Sistema para "${clientName}" criado com sucesso!</h3>` +
        `<p>A nova planilha está pronta. Agora você precisa:</p>` +
        `<ol>` +
        `  <li>Abrir o novo template de relatório para <b><a href="${newTemplateFile.getUrl()}" target="_blank">alterar a logo do cliente</a></b>.</li>` +
        `  <li>Abrir a <b><a href="${newSpreadsheetUrl}" target="_blank">nova planilha</a></b>, ir em "Extensões > Apps Script", e fazer o deploy da nova versão da API.</li>` +
        `  <li>Usar o novo link da API no sistema front-end do cliente.</li>` +
        `</ol>`
      )
      .setWidth(500)
      .setHeight(300);
    ui.showModalDialog(htmlOutput, 'Próximos Passos');
    originalSpreadsheet.toast('Cópia concluída!', 'Sucesso!', 5);

  } catch (e) {
    Logger.log(e);
    ui.alert('Ocorreu um Erro', `Não foi possível criar a cópia: ${e.message}`, ui.ButtonSet.OK);
    SpreadsheetApp.getActiveSpreadsheet().toast('Falha na cópia.', 'Erro', 5);
  }
}

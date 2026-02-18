// ============================================================
// GOOGLE APPS SCRIPT — EMDA Conexão Moda v4
// + Email automático de notificação
// + Salvar foto no Google Drive
// + Sem coluna ano_conclusao (já está nos cursos)
// ============================================================

const SPREADSHEET_ID = '1oj57-yAspnZZGdjCGXQbDuySAofWHYJVnb9Rtn1cyCY';
const SHEET_NAME = 'Currículos';

// Email(s) que receberão as notificações (separar por vírgula para múltiplos)
const NOTIFY_EMAIL = 'fernandadeniseaguiar@gmail.com';

// Nome da pasta no Drive para salvar fotos
const FOTOS_FOLDER_NAME = 'EMDA - Fotos Currículos';

// ============================================================
// GET — Verificação de duplicatas
// ============================================================

function doGet(e) {
  try {
    var action = e.parameter.action;
    
    if (action === 'check') {
      return handleDuplicateCheck(e);
    }
    
    if (action === 'list') {
      return handleList(e);
    }
    
    if (action === 'delete') {
      return handleDelete(e);
    }
    
    if (action === 'changepin') {
      return handleChangePin(e);
    }
    
    return jsonResponse({ status: 'ok', message: 'EMDA Conexão Moda API v5' });
    
  } catch (error) {
    return jsonResponse({ error: error.message });
  }
}

function handleDuplicateCheck(e) {
  var field = e.parameter.field;
  var value = (e.parameter.value || '').trim().toLowerCase();
  
  if (!field || !value) {
    return jsonResponse({ found: false });
  }
  
  var sheet = getSheet();
  if (!sheet) {
    return jsonResponse({ found: false, error: 'Planilha não configurada' });
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colIndex = headers.indexOf(field);
  if (colIndex === -1) {
    return jsonResponse({ found: false });
  }
  
  for (var i = 1; i < data.length; i++) {
    var cellValue = String(data[i][colIndex] || '').trim().toLowerCase();
    
    if (field === 'whatsapp') {
      cellValue = cellValue.replace(/\D/g, '');
      var searchValue = value.replace(/\D/g, '');
      if (cellValue === searchValue && cellValue.length >= 10) {
        return jsonResponse({
          found: true,
          data: {
            nome: data[i][headers.indexOf('nome')] || '',
            timestamp: data[i][headers.indexOf('timestamp')] || ''
          }
        });
      }
    } else {
      if (cellValue === value) {
        return jsonResponse({
          found: true,
          data: {
            nome: data[i][headers.indexOf('nome')] || '',
            timestamp: data[i][headers.indexOf('timestamp')] || ''
          }
        });
      }
    }
  }
  
  return jsonResponse({ found: false });
}

// ============================================================
// GET — Listar todos os cadastros (admin)
// ============================================================

function handleList(e) {
  var pin = e.parameter.pin;
  var biotoken = e.parameter.biotoken;
  if (pin !== getAdminPin() && biotoken !== 'emda-bio-auth') {
    return jsonResponse({ error: 'Não autorizado' });
  }
  
  var sheet = getSheet();
  if (!sheet) return jsonResponse({ data: [] });
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var results = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j] || '';
    }
    row._row = i + 1; // número da linha na planilha
    results.push(row);
  }
  
  return jsonResponse({ data: results });
}

// ============================================================
// GET — Excluir cadastro (admin)
// ============================================================

function handleDelete(e) {
  var pin = e.parameter.pin;
  var biotoken = e.parameter.biotoken;
  if (pin !== getAdminPin() && biotoken !== 'emda-bio-auth') {
    return jsonResponse({ error: 'Não autorizado' });
  }
  
  var rowNum = parseInt(e.parameter.row);
  if (!rowNum || rowNum < 2) {
    return jsonResponse({ error: 'Linha inválida' });
  }
  
  var sheet = getSheet();
  if (!sheet) return jsonResponse({ error: 'Planilha não encontrada' });
  
  // Pegar link da foto antes de deletar
  var fotoLink = sheet.getRange(rowNum, 14).getValue(); // coluna N = foto_link
  
  // Deletar foto do Drive se existir
  if (fotoLink && fotoLink.indexOf('drive.google.com') !== -1) {
    try {
      var fileId = fotoLink.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (fileId && fileId[1]) {
        DriveApp.getFileById(fileId[1]).setTrashed(true);
      }
    } catch (err) {
      Logger.log('Erro ao deletar foto: ' + err.message);
    }
  }
  
  // Deletar linha
  sheet.deleteRow(rowNum);
  
  return jsonResponse({ success: true });
}

// ============================================================
// POST — Salvar currículo + Email + Foto no Drive
// ============================================================

function doPost(e) {
  try {
    var data;
    
    if (e.parameter && e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      return jsonResponse({ success: false, error: 'Sem dados recebidos' });
    }
    
    var sheet = getSheet();
    if (!sheet) {
      return jsonResponse({ success: false, error: 'Planilha não configurada' });
    }
    
    // Salvar foto no Google Drive (se existir)
    var fotoLink = '';
    if (data.foto_base64 && data.foto_base64.startsWith('data:image')) {
      fotoLink = salvarFotoNoDrive(data.foto_base64, data.nome);
    }
    
    // Salvar na planilha (SEM ano_conclusao — já está dentro de cursos)
    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.nome || '',
      data.email || '',
      data.whatsapp || '',
      data.cidade || '',
      data.estado || '',
      data.cursos || '',
      data.experiencia || '',
      data.portfolio || '',
      data.instagram || '',
      data.linkedin || '',
      data.sobre || '',
      data.foto || 'Não',
      fotoLink
    ]);
    
    // Enviar email de notificação
    enviarEmailNotificacao(data, fotoLink);
    
    return jsonResponse({ success: true });
    
  } catch (error) {
    return jsonResponse({ success: false, error: error.message });
  }
}

// ============================================================
// Salvar Foto no Google Drive
// ============================================================

function salvarFotoNoDrive(base64Data, nomeAluno) {
  try {
    var folder = getOrCreateFolder(FOTOS_FOLDER_NAME);
    
    var parts = base64Data.split(',');
    var mimeMatch = parts[0].match(/data:(image\/\w+);base64/);
    
    if (!mimeMatch || !parts[1]) {
      return '';
    }
    
    var mimeType = mimeMatch[1];
    var extension = mimeType.split('/')[1].replace('jpeg', 'jpg');
    var imageData = Utilities.base64Decode(parts[1]);
    var blob = Utilities.newBlob(imageData, mimeType);
    
    var timestamp = new Date().toISOString().slice(0, 10);
    var nomeClean = (nomeAluno || 'sem-nome').replace(/[^a-zA-ZÀ-ú\s]/g, '').replace(/\s+/g, '-').substring(0, 30);
    var fileName = nomeClean + '_' + timestamp + '.' + extension;
    
    blob.setName(fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
    
  } catch (error) {
    Logger.log('Erro ao salvar foto: ' + error.message);
    return 'Erro: ' + error.message;
  }
}

function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

// ============================================================
// Email de Notificação
// ============================================================

function enviarEmailNotificacao(data, fotoLink) {
  try {
    var cursos = (data && data.cursos) ? data.cursos : 'Não informado';
    var cidade = (data && data.cidade) ? (data.cidade + '/' + data.estado) : 'Não informada';
    var dataHora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    var nome = (data && data.nome) ? data.nome : 'Sem nome';
    var email = (data && data.email) ? data.email : '-';
    var whatsapp = (data && data.whatsapp) ? data.whatsapp : '-';
    var experiencia = (data && data.experiencia) ? data.experiencia : '';
    var instagram = (data && data.instagram) ? data.instagram.replace(/^@+/, '') : '';
    var portfolio = (data && data.portfolio) ? data.portfolio : '';
    var linkedin = (data && data.linkedin) ? data.linkedin : '';
    var sobre = (data && data.sobre) ? data.sobre : '';
    
    // Formatar cursos para email (trocar \n por <br>)
    var cursosHtml = cursos.replace(/\n/g, '<br>');
    
    var assunto = '📋 Novo Currículo - ' + nome;
    
    var corpo = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;">';
    
    // Header
    corpo += '<div style="background:#000;padding:24px 32px;text-align:center;">';
    corpo += '<h1 style="color:#C9A962;font-size:18px;font-weight:400;letter-spacing:2px;margin:0;">CONEXÃO MODA</h1>';
    corpo += '<p style="color:rgba(255,255,255,0.5);font-size:11px;margin:4px 0 0 0;letter-spacing:1px;">ESCOLA DE MODA DENISE AGUIAR</p>';
    corpo += '</div>';
    
    // Body
    corpo += '<div style="padding:32px;">';
    corpo += '<p style="color:#666;font-size:13px;margin:0 0 24px 0;">Novo currículo cadastrado em <strong>' + dataHora + '</strong></p>';
    
    // Tabela
    corpo += '<table style="width:100%;border-collapse:collapse;margin-bottom:24px;">';
    corpo += montarLinha('Nome', nome, true);
    corpo += montarLinha('Email', '<a href="mailto:' + email + '" style="color:#C9A962;text-decoration:none;">' + email + '</a>', false);
    corpo += montarLinha('WhatsApp', whatsapp, true);
    corpo += montarLinha('Cidade', cidade, false);
    corpo += montarLinha('Cursos', cursosHtml, true);
    if (experiencia) corpo += montarLinha('Experiência', experiencia, false);
    if (instagram) corpo += montarLinha('Instagram', '<a href="https://instagram.com/' + instagram + '" style="color:#C9A962;text-decoration:none;">@' + instagram + '</a>', true);
    if (portfolio) corpo += montarLinha('Portfólio', '<a href="' + portfolio + '" style="color:#C9A962;text-decoration:none;">' + portfolio + '</a>', false);
    if (linkedin) corpo += montarLinha('LinkedIn', '<a href="' + linkedin + '" style="color:#C9A962;text-decoration:none;">Ver perfil</a>', true);
    if (sobre) corpo += montarLinha('Sobre', sobre, false);
    if (fotoLink) corpo += montarLinha('Foto', '<a href="' + fotoLink + '" style="color:#C9A962;text-decoration:none;">📷 Ver foto</a>', true);
    corpo += '</table>';
    
    // Botão
    corpo += '<div style="text-align:center;margin-top:24px;">';
    corpo += '<a href="https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit" ';
    corpo += 'style="display:inline-block;background:#000;color:#C9A962;padding:12px 32px;text-decoration:none;border-radius:8px;font-size:13px;">Abrir Planilha</a>';
    corpo += '</div></div>';
    
    // Footer
    corpo += '<div style="background:#f8f7f5;padding:16px 32px;text-align:center;border-top:1px solid #eee;">';
    corpo += '<p style="color:#999;font-size:11px;margin:0;">Escola de Moda Denise Aguiar — Conexão Moda</p>';
    corpo += '</div></div>';
    
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: assunto,
      htmlBody: corpo
    });
    
    Logger.log('Email enviado para ' + NOTIFY_EMAIL);
    
  } catch (error) {
    Logger.log('Erro ao enviar email: ' + error.message);
  }
}

function montarLinha(label, valor, destacar) {
  var bg = destacar ? 'background:#f8f7f5;' : '';
  var html = '<tr>';
  html += '<td style="padding:10px 12px;' + bg + 'border-bottom:1px solid #eee;width:130px;color:#999;font-size:12px;text-transform:uppercase;letter-spacing:0.5px;vertical-align:top;">' + label + '</td>';
  html += '<td style="padding:10px 12px;' + bg + 'border-bottom:1px solid #eee;font-size:14px;color:#333;">' + valor + '</td>';
  html += '</tr>';
  return html;
}

// ============================================================
// PIN de Administrador (armazenado em Script Properties)
// ============================================================

var DEFAULT_PIN = '2026@Tifannypaes';

function getAdminPin() {
  var props = PropertiesService.getScriptProperties();
  var pin = props.getProperty('ADMIN_PIN');
  if (!pin) {
    props.setProperty('ADMIN_PIN', DEFAULT_PIN);
    return DEFAULT_PIN;
  }
  return pin;
}

function handleChangePin(e) {
  var currentPin = e.parameter.current;
  var newPin = e.parameter.newpin;
  
  if (currentPin !== getAdminPin()) {
    return jsonResponse({ error: 'Senha atual incorreta' });
  }
  
  if (!newPin || newPin.length < 4) {
    return jsonResponse({ error: 'Nova senha deve ter pelo menos 4 caracteres' });
  }
  
  var props = PropertiesService.getScriptProperties();
  props.setProperty('ADMIN_PIN', newPin);
  
  return jsonResponse({ success: true });
}

// ============================================================
// Helpers
// ============================================================

function getSheet() {
  if (!SPREADSHEET_ID) return null;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// RODE UMA VEZ: Recriar headers e formatar planilha
// ============================================================

function reformatarPlanilha() {
  var sheet = getSheet();
  
  // Novos headers (sem ano_conclusao, portfolio antes de instagram)
  var headers = [
    'timestamp', 'nome', 'email', 'whatsapp', 'cidade', 'estado',
    'cursos', 'experiencia', 'portfolio', 'instagram', 
    'linkedin', 'sobre', 'foto', 'foto_link'
  ];
  
  // Atualizar header (linha 1)
  for (var i = 0; i < headers.length; i++) {
    sheet.getRange(1, i + 1).setValue(headers[i]);
  }
  
  // Limpar colunas extras (se existir coluna 15 antiga)
  if (sheet.getMaxColumns() > 14) {
    var extra = sheet.getRange(1, 15, 1, sheet.getMaxColumns() - 14);
    extra.clear();
  }
  
  // Formatação do header
  var headerRange = sheet.getRange(1, 1, 1, 14);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#000000');
  headerRange.setFontColor('#C9A962');
  headerRange.setFontSize(9);
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  
  // Larguras otimizadas
  sheet.setColumnWidth(1, 160);   // timestamp
  sheet.setColumnWidth(2, 220);   // nome
  sheet.setColumnWidth(3, 220);   // email
  sheet.setColumnWidth(4, 140);   // whatsapp
  sheet.setColumnWidth(5, 140);   // cidade
  sheet.setColumnWidth(6, 50);    // estado
  sheet.setColumnWidth(7, 280);   // cursos
  sheet.setColumnWidth(8, 250);   // experiencia
  sheet.setColumnWidth(9, 200);   // portfolio
  sheet.setColumnWidth(10, 140);  // instagram
  sheet.setColumnWidth(11, 200);  // linkedin
  sheet.setColumnWidth(12, 250);  // sobre
  sheet.setColumnWidth(13, 50);   // foto
  sheet.setColumnWidth(14, 300);  // foto_link
  
  // Wrap text em colunas com texto longo
  sheet.getRange('G:G').setWrap(true);   // cursos
  sheet.getRange('H:H').setWrap(true);   // experiencia
  sheet.getRange('L:L').setWrap(true);   // sobre
  sheet.getRange('N:N').setWrap(true);   // foto_link
  
  // Alinhamento vertical no topo
  sheet.getRange('A:N').setVerticalAlignment('top');
  
  // Altura automática da linha de dados
  sheet.setRowHeight(1, 30);
  
  Logger.log('Planilha reformatada com sucesso!');
  Logger.log('IMPORTANTE: Se havia dados antigos com 15 colunas, reorganize manualmente a linha existente.');
}

// ============================================================
// Teste de email
// ============================================================

function testarEmail() {
  enviarEmailNotificacao({
    nome: 'Teste Email',
    email: 'teste@teste.com',
    whatsapp: '(31) 99999-9999',
    cidade: 'Belo Horizonte',
    estado: 'MG',
    cursos: 'Design de Moda (2025)\nTecidos (2026)',
    experiencia: '',
    instagram: 'teste_insta',
    portfolio: '',
    linkedin: '',
    sobre: ''
  }, '');
}

// ============================================================
// Teste geral
// ============================================================

function testar() {
  var sheet = getSheet();
  if (sheet) {
    Logger.log('Planilha OK! Linhas: ' + sheet.getLastRow());
    Logger.log('URL: ' + sheet.getParent().getUrl());
  } else {
    Logger.log('Planilha não encontrada.');
  }
}

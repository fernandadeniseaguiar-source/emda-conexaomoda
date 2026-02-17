// ============================================================
// GOOGLE APPS SCRIPT — EMDA Banco de Talentos v3
// + Email automático de notificação
// + Salvar foto no Google Drive
// ============================================================

const SPREADSHEET_ID = '1oj57-yAspnZZGdjCGXQbDuySAofWHYJVnb9Rtn1cyCY';
const SHEET_NAME = 'Currículos';

// Email que receberá as notificações
const NOTIFY_EMAIL = 'fernandadeniseaguiar@gmail.com';

// Nome da pasta no Drive para salvar fotos
const FOTOS_FOLDER_NAME = 'EMDA - Fotos Currículos';

// ============================================================
// GET — Verificação de duplicatas
// ============================================================

function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === 'check') {
      return handleDuplicateCheck(e);
    }
    
    return jsonResponse({ status: 'ok', message: 'EMDA Banco de Talentos API v3' });
    
  } catch (error) {
    return jsonResponse({ error: error.message });
  }
}

function handleDuplicateCheck(e) {
  const field = e.parameter.field;
  const value = (e.parameter.value || '').trim().toLowerCase();
  
  if (!field || !value) {
    return jsonResponse({ found: false });
  }
  
  const sheet = getSheet();
  if (!sheet) {
    return jsonResponse({ found: false, error: 'Planilha não configurada' });
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const colIndex = headers.indexOf(field);
  if (colIndex === -1) {
    return jsonResponse({ found: false });
  }
  
  for (let i = 1; i < data.length; i++) {
    let cellValue = String(data[i][colIndex] || '').trim().toLowerCase();
    
    if (field === 'whatsapp') {
      cellValue = cellValue.replace(/\D/g, '');
      const searchValue = value.replace(/\D/g, '');
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
// POST — Salvar currículo + Email + Foto no Drive
// ============================================================

function doPost(e) {
  try {
    let data;
    
    // Aceitar dados via form (campo payload) ou JSON direto
    if (e.parameter && e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      return jsonResponse({ success: false, error: 'Sem dados recebidos' });
    }
    
    const sheet = getSheet();
    if (!sheet) {
      return jsonResponse({ success: false, error: 'Planilha não configurada' });
    }
    
    // Salvar foto no Google Drive (se existir)
    let fotoLink = '';
    if (data.foto_base64 && data.foto_base64.startsWith('data:image')) {
      fotoLink = salvarFotoNoDrive(data.foto_base64, data.nome);
    }
    
    // Salvar na planilha
    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.nome || '',
      data.email || '',
      data.whatsapp || '',
      data.cidade || '',
      data.estado || '',
      data.cursos || '',
      data.ano_conclusao || '',
      data.experiencia || '',
      data.instagram || '',
      data.portfolio || '',
      data.linkedin || '',
      data.sobre || '',
      data.foto || 'Não',
      fotoLink // Link da foto no Drive
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
    // Pegar ou criar pasta de fotos
    const folder = getOrCreateFolder(FOTOS_FOLDER_NAME);
    
    // Extrair tipo e dados do base64
    // formato: data:image/jpeg;base64,/9j/4AAQ...
    const parts = base64Data.split(',');
    const mimeMatch = parts[0].match(/data:(image\/\w+);base64/);
    
    if (!mimeMatch || !parts[1]) {
      return '';
    }
    
    const mimeType = mimeMatch[1];
    const extension = mimeType.split('/')[1].replace('jpeg', 'jpg');
    const imageData = Utilities.base64Decode(parts[1]);
    const blob = Utilities.newBlob(imageData, mimeType);
    
    // Nome do arquivo: NomeAluno_Data.extensão
    const timestamp = new Date().toISOString().slice(0, 10);
    const nomeClean = (nomeAluno || 'sem-nome').replace(/[^a-zA-ZÀ-ú\s]/g, '').replace(/\s+/g, '-').substring(0, 30);
    const fileName = `${nomeClean}_${timestamp}.${extension}`;
    
    blob.setName(fileName);
    
    // Salvar no Drive
    const file = folder.createFile(blob);
    
    // Tornar acessível via link
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
    
  } catch (error) {
    Logger.log('Erro ao salvar foto: ' + error.message);
    return 'Erro: ' + error.message;
  }
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
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
    var anoConclusao = (data && data.ano_conclusao) ? data.ano_conclusao : '-';
    var experiencia = (data && data.experiencia) ? data.experiencia : '';
    var instagram = (data && data.instagram) ? data.instagram : '';
    var portfolio = (data && data.portfolio) ? data.portfolio : '';
    var linkedin = (data && data.linkedin) ? data.linkedin : '';
    var sobre = (data && data.sobre) ? data.sobre : '';
    
    var assunto = '📋 Novo Currículo - ' + nome;
    
    var corpo = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;">';
    
    // Header
    corpo += '<div style="background:#000;padding:24px 32px;text-align:center;">';
    corpo += '<h1 style="color:#C9A962;font-size:18px;font-weight:400;letter-spacing:2px;margin:0;">BANCO DE TALENTOS</h1>';
    corpo += '<p style="color:rgba(255,255,255,0.5);font-size:11px;margin:4px 0 0 0;letter-spacing:1px;">ESCOLA DE MODA DENISE AGUIAR</p>';
    corpo += '</div>';
    
    // Body
    corpo += '<div style="padding:32px;">';
    corpo += '<p style="color:#666;font-size:13px;margin:0 0 24px 0;">Novo currículo cadastrado em <strong>' + dataHora + '</strong></p>';
    
    // Tabela de dados
    corpo += '<table style="width:100%;border-collapse:collapse;margin-bottom:24px;">';
    corpo += montarLinha('Nome', nome, true, true);
    corpo += montarLinha('Email', '<a href="mailto:' + email + '" style="color:#C9A962;text-decoration:none;">' + email + '</a>', false, true);
    corpo += montarLinha('WhatsApp', whatsapp, true, true);
    corpo += montarLinha('Cidade', cidade, false, true);
    corpo += montarLinha('Cursos', cursos, true, true);
    corpo += montarLinha('Conclusão', anoConclusao, false, true);
    
    if (experiencia) corpo += montarLinha('Experiência', experiencia, true, true);
    if (instagram) corpo += montarLinha('Instagram', '<a href="https://instagram.com/' + instagram + '" style="color:#C9A962;text-decoration:none;">@' + instagram + '</a>', false, true);
    if (portfolio) corpo += montarLinha('Portfólio', '<a href="' + portfolio + '" style="color:#C9A962;text-decoration:none;">' + portfolio + '</a>', true, true);
    if (linkedin) corpo += montarLinha('LinkedIn', '<a href="' + linkedin + '" style="color:#C9A962;text-decoration:none;">Ver perfil</a>', false, true);
    if (sobre) corpo += montarLinha('Sobre', sobre, true, true);
    if (fotoLink) corpo += montarLinha('Foto', '<a href="' + fotoLink + '" style="color:#C9A962;text-decoration:none;">📷 Ver foto</a>', false, true);
    
    corpo += '</table>';
    
    // Botão planilha
    corpo += '<div style="text-align:center;margin-top:24px;">';
    corpo += '<a href="https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit" ';
    corpo += 'style="display:inline-block;background:#000;color:#C9A962;padding:12px 32px;text-decoration:none;border-radius:8px;font-size:13px;font-weight:500;letter-spacing:0.5px;">Abrir Planilha Completa</a>';
    corpo += '</div>';
    
    corpo += '</div>';
    
    // Footer
    corpo += '<div style="background:#f8f7f5;padding:16px 32px;text-align:center;border-top:1px solid #eee;">';
    corpo += '<p style="color:#999;font-size:11px;margin:0;">Escola de Moda Denise Aguiar — Banco de Talentos</p>';
    corpo += '</div>';
    
    corpo += '</div>';
    
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

function montarLinha(label, valor, destacar, mostrar) {
  if (!mostrar) return '';
  var bg = destacar ? 'background:#f8f7f5;' : '';
  var html = '<tr>';
  html += '<td style="padding:10px 12px;' + bg + 'border-bottom:1px solid #eee;width:130px;color:#999;font-size:12px;text-transform:uppercase;letter-spacing:0.5px;vertical-align:top;">' + label + '</td>';
  html += '<td style="padding:10px 12px;' + bg + 'border-bottom:1px solid #eee;font-size:14px;color:#333;">' + valor + '</td>';
  html += '</tr>';
  return html;
}

// ============================================================
// Helpers
// ============================================================

function getSheet() {
  if (!SPREADSHEET_ID) return null;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// Teste
// ============================================================

function testar() {
  const sheet = getSheet();
  if (sheet) {
    Logger.log('Planilha OK! Linhas: ' + sheet.getLastRow());
    Logger.log('URL: ' + sheet.getParent().getUrl());
  } else {
    Logger.log('Planilha não encontrada.');
  }
}

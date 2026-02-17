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
    const data = JSON.parse(e.postData.contents);
    
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
    const cursos = data.cursos || 'Não informado';
    const cidade = data.cidade ? data.cidade + '/' + data.estado : 'Não informada';
    const dataHora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    
    const assunto = '📋 Novo Currículo - ' + (data.nome || 'Sem nome');
    
    const corpo = `
<div style="font-family: 'Helvetica Neue', Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #ffffff;">
  
  <!-- Header -->
  <div style="background: #000000; padding: 24px 32px; text-align: center;">
    <h1 style="color: #C9A962; font-size: 18px; font-weight: 400; letter-spacing: 2px; margin: 0;">
      BANCO DE TALENTOS
    </h1>
    <p style="color: rgba(255,255,255,0.5); font-size: 11px; margin: 4px 0 0 0; letter-spacing: 1px;">
      ESCOLA DE MODA DENISE AGUIAR
    </p>
  </div>
  
  <!-- Body -->
  <div style="padding: 32px;">
    
    <p style="color: #666; font-size: 13px; margin: 0 0 24px 0;">
      Novo currículo cadastrado em <strong>${dataHora}</strong>
    </p>
    
    <!-- Dados do Aluno -->
    <table style="width: 100%; border-collapse: collapse; margin-bottom: 24px;">
      <tr>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Nome</td>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; font-size: 14px; font-weight: 600; color: #333;">${data.nome || '-'}</td>
      </tr>
      <tr>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Email</td>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; font-size: 14px; color: #333;">
          <a href="mailto:${data.email}" style="color: #C9A962; text-decoration: none;">${data.email || '-'}</a>
        </td>
      </tr>
      <tr>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">WhatsApp</td>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; font-size: 14px; color: #333;">
          <a href="https://wa.me/55${(data.whatsapp || '').replace(/\\D/g, '')}" style="color: #C9A962; text-decoration: none;">${data.whatsapp || '-'}</a>
        </td>
      </tr>
      <tr>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Cidade</td>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; font-size: 14px; color: #333;">${cidade}</td>
      </tr>
      <tr>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Cursos</td>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; font-size: 14px; color: #333; font-weight: 500;">${cursos}</td>
      </tr>
      <tr>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Conclusão</td>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; font-size: 14px; color: #333;">${data.ano_conclusao || '-'}</td>
      </tr>
      ${data.experiencia ? `
      <tr>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; vertical-align: top;">Experiência</td>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; font-size: 13px; color: #333; line-height: 1.5;">${data.experiencia}</td>
      </tr>
      ` : ''}
      ${data.instagram ? `
      <tr>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Instagram</td>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; font-size: 14px; color: #333;">
          <a href="https://instagram.com/${data.instagram}" style="color: #C9A962; text-decoration: none;">@${data.instagram}</a>
        </td>
      </tr>
      ` : ''}
      ${data.portfolio ? `
      <tr>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Portfólio</td>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; font-size: 14px; color: #333;">
          <a href="${data.portfolio}" style="color: #C9A962; text-decoration: none;">${data.portfolio}</a>
        </td>
      </tr>
      ` : ''}
      ${data.linkedin ? `
      <tr>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">LinkedIn</td>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; font-size: 14px; color: #333;">
          <a href="${data.linkedin}" style="color: #C9A962; text-decoration: none;">Ver perfil</a>
        </td>
      </tr>
      ` : ''}
      ${data.sobre ? `
      <tr>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; vertical-align: top;">Sobre</td>
        <td style="padding: 10px 12px; background: #f8f7f5; border-bottom: 1px solid #eee; font-size: 13px; color: #333; line-height: 1.5;">${data.sobre}</td>
      </tr>
      ` : ''}
      ${fotoLink ? `
      <tr>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; width: 130px; color: #999; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Foto</td>
        <td style="padding: 10px 12px; border-bottom: 1px solid #eee; font-size: 14px; color: #333;">
          <a href="${fotoLink}" style="color: #C9A962; text-decoration: none;">📷 Ver foto</a>
        </td>
      </tr>
      ` : ''}
    </table>
    
    <!-- Link para planilha -->
    <div style="text-align: center; margin-top: 24px;">
      <a href="https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit" 
         style="display: inline-block; background: #000; color: #C9A962; padding: 12px 32px; text-decoration: none; border-radius: 8px; font-size: 13px; font-weight: 500; letter-spacing: 0.5px;">
        Abrir Planilha Completa
      </a>
    </div>
    
  </div>
  
  <!-- Footer -->
  <div style="background: #f8f7f5; padding: 16px 32px; text-align: center; border-top: 1px solid #eee;">
    <p style="color: #999; font-size: 11px; margin: 0;">
      Escola de Moda Denise Aguiar — Banco de Talentos
    </p>
  </div>
  
</div>`;
    
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

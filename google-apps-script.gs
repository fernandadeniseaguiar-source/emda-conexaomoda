// ============================================================
// GOOGLE APPS SCRIPT — EMDA Banco de Talentos
// Cole este código em https://script.google.com
// ============================================================
//
// INSTRUÇÕES:
// 1. Abra https://script.google.com e crie um novo projeto
// 2. Cole este código inteiro
// 3. Clique em "Implantar" → "Nova implantação"
// 4. Tipo: "App da Web"
// 5. Executar como: "Eu" 
// 6. Quem tem acesso: "Qualquer pessoa"
// 7. Copie a URL gerada e cole no CONFIG.GOOGLE_SCRIPT_URL do app.js
//
// PLANILHA:
// O script cria automaticamente uma planilha chamada "EMDA - Banco de Talentos"
// na sua conta do Google Drive na primeira execução.
// ============================================================

// ID da planilha (será preenchido automaticamente na primeira execução)
let SPREADSHEET_ID = '';

// Nome da aba
const SHEET_NAME = 'Currículos';

// ============================================================
// GET — Verificação de duplicatas
// URL: ?action=check&field=nome&value=João Silva
// ============================================================

function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === 'check') {
      return handleDuplicateCheck(e);
    }
    
    // Default: status
    return jsonResponse({ status: 'ok', message: 'EMDA Banco de Talentos API' });
    
  } catch (error) {
    return jsonResponse({ error: error.message }, 500);
  }
}

function handleDuplicateCheck(e) {
  const field = e.parameter.field; // 'nome', 'email', 'whatsapp'
  const value = (e.parameter.value || '').trim().toLowerCase();
  
  if (!field || !value) {
    return jsonResponse({ found: false, error: 'Campo e valor são obrigatórios' });
  }
  
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Mapear campos para colunas da planilha
  const fieldColumnMap = {
    'nome': 'nome',
    'email': 'email',
    'whatsapp': 'whatsapp'
  };
  
  const columnName = fieldColumnMap[field];
  if (!columnName) {
    return jsonResponse({ found: false, error: 'Campo inválido' });
  }
  
  const colIndex = headers.indexOf(columnName);
  if (colIndex === -1) {
    return jsonResponse({ found: false });
  }
  
  // Buscar nas linhas (pular header)
  for (let i = 1; i < data.length; i++) {
    let cellValue = String(data[i][colIndex] || '').trim().toLowerCase();
    
    // Para WhatsApp, comparar só números
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
      // Para nome e email, comparação case-insensitive
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
// POST — Salvar currículo
// ============================================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    const sheet = getOrCreateSheet();
    
    // Adicionar linha
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
      // Foto base64 fica muito grande para a planilha — salvar separadamente se necessário
      data.foto_base64 ? 'Sim (dados na célula)' : ''
    ]);
    
    return jsonResponse({ success: true, message: 'Currículo salvo com sucesso' });
    
  } catch (error) {
    return jsonResponse({ success: false, error: error.message }, 500);
  }
}

// ============================================================
// Helpers
// ============================================================

function getOrCreateSheet() {
  let ss;
  
  // Tentar abrir planilha existente
  if (SPREADSHEET_ID) {
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      // ID inválido, criar nova
      ss = null;
    }
  }
  
  // Se não existe, buscar pelo nome ou criar
  if (!ss) {
    const files = DriveApp.getFilesByName('EMDA - Banco de Talentos');
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create('EMDA - Banco de Talentos');
      
      // Criar headers
      const sheet = ss.getActiveSheet();
      sheet.setName(SHEET_NAME);
      sheet.appendRow([
        'timestamp', 'nome', 'email', 'whatsapp', 'cidade', 'estado',
        'cursos', 'ano_conclusao', 'experiencia', 'instagram', 
        'portfolio', 'linkedin', 'sobre', 'foto', 'foto_base64'
      ]);
      
      // Formatar header
      const headerRange = sheet.getRange(1, 1, 1, 15);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#000000');
      headerRange.setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      
      // Ajustar largura das colunas
      sheet.setColumnWidth(1, 180); // timestamp
      sheet.setColumnWidth(2, 200); // nome
      sheet.setColumnWidth(3, 200); // email
      sheet.setColumnWidth(4, 150); // whatsapp
      sheet.setColumnWidth(7, 300); // cursos
      sheet.setColumnWidth(9, 300); // experiencia
      sheet.setColumnWidth(13, 300); // sobre
      
      Logger.log('Planilha criada: ' + ss.getUrl());
    }
    
    // Salvar o ID para próximas execuções (coloque manualmente após primeira execução)
    // SPREADSHEET_ID = ss.getId();
  }
  
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getActiveSheet();
  }
  
  return sheet;
}

function jsonResponse(data, code) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// Teste manual — rode esta função para testar
// ============================================================

function testar() {
  const sheet = getOrCreateSheet();
  Logger.log('Planilha: ' + sheet.getParent().getUrl());
  Logger.log('Linhas: ' + sheet.getLastRow());
}

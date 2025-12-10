// backend/server.js - VERSÃO COM FONTES MAIORES NAS ETIQUETAS
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const { Readable } = require('stream');

const app = express();
const PORT = process.env.PORT || 3000;

// ✅ CORS SEGURO
const allowedOrigins = [
  'http://localhost:5500',
  'http://127.0.0.1:5500',
  'http://localhost:3000',
  'http://127.0.0.1:3000'
];

app.use(cors({
  origin: function(origin, callback) {
    // Permitir requisições sem origem (como mobile apps ou curl)
    if (!origin) return callback(null, true);
    
    if (allowedOrigins.indexOf(origin) === -1) {
      const msg = `Origem ${origin} não permitida pelo CORS`;
      console.warn('⚠️ Tentativa de acesso de origem não permitida:', origin);
      return callback(new Error(msg), false);
    }
    return callback(null, true);
  },
  credentials: true,
  methods: ['GET', 'POST', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

// Middlewares
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

// ✅ CONFIGURAÇÃO
const frontendPath = path.join(__dirname, '../frontend');
console.log('📁 Frontend path:', frontendPath);
app.use(express.static(frontendPath));

// ✅ CONFIGURAÇÃO DE UPLOAD
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { 
    fileSize: 50 * 1024 * 1024, // 50MB
    files: 1
  },
  fileFilter: (req, file, cb) => {
    // Validar tipo de arquivo
    const allowedTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel'
    ];
    
    if (allowedTypes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error('Apenas arquivos Excel (.xlsx, .xls) são permitidos'));
    }
  }
});

// ✅ GESTOR DE PLANILHAS COM LIMPEZA AUTOMÁTICA
class SpreadsheetManager {
  constructor() {
    this.data = new Map();
    this.timeouts = new Map();
    this.MAX_AGE = 60 * 60 * 1000; // 1 hora
  }

  set(id, spreadsheetData) {
    this.data.set(id, spreadsheetData);
    
    // Limpar automaticamente após 1 hora
    this.setCleanupTimeout(id);
  }

  setCleanupTimeout(id) {
    // Limpar timeout anterior se existir
    if (this.timeouts.has(id)) {
      clearTimeout(this.timeouts.get(id));
    }
    
    // Configurar novo timeout
    const timeout = setTimeout(() => {
      this.delete(id);
      console.log(`🗑️ Planilha ${id} removida por inatividade`);
    }, this.MAX_AGE);
    
    this.timeouts.set(id, timeout);
  }

  get(id) {
    const spreadsheet = this.data.get(id);
    if (spreadsheet) {
      // Resetar timeout quando acessado
      this.setCleanupTimeout(id);
    }
    return spreadsheet;
  }

  has(id) {
    return this.data.has(id);
  }

  delete(id) {
    if (this.timeouts.has(id)) {
      clearTimeout(this.timeouts.get(id));
      this.timeouts.delete(id);
    }
    return this.data.delete(id);
  }

  getAll() {
    return Array.from(this.data.entries()).map(([id, data]) => ({
      id,
      ...data
    }));
  }

  clearAll() {
    this.timeouts.forEach(timeout => clearTimeout(timeout));
    this.timeouts.clear();
    this.data.clear();
  }
}

const spreadsheetManager = new SpreadsheetManager();

// ========== FUNÇÕES AUXILIARES ==========

function getColumnLetter(columnNumber) {
  let dividend = columnNumber;
  let columnLetter = '';
  
  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    columnLetter = String.fromCharCode(65 + modulo) + columnLetter;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  
  return columnLetter || 'A';
}

function columnLetterToNumber(letter) {
  let number = 0;
  for (let i = 0; i < letter.length; i++) {
    number = number * 26 + (letter.charCodeAt(i) - 64);
  }
  return number;
}

// ✅ FUNÇÃO: Obter células mescladas CORRETAMENTE (compatível com todas versões do ExcelJS)
function getMergedCells(worksheet) {
  try {
    let mergedCells = [];
    
    // Versão 4.0+ do ExcelJS
    if (worksheet._merges && Array.isArray(worksheet._merges)) {
      mergedCells = worksheet._merges;
    } 
    // Versões mais antigas
    else if (worksheet.model && worksheet.model.merges) {
      mergedCells = worksheet.model.merges;
    } 
    // Outros formatos possíveis
    else if (worksheet.merges) {
      if (worksheet.merges instanceof Map) {
        mergedCells = Array.from(worksheet.merges.values());
      } else if (Array.isArray(worksheet.merges)) {
        mergedCells = worksheet.merges;
      }
    }
    
    // Converter para formato consistente
    const formattedMerges = [];
    
    for (const merge of mergedCells) {
      try {
        let top, left, bottom, right;
        
        if (typeof merge === 'string') {
          // Formato "A1:B2"
          const [topLeft, bottomRight] = merge.split(':');
          top = parseInt(topLeft.replace(/[A-Z]/g, ''));
          left = columnLetterToNumber(topLeft.replace(/\d/g, ''));
          bottom = parseInt(bottomRight.replace(/[A-Z]/g, ''));
          right = columnLetterToNumber(bottomRight.replace(/\d/g, ''));
        } 
        else if (merge && typeof merge === 'object') {
          // Formato objeto do ExcelJS
          if (merge.s && merge.e) {
            // { s: { r: 1, c: 1 }, e: { r: 2, c: 2 } }
            top = merge.s.r + 1; // ExcelJS usa base 0
            left = merge.s.c + 1;
            bottom = merge.e.r + 1;
            right = merge.e.c + 1;
          } 
          else if (merge.min && merge.max) {
            // { min: { row: 1, col: 1 }, max: { row: 2, col: 2 } }
            top = merge.min.row;
            left = merge.min.col;
            bottom = merge.max.row;
            right = merge.max.col;
          }
          else {
            // { top, left, bottom, right }
            top = merge.top || merge.row || merge.r;
            left = merge.left || merge.col || merge.c;
            bottom = merge.bottom || merge.row2 || merge.r2;
            right = merge.right || merge.col2 || merge.c2;
          }
        }
        
        if (top && left && bottom && right) {
          formattedMerges.push({
            top: parseInt(top),
            left: parseInt(left),
            bottom: parseInt(bottom),
            right: parseInt(right)
          });
        }
      } catch (error) {
        console.warn('⚠️ Erro ao formatar célula mesclada:', error.message);
      }
    }
    
    return formattedMerges;
    
  } catch (error) {
    console.warn('⚠️ Não foi possível obter células mescladas:', error.message);
    return [];
  }
}

// ✅ FUNÇÃO: Verificar se célula está em área mesclada (ROBUSTA)
function isCellInMergedArea(rowNum, colNum, mergedCells) {
  if (!mergedCells || !Array.isArray(mergedCells) || mergedCells.length === 0) {
    return false;
  }
  
  for (const merge of mergedCells) {
    if (!merge || typeof merge !== 'object') continue;
    
    if (rowNum >= merge.top && rowNum <= merge.bottom &&
        colNum >= merge.left && colNum <= merge.right) {
      return true;
    }
  }
  return false;
}

// ✅ FUNÇÃO: Verificar se linha tem células mescladas (ROBUSTA)
function hasMergedCellsInRow(rowNum, mergedCells) {
  if (!mergedCells || !Array.isArray(mergedCells) || mergedCells.length === 0) {
    return false;
  }
  
  for (const merge of mergedCells) {
    if (!merge || typeof merge !== 'object') continue;
    
    if (rowNum >= merge.top && rowNum <= merge.bottom) {
      return true;
    }
  }
  return false;
}

// ✅ FUNÇÃO: Obter valor de célula CORRETAMENTE
function getCellValue(cell) {
  try {
    if (!cell || cell.value === undefined || cell.value === null) {
      return null;
    }
    
    const value = cell.value;
    
    // Tratar diferentes tipos
    if (value instanceof Date) {
      return value.toLocaleDateString('pt-BR');
    }
    
    if (typeof value === 'object') {
      // Se for rich text
      if (value.richText) {
        return value.richText.map(rt => rt.text).join('');
      }
      // Se for fórmula
      if (cell.formula) {
        return cell.result !== undefined ? cell.result : String(value);
      }
      // Se for hyperlink
      if (value.text && value.hyperlink) {
        return value.text;
      }
      return String(value);
    }
    
    // Strings - preservar quebras de linha
    if (typeof value === 'string') {
      return value.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    }
    
    // Números
    if (typeof value === 'number') {
      // Verificar se é data serial do Excel
      if (cell.numFmt && (cell.numFmt.includes('d') || cell.numFmt.includes('m') || cell.numFmt.includes('y'))) {
        try {
          const date = ExcelJS.DateTime.fromExcelSerial(value);
          return date.toLocaleDateString('pt-BR');
        } catch (e) {
          return value;
        }
      }
      return value;
    }
    
    return value;
    
  } catch (error) {
    console.warn('⚠️ Erro ao obter valor da célula:', error.message);
    return null;
  }
}

// ✅ FUNÇÃO: Encontrar linha de cabeçalhos (INTELIGENTE)
function findHeaderRow(worksheet, mergedCells) {
  console.log('🔍 Procurando linha de cabeçalhos...');
  
  // Verificar primeiras 15 linhas
  for (let rowNum = 1; rowNum <= Math.min(15, worksheet.rowCount); rowNum++) {
    
    // Pular linhas com células mescladas (se houver informação de merges)
    if (mergedCells && mergedCells.length > 0 && hasMergedCellsInRow(rowNum, mergedCells)) {
      console.log(`   ⏭️  Linha ${rowNum} pulada (contém células mescladas)`);
      continue;
    }
    
    const row = worksheet.getRow(rowNum);
    let textCells = 0;
    let totalCells = 0;
    
    // Verificar primeiras 30 colunas
    for (let col = 1; col <= Math.min(30, worksheet.columnCount); col++) {
      
      // Pular células em áreas mescladas (se houver informação)
      if (mergedCells && mergedCells.length > 0 && isCellInMergedArea(rowNum, col, mergedCells)) {
        continue;
      }
      
      try {
        const cell = row.getCell(col);
        const value = getCellValue(cell);
        
        if (value !== null && value !== undefined && value !== '') {
          totalCells++;
          // Cabeçalhos geralmente são texto não numérico
          if (typeof value === 'string' && isNaN(value.replace(/\s/g, ''))) {
            textCells++;
          }
        }
      } catch (error) {
        // Ignorar erro
      }
    }
    
    // Se encontrou pelo menos 2 células de texto, é provavelmente cabeçalho
    if (textCells >= 2 && totalCells >= 2) {
      console.log(`✅ Linha ${rowNum} identificada como cabeçalhos`);
      console.log(`   📊 ${textCells} textos, ${totalCells} células válidas`);
      return rowNum;
    }
  }
  
  // Fallback: primeira linha sem merges ou linha 1
  console.log('⚠️  Usando linha 1 como cabeçalhos (fallback)');
  return 1;
}

// ✅ FUNÇÃO: Extrair nomes de colunas (ROBUSTA)
function extractColumnNames(worksheet, headerRowNum, mergedCells) {
  console.log(`📋 Extraindo nomes das colunas (linha ${headerRowNum})...`);
  
  const headerRow = worksheet.getRow(headerRowNum);
  const columns = [];
  
  // Encontrar última coluna com dados
  let lastColumn = 1;
  for (let col = 1; col <= Math.min(100, worksheet.columnCount); col++) {
    try {
      // Pular células em áreas mescladas
      if (mergedCells && mergedCells.length > 0 && isCellInMergedArea(headerRowNum, col, mergedCells)) {
        continue;
      }
      
      const cell = headerRow.getCell(col);
      const value = getCellValue(cell);
      
      if (value !== null && value !== undefined && value !== '') {
        lastColumn = col;
      }
    } catch (error) {
      // Ignorar erro
    }
  }
  
  console.log(`📊 Última coluna com dados: ${lastColumn}`);
  
  // Extrair nomes
  for (let col = 1; col <= lastColumn; col++) {
    try {
      // Pular células em áreas mescladas
      if (mergedCells && mergedCells.length > 0 && isCellInMergedArea(headerRowNum, col, mergedCells)) {
        console.log(`   ⏭️  Coluna ${col} ignorada (em área mesclada)`);
        continue;
      }
      
      const cell = headerRow.getCell(col);
      const cellValue = getCellValue(cell);
      
      let columnName = '';
      let isGenerated = false;
      
      if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
        columnName = String(cellValue).trim();
        // Limpar nome da coluna
        columnName = columnName
          .replace(/[\n\r\t]/g, ' ')
          .replace(/\s+/g, ' ')
          .replace(/[\\/:*?"<>|]/g, '')
          .trim();
      }
      
      // Se não tem nome válido
      if (!columnName || columnName === '' || columnName === 'null' || columnName === 'undefined') {
        isGenerated = true;
        columnName = `Coluna_${getColumnLetter(col)}`;
        console.log(`   📝 Coluna ${col}: "${columnName}" (nome gerado)`);
      } else {
        console.log(`   ✅ Coluna ${col}: "${columnName}"`);
      }
      
      columns.push({
        name: columnName,
        key: `col_${col}`,
        originalColumn: col,
        isGenerated: isGenerated,
        hasMultipleLines: false
      });
      
    } catch (error) {
      console.log(`   ❌ Coluna ${col}: Erro - ${error.message}`);
    }
  }
  
  return columns;
}

// ✅ FUNÇÃO PRINCIPAL: Processar Excel (COM FALLBACK)
async function processExcel(buffer, ignoreMerges = true) {
  try {
    console.log('\n' + '='.repeat(60));
    console.log('🚀 PROCESSANDO PLANILHA');
    console.log('='.repeat(60));
    
    const workbook = new ExcelJS.Workbook();
    const stream = new Readable();
    stream.push(buffer);
    stream.push(null);
    
    await workbook.xlsx.read(stream);
    const worksheet = workbook.worksheets[0];
    
    console.log(`📄 Planilha: ${worksheet.name}`);
    console.log(`📊 Dimensões: ${worksheet.rowCount} linhas × ${worksheet.columnCount} colunas`);
    
    // Tentar obter células mescladas
    const mergedCells = ignoreMerges ? getMergedCells(worksheet) : [];
    console.log(`🔗 ${mergedCells.length} áreas mescladas detectadas`);
    
    // Encontrar linha de cabeçalhos
    const headerRowNum = findHeaderRow(worksheet, mergedCells);
    
    // Extrair nomes das colunas
    const columns = extractColumnNames(worksheet, headerRowNum, mergedCells);
    
    if (columns.length === 0) {
      throw new Error('Nenhuma coluna válida encontrada!');
    }
    
    console.log(`📋 ${columns.length} colunas identificadas`);
    
    // Ler dados
    console.log('\n📄 LENDO DADOS...');
    const data = [];
    let rowsProcessed = 0;
    let rowsSkippedDueToMerges = 0;
    let multiLineCells = 0;
    
    // Limitar a 5000 linhas para performance
    const maxRows = Math.min(worksheet.rowCount, headerRowNum + 5000);
    
    for (let rowNum = headerRowNum + 1; rowNum <= maxRows; rowNum++) {
      
      // Pular linhas com células mescladas (se ignorar merges)
      if (ignoreMerges && mergedCells.length > 0 && hasMergedCellsInRow(rowNum, mergedCells)) {
        rowsSkippedDueToMerges++;
        continue;
      }
      
      const row = worksheet.getRow(rowNum);
      const rowData = { 
        _id: `row_${rowNum}`, 
        _rowNumber: rowNum, 
        _hasMultiLine: false 
      };
      let hasData = false;
      let rowHasMultiLine = false;
      
      for (const column of columns) {
        try {
          const colNum = column.originalColumn;
          
          // Pular células em áreas mescladas
          if (ignoreMerges && mergedCells.length > 0 && isCellInMergedArea(rowNum, colNum, mergedCells)) {
            rowData[column.key] = null;
            continue;
          }
          
          const cell = row.getCell(colNum);
          const cellValue = getCellValue(cell);
          
          rowData[column.key] = cellValue;
          
          if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
            hasData = true;
            
            // Verificar múltiplas linhas
            if (typeof cellValue === 'string' && cellValue.includes('\n')) {
              multiLineCells++;
              rowHasMultiLine = true;
              column.hasMultipleLines = true;
            }
          }
        } catch (error) {
          rowData[column.key] = null;
        }
      }
      
      if (hasData) {
        rowData._hasMultiLine = rowHasMultiLine;
        data.push(rowData);
        rowsProcessed++;
      }
      
      // Mostrar progresso a cada 100 linhas
      if (rowsProcessed % 100 === 0 && rowsProcessed > 0) {
        console.log(`   📊 ${rowsProcessed} linhas processadas...`);
      }
    }
    
    // Resultado final
    console.log('\n' + '='.repeat(60));
    console.log('✅ PROCESSAMENTO CONCLUÍDO!');
    console.log('='.repeat(60));
    console.log(`📊 ${data.length} linhas de dados válidas`);
    console.log(`📋 ${columns.length} colunas disponíveis`);
    
    if (rowsSkippedDueToMerges > 0) {
      console.log(`🚫 ${rowsSkippedDueToMerges} linhas ignoradas (com células mescladas)`);
    }
    
    if (multiLineCells > 0) {
      console.log(`🔤 ${multiLineCells} células com múltiplas linhas`);
    }
    
    return {
      columns,
      data,
      metadata: {
        totalRows: worksheet.rowCount,
        dataRows: data.length,
        rowsSkipped: rowsSkippedDueToMerges,
        headerRow: headerRowNum,
        mergedAreas: mergedCells.length,
        multiLineCells: multiLineCells,
        fileName: worksheet.name
      }
    };
    
  } catch (error) {
    console.error('\n❌ ERRO NO PROCESSAMENTO:', error.message);
    console.error('Stack trace:', error.stack);
    throw error;
  }
}

// ✅ FUNÇÃO: Gerar PDF COM FONTES MAIORES
async function generateLabelsPDF(data, selectedColumns, allColumns, copiesPerLabel = 1) {
  return new Promise((resolve, reject) => {
    try {
      const PDFDocument = require('pdfkit');
      
      // Tamanho AUMENTADO para acomodar fontes maiores (120mm x 160mm)
      const pageWidth = 120 * 2.83465;  // Aumentado para acomodar fontes maiores
      const pageHeight = 160 * 2.83465; // Aumentado para acomodar fontes maiores
      
      const doc = new PDFDocument({ 
        margin: 0, 
        size: [pageWidth, pageHeight],
        info: {
          Title: 'Etiquetas Geradas',
          Author: 'Sistema de Etiquetas',
          Creator: 'Excel to Labels System',
          CreationDate: new Date()
        }
      });
      
      const buffers = [];
      doc.on('data', buffer => buffers.push(buffer));
      doc.on('end', () => {
        const pdfData = Buffer.concat(buffers);
        resolve(pdfData);
      });
      
      doc.on('error', (error) => {
        reject(error);
      });
      
      const margin = 15;
      const labelWidth = pageWidth - (margin * 2);
      const labelHeight = pageHeight - (margin * 2);
      
      let labelCount = 0;
      const totalLabels = data.length * copiesPerLabel;
      
      // Processar cada linha
      for (const row of data) {
        for (let copyIndex = 0; copyIndex < copiesPerLabel; copyIndex++) {
          if (labelCount > 0) {
            doc.addPage();
          }
          
          labelCount++;
          
          // Desenhar etiqueta COM FONTES MAIORES
          drawLabel(doc, row, selectedColumns, allColumns, 
                   margin, margin, labelWidth, labelHeight, 
                   labelCount, totalLabels);
        }
      }

      doc.end();
      
    } catch (error) {
      console.error('❌ Erro ao gerar PDF:', error);
      reject(error);
    }
  });
}

// ✅ FUNÇÃO: Desenhar etiqueta COM FONTES MAIORES (MODIFICADA)
function drawLabel(doc, row, selectedColumns, allColumns, x, y, width, height, currentNumber, totalLabels) {
  const padding = 12;
  const leftX = x + padding;
  const textWidth = width - (padding * 2);
  let currentY = y + padding;

  // Borda da etiqueta
  doc.rect(x, y, width, height)
     .strokeColor('#4f46e5')
     .lineWidth(1.5)
     .stroke();
  
  // Cabeçalho da etiqueta - FONTE MAIOR
  doc.fontSize(11)  // Aumentado para melhor legibilidade
     .font('Helvetica-Bold')
     .fillColor('#4f46e5')
     .text(`ETIQUETA ${currentNumber}/${totalLabels}`, leftX, currentY, {
       width: textWidth,
       align: 'center',
       lineGap: 3
     });
  
  currentY += 16;  // Aumentado espaçamento
  
  // Linha divisória após cabeçalho
  doc.moveTo(leftX, currentY)
     .lineTo(leftX + textWidth, currentY)
     .strokeColor('#e2e8f0')
     .lineWidth(0.5)
     .stroke();
  
  currentY += 12;
  
  // Conteúdo da etiqueta COM FONTES MAIORES
  selectedColumns.forEach(columnKey => {
    const column = allColumns.find(col => col.key === columnKey);
    if (!column) return;
    
    const rawValue = row[columnKey];
    
    if (rawValue === null || rawValue === undefined || rawValue === '') {
      return;
    }
    
    let valueStr = String(rawValue).trim();
    
    // Nome da coluna - FONTE MAIOR
    doc.fontSize(10)  // Aumentado para melhor legibilidade
       .font('Helvetica-Bold')
       .fillColor('#333333')
       .text(`${column.name}:`, leftX, currentY, {
         width: textWidth,
         align: 'left',
         lineGap: 2
       });
    
    currentY += 12;
    
    // Valor (com suporte a múltiplas linhas) - FONTE BEM MAIOR
    if (valueStr.includes('\n')) {
      const lines = valueStr.split('\n');
      lines.forEach((line, index) => {
        if (line.trim()) {
          doc.font('Helvetica')
             .fontSize(13)  // FONTE BEM MAIOR para fácil leitura
             .fillColor('#000000')
             .text(line.trim(), leftX + 5, currentY, {
               width: textWidth - 10,
               align: 'left',
               lineGap: 4
             });
          currentY += 15;
        }
      });
    } else {
      doc.font('Helvetica')
         .fontSize(13)  // FONTE BEM MAIOR para fácil leitura
         .fillColor('#000000')
         .text(valueStr, leftX + 5, currentY, {
           width: textWidth - 10,
           align: 'left',
           lineGap: 4
         });
      
      // Calcular altura do texto para ajustar espaçamento
      const textHeight = doc.heightOfString(valueStr, {
        width: textWidth - 10,
        lineGap: 4
      });
      
      currentY += textHeight + 8;
    }
    
    currentY += 10;  // Espaço entre campos
  });
  
  // Linha divisória no final
  if (selectedColumns.length > 0) {
    currentY = Math.min(currentY, y + height - 15);
    doc.moveTo(leftX, currentY)
       .lineTo(leftX + textWidth, currentY)
       .strokeColor('#e2e8f0')
       .lineWidth(1)
       .stroke();
  }
}

// ========== ROTAS DA API ==========

// ✅ Rota de teste
app.get('/api/test', (req, res) => {
  res.json({ 
    success: true,
    message: 'Backend funcionando!',
    version: '2.1.0',
    timestamp: new Date().toISOString(),
    environment: process.env.NODE_ENV || 'development',
    features: 'Fontes maiores nas etiquetas'
  });
});

// ✅ Rota de saúde
app.get('/api/health', (req, res) => {
  res.json({
    success: true,
    status: 'healthy',
    uptime: process.uptime(),
    memory: process.memoryUsage(),
    spreadsheets: spreadsheetManager.data.size
  });
});

// ✅ Rota de upload
app.post('/api/upload', upload.single('spreadsheet'), async (req, res) => {
  console.log('\n📥 UPLOAD RECEBIDO');
  
  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        message: 'Nenhum arquivo enviado'
      });
    }

    console.log(`📁 Arquivo: ${req.file.originalname} (${(req.file.size / 1024 / 1024).toFixed(2)} MB)`);
    
    let processedData;
    let processingMethod = 'padrão';
    
    try {
      // Primeiro tenta processar ignorando merges
      processedData = await processExcel(req.file.buffer, true);
      processingMethod = 'ignorando células mescladas';
    } catch (error) {
      console.warn('⚠️ Erro ao processar com ignorar merges:', error.message);
      console.log('🔄 Tentando processamento simples...');
      
      // Se falhar, tenta processamento simplificado
      try {
        processedData = await processExcel(req.file.buffer, false);
        processingMethod = 'modo simples';
      } catch (simpleError) {
        console.error('❌ Falha no processamento simples:', simpleError.message);
        throw simpleError;
      }
    }
    
    const spreadsheetId = `spreadsheet_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    
    const spreadsheetData = {
      id: spreadsheetId,
      originalName: req.file.originalname,
      uploadDate: new Date().toISOString(),
      data: processedData.data,
      columns: processedData.columns,
      metadata: {
        ...processedData.metadata,
        processingMethod: processingMethod,
        fileSize: req.file.size
      }
    };

    spreadsheetManager.set(spreadsheetId, spreadsheetData);

    console.log(`✅ Upload concluído: ID ${spreadsheetId}`);
    console.log(`   📊 ${processedData.data.length} linhas, ${processedData.columns.length} colunas`);
    
    res.json({
      success: true,
      message: `Planilha processada com sucesso! (${processingMethod})`,
      details: {
        rows: processedData.data.length,
        columns: processedData.columns.length,
        rowsSkipped: processedData.metadata.rowsSkipped || 0,
        mergedAreas: processedData.metadata.mergedAreas || 0,
        multiLineCells: processedData.metadata.multiLineCells || 0
      },
      data: {
        id: spreadsheetData.id,
        fileName: spreadsheetData.originalName,
        records: spreadsheetData.data.length,
        columns: spreadsheetData.columns.map(c => ({
          name: c.name,
          key: c.key,
          isGenerated: c.isGenerated,
          hasMultipleLines: c.hasMultipleLines
        }))
      }
    });

  } catch (error) {
    console.error('❌ ERRO NO UPLOAD:', error.message);
    
    let errorMessage = 'Erro ao processar planilha';
    let statusCode = 500;
    
    if (error.message.includes('formato') || error.message.includes('Excel')) {
      errorMessage = 'Formato de arquivo inválido. Use .xlsx ou .xls';
      statusCode = 400;
    } else if (error.message.includes('tamanho')) {
      errorMessage = 'Arquivo muito grande. Tamanho máximo: 50MB';
      statusCode = 400;
    }
    
    res.status(statusCode).json({
      success: false,
      message: `${errorMessage}: ${error.message}`,
      error: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
});

// ✅ Rota: Listar planilhas
app.get('/api/spreadsheets', (req, res) => {
  try {
    const files = spreadsheetManager.getAll().map(fileData => ({
      id: fileData.id,
      fileName: fileData.originalName,
      uploadDate: fileData.uploadDate,
      records: fileData.data ? fileData.data.length : 0,
      columns: fileData.columns ? fileData.columns.length : 0,
      rowsSkipped: fileData.metadata?.rowsSkipped || 0,
      mergedAreas: fileData.metadata?.mergedAreas || 0,
      formattedDate: new Date(fileData.uploadDate).toLocaleDateString('pt-BR') + ' ' + 
                    new Date(fileData.uploadDate).toLocaleTimeString('pt-BR').slice(0, 5)
    }));
    
    res.json({ 
      success: true,
      spreadsheets: files,
      total: files.length
    });
    
  } catch (error) {
    console.error('❌ Erro ao listar planilhas:', error);
    res.status(500).json({
      success: false,
      message: 'Erro interno ao listar planilhas'
    });
  }
});

// ✅ Rota: Obter dados da planilha
app.get('/api/spreadsheets/:spreadsheetId/data', (req, res) => {
  try {
    const spreadsheetId = req.params.spreadsheetId;
    const limit = parseInt(req.query.limit) || 50;
    const page = parseInt(req.query.page) || 1;
    const offset = (page - 1) * limit;
    
    const spreadsheetData = spreadsheetManager.get(spreadsheetId);
    
    if (!spreadsheetData) {
      return res.status(404).json({
        success: false,
        message: 'Planilha não encontrada'
      });
    }

    const totalRecords = spreadsheetData.data.length;
    const totalPages = Math.ceil(totalRecords / limit);
    const paginatedData = spreadsheetData.data.slice(offset, offset + limit);

    const formattedData = paginatedData.map((row, index) => ({
      index: offset + index + 1,
      data: row,
      hasMultiLine: row._hasMultiLine || false,
      preview: spreadsheetData.columns
        .slice(0, 3)
        .map(col => {
          const value = row[col.key];
          if (value === null || value === undefined || value === '') {
            return `${col.name}: [VAZIO]`;
          }
          const strValue = String(value);
          return `${col.name}: ${strValue.length > 25 ? strValue.substring(0, 22) + '...' : strValue}`;
        })
        .join(' | ')
    }));

    res.json({
      success: true,
      data: formattedData,
      pagination: {
        page,
        limit,
        totalRecords,
        totalPages,
        hasNextPage: page < totalPages,
        hasPrevPage: page > 1
      },
      columns: spreadsheetData.columns.map(c => ({
        name: c.name,
        key: c.key,
        isGenerated: c.isGenerated,
        hasMultipleLines: c.hasMultipleLines || false
      })),
      metadata: {
        fileName: spreadsheetData.originalName,
        rowsSkipped: spreadsheetData.metadata?.rowsSkipped || 0,
        mergedAreas: spreadsheetData.metadata?.mergedAreas || 0,
        processingMethod: spreadsheetData.metadata?.processingMethod || 'padrão'
      }
    });

  } catch (error) {
    console.error('❌ Erro ao obter dados:', error);
    res.status(500).json({
      success: false,
      message: 'Erro ao obter dados da planilha'
    });
  }
});

// ✅ Rota: Gerar PDF
app.get('/api/generate-pdf/:spreadsheetId', async (req, res) => {
  try {
    const spreadsheetId = req.params.spreadsheetId;
    const selectedColumns = req.query.columns ? req.query.columns.split(',') : [];
    const copies = parseInt(req.query.copies) || 1;
    const selectedRows = req.query.rows ? req.query.rows.split(',').map(Number) : [];
    
    const spreadsheetData = spreadsheetManager.get(spreadsheetId);
    
    if (!spreadsheetData) {
      return res.status(404).json({
        success: false,
        message: 'Planilha não encontrada'
      });
    }

    if (selectedColumns.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'Selecione pelo menos uma coluna'
      });
    }

    // Filtrar dados se houver seleção de linhas
    let dataToUse = spreadsheetData.data;
    if (selectedRows.length > 0) {
      dataToUse = spreadsheetData.data.filter((row, index) => 
        selectedRows.includes(index)
      );
    }

    if (dataToUse.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'Nenhum dado disponível para gerar etiquetas'
      });
    }

    console.log(`📄 Gerando PDF: ${dataToUse.length} etiquetas, ${copies} cópias cada`);
    console.log(`📏 Tamanho da página: 120mm x 160mm`);
    console.log(`🔤 Tamanhos das fontes: Cabeçalho(11), Nome coluna(10), Valor(13)`);
    
    const pdfBuffer = await generateLabelsPDF(
      dataToUse, 
      selectedColumns, 
      spreadsheetData.columns,
      copies
    );
    
    const fileName = `etiquetas_${spreadsheetData.originalName.replace(/\.[^/.]+$/, "")}_${Date.now()}.pdf`;
    
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('Content-Length', pdfBuffer.length);
    res.setHeader('Cache-Control', 'no-store');
    
    res.send(pdfBuffer);

  } catch (error) {
    console.error('❌ ERRO AO GERAR PDF:', error);
    res.status(500).json({
      success: false,
      message: 'Erro ao gerar PDF: ' + error.message
    });
  }
});

// ✅ Rota: Excluir planilha
app.delete('/api/spreadsheets/:spreadsheetId', (req, res) => {
  try {
    const spreadsheetId = req.params.spreadsheetId;
    
    const spreadsheet = spreadsheetManager.get(spreadsheetId);
    if (!spreadsheet) {
      return res.status(404).json({
        success: false,
        message: 'Planilha não encontrada'
      });
    }

    const fileName = spreadsheet.originalName;
    const deleted = spreadsheetManager.delete(spreadsheetId);
    
    if (deleted) {
      console.log(`🗑️ Planilha excluída: ${fileName} (ID: ${spreadsheetId})`);
      res.json({
        success: true,
        message: `Planilha "${fileName}" excluída com sucesso!`
      });
    } else {
      throw new Error('Falha ao excluir planilha');
    }

  } catch (error) {
    console.error('❌ Erro ao excluir:', error);
    res.status(500).json({
      success: false,
      message: 'Erro ao excluir planilha'
    });
  }
});

// ✅ Rota de debug/informações
app.get('/api/debug/:spreadsheetId', (req, res) => {
  try {
    const spreadsheetId = req.params.spreadsheetId;
    const spreadsheetData = spreadsheetManager.get(spreadsheetId);
    
    if (!spreadsheetData) {
      return res.status(404).json({ 
        success: false,
        message: 'Planilha não encontrada' 
      });
    }
    
    res.json({
      success: true,
      fileName: spreadsheetData.originalName,
      uploadDate: spreadsheetData.uploadDate,
      records: spreadsheetData.data.length,
      columns: spreadsheetData.columns.length,
      sampleData: spreadsheetData.data.slice(0, 2),
      columnsPreview: spreadsheetData.columns.slice(0, 10).map(col => ({
        name: col.name,
        key: col.key,
        isGenerated: col.isGenerated,
        hasMultipleLines: col.hasMultipleLines
      })),
      metadata: spreadsheetData.metadata
    });
    
  } catch (error) {
    res.status(500).json({ 
      success: false,
      message: error.message 
    });
  }
});

// ✅ Limpar todas as planilhas (apenas para desenvolvimento)
app.delete('/api/clear-all', (req, res) => {
  if (process.env.NODE_ENV === 'production') {
    return res.status(403).json({
      success: false,
      message: 'Esta rota não está disponível em produção'
    });
  }
  
  const count = spreadsheetManager.data.size;
  spreadsheetManager.clearAll();
  
  res.json({
    success: true,
    message: `${count} planilhas removidas`,
    count
  });
});

// ========== ROTAS PARA PÁGINAS HTML ==========
app.get('/', (req, res) => {
  res.sendFile(path.join(frontendPath, 'index.html'));
});

app.get('/upload', (req, res) => {
  res.sendFile(path.join(frontendPath, 'index.html'));
});

app.get('/generate', (req, res) => {
  res.sendFile(path.join(frontendPath, 'index.html'));
});

// ✅ Rota catch-all para API
app.all('/api/*', (req, res) => {
  res.status(404).json({
    success: false,
    message: 'Rota da API não encontrada',
    path: req.path,
    method: req.method
  });
});

// ✅ Rota catch-all para frontend
app.get('*', (req, res) => {
  res.sendFile(path.join(frontendPath, 'index.html'));
});

// ✅ Middleware de erro global
app.use((error, req, res, next) => {
  console.error('💥 ERRO GLOBAL:', error);
  
  if (error instanceof multer.MulterError) {
    return res.status(400).json({
      success: false,
      message: `Erro no upload: ${error.message}`
    });
  }
  
  res.status(500).json({
    success: false,
    message: 'Erro interno no servidor',
    error: process.env.NODE_ENV === 'development' ? error.message : undefined
  });
});

// ========== INICIAR SERVIDOR ==========
app.listen(PORT, () => {
  console.log('\n' + '='.repeat(60));
  console.log('🚀 SISTEMA DE ETIQUETAS - VERSÃO 2.1');
  console.log('='.repeat(60));
  console.log('✅ Fontes maiores para melhor legibilidade');
  console.log('✅ Tamanho de página: 120mm x 160mm');
  console.log('✅ Tamanhos de fonte:');
  console.log('   • Cabeçalho: 11pt');
  console.log('   • Nome da coluna: 10pt');
  console.log('   • Valor: 13pt (BEM MAIOR)');
  console.log(`📍 Servidor rodando em: http://localhost:${PORT}`);
  console.log('📧 API Endpoints disponíveis:');
  console.log(`   • http://localhost:${PORT}/api/test`);
  console.log(`   • http://localhost:${PORT}/api/upload`);
  console.log(`   • http://localhost:${PORT}/api/generate-pdf`);
  console.log('='.repeat(60));
  
  // Limpar planilhas antigas ao iniciar
  spreadsheetManager.clearAll();
  console.log('🧹 Cache de planilhas limpo');
});

module.exports = app;
// backend/server.js - VERSÃO QUE REMOVE TOTALMENTE LINHAS COM CÉLULAS MESCLADAS
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const { Readable } = require('stream');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// ✅ CORS BÁSICO
app.use(cors());

// Middlewares
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// ✅ CONFIGURAÇÃO SIMPLES
const frontendPath = path.join(__dirname, '../frontend');
console.log('📁 Frontend path:', frontendPath);

// Servir arquivos estáticos do frontend
app.use(express.static(frontendPath));

// ✅ CONFIGURAÇÃO DE UPLOAD
const upload = multer({ storage: multer.memoryStorage() });

// ✅ VARIÁVEL GLOBAL PARA ARMAZENAR DADOS
let spreadsheetsData = new Map();

// ========== ROTAS DA API ==========

// Rota de teste
app.get('/api/test', (req, res) => {
  res.json({ 
    success: true,
    message: 'Backend funcionando!',
    timestamp: new Date().toISOString()
  });
});

// Rota de UPLOAD
app.post('/api/upload', upload.single('spreadsheet'), async (req, res) => {
  try {
    console.log('📥 Recebendo upload...');
    
    if (!req.file) {
      return res.status(400).json({
        success: false,
        message: 'Nenhum arquivo enviado'
      });
    }

    console.log('📥 Processando arquivo:', req.file.originalname);

    // Processar o Excel
    const processedData = await processExcelFromBuffer(req.file.buffer);

    // Gerar ID único
    const spreadsheetId = Date.now().toString();
    
    // Salvar dados em memória
    const tempData = {
      id: spreadsheetId,
      originalName: req.file.originalname,
      uploadDate: new Date().toISOString(),
      data: processedData.data,
      columns: processedData.columns
    };

    spreadsheetsData.set(spreadsheetId, tempData);

    console.log('✅ Arquivo processado:', processedData.data.length, 'registros (sem linhas mescladas)');

    res.json({
      success: true,
      message: `Planilha processada com sucesso! ${processedData.data.length} registros.`,
      data: {
        id: tempData.id,
        fileName: req.file.originalname,
        records: processedData.data.length,
        columns: processedData.columns
      }
    });

  } catch (error) {
    console.error('❌ Erro ao processar planilha:', error);
    res.status(500).json({
      success: false,
      message: 'Erro ao processar planilha: ' + error.message
    });
  }
});

// Rota para listar planilhas
app.get('/api/spreadsheets', (req, res) => {
  try {
    const files = Array.from(spreadsheetsData.values()).map(fileData => ({
      id: fileData.id,
      fileName: fileData.originalName,
      uploadDate: fileData.uploadDate,
      records: fileData.data ? fileData.data.length : 0,
      columns: fileData.columns || []
    }));
    
    res.json({ 
      success: true,
      spreadsheets: files 
    });
    
  } catch (error) {
    console.error('❌ Erro ao listar planilhas:', error);
    res.status(500).json({
      success: false,
      message: 'Erro interno ao listar planilhas'
    });
  }
});

// Rota para obter dados da planilha (com limite)
app.get('/api/spreadsheets/:spreadsheetId/data', (req, res) => {
  try {
    const spreadsheetId = req.params.spreadsheetId;
    const limit = parseInt(req.query.limit) || 50;

    const spreadsheetData = spreadsheetsData.get(spreadsheetId);
    
    if (!spreadsheetData) {
      return res.status(404).json({
        success: false,
        message: 'Planilha não encontrada'
      });
    }

    const limitedData = spreadsheetData.data.slice(0, limit).map((row, index) => ({
      index: index,
      data: row,
      preview: Object.keys(row)
        .filter(key => key !== '_id')
        .slice(0, 3)
        .map(key => {
          const column = spreadsheetData.columns.find(col => col.key === key);
          const columnName = column ? column.name : key;
          const value = formatValue(row[key]);
          return `${columnName}: ${value}`;
        })
        .join(' | ')
    }));

    res.json({
      success: true,
      data: limitedData,
      totalRecords: spreadsheetData.data.length,
      columns: spreadsheetData.columns
    });

  } catch (error) {
    console.error('❌ Erro ao obter dados da planilha:', error);
    res.status(500).json({
      success: false,
      message: 'Erro ao obter dados'
    });
  }
});

// ✅ ROTA PARA GERAR PDF COM FILTRO DE LINHAS
app.get('/api/generate-pdf/:spreadsheetId', async (req, res) => {
  try {
    const spreadsheetId = req.params.spreadsheetId;
    const selectedColumns = req.query.columns ? req.query.columns.split(',') : [];
    const copies = parseInt(req.query.copies) || 1;
    
    // ✅ Filtrar apenas linhas selecionadas
    let selectedRows = [];
    if (req.query.rows && req.query.rows !== '') {
      selectedRows = req.query.rows.split(',').map(num => {
        return parseInt(num.trim());
      }).filter(num => !isNaN(num));
    }

    console.log('=== GERANDO ETIQUETAS ===');
    console.log('Planilha ID:', spreadsheetId);
    console.log('Colunas selecionadas:', selectedColumns.length);
    console.log('Linhas selecionadas:', selectedRows.length > 0 ? selectedRows : 'TODAS');
    console.log('Cópias por etiqueta:', copies);

    const spreadsheetData = spreadsheetsData.get(spreadsheetId);
    
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

    const pdfBuffer = await generateLabelsPDF(
      spreadsheetData.data, 
      selectedColumns, 
      spreadsheetData.columns,
      copies,
      selectedRows
    );
    
    const fileName = `etiquetas-${Date.now()}.pdf`;
    
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.send(pdfBuffer);
    
    console.log(`✅ PDF "${fileName}" enviado para download`);

  } catch (error) {
    console.error('❌ Erro ao gerar PDF:', error);
    res.status(500).json({
      success: false,
      message: 'Erro ao gerar PDF: ' + error.message
    });
  }
});

// ========== FUNÇÕES PRINCIPAIS ==========

// ✅ FUNÇÃO QUE REMOVE TOTALMENTE LINHAS COM CÉLULAS MESCLADAS
async function processExcelFromBuffer(buffer) {
  try {
    const workbook = new ExcelJS.Workbook();
    const stream = new Readable();
    stream.push(buffer);
    stream.push(null);
    
    await workbook.xlsx.read(stream);
    const worksheet = workbook.worksheets[0];
    
    const data = [];
    const columns = [];

    console.log(`📊 Processando: ${worksheet.name} (${worksheet.rowCount} linhas, ${worksheet.columnCount} colunas)`);

    // ========== DETECTAR TODAS AS LINHAS COM CÉLULAS MESCLADAS ==========
    const rowsWithMergedCells = new Set();
    
    // Método 1: Verificar worksheet._merges (mais confiável)
    if (worksheet._merges && Array.isArray(worksheet._merges)) {
      console.log(`🔗 Encontradas ${worksheet._merges.length} áreas mescladas`);
      
      worksheet._merges.forEach(merge => {
        // Marcar TODAS as linhas dentro de cada área mesclada
        for (let row = merge.top; row <= merge.bottom; row++) {
          rowsWithMergedCells.add(row);
          console.log(`   🚫 Linha ${row} marcada para remoção (contém células mescladas)`);
        }
      });
    } else {
      console.log('ℹ️ Nenhuma área mesclada encontrada na propriedade _merges');
    }
    
    // Método 2: Verificar célula por célula (backup)
    console.log(`🔍 Varrendo células para verificar merges...`);
    const checkLimit = Math.min(1000, worksheet.rowCount);
    
    for (let row = 1; row <= checkLimit; row++) {
      for (let col = 1; col <= worksheet.columnCount; col++) {
        try {
          const cell = worksheet.getCell(row, col);
          
          // Verificar se célula está mesclada
          if (cell.isMerged || cell.type === 7) { // type 7 pode indicar célula mesclada
            if (!rowsWithMergedCells.has(row)) {
              rowsWithMergedCells.add(row);
              console.log(`   🚫 Linha ${row} tem célula mesclada em coluna ${col}`);
            }
            break; // Já sabemos que a linha tem merge
          }
        } catch (e) {
          // Ignorar erros
        }
      }
    }
    
    console.log(`🚫 Total de linhas com células mescladas: ${rowsWithMergedCells.size}`);

    // ========== ENCONTRAR LINHA DE CABEÇALHOS (SEM CÉLULAS MESCLADAS) ==========
    let headerRowNum = 1;
    let foundValidHeaderRow = false;
    
    // Procurar linha de cabeçalhos que NÃO tem células mescladas
    for (let row = 1; row <= Math.min(50, worksheet.rowCount); row++) {
      if (rowsWithMergedCells.has(row)) {
        console.log(`⏭️ Linha ${row} pulada para cabeçalhos (tem células mescladas)`);
        continue;
      }
      
      const rowData = worksheet.getRow(row);
      let headerCount = 0;
      
      // Contar quantas células nesta linha têm texto (potenciais cabeçalhos)
      for (let col = 1; col <= Math.min(20, worksheet.columnCount); col++) {
        const cell = rowData.getCell(col);
        const value = cell.value;
        
        if (value !== undefined && value !== null && 
            value.toString().trim() !== '' &&
            typeof value === 'string') {
          headerCount++;
        }
      }
      
      // Se tem pelo menos 2 células com texto, considera como linha de cabeçalhos
      if (headerCount >= 2) {
        headerRowNum = row;
        foundValidHeaderRow = true;
        console.log(`📍 Linha ${row} selecionada como cabeçalhos (${headerCount} cabeçalhos encontrados)`);
        break;
      }
    }
    
    if (!foundValidHeaderRow) {
      // Fallback: usar primeira linha sem merges
      for (let row = 1; row <= worksheet.rowCount; row++) {
        if (!rowsWithMergedCells.has(row)) {
          headerRowNum = row;
          console.log(`🔄 Usando linha ${row} como cabeçalhos (fallback)`);
          break;
        }
      }
    }

    // ========== LER CABEÇALHOS ==========
    console.log(`\n📋 Lendo cabeçalhos da linha ${headerRowNum}:`);
    const headerRow = worksheet.getRow(headerRowNum);
    
    for (let col = 1; col <= worksheet.columnCount; col++) {
      const cell = headerRow.getCell(col);
      const cellValue = cell.value;
      
      if (cellValue !== undefined && cellValue !== null) {
        const headerName = cellValue.toString().trim();
        
        if (headerName !== '') {
          columns.push({
            name: headerName,
            key: `col_${col}`,
            originalColumn: col
          });
          console.log(`✅ Coluna ${col}: "${headerName}"`);
        }
      }
    }

    console.log(`📊 Total de colunas encontradas: ${columns.length}`);
    
    if (columns.length === 0) {
      throw new Error('Nenhum cabeçalho válido encontrado');
    }

    // ========== LER DADOS (REMOVENDO TOTALMENTE LINHAS COM CÉLULAS MESCLADAS) ==========
    console.log(`\n📄 Lendo dados (REMOVENDO linhas com células mescladas)...`);
    let totalRowsRead = 0;
    let totalRowsSkipped = 0;
    
    // Começar da linha APÓS os cabeçalhos
    const startDataRow = headerRowNum + 1;
    
    for (let rowNum = startDataRow; rowNum <= worksheet.rowCount; rowNum++) {
      // ✅ REMOVER COMPLETAMENTE se a linha tem células mescladas
      if (rowsWithMergedCells.has(rowNum)) {
        totalRowsSkipped++;
        if (totalRowsSkipped <= 5) {
          console.log(`   🗑️ LINHA ${rowNum} REMOVIDA (contém células mescladas)`);
        }
        continue; // PULAR COMPLETAMENTE ESTA LINHA
      }
      
      const row = worksheet.getRow(rowNum);
      const rowData = { _id: `row_${rowNum}`, _originalRow: rowNum };
      let hasData = false;

      // Ler todas as colunas
      columns.forEach(colInfo => {
        const cellValue = row.getCell(colInfo.originalColumn).value;
        rowData[colInfo.key] = cellValue;
        
        if (cellValue !== undefined && cellValue !== null && 
            cellValue.toString().trim() !== '') {
          hasData = true;
        }
      });

      if (hasData) {
        data.push(rowData);
        totalRowsRead++;
        
        // Mostrar progresso
        if (totalRowsRead % 100 === 0) {
          console.log(`   📈 Processadas ${totalRowsRead} linhas válidas...`);
        }
      } else {
        totalRowsSkipped++;
      }
    }

    console.log(`\n🎉 PROCESSAMENTO FINALIZADO!`);
    console.log(`✅ Linhas válidas processadas: ${data.length}`);
    console.log(`🚫 Linhas REMOVIDAS (com células mescladas): ${totalRowsSkipped}`);
    console.log(`📋 Colunas: ${columns.length}`);
    
    // DEBUG: Mostrar exemplos das primeiras linhas processadas
    console.log('\n🔍 Primeiras 3 linhas processadas:');
    data.slice(0, 3).forEach((row, idx) => {
      console.log(`   Linha ${idx + 1} (original: ${row._originalRow}):`, 
        Object.keys(row)
          .filter(key => key.startsWith('col_'))
          .slice(0, 3)
          .map(key => {
            const col = columns.find(c => c.key === key);
            return col ? `${col.name}: ${formatValue(row[key])}` : '';
          })
          .filter(x => x)
          .join(' | ')
      );
    });
    
    return { columns, data };
    
  } catch (error) {
    console.error('❌ Erro no processExcelFromBuffer:', error);
    throw error;
  }
}

// ✅ FUNÇÃO AUXILIAR: Formatar valores
function formatValue(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) return value.toLocaleDateString('pt-BR');
  if (typeof value === 'string' && value.length > 50) return value.substring(0, 47) + '...';
  return String(value);
}

// ✅ FUNÇÃO: Gerar PDF apenas com linhas selecionadas
async function generateLabelsPDF(data, selectedColumns, allColumns, copiesPerLabel = 1, selectedRows = []) {
  return new Promise((resolve, reject) => {
    try {
      console.log('=== GERANDO PDF ===');
      console.log('Total de linhas disponíveis:', data.length);
      console.log('Linhas selecionadas para gerar:', selectedRows);
      console.log('Cópias por linha:', copiesPerLabel);
      
      // ✅ FILTRAR APENAS AS LINHAS SELECIONADAS
      let dataToProcess = data;
      
      if (selectedRows && selectedRows.length > 0) {
        dataToProcess = selectedRows
          .filter(index => index >= 0 && index < data.length)
          .map(index => data[index]);
        
        console.log(`✅ Filtrando: ${selectedRows.length} linhas selecionadas -> ${dataToProcess.length} linhas para processar`);
        
        if (dataToProcess.length === 0) {
          throw new Error('Nenhuma linha válida selecionada para gerar etiquetas');
        }
      } else {
        console.log('⚠️ Nenhuma linha específica selecionada. Gerando TODAS as linhas.');
      }
      
      const PDFDocument = require('pdfkit');
      
      const pageWidth = 110 * 2.83465;
      const pageHeight = pageWidth;
      
      const doc = new PDFDocument({ 
        margin: 0, 
        size: [pageWidth, pageHeight]
      });
      
      const buffers = [];
      doc.on('data', buffers.push.bind(buffers));
      doc.on('end', () => {
        const pdfData = Buffer.concat(buffers);
        resolve(pdfData);
      });

      const margin = 10;
      const labelWidth = pageWidth - (margin * 2);
      const labelHeight = pageHeight - (margin * 2);

      let labelCount = 0;
      const totalLabels = dataToProcess.length * copiesPerLabel;

      console.log(`📄 Gerando ${totalLabels} etiquetas...`);

      dataToProcess.forEach((row, rowIndex) => {
        for (let copyIndex = 0; copyIndex < copiesPerLabel; copyIndex++) {
          if (labelCount > 0) doc.addPage();
          
          labelCount++;
          drawSimpleLabel(doc, row, selectedColumns, allColumns, margin, margin, labelWidth, labelHeight, labelCount, totalLabels);
        }
      });

      doc.end();

      console.log(`✅ PDF gerado com sucesso! ${labelCount} etiquetas criadas.`);

    } catch (error) {
      console.error('❌ Erro ao gerar PDF:', error);
      reject(error);
    }
  });
}

// ✅ FUNÇÃO: Desenhar etiqueta
function drawSimpleLabel(doc, row, selectedColumns, allColumns, x, y, width, height, currentNumber, totalLabels) {
  // Borda
  doc.rect(x, y, width - 1, height - 1)
     .strokeColor('#cccccc')
     .lineWidth(1)
     .stroke();

  const padding = 20;
  const leftX = x + padding;
  const textWidth = width - (padding * 2);
  let currentY = y + padding;

  // Cabeçalho
  doc.fontSize(16)
     .font('Helvetica-Bold')
     .fillColor('#333333')
     .text(`ETIQUETA ${currentNumber}/${totalLabels}`, leftX, currentY, {
       width: textWidth,
       align: 'left'
     });
  
  currentY += 30;

  // Conteúdo
  doc.fontSize(12)
     .font('Helvetica')
     .fillColor('#000000');

  selectedColumns.forEach(columnKey => {
    const column = allColumns.find(col => col.key === columnKey);
    if (column) {
      const value = formatValue(row[columnKey]);
      if (value && value !== 'undefined' && value !== 'null') {
        const text = `${column.name}: ${value}`;
        if (currentY < y + height - 20) {
          doc.text(text, leftX, currentY, { width: textWidth, align: 'left' });
          currentY += 20;
        }
      }
    }
  });

  // Rodapé
  const footerY = y + height - 15;
  doc.fontSize(10)
     .fillColor('#666666')
     .text(`${currentNumber}/${totalLabels}`, x + width - padding - 10, footerY, {
       align: 'right'
     });
}

// ========== ROTAS PARA DEBUG ==========

// Rota para verificar exatamente quais linhas têm merges
app.post('/api/check-merged-rows', upload.single('spreadsheet'), async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);
    const worksheet = workbook.worksheets[0];
    
    const rowsWithMerges = new Set();
    const mergeDetails = [];
    
    // Detectar merges
    if (worksheet._merges && Array.isArray(worksheet._merges)) {
      worksheet._merges.forEach(merge => {
        const mergeInfo = {
          fromRow: merge.top,
          toRow: merge.bottom,
          fromCol: merge.left,
          toCol: merge.right,
          totalRows: merge.bottom - merge.top + 1,
          totalCols: merge.right - merge.left + 1
        };
        
        mergeDetails.push(mergeInfo);
        
        // Marcar todas as linhas nesta área mesclada
        for (let row = merge.top; row <= merge.bottom; row++) {
          rowsWithMerges.add(row);
        }
      });
    }
    
    // Amostra de linhas para ver o que será processado
    const sampleRows = [];
    const maxSample = Math.min(30, worksheet.rowCount);
    
    for (let row = 1; row <= maxSample; row++) {
      const rowData = worksheet.getRow(row);
      const hasMerge = rowsWithMerges.has(row);
      const sample = { 
        row, 
        hasMerge,
        status: hasMerge ? '🚫 SERÁ REMOVIDA' : '✅ SERÁ PROCESSADA'
      };
      
      // Pegar valores das primeiras 3 colunas
      for (let col = 1; col <= Math.min(3, worksheet.columnCount); col++) {
        sample[`col${col}`] = rowData.getCell(col).value || '';
      }
      
      sampleRows.push(sample);
    }
    
    res.json({
      fileName: req.file.originalname,
      totalRows: worksheet.rowCount,
      rowsWithMerges: Array.from(rowsWithMerges).sort((a,b) => a-b),
      mergeDetails: mergeDetails,
      sampleRows: sampleRows,
      summary: {
        totalRowsWithMerges: rowsWithMerges.size,
        willBeProcessed: worksheet.rowCount - rowsWithMerges.size,
        percentageRemoved: ((rowsWithMerges.size / worksheet.rowCount) * 100).toFixed(1) + '%'
      }
    });
    
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ========== ROTAS PARA AS PÁGINAS HTML ==========
app.get('/', (req, res) => {
  res.sendFile(path.join(frontendPath, 'index.html'));
});

app.get('/upload', (req, res) => {
  res.sendFile(path.join(frontendPath, 'index.html'));
});

app.get('/generate', (req, res) => {
  res.sendFile(path.join(frontendPath, 'index.html'));
});

// ✅ ROTA CATCH-ALL SEGURA PARA SPA
app.get('*', (req, res) => {
  if (req.path.startsWith('/api/')) {
    return res.status(404).json({
      success: false,
      message: 'Rota da API não encontrada'
    });
  }
  
  res.sendFile(path.join(frontendPath, 'index.html'));
});

// ========== INICIAR SERVIDOR ==========
app.listen(PORT, () => {
  console.log('===================================');
  console.log('🎉 SISTEMA DE ETIQUETAS RODANDO!');
  console.log('✅ Linhas com células mescladas serão REMOVIDAS COMPLETAMENTE');
  console.log(`📍 Local: http://localhost:${PORT}`);
  console.log('===================================');
});

module.exports = app;


// backend/server.js - VERSÃO COM FONTES GIGANTES ALINHADAS À ESQUERDA
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Middlewares
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../frontend')));

// Configurar upload com Multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = './uploads';
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir);
      console.log('📁 Pasta uploads criada');
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueName = Date.now() + '-' + Math.round(Math.random() * 1E9) + path.extname(file.originalname);
    cb(null, uniqueName);
  }
});

const upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    const allowedTypes = ['.xlsx', '.xls'];
    const fileExt = path.extname(file.originalname).toLowerCase();
    
    if (allowedTypes.includes(fileExt)) {
      cb(null, true);
    } else {
      cb(new Error('Apenas arquivos Excel (.xlsx, .xls) são permitidos!'), false);
    }
  },
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB
  }
});

// ========== ROTAS PARA PÁGINAS ==========
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '../frontend/index.html'));
});

app.get('/upload', (req, res) => {
  res.sendFile(path.join(__dirname, '../frontend/upload.html'));
});

app.get('/generate', (req, res) => {
  res.sendFile(path.join(__dirname, '../frontend/generate.html'));
});

// ========== API ROUTES ==========
app.get('/api/test', (req, res) => {
  res.json({ 
    status: 'OK', 
    message: 'Backend funcionando!',
    timestamp: new Date().toISOString()
  });
});

// Rota de UPLOAD REAL
app.post('/api/upload', upload.single('spreadsheet'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        message: 'Nenhum arquivo enviado'
      });
    }

    console.log('📥 Processando arquivo:', req.file.originalname);

    // Processar o Excel
    const processedData = await processExcel(req.file.path);

    // Salvar dados processados
    const tempData = {
      id: Date.now(),
      originalName: req.file.originalname,
      uploadDate: new Date().toISOString(),
      data: processedData.data,
      columns: processedData.columns,
      filePath: req.file.path
    };

    // Criar pasta temp se não existir
    if (!fs.existsSync('./temp')) {
      fs.mkdirSync('./temp');
      console.log('📁 Pasta temp criada');
    }

    // Salvar arquivo temporário
    const tempFilePath = `./temp/data-${tempData.id}.json`;
    fs.writeFileSync(tempFilePath, JSON.stringify(tempData, null, 2));

    console.log('✅ Arquivo processado:', processedData.data.length, 'registros');

    res.json({
      success: true,
      message: `✅ Planilha processada com sucesso! ${processedData.data.length} registros encontrados.`,
      data: {
        id: tempData.id,
        fileName: req.file.originalname,
        records: processedData.data.length,
        columns: processedData.columns,
        uploadDate: tempData.uploadDate
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

// Rota para listar planilhas processadas
app.get('/api/spreadsheets', (req, res) => {
  try {
    console.log('📋 Listando planilhas processadas...');
    const tempDir = './temp';
    
    if (!fs.existsSync(tempDir)) {
      console.log('ℹ️ Pasta temp não existe, retornando array vazio');
      return res.json({ 
        success: true, 
        spreadsheets: [] 
      });
    }

    const files = fs.readdirSync(tempDir)
      .filter(file => file.startsWith('data-') && file.endsWith('.json'))
      .map(file => {
        try {
          const filePath = path.join(tempDir, file);
          const fileData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
          
          return {
            id: fileData.id,
            fileName: fileData.originalName,
            uploadDate: fileData.uploadDate,
            records: fileData.data ? fileData.data.length : 0,
            columns: fileData.columns || []
          };
        } catch (error) {
          console.error('❌ Erro ao ler arquivo', file, error);
          return null;
        }
      })
      .filter(Boolean);

    console.log(`✅ ${files.length} planilhas encontradas`);
    
    res.json({ 
      success: true,
      spreadsheets: files 
    });
    
  } catch (error) {
    console.error('❌ Erro ao listar planilhas:', error);
    res.status(500).json({
      success: false,
      message: 'Erro interno ao listar planilhas: ' + error.message
    });
  }
});

// ✅ ROTA: Excluir planilha
app.delete('/api/spreadsheets/:spreadsheetId', (req, res) => {
  try {
    const spreadsheetId = req.params.spreadsheetId;
    console.log('🗑️  Solicitada exclusão da planilha:', spreadsheetId);

    const tempFilePath = `./temp/data-${spreadsheetId}.json`;
    let spreadsheetData = null;

    if (fs.existsSync(tempFilePath)) {
      spreadsheetData = JSON.parse(fs.readFileSync(tempFilePath, 'utf8'));
    }

    const deletionResults = [];

    if (fs.existsSync(tempFilePath)) {
      fs.unlinkSync(tempFilePath);
      deletionResults.push('✅ Arquivo temporário excluído');
      console.log('✅ Arquivo temporário excluído:', tempFilePath);
    } else {
      deletionResults.push('⚠️ Arquivo temporário não encontrado');
    }

    if (spreadsheetData && spreadsheetData.filePath && fs.existsSync(spreadsheetData.filePath)) {
      fs.unlinkSync(spreadsheetData.filePath);
      deletionResults.push('✅ Arquivo original excluído');
      console.log('✅ Arquivo original excluído:', spreadsheetData.filePath);
    } else {
      deletionResults.push('⚠️ Arquivo original não encontrado ou já excluído');
    }

    res.json({
      success: true,
      message: `Planilha excluída com sucesso!`,
      details: deletionResults,
      deletedId: spreadsheetId
    });

    console.log(`✅ Planilha ${spreadsheetId} excluída com sucesso`);

  } catch (error) {
    console.error('❌ Erro ao excluir planilha:', error);
    res.status(500).json({
      success: false,
      message: 'Erro ao excluir planilha: ' + error.message
    });
  }
});

// ✅ NOVA ROTA: Obter dados específicos da planilha para seleção
app.get('/api/spreadsheets/:spreadsheetId/data', (req, res) => {
  try {
    const spreadsheetId = req.params.spreadsheetId;
    const limit = parseInt(req.query.limit) || 50;
    
    console.log('📊 Solicitando dados da planilha:', spreadsheetId);

    const tempFilePath = `./temp/data-${spreadsheetId}.json`;
    
    if (!fs.existsSync(tempFilePath)) {
      return res.status(404).json({
        success: false,
        message: 'Planilha não encontrada'
      });
    }

    const spreadsheetData = JSON.parse(fs.readFileSync(tempFilePath, 'utf8'));
    
    if (!spreadsheetData.data || spreadsheetData.data.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'Planilha não contém dados'
      });
    }

    // Limitar quantidade de dados para performance
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
      columns: spreadsheetData.columns,
      limited: limitedData.length < spreadsheetData.data.length,
      message: limitedData.length < spreadsheetData.data.length 
        ? `Mostrando ${limitedData.length} de ${spreadsheetData.data.length} registros` 
        : `Mostrando todos os ${spreadsheetData.data.length} registros`
    });

  } catch (error) {
    console.error('❌ Erro ao obter dados da planilha:', error);
    res.status(500).json({
      success: false,
      message: 'Erro ao obter dados: ' + error.message
    });
  }
});

// ========== FUNÇÕES AUXILIARES ==========

async function processExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.worksheets[0];
  
  const data = [];
  const columns = [];

  worksheet.getRow(1).eachCell((cell, colNumber) => {
    columns.push({ 
      name: cell.value?.toString() || `Coluna ${colNumber}`, 
      key: `col_${colNumber}` 
    });
  });

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const rowData = { _id: `row_${rowNumber}` };
    row.eachCell((cell, colNumber) => {
      rowData[`col_${colNumber}`] = cell.value;
    });

    if (Object.values(rowData).some(val => val !== undefined && val !== null && val !== '')) {
      data.push(rowData);
    }
  });

  return { 
    columns: columns, 
    data: data, 
    totalRows: data.length,
    sheetName: worksheet.name 
  };
}

function formatValue(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  
  if (value instanceof Date) {
    return value.toLocaleDateString('pt-BR');
  }
  
  if (typeof value === 'number') {
    if (value.toString().includes('.') && value >= 1) {
      return 'R$ ' + value.toFixed(2).replace('.', ',');
    }
    return value.toString();
  }
  
  if (typeof value === 'string') {
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toLocaleDateString('pt-BR');
    }
    
    if (value.length > 50) {
      return value.substring(0, 47) + '...';
    }
    
    return value;
  }
  
  return String(value);
}

// ✅ FUNÇÃO ATUALIZADA: Obter configurações de estilo com FONTES GIGANTES ALINHADAS À ESQUERDA
function getLabelStyle(template) {
  const baseStyles = {
    simple: {
      background: null,
      borderColor: '#cccccc',
      borderWidth: 1,
      padding: 20, // ✅ MAIS ESPAÇO À ESQUERDA
      showHeader: true,
      headerHeight: 25,
      headerBackground: '#f8f9fa',
      headerFont: 'Helvetica-Bold',
      headerFontSize: 18, // ✅ FONTE GRANDE MAS NÃO GIGANTE
      headerColor: '#333333',
      bodyFont: 'Helvetica-Bold',
      bodyFontSize: 16, // ✅ FONTE GRANDE MAS NÃO GIGANTE
      textColor: '#000000',
      textLight: '#666666',
      primaryColor: '#333333',
      accentColor: '#e53e3e',
      footerColor: '#666666',
      dividerColor: '#dddddd',
      dividerWidth: 0.5,
      lineHeight: 20,
      lineSpacing: 8,
      highlightImportant: true,
      alignLeft: true // ✅ SEMPRE ALINHAR À ESQUERDA
    },
    modern: {
      background: '#ffffff',
      borderColor: '#6366f1',
      borderWidth: 1.5,
      padding: 20,
      showHeader: true,
      headerHeight: 28,
      headerBackground: '#6366f1',
      headerFont: 'Helvetica-Bold',
      headerFontSize: 18,
      headerColor: '#ffffff',
      bodyFont: 'Helvetica-Bold',
      bodyFontSize: 16,
      textColor: '#1f2937',
      textLight: '#6b7280',
      primaryColor: '#6366f1',
      accentColor: '#10b981',
      footerColor: '#9ca3af',
      dividerColor: '#e5e7eb',
      dividerWidth: 0.8,
      lineHeight: 20,
      lineSpacing: 8,
      highlightImportant: true,
      alignLeft: true
    },
    minimal: {
      background: '#ffffff',
      borderColor: '#e5e7eb',
      borderWidth: 0.5,
      padding: 20,
      showHeader: false,
      headerHeight: 0,
      headerBackground: null,
      headerFont: 'Helvetica-Bold',
      headerFontSize: 16,
      headerColor: '#374151',
      bodyFont: 'Helvetica-Bold',
      bodyFontSize: 16,
      textColor: '#374151',
      textLight: '#6b7280',
      primaryColor: '#374151',
      accentColor: '#059669',
      footerColor: '#9ca3af',
      dividerColor: '#f3f4f6',
      dividerWidth: 0.3,
      lineHeight: 20,
      lineSpacing: 8,
      highlightImportant: false,
      alignLeft: true
    },
    colorful: {
      background: '#f0f9ff',
      borderColor: '#0ea5e9',
      borderWidth: 1.5,
      padding: 20,
      showHeader: true,
      headerHeight: 26,
      headerBackground: '#0ea5e9',
      headerFont: 'Helvetica-Bold',
      headerFontSize: 18,
      headerColor: '#ffffff',
      bodyFont: 'Helvetica-Bold',
      bodyFontSize: 16,
      textColor: '#0c4a6e',
      textLight: '#0369a1',
      primaryColor: '#0ea5e9',
      accentColor: '#f59e0b',
      footerColor: '#38bdf8',
      dividerColor: '#bae6fd',
      dividerWidth: 0.6,
      lineHeight: 20,
      lineSpacing: 8,
      highlightImportant: true,
      alignLeft: true
    },
    // ✅ TEMPLATE SW OSAKA ALINHADO À ESQUERDA
    swosaka: {
      background: '#ffffff',
      borderColor: '#000000',
      borderWidth: 2,
      padding: 20,
      showHeader: false,
      headerHeight: 0,
      headerBackground: null,
      headerFont: 'Helvetica-Bold',
      headerFontSize: 20,
      headerColor: '#000000',
      bodyFont: 'Helvetica-Bold',
      bodyFontSize: 16,
      textColor: '#000000',
      textLight: '#000000',
      primaryColor: '#000000',
      accentColor: '#000000',
      footerColor: '#000000',
      dividerColor: '#000000',
      dividerWidth: 1.2,
      lineHeight: 22,
      lineSpacing: 10,
      highlightImportant: false,
      largeFontSize: 22,
      xlargeFontSize: 24,
      boldFont: 'Helvetica-Bold',
      normalFont: 'Helvetica-Bold',
      alignLeft: true
    }
  };

  return baseStyles[template] || baseStyles.simple;
}

// ✅ FUNÇÃO ATUALIZADA: Desenhar etiqueta SW OSAKA ALINHADA À ESQUERDA
function drawSWOsakaLabel(doc, row, selectedColumns, allColumns, x, y, width, height, currentNumber, totalLabels, styles) {
  const leftX = x + styles.padding; // ✅ SEMPRE À ESQUERDA
  const textWidth = width - (styles.padding * 2);
  
  doc.font(styles.boldFont)
     .fillColor(styles.textColor);

  // SW OSAKA - ALINHADO À ESQUERDA
  doc.fontSize(styles.xlargeFontSize)
     .text('SW OSAKA', leftX, y + 25, {
       width: textWidth,
       align: 'left'
     });

  // O-2743 - ALINHADO À ESQUERDA
  doc.fontSize(styles.largeFontSize)
     .text('O-2743', leftX, y + 55, {
       width: textWidth,
       align: 'left'
     });

  // Linha divisória - DE PONTA A PONTA
  doc.moveTo(x + 15, y + 85)
     .lineTo(x + width - 15, y + 85)
     .strokeColor(styles.dividerColor)
     .lineWidth(styles.dividerWidth)
     .stroke();

  // Pagina 5 - ALINHADO À ESQUERDA
  doc.fontSize(styles.bodyFontSize)
     .text('Pagina 5', leftX, y + 95);

  // 519 - ALINHADO À DIREITA (para contraste)
  doc.fontSize(styles.bodyFontSize)
     .text('519', x + width - 40, y + 95, {
       align: 'right'
     });

  // Linha divisória - DE PONTA A PONTA
  doc.moveTo(x + 15, y + 115)
     .lineTo(x + width - 15, y + 115)
     .strokeColor(styles.dividerColor)
     .lineWidth(styles.dividerWidth)
     .stroke();

  // NWA-4J71 - ALINHADO À ESQUERDA
  doc.fontSize(styles.xlargeFontSize)
     .text('NWA-4J71', leftX, y + 130, {
       width: textWidth,
       align: 'left'
     });

  // Numeração X/Y - ALINHADO À DIREITA (inferior)
  doc.fontSize(styles.bodyFontSize - 2)
     .font(styles.normalFont)
     .text(`${currentNumber}/${totalLabels}`, x + width - 30, y + height - 25, {
       align: 'right'
     });

  // Se houver colunas selecionadas, mostrar dados dinâmicos ALINHADOS À ESQUERDA
  if (selectedColumns.length > 0) {
    let dataY = y + 160;
    selectedColumns.forEach((columnKey, index) => {
      if (index < 3) { // ✅ AGORA CABEM ATÉ 3 CAMPOS
        const column = allColumns.find(col => col.key === columnKey);
        if (column) {
          const value = formatValue(row[columnKey]);
          if (value && value !== 'undefined' && value !== 'null') {
            doc.fontSize(styles.bodyFontSize - 2)
               .font(styles.normalFont)
               .text(`${column.name}: ${value}`, leftX, dataY, {
                 width: textWidth,
                 align: 'left'
               });
            dataY += 20;
          }
        }
      }
    });
  }
}

// ✅ FUNÇÃO COMPLETAMENTE REFEITA: Desenhar etiqueta com FONTES GRANDES ALINHADAS À ESQUERDA
function drawLabel(doc, row, selectedColumns, allColumns, x, y, width, height, currentNumber, totalLabels, template = 'simple') {
  
  const styles = getLabelStyle(template);
  
  if (styles.background) {
    doc.rect(x, y, width - 1, height - 1)
       .fill(styles.background);
  }

  doc.rect(x, y, width - 1, height - 1)
     .strokeColor(styles.borderColor)
     .lineWidth(styles.borderWidth)
     .stroke();

  // ✅ LAYOUT ESPECÍFICO PARA SW OSAKA
  if (template === 'swosaka') {
    drawSWOsakaLabel(doc, row, selectedColumns, allColumns, x, y, width, height, currentNumber, totalLabels, styles);
    return;
  }

  const leftX = x + styles.padding; // ✅ POSIÇÃO FIXA À ESQUERDA
  const textWidth = width - (styles.padding * 2);
  
  let currentY = y + styles.padding;

  if (styles.showHeader) {
    if (styles.headerBackground) {
      doc.rect(x, y, width - 1, styles.headerHeight)
         .fill(styles.headerBackground);
    }

    const headerText = `ETIQUETA ${currentNumber}/${totalLabels}`;
    doc.fontSize(styles.headerFontSize)
       .font(styles.headerFont)
       .fillColor(styles.headerColor)
       .text(headerText, leftX, currentY, {
         width: textWidth,
         align: 'left' // ✅ ALINHADO À ESQUERDA
       });
    
    currentY += styles.headerHeight + 8;
    
    // Linha divisória - MAIS CURTA para não cortar texto
    doc.moveTo(leftX, currentY - 3)
       .lineTo(leftX + (textWidth * 0.8), currentY - 3)
       .strokeColor(styles.dividerColor)
       .lineWidth(styles.dividerWidth)
       .stroke();
    
    currentY += 5;
  } else {
    // Header alternativo ALINHADO À ESQUERDA
    doc.fontSize(16)
       .font('Helvetica-Bold')
       .fillColor(styles.primaryColor)
       .text(`${currentNumber}/${totalLabels}`, leftX, currentY, {
         width: textWidth,
         align: 'left' // ✅ ALINHADO À ESQUERDA
       });
    
    currentY += 22;
  }

  doc.font(styles.bodyFont)
     .fontSize(styles.bodyFontSize)
     .fillColor(styles.textColor);

  // ✅ CONTEÚDO PRINCIPAL SEMPRE À ESQUERDA
  selectedColumns.forEach(columnKey => {
    const column = allColumns.find(col => col.key === columnKey);
    if (column) {
      const value = formatValue(row[columnKey]);
      
      if (value && value !== 'undefined' && value !== 'null') {
        const text = `${column.name}: ${value}`;
        
        // Verificar se cabe na etiqueta
        if (currentY < y + height - 30) {
          if (styles.highlightImportant && (column.name.toLowerCase().includes('preço') || column.name.toLowerCase().includes('valor'))) {
            doc.font('Helvetica-Bold')
               .fillColor(styles.accentColor);
          }
          
          // ✅ TEXTO SEMPRE ALINHADO À ESQUERDA
          doc.text(text, leftX, currentY, {
            width: textWidth,
            align: 'left' // ✅ ALINHADO À ESQUERDA
          });
          
          doc.font(styles.bodyFont)
             .fillColor(styles.textColor);
          
          currentY += styles.lineHeight;
        }
      }
    }
  });

  // Footer ALINHADO À DIREITA (para não atrapalhar o conteúdo)
  const footerY = y + height - 20;
  doc.fontSize(10)
     .font('Helvetica')
     .fillColor(styles.footerColor)
     .text(`${currentNumber}/${totalLabels}`, x + width - styles.padding - 10, footerY, {
       align: 'right'
     });

  if (selectedColumns.length === 0) {
    // Mensagem ALINHADA À ESQUERDA quando não há colunas
    doc.fontSize(14)
       .fillColor(styles.textLight)
       .text('Nenhuma coluna selecionada', leftX, currentY, {
         width: textWidth,
         align: 'left' // ✅ ALINHADO À ESQUERDA
       });
  }
}

// ✅ FUNÇÃO ATUALIZADA: Gerar PDF com FONTES GRANDES ALINHADAS À ESQUERDA
async function generateLabelsPDF(data, selectedColumns, allColumns, template = 'simple', copiesPerLabel = 1, selectedRows = null) {
  return new Promise((resolve, reject) => {
    try {
      const PDFDocument = require('pdfkit');
      
      const pageWidth = 110 * 2.83465;
      const pageHeight = pageWidth;
      
      const doc = new PDFDocument({ 
        margin: 0, 
        size: [pageWidth, pageHeight],
        layout: 'portrait'
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

      console.log(`📐 Dimensões da página: ${pageWidth.toFixed(1)}pt × ${pageHeight.toFixed(1)}pt (110mm × 110mm)`);
      console.log(`📐 Dimensões da etiqueta: ${labelWidth.toFixed(1)}pt × ${labelHeight.toFixed(1)}pt`);
      console.log(`🔠 FONTES: GRANDES E ALINHADAS À ESQUERDA`);

      let globalLabelCount = 0;

      let filteredData = data;
      if (selectedRows && selectedRows.length > 0) {
        filteredData = data.filter((row, index) => selectedRows.includes(index));
        console.log(`📋 Filtrados ${filteredData.length} de ${data.length} registros`);
      }

      const totalLabels = filteredData.length * copiesPerLabel;

      console.log(`📊 Gerando ${totalLabels} etiquetas (${filteredData.length} registros × ${copiesPerLabel} cópias cada)`);

      filteredData.forEach((row, rowIndex) => {
        for (let copyIndex = 0; copyIndex < copiesPerLabel; copyIndex++) {
          const currentNumber = globalLabelCount + 1;
          
          if (globalLabelCount > 0) {
            doc.addPage();
          }
          
          const x = margin;
          const y = margin;

          drawLabel(doc, row, selectedColumns, allColumns, x, y, labelWidth, labelHeight, currentNumber, totalLabels, template);
          globalLabelCount++;
        }
      });

      doc.end();

    } catch (error) {
      reject(error);
    }
  });
}

// ========== ROTAS DE PDF ==========

app.get('/api/generate-pdf/:spreadsheetId', async (req, res) => {
  try {
    const spreadsheetId = req.params.spreadsheetId;
    const selectedColumns = req.query.columns ? req.query.columns.split(',') : [];
    const template = req.query.template || 'simple';
    const copies = parseInt(req.query.copies) || 1;
    const selectedRows = req.query.rows ? req.query.rows.split(',').map(Number) : null;
    
    console.log('📄 Gerando PDF para planilha:', spreadsheetId);
    console.log('📋 Colunas selecionadas:', selectedColumns);
    console.log('🎨 Template selecionado:', template);
    console.log('🔢 Cópias por registro:', copies);
    console.log('📝 Registros selecionados:', selectedRows || 'Todos');

    const tempFilePath = `./temp/data-${spreadsheetId}.json`;
    
    if (!fs.existsSync(tempFilePath)) {
      return res.status(404).json({
        success: false,
        message: 'Planilha não encontrada'
      });
    }

    const spreadsheetData = JSON.parse(fs.readFileSync(tempFilePath, 'utf8'));
    
    if (!spreadsheetData.data || spreadsheetData.data.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'Planilha não contém dados para gerar etiquetas'
      });
    }

    if (copies < 1 || copies > 1000) {
      return res.status(400).json({
        success: false,
        message: 'Número de cópias deve estar entre 1 e 1000'
      });
    }

    if (selectedRows) {
      const invalidRows = selectedRows.filter(rowIndex => rowIndex < 0 || rowIndex >= spreadsheetData.data.length);
      if (invalidRows.length > 0) {
        return res.status(400).json({
          success: false,
          message: `Índices de linha inválidos: ${invalidRows.join(', ')}. A planilha tem ${spreadsheetData.data.length} registros (índices 0 a ${spreadsheetData.data.length - 1}).`
        });
      }
    }

    const totalRecords = selectedRows ? selectedRows.length : spreadsheetData.data.length;
    const totalEtiquetas = totalRecords * copies;

    console.log(`📊 Total de etiquetas a gerar: ${totalEtiquetas} (${totalRecords} registros × ${copies} cópias)`);
    console.log(`📐 Formato: 1 ETIQUETA POR PÁGINA de 110mm × 110mm`);
    console.log(`🔠 FONTES: GRANDES E ALINHADAS À ESQUERDA`);

    const pdfBuffer = await generateLabelsPDF(
      spreadsheetData.data, 
      selectedColumns, 
      spreadsheetData.columns,
      template,
      copies,
      selectedRows
    );
    
    const fileName = selectedRows 
      ? `etiquetas-ESQUERDA-${spreadsheetData.originalName.replace('.xlsx', '').replace('.xls', '')}-${selectedRows.length}registros-${copies}x.pdf`
      : `etiquetas-ESQUERDA-${spreadsheetData.originalName.replace('.xlsx', '').replace('.xls', '')}-${copies}x.pdf`;
    
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    
    res.send(pdfBuffer);
    
    console.log('✅ PDF gerado com sucesso!');
    console.log(`🏷️ ${totalEtiquetas} etiquetas geradas em páginas de 110mm`);
    console.log(`🔠 FONTES GRANDES E ALINHADAS À ESQUERDA aplicadas com sucesso!`);

  } catch (error) {
    console.error('❌ Erro ao gerar PDF:', error);
    res.status(500).json({
      success: false,
      message: 'Erro ao gerar PDF: ' + error.message
    });
  }
});

// ✅ ROTA DE TESTE: PDF com fontes ALINHADAS À ESQUERDA
app.get('/api/test-pdf', (req, res) => {
  try {
    const PDFDocument = require('pdfkit');
    
    const pageWidth = 110 * 2.83465;
    const pageHeight = pageWidth;
    
    const doc = new PDFDocument({ 
      margin: 0, 
      size: [pageWidth, pageHeight],
      layout: 'portrait'
    });
    
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="teste-etiquetas-ESQUERDA.pdf"');
    
    doc.pipe(res);
    
    const leftX = 20;
    
    doc.fontSize(18)
       .text('🎉 SISTEMA DE ETIQUETAS', leftX, 30, {
         align: 'left'
       });
    
    doc.fontSize(14)
       .text('FONTES GRANDES ALINHADAS À ESQUERDA', leftX, 60, {
         align: 'left'
       });
    
    doc.fontSize(12)
       .text(`Gerado em: ${new Date().toLocaleString('pt-BR')}`, leftX, 90, {
         align: 'left'
       });
    
    doc.fontSize(11)
       .text('✅ Funcionalidades:', leftX, 120, { align: 'left' })
       .text('   • Páginas de 110mm × 110mm', leftX, 140, { align: 'left' })
       .text('   • 1 etiqueta por página', leftX, 160, { align: 'left' })
       .text('   • FONTES GRANDES', leftX, 180, { align: 'left' })
       .text('   • TEXTO ALINHADO À ESQUERDA', leftX, 200, { align: 'left' })
       .text('   • Nenhum corte de texto', leftX, 220, { align: 'left' });
    
    doc.end();
    
    console.log('✅ PDF de teste ALINHADO À ESQUERDA gerado com sucesso!');
    
  } catch (error) {
    console.error('❌ Erro ao gerar PDF de teste:', error);
    res.status(500).json({ 
      success: false,
      error: 'Erro ao gerar PDF de teste: ' + error.message 
    });
  }
});

// Rota para debug
app.get('/api/debug/files', (req, res) => {
  try {
    const tempDir = './temp';
    const uploadDir = './uploads';
    
    const debugInfo = {
      tempExists: fs.existsSync(tempDir),
      uploadExists: fs.existsSync(uploadDir),
      tempFiles: fs.existsSync(tempDir) ? fs.readdirSync(tempDir) : [],
      uploadFiles: fs.existsSync(uploadDir) ? fs.readdirSync(uploadDir) : [],
      serverTime: new Date().toISOString()
    };
    
    console.log('🔍 Debug info:', debugInfo);
    res.json(debugInfo);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Rota para limpar arquivos temporários
app.delete('/api/cleanup', (req, res) => {
  try {
    console.log('🧹 Iniciando limpeza de arquivos temporários...');
    
    const tempDir = './temp';
    const uploadDir = './uploads';
    
    let deletedFiles = 0;
    
    if (fs.existsSync(tempDir)) {
      const tempFiles = fs.readdirSync(tempDir);
      tempFiles.forEach(file => {
        const filePath = path.join(tempDir, file);
        fs.unlinkSync(filePath);
        deletedFiles++;
      });
      console.log(`✅ ${tempFiles.length} arquivos temporários excluídos`);
    }
    
    if (fs.existsSync(uploadDir)) {
      const uploadFiles = fs.readdirSync(uploadDir);
      uploadFiles.forEach(file => {
        const filePath = path.join(uploadDir, file);
        fs.unlinkSync(filePath);
        deletedFiles++;
      });
      console.log(`✅ ${uploadFiles.length} arquivos de upload excluídos`);
    }
    
    res.json({
      success: true,
      message: `Limpeza concluída! ${deletedFiles} arquivos excluídos.`,
      deletedFiles: deletedFiles
    });
    
  } catch (error) {
    console.error('❌ Erro na limpeza:', error);
    res.status(500).json({
      success: false,
      message: 'Erro na limpeza: ' + error.message
    });
  }
});

// Iniciar servidor
app.listen(PORT, () => {
  console.log('============================================');
  console.log('🎉 SISTEMA DE ETIQUETAS RODANDO!');
  console.log('🔠 VERSÃO COM FONTES ALINHADAS À ESQUERDA!');
  console.log('📍 Início: http://localhost:' + PORT);
  console.log('📍 Upload: http://localhost:' + PORT + '/upload');
  console.log('📍 Gerar: http://localhost:' + PORT + '/generate');
  console.log('📍 API Test: http://localhost:' + PORT + '/api/test');
  console.log('============================================');
  console.log('📋 FUNCIONALIDADES DISPONÍVEIS:');
  console.log('   ✅ Upload de planilhas Excel (.xlsx, .xls)');
  console.log('   ✅ 5 templates de etiquetas diferentes');
  console.log('   ✅ ETIQUETAS EM PÁGINAS DE 110mm × 110mm');
  console.log('   ✅ 1 etiqueta grande por página');
  console.log('   ✅ Template SW OSAKA (layout específico)');
  console.log('   ✅ Seleção de registros específicos');
  console.log('   ✅ Múltiplas cópias por registro (1-1000)');
  console.log('   ✅ 🔠 FONTES GRANDES para máxima legibilidade');
  console.log('   ✅ ✅ TEXTO ALINHADO À ESQUERDA');
  console.log('   ✅ ✅ NENHUM CORTE DE TEXTO');
  console.log('============================================');
  console.log('🔠 CONFIGURAÇÃO DAS FONTES:');
  console.log('   • Simple: Body 16pt, Header 18pt');
  console.log('   • Modern: Body 16pt, Header 18pt');
  console.log('   • Minimal: Body 16pt');
  console.log('   • Colorful: Body 16pt, Header 18pt');
  console.log('   • SW OSAKA: Principais 22-24pt');
  console.log('   • ✅ TODO TEXTO ALINHADO À ESQUERDA');
  console.log('   • ✅ MARGEM AMPLA PARA EVITAR CORTES');
  console.log('   • ✅ NENHUMA LETRA CORTADA');
  console.log('============================================');
});
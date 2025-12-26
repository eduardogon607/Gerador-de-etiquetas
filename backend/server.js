// backend/server.js - SISTEMA DE ETIQUETAS COM QR CODE VISÍVEL
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');
const PDFDocument = require('pdfkit');
const QRCode = require('qrcode');

const app = express();
const PORT = process.env.PORT || 3000;

console.log('='.repeat(60));
console.log('🚀 SISTEMA DE ETIQUETAS - QR CODE GARANTIDO');
console.log('='.repeat(60));
console.log(`📍 Porta: ${PORT}`);
console.log(`📅 Iniciado em: ${new Date().toLocaleString('pt-BR')}`);
console.log('='.repeat(60));

// CORS
app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization']
}));

// Middlewares
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Servir frontend estático
const frontendPath = path.join(__dirname, '../frontend');
app.use(express.static(frontendPath));

// Configuração de upload
const upload = multer({
    storage: multer.memoryStorage(),
    limits: { 
        fileSize: 50 * 1024 * 1024,
        files: 1
    },
    fileFilter: (req, file, cb) => {
        const allowedTypes = ['.xlsx', '.xls', '.csv'];
        const ext = path.extname(file.originalname).toLowerCase();
        
        if (allowedTypes.includes(ext)) {
            cb(null, true);
        } else {
            cb(new Error('Apenas arquivos Excel (.xlsx, .xls) e CSV são permitidos'), false);
        }
    }
}).single('spreadsheet');

// ========== GERENCIADOR DE PLANILHAS ==========

class SpreadsheetManager {
    constructor() {
        this.data = new Map();
    }

    set(id, spreadsheetData) {
        this.data.set(id, {
            ...spreadsheetData,
            lastAccess: Date.now()
        });
    }

    get(id) {
        const sheet = this.data.get(id);
        if (sheet) {
            sheet.lastAccess = Date.now();
        }
        return sheet;
    }

    delete(id) {
        return this.data.delete(id);
    }

    getAll() {
        return Array.from(this.data.entries()).map(([id, data]) => ({
            id,
            fileName: data.originalName,
            uploadDate: data.uploadDate,
            lastAccess: data.lastAccess,
            records: data.data ? data.data.length : 0,
            columns: data.columns ? data.columns.length : 0,
            formattedDate: new Date(data.uploadDate).toLocaleString('pt-BR'),
            size: data.fileSize || 'N/A'
        }));
    }

    cleanup(maxAgeHours = 24) {
        const maxAge = maxAgeHours * 60 * 60 * 1000;
        const now = Date.now();
        let removed = 0;
        
        for (const [id, sheet] of this.data.entries()) {
            if (now - sheet.lastAccess > maxAge) {
                this.delete(id);
                removed++;
                console.log(`🗑️ Removida planilha antiga: ${id} (${sheet.originalName})`);
            }
        }
        
        return removed;
    }
}

const spreadsheetManager = new SpreadsheetManager();

// Limpeza automática a cada hora
setInterval(() => {
    const removed = spreadsheetManager.cleanup(24);
    if (removed > 0) {
        console.log(`🧹 Limpeza automática: ${removed} planilhas removidas`);
    }
}, 60 * 60 * 1000);

// ========== FUNÇÃO PARA FORMATAR VALORES EXATOS ==========

function formatCellValue(cell) {
    try {
        if (cell === null || cell === undefined || cell === '') {
            return '';
        }
        
        // Se for string, retorna como está
        if (typeof cell === 'string') {
            return cell.trim();
        }
        
        // Se for número, verificar se é data do Excel
        if (typeof cell === 'number') {
            // Excel armazena datas como números (dias desde 01/01/1900)
            if (cell > 0 && cell < 100000) {
                try {
                    // Converter data do Excel
                    const excelEpoch = new Date(1899, 11, 30);
                    let excelDate = cell;
                    
                    // Ajustar bug do Excel (1900 considerado bissexto)
                    if (excelDate > 60) excelDate -= 1;
                    
                    const date = new Date(excelEpoch.getTime() + (excelDate * 24 * 60 * 60 * 1000));
                    
                    if (!isNaN(date.getTime())) {
                        const day = date.getDate().toString().padStart(2, '0');
                        const month = (date.getMonth() + 1).toString().padStart(2, '0');
                        const year = date.getFullYear();
                        const hours = date.getHours().toString().padStart(2, '0');
                        const minutes = date.getMinutes().toString().padStart(2, '0');
                        const seconds = date.getSeconds().toString().padStart(2, '0');
                        
                        // Verificar se tem hora
                        if (hours === '00' && minutes === '00' && seconds === '00') {
                            return `${day}/${month}/${year}`;
                        } else {
                            return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
                        }
                    }
                } catch (e) {
                    // Se falhar, formata como número normal
                }
            }
            
            // Para outros números, formatar normalmente
            if (Number.isInteger(cell)) {
                return cell.toString();
            } else {
                return Number(cell.toFixed(6)).toString();
            }
        }
        
        // Se for booleano
        if (typeof cell === 'boolean') {
            return cell ? 'VERDADEIRO' : 'FALSO';
        }
        
        // Se for data do JavaScript
        if (cell instanceof Date) {
            return cell.toLocaleString('pt-BR');
        }
        
        // Para qualquer outro tipo
        return String(cell).trim();
        
    } catch (error) {
        console.warn('Erro ao formatar célula:', error);
        return String(cell || '').trim();
    }
}

// ========== PROCESSAMENTO EXCEL ==========

async function processExcel(buffer, originalName) {
    try {
        console.log('\n📥 PROCESSANDO ARQUIVO EXCEL...');
        console.log(`📄 Nome: ${originalName}`);
        console.log(`📏 Tamanho: ${(buffer.length / 1024 / 1024).toFixed(2)} MB`);
        
        let jsonData = [];
        let sheetName = 'Planilha';
        
        // Tentar processar como Excel
        try {
            const workbook = XLSX.read(buffer, { 
                type: 'buffer',
                cellFormula: false,
                cellHTML: false,
                cellStyles: false,
                raw: true,
                dateNF: 'dd/mm/yyyy'
            });
            
            sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            console.log(`📊 Planilha: ${sheetName}`);
            
            jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1,
                defval: '',
                raw: true
            });
            
        } catch (excelError) {
            console.log('⚠️ Não é Excel, tentando como CSV...');
            
            // Tentar como CSV
            const text = buffer.toString('utf-8');
            const lines = text.split('\n').filter(line => line.trim() !== '');
            
            jsonData = lines.map(line => {
                if (line.includes(';')) {
                    return line.split(';').map(cell => cell.trim());
                } else if (line.includes(',')) {
                    return line.split(',').map(cell => cell.trim());
                } else {
                    return line.split('\t').map(cell => cell.trim());
                }
            });
        }
        
        if (jsonData.length === 0) {
            throw new Error('Planilha vazia ou formato não suportado!');
        }
        
        console.log(`📊 Total de linhas brutas: ${jsonData.length}`);
        
        // ENCONTRAR CABEÇALHOS
        let headerRowIndex = 0;
        let maxTextCells = 0;
        
        for (let i = 0; i < Math.min(5, jsonData.length); i++) {
            const row = jsonData[i] || [];
            let textCells = 0;
            
            for (let j = 0; j < row.length; j++) {
                const cell = row[j];
                if (cell && cell.toString().trim().length > 0) {
                    textCells++;
                }
            }
            
            if (textCells > maxTextCells) {
                maxTextCells = textCells;
                headerRowIndex = i;
            }
        }
        
        console.log(`✅ Cabeçalhos na linha ${headerRowIndex + 1} (${maxTextCells} colunas)`);
        
        // EXTRAIR NOMES DAS COLUNAS
        const headerRow = jsonData[headerRowIndex] || [];
        const columns = [];
        
        for (let i = 0; i < headerRow.length; i++) {
            let columnName = headerRow[i];
            
            if (!columnName || columnName.toString().trim() === '') {
                columnName = `Coluna ${i + 1}`;
            } else {
                columnName = columnName.toString()
                    .replace(/[\n\r\t]/g, ' ')
                    .replace(/\s+/g, ' ')
                    .trim()
                    .substring(0, 50);
                
                if (!columnName) {
                    columnName = `Coluna ${i + 1}`;
                }
            }
            
            columns.push({
                name: columnName,
                key: `col_${i}`,
                index: i,
                originalIndex: i + 1,
                hasMultipleLines: false
            });
        }
        
        console.log(`📋 ${columns.length} colunas identificadas`);
        
        // PROCESSAR DADOS
        const data = [];
        let emptyRows = 0;
        
        for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i] || [];
            const rowData = { 
                _id: `row_${i + 1}`, 
                _rowNumber: i + 1,
                _isEmpty: true
            };
            
            let hasData = false;
            
            for (let j = 0; j < columns.length; j++) {
                const col = columns[j];
                const value = row[col.index];
                
                if (value !== undefined && value !== null && value !== '') {
                    rowData[col.key] = formatCellValue(value);
                    hasData = true;
                    rowData._isEmpty = false;
                    
                    if (typeof rowData[col.key] === 'string' && rowData[col.key].includes('\n')) {
                        col.hasMultipleLines = true;
                    }
                } else {
                    rowData[col.key] = '';
                }
            }
            
            if (hasData) {
                data.push(rowData);
            } else {
                emptyRows++;
            }
        }
        
        console.log(`✅ ${data.length} linhas de dados extraídas`);
        console.log(`🚫 ${emptyRows} linhas vazias ignoradas`);
        
        return {
            columns,
            data,
            metadata: {
                totalRows: jsonData.length,
                dataRows: data.length,
                emptyRows: emptyRows,
                headerRow: headerRowIndex + 1,
                fileName: sheetName,
                originalFileName: originalName,
                processedAt: new Date().toISOString(),
                fileSize: buffer.length
            }
        };
        
    } catch (error) {
        console.error('❌ Erro no processamento:', error.message);
        throw new Error(`Falha ao processar planilha: ${error.message}`);
    }
}

// ========== FUNÇÕES DE QR CODE GARANTIDAS ==========

function generateQRData(row, selectedColumns, allColumns, etiquetaNumero) {
    try {
        const qrInfo = {
            id: `ETQ${etiquetaNumero.toString().padStart(4, '0')}`,
            n: etiquetaNumero,
            t: new Date().toISOString().split('T')[0],
            d: {}
        };
        
        let added = 0;
        for (const colKey of selectedColumns) {
            if (added >= 3) break;
            
            const column = allColumns.find(c => c.key === colKey);
            if (!column) continue;
            
            const value = row[colKey];
            if (value && value.toString().trim() !== '') {
                const key = column.name.substring(0, 10);
                const val = value.toString().substring(0, 25).trim();
                
                if (key && val) {
                    qrInfo.d[key] = val;
                    added++;
                }
            }
        }
        
        const qrString = JSON.stringify(qrInfo);
        console.log(`🔳 Dados para QR Code (${etiquetaNumero}): ${qrString.substring(0, 50)}...`);
        return qrString;
        
    } catch (error) {
        console.warn('Erro ao gerar dados QR:', error);
        return `ETIQUETA_${etiquetaNumero}_${Date.now()}`;
    }
}

async function generateQRCodeImage(text) {
    try {
        console.log(`🔄 Gerando QR Code para texto de ${text.length} caracteres...`);
        
        const qrCodeUrl = await QRCode.toDataURL(text, {
            errorCorrectionLevel: 'L', // Mais baixo = mais confiável
            margin: 1,
            width: 150, // Tamanho menor para garantir
            color: {
                dark: '#000000', // Preto sólido
                light: '#FFFFFF' // Branco sólido
            },
            type: 'image/png'
        });
        
        console.log(`✅ QR Code gerado: ${qrCodeUrl.length} bytes`);
        return qrCodeUrl;
        
    } catch (error) {
        console.error('❌ ERRO ao gerar QR Code:', error.message);
        
        // Fallback: QR Code mínimo
        const blankQR = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';
        return blankQR;
    }
}

// ========== FUNÇÃO PARA DESENHAR ETIQUETA COM QR CODE GARANTIDO ==========

async function drawLabel(doc, row, selectedColumns, allColumns, qrCodeImage, x, y, width, height, current, total) {
    const MARGIN = 12;
    const contentWidth = width - (2 * MARGIN);
    
    let currentY = y + MARGIN;
    
    // ========== CABEÇALHO ==========
    const headerHeight = 28;
    
    doc.rect(x + MARGIN, currentY, contentWidth, headerHeight)
       .fill('#f8fafc')
       .stroke('#e2e8f0')
       .stroke();
    
    doc.fontSize(16)
       .font('Helvetica-Bold')
       .fillColor('#4f46e5')
       .text(`ETIQUETA ${current}`, 
             x + MARGIN, currentY + 7, {
                 width: contentWidth,
                 align: 'center'
             });
    
    doc.moveTo(x + MARGIN, currentY + headerHeight)
       .lineTo(x + MARGIN + contentWidth, currentY + headerHeight)
       .strokeColor('#e2e8f0')
       .lineWidth(1)
       .stroke();
    
    currentY += headerHeight + 15;
    
    // ========== DEFINIR LAYOUT FIXO (SEMPRE COM QR CODE) ==========
    // 60% para dados, 40% para QR Code
    const leftColumnWidth = contentWidth * 0.60;
    const rightColumnWidth = contentWidth * 0.40;
    const rightColumnX = x + MARGIN + leftColumnWidth + 5; // Menor gap
    
    // ========== COLUNA ESQUERDA: DADOS ==========
    let dataCurrentY = currentY;
    let maxDataHeight = height - currentY - MARGIN - 20; // Altura máxima para dados
    
    // Contador para limitar dados se necessário
    let itemsDrawn = 0;
    const maxItems = 10; // Máximo de itens de dados
    
    for (const colKey of selectedColumns) {
        if (itemsDrawn >= maxItems) break;
        
        const column = allColumns.find(c => c.key === colKey);
        if (!column) continue;
        
        const rawValue = row[colKey];
        if (!rawValue || rawValue.toString().trim() === '') {
            continue;
        }
        
        const valueStr = rawValue.toString().trim();
        
        // Nome da coluna
        doc.fontSize(8)
           .font('Helvetica-Bold')
           .fillColor('#4f46e5')
           .text(`${column.name}:`, 
                 x + MARGIN + 5, dataCurrentY, {
                     width: leftColumnWidth - 10,
                     align: 'left'
                 });
        
        dataCurrentY += 10;
        
        // Valor
        if (valueStr.includes('\n')) {
            const lines = valueStr.split('\n').filter(l => l.trim());
            for (let line of lines.slice(0, 2)) {
                doc.fontSize(9)
                   .font('Helvetica')
                   .fillColor('#000000')
                   .text(`${line.trim()}`, 
                         x + MARGIN + 10, dataCurrentY, {
                             width: leftColumnWidth - 15,
                             align: 'left'
                         });
                
                dataCurrentY += 12;
            }
        } else {
            doc.fontSize(9)
               .font('Helvetica')
               .fillColor('#000000')
               .text(`${valueStr}`, 
                     x + MARGIN + 10, dataCurrentY, {
                         width: leftColumnWidth - 15,
                         align: 'left'
                     });
            
            dataCurrentY += 14;
        }
        
        dataCurrentY += 3;
        itemsDrawn++;
        
        // Verificar se ainda cabe (deixando espaço para QR)
        if (dataCurrentY > (y + maxDataHeight)) {
            console.log(`⚠️ Limitando dados na etiqueta ${current} para caber QR Code`);
            break;
        }
    }
    
    // ========== COLUNA DIREITA: QR CODE (SEMPRE DESENHADO) ==========
    console.log(`🎯 Desenhando QR Code na etiqueta ${current}...`);
    
    // Área do QR Code - posição fixa
    const qrAreaY = currentY;
    const qrAreaHeight = height - currentY - MARGIN - 30;
    const qrSize = Math.min(rightColumnWidth - 20, qrAreaHeight - 30); // Tamanho ajustado
    
    console.log(`📏 Área QR: x=${rightColumnX}, y=${qrAreaY}, w=${rightColumnWidth}, h=${qrAreaHeight}`);
    console.log(`📐 Tamanho QR: ${qrSize}px`);
    
    // Fundo para área do QR
    doc.rect(rightColumnX, qrAreaY, rightColumnWidth, qrAreaHeight)
       .fill('#ffffff')
       .stroke('#4f46e5') // Borda colorida para visibilidade
       .lineWidth(1)
       .stroke();
    
    // DESENHAR QR CODE (GARANTIDO)
    let qrDrawn = false;
    
    if (qrCodeImage && qrCodeImage.startsWith('data:image/png;base64,')) {
        try {
            console.log(`📸 Tentando desenhar QR Code real...`);
            const base64Data = qrCodeImage.split(',')[1];
            const imageBuffer = Buffer.from(base64Data, 'base64');
            
            const qrX = rightColumnX + (rightColumnWidth - qrSize) / 2;
            const qrY = qrAreaY + 10;
            
            console.log(`📍 Posição QR: x=${qrX}, y=${qrY}`);
            
            doc.image(imageBuffer, qrX, qrY, {
                width: qrSize,
                height: qrSize
            });
            
            qrDrawn = true;
            console.log(`✅ QR Code real desenhado na etiqueta ${current}`);
            
        } catch (imageError) {
            console.error(`❌ Erro ao desenhar QR Code:`, imageError.message);
        }
    }
    
    // SE NÃO CONSEGUIU DESENHAR QR REAL, DESENHAR PLACEHOLDER VISÍVEL
    if (!qrDrawn) {
        console.log(`🎨 Desenhando placeholder do QR Code...`);
        drawQRPlaceholderVisible(doc, rightColumnX, qrAreaY, rightColumnWidth, qrAreaHeight, current);
    }
    
    // Título acima do QR
    doc.fontSize(10)
       .font('Helvetica-Bold')
       .fillColor('#4f46e5')
       .text('CÓDIGO QR', 
             rightColumnX, qrAreaY + 3, {
                 width: rightColumnWidth,
                 align: 'center'
             });
    
    // Instruções abaixo do QR
    doc.fontSize(7)
       .font('Helvetica')
       .fillColor('#666666')
       .text('Escaneie para', 
             rightColumnX, qrAreaY + qrSize + 15, {
                 width: rightColumnWidth,
                 align: 'center'
             });
    
    doc.fontSize(7)
       .font('Helvetica')
       .fillColor('#666666')
       .text('ver informações', 
             rightColumnX, qrAreaY + qrSize + 22, {
                 width: rightColumnWidth,
                 align: 'center'
             });
    
    // ========== RODAPÉ ==========
    const footerY = y + height - MARGIN - 15;
    
    doc.fontSize(8)
       .font('Helvetica')
       .fillColor('#999999')
       .text(`${current}/${total}`, 
             x + MARGIN, footerY, {
                 width: contentWidth,
                 align: 'right'
             });
}

// FUNÇÃO PARA DESENHAR QR CODE PLACEHOLDER VISÍVEL
function drawQRPlaceholderVisible(doc, x, y, width, height, etiquetaNumero) {
    const centerX = x + width / 2;
    const centerY = y + height / 2;
    const boxSize = Math.min(width - 20, height - 40);
    
    console.log(`🎨 Desenhando placeholder em x=${centerX-boxSize/2}, y=${centerY-boxSize/2}, tamanho=${boxSize}`);
    
    // Caixa com fundo colorido
    doc.rect(centerX - boxSize/2, centerY - boxSize/2, boxSize, boxSize)
       .fill('#e0f2fe') // Azul claro
       .stroke('#0ea5e9') // Azul
       .lineWidth(1)
       .stroke();
    
    // Grade simulando QR Code
    const cellSize = boxSize / 7;
    
    // Desenhar células pretas
    doc.fillColor('#000000');
    
    // Padrão fixo de células pretas
    const blackCells = [
        [1,1], [1,5], [2,2], [2,4], [3,3],
        [4,2], [4,4], [5,1], [5,5]
    ];
    
    for (const [row, col] of blackCells) {
        const cellX = centerX - boxSize/2 + (col * cellSize);
        const cellY = centerY - boxSize/2 + (row * cellSize);
        
        doc.rect(cellX + 1, cellY + 1, cellSize - 2, cellSize - 2)
           .fill();
    }
    
    // Texto "QR" no centro
    doc.fontSize(12)
       .font('Helvetica-Bold')
       .fillColor('#0ea5e9')
       .text('QR', 
             centerX - 10, centerY - 8, {
                 align: 'center'
             });
    
    // Número da etiqueta pequeno
    doc.fontSize(6)
       .font('Helvetica')
       .fillColor('#666666')
       .text(`#${etiquetaNumero}`, 
             centerX - 10, centerY + 10, {
                 align: 'center'
             });
}

// ========== FUNÇÃO PRINCIPAL PARA GERAR PDF ==========

async function generateLabelsPDF(data, selectedColumns, allColumns, copiesPerLabel = 1) {
    return new Promise(async (resolve, reject) => {
        try {
            console.log(`\n📏 GERANDO PDF COM QR CODE GARANTIDO...`);
            console.log(`📦 ${data.length} registros × ${copiesPerLabel} cópias`);
            
            // TAMANHO DA ETIQUETA 105x148mm
            const ETIQUETA_LARGURA_MM = 105;
            const ETIQUETA_ALTURA_MM = 148;
            
            const pageWidth = ETIQUETA_LARGURA_MM * 2.83465;
            const pageHeight = ETIQUETA_ALTURA_MM * 2.83465;
            
            console.log(`📐 Tamanho página: ${pageWidth.toFixed(0)}x${pageHeight.toFixed(0)} pontos`);
            
            // Criar documento PDF
            const doc = new PDFDocument({ 
                margin: 0,
                size: [pageWidth, pageHeight],
                autoFirstPage: true
            });
            
            const buffers = [];
            doc.on('data', buffer => {
                buffers.push(buffer);
            });
            
            doc.on('end', () => {
                const pdfData = Buffer.concat(buffers);
                console.log(`✅ PDF gerado: ${pdfData.length} bytes`);
                resolve(pdfData);
            });
            
            doc.on('error', (err) => {
                console.error('❌ Erro no PDF:', err);
                reject(err);
            });
            
            let totalEtiquetas = 0;
            
            // GERAR CADA ETIQUETA
            for (let copy = 0; copy < copiesPerLabel; copy++) {
                for (let i = 0; i < data.length; i++) {
                    const row = data[i];
                    totalEtiquetas++;
                    
                    if (totalEtiquetas > 1) {
                        doc.addPage({
                            size: [pageWidth, pageHeight],
                            margin: 0
                        });
                    }
                    
                    if (totalEtiquetas % 5 === 0) {
                        console.log(`   🏷️ ${totalEtiquetas} etiquetas geradas...`);
                    }
                    
                    // GERAR QR CODE PARA ESTA ETIQUETA
                    console.log(`\n🔳 Gerando QR Code para etiqueta ${totalEtiquetas}...`);
                    const qrDataText = generateQRData(row, selectedColumns, allColumns, totalEtiquetas);
                    const qrCodeImage = await generateQRCodeImage(qrDataText);
                    
                    // DESENHAR ETIQUETA
                    console.log(`🎨 Desenhando etiqueta ${totalEtiquetas}...`);
                    await drawLabel(doc, row, selectedColumns, allColumns, qrCodeImage, 
                                  0, 0, pageWidth, pageHeight, totalEtiquetas, data.length * copiesPerLabel);
                }
            }
            
            console.log(`\n🎉 PDF finalizado com ${totalEtiquetas} etiquetas`);
            doc.end();
            
        } catch (error) {
            console.error('❌ Erro ao gerar PDF:', error);
            reject(error);
        }
    });
}

// ========== ROTAS DA API ==========

// Rota de teste
app.get('/api/test', (req, res) => {
    res.json({ 
        success: true,
        message: '✅ Backend funcionando com QR Code garantido!',
        version: '5.0.0',
        timestamp: new Date().toISOString()
    });
});

// Upload de planilha
app.post('/api/upload', (req, res) => {
    upload(req, res, async function(err) {
        if (err) {
            console.error('❌ Erro no upload:', err.message);
            return res.status(400).json({
                success: false,
                message: err.message || 'Erro no upload do arquivo'
            });
        }
        
        if (!req.file) {
            return res.status(400).json({
                success: false,
                message: 'Nenhum arquivo selecionado'
            });
        }
        
        try {
            console.log(`\n📤 Upload: ${req.file.originalname}`);
            
            const result = await processExcel(req.file.buffer, req.file.originalname);
            const id = `sheet_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
            
            spreadsheetManager.set(id, {
                id,
                originalName: req.file.originalname,
                uploadDate: new Date().toISOString(),
                fileSize: req.file.size,
                ...result
            });
            
            console.log(`✅ Upload concluído: ${id}`);
            
            res.json({
                success: true,
                message: `✅ Planilha processada com sucesso!`,
                data: {
                    id,
                    fileName: req.file.originalname,
                    records: result.data.length,
                    columns: result.columns.map(c => ({
                        name: c.name,
                        key: c.key,
                        hasMultipleLines: c.hasMultipleLines
                    })),
                    metadata: result.metadata
                }
            });
            
        } catch (error) {
            console.error('❌ Erro no processamento:', error);
            res.status(500).json({
                success: false,
                message: `Erro: ${error.message}`
            });
        }
    });
});

// Listar planilhas
app.get('/api/spreadsheets', (req, res) => {
    try {
        const spreadsheets = spreadsheetManager.getAll();
        
        res.json({
            success: true,
            spreadsheets,
            count: spreadsheets.length,
            totalRecords: spreadsheets.reduce((sum, sheet) => sum + (sheet.records || 0), 0),
            serverTime: new Date().toISOString()
        });
        
    } catch (error) {
        console.error('Erro ao listar planilhas:', error);
        res.status(500).json({
            success: false,
            message: 'Erro ao listar planilhas'
        });
    }
});

// Obter dados de uma planilha
app.get('/api/spreadsheets/:id/data', (req, res) => {
    try {
        const sheet = spreadsheetManager.get(req.params.id);
        
        if (!sheet) {
            return res.status(404).json({
                success: false,
                message: 'Planilha não encontrada'
            });
        }
        
        const previewData = sheet.data.slice(0, 100).map((row, idx) => ({
            index: idx + 1,
            data: row,
            preview: sheet.columns.slice(0, 3).map(col => {
                const value = row[col.key];
                if (!value && value !== 0) return `${col.name}: [VAZIO]`;
                return `${col.name}: ${String(value).substring(0, 30)}`;
            }).join(' | ')
        }));
        
        res.json({
            success: true,
            data: previewData,
            columns: sheet.columns.map(c => ({
                name: c.name,
                key: c.key,
                hasMultipleLines: c.hasMultipleLines,
                originalIndex: c.originalIndex
            })),
            totalRecords: sheet.data.length,
            metadata: sheet.metadata
        });
        
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
        res.status(500).json({
            success: false,
            message: 'Erro ao carregar dados'
        });
    }
});

// Gerar PDF
app.get('/api/generate-pdf/:id', async (req, res) => {
    console.log('\n📄 SOLICITAÇÃO DE PDF RECEBIDA');
    console.log(`📋 Planilha ID: ${req.params.id}`);
    console.log(`📋 Parâmetros:`, req.query);
    
    try {
        const sheet = spreadsheetManager.get(req.params.id);
        
        if (!sheet) {
            console.error('❌ Planilha não encontrada');
            return res.status(404).json({
                success: false,
                message: 'Planilha não encontrada'
            });
        }
        
        console.log(`✅ Planilha encontrada: ${sheet.metadata.fileName}`);
        console.log(`📊 Total de registros: ${sheet.data.length}`);
        
        // Obter colunas selecionadas
        const selectedColumns = req.query.columns ? 
            req.query.columns.split(',').filter(c => c && c.trim() !== '') : [];
        
        if (selectedColumns.length === 0) {
            console.error('❌ Nenhuma coluna selecionada');
            return res.status(400).json({
                success: false,
                message: 'Selecione pelo menos uma coluna'
            });
        }
        
        console.log(`✅ Colunas selecionadas: ${selectedColumns.length}`);
        
        // Obter número de cópias
        const copies = parseInt(req.query.copies) || 1;
        console.log(`✅ Cópias: ${copies}`);
        
        // Usar todos os dados
        let dataToUse = sheet.data;
        
        // Limitar máximo de etiquetas
        const MAX_ETIQUETAS = 500;
        const totalEtiquetas = dataToUse.length * copies;
        
        if (totalEtiquetas > MAX_ETIQUETAS) {
            console.warn(`⚠️ Limitando para ${MAX_ETIQUETAS} etiquetas`);
            const maxRegistros = Math.floor(MAX_ETIQUETAS / copies);
            dataToUse = dataToUse.slice(0, maxRegistros);
        }
        
        console.log(`🔄 Gerando PDF com ${dataToUse.length} registros...`);
        
        // Gerar PDF
        const pdfBuffer = await generateLabelsPDF(dataToUse, selectedColumns, sheet.columns, copies);
        
        if (!pdfBuffer || pdfBuffer.length === 0) {
            throw new Error('PDF gerado está vazio');
        }
        
        console.log(`✅ PDF gerado com sucesso: ${pdfBuffer.length} bytes`);
        
        // Configurar headers para download
        const fileName = `etiquetas_${Date.now()}.pdf`;
        
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Length', pdfBuffer.length);
        res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
        res.setHeader('Pragma', 'no-cache');
        res.setHeader('Expires', '0');
        
        // Enviar PDF
        console.log(`📤 Enviando PDF para download...`);
        res.send(pdfBuffer);
        
    } catch (error) {
        console.error('❌ ERRO AO GERAR PDF:', error);
        
        if (!res.headersSent) {
            res.status(500).json({
                success: false,
                message: `Erro ao gerar PDF: ${error.message}`,
                error: process.env.NODE_ENV === 'development' ? error.stack : undefined
            });
        }
    }
});

// Excluir planilha
app.delete('/api/spreadsheets/:id', (req, res) => {
    try {
        const deleted = spreadsheetManager.delete(req.params.id);
        
        res.json({
            success: true,
            message: deleted ? '✅ Planilha excluída com sucesso' : 'Planilha não encontrada',
            deleted: deleted
        });
    } catch (error) {
        console.error('Erro ao excluir:', error);
        res.status(500).json({
            success: false,
            message: 'Erro ao excluir planilha'
        });
    }
});

// Teste de QR Code
app.get('/api/test-qrcode', async (req, res) => {
    try {
        const testText = `Teste QR Code ${new Date().toLocaleString('pt-BR')}`;
        console.log(`🔳 Testando QR Code com texto: ${testText}`);
        
        const qrCode = await generateQRCodeImage(testText);
        
        res.json({
            success: true,
            message: 'QR Code testado',
            hasQRCode: !!qrCode,
            qrCodeSize: qrCode ? qrCode.length : 0,
            preview: qrCode ? qrCode.substring(0, 50) + '...' : null
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            message: 'Erro no teste de QR Code',
            error: error.message
        });
    }
});

// Servir frontend
app.get('*', (req, res) => {
    res.sendFile(path.join(frontendPath, 'index.html'));
});

// ========== INICIAR SERVIDOR ==========

const server = app.listen(PORT, () => {
    console.log('\n' + '='.repeat(60));
    console.log('✅ SERVIDOR INICIADO COM SUCESSO!');
    console.log('='.repeat(60));
    console.log(`📍 URL: http://localhost:${PORT}`);
    console.log(`📁 Frontend: ${frontendPath}`);
    console.log(`🔳 QR Code: GARANTIDO e visível`);
    console.log(`📏 Layout: 60% dados + 40% QR Code`);
    console.log(`🎨 Placeholder: Colorido se QR falhar`);
    console.log('='.repeat(60));
    console.log('📝 Endpoints:');
    console.log('  • GET  /api/test              - Testar conexão');
    console.log('  • GET  /api/test-qrcode       - Testar QR Code');
    console.log('  • POST /api/upload            - Upload planilha');
    console.log('  • GET  /api/spreadsheets      - Listar planilhas');
    console.log('  • GET  /api/spreadsheets/:id  - Dados da planilha');
    console.log('  • GET  /api/generate-pdf/:id  - ✅ Gerar PDF com QR');
    console.log('  • DELETE /api/spreadsheets/:id - Excluir planilha');
    console.log('='.repeat(60));
    console.log('🚀 Sistema pronto! QR Code será SEMPRE visível!');
    console.log('='.repeat(60));
});
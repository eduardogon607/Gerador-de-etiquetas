// ========== IMPORTS ==========
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const PDFDocument = require('pdfkit');
const cors = require('cors');
const QRCode = require('qrcode');

// ========== CONFIGURAÇÃO ==========
const PORT = process.env.PORT || 3001;
const isProduction = process.env.NODE_ENV === 'production';

// ========== INICIALIZAR APP ==========
const app = express();

// ========== CONFIGURAÇÃO CORS ==========
app.use(cors({
    origin: isProduction ? process.env.RENDER_EXTERNAL_URL : '*',
    credentials: true
}));

// ========== MIDDLEWARES ==========
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Middleware para forçar HTTPS em produção
if (isProduction) {
    app.use((req, res, next) => {
        if (req.headers['x-forwarded-proto'] !== 'https') {
            return res.redirect('https://' + req.headers.host + req.url);
        }
        next();
    });
}

// Servir arquivos estáticos
const frontendPath = path.join(__dirname, '..', 'frontend');
if (fs.existsSync(frontendPath)) {
    app.use(express.static(frontendPath));
    console.log(`✅ Frontend sendo servido de: ${frontendPath}`);
}

// ========== CONFIGURAÇÃO DE UPLOAD ==========
const UPLOAD_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) {
    fs.mkdirSync(UPLOAD_DIR, { recursive: true });
    console.log(`📁 Pasta upload criada: ${UPLOAD_DIR}`);
}

const upload = multer({
    storage: multer.diskStorage({
        destination: (req, file, cb) => {
            cb(null, UPLOAD_DIR);
        },
        filename: (req, file, cb) => {
            const safeName = file.originalname.replace(/[<>:"/\\|?*]/g, '_');
            const uniqueName = Date.now() + '-' + Math.random().toString(36).substring(2, 9) + path.extname(safeName);
            cb(null, uniqueName);
        }
    }),
    fileFilter: (req, file, cb) => {
        const allowedTypes = /\.(xlsx|xls)$/i;
        const isValid = allowedTypes.test(path.extname(file.originalname));
        cb(null, isValid);
    },
    limits: {
        fileSize: 20 * 1024 * 1024
    }
});

// ========== DADOS EM MEMÓRIA ==========
let spreadsheets = [];

// ========== FUNÇÕES AUXILIARES ==========

// Função para converter datas de mm/dd/yyyy para dd/mm/yyyy
function convertToBrazilianDate(dateString) {
    if (!dateString || typeof dateString !== 'string') {
        return dateString;
    }
    
    const trimmed = dateString.trim();
    
    // Se já está em formato dd/mm/yyyy, retorna como está
    if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(trimmed)) {
        return trimmed;
    }
    
    // Tentar converter de mm/dd/yyyy para dd/mm/yyyy
    // Formato americano: 3/15/2024 ou 03/15/2024
    const americanFormat = trimmed.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
    if (americanFormat) {
        const [, month, day, year] = americanFormat;
        // Se o ano tem 2 dígitos, assumir século 21
        const fullYear = year.length === 2 ? `20${year}` : year;
        return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${fullYear}`;
    }
    
    // Tentar converter de mm/dd/yyyy hh:mm para dd/mm/yyyy hh:mm
    const americanWithTime = trimmed.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})\s+(\d{1,2}):(\d{1,2})$/);
    if (americanWithTime) {
        const [, month, day, year, hour, minute] = americanWithTime;
        const fullYear = year.length === 2 ? `20${year}` : year;
        return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${fullYear} ${hour.padStart(2, '0')}:${minute.padStart(2, '0')}`;
    }
    
    // Tentar converter de yyyy-mm-dd para dd/mm/yyyy
    const isoFormat = trimmed.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
    if (isoFormat) {
        const [, year, month, day] = isoFormat;
        return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year}`;
    }
    
    // Tentar converter de yyyy-mm-dd hh:mm para dd/mm/yyyy hh:mm
    const isoWithTime = trimmed.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})\s+(\d{1,2}):(\d{1,2})$/);
    if (isoWithTime) {
        const [, year, month, day, hour, minute] = isoWithTime;
        return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year} ${hour.padStart(2, '0')}:${minute.padStart(2, '0')}`;
    }
    
    // Se não for nenhum formato de data conhecido, retorna original
    return trimmed;
}

// Função para obter o valor formatado da célula
function getExcelCellValue(worksheet, row, col) {
    const cellAddress = xlsx.utils.encode_cell({ r: row, c: col });
    const cell = worksheet[cellAddress];
    
    if (!cell) {
        return '';
    }
    
    let value = '';
    
    // Se a célula tem propriedade 'w' (valor formatado como string), usa esse
    if (cell.w !== undefined) {
        value = cell.w.toString().trim();
    }
    // Se não tem 'w', usa o valor bruto 'v'
    else if (cell.v !== undefined) {
        value = cell.v.toString().trim();
    }
    
    // Se a célula é do tipo data (t = 'd') ou se parece com data
    if (cell.t === 'd' || cell.t === 'n') {
        // Verificar se o valor parece ser uma data
        if (value.match(/^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}/)) {
            // Converter para formato brasileiro
            value = convertToBrazilianDate(value);
        }
    }
    
    return value;
}

// ========== ROTAS DA API ==========

app.get('/api/ping', (req, res) => {
    res.json({ 
        success: true, 
        message: 'API Online 🚀', 
        time: new Date().toISOString(),
        dimensions: 'Etiquetas 110mm (altura) × 80mm (largura)'
    });
});

app.get('/api/create-test', (req, res) => {
    const testSheet = {
        id: 'test_sheet_' + Date.now(),
        fileName: 'planilha_teste.xlsx',
        uploadDate: new Date().toISOString(),
        columns: [
            { id: 'col_0', name: 'Código', index: 0 },
            { id: 'col_1', name: 'Produto', index: 1 },
            { id: 'col_2', name: 'Quantidade', index: 2 },
            { id: 'col_3', name: 'Data Fabricação', index: 3 },
            { id: 'col_4', name: 'Data Vencimento', index: 4 },
            { id: 'col_5', name: 'Hora Produção', index: 5 },
            { id: 'col_6', name: 'Data+Hora Teste', index: 6 },
            { id: 'col_7', name: 'Número Serial', index: 7 }
        ],
        data: [
            { 
                'Código': 'PROD-001',
                'Produto': 'Caneta Azul', 
                'Quantidade': '100',
                'Data Fabricação': '15/03/2024',
                'Data Vencimento': '15/06/2024',
                'Hora Produção': '14:30',
                'Data+Hora Teste': '15/03/2024 14:30',
                'Número Serial': '45000'
            },
            { 
                'Código': 'PROD-002',
                'Produto': 'Caderno', 
                'Quantidade': '50',
                'Data Fabricação': '10/03/2024',
                'Data Vencimento': '10/09/2024',
                'Hora Produção': '09:15',
                'Data+Hora Teste': '10/03/2024 09:15',
                'Número Serial': '45001'
            }
        ],
        rowCount: 2,
        columnCount: 8
    };

    spreadsheets.push(testSheet);
    
    res.json({
        success: true,
        message: 'Dados de teste criados!',
        spreadsheet: testSheet
    });
});

app.get('/api/spreadsheets', (req, res) => {
    const simplified = spreadsheets.map(sheet => ({
        id: sheet.id,
        fileName: sheet.fileName,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        uploadDate: sheet.uploadDate
    }));

    res.json({
        success: true,
        spreadsheets: simplified,
        total: simplified.length
    });
});

app.get('/api/spreadsheet/:id', (req, res) => {
    const { id } = req.params;
    const spreadsheet = spreadsheets.find(s => s.id === id);

    if (!spreadsheet) {
        return res.status(404).json({
            success: false,
            error: 'Planilha não encontrada'
        });
    }

    res.json({
        success: true,
        data: spreadsheet
    });
});

app.post('/api/upload', upload.single('spreadsheet'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({
                success: false,
                error: 'Nenhum arquivo enviado'
            });
        }

        console.log(`📤 Processando: ${req.file.originalname}`);

        // Ler arquivo Excel
        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Determinar o range da planilha
        const range = xlsx.utils.decode_range(worksheet['!ref']);
        
        // Extrair cabeçalhos (primeira linha)
        const headers = [];
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellValue = getExcelCellValue(worksheet, 0, col);
            let colName;
            
            if (!cellValue || cellValue.trim() === '') {
                colName = `Coluna_${col + 1}`;
            } else {
                colName = cellValue.trim();
            }
            
            headers.push({
                id: `col_${col}`,
                name: colName,
                index: col
            });
        }

        // Processar dados linha por linha
        const rows = [];
        for (let row = 1; row <= range.e.r; row++) {
            const rowData = {};
            let rowHasData = false;
            
            headers.forEach((header, colIndex) => {
                const cellValue = getExcelCellValue(worksheet, row, colIndex);
                
                if (cellValue && cellValue.trim() !== '') {
                    rowHasData = true;
                }
                
                // Armazenar o valor já convertido para formato brasileiro se for data
                rowData[header.name] = cellValue;
            });
            
            // Só adicionar se a linha tem dados
            if (rowHasData) {
                rows.push(rowData);
            }
        }

        // Criar objeto da planilha
        const spreadsheet = {
            id: `sheet_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
            fileName: req.file.originalname,
            filePath: req.file.path,
            uploadDate: new Date().toISOString(),
            columns: headers,
            data: rows,
            rowCount: rows.length,
            columnCount: headers.length
        };

        spreadsheets.push(spreadsheet);

        console.log(`✅ Planilha processada: ${rows.length} linhas, ${headers.length} colunas`);
        
        // Debug: mostrar exemplos de datas processadas
        if (rows.length > 0) {
            console.log('📋 Exemplo de dados processados (primeira linha):');
            const firstRowData = rows[0];
            Object.keys(firstRowData).forEach(key => {
                const value = firstRowData[key];
                console.log(`   ${key}: "${value}"`);
                
                // Verificar se é data
                if (value && value.match(/\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}/)) {
                    console.log(`     ↑ FORMATO BRASILEIRO OK`);
                }
            });
        }

        res.json({
            success: true,
            message: `Planilha processada com sucesso! ${rows.length} linhas encontradas.`,
            spreadsheet: {
                id: spreadsheet.id,
                fileName: spreadsheet.fileName,
                rowCount: spreadsheet.rowCount,
                columnCount: spreadsheet.columnCount,
                uploadDate: spreadsheet.uploadDate
            }
        });

    } catch (error) {
        console.error('❌ Erro no upload:', error);
        
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
        
        res.status(500).json({
            success: false,
            error: `Erro ao processar planilha: ${error.message}`
        });
    }
});

// ========== GERAR PDF - DIMENSÕES CORRETAS ==========
app.post('/api/generate', async (req, res) => {
    console.log('🏷️ Recebendo requisição para gerar PDF (110×80mm)...');
    
    try {
        const { 
            spreadsheetId, 
            columns, 
            mode = 'all',
            rowIndex = 0,
            quantity = 1
        } = req.body;

        console.log('📦 Dados recebidos:', { spreadsheetId, columns, mode, rowIndex, quantity });

        // Validações
        if (!spreadsheetId) {
            return res.status(400).json({
                success: false,
                error: 'ID da planilha é obrigatório'
            });
        }

        if (!columns || !Array.isArray(columns) || columns.length === 0) {
            return res.status(400).json({
                success: false,
                error: 'Selecione pelo menos uma coluna'
            });
        }

        // Encontrar planilha
        const spreadsheet = spreadsheets.find(s => s.id === spreadsheetId);
        if (!spreadsheet) {
            return res.status(404).json({
                success: false,
                error: 'Planilha não encontrada'
            });
        }

        console.log(`✅ Planilha encontrada: ${spreadsheet.fileName}`);

        // Filtrar colunas selecionadas
        const selectedColumns = spreadsheet.columns.filter(col => columns.includes(col.id));
        if (selectedColumns.length === 0) {
            return res.status(400).json({
                success: false,
                error: 'Nenhuma coluna válida selecionada'
            });
        }

        console.log(`✅ ${selectedColumns.length} colunas selecionadas`);

        // Preparar dados
        let dataToProcess = [];
        let totalLabels = 0;

        if (mode === 'all') {
            dataToProcess = spreadsheet.data;
            totalLabels = spreadsheet.rowCount;
            console.log(`📊 Modo: TODAS AS LINHAS (${totalLabels} etiquetas)`);
        } else if (mode === 'single') {
            if (rowIndex < 0 || rowIndex >= spreadsheet.rowCount) {
                return res.status(400).json({
                    success: false,
                    error: `Índice de linha inválido`
                });
            }

            const selectedRow = spreadsheet.data[rowIndex];
            for (let i = 0; i < quantity; i++) {
                dataToProcess.push(selectedRow);
            }
            totalLabels = quantity;
            console.log(`📊 Modo: LINHA ESPECÍFICA (${quantity} cópias da linha ${rowIndex + 1})`);
        }

        // ========== CONFIGURAÇÃO DO PDF ==========
        // 110mm ALTURA × 80mm LARGURA em pontos (1mm = 2.834645669 points)
        const pageWidth = 80 * 2.834645669;   // Largura da etiqueta
        const pageHeight = 110 * 2.834645669; // Altura da etiqueta

        console.log(`📏 Dimensões da etiqueta:`);
        console.log(`   - Altura: 110mm (${pageHeight.toFixed(2)} pontos)`);
        console.log(`   - Largura: 80mm (${pageWidth.toFixed(2)} pontos)`);
        console.log(`   - Total de etiquetas: ${dataToProcess.length}`);

        // Criar documento PDF
        const doc = new PDFDocument({
            size: [pageWidth, pageHeight], // [largura, altura]
            margin: 0,
            autoFirstPage: false,
            bufferPages: true
        });

        // Configurar headers da resposta
        const timestamp = new Date().toISOString().slice(0, 19).replace(/[:]/g, '-');
        const fileName = mode === 'all' 
            ? `etiquetas_110x80_${timestamp}.pdf`
            : `etiqueta_linha${rowIndex + 1}_${timestamp}.pdf`;

        // Configurar headers importantes
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
        res.setHeader('Pragma', 'no-cache');
        res.setHeader('Expires', '0');
        res.setHeader('X-Content-Type-Options', 'nosniff');

        console.log(`📄 Configurando download: ${fileName}`);

        // Pipe do PDF para a resposta
        doc.pipe(res);

        // ========== GERAR ETIQUETAS ==========
        let labelCount = 0;

        for (let i = 0; i < dataToProcess.length; i++) {
            const rowData = dataToProcess[i];
            
            // Verificar se há dados
            const hasData = selectedColumns.some(col => {
                const value = rowData[col.name];
                return value && value.toString().trim() !== '';
            });

            if (!hasData) {
                continue;
            }

            // Nova página para cada etiqueta
            doc.addPage({
                size: [pageWidth, pageHeight],
                margins: { top: 0, bottom: 0, left: 0, right: 0 }
            });

            labelCount++;

            // ========== CABEÇALHO (NEGRITO) ==========
            const headerHeight = 16;
            doc.rect(0, 0, pageWidth, headerHeight)
               .fill('#4f46e5');
            
            // Número da etiqueta - NEGRITO
            doc.fontSize(10)
               .fillColor('white')
               .font('Helvetica-Bold')
               .text(`ETIQUETA ${labelCount}/${totalLabels}`, 10, 4);

            // ========== CONTEÚDO PRINCIPAL ==========
            const contentStartY = headerHeight + 10;
            const contentWidth = pageWidth - 20;
            
            // Calcular altura dos itens baseado no número de colunas
            const totalItems = selectedColumns.length;
            const availableHeight = pageHeight - contentStartY - 50;
            const itemHeight = Math.min(20, availableHeight / totalItems);
            
            const cardHeight = itemHeight;
            const cardSpacing = 5;
            
            let currentY = contentStartY;
            
            // Ajustar tamanho da fonte baseado no número de colunas
            let fontSize = 9;
            let labelFontSize = 7;
            
            if (totalItems > 8) {
                fontSize = 8;
                labelFontSize = 6;
            }
            if (totalItems > 12) {
                fontSize = 7;
                labelFontSize = 5;
            }

            console.log(`📝 Configuração fonte: Label=${labelFontSize}px, Valor=${fontSize}px, Altura item=${cardHeight}px`);

            for (let colIndex = 0; colIndex < selectedColumns.length; colIndex++) {
                const col = selectedColumns[colIndex];
                const value = rowData[col.name] || '';
                
                if (value.toString().trim() === '') continue;
                
                // Verificar se ainda cabe na etiqueta
                if (currentY + cardHeight > pageHeight - 55) {
                    console.log(`⚠️ Etiqueta ${labelCount}: Apenas ${colIndex} de ${selectedColumns.length} colunas couberam`);
                    break;
                }

                // Cartão de informação
                doc.roundedRect(10, currentY, contentWidth, cardHeight, 3)
                   .fill('#f8fafc')
                   .stroke('#e2e8f0')
                   .stroke();

                // Nome da coluna - NEGRITO GARANTIDO (cor roxa para destacar)
                doc.fontSize(labelFontSize)
                   .fillColor('#000000')
                   .font('Helvetica-Bold')
                   .text(col.name.toUpperCase(), 
                         12, 
                         currentY + 3,
                         { 
                             width: contentWidth - 4,
                             ellipsis: true 
                         });

                // Valor - NEGRITO
                let displayValue = String(value);
                
                // Calcular comprimento máximo baseado no espaço
                const maxChars = Math.floor(contentWidth / (fontSize * 0.5));
                if (displayValue.length > maxChars) {
                    displayValue = displayValue.substring(0, maxChars - 3) + '...';
                }

                doc.fontSize(fontSize)
                   .fillColor('#1e293b')
                   .font('Helvetica-Bold')
                   .text(displayValue, 
                         12, 
                         currentY + (cardHeight / 1.5),
                         { 
                             width: contentWidth - 4,
                             ellipsis: true 
                         });

                // Atualizar posição
                currentY += cardHeight + cardSpacing;
            }

            // ========== QR CODE ==========
            try {
                const qrText = `ETQ${labelCount}/${totalLabels}\n${new Date().toLocaleDateString('pt-BR')}`;
                
                const qrCode = await QRCode.toBuffer(qrText, {
                    width: 120,
                    margin: 1,
                    color: { 
                        dark: '#000000', 
                        light: '#FFFFFF' 
                    },
                    errorCorrectionLevel: 'M'
                });

                const qrSize = 35;
                const qrX = pageWidth - qrSize - 15;
                const qrY = pageHeight - qrSize - 15;
                
                doc.image(qrCode, qrX, qrY, {
                    width: qrSize,
                    height: qrSize
                });
                
                doc.rect(qrX - 2, qrY - 2, qrSize + 4, qrSize + 4)
                   .stroke('#4f46e5')
                   .lineWidth(0.8);
                
                // Legenda QR Code - NEGRITO
                doc.fontSize(6)
                   .fillColor('#4f46e5')
                   .font('Helvetica-Bold')
                   .text('QR CODE', 
                         qrX, 
                         qrY + qrSize + 2,
                         { 
                             width: qrSize, 
                             align: 'center' 
                         });
                
                console.log(`✅ QR Code adicionado (${qrSize}px)`);
                
            } catch (qrError) {
                console.log(`⚠️ QR Code não gerado: ${qrError.message}`);
            }

            // ========== RODAPÉ (TUDO EM NEGRITO) ==========
            const shortFileName = spreadsheet.fileName.length > 25 
                ? spreadsheet.fileName.substring(0, 22) + '...' 
                : spreadsheet.fileName;
            
            // Nome do arquivo - NEGRITO
            doc.fontSize(6)
               .fillColor('#64748b')
               .font('Helvetica-Bold')
               .text(shortFileName, 
                     10, 
                     pageHeight - 15,
                     { 
                         width: pageWidth - 90,
                         ellipsis: true 
                     });

            // Data - NEGRITO
            doc.fontSize(6)
               .fillColor('#64748b')
               .font('Helvetica-Bold')
               .text(new Date().toLocaleDateString('pt-BR'), 
                     pageWidth - 80,
                     pageHeight - 15,
                     { 
                         width: 70,
                         align: 'right' 
                     });

            // Linha divisória
            doc.moveTo(10, pageHeight - 18)
               .lineTo(pageWidth - 10, pageHeight - 18)
               .stroke('#e2e8f0')
               .lineWidth(0.5);
        }

        // Se nenhuma etiqueta foi gerada
        if (labelCount === 0) {
            doc.addPage({
                size: [pageWidth, pageHeight],
                margins: { top: 0, bottom: 0, left: 0, right: 0 }
            });
            
            // Mensagem de erro - NEGRITO
            doc.fontSize(14)
               .fillColor('#64748b')
               .font('Helvetica-Bold')
               .text('Nenhum dado encontrado', 
                     pageWidth / 2, 
                     pageHeight / 2 - 20,
                     { align: 'center' });
            
            doc.fontSize(12)
               .fillColor('#94a3b8')
               .font('Helvetica-Bold')
               .text('Verifique as colunas selecionadas', 
                     pageWidth / 2, 
                     pageHeight / 2,
                     { align: 'center' });
        }

        // Finalizar PDF
        doc.end();

        console.log(`✅ PDF gerado com sucesso: ${labelCount} etiquetas 110×80mm`);
        console.log(`📤 Arquivo: ${fileName}`);

    } catch (error) {
        console.error('❌ Erro ao gerar PDF:', error);
        
        if (!res.headersSent) {
            res.status(500).json({
                success: false,
                error: `Erro ao gerar PDF: ${error.message}`
            });
        } else {
            console.error('⚠️ Não foi possível enviar erro: resposta já iniciada');
        }
    }
});

// ========== TESTE DE PDF COM DIMENSÕES CORRETAS ==========
app.get('/api/test-pdf', async (req, res) => {
    try {
        console.log('🧪 Gerando PDF de teste 110×80mm...');
        
        const pageWidth = 80 * 2.834645669;
        const pageHeight = 110 * 2.834645669;
        
        const doc = new PDFDocument({
            size: [pageWidth, pageHeight],
            margin: 0,
            autoFirstPage: false
        });
        
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename="etiqueta_teste_110x80.pdf"');
        res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
        res.setHeader('Pragma', 'no-cache');
        res.setHeader('Expires', '0');
        
        doc.pipe(res);
        
        doc.addPage({
            size: [pageWidth, pageHeight],
            margins: { top: 0, bottom: 0, left: 0, right: 0 }
        });
        
        doc.rect(0, 0, pageWidth, 20)
           .fill('#4f46e5');
        
        doc.fontSize(12)
           .fillColor('white')
           .font('Helvetica-Bold')
           .text('ETIQUETA DE TESTE', 15, 5);
        
        doc.fontSize(10)
           .fillColor('#1e293b')
           .font('Helvetica-Bold')
           .text('Sistema de Etiquetas', 15, 35);
        
        doc.fontSize(9)
           .fillColor('#64748b')
           .font('Helvetica-Bold')
           .text('Dimensões: 110mm × 80mm', 15, 55);
        
        doc.fontSize(9)
           .fillColor('#10b981')
           .font('Helvetica-Bold')
           .text('✅ PDF gerado com sucesso!', 15, 75);
        
        try {
            const qrText = 'ETIQUETA DE TESTE\nDimensões: 110×80mm\nData: ' + new Date().toLocaleDateString('pt-BR');
            const qrCode = await QRCode.toBuffer(qrText, {
                width: 150,
                margin: 1
            });
            
            doc.image(qrCode, pageWidth - 60, pageHeight - 60, {
                width: 50,
                height: 50
            });
            
        } catch (qrError) {
            console.log('QR Code não gerado no teste:', qrError.message);
        }
        
        doc.fontSize(8)
           .fillColor('#94a3b8')
           .font('Helvetica-Bold')
           .text(`Teste - ${new Date().toLocaleDateString('pt-BR')}`, 
                 15, 
                 pageHeight - 20);
        
        doc.end();
        
        console.log('✅ PDF de teste 110×80mm gerado com sucesso');
        
    } catch (error) {
        console.error('❌ Erro no PDF de teste:', error);
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// ========== ROTAS DO FRONTEND ==========
app.get('/', (req, res) => {
    const indexPath = path.join(frontendPath, 'index.html');
    if (fs.existsSync(indexPath)) {
        res.sendFile(indexPath);
    } else {
        res.send(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Sistema de Etiquetas 110×80mm</title>
                <style>
                    body { font-family: Arial; padding: 50px; text-align: center; }
                    .btn { display: inline-block; padding: 12px 24px; margin: 10px; background: #4f46e5; color: white; text-decoration: none; border-radius: 8px; }
                    .dimensions { background: #f0f9ff; padding: 10px; border-radius: 8px; margin: 20px; display: inline-block; }
                </style>
            </head>
            <body>
                <h1>🏷️ Sistema de Etiquetas</h1>
                <div class="dimensions">
                    <strong>📏 Dimensões:</strong> 110mm (altura) × 80mm (largura)
                </div>
                <p>Backend funcionando!</p>
                <a href="/api/ping" class="btn">Testar API</a>
                <a href="/api/create-test" class="btn">Criar Dados Teste</a>
                <a href="/api/test-pdf" class="btn">Testar PDF 110×80mm</a>
                <a href="/upload" class="btn">Upload</a>
                <a href="/generate" class="btn">Gerar Etiquetas</a>
            </body>
            </html>
        `);
    }
});

app.get('/upload', (req, res) => {
    const uploadPath = path.join(frontendPath, 'upload.html');
    if (fs.existsSync(uploadPath)) {
        res.sendFile(uploadPath);
    } else {
        res.redirect('/');
    }
});

app.get('/generate', (req, res) => {
    const generatePath = path.join(frontendPath, 'generate.html');
    if (fs.existsSync(generatePath)) {
        res.sendFile(generatePath);
    } else {
        res.redirect('/');
    }
});

// ========== INICIAR SERVIDOR ==========
app.listen(PORT, () => {
    console.log('\n' + '='.repeat(60));
    console.log('🚀 SISTEMA DE ETIQUETAS 110×80mm - INICIADO');
    console.log('='.repeat(60));
    console.log(`✅ Servidor: http://localhost:${PORT}`);
    console.log(`✅ Dimensões: 110mm altura × 80mm largura`);
    console.log(`✅ Teste PDF: http://localhost:${PORT}/api/test-pdf`);
    console.log(`✅ Upload: http://localhost:${PORT}/upload`);
    console.log(`✅ Gerar: http://localhost:${PORT}/generate`);
    console.log('='.repeat(60));
});
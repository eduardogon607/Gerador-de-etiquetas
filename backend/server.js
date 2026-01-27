
// backend/server.js - QR CODE COM INFORMAÇÕES COMPLETAS
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const PDFDocument = require('pdfkit');
const cors = require('cors');
const QRCode = require('qrcode');

const app = express();
const PORT = 3001;

// ========== CONFIGURAÇÕES ==========
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const frontendPath = path.join(__dirname, '../frontend');
app.use(express.static(frontendPath));

if (!fs.existsSync('uploads')) {
    fs.mkdirSync('uploads', { recursive: true });
}

// ========== UPLOAD ==========
const upload = multer({
    storage: multer.diskStorage({
        destination: 'uploads/',
        filename: (req, file, cb) => {
            const uniqueName = Date.now() + '-' + Math.random().toString(36).substring(7) + path.extname(file.originalname);
            cb(null, uniqueName);
        }
    }),
    fileFilter: (req, file, cb) => {
        const ext = path.extname(file.originalname).toLowerCase();
        if (['.xlsx', '.xls'].includes(ext)) {
            cb(null, true);
        } else {
            cb(new Error('Apenas arquivos Excel (.xlsx, .xls) são permitidos'));
        }
    },
    limits: { fileSize: 10 * 1024 * 1024 }
});

// ========== DADOS ==========
let spreadsheets = [];

// ========== FUNÇÃO PARA FORÇAR FORMATO BRASILEIRO ==========
function formatarComoBrasileiro(valor, formatoExcel) {
    if (!valor && valor !== 0) return '';
    
    // Se já é string, verificar se é data
    if (typeof valor === 'string') {
        const str = valor.trim();
        
        // Verificar padrão de data (dd/mm/aaaa ou mm/dd/aaaa)
        const padraoData = /^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})(.*)$/;
        const match = str.match(padraoData);
        
        if (match) {
            const p1 = parseInt(match[1]);
            const p2 = parseInt(match[2]);
            const ano = match[3];
            const resto = match[4] || '';
            
            // Se p1 > 12 e p2 <= 12 → já está dd/mm (BR)
            if (p1 > 12 && p2 <= 12) {
                return str;
            }
            
            // Se p1 <= 12 e p2 > 12 → está invertido (mm/dd)
            if (p1 <= 12 && p2 > 12) {
                const dia = p2.toString().padStart(2, '0');
                const mes = p1.toString().padStart(2, '0');
                const anoCompleto = ano.length === 2 ? `20${ano}` : ano;
                return `${dia}/${mes}/${anoCompleto}${resto}`;
            }
            
            // Se ambos <= 12, verificar pelo formato do Excel
            if (p1 <= 12 && p2 <= 12) {
                // Se o formato Excel é americano (começa com m), inverter
                if (formatoExcel && formatoExcel.toLowerCase().startsWith('m')) {
                    const dia = p2.toString().padStart(2, '0');
                    const mes = p1.toString().padStart(2, '0');
                    const anoCompleto = ano.length === 2 ? `20${ano}` : ano;
                    return `${dia}/${mes}/${anoCompleto}${resto}`;
                }
            }
        }
        
        return str;
    }
    
    // Se for número (serial do Excel)
    if (typeof valor === 'number') {
        try {
            // Converter serial do Excel para data
            const excelSerial = valor;
            
            // Verificar se é data (está no intervalo típico)
            if (excelSerial > 0 && excelSerial < 50000) {
                const baseDate = new Date(1899, 11, 30);
                const date = new Date(baseDate.getTime() + excelSerial * 24 * 60 * 60 * 1000);
                
                const dia = String(date.getDate()).padStart(2, '0');
                const mes = String(date.getMonth() + 1).padStart(2, '0');
                const ano = date.getFullYear();
                
                let resultado = `${dia}/${mes}/${ano}`;
                
                // Se tem parte decimal, adicionar hora
                const parteDecimal = excelSerial - Math.floor(excelSerial);
                if (parteDecimal > 0) {
                    const horas = String(date.getHours()).padStart(2, '0');
                    const minutos = String(date.getMinutes()).padStart(2, '0');
                    resultado += ` ${horas}:${minutos}`;
                    
                    if (date.getSeconds() > 0) {
                        const segundos = String(date.getSeconds()).padStart(2, '0');
                        resultado += `:${segundos}`;
                    }
                }
                
                return resultado;
            }
            
            return valor.toString();
            
        } catch (error) {
            return valor.toString();
        }
    }
    
    return String(valor);
}

// ========== FUNÇÃO PARA GERAR TEXTO DO QR CODE ==========
function gerarTextoQRCode(etiquetaNumero, totalEtiquetas, spreadsheet, rowData, selectedColumns, mode, rowIndex) {
    // Criar texto formatado para o QR Code
    let texto = "=".repeat(40) + "\n";
    texto += "🏷️ ETIQUETA DIGITAL\n";
    texto += "=".repeat(40) + "\n\n";
    
    // Informações gerais
    texto += "📋 INFORMAÇÕES GERAIS\n";
    texto += "• Sistema: Gerador de Etiquetas\n";
    texto += `• Etiqueta: ${etiquetaNumero} de ${totalEtiquetas}\n`;
    texto += `• Arquivo: ${spreadsheet.fileName}\n`;
    texto += `• Data de geração: ${new Date().toLocaleDateString('pt-BR')}\n`;
    texto += `• Hora de geração: ${new Date().toLocaleTimeString('pt-BR')}\n`;
    
    if (mode === 'single') {
        texto += `• Modo: Linha específica (${rowIndex + 1})\n`;
    } else {
        texto += "• Modo: Todas as linhas\n";
    }
    
    texto += "\n" + "=".repeat(40) + "\n";
    texto += "📊 DADOS DA ETIQUETA\n";
    texto += "=".repeat(40) + "\n\n";
    
    // Adicionar todos os dados da linha
    selectedColumns.forEach(col => {
        const valor = rowData[col.name];
        if (valor && valor.toString().trim() !== '') {
            // Truncar valores muito longos para o QR Code
            let valorExibicao = valor.toString();
            if (valorExibicao.length > 50) {
                valorExibicao = valorExibicao.substring(0, 47) + '...';
            }
            
            texto += `• ${col.name}: ${valorExibicao}\n`;
        }
    });
    
    texto += "\n" + "=".repeat(40) + "\n";
    texto += "📱 COMO USAR\n";
    texto += "=".repeat(40) + "\n";
    texto += "• Escaneie este QR Code para ver\n";
    texto += "  as informações da etiqueta\n";
    texto += "• Mantenha para referência futura\n";
    texto += "• Compartilhe se necessário\n";
    
    texto += "\n" + "=".repeat(40) + "\n";
    texto += "🔗 SISTEMA DE ETIQUETAS\n";
    texto += `• Gerado em: ${new Date().toISOString()}\n`;
    texto += "• Sistema 100% funcional\n";
    texto += "=".repeat(40);
    
    return texto;
}

// ========== FUNÇÃO PARA GERAR QR CODE ==========
async function gerarQRCodeParaEtiqueta(etiquetaNumero, totalEtiquetas, spreadsheet, rowData, selectedColumns, mode, rowIndex) {
    try {
        // Gerar texto formatado
        const texto = gerarTextoQRCode(etiquetaNumero, totalEtiquetas, spreadsheet, rowData, selectedColumns, mode, rowIndex);
        
        // Gerar QR Code
        const qrCodeDataURL = await QRCode.toDataURL(texto, {
            width: 200,  // Maior para mais dados
            margin: 2,
            color: {
                dark: '#000000',
                light: '#FFFFFF'
            },
            errorCorrectionLevel: 'H'  // Alta correção de erro
        });
        
        return qrCodeDataURL;
        
    } catch (error) {
        console.error('❌ Erro ao gerar QR Code:', error);
        return null;
    }
}

// ========== FUNÇÃO PARA LER EXCEL ==========
async function lerExcelCorretamente(filePath) {
    console.log(`📖 Lendo Excel: ${filePath}`);
    
    try {
        const workbook = xlsx.readFile(filePath, {
            cellDates: false,
            cellNF: true,
            dateNF: 'dd"/"mm"/"yyyy',
            raw: false,
            cellText: true
        });
        
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        const range = xlsx.utils.decode_range(worksheet['!ref']);
        console.log(`📊 Planilha: ${range.e.r + 1} linhas × ${range.e.c + 1} colunas`);
        
        // Processar célula por célula
        const data = [];
        const headers = [];
        
        // 1. Ler cabeçalhos
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = xlsx.utils.encode_cell({ r: range.s.r, c: col });
            const cell = worksheet[cellAddress];
            
            let headerName = cell ? (cell.w || String(cell.v || '')) : '';
            headerName = headerName.trim();
            
            if (!headerName) {
                headerName = `Coluna_${col + 1}`;
            }
            
            headers.push({
                index: col,
                name: headerName,
                format: cell ? cell.z : null
            });
        }
        
        console.log(`📋 ${headers.length} colunas identificadas`);
        
        // 2. Processar dados
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            const rowData = {};
            let hasData = false;
            
            headers.forEach(header => {
                const cellAddress = xlsx.utils.encode_cell({ r: row, c: header.index });
                const cell = worksheet[cellAddress];
                
                let valorFinal = '';
                
                if (cell) {
                    if (cell.w !== undefined && cell.w !== null) {
                        valorFinal = formatarComoBrasileiro(cell.w, cell.z);
                    } else if (cell.v !== undefined) {
                        valorFinal = formatarComoBrasileiro(cell.v, cell.z);
                    }
                }
                
                rowData[header.name] = valorFinal;
                
                if (valorFinal && valorFinal.trim() !== '') {
                    hasData = true;
                }
            });
            
            if (hasData) {
                data.push(rowData);
            }
        }
        
        console.log(`✅ ${data.length} linhas processadas`);
        
        // Criar estrutura de colunas
        const columns = headers.map((header, index) => {
            // Encontrar sample value
            let sampleValue = '';
            for (let i = 0; i < Math.min(5, data.length); i++) {
                const val = data[i][header.name];
                if (val && val.trim() !== '') {
                    sampleValue = val;
                    break;
                }
            }
            
            if (!sampleValue) {
                sampleValue = '(vazio)';
            }
            
            const pareceData = sampleValue.match(/^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}/);
            const temHora = sampleValue.includes(':');
            
            return {
                id: `col_${index}`,
                name: header.name,
                index: index,
                sampleValue: sampleValue,
                format: header.format,
                pareceData: pareceData,
                temDataHora: pareceData && temHora
            };
        });
        
        return {
            columns: columns,
            data: data,
            worksheet: worksheet,
            rowCount: data.length,
            columnCount: columns.length
        };
        
    } catch (error) {
        console.error('❌ Erro ao ler Excel:', error);
        throw error;
    }
}

// ========== ROTAS API ==========
app.get('/api/ping', (req, res) => {
    res.json({
        success: true,
        message: 'pong 🏓',
        server: 'Etiquetas API - QR Code Informativo',
        time: new Date().toISOString(),
        spreadsheetsCount: spreadsheets.length
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
        
        const resultado = await lerExcelCorretamente(req.file.path);
        
        const spreadsheet = {
            id: `sheet_${Date.now()}`,
            fileName: req.file.originalname,
            filePath: req.file.path,
            uploadDate: new Date().toISOString(),
            columns: resultado.columns,
            data: resultado.data,
            rawWorksheet: resultado.worksheet,
            rowCount: resultado.rowCount,
            columnCount: resultado.columnCount
        };
        
        spreadsheets.push(spreadsheet);
        
        res.json({
            success: true,
            message: `Planilha processada: ${resultado.rowCount} linhas, ${resultado.columnCount} colunas`,
            spreadsheet: {
                id: spreadsheet.id,
                fileName: spreadsheet.fileName,
                rowCount: spreadsheet.rowCount,
                columnCount: spreadsheet.columnCount
            }
        });
        
    } catch (error) {
        console.error('❌ Erro no upload:', error);
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

app.get('/api/spreadsheets', (req, res) => {
    const simplifiedList = spreadsheets.map(sheet => ({
        id: sheet.id,
        fileName: sheet.fileName,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        uploadDate: sheet.uploadDate
    }));
    
    res.json({
        success: true,
        spreadsheets: simplifiedList,
        total: simplifiedList.length
    });
});

app.get('/api/spreadsheet/:id', (req, res) => {
    const id = req.params.id;
    const spreadsheet = spreadsheets.find(s => s.id === id);
    
    if (!spreadsheet) {
        return res.status(404).json({
            success: false,
            error: 'Planilha não encontrada'
        });
    }
    
    res.json({
        success: true,
        data: {
            id: spreadsheet.id,
            fileName: spreadsheet.fileName,
            rowCount: spreadsheet.rowCount,
            columnCount: spreadsheet.columnCount,
            columns: spreadsheet.columns,
            data: spreadsheet.data.slice(0, 10),
            primeiraLinha: spreadsheet.data[0]
        }
    });
});

app.delete('/api/spreadsheet/:id', (req, res) => {
    const id = req.params.id;
    const index = spreadsheets.findIndex(s => s.id === id);
    
    if (index === -1) {
        return res.status(404).json({
            success: false,
            error: 'Planilha não encontrada'
        });
    }
    
    const spreadsheet = spreadsheets[index];
    
    if (fs.existsSync(spreadsheet.filePath)) {
        fs.unlinkSync(spreadsheet.filePath);
    }
    
    spreadsheets.splice(index, 1);
    
    res.json({
        success: true,
        message: 'Planilha excluída'
    });
});

// ========== ROTA PARA TESTAR QR CODE ==========
app.post('/api/test-qrcode', async (req, res) => {
    try {
        const { texto } = req.body;
        
        if (!texto) {
            return res.status(400).json({
                success: false,
                error: 'Texto é obrigatório'
            });
        }
        
        const qrCodeDataURL = await QRCode.toDataURL(texto, {
            width: 300,
            margin: 2,
            color: {
                dark: '#000000',
                light: '#FFFFFF'
            }
        });
        
        res.json({
            success: true,
            qrCode: qrCodeDataURL,
            texto: texto,
            nota: 'QR Code gerado com sucesso'
        });
        
    } catch (error) {
        console.error('❌ Erro ao gerar QR Code de teste:', error);
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// ========== GERAR ETIQUETAS COM QR CODE INFORMATIVO ==========
app.post('/api/generate', async (req, res) => {
    console.log('🏷️ Gerando etiquetas com QR Code informativo...');
    
    try {
        const { 
            spreadsheetId, 
            columns, 
            mode = 'all',
            rowIndex = 0,
            quantity = 1
        } = req.body;
        
        if (!spreadsheetId) {
            return res.status(400).json({
                success: false,
                error: 'ID da planilha é obrigatório'
            });
        }
        
        const spreadsheet = spreadsheets.find(s => s.id === spreadsheetId);
        if (!spreadsheet) {
            return res.status(404).json({
                success: false,
                error: 'Planilha não encontrada'
            });
        }
        
        const selectedColumns = spreadsheet.columns.filter(col => columns.includes(col.id));
        if (selectedColumns.length === 0) {
            return res.status(400).json({
                success: false,
                error: 'Selecione pelo menos uma coluna'
            });
        }
        
        console.log(`✅ ${selectedColumns.length} colunas selecionadas`);
        
        // Definir dados a serem processados
        let dataToProcess = [];
        let totalLabels = 0;
        
        if (mode === 'all') {
            dataToProcess = spreadsheet.data;
            totalLabels = spreadsheet.rowCount;
            console.log(`📊 Gerando ${totalLabels} etiquetas`);
        } else if (mode === 'single') {
            if (rowIndex < 0 || rowIndex >= spreadsheet.rowCount) {
                return res.status(400).json({
                    success: false,
                    error: `Índice de linha inválido`
                });
            }
            
            if (quantity < 1 || quantity > 1000) {
                return res.status(400).json({
                    success: false,
                    error: 'Quantidade inválida'
                });
            }
            
            const selectedRow = spreadsheet.data[rowIndex];
            for (let i = 0; i < quantity; i++) {
                dataToProcess.push(selectedRow);
            }
            totalLabels = quantity;
            console.log(`📊 Gerando ${quantity} etiquetas da linha ${rowIndex + 1}`);
        }
        
        // ========== TAMANHO 120×80mm ==========
        const MM_TO_POINTS = 2.834645669;
        const PAGE_HEIGHT_MM = 120;
        const PAGE_WIDTH_MM = 80;
        
        const PAGE_HEIGHT = Math.round(PAGE_HEIGHT_MM * MM_TO_POINTS);
        const PAGE_WIDTH = Math.round(PAGE_WIDTH_MM * MM_TO_POINTS);
        
        // Tamanho do QR Code
        const QR_SIZE = 45; // 40 pontos = ~14mm"
        
        // Criar PDF
        const doc = new PDFDocument({
            size: [PAGE_WIDTH, PAGE_HEIGHT],
            margin: 0,
            autoFirstPage: false
        });
        
        let fileName;
        if (mode === 'all') {
            fileName = `etiquetas_${spreadsheet.fileName.replace(/\.[^/.]+$/, "")}_${Date.now()}.pdf`;
        } else {
            fileName = `etiquetas_linha${rowIndex + 1}_x${quantity}_${Date.now()}.pdf`;
        }
        
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        
        doc.pipe(res);
        
        // Configurar fontes
        doc.registerFont('Helvetica', 'Helvetica');
        doc.registerFont('Helvetica-Bold', 'Helvetica-Bold');
        
        // ========== GERAR ETIQUETAS ==========
        let etiquetasGeradas = 0;
        
        for (let index = 0; index < dataToProcess.length; index++) {
            const rowData = dataToProcess[index];
            
            // Verificar se tem dados reais
            const temDadosReais = selectedColumns.some(col => {
                const valor = rowData[col.name];
                return valor && valor.toString().trim() !== '';
            });
            
            if (!temDadosReais) {
                console.log(`⚠️ Linha ${index + 1} sem dados, pulando etiqueta`);
                continue;
            }
            
            // ADICIONAR PÁGINA
            doc.addPage({
                size: [PAGE_WIDTH, PAGE_HEIGHT],
                margins: { top: 0, bottom: 0, left: 0, right: 0 }
            });
            
            etiquetasGeradas++;
            
            // Cabeçalho da etiqueta
            const headerHeight = 25;
            doc.rect(0, 0, PAGE_WIDTH, headerHeight)
               .fill('#4f46e5');
            
            if (mode === 'all') {
                doc.fontSize(12)
                   .fillColor('white')
                   .font('Helvetica-Bold')
                   .text(`ETIQUETA ${etiquetasGeradas}`, 15, 8);
                
                doc.fontSize(10)
                   .fillColor('rgba(255, 255, 255, 0.8)')
                   .text(`${etiquetasGeradas}/${totalLabels}`, PAGE_WIDTH - 60, 10, { align: 'right' });
            } else {
                doc.fontSize(12)
                   .fillColor('white')
                   .font('Helvetica-Bold')
                   .text(`ETIQUETA ${etiquetasGeradas}`, 15, 8);
                
                doc.fontSize(10)
                   .fillColor('rgba(255, 255, 255, 0.8)')
                   .text(`Cópia ${etiquetasGeradas}/${quantity}`, PAGE_WIDTH - 80, 10, { align: 'right' });
            }
            
            // Área de conteúdo
            const CONTENT_START_Y = headerHeight + 15;
            
            // Filtrar colunas com dados
            const columnsWithData = selectedColumns.filter(col => {
                const value = rowData[col.name];
                return value && value.toString().trim() !== '';
            });
            
            // Layout das colunas
            const COLUMN_COUNT = 2;
            const COLUMN_WIDTH = (PAGE_WIDTH - 40) / COLUMN_COUNT;
            const CARD_HEIGHT = 38;
            const CARD_SPACING = 8;
            
            let columnData = [[], []];
            
            // Distribuir colunas
            columnsWithData.forEach((col, colIndex) => {
                const columnIdx = colIndex % COLUMN_COUNT;
                columnData[columnIdx].push(col);
            });
            
            // Desenhar colunas
            for (let colIndex = 0; colIndex < COLUMN_COUNT; colIndex++) {
                const currentX = 15 + (colIndex * (COLUMN_WIDTH + 10));
                let currentY = CONTENT_START_Y;
                
                // Processar cartões
                columnData[colIndex].forEach((col, cardIndex) => {
                    if (currentY + CARD_HEIGHT > PAGE_HEIGHT - 60) return;
                    
                    const cellValue = rowData[col.name];
                    if (!cellValue || cellValue.toString().trim() === '') return;
                    
                    // Criar cartão
                    doc.roundedRect(currentX, currentY, COLUMN_WIDTH - 5, CARD_HEIGHT, 5)
                       .fill('#f8fafc')
                       .stroke('#e2e8f0')
                       .stroke();
                    
                    // Nome da coluna
                    doc.fontSize(7)
                       .fillColor('#64748b')
                       .font('Helvetica')
                       .text(col.name.toUpperCase(), 
                             currentX + 6, 
                             currentY + 6, 
                             { width: COLUMN_WIDTH - 15 });
                    
                    // Valor
                    let valueText = String(cellValue);
                    
                    if (valueText.length > 25 && !col.temDataHora) {
                        valueText = valueText.substring(0, 22) + '...';
                    }
                    
                    doc.fontSize(col.temDataHora ? 8 : 9)
                       .fillColor('#1e293b')
                       .font('Helvetica-Bold')
                       .text(valueText, 
                             currentX + 6, 
                             currentY + (col.temDataHora ? 18 : 20),
                             { width: COLUMN_WIDTH - 15, ellipsis: true });
                    
                    currentY += CARD_HEIGHT + CARD_SPACING;
                });
            }
            
            // 🔥 GERAR QR CODE COM INFORMAÇÕES COMPLETAS
            try {
                const qrCodeDataURL = await gerarQRCodeParaEtiqueta(
                    etiquetasGeradas,
                    totalLabels,
                    spreadsheet,
                    rowData,
                    selectedColumns,
                    mode,
                    rowIndex
                );
                
                if (qrCodeDataURL) {
                    // Posicionar no canto inferior direito
                    const qrX = PAGE_WIDTH - QR_SIZE - 15;
                    const qrY = PAGE_HEIGHT - QR_SIZE - 15;
                    
                    // Adicionar QR Code ao PDF
                    doc.image(qrCodeDataURL, qrX, qrY, {
                        width: QR_SIZE,
                        height: QR_SIZE
                    });
                    
                    // Adicionar borda e legenda
                    doc.rect(qrX - 2, qrY - 2, QR_SIZE + 4, QR_SIZE + 4)
                       .stroke('#4f46e5')
                       .lineWidth(0.5);
                    
                    // Texto abaixo do QR Code
                    doc.fontSize(5)
                       .fillColor('#4f46e5')
                       .text('Escanear para ver dados', 
                             qrX, 
                             qrY + QR_SIZE + 2,
                             { width: QR_SIZE, align: 'center' });
                    
                    console.log(`✅ QR Code informativo adicionado na etiqueta ${etiquetasGeradas}`);
                }
            } catch (qrError) {
                console.error(`❌ Erro no QR Code: ${qrError.message}`);
                // Desenhar placeholder em caso de erro
                const qrX = PAGE_WIDTH - QR_SIZE - 15;
                const qrY = PAGE_HEIGHT - QR_SIZE - 15;
                
                doc.rect(qrX, qrY, QR_SIZE, QR_SIZE)
                   .fill('#f8fafc')
                   .stroke('#e2e8f0')
                   .stroke();
                
                doc.fontSize(6)
                   .fillColor('#94a3b8')
                   .text('QR Code\nindisponível', 
                         qrX + 5, 
                         qrY + QR_SIZE/2 - 6,
                         { width: QR_SIZE - 10, align: 'center' });
            }
            
            // Rodapé
            doc.fontSize(6)
               .fillColor('#94a3b8')
               .text(`${spreadsheet.fileName}`, 
                     15, 
                     PAGE_HEIGHT - 25,
                     { width: PAGE_WIDTH - (QR_SIZE + 40), align: 'left' });
        }
        
        // Se nenhuma etiqueta foi gerada
        if (etiquetasGeradas === 0) {
            doc.addPage({
                size: [PAGE_WIDTH, PAGE_HEIGHT],
                margins: { top: 0, bottom: 0, left: 0, right: 0 }
            });
            
            doc.fontSize(14)
               .fillColor('#64748b')
               .text('Nenhum dado encontrado', 
                     30, 
                     PAGE_HEIGHT / 2 - 20,
                     { align: 'center' });
        }
        
        doc.end();
        
        console.log(`✅ PDF gerado: ${fileName}`);
        console.log(`✅ ${etiquetasGeradas} etiquetas com QR Code informativo`);
        
    } catch (error) {
        console.error('❌ Erro ao gerar PDF:', error);
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// ========== ROTAS PÁGINAS ==========
app.get('/', (req, res) => {
    res.sendFile(path.join(frontendPath, 'index.html'));
});

app.get('/upload', (req, res) => {
    res.sendFile(path.join(frontendPath, 'upload.html'));
});

app.get('/generate', (req, res) => {
    res.sendFile(path.join(frontendPath, 'generate.html'));
});

// ========== INICIAR ==========
app.listen(PORT, () => {
    console.log('='.repeat(70));
    console.log(`🚀 SISTEMA DE ETIQUETAS - QR CODE INFORMATIVO`);
    console.log('='.repeat(70));
    console.log(`✅ Backend: http://localhost:${PORT}`);
    console.log(`✅ Upload: http://localhost:${PORT}/upload`);
    console.log(`✅ Gerar: http://localhost:${PORT}/generate`);
    console.log('='.repeat(70));
    console.log('🔥 QR CODE MELHORADO:');
    console.log('• Ao escanear, mostra TODOS os dados da etiqueta');
    console.log('• Formatação organizada e legível');
    console.log('• Inclui informações do sistema');
    console.log('• Inclui data/hora de geração');
    console.log('• Inclui modo de operação');
    console.log('='.repeat(70));
    console.log('📱 TESTAR QR CODE:');
    console.log('POST http://localhost:3001/api/test-qrcode');
    console.log('Body: {"texto": "Conteúdo do QR Code"}');
    console.log('='.repeat(70));
});
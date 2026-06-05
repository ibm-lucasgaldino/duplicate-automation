// ===== HAMBURGER MENU FUNCTIONALITY =====
document.addEventListener('DOMContentLoaded', function() {
    const hamburgerBtn = document.getElementById('hamburgerBtn');
    const sidebar = document.getElementById('sidebar');
    const sidebarOverlay = document.getElementById('sidebarOverlay');
    const closeBtn = document.getElementById('closeBtn');

    // Open sidebar
    hamburgerBtn.addEventListener('click', function() {
        sidebar.classList.add('active');
        sidebarOverlay.classList.add('active');
    });

    // Close sidebar
    closeBtn.addEventListener('click', function() {
        sidebar.classList.remove('active');
        sidebarOverlay.classList.remove('active');
    });

    // Close sidebar when clicking overlay
    sidebarOverlay.addEventListener('click', function() {
        sidebar.classList.remove('active');
        sidebarOverlay.classList.remove('active');
    });

    // Close sidebar on ESC key
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && sidebar.classList.contains('active')) {
            sidebar.classList.remove('active');
            sidebarOverlay.classList.remove('active');
        }
    });
});

// ===== EXISTING CODE BELOW =====

// Função para alternar entre tabs
function openTab(event, tabName) {
    // Esconder todos os conteúdos das tabs
    const tabContents = document.getElementsByClassName('tab-content');
    for (let i = 0; i < tabContents.length; i++) {
        tabContents[i].classList.remove('active');
    }

    // Remover classe active de todos os botões
    const tabButtons = document.getElementsByClassName('tab-button');
    for (let i = 0; i < tabButtons.length; i++) {
        tabButtons[i].classList.remove('active');
    }

    // Mostrar o conteúdo da tab selecionada
    document.getElementById(tabName).classList.add('active');
    
    // Adicionar classe active ao botão clicado
    event.currentTarget.classList.add('active');
}

// Creator ID to Name Mapping
const creatorMapping = {
    '015617631': 'Lucas',
    '077092631': 'Rodrigo',
    '118046631': 'Ohana',
    '116645631': 'Loane',
    '017952631': 'Marcel',
    '075517631': 'Gabriel Borges',
    '117681631': 'Erica',
    '063683631': 'Paulo',
    '000944631': 'Julia',
    '09029U744': 'Hamsa',
    '005I8O744': 'Ragu T',
    '003MNS744': 'D Reddy',
    '06225W744': 'Sujinraja R'
};

// Allowed Creator IDs
const allowedCreatorIDs = Object.keys(creatorMapping);

// Variáveis globais
let irFile = null;
let workbookData = null;

// Função para lidar com o upload do arquivo IR
function handleIRFileUpload(event) {
    const file = event.target.files[0];
    
    if (!file) {
        return;
    }

    // Verificar se é um arquivo Excel
    const validExtensions = ['.xlsx', '.xls'];
    const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    
    if (!validExtensions.includes(fileExtension)) {
        alert('Por favor, selecione um arquivo Excel válido (.xlsx ou .xls)');
        event.target.value = '';
        return;
    }

    // Armazenar o arquivo
    irFile = file;

    // Exibir informações do arquivo
    displayIRFileInfo(file);

    // Processar o arquivo
    processIRFile(file);
}

// Função para exibir informações do arquivo
function displayIRFileInfo(file) {
    const fileInfo = document.getElementById('irFileInfo');
    const fileName = document.getElementById('irFileName');
    const fileSize = document.getElementById('irFileSize');

    // Formatar tamanho do arquivo
    const sizeInKB = (file.size / 1024).toFixed(2);
    const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
    const formattedSize = file.size > 1024 * 1024 
        ? `${sizeInMB} MB` 
        : `${sizeInKB} KB`;

    fileName.textContent = file.name;
    fileSize.textContent = formattedSize;
    fileInfo.style.display = 'block';
}

// Função para limpar o arquivo selecionado
function clearIRFile() {
    irFile = null;
    workbookData = null;
    document.getElementById('irFileInput').value = '';
    document.getElementById('irFileInfo').style.display = 'none';
    document.getElementById('irProcessing').style.display = 'none';
    document.getElementById('irResults').style.display = 'none';
}

// Função para processar o arquivo IR
function processIRFile(file) {
    const processingSection = document.getElementById('irProcessing');
    const resultsSection = document.getElementById('irResults');
    
    processingSection.style.display = 'block';
    resultsSection.style.display = 'none';
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Get the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON with header row
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            workbookData = jsonData;
            
            // Prompt user for reference month
            promptForReferenceMonth();
            
        } catch (error) {
            processingSection.style.display = 'none';
            alert('Erro ao processar o arquivo Excel: ' + error.message);
            console.error('Error processing Excel:', error);
        }
    };
    
    reader.onerror = function(error) {
        processingSection.style.display = 'none';
        alert('Erro ao ler o arquivo: ' + error);
        console.error('Error reading file:', error);
    };
    
    reader.readAsArrayBuffer(file);
}

function promptForReferenceMonth() {
    const processingSection = document.getElementById('irProcessing');
    processingSection.style.display = 'none';
    
    // Create modal for month input
    const modal = document.createElement('div');
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.5);
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 1000;
    `;
    
    const modalContent = document.createElement('div');
    modalContent.style.cssText = `
        background: white;
        padding: 30px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        max-width: 400px;
        width: 90%;
    `;
    
    modalContent.innerHTML = `
        <h3 style="margin-top: 0; color: #667eea;">Mês de Referência</h3>
        <p style="margin-bottom: 20px; color: #666;">Informe o mês de referência para consulta de duplicidades (Payment Date):</p>
        <input type="month" id="referenceMonth" style="width: 100%; padding: 10px; border: 2px solid #667eea; border-radius: 5px; font-size: 16px; margin-bottom: 20px;">
        <div style="display: flex; gap: 10px; justify-content: flex-end;">
            <button onclick="cancelMonthSelection()" style="padding: 10px 20px; background: #ccc; border: none; border-radius: 5px; cursor: pointer; font-size: 14px;">Cancelar</button>
            <button onclick="confirmMonthSelection()" style="padding: 10px 20px; background: #667eea; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 14px;">Confirmar</button>
        </div>
    `;
    
    modal.appendChild(modalContent);
    document.body.appendChild(modal);
    
    // Store modal reference
    window.currentModal = modal;
    
    // Focus on input
    setTimeout(() => {
        document.getElementById('referenceMonth').focus();
    }, 100);
}

function cancelMonthSelection() {
    if (window.currentModal) {
        document.body.removeChild(window.currentModal);
        window.currentModal = null;
    }
    clearIRFile();
}

function confirmMonthSelection() {
    const monthInput = document.getElementById('referenceMonth');
    const referenceMonth = monthInput.value;
    
    if (!referenceMonth) {
        alert('Por favor, selecione um mês de referência.');
        return;
    }
    
    // Remove modal
    if (window.currentModal) {
        document.body.removeChild(window.currentModal);
        window.currentModal = null;
    }
    
    // Process duplicates with the selected month
    processDuplicates(referenceMonth);
}

function processDuplicates(referenceMonth) {
    const processingSection = document.getElementById('irProcessing');
    const resultsSection = document.getElementById('irResults');
    
    processingSection.style.display = 'block';
    
    setTimeout(() => {
        try {
            const duplicates = findDuplicates(referenceMonth);
            displayResults(duplicates);
            
            processingSection.style.display = 'none';
            resultsSection.style.display = 'block';
        } catch (error) {
            processingSection.style.display = 'none';
            alert('Erro ao processar duplicatas: ' + error.message);
            console.error('Error processing duplicates:', error);
        }
    }, 500);
}

function findDuplicates(referenceMonth) {
    if (!workbookData || workbookData.length < 2) {
        throw new Error('Dados do Excel não encontrados ou inválidos');
    }
    
    const [year, month] = referenceMonth.split('-').map(Number);
    const results = [];
    duplicateRowNumbers = []; // Reset global variable
    referenceMonthGlobal = { year, month }; // Store reference month globally
    
    console.log('Reference Year:', year, 'Month:', month);
    console.log('Total rows:', workbookData.length);
    
    // Skip header row (index 0), start from row 1
    for (let i = 1; i < workbookData.length; i++) {
        const row = workbookData[i];
        
        // Column indices (0-based): A=0, B=1, C=2, Q=16, V=21, W=22, X=23, Y=24
        const billingDoc = row[0] || '';
        const paymentDate = row[1] || '';
        const sapOrder = row[2] || '';
        const reason = row[16] || '';
        const creatorID = String(row[21] || '').trim();
        const possibleDuplicates = String(row[22] || '').toLowerCase().trim();
        const columnX = String(row[23] || '').toLowerCase().trim();
        const columnY = String(row[24] || '').toLowerCase().trim();
        
        // Debug specific rows
        if (i + 1 === 4686 || i + 1 === 4698 || i + 1 === 4699 || i + 1 === 8301) {
            console.log(`Row ${i + 1}:`, {
                billingDoc,
                paymentDate,
                creatorID,
                possibleDuplicates,
                columnX,
                columnY,
                matchesMonth: matchesReferenceMonth(paymentDate, year, month),
                isAllowedCreator: allowedCreatorIDs.includes(creatorID),
                hasPossibleDuplicate: possibleDuplicates.includes('possible duplicate'),
                hasNotDuplicateX: columnX.includes('not duplicate'),
                hasNotDuplicateY: columnY.includes('not duplicate')
            });
        }
        
        // Filter 1: Check if payment date matches reference month
        if (!matchesReferenceMonth(paymentDate, year, month)) {
            continue;
        }
        
        // Filter 2: Check if Creator ID is in allowed list
        if (!allowedCreatorIDs.includes(creatorID)) {
            continue;
        }
        
        // Filter 3: Check Possible Duplicates (Column W) - must contain "possible duplicate"
        if (!possibleDuplicates.includes('possible duplicate')) {
            continue;
        }
        
        // Filter 4: Exclude rows with "Not Duplicate" in Column X
        if (columnX.includes('not duplicate')) {
            continue;
        }
        
        // Filter 5: Exclude rows with "Not Duplicate" in Column Y
        if (columnY.includes('not duplicate')) {
            continue;
        }
        
        // This row is a potential duplicate
        const rowNumber = i + 1;
        duplicateRowNumbers.push(rowNumber); // Store row number
        
        results.push({
            rowNumber: rowNumber, // Excel row number (1-based)
            billingDoc: billingDoc,
            sapOrder: sapOrder,
            reason: reason,
            creatorID: creatorID,
            creatorName: creatorMapping[creatorID] || 'Desconhecido'
        });
    }
    
    console.log('Results found:', results.length);
    return results;
}

function matchesReferenceMonth(dateValue, refYear, refMonth) {
    if (!dateValue) return false;
    
    // Try to parse the date
    let date;
    
    // Check if it's an Excel serial number
    if (typeof dateValue === 'number') {
        date = XLSX.SSF.parse_date_code(dateValue);
        if (date) {
            console.log('Parsed date:', date, 'Expected:', refYear, refMonth);
            return date.y === refYear && date.m === refMonth;
        }
    }
    
    // Try to parse as string date
    if (typeof dateValue === 'string') {
        // Try various date formats
        const parsedDate = new Date(dateValue);
        if (!isNaN(parsedDate.getTime())) {
            return parsedDate.getFullYear() === refYear && (parsedDate.getMonth() + 1) === refMonth;
        }
    }
    
    return false;
}

function displayResults(duplicates) {
    const resultsContent = document.getElementById('irResultsContent');
    
    // IDs to highlight
    const highlightIDs = [
        '015617631',
        '077092631',
        '118046631',
        '116645631',
        '017952631',
        '075517631',
        '117681631',
        '063683631',
        '000944631',
        '09029U744',
        '005I8O744',
        '003MNS744',
        '06225W744'
    ];
    
    // Reason codes to highlight
    const highlightReasons = ['Z08', 'Z14', 'Z16', 'Z10'];
    
    if (duplicates.length === 0) {
        resultsContent.innerHTML = `
            <div style="text-align: center; padding: 40px; color: #28a745;">
                <h3>✓ Nenhuma duplicata encontrada!</h3>
                <p>Não foram encontradas possíveis duplicatas com os critérios especificados.</p>
                <button onclick="generateExcelWithComments()" style="margin-top: 20px; padding: 12px 24px; background: #667eea; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; font-weight: bold;">
                    📥 Gerar Excel com Comentários
                </button>
            </div>
        `;
        return;
    }
    
    let html = `
        <div style="margin-bottom: 20px;">
            <h3 style="color: #667eea; margin-bottom: 10px;">Possíveis Duplicatas Encontradas: ${duplicates.length}</h3>
            <p style="color: #666; margin-bottom: 20px;">Insira comentários para cada linha identificada como possível duplicata:</p>
            <div style="display: flex; gap: 20px; margin-bottom: 15px; flex-wrap: wrap;">
                <div style="display: flex; align-items: center; gap: 8px;">
                    <div style="width: 20px; height: 20px; background: #fff9c4; border: 2px solid #ffc107; border-radius: 3px;"></div>
                    <span style="font-size: 14px; color: #666;">Creator ID destacado</span>
                </div>
                <div style="display: flex; align-items: center; gap: 8px;">
                    <div style="width: 20px; height: 20px; background: #d4edda; border: 2px solid #28a745; border-radius: 3px;"></div>
                    <span style="font-size: 14px; color: #666;">Reason Code destacado (Z08/Z14/Z16/Z10)</span>
                </div>
            </div>
        </div>
        <div style="overflow-x: auto;">
            <table style="width: 100%; border-collapse: collapse; background: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                <thead>
                    <tr style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;">
                        <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Linha Excel</th>
                        <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Billing Document</th>
                        <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">SAP Order</th>
                        <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Reason</th>
                        <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Creator ID</th>
                        <th style="padding: 12px; text-align: left; border: 1px solid #ddd;">Creator Name</th>
                        <th style="padding: 12px; text-align: left; border: 1px solid #ddd; min-width: 250px;">Comentário</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    duplicates.forEach((dup, index) => {
        const isHighlightedCreator = highlightIDs.includes(dup.creatorID);
        const isHighlightedReason = highlightReasons.includes(dup.reason);
        let rowStyle;
        
        if (isHighlightedCreator) {
            rowStyle = 'background: #fff9c4; border-left: 4px solid #ffc107;';
        } else {
            rowStyle = index % 2 === 0 ? 'background: #f8f9fa;' : 'background: white;';
        }
        
        const reasonStyle = isHighlightedReason
            ? 'padding: 12px; border: 1px solid #ddd; background: #d4edda; font-weight: bold; color: #155724;'
            : 'padding: 12px; border: 1px solid #ddd;';
        
        html += `
            <tr style="${rowStyle}">
                <td style="padding: 12px; border: 1px solid #ddd; font-weight: bold; color: #667eea;">${dup.rowNumber}</td>
                <td style="padding: 12px; border: 1px solid #ddd;">${dup.billingDoc}</td>
                <td style="padding: 12px; border: 1px solid #ddd;">${dup.sapOrder}</td>
                <td style="${reasonStyle}">${dup.reason}</td>
                <td style="padding: 12px; border: 1px solid #ddd; ${isHighlightedCreator ? 'font-weight: bold; color: #f57c00;' : ''}">${dup.creatorID}</td>
                <td style="padding: 12px; border: 1px solid #ddd; font-weight: 500;">${dup.creatorName}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">
                    <input type="text"
                           id="comment_${dup.rowNumber}"
                           data-row="${dup.rowNumber}"
                           placeholder="Insira o comentário..."
                           style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px;">
                </td>
            </tr>
        `;
    });
    
    html += `
                </tbody>
            </table>
        </div>
        <div style="margin-top: 20px; padding: 15px; background: #d1ecf1; border-left: 4px solid #0c5460; border-radius: 5px;">
            <strong>ℹ️ Informação:</strong> Após inserir os comentários, clique no botão abaixo para gerar o arquivo Excel atualizado. As linhas sem comentário e que não são duplicatas receberão automaticamente o comentário: "Not duplicate according QMF report or out of SW VAR validation criteria."
        </div>
        <div style="margin-top: 20px; text-align: center;">
            <button onclick="generateExcelWithComments()" style="padding: 12px 24px; background: #667eea; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; font-weight: bold; box-shadow: 0 2px 4px rgba(0,0,0,0.2);">
                📥 Gerar Excel com Comentários
            </button>
        </div>
    `;
    
    resultsContent.innerHTML = html;
}

// Adicionar suporte para drag and drop
document.addEventListener('DOMContentLoaded', function() {
    const fileUploadLabel = document.querySelector('.file-upload-label');
    
    if (fileUploadLabel) {
        // Prevenir comportamento padrão para drag and drop
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            fileUploadLabel.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        // Destacar área de drop quando arquivo está sobre ela
        ['dragenter', 'dragover'].forEach(eventName => {
            fileUploadLabel.addEventListener(eventName, highlight, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            fileUploadLabel.addEventListener(eventName, unhighlight, false);
        });

        function highlight(e) {
            fileUploadLabel.style.borderColor = '#667eea';
            fileUploadLabel.style.background = '#edf2f7';
        }

        function unhighlight(e) {
            fileUploadLabel.style.borderColor = '#cbd5e0';
            fileUploadLabel.style.background = '#f7fafc';
        }

        // Lidar com o drop do arquivo
        fileUploadLabel.addEventListener('drop', handleDrop, false);

        function handleDrop(e) {
            const dt = e.dataTransfer;
            const files = dt.files;

            if (files.length > 0) {
                const fileInput = document.getElementById('irFileInput');
                fileInput.files = files;
                
                // Disparar evento de change manualmente
                const event = new Event('change', { bubbles: true });
                fileInput.dispatchEvent(event);
            }
        }
    }
});

// Made with Bob

// Global variables to store duplicate row numbers and reference month
let duplicateRowNumbers = [];
let referenceMonthGlobal = null;

function generateExcelWithComments() {
    if (!workbookData || workbookData.length === 0) {
        alert('Nenhum dado disponível para gerar o Excel.');
        return;
    }
    
    // Collect comments from input fields
    const comments = {};
    const hasManualComment = {};
    duplicateRowNumbers.forEach(rowNum => {
        const input = document.getElementById(`comment_${rowNum}`);
        if (input) {
            const comment = input.value.trim();
            comments[rowNum] = comment;
            hasManualComment[rowNum] = comment.length > 0;
        }
    });
    
    // Create a copy of the workbook data
    const updatedData = workbookData.map((row, index) => {
        if (index === 0) {
            // Header row - add title to column Z
            const headerRow = [...row];
            // Ensure row has enough columns (at least 26 for column Z)
            while (headerRow.length < 26) {
                headerRow.push('');
            }
            headerRow[25] = 'BPSO COMMENTS'; // Column Z
            return headerRow;
        }
        
        const rowNumber = index + 1;
        const newRow = [...row];
        
        // Ensure row has enough columns (at least 26 for column Z)
        while (newRow.length < 26) {
            newRow.push('');
        }
        
        // Column Z index is 25 (0-based)
        if (duplicateRowNumbers.includes(rowNumber)) {
            // This is a duplicate row - use user's comment if provided
            if (comments[rowNumber]) {
                newRow[25] = comments[rowNumber];
            }
        } else {
            // Not a duplicate - add default comment if column Z is empty
            const currentComment = String(newRow[25] || '').trim();
            if (!currentComment) {
                // Check if this row is in the reference month
                const paymentDate = newRow[1]; // Column B (Payment Date)
                const isInReferencePeriod = referenceMonthGlobal &&
                    matchesReferenceMonth(paymentDate, referenceMonthGlobal.year, referenceMonthGlobal.month);
                
                if (isInReferencePeriod) {
                    newRow[25] = 'Not duplicate according QMF report or out of SW VAR validation criteria.';
                } else {
                    newRow[25] = 'Line not analyse in this report period.';
                }
            }
        }
        
        return newRow;
    });
    
    // Create new workbook
    const ws = XLSX.utils.aoa_to_sheet(updatedData);
    
    // Mark cells with manual comments by adding a prefix
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
        const rowNumber = R + 1;
        if (hasManualComment[rowNumber]) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: 25 }); // Column Z
            if (ws[cellAddress] && ws[cellAddress].v) {
                // Add a marker to indicate manual comment (you can format this in Excel later)
                ws[cellAddress].v = '⭐ ' + ws[cellAddress].v;
            }
        }
    }
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    
    // Generate file name with timestamp
    const now = new Date();
    const timestamp = now.toISOString().replace(/[:.]/g, '-').slice(0, -5);
    const fileName = `IR_Duplicates_Updated_${timestamp}.xlsx`;
    
    // Download the file
    XLSX.writeFile(wb, fileName);
    
    alert('✅ Arquivo Excel gerado com sucesso!\n\nOs comentários foram adicionados na coluna Z:\n- Linhas duplicatas com comentário manual: marcadas com ⭐\n- Demais linhas: comentário padrão\n- Formato de data preservado nas colunas B, D, T');
}

// ==================== DUPE SAAS FUNCTIONALITY ====================

// Global variables for SaaS files
let saasQmfFile = null;
let saasCustomFile = null;
let saasDrFile = null;
let saasQmfData = null;
let saasCustomData = null;
let saasDrData = null;
let saasComments = {}; // Store comments for each row

// Handle QMF Cash (SaaS) file upload
function handleSaaSQmfFileUpload(event) {
    const file = event.target.files[0];
    
    if (!file) {
        return;
    }

    // Verify it's an Excel file
    const validExtensions = ['.xlsx', '.xls'];
    const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    
    if (!validExtensions.includes(fileExtension)) {
        alert('Por favor, selecione um arquivo Excel válido (.xlsx ou .xls)');
        event.target.value = '';
        return;
    }

    // Store the file
    saasQmfFile = file;

    // Display file info
    displaySaaSQmfFileInfo(file);

    // Process the file
    processSaaSQmfFile(file);
}

// Display QMF file info
function displaySaaSQmfFileInfo(file) {
    const uploadLabel = document.querySelector('#saasQmfFileInput + .saas-upload-label');
    const status = document.getElementById('saasQmfStatus');
    const fileName = document.getElementById('saasQmfFileName');
    const fileSize = document.getElementById('saasQmfFileSize');

    const sizeInKB = (file.size / 1024).toFixed(2);
    const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
    const formattedSize = file.size > 1024 * 1024
        ? `${sizeInMB} MB`
        : `${sizeInKB} KB`;

    fileName.textContent = file.name;
    fileSize.textContent = formattedSize;
    
    uploadLabel.style.display = 'none';
    status.style.display = 'flex';
}

// Clear QMF file
function clearSaaSQmfFile() {
    saasQmfFile = null;
    saasQmfData = null;
    const uploadLabel = document.querySelector('#saasQmfFileInput + .saas-upload-label');
    const status = document.getElementById('saasQmfStatus');
    
    document.getElementById('saasQmfFileInput').value = '';
    uploadLabel.style.display = 'flex';
    status.style.display = 'none';
    checkAllSaaSFilesLoaded();
}

// Process QMF file
function processSaaSQmfFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            saasQmfData = jsonData;
            console.log('QMF Cash file loaded:', jsonData.length, 'rows');
            
            checkAllSaaSFilesLoaded();
            
        } catch (error) {
            alert('Erro ao processar o arquivo QMF Cash: ' + error.message);
            console.error('Error processing QMF file:', error);
        }
    };
    
    reader.onerror = function(error) {
        alert('Erro ao ler o arquivo QMF Cash: ' + error);
        console.error('Error reading QMF file:', error);
    };
    
    reader.readAsArrayBuffer(file);
}

// Handle Custom Report file upload
function handleSaaSCustomFileUpload(event) {
    const file = event.target.files[0];
    
    if (!file) {
        return;
    }

    const validExtensions = ['.csv'];
    const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    
    if (!validExtensions.includes(fileExtension)) {
        alert('Por favor, selecione um arquivo CSV válido (.csv)');
        event.target.value = '';
        return;
    }

    saasCustomFile = file;
    displaySaaSCustomFileInfo(file);
    processSaaSCustomFile(file);
}

// Display Custom Report file info
function displaySaaSCustomFileInfo(file) {
    const uploadLabel = document.querySelector('#saasCustomFileInput + .saas-upload-label');
    const status = document.getElementById('saasCustomStatus');
    const fileName = document.getElementById('saasCustomFileName');
    const fileSize = document.getElementById('saasCustomFileSize');

    const sizeInKB = (file.size / 1024).toFixed(2);
    const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
    const formattedSize = file.size > 1024 * 1024
        ? `${sizeInMB} MB`
        : `${sizeInKB} KB`;

    fileName.textContent = file.name;
    fileSize.textContent = formattedSize;
    
    uploadLabel.style.display = 'none';
    status.style.display = 'flex';
}

// Clear Custom Report file
function clearSaaSCustomFile() {
    saasCustomFile = null;
    saasCustomData = null;
    const uploadLabel = document.querySelector('#saasCustomFileInput + .saas-upload-label');
    const status = document.getElementById('saasCustomStatus');
    
    document.getElementById('saasCustomFileInput').value = '';
    uploadLabel.style.display = 'flex';
    status.style.display = 'none';
    checkAllSaaSFilesLoaded();
}

// Process Custom Report file (CSV)
function processSaaSCustomFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const csvText = e.target.result;
            const jsonData = parseCSV(csvText);
            
            saasCustomData = jsonData;
            console.log('Custom Report CSV file loaded:', jsonData.length, 'rows');
            
            checkAllSaaSFilesLoaded();
            
        } catch (error) {
            alert('Erro ao processar o Custom Report CSV: ' + error.message);
            console.error('Error processing Custom Report CSV:', error);
        }
    };
    
    reader.onerror = function(error) {
        alert('Erro ao ler o Custom Report CSV: ' + error);
        console.error('Error reading Custom Report CSV:', error);
    };
    
    reader.readAsText(file);
}

// Parse CSV to array format (similar to Excel sheet_to_json)
function parseCSV(csvText) {
    const lines = csvText.split(/\r?\n/);
    const result = [];
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue; // Skip empty lines
        
        const row = [];
        let currentField = '';
        let insideQuotes = false;
        
        for (let j = 0; j < line.length; j++) {
            const char = line[j];
            const nextChar = line[j + 1];
            
            if (char === '"') {
                if (insideQuotes && nextChar === '"') {
                    // Escaped quote
                    currentField += '"';
                    j++; // Skip next quote
                } else {
                    // Toggle quote state
                    insideQuotes = !insideQuotes;
                }
            } else if (char === ',' && !insideQuotes) {
                // Field separator
                row.push(currentField);
                currentField = '';
            } else {
                currentField += char;
            }
        }
        
        // Add last field
        row.push(currentField);
        result.push(row);
    }
    
    return result;
}

// Handle DR Report file upload
function handleSaaSDrFileUpload(event) {
    const file = event.target.files[0];
    
    if (!file) {
        return;
    }

    const validExtensions = ['.xlsx', '.xls'];
    const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    
    if (!validExtensions.includes(fileExtension)) {
        alert('Por favor, selecione um arquivo Excel válido (.xlsx ou .xls)');
        event.target.value = '';
        return;
    }

    saasDrFile = file;
    displaySaaSDrFileInfo(file);
    processSaaSDrFile(file);
}

// Display DR Report file info
function displaySaaSDrFileInfo(file) {
    const uploadLabel = document.querySelector('#saasDrFileInput + .saas-upload-label');
    const status = document.getElementById('saasDrStatus');
    const fileName = document.getElementById('saasDrFileName');
    const fileSize = document.getElementById('saasDrFileSize');

    const sizeInKB = (file.size / 1024).toFixed(2);
    const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
    const formattedSize = file.size > 1024 * 1024
        ? `${sizeInMB} MB`
        : `${sizeInKB} KB`;

    fileName.textContent = file.name;
    fileSize.textContent = formattedSize;
    
    uploadLabel.style.display = 'none';
    status.style.display = 'flex';
}

// Clear DR Report file
function clearSaaSDrFile() {
    saasDrFile = null;
    saasDrData = null;
    const uploadLabel = document.querySelector('#saasDrFileInput + .saas-upload-label');
    const status = document.getElementById('saasDrStatus');
    
    document.getElementById('saasDrFileInput').value = '';
    uploadLabel.style.display = 'flex';
    status.style.display = 'none';
    checkAllSaaSFilesLoaded();
}

// Process DR Report file (Excel)
function processSaaSDrFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            saasDrData = jsonData;
            console.log('DR Report Excel file loaded:', jsonData.length, 'rows');
            
            checkAllSaaSFilesLoaded();
            
        } catch (error) {
            alert('Erro ao processar o DR Report Excel: ' + error.message);
            console.error('Error processing DR Report Excel:', error);
        }
    };
    
    reader.onerror = function(error) {
        alert('Erro ao ler o DR Report Excel: ' + error);
        console.error('Error reading DR Report Excel:', error);
    };
    
    reader.readAsArrayBuffer(file);
}

// Check if all three files are loaded and process
function checkAllSaaSFilesLoaded() {
    const resultsSection = document.getElementById('saasResults');
    resultsSection.style.display = 'none';
    
    if (saasQmfData && saasCustomData && saasDrData) {
        // All three files loaded - prompt for reference month
        promptForSaaSReferenceMonth();
    }
}

// Prompt for reference month for SaaS
function promptForSaaSReferenceMonth() {
    const processingSection = document.getElementById('saasProcessing');
    processingSection.style.display = 'none';
    
    const modal = document.createElement('div');
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.5);
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 1000;
    `;
    
    const modalContent = document.createElement('div');
    modalContent.style.cssText = `
        background: white;
        padding: 30px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        max-width: 400px;
        width: 90%;
    `;
    
    modalContent.innerHTML = `
        <h3 style="margin-top: 0; color: #667eea;">Mês de Referência - SaaS</h3>
        <p style="margin-bottom: 20px; color: #666;">Informe o mês de referência para consulta de duplicidades (Payment Date):</p>
        <input type="month" id="saasReferenceMonth" style="width: 100%; padding: 10px; border: 2px solid #667eea; border-radius: 5px; font-size: 16px; margin-bottom: 20px;">
        <div style="display: flex; gap: 10px; justify-content: flex-end;">
            <button onclick="cancelSaaSMonthSelection()" style="padding: 10px 20px; background: #ccc; border: none; border-radius: 5px; cursor: pointer; font-size: 14px;">Cancelar</button>
            <button onclick="confirmSaaSMonthSelection()" style="padding: 10px 20px; background: #667eea; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 14px;">Confirmar</button>
        </div>
    `;
    
    modal.appendChild(modalContent);
    document.body.appendChild(modal);
    
    window.currentSaaSModal = modal;
    
    setTimeout(() => {
        document.getElementById('saasReferenceMonth').focus();
    }, 100);
}

// Cancel SaaS month selection
function cancelSaaSMonthSelection() {
    if (window.currentSaaSModal) {
        document.body.removeChild(window.currentSaaSModal);
        window.currentSaaSModal = null;
    }
    clearSaaSQmfFile();
    clearSaaSCustomFile();
    clearSaaSDrFile();
}

// Confirm SaaS month selection
function confirmSaaSMonthSelection() {
    const monthInput = document.getElementById('saasReferenceMonth');
    const referenceMonth = monthInput.value;
    
    if (!referenceMonth) {
        alert('Por favor, selecione um mês de referência.');
        return;
    }
    
    if (window.currentSaaSModal) {
        document.body.removeChild(window.currentSaaSModal);
        window.currentSaaSModal = null;
    }
    
    processSaaSDuplicates(referenceMonth);
}

// Process SaaS duplicates
function processSaaSDuplicates(referenceMonth) {
    const processingSection = document.getElementById('saasProcessing');
    const resultsSection = document.getElementById('saasResults');
    
    processingSection.style.display = 'block';
    
    setTimeout(() => {
        try {
            const duplicates = findSaaSDuplicates(referenceMonth);
            displaySaaSResults(duplicates, referenceMonth);
            
            processingSection.style.display = 'none';
            resultsSection.style.display = 'block';
        } catch (error) {
            processingSection.style.display = 'none';
            alert('Erro ao processar duplicatas SaaS: ' + error.message);
            console.error('Error processing SaaS duplicates:', error);
        }
    }, 500);
}

// Find SaaS duplicates based on filters
function findSaaSDuplicates(referenceMonth) {
    if (!saasQmfData || saasQmfData.length < 2) {
        throw new Error('Dados do Report QMF Cash não encontrados ou inválidos');
    }
    
    console.log('🔵 Starting findSaaSDuplicates');
    console.log('QMF Data rows:', saasQmfData.length);
    console.log('Custom Report rows:', saasCustomData ? saasCustomData.length : 'NOT LOADED');
    console.log('DR Report rows:', saasDrData ? saasDrData.length : 'NOT LOADED');
    
    const [year, month] = referenceMonth.split('-').map(Number);
    const results = [];
    
    // Get header row for column names
    const headerRow = saasQmfData[0];
    const columnNames = {
        A: headerRow[0] || 'Col A',
        B: headerRow[1] || 'Col B',
        C: headerRow[2] || 'Col C',
        D: headerRow[3] || 'Col D',
        E: headerRow[4] || 'Col E',
        F: headerRow[5] || 'Col F',
        G: headerRow[6] || 'Col G',
        L: headerRow[11] || 'Col L',
        M: headerRow[12] || 'Col M'
    };
    
    console.log('SaaS - Reference Year:', year, 'Month:', month);
    console.log('SaaS - Total rows:', saasQmfData.length);
    console.log('SaaS - Column Names:', columnNames);
    
    // Skip header row (index 0), start from row 1
    for (let i = 1; i < saasQmfData.length; i++) {
        const row = saasQmfData[i];
        
        // Column indices (0-based):
        // A=0, B=1, C=2, D=3, E=4, F=5, G=6, L=11, M=12, R=17, S=18
        const columnA = row[0] || '';  // Will display
        const columnB = row[1] || '';  // Will display
        const columnC = row[2] || '';  // Will display
        const columnD = row[3] || '';  // Payment Date - filter by reference month
        const columnE = row[4] || '';  // Will display
        const columnF = row[5] || '';  // Supplier - for matching
        const columnG = row[6] || '';  // Value - for matching
        const columnL = row[11] || ''; // Will display
        const columnM = row[12] || ''; // Will display
        const columnR = String(row[17] || '').toLowerCase().trim(); // Possible Duplicate filter
        const columnS = String(row[18] || '').toLowerCase().trim(); // Not Duplicate exclusion
        
        // Filter 1: Check if payment date (Column D) matches reference month
        if (!matchesReferenceMonth(columnD, year, month)) {
            continue;
        }
        
        // Filter 2: Column R must contain "possible duplicate"
        if (!columnR.includes('possible duplicate')) {
            continue;
        }
        
        // Filter 3: Exclude rows with "not duplicate" in Column S
        if (columnS.includes('not duplicate')) {
            continue;
        }
        
        // This row is a potential duplicate
        const rowNumber = i + 1;
        
        // Find previous months' rows with same value (G) and supplier (F)
        const previousMatches = findPreviousMonthMatches(columnF, columnG, columnD, year, month);
        
        // Determine License Type based on Payment Reference (Column A)
        const licenseType = determineLicenseType(columnA, previousMatches);
        
        results.push({
            rowNumber: rowNumber,
            columnA: columnA,
            columnB: columnB,
            columnC: columnC,
            columnD: columnD,
            columnE: columnE,
            columnF: columnF,
            columnG: columnG,
            columnL: columnL,
            columnM: columnM,
            previousMatches: previousMatches,
            columnNames: columnNames,
            licenseType: licenseType
        });
    }
    
    console.log('SaaS duplicates found:', results.length);
    return results;
}

// Determine License Type by checking Payment Reference in Custom Report and DR Report
function determineLicenseType(paymentReference, previousMatches) {
    // Collect all payment references (current + previous matches from Column A)
    const allPaymentRefs = [paymentReference];
    if (previousMatches && previousMatches.length > 0) {
        previousMatches.forEach(match => {
            if (match.columnA) {
                allPaymentRefs.push(match.columnA);
            }
        });
    }
    
    console.log('=== Checking License Type ===');
    console.log('Payment References to check:', allPaymentRefs);
    
    let foundInCustomReport = false;
    let foundInDrReport = false;
    
    // Find "Payment - Payment Reference" column index in Custom Report (header row)
    let paymentRefColumnIndex = -1;
    if (saasCustomData && saasCustomData.length > 0) {
        const headerRow = saasCustomData[0];
        
        for (let i = 0; i < headerRow.length; i++) {
            const colName = String(headerRow[i]).trim();
            
            if (colName === 'Payment - Payment Reference') {
                paymentRefColumnIndex = i;
                console.log('✓✓✓ Found "Payment - Payment Reference" column at index:', i);
                break;
            }
        }
        
        if (paymentRefColumnIndex === -1) {
            console.log('⚠️⚠️⚠️ "Payment - Payment Reference" column NOT found');
        }
    }
    
    // Check each payment reference in Custom Report BP column
    for (const payRef of allPaymentRefs) {
        const payRefStr = String(payRef).trim();
        if (!payRefStr) {
            console.log('⊘ Skipping empty payment reference');
            continue;
        }
        
        console.log(`\n→ Checking payment reference: "${payRefStr}" (length: ${payRefStr.length})`);
        
        // Search in Custom Report - specifically in "Payment - Payment Reference" column
        if (saasCustomData && paymentRefColumnIndex >= 0 && !foundInCustomReport) {
            console.log(`  Searching in ${saasCustomData.length - 1} rows (skipping header)...`);
            
            // Special check for BR172669
            if (payRefStr === 'BR172669') {
                console.log('  🔍 SPECIAL CHECK for BR172669');
                const row4692 = saasCustomData[4691]; // 0-based, so 4692-1
                if (row4692) {
                    const paymentRefValue = String(row4692[paymentRefColumnIndex] || '').trim();
                    console.log(`  Row 4692 Payment Reference value: "${paymentRefValue}" (length: ${paymentRefValue.length})`);
                    console.log(`  Match: ${paymentRefValue === payRefStr}`);
                }
            }
            
            for (let i = 1; i < saasCustomData.length; i++) { // Skip header row
                const row = saasCustomData[i];
                const paymentRefValue = String(row[paymentRefColumnIndex] || '').trim();
                
                // Extra logging for BR172669
                if (payRefStr === 'BR172669' && paymentRefValue.includes('BR172669')) {
                    console.log(`  Found similar value at row ${i + 1}: "${paymentRefValue}"`);
                }
                
                if (paymentRefValue === payRefStr) {
                    foundInCustomReport = true;
                    console.log('✓✓✓ MATCH FOUND in Custom Report!');
                    console.log('  - Row:', i + 1);
                    console.log('  - Payment Reference Value:', paymentRefValue);
                    console.log('  - Payment Ref from QMF:', payRefStr);
                    break;
                }
            }
            
            if (!foundInCustomReport) {
                console.log('  ✗ Not found in Custom Report "Payment - Payment Reference" column');
            }
        }
        
        // If found, no need to continue checking other references
        if (foundInCustomReport) break;
    }
    
    console.log('=== Result: Custom Report (Renewal):', foundInCustomReport, '===\n');
    
    // Check DR Report for New License - Column AL (index 37)
    console.log('=== Checking DR Report (New License) - Column AL ===');
    
    // Check each payment reference in DR Report Column AL
    for (const payRef of allPaymentRefs) {
        const payRefStr = String(payRef).trim();
        if (!payRefStr) continue;
        
        console.log(`\n→ Checking payment reference in DR Report Column AL: "${payRefStr}"`);
        
        // Search in DR Report - specifically in Column AL (index 37)
        if (saasDrData && saasDrData.length > 1 && !foundInDrReport) {
            console.log(`  Searching in ${saasDrData.length - 1} rows (skipping header)...`);
            
            for (let i = 1; i < saasDrData.length; i++) { // Skip header row
                const row = saasDrData[i];
                const columnAL = String(row[37] || '').trim(); // Column AL = index 37
                
                // Check if Column AL contains the payment reference
                // Can be exact match or part of the cell (for sub-lines with +)
                if (columnAL && columnAL.includes(payRefStr)) {
                    foundInDrReport = true;
                    console.log('✓✓✓ MATCH FOUND in DR Report Column AL!');
                    console.log('  - Row:', i + 1);
                    console.log('  - Column AL Value:', columnAL);
                    console.log('  - Payment Ref from QMF:', payRefStr);
                    break;
                }
            }
            
            if (!foundInDrReport) {
                console.log('  ✗ Not found in DR Report Column AL');
            }
        }
        
        // If found, no need to continue checking other references
        if (foundInDrReport) break;
    }
    
    console.log('=== Result: DR Report (New License):', foundInDrReport, '===\n');
    
    // Determine classification based on findings
    if (foundInCustomReport && foundInDrReport) {
        return 'Referência em conflito de licenças';
    } else if (foundInCustomReport) {
        return 'Renewal';
    } else if (foundInDrReport) {
        return 'New License';
    } else {
        return 'Referência não encontrada';
    }
}

// Find rows from previous months with same supplier and value
function findPreviousMonthMatches(supplier, value, currentDate, refYear, refMonth) {
    const matches = [];
    
    // Skip header row
    for (let i = 1; i < saasQmfData.length; i++) {
        const row = saasQmfData[i];
        
        const rowSupplier = row[5] || '';  // Column F
        const rowValue = row[6] || '';     // Column G
        const rowDate = row[3] || '';      // Column D
        
        // Check if supplier and value match
        if (String(rowSupplier).trim() !== String(supplier).trim()) {
            continue;
        }
        
        if (String(rowValue).trim() !== String(value).trim()) {
            continue;
        }
        
        // Check if date is from a previous month
        if (!isFromPreviousMonth(rowDate, currentDate, refYear, refMonth)) {
            continue;
        }
        
        // This is a match from a previous month
        matches.push({
            rowNumber: i + 1,
            columnA: row[0] || '',
            columnB: row[1] || '',
            columnC: row[2] || '',
            columnD: rowDate,
            columnE: row[4] || '',
            columnF: rowSupplier,
            columnG: rowValue,
            columnL: row[11] || '',
            columnM: row[12] || ''
        });
    }
    
    return matches;
}

// Check if a date is from a previous month (before the reference month)
function isFromPreviousMonth(dateValue, currentDate, refYear, refMonth) {
    if (!dateValue) return false;
    
    let date;
    
    // Check if it's an Excel serial number
    if (typeof dateValue === 'number') {
        date = XLSX.SSF.parse_date_code(dateValue);
        if (date) {
            // Compare year and month
            if (date.y < refYear) return true;
            if (date.y === refYear && date.m < refMonth) return true;
            return false;
        }
    }
    
    // Try to parse as string date
    if (typeof dateValue === 'string') {
        const parsedDate = new Date(dateValue);
        if (!isNaN(parsedDate.getTime())) {
            const year = parsedDate.getFullYear();
            const month = parsedDate.getMonth() + 1;
            
            if (year < refYear) return true;
            if (year === refYear && month < refMonth) return true;
            return false;
        }
    }
    
    return false;
}

// Get style for License Type cell
function getLicenseTypeStyle(licenseType) {
    switch(licenseType) {
        case 'Renewal':
            return 'background: #d1fae5; color: #065f46; font-weight: bold;';
        case 'New License':
            return 'background: #dbeafe; color: #1e40af; font-weight: bold;';
        case 'Referência em conflito de licenças':
            return 'background: #fee2e2; color: #991b1b; font-weight: bold;';
        case 'Referência não encontrada':
            return 'background: #fef3c7; color: #92400e; font-weight: bold;';
        default:
            return '';
    }
}

// Display SaaS results
function displaySaaSResults(duplicates, referenceMonth) {
    const resultsContent = document.getElementById('saasResultsContent');
    
    // Store duplicates globally for export function
    window.lastSaaSDuplicates = duplicates;
    window.lastSaaSReferenceMonth = referenceMonth;
    
    if (duplicates.length === 0) {
        resultsContent.innerHTML = `
            <div style="text-align: center; padding: 40px; color: #28a745;">
                <h3>✓ Nenhuma duplicata encontrada!</h3>
                <p>Não foram encontradas possíveis duplicatas SaaS com os critérios especificados para o mês de referência ${referenceMonth}.</p>
            </div>
        `;
        return;
    }
    
    // Get column names from first duplicate (they all have the same)
    const colNames = duplicates[0].columnNames;
    
    let html = `
        <div style="margin-bottom: 20px;">
            <h3 style="color: #667eea; margin-bottom: 10px;">Possíveis Duplicatas SaaS Encontradas: ${duplicates.length}</h3>
            <p style="color: #666; margin-bottom: 20px;">Mês de Referência: <strong>${referenceMonth}</strong></p>
            <p style="color: #666; font-size: 0.9rem;">💡 Clique no botão <strong>+</strong> para ver linhas de meses anteriores com mesmo valor e fornecedor</p>
        </div>
        <div style="overflow-x: auto; max-width: 100%;">
            <table style="width: 100%; border-collapse: collapse; background: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1); font-size: 0.75rem; table-layout: auto;">
                <thead>
                    <tr style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;">
                        <th style="padding: 6px 4px; text-align: center; border: 1px solid #ddd; width: 35px; font-size: 0.75rem; line-height: 1.2;">+</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; width: 50px; font-size: 0.75rem; line-height: 1.2;">Linha</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; max-width: 80px; font-size: 0.75rem; line-height: 1.2; white-space: normal; word-wrap: break-word;">${colNames.A}</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; max-width: 100px; font-size: 0.75rem; line-height: 1.2; white-space: normal; word-wrap: break-word;">${colNames.B}</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; max-width: 80px; font-size: 0.75rem; line-height: 1.2; white-space: normal; word-wrap: break-word;">${colNames.C}</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; width: 85px; font-size: 0.75rem; line-height: 1.2; white-space: normal; word-wrap: break-word;">${colNames.D}</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; max-width: 80px; font-size: 0.75rem; line-height: 1.2; white-space: normal; word-wrap: break-word;">${colNames.E}</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; max-width: 120px; font-size: 0.75rem; line-height: 1.2; white-space: normal; word-wrap: break-word;">${colNames.F}</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; width: 80px; font-size: 0.75rem; line-height: 1.2; white-space: normal; word-wrap: break-word;">${colNames.G}</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; max-width: 100px; font-size: 0.75rem; line-height: 1.2; white-space: normal; word-wrap: break-word;">${colNames.L}</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; max-width: 100px; font-size: 0.75rem; line-height: 1.2; white-space: normal; word-wrap: break-word;">${colNames.M}</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; width: 120px; font-size: 0.75rem; line-height: 1.2; white-space: normal; word-wrap: break-word;">License Type</th>
                        <th style="padding: 6px 4px; text-align: left; border: 1px solid #ddd; width: 180px; font-size: 0.75rem; line-height: 1.2;">Comentário</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    duplicates.forEach((dup, index) => {
        const rowStyle = index % 2 === 0 ? 'background: #f8f9fa;' : 'background: white;';
        const uniqueId = `saas_dup_${index}`;
        
        // Format date if it's a number (Excel serial date)
        let formattedDate = dup.columnD;
        if (typeof dup.columnD === 'number') {
            const date = XLSX.SSF.parse_date_code(dup.columnD);
            if (date) {
                formattedDate = `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`;
            }
        }
        
        const hasMatches = dup.previousMatches && dup.previousMatches.length > 0;
        const buttonStyle = hasMatches
            ? 'background: #667eea; color: white; cursor: pointer;'
            : 'background: #e2e8f0; color: #a0aec0; cursor: not-allowed;';
        
        html += `
            <tr style="${rowStyle}">
                <td style="padding: 4px 2px; border: 1px solid #ddd; text-align: center;">
                    <button onclick="toggleSaaSPreviousMatches('${uniqueId}')"
                            style="width: 24px; height: 24px; border: none; border-radius: 4px; font-size: 14px; font-weight: bold; ${buttonStyle}"
                            ${!hasMatches ? 'disabled' : ''}>
                        +
                    </button>
                    ${hasMatches ? `<div style="font-size: 0.65rem; color: #667eea; margin-top: 1px;">(${dup.previousMatches.length})</div>` : ''}
                </td>
                <td style="padding: 6px 4px; border: 1px solid #ddd; font-weight: bold; color: #667eea; font-size: 0.75rem;">${dup.rowNumber}</td>
                <td style="padding: 6px 4px; border: 1px solid #ddd; font-size: 0.75rem; max-width: 80px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${dup.columnA}">${dup.columnA}</td>
                <td style="padding: 6px 4px; border: 1px solid #ddd; font-size: 0.75rem; max-width: 100px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${dup.columnB}">${dup.columnB}</td>
                <td style="padding: 6px 4px; border: 1px solid #ddd; font-size: 0.75rem; max-width: 80px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${dup.columnC}">${dup.columnC}</td>
                <td style="padding: 6px 4px; border: 1px solid #ddd; font-size: 0.75rem;">${formattedDate}</td>
                <td style="padding: 6px 4px; border: 1px solid #ddd; font-size: 0.75rem; max-width: 80px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${dup.columnE}">${dup.columnE}</td>
                <td style="padding: 6px 4px; border: 1px solid #ddd; font-size: 0.75rem; max-width: 120px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${dup.columnF}">${dup.columnF}</td>
                <td style="padding: 6px 4px; border: 1px solid #ddd; font-weight: bold; font-size: 0.75rem;">${dup.columnG}</td>
                <td style="padding: 6px 4px; border: 1px solid #ddd; font-size: 0.75rem; max-width: 100px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${dup.columnL}">${dup.columnL}</td>
                <td style="padding: 6px 4px; border: 1px solid #ddd; font-size: 0.75rem; max-width: 100px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${dup.columnM}">${dup.columnM}</td>
                <td style="padding: 6px 4px; border: 1px solid #ddd; font-size: 0.7rem; ${getLicenseTypeStyle(dup.licenseType)}">${dup.licenseType}</td>
                <td style="padding: 4px; border: 1px solid #ddd;">
                    <input type="text"
                           id="comment_${index}"
                           placeholder="Comentário..."
                           value="${saasComments[index] || ''}"
                           style="width: 100%; padding: 4px; border: 1px solid #ddd; border-radius: 3px; font-size: 0.7rem; box-sizing: border-box;"
                           onchange="saveSaaSComment(${index}, this.value)">
                </td>
            </tr>
        `;
        
        // Add expandable row for previous matches
        if (hasMatches) {
            html += `
                <tr id="${uniqueId}" style="display: none;">
                    <td colspan="13" style="padding: 0; border: 1px solid #ddd;">
                        <div style="background: #fffbeb; padding: 15px; border-left: 4px solid #f59e0b;">
                            <h4 style="margin: 0 0 10px 0; color: #92400e; font-size: 0.95rem;">📅 Linhas de Meses Anteriores (${dup.previousMatches.length})</h4>
                            <div style="overflow-x: auto;">
                                <table style="width: 100%; border-collapse: collapse; font-size: 0.8rem;">
                                    <thead>
                                        <tr style="background: #fef3c7;">
                                            <th style="padding: 6px; text-align: left; border: 1px solid #fde68a; font-size: 0.8rem;">Linha</th>
                                            <th style="padding: 6px; text-align: left; border: 1px solid #fde68a; font-size: 0.8rem;">${dup.columnNames.A}</th>
                                            <th style="padding: 6px; text-align: left; border: 1px solid #fde68a; font-size: 0.8rem;">${dup.columnNames.B}</th>
                                            <th style="padding: 6px; text-align: left; border: 1px solid #fde68a; font-size: 0.8rem;">${dup.columnNames.C}</th>
                                            <th style="padding: 6px; text-align: left; border: 1px solid #fde68a; font-size: 0.8rem;">${dup.columnNames.D}</th>
                                            <th style="padding: 6px; text-align: left; border: 1px solid #fde68a; font-size: 0.8rem;">${dup.columnNames.E}</th>
                                            <th style="padding: 6px; text-align: left; border: 1px solid #fde68a; font-size: 0.8rem;">${dup.columnNames.F}</th>
                                            <th style="padding: 6px; text-align: left; border: 1px solid #fde68a; font-size: 0.8rem;">${dup.columnNames.G}</th>
                                            <th style="padding: 6px; text-align: left; border: 1px solid #fde68a; font-size: 0.8rem;">${dup.columnNames.L}</th>
                                            <th style="padding: 6px; text-align: left; border: 1px solid #fde68a; font-size: 0.8rem;">${dup.columnNames.M}</th>
                                        </tr>
                                    </thead>
                                    <tbody>
            `;
            
            dup.previousMatches.forEach((match, matchIndex) => {
                let matchFormattedDate = match.columnD;
                if (typeof match.columnD === 'number') {
                    const date = XLSX.SSF.parse_date_code(match.columnD);
                    if (date) {
                        matchFormattedDate = `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`;
                    }
                }
                
                const matchRowStyle = matchIndex % 2 === 0 ? 'background: white;' : 'background: #fefce8;';
                
                html += `
                    <tr style="${matchRowStyle}">
                        <td style="padding: 6px; border: 1px solid #fde68a; color: #92400e; font-weight: bold;">${match.rowNumber}</td>
                        <td style="padding: 6px; border: 1px solid #fde68a;">${match.columnA}</td>
                        <td style="padding: 6px; border: 1px solid #fde68a;">${match.columnB}</td>
                        <td style="padding: 6px; border: 1px solid #fde68a;">${match.columnC}</td>
                        <td style="padding: 6px; border: 1px solid #fde68a;">${matchFormattedDate}</td>
                        <td style="padding: 6px; border: 1px solid #fde68a;">${match.columnE}</td>
                        <td style="padding: 6px; border: 1px solid #fde68a;">${match.columnF}</td>
                        <td style="padding: 6px; border: 1px solid #fde68a; font-weight: bold;">${match.columnG}</td>
                        <td style="padding: 6px; border: 1px solid #fde68a;">${match.columnL}</td>
                        <td style="padding: 6px; border: 1px solid #fde68a;">${match.columnM}</td>
                    </tr>
                `;
            });
            
            html += `
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </td>
                </tr>
            `;
        }
    });
    
    html += `
                </tbody>
            </table>
        </div>
        <div style="margin-top: 20px; display: flex; gap: 15px; align-items: center;">
            <button onclick="exportSaaSWithComments()"
                    style="padding: 12px 24px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 8px; font-size: 1rem; font-weight: 600; cursor: pointer; box-shadow: 0 4px 6px rgba(0,0,0,0.1); transition: all 0.3s ease;"
                    onmouseover="this.style.transform='translateY(-2px)'; this.style.boxShadow='0 6px 12px rgba(0,0,0,0.15)'"
                    onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 4px 6px rgba(0,0,0,0.1)'">
                📥 Exportar Report QMF com Comentários
            </button>
        </div>
        <div style="margin-top: 15px; padding: 15px; background: #fef3c7; border-left: 4px solid #f59e0b; border-radius: 5px;">
            <strong>ℹ️ Informação:</strong> O botão acima exporta o Report QMF original com comentários adicionados na coluna T.
        </div>
    `;
    
    resultsContent.innerHTML = html;
}

// Toggle display of previous matches
function toggleSaaSPreviousMatches(rowId) {
    const row = document.getElementById(rowId);
    if (row) {
        row.style.display = row.style.display === 'none' ? 'table-row' : 'none';
    }
}

// Save comment for a specific row
function saveSaaSComment(index, comment) {
    saasComments[index] = comment;
    console.log(`Comment saved for row ${index}:`, comment);
}

// Export SaaS Report QMF with comments
function exportSaaSWithComments() {
    if (!saasQmfData || saasQmfData.length === 0) {
        alert('Nenhum dado do Report QMF disponível para exportar.');
        return;
    }
    
    // Get the reference month and year from the last processing
    const referenceMonth = window.lastSaaSReferenceMonth;
    if (!referenceMonth) {
        alert('Mês de referência não encontrado. Por favor, processe as duplicatas primeiro.');
        return;
    }
    
    const [year, month] = referenceMonth.split('-').map(Number);
    
    // Create a copy of the QMF data
    const exportData = saasQmfData.map(row => [...row]);
    
    // Add "Comments" header in column T (index 19)
    if (exportData.length > 0) {
        exportData[0][19] = 'Comments';
    }
    
    // Create a map of duplicate row numbers for quick lookup
    const duplicateRows = new Set();
    const duplicateRowComments = {};
    
    // Get all duplicates from the last processing
    if (window.lastSaaSDuplicates) {
        window.lastSaaSDuplicates.forEach((dup, index) => {
            duplicateRows.add(dup.rowNumber);
            // Use user comment if available, otherwise use default
            duplicateRowComments[dup.rowNumber] = saasComments[index] || 'Possible duplicate';
        });
    }
    
    // Process each row (skip header)
    for (let i = 1; i < exportData.length; i++) {
        const row = exportData[i];
        const rowNumber = i + 1;
        
        // Check if this row is a duplicate
        if (duplicateRows.has(rowNumber)) {
            // Add user comment or default duplicate comment
            row[19] = duplicateRowComments[rowNumber];
        } else {
            // Check if row is from the reference month
            const rowDate = row[3]; // Column D - Payment Date
            
            if (rowDate) {
                let isReferenceMonth = false;
                
                // Check if it's an Excel serial number
                if (typeof rowDate === 'number') {
                    const date = XLSX.SSF.parse_date_code(rowDate);
                    if (date && date.y === year && date.m === month) {
                        isReferenceMonth = true;
                    }
                } else {
                    // Try to parse as string date
                    const dateStr = String(rowDate);
                    if (dateStr.includes(`${year}-${String(month).padStart(2, '0')}`)) {
                        isReferenceMonth = true;
                    }
                }
                
                if (isReferenceMonth) {
                    // Row is from reference month but not a duplicate
                    row[19] = 'Not a dupe according report criteria';
                } else {
                    // Row is from a different month
                    row[19] = 'Line not analyse in the report period';
                }
            } else {
                // No date available
                row[19] = 'Line not analyse in the report period';
            }
        }
    }
    
    // Create workbook and worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(exportData);
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Report QMF with Comments');
    
    // Generate filename with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    const filename = `Report_QMF_with_Comments_${timestamp}.xlsx`;
    
    // Export file
    XLSX.writeFile(wb, filename);
    
    console.log(`Exported ${exportData.length - 1} rows with comments to ${filename}`);
    alert(`Arquivo exportado com sucesso: ${filename}`);
}

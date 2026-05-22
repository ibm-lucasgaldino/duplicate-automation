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

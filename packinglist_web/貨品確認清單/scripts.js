// 全域變數
let processedData = [];
let startTime = 0;

// DOM 元素
const fileInput = document.getElementById('excelFile');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const progressContainer = document.getElementById('progressContainer');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const previewSection = document.getElementById('previewSection');
const previewTable = document.getElementById('previewTable');
const tableHeader = document.getElementById('tableHeader');
const tableBody = document.getElementById('tableBody');
const summaryInfo = document.getElementById('summaryInfo');
const totalCount = document.getElementById('totalCount');
const processTime = document.getElementById('processTime');
const downloadBtn = document.getElementById('downloadBtn');
const outputFileName = document.getElementById('outputFileName');
const errorContainer = document.getElementById('errorContainer');
const errorMessage = document.getElementById('errorMessage');

// 事件監聽器
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
});

function initializeEventListeners() {
    // 檔案上傳事件
    fileInput.addEventListener('change', handleFileUpload);
    
    // 下載按鈕事件
    downloadBtn.addEventListener('click', downloadExcel);
    
    // 檔案名稱輸入事件
    outputFileName.addEventListener('input', function() {
        // 移除不合法的檔案名稱字符
        this.value = this.value.replace(/[<>:"/\\|?*]/g, '');
    });
}

// 處理檔案上傳
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    // 驗證檔案類型
    if (!isValidExcelFile(file)) {
        showError('請選擇有效的 Excel 檔案 (.xlsx 或 .xls)');
        return;
    }
    
    // 顯示檔案資訊
    showFileInfo(file);
    
    // 隱藏錯誤和預覽區域
    hideError();
    hidePreview();
    
    try {
        startTime = Date.now();
        await processExcelFile(file);
    } catch (error) {
        console.error('處理檔案時發生錯誤:', error);
        showError(`處理檔案時發生錯誤: ${error.message}`);
    }
}

// 驗證是否為有效的 Excel 檔案
function isValidExcelFile(file) {
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    const validExtensions = ['.xlsx', '.xls'];
    
    return validTypes.includes(file.type) || 
           validExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
}

// 顯示檔案資訊
function showFileInfo(file) {
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileInfo.style.display = 'flex';
}

// 格式化檔案大小
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 處理 Excel 檔案
async function processExcelFile(file) {
    showProgress(0, '讀取檔案中...');
    
    try {
        // 讀取檔案
        const arrayBuffer = await file.arrayBuffer();
        showProgress(25, '解析 Excel 檔案...');
        
        // 解析 Excel
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        showProgress(50, '處理資料中...');
        
        // 處理資料
        const extractedData = extractDataFromWorkbook(workbook);
        showProgress(75, '生成預覽...');
        
        // 生成預覽
        displayPreview(extractedData);
        showProgress(100, '處理完成!');
        
        // 隱藏進度條，顯示預覽
        setTimeout(() => {
            hideProgress();
            showPreview();
        }, 500);
        
    } catch (error) {
        hideProgress();
        throw error;
    }
}

// 從工作簿中提取資料
function extractDataFromWorkbook(workbook) {
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    // 尋找資料表格的開始行（包含"序號"的行）
    let headerRowIndex = -1;
    const targetColumns = ['序號', '廠商', '貨名', '箱規', '數量', '報關數量', '包裝方式'];
    
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && Array.isArray(row)) {
            // 檢查是否包含目標欄位
            const hasTargetColumns = targetColumns.some(col => 
                row.some(cell => cell && cell.toString().includes(col))
            );
            if (hasTargetColumns) {
                headerRowIndex = i;
                break;
            }
        }
    }
    
    if (headerRowIndex === -1) {
        throw new Error('找不到包含必要欄位的資料表格');
    }
    
    // 提取標題行
    const headerRow = rawData[headerRowIndex];
    
    // 找到目標欄位的索引
    const columnIndexes = findColumnIndexes(headerRow);
    
    // 提取資料行
    const dataRows = rawData.slice(headerRowIndex + 1).filter(row => 
        row && row.length > 0 && row[columnIndexes.序號]
    );
    
    // 處理資料
    const processedRows = dataRows.map((row, index) => ({
        序號: row[columnIndexes.序號] || (index + 1),
        廠商: row[columnIndexes.廠商] || '',
        貨名: row[columnIndexes.貨名] || '',
        箱規: row[columnIndexes.箱規] || '',
        '數量(箱)': extractQuantity(row[columnIndexes['數量(箱)']]) || '',
        報關數量: row[columnIndexes.報關數量] || '',
        包裝方式: row[columnIndexes.包裝方式] || '',
        物流號: '' // 新增的空白物流號欄位
    }));
    
    processedData = processedRows;
    return processedRows;
}

// 找到欄位索引
function findColumnIndexes(headerRow) {
    const indexes = {};
    const columnMapping = {
        '序號': ['序號'],
        '廠商': ['廠商'],
        '貨名': ['貨名'],
        '箱規': ['箱規'],
        '數量(箱)': ['數量(箱)', '數量', '數量(箱)報關數量'],
        '報關數量': ['報關數量', '數量(箱)報關數量'],
        '包裝方式': ['包裝方式']
    };
    
    Object.keys(columnMapping).forEach(key => {
        const possibleNames = columnMapping[key];
        for (let i = 0; i < headerRow.length; i++) {
            const cellValue = headerRow[i];
            if (cellValue) {
                const cellStr = cellValue.toString().trim();
                if (possibleNames.some(name => cellStr.includes(name))) {
                    indexes[key] = i;
                    break;
                }
            }
        }
    });
    
    return indexes;
}

// 提取數量資訊（處理合併欄位）
function extractQuantity(cellValue) {
    if (!cellValue) return '';
    
    const str = cellValue.toString();
    // 如果包含報關數量，提取第一個數字
    if (str.includes('報關數量') || str.includes('\r\n') || str.includes('\n')) {
        const match = str.match(/(\d+)/);
        return match ? match[1] : str;
    }
    
    return str;
}

// 顯示預覽
function displayPreview(data) {
    // 清空表格
    tableHeader.innerHTML = '';
    tableBody.innerHTML = '';
    
    if (data.length === 0) {
        throw new Error('沒有找到有效的資料');
    }
    
    // 建立標題行
    const headers = ['序號', '廠商', '貨名', '箱規', '數量(箱)', '報關數量', '包裝方式', '物流號'];
    const headerRow = document.createElement('tr');
    
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    
    tableHeader.appendChild(headerRow);
    
    // 建立資料行
    data.forEach(row => {
        const tr = document.createElement('tr');
        
        headers.forEach(header => {
            const td = document.createElement('td');
            td.textContent = row[header] || '';
            tr.appendChild(td);
        });
        
        tableBody.appendChild(tr);
    });
    
    // 更新摘要資訊
    totalCount.textContent = data.length;
    const endTime = Date.now();
    const processingTime = ((endTime - startTime) / 1000).toFixed(1);
    processTime.textContent = processingTime;
}

// 下載 Excel 檔案 (使用 ExcelJS)
async function downloadExcel() {
    if (processedData.length === 0) {
        showError('沒有可下載的資料');
        return;
    }
    
    try {
        // 建立新的工作簿
        const workbook = new ExcelJS.Workbook();
        
        // 設定工作簿屬性
        workbook.creator = '嘉鴻精密科技股份有限公司';
        workbook.created = new Date();
        workbook.modified = new Date();
        workbook.lastPrinted = new Date();
        
        // 建立工作表
        const worksheet = workbook.addWorksheet('貨品確認清單', {
            pageSetup: {
                paperSize: 9, // A4
                orientation: 'landscape',
                fitToPage: true,
                margins: {
                    left: 0.7, right: 0.7,
                    top: 0.75, bottom: 0.75,
                    header: 0.3, footer: 0.3
                }
            }
        });
        
        // 準備表頭
        const headers = ['序號', '廠商', '貨名', '箱規', '數量(箱)', '報關數量', '包裝方式', '物流號'];
        
        // 添加標題行
        const headerRow = worksheet.addRow(headers);
        
        // 設定標題行樣式
        headerRow.eachCell((cell, colNumber) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFBFBFBF' }
            };
            cell.font = {
                name: '微軟正黑體',
                size: 12,
                bold: true,
                color: { argb: 'FF000000' }
            };
            cell.alignment = {
                horizontal: 'center',
                vertical: 'middle'
            };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            };
        });
        
        // 設定標題行高度
        headerRow.height = 25;
        
        // 添加資料行
        processedData.forEach((rowData, index) => {
            const row = worksheet.addRow([
                rowData['序號'] || (index + 1),
                rowData['廠商'] || '',
                rowData['貨名'] || '',
                rowData['箱規'] || '',
                rowData['數量(箱)'] || '',
                rowData['報關數量'] || '',
                rowData['包裝方式'] || '',
                rowData['物流號'] || ''
            ]);
            
            // 設定資料行樣式
            row.eachCell((cell, colNumber) => {
                cell.font = {
                    name: '微軟正黑體',
                    size: 11
                };
                
                // 序號欄位特殊樣式
                if (colNumber === 1) {
                    cell.font.bold = true;
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFE8F4FD' }
                    };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } 
                // 物流號欄位特殊樣式
                else if (colNumber === 8) {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFFF9E6' }
                    };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FF000000' } },
                        left: { style: 'thin', color: { argb: 'FF000000' } },
                        bottom: { style: 'thin', color: { argb: 'FF000000' } },
                        right: { style: 'thick', color: { argb: 'FF4472C4' } }
                    };
                }
                // 其他欄位樣式
                else {
                    // 交替行背景色
                    if (index % 2 === 0) {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFF8F9FA' }
                        };
                    } else {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFFFFFF' }
                        };
                    }
                    cell.alignment = { 
                        horizontal: 'left', 
                        vertical: 'middle',
                        wrapText: true
                    };
                }
                
                // 為所有儲存格添加邊框 (除非已經設定過)
                if (!cell.border) {
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FF000000' } },
                        left: { style: 'thin', color: { argb: 'FF000000' } },
                        bottom: { style: 'thin', color: { argb: 'FF000000' } },
                        right: { style: 'thin', color: { argb: 'FF000000' } }
                    };
                }
            });
            
            // 設定資料行高度
            row.height = 20;
        });
        
        // 設定欄位寬度
        const columnWidths = [12, 20, 40, 30, 15, 15, 15, 20]; // 序號, 廠商, 貨名, 箱規, 數量(箱), 報關數量, 包裝方式, 物流號
        columnWidths.forEach((width, index) => {
            worksheet.getColumn(index + 1).width = width;
        });
        
        // 凍結標題行
        worksheet.views = [
            { state: 'frozen', ySplit: 1 }
        ];
        
        // 生成檔案名稱
        const filename = (outputFileName.value.trim() || '貨品確認清單') + '.xlsx';
        
        // 生成並下載檔案
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        // 建立下載連結
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        // 顯示成功訊息
        showTemporaryMessage('檔案下載成功！', 'success');
        
    } catch (error) {
        console.error('下載檔案時發生錯誤:', error);
        showError(`下載檔案時發生錯誤: ${error.message}`);
    }
}

// 顯示進度
function showProgress(percentage, message) {
    progressContainer.style.display = 'block';
    progressFill.style.width = percentage + '%';
    progressText.textContent = message;
}

// 隱藏進度
function hideProgress() {
    progressContainer.style.display = 'none';
}

// 顯示預覽區域
function showPreview() {
    previewSection.style.display = 'block';
}

// 隱藏預覽區域
function hidePreview() {
    previewSection.style.display = 'none';
}

// 顯示錯誤訊息
function showError(message) {
    errorMessage.textContent = message;
    errorContainer.style.display = 'block';
    hideProgress();
}

// 隱藏錯誤訊息
function hideError() {
    errorContainer.style.display = 'none';
}

// 顯示臨時訊息
function showTemporaryMessage(message, type = 'info') {
    const messageDiv = document.createElement('div');
    messageDiv.className = `temp-message temp-message-${type}`;
    messageDiv.innerHTML = `
        <div class="temp-message-content">
            <span>${message}</span>
            <button class="temp-message-close" onclick="this.parentElement.parentElement.remove()">×</button>
        </div>
    `;
    
    // 添加樣式
    messageDiv.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: ${type === 'success' ? '#4CAF50' : '#2196F3'};
        color: white;
        padding: 15px 20px;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 1000;
        animation: slideInRight 0.3s ease;
    `;
    
    document.body.appendChild(messageDiv);
    
    // 3秒後自動移除
    setTimeout(() => {
        if (messageDiv.parentElement) {
            messageDiv.remove();
        }
    }, 3000);
}

// 添加 CSS 動畫
const style = document.createElement('style');
style.textContent = `
    @keyframes slideInRight {
        from {
            transform: translateX(100%);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
    
    .temp-message-content {
        display: flex;
        align-items: center;
        gap: 15px;
    }
    
    .temp-message-close {
        background: none;
        border: none;
        color: white;
        font-size: 18px;
        font-weight: bold;
        cursor: pointer;
        padding: 0;
        width: 20px;
        height: 20px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 50%;
    }
    
    .temp-message-close:hover {
        background: rgba(255,255,255,0.2);
    }
`;
document.head.appendChild(style);
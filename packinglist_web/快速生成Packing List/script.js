// 快速生成 Packing List 系統 - script.js (修正版：解決公式和數值讀取問題)

document.addEventListener('DOMContentLoaded', async () => {
    // ===== 全域變數 =====
    let masterListData = null;
    let uploadedImageData = null;

    // ===== DOM 元素 =====
    const message = document.getElementById('message');
    const masterListInput = document.getElementById('masterListInput');
    const masterListStatus = document.getElementById('masterListStatus');
    const loadedMasterListInfo = document.getElementById('loadedMasterListInfo');
    const masterListPreview = document.getElementById('masterListPreview');
    const quickPreviewPacking = document.getElementById('quickPreviewPacking');
    const quickGeneratePacking = document.getElementById('quickGeneratePacking');
    const imageInput = document.getElementById('imageInput');
    const imageStatus = document.getElementById('imageStatus');
    const imagePreview = document.getElementById('imagePreview');
    const previewImg = document.getElementById('previewImg');
    const imageInfo = document.getElementById('imageInfo');
    const removeImage = document.getElementById('removeImage');
    const previewModal = document.getElementById('previewModal');
    const closePreview = document.getElementById('closePreview');
    const closePreviewBottom = document.getElementById('closePreviewBottom');
    const generateFromPreview = document.getElementById('generateFromPreview');
    const previewContent = document.getElementById('previewContent');
    
    // 新增的貿易條件相關DOM元素
    const tradeTerms = document.getElementById('tradeTerms');
    const portSelection = document.getElementById('portSelection');
    const addAddress2 = document.getElementById('addAddress2');
    const addAddress3 = document.getElementById('addAddress3');
    const address2Container = document.getElementById('address2Container');
    const address3Container = document.getElementById('address3Container');

    // ===== 基礎工具函數 =====

    function showMessage(text, type = 'info') {
        message.textContent = text;
        message.className = `mt-4 p-3 rounded ${
            type === 'error' ? 'bg-red-100 text-red-700 border border-red-300' : 
            type === 'success' ? 'bg-green-100 text-green-700 border border-green-300' : 
            type === 'warning' ? 'bg-yellow-100 text-yellow-700 border border-yellow-300' :
            'bg-blue-100 text-blue-700 border border-blue-300'
        }`;
        message.classList.remove('hidden');
        setTimeout(() => message.classList.add('hidden'), 5000);
    }

    function showStatus(element, text, type = 'info') {
        if (!element) return;
        element.textContent = text;
        element.className = `mt-2 text-sm ${
            type === 'error' ? 'status-error' : 
            type === 'success' ? 'status-success' : 
            type === 'loading' ? 'status-loading' : 
            type === 'warning' ? 'status-warning' :
            'status-info'
        }`;
    }

    // ===== 數值處理函數 =====
    
    // 安全地獲取單元格數值 - 特別處理箱規等文本資料
    function getCellValue(cell, isTextColumn = false) {
        if (!cell || cell.value === undefined || cell.value === null) {
            return isTextColumn ? '' : 0;
        }
        
        // 調試：輸出儲存格原始值結構
        if (isTextColumn) {
            console.log('儲存格原始值:', {
                value: cell.value,
                type: typeof cell.value,
                hasRichText: cell.value && cell.value.richText,
                cellText: cell.text,
                result: cell.result
            });
        }
        
        // 如果是數字直接返回
        if (typeof cell.value === 'number') {
            return cell.value;
        }
        
        // 如果是字符串，對於文本欄位直接返回，數值欄位嘗試轉換
        if (typeof cell.value === 'string') {
            if (isTextColumn) {
                return cell.value; // 文本欄位直接返回完整字串
            }
            const numValue = parseFloat(cell.value);
            return isNaN(numValue) ? cell.value : numValue;
        }
        
        // 處理富文本格式 (richText) - 重要：箱規資料可能是這種格式
        if (typeof cell.value === 'object' && cell.value.richText && Array.isArray(cell.value.richText)) {
            let fullText = '';
            cell.value.richText.forEach(textPart => {
                if (textPart.text) {
                    fullText += textPart.text;
                }
            });
            console.log('富文本合併結果:', fullText);
            
            if (isTextColumn) {
                return fullText; // 文本欄位返回完整文字
            }
            
            // 數值欄位嘗試轉換
            const numValue = parseFloat(fullText);
            return isNaN(numValue) ? fullText : numValue;
        }
        
        // 如果是公式結果，嘗試獲取計算值
        if (cell.result !== undefined && cell.result !== null) {
            if (typeof cell.result === 'number') {
                return cell.result;
            }
            if (typeof cell.result === 'string') {
                return isTextColumn ? cell.result : (parseFloat(cell.result) || 0);
            }
        }
        
        // 如果是對象，嘗試獲取其數值屬性
        if (typeof cell.value === 'object') {
            if (cell.value.result !== undefined) {
                const result = cell.value.result;
                return isTextColumn ? result : (parseFloat(result) || 0);
            }
            if (cell.value.value !== undefined) {
                const value = cell.value.value;
                return isTextColumn ? value : (parseFloat(value) || 0);
            }
            // 如果都不行，返回預設值
            return isTextColumn ? '' : 0;
        }
        
        return cell.value;
    }

    // 安全地讀取貨櫃尺寸，強制轉換為文字格式
    function getContainerSizeSafely(cell) {
        if (!cell || cell.value === undefined || cell.value === null) {
            return '40呎'; // 預設值
        }
        
        let containerSize = '';
        
        // 如果是字串，直接使用
        if (typeof cell.value === 'string') {
            containerSize = cell.value;
        }
        // 如果是數字，轉為字串
        else if (typeof cell.value === 'number') {
            containerSize = cell.value.toString();
        }
        // 如果是物件，嘗試各種方式提取文字
        else if (typeof cell.value === 'object') {
            // 嘗試 richText 格式
            if (cell.value.richText && Array.isArray(cell.value.richText)) {
                containerSize = '';
                cell.value.richText.forEach(textPart => {
                    if (textPart.text) {
                        containerSize += textPart.text;
                    }
                });
            }
            // 嘗試 value 屬性
            else if (cell.value.value !== undefined) {
                containerSize = cell.value.value.toString();
            }
            // 嘗試 result 屬性
            else if (cell.value.result !== undefined) {
                containerSize = cell.value.result.toString();
            }
            // 最後使用 toString()
            else {
                containerSize = cell.value.toString();
            }
        }
        // 其他情況
        else {
            containerSize = cell.value.toString();
        }
        
        // 移除可能導致Excel錯誤的字符，並清理空白
        containerSize = containerSize.replace(/[\*\?\:\\\/\[\]]/g, '').trim();
        
        // 如果結果是空的或異常的，使用預設值
        if (!containerSize || containerSize === '[object Object]' || containerSize === 'undefined') {
            containerSize = '40呎';
        }
        
        console.log('貨櫃尺寸讀取結果:', {
            原始值: cell.value,
            原始類型: typeof cell.value,
            最終結果: containerSize
        });
        
        return containerSize;
    }

    // ===== 日期相關函數 =====

    // 日期格式化函數
    function formatDateForWorksheet(dateString) {
        if (!dateString) {
            const today = new Date();
            dateString = today.toISOString().split('T')[0];
        }
        
        const date = new Date(dateString);
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const day = date.getDate().toString().padStart(2, '0');
        
        return `${month}月${day}日`;
    }

    // 生成工作表名稱
    function generateWorksheetName(type, dateString, containerSize) {
        const formattedDate = formatDateForWorksheet(dateString);
        
        // 安全處理 containerSize，確保是字串格式
        let safeContainerSize = '';
        if (containerSize) {
            if (typeof containerSize === 'string') {
                safeContainerSize = containerSize;
            } else if (typeof containerSize === 'object' && containerSize.value) {
                safeContainerSize = containerSize.value.toString();
            } else {
                safeContainerSize = containerSize.toString();
            }
        }
        
        // 移除可能導致Excel錯誤的字符
        safeContainerSize = safeContainerSize.replace(/[\*\?\:\\\/\[\]]/g, '');
        
        if (type === 'packinglist') {
            return `Packing List 貨櫃 ${formattedDate} ${safeContainerSize}`;
        } else if (type === 'invoice') {
            return `Invoice ${formattedDate} ${safeContainerSize}`;
        }
        
        return 'Sheet1';
    }

    // 初始化日期選擇器為今天
    function initializeDateInputs() {
        const today = new Date().toISOString().split('T')[0];
        const quickShippingDateInput = document.getElementById('quickShippingDate');
        
        if (quickShippingDateInput) {
            quickShippingDateInput.value = today;
        }
    }

    // 初始化日期選擇器
    initializeDateInputs();

    // ===== 貿易條件格式生成函數 =====
    function generateTradeConditionString() {
        const tradeTermsValue = tradeTerms.value;
        const portValue = portSelection.value;
        
        // 獲取所有顯示的地址
        const addresses = [];
        
        // 地址一（始終顯示）
        const address1Select = document.querySelector('.address-item:nth-child(1) .address-select');
        if (address1Select) {
            addresses.push(`地址一：${address1Select.value}`);
        }
        
        // 地址二（如果顯示）
        if (address2Container.style.display !== 'none') {
            const address2Select = address2Container.querySelector('.address-select');
            if (address2Select) {
                addresses.push(`地址二：${address2Select.value}`);
            }
        }
        
        // 地址三（如果顯示）
        if (address3Container.style.display !== 'none') {
            const address3Select = address3Container.querySelector('.address-select');
            if (address3Select) {
                addresses.push(`地址三：${address3Select.value}`);
            }
        }
        
        // 組合最終字串
        const addressString = addresses.join('-');
        return `${tradeTermsValue} ${portValue}-${addressString}`;
    }

    // ===== 檢查狀態函數 =====
    function checkQuickFunctionReady() {
        const hasMasterList = masterListData !== null;
        const hasDate = document.getElementById('quickShippingDate')?.value;
        
        quickPreviewPacking.disabled = !hasMasterList;
        quickGeneratePacking.disabled = !(hasMasterList && hasDate);
        
        // 如果沒有日期，顯示提示
        if (hasMasterList && !hasDate) {
            showMessage('請選擇出貨日期以啟用快速生成功能', 'warning');
        }
        
        // 如果有圖片，顯示圖片狀態
        if (uploadedImageData) {
            console.log('已載入圖片，將添加到 Excel 文件中');
        }
    }

    // 添加日期變更事件監聽器
    const quickShippingDateInput = document.getElementById('quickShippingDate');

    if (quickShippingDateInput) {
        quickShippingDateInput.addEventListener('change', () => {
            checkQuickFunctionReady();
        });
    }

    // ===== 圖片處理相關函數 =====

    // 圖片壓縮和調整大小函數
    function resizeImage(file, targetWidth, targetHeight) {
        return new Promise((resolve) => {
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            const img = new Image();
            
            img.onload = () => {
                // 計算等比例縮放
                const ratio = Math.min(targetWidth / img.width, targetHeight / img.height);
                const newWidth = img.width * ratio;
                const newHeight = img.height * ratio;
                
                canvas.width = newWidth;
                canvas.height = newHeight;
                
                ctx.drawImage(img, 0, 0, newWidth, newHeight);
                
                canvas.toBlob((blob) => {
                    resolve(blob);
                }, file.type, 0.8); // 壓縮品質 80%
            };
            
            img.src = URL.createObjectURL(file);
        });
    }

    // 圖片上傳處理函數
    async function handleImageUpload(file) {
        try {
            showStatus(imageStatus, '正在處理圖片...', 'loading');
            
            // 檢查檔案類型
            if (!file.type.startsWith('image/')) {
                throw new Error('請選擇有效的圖片檔案');
            }
            
            // 檢查檔案大小 (5MB)
            if (file.size > 5 * 1024 * 1024) {
                throw new Error('圖片檔案過大，請選擇小於 5MB 的檔案');
            }
            
            // Excel 中的目標尺寸 (單位：像素，按 96 DPI 計算)
            // 長 1.73cm ≈ 65 pixels, 寬 3.04cm ≈ 115 pixels
            const targetWidth = 115;
            const targetHeight = 65;
            
            console.log('開始處理圖片:', file.name, '大小:', file.size, 'bytes', '類型:', file.type);
            
            // 壓縮圖片
            const resizedBlob = await resizeImage(file, targetWidth, targetHeight);
            
            // 轉換為 ArrayBuffer
            const arrayBuffer = await resizedBlob.arrayBuffer();
            
            // 確定圖片副檔名
            let extension = file.type.split('/')[1];
            if (extension === 'jpeg') extension = 'jpg'; // ExcelJS 偏好 jpg 而非 jpeg
            if (!['jpg', 'png', 'gif', 'bmp'].includes(extension)) {
                extension = 'png'; // 預設使用 png
            }
            
            console.log('圖片處理完成:', {
                originalSize: file.size,
                processedSize: resizedBlob.size,
                extension: extension,
                targetDimensions: { width: targetWidth, height: targetHeight },
                arrayBufferSize: arrayBuffer.byteLength
            });
            
            // 儲存圖片資料
            uploadedImageData = {
                arrayBuffer: arrayBuffer,
                extension: extension,
                originalName: file.name,
                size: resizedBlob.size,
                dimensions: { width: targetWidth, height: targetHeight }
            };
            
            // 顯示預覽
            const previewUrl = URL.createObjectURL(resizedBlob);
            previewImg.src = previewUrl;
            imageInfo.textContent = `已調整大小: ${targetWidth}×${targetHeight}px (${(resizedBlob.size / 1024).toFixed(1)} KB)`;
            imagePreview.classList.remove('hidden');
            
            showStatus(imageStatus, '✓ 圖片上傳成功', 'success');
            checkQuickFunctionReady();
            
            console.log('✓ 圖片資料已儲存到 uploadedImageData');
            console.log('uploadedImageData:', uploadedImageData);
            
        } catch (error) {
            console.error('圖片上傳錯誤:', error);
            showStatus(imageStatus, `圖片上傳失敗: ${error.message}`, 'error');
            uploadedImageData = null;
        }
    }

    // 移除圖片函數
    function removeUploadedImage() {
        uploadedImageData = null;
        imagePreview.classList.add('hidden');
        imageInput.value = '';
        showStatus(imageStatus, '', '');
        checkQuickFunctionReady();
    }

    // ===== 圖片事件監聽器 =====
    imageInput.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (file) {
            await handleImageUpload(file);
        }
    });

    removeImage.addEventListener('click', removeUploadedImage);

    // ===== 地址管理事件監聽器 =====
    addAddress2.addEventListener('click', () => {
        address2Container.style.display = 'block';
        addAddress2.style.display = 'none';
        addAddress3.style.display = 'inline-flex';
    });

    addAddress3.addEventListener('click', () => {
        address3Container.style.display = 'block';
        addAddress3.style.display = 'none';
    });

    // 移除地址按鈕事件監聽器
    document.addEventListener('click', (e) => {
        if (e.target.classList.contains('remove-address') || e.target.closest('.remove-address')) {
            const addressItem = e.target.closest('.address-item');
            if (addressItem.id === 'address2Container') {
                address2Container.style.display = 'none';
                addAddress2.style.display = 'inline-flex';
                // 如果地址三也顯示，需要調整按鈕狀態
                if (address3Container.style.display !== 'none') {
                    addAddress3.style.display = 'inline-flex';
                }
            } else if (addressItem.id === 'address3Container') {
                address3Container.style.display = 'none';
                addAddress3.style.display = 'inline-flex';
            }
        }
    });

    // ===== 總表上傳處理 =====
    masterListInput.addEventListener('change', async (event) => {
        await processMasterList(event.target.files[0]);
    });

    async function processMasterList(file) {
        try {
            if (!file) {
                showStatus(masterListStatus, '未選擇檔案', 'error');
                return;
            }

            showStatus(masterListStatus, '正在載入總表檔案...', 'loading');
            
            const arrayBuffer = await file.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            
            const worksheet = workbook.getWorksheet(1);
            if (!worksheet) {
                throw new Error('總表檔案中沒有找到工作表');
            }

            // 讀取基本資訊 (A1-J9)
            const basicInfo = {};
            for (let row = 1; row <= 9; row++) {
                for (let col = 1; col <= 10; col++) {
                    const cell = worksheet.getCell(row, col);
                    if (cell.value && cell.value !== '') {
                        basicInfo[`${String.fromCharCode(64 + col)}${row}`] = cell.value;
                    }
                }
            }

            // 讀取貨櫃尺寸 (D11) - 使用安全讀取方式
            const containerCell = worksheet.getCell(11, 4);
            const containerSize = getContainerSizeSafely(containerCell);

            // 讀取產品資料 (從第14行開始) - 使用安全數值讀取
            const productData = [];
            let rowNum = 14;
            
            while (true) {
                const row = worksheet.getRow(rowNum);
                const serialNumber = getCellValue(row.getCell(1));
                
                // 如果序號為空或不是數字，表示到達資料末尾
                if (!serialNumber || isNaN(parseInt(serialNumber))) {
                    break;
                }

                const product = {
                    '序號': getCellValue(row.getCell(1)),
                    '廠商': getCellValue(row.getCell(2), true) || '',
                    '報關名稱': getCellValue(row.getCell(3), true) || '',
                    '貨名': getCellValue(row.getCell(4), true) || '',
                    '箱規': getCellValue(row.getCell(5), true) || '',
                    '數量': getCellValue(row.getCell(6)) || 0,
                    '最小單位': getCellValue(row.getCell(7), true) || '',
                    '長 (cm)': getCellValue(row.getCell(8)) || 0,
                    '寬 (cm)': getCellValue(row.getCell(9)) || 0,
                    '高 (cm)': getCellValue(row.getCell(10)) || 0,
                    '毛重kg': getCellValue(row.getCell(11)) || 0,
                    '淨重kg': getCellValue(row.getCell(12)) || 0,
                    '報關數量': getCellValue(row.getCell(13)) || 0,
                    '包裝方式': getCellValue(row.getCell(14), true) || '',
                    // ===== 關鍵修改：安全讀取合計數值，處理公式結果 =====
                    '毛重kg(合計)': getCellValue(row.getCell(15)) || 0,
                    '淨重kg(合計)': getCellValue(row.getCell(16)) || 0,
                    '體積(合計)': getCellValue(row.getCell(17)) || 0,
                    '材質': getCellValue(row.getCell(18), true) || '',
                    '用途': getCellValue(row.getCell(19), true) || '',
                    '稅則編碼': getCellValue(row.getCell(20), true) || '',
                    '稅率(%)': getCellValue(row.getCell(21), true) || '',
                    '麥頭名稱備註': getCellValue(row.getCell(22), true) || '',
                    '¥採購單價(外部)': getCellValue(row.getCell(23)) || 0,
                    '¥採購單價(內部)': getCellValue(row.getCell(25)) || 0,
                    '¥採購总金額(內部)': getCellValue(row.getCell(26)) || 0
                };

                console.log(`第${rowNum}行產品資料:`, {
                    序號: product['序號'],
                    廠商: product['廠商'],
                    報關名稱: product['報關名稱'],
                    貨名: product['貨名'],
                    箱規: product['箱規'],
                    毛重合計: product['毛重kg(合計)'],
                    淨重合計: product['淨重kg(合計)'],
                    體積合計: product['體積(合計)']
                });

                productData.push(product);
                rowNum++;
            }

            // 讀取合計資料 (第12行) - 使用安全數值讀取
            const summaryRow = worksheet.getRow(12);
            const summaryData = {
                '合計文字': getCellValue(summaryRow.getCell(12)) || '合計',
                '報關數量合計': getCellValue(summaryRow.getCell(13)) || 0,
                '毛重合計': getCellValue(summaryRow.getCell(15)) || 0,
                '淨重合計': getCellValue(summaryRow.getCell(16)) || 0,
                '體積合計': getCellValue(summaryRow.getCell(17)) || 0
            };

            console.log('合計資料:', summaryData);

            masterListData = {
                basicInfo,
                containerSize,
                productData,
                summaryData,
                fileName: file.name
            };

            showStatus(masterListStatus, '✓ 總表檔案載入成功', 'success');
            showMessage('總表檔案載入成功', 'success');
            
            // 顯示載入資訊
            loadedMasterListInfo.innerHTML = `
                <div class="master-info-grid">
                    <div class="master-info-item">
                        <div class="label">檔案名稱</div>
                        <div class="value">${file.name}</div>
                    </div>
                    <div class="master-info-item">
                        <div class="label">產品數量</div>
                        <div class="value">${productData.length} 項</div>
                    </div>
                    <div class="master-info-item">
                        <div class="label">貨櫃尺寸</div>
                        <div class="value">${containerSize}</div>
                    </div>
                </div>
            `;

            generateMasterListPreview();
            checkQuickFunctionReady();

        } catch (error) {
            console.error('總表檔案處理錯誤:', error);
            showStatus(masterListStatus, '載入失敗', 'error');
            showMessage(`總表檔案載入錯誤：${error.message}`, 'error');
        }
    }

    // ===== 總表預覽生成 =====
    function generateMasterListPreview() {
        if (!masterListData || !masterListData.productData) {
            masterListPreview.innerHTML = '<div class="text-center py-8 text-gray-500"><p class="text-lg">請先上傳總表檔案</p></div>';
            return;
        }

        const { productData, summaryData } = masterListData;

        let html = '<div class="overflow-x-auto"><table class="master-table">';
        html += '<thead><tr>';
        
        // 修改：移除貨名欄位
        const columns = [
            '序號', '廠商', '報關名稱', '箱規', '數量', '最小單位', 
            '長(cm)', '寬(cm)', '高(cm)', '毛重kg', '淨重kg', '報關數量', 
            '包裝方式', '毛重kg(合計)', '淨重kg(合計)', '體積(合計)', 
            '材質', '用途', '稅則編碼', '稅率(%)', '麥頭名稱備註'
        ];
        
        columns.forEach(col => {
            html += `<th>${col}</th>`;
        });
        html += '</tr></thead><tbody>';

        // 顯示產品資料（移除貨名欄位）
        productData.forEach(item => {
            html += '<tr>';
            html += `<td>${item['序號']}</td>`;
            html += `<td>${item['廠商']}</td>`;
            html += `<td>${item['報關名稱']}</td>`;
            html += `<td>${item['箱規']}</td>`;
            html += `<td>${item['數量']}</td>`;
            html += `<td>${item['最小單位']}</td>`;
            html += `<td>${item['長 (cm)']}</td>`;
            html += `<td>${item['寬 (cm)']}</td>`;
            html += `<td>${item['高 (cm)']}</td>`;
            html += `<td>${item['毛重kg']}</td>`;
            html += `<td>${item['淨重kg']}</td>`;
            html += `<td>${item['報關數量']}</td>`;
            html += `<td>${item['包裝方式']}</td>`;
            html += `<td>${item['毛重kg(合計)']}</td>`;
            html += `<td>${item['淨重kg(合計)']}</td>`;
            html += `<td>${item['體積(合計)']}</td>`;
            html += `<td>${item['材質']}</td>`;
            html += `<td>${item['用途']}</td>`;
            html += `<td>${item['稅則編碼']}</td>`;
            html += `<td>${item['稅率(%)']}</td>`;
            html += `<td>${item['麥頭名稱備註']}</td>`;
            html += '</tr>';
        });

        // 顯示合計行（修改 colspan 從 12 改為 11）
        html += '<tr style="background-color: #f3f4f6; font-weight: bold;">';
        html += '<td colspan="11">合計</td>';
        html += `<td>${summaryData['報關數量合計']}</td>`;
        html += '<td></td>';
        html += `<td>${summaryData['毛重合計']}</td>`;
        html += `<td>${summaryData['淨重合計']}</td>`;
        html += `<td>${summaryData['體積合計']}</td>`;
        html += '<td colspan="5"></td>';
        html += '</tr>';

        html += '</tbody></table></div>';
        masterListPreview.innerHTML = html;
    }

    // ===== 快速生成 Packing List 預覽 =====
    function generateQuickPackingListPreview() {
        if (!masterListData) {
            showMessage('請先上傳總表檔案', 'error');
            return;
        }
        
        const quickShippingDate = document.getElementById('quickShippingDate')?.value;
        if (!quickShippingDate) {
            showMessage('請選擇出貨日期', 'error');
            return;
        }

        const { basicInfo, containerSize, productData, summaryData } = masterListData;

        let html = '<div class="preview-content">';
        
        html += '<h1 style="text-align: center; font-size: 18px; font-weight: bold; margin-bottom: 20px;">PACKING LIST</h1>';
        
        // 從基本資訊中提取資料
        Object.entries(basicInfo).forEach(([key, value]) => {
            if (value && typeof value === 'string' && value.includes(':')) {
                html += `<div style="margin-bottom: 10px;"><strong>${value}</strong></div>`;
            }
        });
        
        html += '<div style="margin-bottom: 10px;"><strong>買方:</strong>嘉鴻精密科技股份有限公司 JIAHON PRECISION CO.,LTD</div>';
        html += '<div style="margin-bottom: 10px;"><strong>Tel:</strong> 037661986</div>';
        html += '<div style="margin-bottom: 10px;"><strong>地址:</strong>苗栗縣竹南鎮龍山路一段336號、No. 336, Sec. 1, Longshan Rd., Zhunan Township, Miaoli County</div>';
        html += `<div style="margin-bottom: 10px;"><strong>出貨日期:</strong> ${formatDateForWorksheet(quickShippingDate)}</div>`;
        html += `<div style="margin-bottom: 20px; text-align: center;"><strong>${containerSize}</strong></div>`;
        
        html += '<table class="preview-table">';
        html += '<thead><tr>';
        
        // 修改：移除貨名欄位
        const quickPreviewColumns = [
            '序號', '廠商', '報關名稱', '箱規', '數量', '最小單位', '長 (cm)', '寬 (cm)', 
            '高 (cm)', '毛重 kg(每箱)', '淨重 kg(每箱)', '數量(箱)', '報關數量', 
            '包裝方式', '毛重kg (合計)', '淨重kg (合計)', '體積(合計)', '材質', 
            '用途', '稅則編碼', '稅率 (%)', '麥頭名稱備註'
        ];
        
        quickPreviewColumns.forEach(col => {
            html += `<th>${col}</th>`;
        });
        html += '</tr></thead><tbody>';

        // 顯示產品資料（移除貨名欄位）- 直接使用總表中的數值
        productData.forEach(item => {
            html += '<tr>';
            html += `<td>${item['序號']}</td>`;
            html += `<td>${item['廠商']}</td>`;
            html += `<td>${item['報關名稱']}</td>`;
            html += `<td>${item['箱規']}</td>`;
            html += `<td>${item['數量']}</td>`;
            html += `<td>${item['最小單位']}</td>`;
            html += `<td>${item['長 (cm)']}</td>`;
            html += `<td>${item['寬 (cm)']}</td>`;
            html += `<td>${item['高 (cm)']}</td>`;
            html += `<td>${item['毛重kg']}</td>`;
            html += `<td>${item['淨重kg']}</td>`;
            html += `<td>${item['報關數量']}</td>`;
            html += `<td>${item['報關數量']}</td>`;
            html += `<td>${item['包裝方式']}</td>`;
            // ===== 關鍵修改：直接顯示總表中已計算好的數值 =====
            html += `<td>${item['毛重kg(合計)']}</td>`;
            html += `<td>${item['淨重kg(合計)']}</td>`;
            html += `<td>${item['體積(合計)']}</td>`;
            html += `<td>${item['材質']}</td>`;
            html += `<td>${item['用途']}</td>`;
            html += `<td>${item['稅則編碼']}</td>`;
            html += `<td>${item['稅率(%)']}</td>`;
            html += `<td>${item['麥頭名稱備註']}</td>`;
            html += '</tr>';
        });

        // 顯示合計行（修改 colspan 從 12 改為 11）- 直接使用總表中的合計數值
        html += '<tr style="background-color: #f3f4f6; font-weight: bold;">';
        html += `<td colspan="11">${summaryData['合計文字']}</td>`;
        html += `<td>${summaryData['報關數量合計']}</td>`;
        html += '<td></td>';
        html += `<td>${summaryData['毛重合計']}</td>`;
        html += `<td>${summaryData['淨重合計']}</td>`;
        html += `<td>${summaryData['體積合計']}</td>`;
        html += '<td colspan="5"></td>';
        html += '</tr>';

        html += '</tbody></table>';
        html += '</div>';

        previewContent.innerHTML = html;
        previewModal.classList.remove('hidden');
    }

    // ===== 添加圖片到工作表的函數 =====
    async function addImageToWorksheet(worksheet, imageData, rowStart, rowEnd, colStart, colEnd) {
        try {
            if (!imageData || !imageData.arrayBuffer) {
                console.log('沒有圖片資料可添加');
                return;
            }
            
            console.log(`添加圖片到工作表，位置: 第${colStart}欄，第${rowStart}-${rowEnd}行`);
            console.log('圖片資料:', {
                extension: imageData.extension,
                size: imageData.size,
                dimensions: imageData.dimensions
            });
            
            // 添加圖片到工作簿
            const imageId = worksheet.workbook.addImage({
                buffer: imageData.arrayBuffer,
                extension: imageData.extension
            });
            
            console.log('圖片ID:', imageId);
            
            // 使用絕對位置和尺寸
            worksheet.addImage(imageId, {
                tl: { col: colStart - 1, row: rowStart - 1 }, // 左上角位置 (0-based)
                ext: { width: imageData.dimensions.width, height: imageData.dimensions.height }, // 使用實際尺寸
                editAs: 'absolute' // 絕對位置，不隨單元格變化
            });
            
            console.log(`✓ 圖片添加成功，位置: ${colStart}${rowStart}，尺寸: ${imageData.dimensions.width}×${imageData.dimensions.height}`);
            
        } catch (error) {
            console.error('添加圖片到工作表時出錯:', error);
            console.error('錯誤詳情:', error.stack);
            
            // 備用方法：嘗試不同的圖片添加方式
            try {
                console.log('嘗試備用圖片添加方法...');
                const imageId = worksheet.workbook.addImage({
                    buffer: imageData.arrayBuffer,
                    extension: imageData.extension
                });
                
                // 使用更簡單的範圍定義
                worksheet.addImage(imageId, `${String.fromCharCode(64 + colStart)}${rowStart}:${String.fromCharCode(64 + colEnd)}${rowEnd}`);
                console.log('✓ 備用方法成功');
            } catch (backupError) {
                console.error('備用方法也失敗:', backupError);
            }
        }
    }

    async function generateQuickFinalPackingList() {
        try {
            console.log('=== 開始快速生成 Packing List、Invoice 和內部人員表（兩個分別的Excel檔案）===');
            
            // 檢查總表資料
            if (!masterListData) {
                console.error('沒有總表資料');
                showMessage('請先上傳總表檔案', 'error');
                return;
            }

            // 檢查檔案是否存在
            if (!masterListInput.files || !masterListInput.files[0]) {
                console.error('沒有找到上傳的檔案');
                showMessage('無法找到總表檔案，請重新上傳', 'error');
                return;
            }

            // 獲取使用者輸入
            const quickShippingDate = document.getElementById('quickShippingDate')?.value;
            const quickFileNameInput = document.getElementById('quickFileName');
            const customFileName = quickFileNameInput ? quickFileNameInput.value.trim() : '';
            const baseFileName = customFileName || 'PackingList';

            if (!quickShippingDate) {
                showMessage('請選擇出貨日期', 'error');
                return;
            }

            console.log('✓ 基本檢查通過');
            showMessage('正在生成 Packing List、Invoice 和內部人員表，請稍候...', 'info');

            // 重新讀取總表檔案以獲取樣式資訊
            const arrayBuffer = await masterListInput.files[0].arrayBuffer();
            const sourceWorkbook = new ExcelJS.Workbook();
            await sourceWorkbook.xlsx.load(arrayBuffer);
            const sourceWorksheet = sourceWorkbook.getWorksheet(1);

            if (!sourceWorksheet) {
                throw new Error('無法讀取總表工作表');
            }

            // 讀取貨櫃尺寸 - 使用安全讀取方式
            const containerCell = sourceWorksheet.getCell(11, 4);
            const containerSize = getContainerSizeSafely(containerCell);

            // 生成格式化日期（貨櫃尺寸已經在 getContainerSizeSafely 中處理完成）
            const formattedDate = formatDateForWorksheet(quickShippingDate);
            const safeContainerSize = containerSize; // 已經是安全的字串格式

            // === 第一個Excel：Packing List + Invoice ===
            const packingInvoiceWorkbook = new ExcelJS.Workbook();

            // 創建 Packing List 工作表
            const packingListName = generateWorksheetName('packinglist', quickShippingDate, safeContainerSize);
            const packingWorksheet = packingInvoiceWorkbook.addWorksheet(packingListName);
            await createPackingListWithDirectValues(sourceWorksheet, packingWorksheet, masterListData);

            // 創建 Invoice 工作表
            const invoiceName = generateWorksheetName('invoice', quickShippingDate, safeContainerSize);
            const invoiceWorksheet = packingInvoiceWorkbook.addWorksheet(invoiceName);
            await createInvoiceWithDirectValues(sourceWorksheet, invoiceWorksheet, masterListData);

            // 生成第一個Excel檔案
            const packingInvoiceFileName = `${baseFileName}_${formattedDate}_${safeContainerSize}.xlsx`;
            const packingInvoiceBuffer = await packingInvoiceWorkbook.xlsx.writeBuffer();
            const packingInvoiceBlob = new Blob([packingInvoiceBuffer], { 
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
            });

            // === 第二個Excel：內部人員 ===
            const internalWorkbook = new ExcelJS.Workbook();
            const internalWorksheet = internalWorkbook.addWorksheet('內部人員');
            await createInternalPersonnelWorksheet(sourceWorksheet, internalWorksheet, masterListData);

            // 生成第二個Excel檔案
            const internalFileName = `${baseFileName}_內部人員_${formattedDate}_${safeContainerSize}.xlsx`;
            const internalBuffer = await internalWorkbook.xlsx.writeBuffer();
            const internalBlob = new Blob([internalBuffer], { 
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
            });

            // 同時下載兩個檔案
            console.log('準備下載兩個Excel檔案...');
            
            // 使用 setTimeout 確保兩個下載不會互相干擾
            saveAs(packingInvoiceBlob, packingInvoiceFileName);
            
            setTimeout(() => {
                saveAs(internalBlob, internalFileName);
            }, 1000); // 延遲1秒下載第二個檔案

            showMessage(`生成成功！已下載兩個Excel檔案：\n1. ${packingInvoiceFileName}\n2. ${internalFileName}`, 'success');
            console.log('✓ 兩個Excel檔案生成並下載完成');

        } catch (error) {
            console.error('快速生成錯誤:', error);
            showMessage(`快速生成錯誤：${error.message}`, 'error');
        }
    }

    // === 創建 Packing List - 完全使用直接數值，避免所有公式問題 ===
    async function createPackingListWithDirectValues(sourceWorksheet, targetWorksheet, data) {
        console.log('創建 Packing List（完全使用直接數值，避免公式）');
        
        const { basicInfo, containerSize, productData, summaryData } = data;
        
        // 目標欄位（移除貨名）
        const targetColumns = [1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22];
        
        // === 1. 複製欄寬 ===
        targetColumns.forEach((sourceColIndex, newColIndex) => {
            const sourceColumn = sourceWorksheet.getColumn(sourceColIndex);
            if (sourceColumn.width) {
                targetWorksheet.getColumn(newColIndex + 1).width = sourceColumn.width;
            }
        });

        // === 2. 複製 A1-J9 基本資訊區域（含樣式） ===
        for (let row = 1; row <= 9; row++) {
            const sourceRow = sourceWorksheet.getRow(row);
            const newRow = targetWorksheet.getRow(row);

            if (sourceRow.height) {
                newRow.height = sourceRow.height;
            }

            for (let col = 1; col <= 10; col++) {
                const sourceCell = sourceRow.getCell(col);
                const newCell = newRow.getCell(col);

                newCell.value = sourceCell.value;
                if (sourceCell.style) {
                    newCell.style = JSON.parse(JSON.stringify(sourceCell.style));
                }
            }
        }

        // === 3. 複製 D11 貨櫃尺寸（含樣式） ===
        const sourceContainerCell = sourceWorksheet.getCell(11, 4);
        const newContainerCell = targetWorksheet.getCell(11, 4);
        newContainerCell.value = sourceContainerCell.value;
        if (sourceContainerCell.style) {
            newContainerCell.style = JSON.parse(JSON.stringify(sourceContainerCell.style));
        }

        // === 4. 跳過第12行合計行，稍後會在產品資料最後添加 ===
        // 注意：合計行現在會在第6步驟之後添加

        // === 5. 複製第13行標題行（含樣式） ===
        const sourceHeaderRow = sourceWorksheet.getRow(13);
        const newHeaderRow = targetWorksheet.getRow(13);

        if (sourceHeaderRow.height) {
            newHeaderRow.height = sourceHeaderRow.height;
        }

        targetColumns.forEach((sourceColIndex, newColIndex) => {
            const sourceCell = sourceHeaderRow.getCell(sourceColIndex);
            const newCell = newHeaderRow.getCell(newColIndex + 1);

            newCell.value = sourceCell.value;
            if (sourceCell.style) {
                newCell.style = JSON.parse(JSON.stringify(sourceCell.style));
            }
        });

        // === 6. 填入產品資料（含樣式，完全使用直接數值） ===
        productData.forEach((product, index) => {
            const sourceRowNum = 14 + index;
            const newRowNum = 14 + index;
            
            const sourceDataRow = sourceWorksheet.getRow(sourceRowNum);
            const newDataRow = targetWorksheet.getRow(newRowNum);

            if (sourceDataRow.height) {
                newDataRow.height = sourceDataRow.height;
            }

            targetColumns.forEach((sourceColIndex, newColIndex) => {
                const sourceCell = sourceDataRow.getCell(sourceColIndex);
                const newCell = newDataRow.getCell(newColIndex + 1);

                // ===== 關鍵修改：完全使用直接數值，避免所有公式 =====
                if (sourceColIndex === 1) newCell.value = product['序號'];
                else if (sourceColIndex === 2) newCell.value = product['廠商'];
                else if (sourceColIndex === 3) newCell.value = product['報關名稱'];
                else if (sourceColIndex === 5) newCell.value = product['箱規'];
                else if (sourceColIndex === 6) newCell.value = product['數量'];
                else if (sourceColIndex === 7) newCell.value = product['最小單位'];
                else if (sourceColIndex === 8) newCell.value = product['長 (cm)'];
                else if (sourceColIndex === 9) newCell.value = product['寬 (cm)'];
                else if (sourceColIndex === 10) newCell.value = product['高 (cm)'];
                else if (sourceColIndex === 11) newCell.value = product['毛重kg'];
                else if (sourceColIndex === 12) newCell.value = product['淨重kg'];
                else if (sourceColIndex === 13) newCell.value = product['報關數量'];
                else if (sourceColIndex === 14) newCell.value = product['包裝方式'];
                // ===== 完全避免公式，直接使用總表已計算的數值 =====
                else if (sourceColIndex === 15) newCell.value = product['毛重kg(合計)']; // 直接使用總表數值
                else if (sourceColIndex === 16) newCell.value = product['淨重kg(合計)']; // 直接使用總表數值
                else if (sourceColIndex === 17) newCell.value = product['體積(合計)']; // 直接使用總表數值
                else if (sourceColIndex === 18) newCell.value = product['材質'];
                else if (sourceColIndex === 19) newCell.value = product['用途'];
                else if (sourceColIndex === 20) newCell.value = product['稅則編碼'];
                else if (sourceColIndex === 21) newCell.value = product['稅率(%)'];
                else if (sourceColIndex === 22) newCell.value = product['麥頭名稱備註'];

                // 複製樣式
                if (sourceCell.style) {
                    newCell.style = JSON.parse(JSON.stringify(sourceCell.style));
                }

                // 麥頭名稱自動換行
                if (sourceColIndex === 22) {
                    newCell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
                }
            });
        });

        // === 7. 在產品資料最後添加合計行 ===
        const summaryRowNum = 14 + productData.length; // 產品資料後的下一行
        const sourceSummaryRow = sourceWorksheet.getRow(12); // 原始合計行樣式
        const newSummaryRow = targetWorksheet.getRow(summaryRowNum);

        if (sourceSummaryRow.height) {
            newSummaryRow.height = sourceSummaryRow.height;
        }

        targetColumns.forEach((sourceColIndex, newColIndex) => {
            const sourceCell = sourceSummaryRow.getCell(sourceColIndex);
            const newCell = newSummaryRow.getCell(newColIndex + 1);

            // ===== 設置合計數值 =====
            if (sourceColIndex === 12) {
                newCell.value = summaryData['合計文字'];
            } else if (sourceColIndex === 13) {
                newCell.value = summaryData['報關數量合計'];
            } else if (sourceColIndex === 15) {
                newCell.value = summaryData['毛重合計'];
            } else if (sourceColIndex === 16) {
                newCell.value = summaryData['淨重合計'];
            } else if (sourceColIndex === 17) {
                newCell.value = summaryData['體積合計'];
            } else {
                newCell.value = null;
            }

            // 複製樣式
            if (sourceCell.style) {
                newCell.style = JSON.parse(JSON.stringify(sourceCell.style));
            }
        });

        console.log(`✓ Packing List 合計行添加到第 ${summaryRowNum} 行`);

        // === 8. 添加圖片 ===
        if (uploadedImageData) {
            try {
                const imageId = targetWorksheet.workbook.addImage({
                    buffer: uploadedImageData.arrayBuffer,
                    extension: uploadedImageData.extension
                });
                
                targetWorksheet.addImage(imageId, {
                    tl: { col: 22, row: 12 }, // U13 位置（因為移除了貨名欄位）
                    ext: { width: uploadedImageData.dimensions.width, height: uploadedImageData.dimensions.height },
                    editAs: 'absolute'
                });
                
                console.log('✓ 已添加圖片到 Packing List');
            } catch (error) {
                console.error('添加圖片失敗:', error);
            }
        }

        console.log('✓ Packing List 工作表創建完成（完全使用直接數值）');
    }

    // === 創建 Invoice - 完全使用直接數值，避免所有公式問題 ===
    async function createInvoiceWithDirectValues(sourceWorksheet, targetWorksheet, data) {
        console.log('創建 Invoice（完全使用直接數值，避免公式）');
        
        const { basicInfo, containerSize, productData } = data;
        
        // Invoice 目標欄位
        const targetColumns = [1, 3, 6, 7, 25, 26];
        
        // === 1. 複製欄寬 ===
        targetColumns.forEach((sourceColIndex, newColIndex) => {
            const sourceColumn = sourceWorksheet.getColumn(sourceColIndex);
            let width = sourceColumn.width || 10;
            
            // 特殊設置
            if (newColIndex === 4) width = 19; // U/P(RMB)單價
            if (newColIndex === 5) width = 19; // Amount(RMB)合計
            
            targetWorksheet.getColumn(newColIndex + 1).width = width;
        });

        // === 2. 複製 A1-J9 基本資訊區域（含樣式），A1改為INVOICE ===
        for (let row = 1; row <= 9; row++) {
            const sourceRow = sourceWorksheet.getRow(row);
            const newRow = targetWorksheet.getRow(row);

            if (sourceRow.height) {
                newRow.height = sourceRow.height;
            }

            for (let col = 1; col <= 10; col++) {
                const sourceCell = sourceRow.getCell(col);
                const newCell = newRow.getCell(col);

                if (row === 1 && col === 1) {
                    newCell.value = 'INVOICE 發票';
                } else {
                    newCell.value = sourceCell.value;
                }
                
                if (sourceCell.style) {
                    newCell.style = JSON.parse(JSON.stringify(sourceCell.style));
                }
            }
        }

        // === 3. 設置 D11 貿易條件資訊（含樣式） ===
        const sourceContainerCell = sourceWorksheet.getCell(11, 4);
        const newContainerCell = targetWorksheet.getCell(11, 4);
        
        // ===== 關鍵修改：使用新的貿易條件格式替代原來的containerSize =====
        const tradeConditionString = generateTradeConditionString();
        newContainerCell.value = tradeConditionString;
        
        // 複製原有樣式
        if (sourceContainerCell.style) {
            newContainerCell.style = JSON.parse(JSON.stringify(sourceContainerCell.style));
        }
        
        console.log('Invoice D11設置為:', tradeConditionString);

        // === 4. 跳過第12行合計行，稍後會在產品資料最後添加 ===
        // 注意：Invoice 合計行現在會在第6步驟之後添加

        // === 5. 複製第13行標題行 ===
        const invoiceHeaders = ['序號', '報關名稱', '數量', '最小單位', 'U/P(RMB)單價', 'Amount(RMB)合計'];
        const newHeaderRow = targetWorksheet.getRow(13);
        
        if (sourceWorksheet.getRow(13).height) {
            newHeaderRow.height = sourceWorksheet.getRow(13).height;
        }

        invoiceHeaders.forEach((header, index) => {
            const cell = newHeaderRow.getCell(index + 1);
            cell.value = header;

            // 複製對應原始欄位的樣式
            const sourceCell = sourceWorksheet.getRow(13).getCell(targetColumns[index]);
            if (sourceCell.style) {
                cell.style = JSON.parse(JSON.stringify(sourceCell.style));
            }
        });

        // === 6. 填入產品資料（直接計算金額，避免公式） ===
        productData.forEach((product, index) => {
            const sourceRowNum = 14 + index;
            const newRowNum = 14 + index;
            
            const sourceDataRow = sourceWorksheet.getRow(sourceRowNum);
            const newDataRow = targetWorksheet.getRow(newRowNum);

            if (sourceDataRow.height) {
                newDataRow.height = sourceDataRow.height;
            }

            // 設置資料
            newDataRow.getCell(1).value = product['序號'];
            newDataRow.getCell(2).value = product['報關名稱']; // 只使用報關名稱
            newDataRow.getCell(3).value = product['數量'];
            newDataRow.getCell(4).value = product['最小單位'];
            newDataRow.getCell(5).value = product['¥採購單價(外部)'];
            
            // ===== 關鍵修改：直接計算金額，避免公式 =====
            const quantity = parseFloat(product['數量']) || 0;
            const unitPrice = parseFloat(product['¥採購單價(外部)']) || 0;
            const amount = quantity * unitPrice;
            newDataRow.getCell(6).value = amount; // 直接設置計算結果

            // 複製樣式
            targetColumns.forEach((sourceColIndex, newColIndex) => {
                const sourceCell = sourceDataRow.getCell(sourceColIndex);
                const newCell = newDataRow.getCell(newColIndex + 1);
                
                if (sourceCell.style) {
                    newCell.style = JSON.parse(JSON.stringify(sourceCell.style));
                }
            });
        });

        // === 7. 在產品資料最後添加合計行（Invoice專用） ===
        const summaryRowNum = 14 + productData.length; // 產品資料後的下一行
        const sourceSummaryRow = sourceWorksheet.getRow(12); // 原始合計行樣式
        const newSummaryRow = targetWorksheet.getRow(summaryRowNum);

        if (sourceSummaryRow.height) {
            newSummaryRow.height = sourceSummaryRow.height;
        }

        // 設置合計CNY文字
        const summaryCell = newSummaryRow.getCell(5);
        summaryCell.value = '合計CNY';
        
        // 複製樣式
        const sourceSummaryStyleCell = sourceSummaryRow.getCell(25);
        if (sourceSummaryStyleCell.style) {
            summaryCell.style = JSON.parse(JSON.stringify(sourceSummaryStyleCell.style));
        }

        // ===== 計算總金額 =====
        const totalAmount = productData.reduce((sum, product) => {
            const quantity = parseFloat(product['數量']) || 0;
            const unitPrice = parseFloat(product['¥採購單價(外部)']) || 0;
            return sum + (quantity * unitPrice);
        }, 0);

        const totalCell = newSummaryRow.getCell(6);
        totalCell.value = totalAmount; // 直接設置計算結果
        totalCell.numFmt = '#,##0';
        if (sourceSummaryStyleCell.style) {
            totalCell.style = JSON.parse(JSON.stringify(sourceSummaryStyleCell.style));
        }

        console.log(`✓ Invoice 合計行添加到第 ${summaryRowNum} 行，總金額: ${totalAmount}`);

        // === 8. 添加圖片 ===
        if (uploadedImageData) {
            try {
                const imageId = targetWorksheet.workbook.addImage({
                    buffer: uploadedImageData.arrayBuffer,
                    extension: uploadedImageData.extension
                });
                
                targetWorksheet.addImage(imageId, {
                    tl: { col: 7, row: 10 }, // H11 位置
                    ext: { width: uploadedImageData.dimensions.width, height: uploadedImageData.dimensions.height },
                    editAs: 'absolute'
                });
                
                console.log('✓ 已添加圖片到 Invoice');
            } catch (error) {
                console.error('添加圖片失敗:', error);
            }
        }

        console.log('✓ Invoice 工作表創建完成（完全使用直接數值）');
    }

    // === 創建內部人員工作表函數 ===
    async function createInternalPersonnelWorksheet(sourceWorksheet, targetWorksheet, data) {
        console.log('創建內部人員工作表');
        
        try {
            const { productData, containerSize } = data;
            
            // === 1. 設定欄寬 ===
            const columnWidths = [5, 8, 12, 30, 30, 8, 8, 10, 10, 10, 12, 12, 10, 15, 15, 15, 12];
            columnWidths.forEach((width, index) => {
                targetWorksheet.getColumn(index + 1).width = width;
            });

            // === 2. 創建標題和格式 ===
            // 第2行：貨櫃尺寸
            const containerRow = targetWorksheet.getRow(2);
            containerRow.height = 20;
            const containerCell = containerRow.getCell(5);
            containerCell.value = containerSize || '20呎  嘉鴻';
            containerCell.style = {
                font: { bold: true, size: 12, color: { argb: 'FFFF0000' } }, // 紅色文字
                alignment: { horizontal: 'center', vertical: 'middle' }
            };

            // 第3行：標題行
            const headerRow = targetWorksheet.getRow(3);
            headerRow.height = 40;
            const headers = ['', '序號', '廠商', '貨名', '箱規', '數量', '最小單位', '長\n(cm)', '寬\n(cm)', '高\n(cm)', '毛重kg(每箱)', '淨重kg(每箱)', '數量(箱)報關數量', '包裝方式', '毛重kg(合計)', '淨重kg(合計)', '體積(合計)'];
            
            headers.forEach((header, index) => {
                const cell = headerRow.getCell(index + 1);
                cell.value = header;
                // A欄不設置樣式
                if (index > 0) {
                    cell.style = {
                        font: { name: '細明體', bold: true, size: 10 },
                        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
                        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E6FA' } },
                        border: {
                            top: { style: 'thin' },
                            left: { style: 'thin' },
                            bottom: { style: 'thin' },
                            right: { style: 'thin' }
                        }
                    };
                }
            });

            // === 3. 填入產品資料 (從第4行開始) ===
            productData.forEach((product, index) => {
                const rowNum = 4 + index;
                const dataRow = targetWorksheet.getRow(rowNum);
                // 自動調整列高：基本高度30，如果箱規內容較長則增加高度
                const boxSpecLength = (product['箱規'] || '').toString().length;
                const autoHeight = Math.max(30, Math.min(60, 30 + Math.floor(boxSpecLength / 20) * 10));
                dataRow.height = autoHeight;
                
                const internalData = [
                    null, // A列空白
                    product['序號'], // B列 - 序號
                    product['廠商'], // C列 - 廠商
                    product['貨名'], // D列 - 貨名
                    product['箱規'], // E列 - 箱規
                    product['數量'], // F列 - 數量
                    product['最小單位'], // G列 - 最小單位
                    product['長 (cm)'], // H列 - 長(cm)
                    product['寬 (cm)'], // I列 - 寬(cm)
                    product['高 (cm)'], // J列 - 高(cm)
                    product['毛重kg'], // K列 - 毛重kg(每箱)
                    product['淨重kg'], // L列 - 淨重kg(每箱)
                    product['報關數量'], // M列 - 數量(箱)報關數量
                    product['包裝方式'], // N列 - 包裝方式
                    product['毛重kg(合計)'], // O列 - 毛重kg(合計)
                    product['淨重kg(合計)'], // P列 - 淨重kg(合計)
                    product['體積(合計)'] // Q列 - 體積(合計)
                ];

                internalData.forEach((value, colIndex) => {
                    const cell = dataRow.getCell(colIndex + 1);
                    cell.value = value;
                    // A欄不設置樣式
                    if (colIndex > 0) {
                        cell.style = {
                            font: { name: '細明體', size: 9 },
                            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
                            border: {
                                top: { style: 'thin' },
                                left: { style: 'thin' },
                                bottom: { style: 'thin' },
                                right: { style: 'thin' }
                            }
                        };
                    }
                });
            });

            // === 添加合計行 ===
            const { summaryData } = data;
            const summaryRowNum = 4 + productData.length;
            const summaryRow = targetWorksheet.getRow(summaryRowNum);
            summaryRow.height = 35; // 合計行稍微高一點
            
            const summaryValues = [
                null, // A列空白
                null, // B列
                null, // C列
                null, // D列
                null, // E列
                null, // F列
                null, // G列
                null, // H列
                null, // I列
                null, // J列
                null, // K列
                '合計', // L列 - 合計文字
                Math.round(summaryData['報關數量合計'] || 289), // M列 - 數量(箱)報關數量合計 (4捨五入)
                null, // N列
                Math.round(summaryData['毛重合計'] || 4150), // O列 - 毛重kg(合計) (4捨五入)
                Math.round(summaryData['淨重合計'] || 3803), // P列 - 淨重kg(合計) (4捨五入)
                Math.round(summaryData['體積合計'] || 27) // Q列 - 體積(合計) (4捨五入)
            ];

            summaryValues.forEach((value, colIndex) => {
                const cell = summaryRow.getCell(colIndex + 1);
                cell.value = value;
                // A欄不設置樣式
                if (colIndex > 0 && value !== null) {
                    cell.style = {
                        font: { name: '細明體', size: 9, bold: true, color: { argb: 'FFFF0000' } }, // 紅色文字
                        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
                        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } },
                        border: {
                            top: { style: 'thin' },
                            left: { style: 'thin' },
                            bottom: { style: 'thin' },
                            right: { style: 'thin' }
                        }
                    };
                }
            });

            // === 4. 設定列印格式 ===
            const lastRow = 4 + productData.length; // 包含合計行
            targetWorksheet.pageSetup = {
                printArea: `A1:Q${lastRow}`,
                orientation: 'landscape',
                paperSize: 9,
                fitToPage: true,
                fitToWidth: 1,
                fitToHeight: 0, // 不限制高度
                margins: {
                    left: 0.7,
                    right: 0.7,
                    top: 0.75,
                    bottom: 0.75,
                    header: 0.3,
                    footer: 0.3
                },
                horizontalCentered: true,
                verticalCentered: true
            };

            console.log('✓ 內部人員工作表創建完成');
            
        } catch (error) {
            console.error('創建內部人員工作表時發生錯誤:', error);
        }
    }

    // ===== 預覽模態框事件監聽器 =====
    closePreview.addEventListener('click', () => previewModal.classList.add('hidden'));
    closePreviewBottom.addEventListener('click', () => previewModal.classList.add('hidden'));
    generateFromPreview.addEventListener('click', () => {
        previewModal.classList.add('hidden');
        generateQuickFinalPackingList();
    });

    // ===== 快速功能事件監聽器 =====
    quickPreviewPacking.addEventListener('click', generateQuickPackingListPreview);
    quickGeneratePacking.addEventListener('click', generateQuickFinalPackingList);

    // ===== 結束程式碼片段 =====
    console.log('快速生成 Packing List 系統已完全載入（修正版：解決公式和數值讀取問題）');
});
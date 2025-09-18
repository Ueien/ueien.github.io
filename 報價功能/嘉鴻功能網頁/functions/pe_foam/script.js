// PE泡棉報價計算器 JavaScript - 雙模式版本

document.addEventListener('DOMContentLoaded', function() {
    // 獲取DOM元素
    const calculateBtn = document.getElementById('calculate-btn');
    const resetBtn = document.getElementById('reset-btn');
    const exportExcelBtn = document.getElementById('export-excel-btn');
    const addRowBtn = document.getElementById('add-row-btn');
    const clearAllBtn = document.getElementById('clear-all-btn');
    const resultSection = document.getElementById('result-section');
    const sheetModeRadio = document.getElementById('sheet-mode');
    const roundModeRadio = document.getElementById('round-mode');
    const dataTableBody = document.getElementById('data-table-body');
    const resultTableBody = document.getElementById('result-table-body');
    const totalItemsSpan = document.getElementById('total-items');
    const grandTotalSpan = document.getElementById('grand-total');

    // 雙模式數據存儲
    const modeData = {
        sheet: {
            inputData: [],
            calculationResults: []
        },
        round: {
            inputData: [],
            calculationResults: []
        }
    };

    let currentMode = 'sheet'; // 當前模式

    // 初始化
    setupInputValidation();
    setupEventListeners();
    initializeDefaultRow();

    function setupEventListeners() {
        addRowBtn.addEventListener('click', addNewRow);
        clearAllBtn.addEventListener('click', clearCurrentModeData);
        calculateBtn.addEventListener('click', calculateAllData);
        resetBtn.addEventListener('click', resetAll);
        exportExcelBtn.addEventListener('click', exportToExcel);

        // 模式切換事件
        sheetModeRadio.addEventListener('change', () => switchMode('sheet'));
        roundModeRadio.addEventListener('change', () => switchMode('round'));

        // 為現有行設置事件監聽器
        setupRowEventListeners();
    }

    function initializeDefaultRow() {
        // 初始化片材模式的默認行
        saveCurrentInputData();
        updateDisplay();
    }

    function switchMode(newMode) {
        if (currentMode === newMode) return;

        // 保存當前模式的輸入數據
        saveCurrentInputData();

        // 切換模式
        currentMode = newMode;

        // 載入新模式的數據
        loadModeData();

        // 更新顯示
        updateDisplay();

        console.log(`已切換到${newMode === 'sheet' ? '片材' : '圓形泡棉'}模式`);
    }

    function saveCurrentInputData() {
        const rows = document.querySelectorAll('.data-row');
        const inputData = [];

        rows.forEach(row => {
            const thickness = row.querySelector('.thickness-input').value;
            const width = row.querySelector('.width-input').value;
            const length = row.querySelector('.length-input').value;
            const rmbCost = row.querySelector('.rmb-cost-input').value;
            const quantity = row.querySelector('.quantity-input').value;

            inputData.push({
                thickness,
                width,
                length,
                rmbCost,
                quantity
            });
        });

        modeData[currentMode].inputData = inputData;
    }

    function loadModeData() {
        const inputData = modeData[currentMode].inputData;

        // 清空現有表格
        dataTableBody.innerHTML = '';

        // 如果沒有數據，創建一個空行
        if (inputData.length === 0) {
            addEmptyRow();
        } else {
            // 載入保存的數據
            inputData.forEach(data => {
                addRowWithData(data);
            });
        }

        setupRowEventListeners();
    }

    function addEmptyRow() {
        const newRow = createRowElement({
            thickness: '',
            width: '',
            length: '',
            rmbCost: '',
            quantity: ''
        });
        dataTableBody.appendChild(newRow);
    }

    function addRowWithData(data) {
        const newRow = createRowElement(data);
        dataTableBody.appendChild(newRow);
    }

    function createRowElement(data) {
        const newRow = document.createElement('tr');
        newRow.className = 'data-row';
        newRow.innerHTML = `
            <td>
                <button type="button" class="delete-row-btn">❌</button>
            </td>
            <td>
                <input type="number" class="thickness-input" step="1" min="1" placeholder="0" value="${data.thickness}">
            </td>
            <td>
                <input type="number" class="width-input" step="1" min="1" placeholder="0" value="${data.width}">
            </td>
            <td>
                <input type="number" class="length-input" step="1" min="1" placeholder="0" value="${data.length}">
            </td>
            <td>
                <input type="number" class="rmb-cost-input" step="0.1" min="0" placeholder="0.0" value="${data.rmbCost}">
            </td>
            <td>
                <input type="number" class="quantity-input" step="1" min="1" placeholder="0" value="${data.quantity}">
            </td>
        `;
        return newRow;
    }

    function updateDisplay() {
        const results = modeData[currentMode].calculationResults;

        if (results.length > 0) {
            displayResults(results);
        } else {
            hideResults();
        }
    }

    function setupRowEventListeners() {
        // 為所有刪除按鈕添加事件監聽器
        document.querySelectorAll('.delete-row-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                deleteRow(this);
            });
        });

        // 為所有輸入框添加驗證
        setupInputValidationForRows();
    }

    function setupInputValidation() {
        // 初始化輸入驗證將在 setupInputValidationForRows 中完成
    }

    function setupInputValidationForRows() {
        // 人民幣成本輸入驗證
        document.querySelectorAll('.rmb-cost-input').forEach(input => {
            input.addEventListener('input', function(e) {
                let value = this.value;
                const cursorPosition = this.selectionStart;

                const cleanValue = value.replace(/[^0-9.]/g, '');
                const parts = cleanValue.split('.');
                let finalValue = parts[0];

                if (parts.length > 1) {
                    finalValue += '.' + (parts[1] ? parts[1].substring(0, 1) : '');
                }

                if (finalValue.startsWith('.')) {
                    finalValue = '0' + finalValue;
                }

                if (this.value !== finalValue) {
                    this.value = finalValue;
                    this.setSelectionRange(cursorPosition, cursorPosition);
                }
            });
        });

        // 整數輸入驗證
        document.querySelectorAll('.thickness-input, .width-input, .length-input, .quantity-input').forEach(input => {
            input.addEventListener('input', function() {
                this.value = this.value.replace(/[^0-9]/g, '');
            });
        });
    }

    function addNewRow() {
        const newRow = createRowElement({
            thickness: '',
            width: '',
            length: '',
            rmbCost: '',
            quantity: ''
        });

        dataTableBody.appendChild(newRow);
        setupRowEventListeners();

        // 動畫效果
        newRow.style.opacity = '0';
        newRow.style.transform = 'translateY(20px)';
        setTimeout(() => {
            newRow.style.transition = 'all 0.3s ease';
            newRow.style.opacity = '1';
            newRow.style.transform = 'translateY(0)';
        }, 10);
    }

    function deleteRow(deleteBtn) {
        const row = deleteBtn.closest('tr');
        if (dataTableBody.children.length > 1) {
            row.style.transition = 'all 0.3s ease';
            row.style.opacity = '0';
            row.style.transform = 'translateX(-100px)';
            setTimeout(() => {
                row.remove();
            }, 300);
        } else {
            alert('至少需要保留一行資料');
        }
    }

    function clearCurrentModeData() {
        if (confirm(`確定要清空${currentMode === 'sheet' ? '片材' : '圓形泡棉'}模式的所有資料嗎？`)) {
            // 清空當前模式的數據
            modeData[currentMode].inputData = [];
            modeData[currentMode].calculationResults = [];

            // 重新載入（會創建一個空行）
            loadModeData();
            hideResults();
        }
    }

    function calculateAllData() {
        // 先保存輸入數據
        saveCurrentInputData();

        const inputData = modeData[currentMode].inputData;
        const results = [];
        let hasValidData = false;

        inputData.forEach((data, index) => {
            const thickness = parseInt(data.thickness);
            const width = parseInt(data.width);
            const length = parseInt(data.length);
            const rmbCost = parseFloat(data.rmbCost);
            const quantity = parseInt(data.quantity);

            // 驗證該行資料
            if (thickness && width && length && rmbCost && quantity &&
                thickness > 0 && width > 0 && length > 0 && rmbCost > 0 && quantity > 0) {

                const result = calculateSingleItem(thickness, width, length, rmbCost, quantity, index + 1);
                results.push(result);
                hasValidData = true;
            }
        });

        if (!hasValidData) {
            alert('請填寫至少一行完整的有效資料');
            return;
        }

        // 保存計算結果
        modeData[currentMode].calculationResults = results;
        displayResults(results);

        // 添加計算動畫效果
        calculateBtn.style.transform = 'scale(0.95)';
        calculateBtn.style.opacity = '0.8';
        setTimeout(() => {
            calculateBtn.style.transform = 'scale(1)';
            calculateBtn.style.opacity = '1';
        }, 150);
    }

    function calculateSingleItem(t, w, l, rmbCost, quantity, rowNumber) {
        const isSheetMode = currentMode === 'sheet';

        // 計算體積
        const volume = t * w * l;

        // 計算運費
        let freight;
        if (isSheetMode) {
            freight = calculateSheetFreight(t, w, l);
        } else {
            freight = calculateRoundFreight(t, w, l);
        }

        // 計算售價和總價
        const multiplier = 8.6;
        const unitPrice = rmbCost * multiplier + freight;
        const totalPrice = unitPrice * quantity;

        return {
            rowNumber,
            thickness: t,
            width: w,
            length: l,
            volume,
            rmbCost,
            multiplier,
            freight,
            unitPrice,
            quantity,
            totalPrice
        };
    }

    // 片材運費計算
    function calculateSheetFreight(t, w, l) {
        const tConverted = t / 1000;
        const wConverted = w / 100;
        const lConverted = l / 100;

        let tFormatted;
        if (t < 10) {
            tFormatted = tConverted.toFixed(3);
        } else {
            tFormatted = tConverted.toFixed(2);
        }

        const wFormatted = wConverted.toFixed(2);
        const lFormatted = lConverted.toFixed(2);

        const freight = parseFloat(tFormatted) * parseFloat(wFormatted) * parseFloat(lFormatted) * 3000;
        return freight;
    }

    // 圓形泡棉運費計算
    function calculateRoundFreight(t, w, l) {
        const step1 = l / 1000;           // L/1000
        const step2 = step1 / 2;          // /2
        const step3 = Math.pow(step2, 2); // ^2
        const step4 = t / 1000;           // t/1000
        const freight = step3 * 3.14 * step4 * 3000;

        return Math.round(freight * 10) / 10;
    }

    function displayResults(results) {
        // 清空結果表格
        resultTableBody.innerHTML = '';

        let grandTotal = 0;

        results.forEach(result => {
            const row = document.createElement('tr');

            row.innerHTML = `
                <td>${result.rowNumber}</td>
                <td>${result.thickness}</td>
                <td>${result.width}</td>
                <td>${result.length}</td>
                <td>${result.volume.toLocaleString()}</td>
                <td class="highlight-cell">${result.rmbCost.toFixed(1)}</td>
                <td>${result.multiplier}</td>
                <td class="highlight-cell">${result.freight.toFixed(2)}</td>
                <td class="highlight-cell">${Math.round(result.unitPrice).toLocaleString()}</td>
                <td>${result.quantity.toLocaleString()}</td>
                <td class="total-cell">${Math.round(result.totalPrice).toLocaleString()}</td>
            `;

            resultTableBody.appendChild(row);
            grandTotal += result.totalPrice;
        });

        // 更新總計資訊
        totalItemsSpan.textContent = results.length;
        grandTotalSpan.textContent = Math.round(grandTotal).toLocaleString() + ' NTD';

        // 顯示結果區域和Excel按鈕
        resultSection.style.display = 'block';

        // 檢查是否有任一模式有計算結果
        const hasAnyResults = modeData.sheet.calculationResults.length > 0 ||
                              modeData.round.calculationResults.length > 0;
        exportExcelBtn.style.display = hasAnyResults ? 'flex' : 'none';

        // 滾動到結果區域
        resultSection.scrollIntoView({
            behavior: 'smooth',
            block: 'start'
        });
    }

    function hideResults() {
        resultSection.style.display = 'none';

        // 檢查是否有任一模式有計算結果
        const hasAnyResults = modeData.sheet.calculationResults.length > 0 ||
                              modeData.round.calculationResults.length > 0;
        exportExcelBtn.style.display = hasAnyResults ? 'flex' : 'none';
    }

    function resetAll() {
        if (confirm('確定要重置所有模式的資料和結果嗎？')) {
            // 清空所有模式的數據
            modeData.sheet.inputData = [];
            modeData.sheet.calculationResults = [];
            modeData.round.inputData = [];
            modeData.round.calculationResults = [];

            // 重置為片材模式
            currentMode = 'sheet';
            sheetModeRadio.checked = true;

            // 重新載入
            loadModeData();
            hideResults();

            window.scrollTo({
                top: 0,
                behavior: 'smooth'
            });
        }
    }

    function exportToExcel() {
        const sheetResults = modeData.sheet.calculationResults;
        const roundResults = modeData.round.calculationResults;

        if (sheetResults.length === 0 && roundResults.length === 0) {
            alert('沒有計算結果可以下載');
            return;
        }

        // 創建工作簿
        const workbook = XLSX.utils.book_new();

        // 表頭
        const headers = [
            '編號', '厚度(t)', '寬度(W)', '長度(L)', '體積',
            '人民幣成本', '倍數', '運費(NTD)', '單個售價(NTD)', '需求量', '總價(NTD)'
        ];

        // 創建片材工作表
        if (sheetResults.length > 0) {
            const sheetData = [headers];

            sheetResults.forEach(result => {
                sheetData.push([
                    result.rowNumber,
                    result.thickness,
                    result.width,
                    result.length,
                    result.volume,
                    parseFloat(result.rmbCost.toFixed(1)),
                    result.multiplier,
                    parseFloat(result.freight.toFixed(2)),
                    Math.round(result.unitPrice),
                    result.quantity,
                    Math.round(result.totalPrice)
                ]);
            });

            // 添加總計行
            const sheetTotal = sheetResults.reduce((sum, result) => sum + result.totalPrice, 0);
            sheetData.push([]);
            sheetData.push(['總計項目數', sheetResults.length, '', '', '', '', '', '', '', '總金額', Math.round(sheetTotal)]);

            const sheetWorksheet = XLSX.utils.aoa_to_sheet(sheetData);
            XLSX.utils.book_append_sheet(workbook, sheetWorksheet, 'PE片材泡棉');
        }

        // 創建圓形工作表
        if (roundResults.length > 0) {
            const roundData = [headers];

            roundResults.forEach(result => {
                roundData.push([
                    result.rowNumber,
                    result.thickness,
                    result.width,
                    result.length,
                    result.volume,
                    parseFloat(result.rmbCost.toFixed(1)),
                    result.multiplier,
                    parseFloat(result.freight.toFixed(2)),
                    Math.round(result.unitPrice),
                    result.quantity,
                    Math.round(result.totalPrice)
                ]);
            });

            // 添加總計行
            const roundTotal = roundResults.reduce((sum, result) => sum + result.totalPrice, 0);
            roundData.push([]);
            roundData.push(['總計項目數', roundResults.length, '', '', '', '', '', '', '', '總金額', Math.round(roundTotal)]);

            const roundWorksheet = XLSX.utils.aoa_to_sheet(roundData);
            XLSX.utils.book_append_sheet(workbook, roundWorksheet, 'PE圓形泡棉');
        }

        // 下載檔案
        const fileName = `PE泡棉報價_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(workbook, fileName);

        // 添加下載動畫效果
        exportExcelBtn.style.transform = 'scale(0.95)';
        setTimeout(() => {
            exportExcelBtn.style.transform = 'scale(1)';
        }, 150);
    }

    // 添加頁面載入動畫
    document.body.style.opacity = '0';
    setTimeout(() => {
        document.body.style.transition = 'opacity 0.5s ease';
        document.body.style.opacity = '1';
    }, 100);
});
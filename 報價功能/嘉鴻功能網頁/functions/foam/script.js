// EVA泡棉報價計算器
class FoamCalculator {
    constructor() {
        this.currentExchangeRate = null; // 儲存當前匯率
        this.initializeElements();
        this.bindEvents();
        this.fetchExchangeRate(); // 初始化時獲取匯率
        console.log('EVA泡棉計算器已啟動');
    }

    initializeElements() {
        this.elements = {
            // 輸入元素
            length: document.getElementById('length'),
            width: document.getElementById('width'),
            thickness: document.getElementById('thickness'),
            exchangeRate: document.getElementById('exchange-rate'),
            unitPrice: document.getElementById('unit-price'),
            quantity: document.getElementById('quantity'),
            containerFreight: document.getElementById('container-freight'),
            seaExpressFreight: document.getElementById('sea-express-freight'),

            // 按鈕
            calculateBtn: document.getElementById('calculate-btn'),
            resetBtn: document.getElementById('reset-btn'),
            downloadBtn: document.getElementById('download-btn'),
            resultSection: document.getElementById('result-section'),
            exchangeRateLabel: document.getElementById('exchange-rate-label')
        };
    }

    bindEvents() {
        // 計算按鈕
        if (this.elements.calculateBtn) {
            this.elements.calculateBtn.addEventListener('click', () => this.calculate());
        }
        
        // 重置按鈕
        if (this.elements.resetBtn) {
            this.elements.resetBtn.addEventListener('click', () => this.reset());
        }
        
        // 下載按鈕
        if (this.elements.downloadBtn) {
            this.elements.downloadBtn.addEventListener('click', () => this.downloadExcel());
        }
        
        // 報價選項點擊事件
        this.bindQuoteOptionEvents();
        
        // 即時計算
        this.bindInputEvents();
        
        // 鍵盤快捷鍵
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.key === 'Enter') {
                this.calculate();
            } else if (e.ctrlKey && e.key === 'r') {
                e.preventDefault();
                this.reset();
            }
        });
    }

    // 綁定報價選項點擊事件
    bindQuoteOptionEvents() {
        const seaTab = document.getElementById('sea-tab');
        const containerTab = document.getElementById('container-tab');
        
        if (seaTab) {
            seaTab.addEventListener('click', () => this.switchTab('sea'));
        }
        
        if (containerTab) {
            containerTab.addEventListener('click', () => this.switchTab('container'));
        }
    }

    // 切換頁籤
    switchTab(type) {
        // 切換頁籤按鈕狀態
        const seaTab = document.getElementById('sea-tab');
        const containerTab = document.getElementById('container-tab');
        const seaPanel = document.getElementById('sea-panel');
        const containerPanel = document.getElementById('container-panel');
        
        // 重置所有頁籤狀態
        seaTab.classList.remove('active');
        containerTab.classList.remove('active');
        seaPanel.classList.remove('active');
        containerPanel.classList.remove('active');
        
        // 切換CBM結果區醒目顯示
        const cbmUnder8Block = document.getElementById('cbm-under-8');
        const cbmOver8Block = document.getElementById('cbm-over-8');
        cbmUnder8Block.classList.remove('highlight');
        cbmOver8Block.classList.remove('highlight');
        
        if (type === 'sea') {
            seaTab.classList.add('active');
            seaPanel.classList.add('active');
            cbmUnder8Block.classList.add('highlight');
        } else {
            containerTab.classList.add('active');
            containerPanel.classList.add('active');
            cbmOver8Block.classList.add('highlight');
        }
    }

    // 選擇報價選項 (保留以兼容，但改為調用switchTab)
    selectQuoteOption(type) {
        this.switchTab(type);
    }

    bindInputEvents() {
        const inputs = [
            this.elements.length,
            this.elements.width,
            this.elements.thickness,
            this.elements.exchangeRate,
            this.elements.unitPrice,
            this.elements.containerFreight,
            this.elements.seaExpressFreight
        ].filter(el => el !== null);

        inputs.forEach(element => {
            element.addEventListener('input', () => this.updateQuantityRange());
        });

        // 泡棉片數特殊處理
        if (this.elements.quantity) {
            // 限制只能輸入整數
            this.elements.quantity.addEventListener('input', (e) => {
                // 移除小數點和非數字字符
                let value = e.target.value.replace(/[^\d]/g, '');
                e.target.value = value;
            });
            
            this.elements.quantity.addEventListener('blur', () => this.calculate());
            this.elements.quantity.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    this.calculate();
                }
                // 阻止輸入小數點和其他非數字字符
                if (e.key === '.' || e.key === ',' || e.key === '-' || e.key === '+') {
                    e.preventDefault();
                }
            });
        }

        // 單價範圍限制處理
        if (this.elements.unitPrice) {
            // 限制只能輸入整數
            this.elements.unitPrice.addEventListener('input', (e) => {
                // 只移除小數點和非數字字符，不限制範圍（在blur時才限制）
                let value = e.target.value.replace(/[^\d]/g, '');
                e.target.value = value;
            });
            
            this.elements.unitPrice.addEventListener('blur', () => {
                // 確保值在範圍內
                const value = parseInt(this.elements.unitPrice.value);
                if (isNaN(value) || value < 16000) {
                    this.elements.unitPrice.value = '16000';
                    this.showMessage('單價最小值為 16,000', 'error');
                } else if (value > 35000) {
                    this.elements.unitPrice.value = '35000';
                    this.showMessage('單價最大值為 35,000', 'error');
                }
                this.updateQuantityRange();
            });
            
            this.elements.unitPrice.addEventListener('keydown', (e) => {
                // 阻止輸入小數點和其他非數字字符
                if (e.key === '.' || e.key === ',' || e.key === '-' || e.key === '+') {
                    e.preventDefault();
                }
            });
        }
    }

    // 獲取人民幣匯率
    async fetchExchangeRate() {
        try {
            console.log('開始獲取人民幣匯率...');
            
            // 使用 Exchange Rates API (免費，無需API key)
            const response = await fetch('https://api.exchangerate-api.com/v4/latest/CNY');
            
            if (!response.ok) {
                throw new Error('無法獲取匯率數據');
            }
            
            const data = await response.json();
            const cnyToTwd = data.rates.TWD;
            
            if (cnyToTwd) {
                // 四捨五入到小數點第二位
                this.currentExchangeRate = Math.round(cnyToTwd * 100) / 100;
                
                // 更新標籤顯示
                if (this.elements.exchangeRateLabel) {
                    this.elements.exchangeRateLabel.textContent = 
                        `匯率(人民幣CNY，目前匯率${this.currentExchangeRate.toFixed(2)})`;
                }
                
                // 可選：自動填入匯率到輸入框
                if (this.elements.exchangeRate && !this.elements.exchangeRate.value) {
                    this.elements.exchangeRate.value = this.currentExchangeRate.toFixed(2);
                }
                
                console.log(`匯率獲取成功: 1 CNY = ${this.currentExchangeRate.toFixed(2)} TWD`);
                
            } else {
                throw new Error('匯率數據格式錯誤');
            }
            
        } catch (error) {
            console.error('獲取匯率失敗:', error);
            
            // 更新標籤顯示錯誤狀態
            if (this.elements.exchangeRateLabel) {
                this.elements.exchangeRateLabel.textContent = '匯率(人民幣CNY，獲取失敗)';
            }
            
            // 顯示錯誤訊息給用戶
            this.showMessage('無法獲取當前匯率，請手動輸入', 'error');
        }
    }

    // 更新片數範圍 - 修改為起訂量+1到優惠量-1
    updateQuantityRange() {
        const values = this.getInputValues();
        
        // 檢查基本輸入是否完整
        const required = ['length', 'width', 'thickness', 'unitPrice'];
        const isValid = required.every(key => values[key] > 0);
        
        if (isValid) {
            const cbm = this.calculateFoamCbm(values.length, values.width, values.thickness);
            const twdCost = Number((values.unitPrice * cbm).toFixed(2));
            
            if (twdCost > 0) {
                // 計算MOQ範圍
                const moq_temp = Math.round(5000 / twdCost);
                const moq2 = Math.round(8 / cbm);
                const moq1 = Math.ceil(5000 / ((twdCost * 2.5 * moq_temp + cbm * values.seaExpressFreight * moq_temp) / moq_temp));
                
                // 修改範圍：起訂量+1 到 優惠量-1
                const minQuantity = moq1 + 1;
                const maxQuantity = moq2 - 1;
                
                // 更新標籤和輸入框屬性
                const quantityLabel = document.getElementById('quantity-label');
                quantityLabel.textContent = `片數 (範圍=${minQuantity}~${maxQuantity})`;

                this.elements.quantity.min = minQuantity;
                this.elements.quantity.max = maxQuantity;
                this.elements.quantity.placeholder = `請輸入片數 (${minQuantity}~${maxQuantity})`;
            }
        } else {
            // 重置為預設
            const quantityLabel = document.getElementById('quantity-label');
            quantityLabel.textContent = '片數';
            this.elements.quantity.removeAttribute('min');
            this.elements.quantity.removeAttribute('max');
            this.elements.quantity.placeholder = '請輸入片數';
        }
    }

    // 獲取輸入值 - 修改預設值為23235
    getInputValues() {
        return {
            length: parseFloat(this.elements.length.value) || 0,
            width: parseFloat(this.elements.width.value) || 0,
            thickness: parseFloat(this.elements.thickness.value) || 0,
            exchangeRate: parseFloat(this.elements.exchangeRate.value) || 0,
            unitPrice: parseFloat(this.elements.unitPrice.value) || 23235,
            quantity: parseInt(this.elements.quantity.value) || 0,
            containerFreight: parseFloat(this.elements.containerFreight.value) || 2000,
            seaExpressFreight: parseFloat(this.elements.seaExpressFreight.value) || 6000
        };
    }

    // 驗證輸入值 - 修改片數範圍驗證
    validateInputs(values) {
        const required = ['length', 'width', 'thickness', 'exchangeRate', 'unitPrice', 'containerFreight', 'seaExpressFreight'];
        const isValid = required.every(key => values[key] > 0);
        if (!isValid) return false;
        
        // 檢查單價範圍
        if (values.unitPrice < 16000 || values.unitPrice > 35000) {
            this.showError('單價需要在 16,000~35,000 範圍內');
            return false;
        }
        
        // 檢查片數範圍 - 使用動態範圍
        const minQuantity = parseInt(this.elements.quantity.min) || 1;
        const maxQuantity = parseInt(this.elements.quantity.max) || Infinity;

        if (values.quantity !== 0 && (values.quantity < minQuantity || values.quantity > maxQuantity)) {
            this.showError(`EVA泡棉片數需要在${minQuantity}~${maxQuantity}片之間`);
            return false;
        }
        return true;
    }

    // 計算功能 - 修改邏輯
    calculate() {
        console.log('開始計算EVA泡棉報價');
        const values = this.getInputValues();
        
        // 檢查片數範圍
        const minQuantity = parseInt(this.elements.quantity.min) || 1;
        const maxQuantity = parseInt(this.elements.quantity.max) || Infinity;

        if (values.quantity !== 0 && (values.quantity < minQuantity || values.quantity > maxQuantity)) {
            this.clearResults();
            
            // 顯示片數範圍提醒
            this.showQuantityRangeError(minQuantity, maxQuantity, values.quantity);
            this.elements.resultSection.classList.add('show');
            return;
        }
        
        if (!this.validateInputs(values)) {
            this.clearResults();
            this.elements.resultSection.classList.remove('show');
            return;
        }

        try {
            const cbm = this.calculateFoamCbm(values.length, values.width, values.thickness);
            const totalCbm = cbm * values.quantity;
            const twdCost = Number((values.unitPrice * cbm).toFixed(2));

            // < 8 CBM計算 (海快運費)
            const moq_temp = Math.round(5000 / twdCost);
            const totalCost1 = moq_temp * twdCost + cbm * values.seaExpressFreight * moq_temp;
            const quote1 = twdCost * 2.5 * moq_temp + cbm * values.seaExpressFreight * moq_temp;
            const unitSalePrice1 = Number((quote1 / moq_temp).toFixed(1));
            const moq1 = Math.ceil(5000 / unitSalePrice1);

            // ≥ 8 CBM計算 (貨櫃運費)
            const moq2 = Math.round(8 / cbm);
            const totalCost2 = moq2 * twdCost;
            const quote2 = twdCost * 1.85 * moq2 + cbm * values.containerFreight * moq2;
            const unitSalePrice2 = Number((quote2 / moq2).toFixed(1));

            // 差額計算
            const quantityDiff = moq2 - moq_temp;
            const unitPriceDiff = Number(unitSalePrice2 - unitSalePrice1).toFixed(2);
            const totalDiff = quantityDiff !== 0 ? unitPriceDiff / quantityDiff : 0;
            
            // 判斷使用哪種運費方式
            const isUsingContainer = values.quantity >= (moq2 - 1); // 優惠量-1就使用貨櫃
            
            // 計算單片報價和運費
            let unitPrice, unitFreight, finalPrice;
            
            if (isUsingContainer) {
                // 使用貨櫃運費
                unitPrice = Math.round(unitSalePrice2 + totalDiff * (values.quantity - moq_temp));
                unitFreight = Math.round(values.containerFreight * cbm);
            } else {
                // 使用海快運費
                unitPrice = Math.round(unitSalePrice1 + totalDiff * (values.quantity - moq_temp));
                unitFreight = Math.round(values.seaExpressFreight * cbm);
            }
            
            finalPrice = unitPrice + unitFreight; // 兩個整數相加，結果也是整數

            this.updateResults({
                values,
                cbm,
                totalCbm,
                twdCost,
                moq_temp,
                moq1,
                totalCost1,
                quote1,
                unitSalePrice1,
                moq2,
                totalCost2,
                quote2,
                unitSalePrice2,
                quantityDiff,
                unitPriceDiff,
                totalDiff,
                unitPrice,
                unitFreight,
                finalPrice,
                isUsingContainer
            });

            this.elements.resultSection.classList.add('show');
            console.log('EVA泡棉計算完成');
        } catch (error) {
            console.error('泡棉計算錯誤:', error);
            this.showError('計算過程中發生錯誤，請檢查輸入值');
        }
    }

    // 顯示片數範圍錯誤
    showQuantityRangeError(minQuantity, maxQuantity, currentQuantity) {
        // 更新海快運費頁籤內容
        document.getElementById('sea-final-quote').textContent = '請檢查輸入片數';
        document.getElementById('sea-unit-quote').textContent = '';
        document.getElementById('sea-unit-freight').textContent = '';
        document.getElementById('sea-quote-formula').textContent = 
            `片數範圍應為 ${minQuantity} ~ ${maxQuantity} 片，您輸入了 ${currentQuantity} 片`;

        // 更新貨櫃運費頁籤內容
        document.getElementById('container-final-quote').textContent = '請檢查輸入片數';
        document.getElementById('container-unit-quote').textContent = '';
        document.getElementById('container-unit-freight').textContent = '';
        document.getElementById('container-quote-formula').textContent = 
            `片數範圍應為 ${minQuantity} ~ ${maxQuantity} 片，您輸入了 ${currentQuantity} 片`;

        // 更新提示文字
        document.getElementById('tip-text').textContent = '請修正片數後重新計算';
        
        // 預設顯示海快運費頁籤
        this.switchTab('sea');
    }

    // 計算泡棉CBM
    calculateFoamCbm(length, width, thickness) {
        return (length * width * thickness) / 1000000000;
    }

    // 更新結果顯示 - 大幅修改
    updateResults({ values, cbm, twdCost, moq_temp, moq1, totalCost1, quote1, unitSalePrice1,
                    moq2, totalCost2, quote2, unitSalePrice2, quantityDiff, unitPriceDiff, 
                    totalDiff, unitPrice, unitFreight, finalPrice, isUsingContainer }) {
       
        const totalCbm = cbm * values.quantity;
        
        // 緊湊版輸入數據顯示
        document.getElementById('display-length').textContent = `${values.length} mm`;
        document.getElementById('display-width').textContent = `${values.width} mm`;
        document.getElementById('display-thickness').textContent = `${values.thickness} mm`;
        document.getElementById('display-exchange-rate').textContent = values.exchangeRate.toFixed(2);
        document.getElementById('display-unit-price').textContent = `NT$ ${values.unitPrice.toLocaleString()}`;
        document.getElementById('display-quantity').textContent = values.quantity.toLocaleString();
        document.getElementById('display-twd-cost').textContent = `NT$ ${twdCost.toFixed(1)}`;
        document.getElementById('display-cbm').textContent = cbm.toFixed(7);
        document.getElementById('display-total-cbm').textContent = totalCbm.toFixed(7);

        // 計算實際使用的單片報價（根據實際片數調整）
        const actualUnitPrice1 = Math.round(unitSalePrice1 + totalDiff * (values.quantity - moq_temp));
        const actualUnitPrice2 = Math.round(unitSalePrice2 + totalDiff * (values.quantity - moq_temp));

        // 8 CBM 內結果顯示
        document.getElementById('moq1').textContent = `${moq1}`;
        document.getElementById('unit-sale-price1').textContent = `NT$ ${actualUnitPrice1}`;
        document.getElementById('unit-freight1').textContent = `NT$ ${Math.round(values.seaExpressFreight * cbm)}`;

        // 8 CBM 以上結果顯示
        document.getElementById('moq2').textContent = `${moq2}`;
        document.getElementById('unit-sale-price2').textContent = `NT$ ${actualUnitPrice2}`;
        document.getElementById('unit-freight2').textContent = `NT$ ${Math.round(values.containerFreight * cbm)}`;

        // 根據是否使用貨櫃運費來決定醒目顯示
        const cbmUnder8Block = document.getElementById('cbm-under-8');
        const cbmOver8Block = document.getElementById('cbm-over-8');
        
        // 計算兩種運費方案的最終報價
        const seaUnitPrice = Math.round(unitSalePrice1 + totalDiff * (values.quantity - moq_temp));
        const seaUnitFreight = Math.round(values.seaExpressFreight * cbm);
        const seaFinalPrice = seaUnitPrice + seaUnitFreight;

        const containerUnitPrice = Math.round(unitSalePrice2 + totalDiff * (values.quantity - moq_temp));
        const containerUnitFreight = Math.round(values.containerFreight * cbm);
        const containerFinalPrice = containerUnitPrice + containerUnitFreight;

        // 更新海快運費頁籤內容
        document.getElementById('sea-final-quote').textContent = `NT$ ${seaFinalPrice}`;
        document.getElementById('sea-unit-quote').textContent = `NT$ ${seaUnitPrice}`;
        document.getElementById('sea-unit-freight').textContent = `NT$ ${seaUnitFreight}`;
        document.getElementById('sea-quote-formula').textContent = 
            `海快運費 = ${seaUnitPrice} + ${seaUnitFreight} = ${seaFinalPrice}`;

        // 更新貨櫃運費頁籤內容
        document.getElementById('container-final-quote').textContent = `NT$ ${containerFinalPrice}`;
        document.getElementById('container-unit-quote').textContent = `NT$ ${containerUnitPrice}`;
        document.getElementById('container-unit-freight').textContent = `NT$ ${containerUnitFreight}`;
        document.getElementById('container-quote-formula').textContent = 
            `貨櫃運費 = ${containerUnitPrice} + ${containerUnitFreight} = ${containerFinalPrice}`;

        // 設置預設頁籤（根據系統推薦，但不顯示推薦標記）
        if (isUsingContainer) {
            // 預設顯示貨櫃運費頁籤
            this.switchTab('container');
            document.getElementById('tip-text').textContent = '根據您的片數建議查看貨櫃運費';
        } else {
            // 預設顯示海快運費頁籤
            this.switchTab('sea');
            document.getElementById('tip-text').textContent = '根據您的片數建議查看海快運費';
        }
    }

    // 清除結果
    clearResults() {
        console.log('清除結果顯示');
        
        // 清除緊湊版輸入顯示
        document.getElementById('display-length').textContent = '0 mm';
        document.getElementById('display-width').textContent = '0 mm';
        document.getElementById('display-thickness').textContent = '0 mm';
        document.getElementById('display-exchange-rate').textContent = '0';
        document.getElementById('display-unit-price').textContent = '0';
        document.getElementById('display-quantity').textContent = '0';
        document.getElementById('display-twd-cost').textContent = 'NT$ 0';
        document.getElementById('display-cbm').textContent = '0.0000000';
        document.getElementById('display-total-cbm').textContent = '0.0000000';
        
        // 清除CBM結果
        document.getElementById('moq1').textContent = '0';
        document.getElementById('unit-sale-price1').textContent = 'NT$ 0';
        document.getElementById('unit-freight1').textContent = 'NT$ 0';
        document.getElementById('moq2').textContent = '0';
        document.getElementById('unit-sale-price2').textContent = 'NT$ 0';
        document.getElementById('unit-freight2').textContent = 'NT$ 0';
        
        // 清除最終報價頁籤內容
        document.getElementById('sea-final-quote').textContent = 'NT$ 0';
        document.getElementById('sea-unit-quote').textContent = 'NT$ 0';
        document.getElementById('sea-unit-freight').textContent = 'NT$ 0';
        document.getElementById('sea-quote-formula').textContent = '海快運費 = 單片報價 + 單片運費';
        
        document.getElementById('container-final-quote').textContent = 'NT$ 0';
        document.getElementById('container-unit-quote').textContent = 'NT$ 0';
        document.getElementById('container-unit-freight').textContent = 'NT$ 0';
        document.getElementById('container-quote-formula').textContent = '貨櫃運費 = 單片報價 + 單片運費';
        
        document.getElementById('tip-text').textContent = '點擊上方頁籤切換查看不同運費方案';
        
        // 重置頁籤狀態，預設顯示海快運費
        this.switchTab('sea');
        
        // 移除醒目樣式
        document.getElementById('cbm-under-8').classList.remove('highlight');
        document.getElementById('cbm-over-8').classList.remove('highlight');
    }

    // 重置所有欄位 - 修改預設值
    reset() {
        console.log('重置所有欄位');
        Object.values(this.elements).forEach(element => {
            if (element && element.tagName === 'INPUT') {
                if (element.id === 'unit-price') {
                    element.value = '23235'; // 修改預設值
                } else if (element.id === 'container-freight') {
                    element.value = '2000';
                } else if (element.id === 'sea-express-freight') {
                    element.value = '6000';
                } else {
                    element.value = '';
                }
            }
        });
        
        // 重置片數標籤和屬性
        const quantityLabel = document.getElementById('quantity-label');
        quantityLabel.textContent = '片數';
        this.elements.quantity.removeAttribute('min');
        this.elements.quantity.removeAttribute('max');
        this.elements.quantity.placeholder = '請輸入片數';
        
        this.clearResults();
        this.elements.resultSection.classList.remove('show');
        
        // 聚焦到第一個輸入欄位
        if (this.elements.containerFreight) this.elements.containerFreight.focus();
        
        this.showMessage('所有欄位已重置', 'success');
    }

    // 下載Excel報表 - 更新數據結構
    downloadExcel() {
        console.log('開始下載Excel...');
        const values = this.getInputValues();
        
        if (!this.validateInputs(values)) {
            this.showError('請先完成所有欄位的計算後再下載報表');
            return;
        }

        // 從顯示元素讀取已計算的數值
        const displayedData = {
            length: document.getElementById('display-length').textContent.replace(' mm', ''),
            width: document.getElementById('display-width').textContent.replace(' mm', ''),
            thickness: document.getElementById('display-thickness').textContent.replace(' mm', ''),
            exchangeRate: document.getElementById('display-exchange-rate').textContent,
            unitPrice: document.getElementById('display-unit-price').textContent.replace('NT$ ', '').replace(/,/g, ''),
            quantity: document.getElementById('display-quantity').textContent.replace(/,/g, ''),
            twdCost: document.getElementById('display-twd-cost').textContent.replace('NT$ ', ''),
            cbm: document.getElementById('display-cbm').textContent,
            totalCbm: document.getElementById('display-total-cbm').textContent,
            
            // CBM結果
            moq1: document.getElementById('moq1').textContent,
            unitSalePrice1: document.getElementById('unit-sale-price1').textContent.replace('NT$ ', ''),
            unitFreight1: document.getElementById('unit-freight1').textContent.replace('NT$ ', ''),
            
            moq2: document.getElementById('moq2').textContent,
            unitSalePrice2: document.getElementById('unit-sale-price2').textContent.replace('NT$ ', ''),
            unitFreight2: document.getElementById('unit-freight2').textContent.replace('NT$ ', ''),
            
            // 最終報價
            finalQuote: document.getElementById('final-quote').textContent.replace('NT$ ', ''),
            unitQuote: document.getElementById('display-unit-quote').textContent.replace('NT$ ', ''),
            unitFreightFinal: document.getElementById('display-unit-freight-final').textContent.replace('NT$ ', ''),
            quoteFormula: document.getElementById('quote-formula').textContent
        };

        this.loadSheetJS().then(() => {
            this.createExcelFile(displayedData);
        }).catch(error => {
            console.error('載入Excel庫失敗:', error);
            this.showError('下載功能初始化失敗，請重試');
        });
    }

    // 載入SheetJS庫
    loadSheetJS() {
        return new Promise((resolve, reject) => {
            if (window.XLSX) {
                resolve();
                return;
            }

            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
            script.onload = () => resolve();
            script.onerror = () => reject(new Error('Failed to load SheetJS'));
            document.head.appendChild(script);
        });
    }

    // 創建Excel文件 - 更新數據結構
    createExcelFile(displayedData) {
        const wb = window.XLSX.utils.book_new();
        const wsData = [
            ['', 'EVA泡棉報價計算結果', '', '', ''],
            ['', '', '', '', ''],
            ['', '輸入數據', '', '數值', ''],
            ['', '長度 (mm)', '', displayedData.length, ''],
            ['', '寬度 (mm)', '', displayedData.width, ''],
            ['', '厚度 (mm)', '', displayedData.thickness, ''],
            ['', '匯率(人民幣CNY)', '', displayedData.exchangeRate, ''],
            ['', '單價 (NT$/CBM)', '', displayedData.unitPrice, ''],
            ['', '片數', '', displayedData.quantity, ''],
            ['', '台幣成本', '', displayedData.twdCost, ''],
            ['', 'CBM', '', displayedData.cbm, ''],
            ['', '總CBM', '', displayedData.totalCbm, ''],
            ['', '', '', '', ''],
            ['', '8 CBM 內結果', '', '', ''],
            ['', 'MOQ1', '', displayedData.moq1, ''],
            ['', '單片報價', '', displayedData.unitSalePrice1, ''],
            ['', '單片運費', '', displayedData.unitFreight1, ''],
            ['', '', '', '', ''],
            ['', '8 CBM 以上結果', '', '', ''],
            ['', 'MOQ2', '', displayedData.moq2, ''],
            ['', '單片報價', '', displayedData.unitSalePrice2, ''],
            ['', '單片運費', '', displayedData.unitFreight2, ''],
            ['', '', '', '', ''],
            ['', '最終報價', '', displayedData.finalQuote, ''],
            ['', '單片報價', '', displayedData.unitQuote, ''],
            ['', '單片運費', '', displayedData.unitFreightFinal, ''],
            ['', '計算公式', '', displayedData.quoteFormula, '']
        ];
        const ws = window.XLSX.utils.aoa_to_sheet(wsData);
        ws['!cols'] = Array(5).fill({ wch: 20 });
        window.XLSX.utils.book_append_sheet(wb, ws, '報價結果');
        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
        const filename = `EVA泡棉報價_${timestamp}.xlsx`;
        try {
            window.XLSX.writeFileXLSX(wb, filename);
        } catch (error) {
            window.XLSX.writeFile(wb, filename);
        }
        this.showMessage('Excel報表下載成功！', 'success');
    }

    // 顯示訊息
    showMessage(message, type = 'info') {
        const messageDiv = document.createElement('div');
        messageDiv.className = `message message-${type}`;
        messageDiv.textContent = message;
        messageDiv.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 15px 20px;
            border-radius: 10px;
            color: white;
            font-weight: 600;
            z-index: 1000;
            animation: slideIn 0.3s ease-out;
            ${this.getMessageStyles(type)}
        `;
        
        document.body.appendChild(messageDiv);
        
        setTimeout(() => {
            messageDiv.style.animation = 'slideOut 0.3s ease-in';
            setTimeout(() => {
                if (document.body.contains(messageDiv)) {
                    document.body.removeChild(messageDiv);
                }
            }, 300);
        }, 3000);
    }

    // 獲取訊息樣式
    getMessageStyles(type) {
        const styles = {
            success: 'background: #1565c0;',
            error: 'background: #d32f2f;',
            info: 'background: #2196f3;'
        };
        return styles[type] || styles.info;
    }

    // 顯示錯誤訊息
    showError(message) {
        this.showMessage(message, 'error');
    }
}

// 當DOM載入完成後初始化計算器
document.addEventListener('DOMContentLoaded', () => {
    console.log('DOM載入完成，開始初始化EVA泡棉計算器...');
    try {
        window.foamCalculator = new FoamCalculator();
        console.log('EVA泡棉報價計算器已啟動');
        console.log('快捷鍵：Ctrl+Enter 計算，Ctrl+R 重置');
    } catch (error) {
        console.error('初始化計算器時發生錯誤:', error);
    }
});
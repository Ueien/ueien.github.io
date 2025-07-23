// 周轉箱報價計算器
class TurnoverCalculator {
    constructor() {
        this.initializeElements();
        this.bindEvents();
        console.log('周轉箱計算器已啟動');
    }

    initializeElements() {
        this.elements = {
            // 輸入元素
            width: document.getElementById('width'),
            depth: document.getElementById('depth'),
            height: document.getElementById('height'),
            rmbCost: document.getElementById('rmb-cost'),
            exchangeRate: document.getElementById('exchange-rate'),
            quantity: document.getElementById('quantity'),
            cbmPrice: document.getElementById('cbm-price'),

            // 按鈕
            calculateBtn: document.getElementById('calculate-btn'),
            resetBtn: document.getElementById('reset-btn'),
            downloadBtn: document.getElementById('download-btn'),
            resultSection: document.getElementById('result-section')
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

    bindInputEvents() {
        const inputs = [
            this.elements.width,
            this.elements.depth,
            this.elements.height,
            this.elements.rmbCost,
            this.elements.exchangeRate,
            this.elements.quantity,
            this.elements.cbmPrice
        ].filter(el => el !== null);

        inputs.forEach(element => {
            element.addEventListener('input', () => this.calculate());
        });
    }

    // 獲取輸入值
    getInputValues() {
        return {
            width: parseFloat(this.elements.width.value) || 0,
            depth: parseFloat(this.elements.depth.value) || 0,
            height: parseFloat(this.elements.height.value) || 0,
            rmbCost: parseFloat(this.elements.rmbCost.value) || 0,
            exchangeRate: parseFloat(this.elements.exchangeRate.value) || 0,
            quantity: parseInt(this.elements.quantity.value) || 0,
            cbmPrice: parseFloat(this.elements.cbmPrice.value) || 2000
        };
    }

    // 驗證輸入值
    validateInputs(values) {
        const required = ['width', 'depth', 'height', 'rmbCost', 'exchangeRate', 'quantity'];
        return required.every(key => values[key] > 0);
    }

    // 計算功能
    calculate() {
        console.log('開始計算周轉箱報價');
        const values = this.getInputValues();
        
        if (!this.validateInputs(values)) {
            this.clearResults();
            this.elements.resultSection.classList.remove('show');
            return;
        }

        try {
            // 更新輸入數據顯示
            this.updateInputDisplay(values);
            
            // 計算台幣成本
            const twdCost = this.calculateTwdCost(values.rmbCost, values.exchangeRate);
            
            // 計算CBM
            const singleCbm = this.calculateSingleCbm(values.width, values.depth, values.height);
            const totalCbm = this.calculateTotalCbm(singleCbm, values.quantity);
            
            // 周轉箱計算
            const pricingTier = this.getTurnoverPricingTier(totalCbm, singleCbm);
            const shippingCost = this.calculateTurnoverShipping(singleCbm, values.cbmPrice);
            const unitPrice = Math.round(twdCost * pricingTier.multiplier + shippingCost);
            const totalPrice = unitPrice * values.quantity;
            
            // 更新顯示
            this.updateResults({
                values,
                twdCost,
                singleCbm,
                totalCbm,
                pricingTier,
                shippingCost,
                unitPrice,
                totalPrice
            });
            
            // 顯示結果區域
            this.elements.resultSection.classList.add('show');
            
            console.log('周轉箱計算完成');
            
        } catch (error) {
            console.error('計算錯誤:', error);
            this.showError('計算過程中發生錯誤，請檢查輸入值');
        }
    }

    // 計算台幣成本
    calculateTwdCost(rmbCost, exchangeRate) {
        return rmbCost * exchangeRate;
    }

    // 計算單個CBM
    calculateSingleCbm(width, depth, height) {
        return (width / 1000) * (depth / 1000) * (height / 1000);
    }

    // 計算總CBM
    calculateTotalCbm(singleCbm, quantity) {
        return singleCbm * quantity;
    }

    // 周轉箱定價層級
    getTurnoverPricingTier(totalCbm, singleCbm) {
        let multiplier, cbmRange, quantityRange;
        
        if (totalCbm < 1) {
            multiplier = 2.35;
            cbmRange = '< 1 CBM';
            const maxQuantity = Math.floor(1 / singleCbm);
            quantityRange = `< ${maxQuantity}個`;
        } else if (totalCbm >= 1 && totalCbm < 3) {
            multiplier = 2.2;
            cbmRange = '1-3 CBM';
            const minQuantity = Math.floor(1 / singleCbm);
            const maxQuantity = Math.floor(3 / singleCbm);
            quantityRange = `${minQuantity}~${maxQuantity}個`;
        } else if (totalCbm >= 3 && totalCbm < 5) {
            multiplier = 2.05;
            cbmRange = '3-5 CBM';
            const minQuantity = Math.floor(3 / singleCbm);
            const maxQuantity = Math.floor(5 / singleCbm);
            quantityRange = `${minQuantity}~${maxQuantity}個`;
        } else if (totalCbm >= 5 && totalCbm < 8) {
            multiplier = 1.9;
            cbmRange = '5-8 CBM';
            const minQuantity = Math.floor(5 / singleCbm);
            const maxQuantity = Math.floor(8 / singleCbm);
            quantityRange = `${minQuantity}~${maxQuantity}個`;
        } else {
            multiplier = 1.85;
            cbmRange = '8-24 CBM';
            const minQuantity = Math.floor(8 / singleCbm);
            const maxQuantity = Math.floor(24 / singleCbm);
            quantityRange = `${minQuantity}~${maxQuantity}個`;
        }
        
        return { 
            multiplier, 
            tier: cbmRange,
            quantityRange: quantityRange
        };
    }

    // 計算周轉箱運費
    calculateTurnoverShipping(singleCbm, cbmPrice) {
        return singleCbm * cbmPrice;
    }

    // 更新輸入顯示
    updateInputDisplay(values) {
        document.getElementById('display-width').textContent = `${values.width} mm`;
        document.getElementById('display-depth').textContent = `${values.depth} mm`;
        document.getElementById('display-height').textContent = `${values.height} mm`;
        document.getElementById('display-cost').textContent = `${values.rmbCost} CNY`;
        document.getElementById('display-rate').textContent = values.exchangeRate;
        document.getElementById('display-quantity').textContent = `${values.quantity} 個`;
        document.getElementById('display-cbm-price').textContent = `${values.cbmPrice} 元/CBM`;
    }

    // 更新結果顯示
    updateResults({ values, twdCost, singleCbm, totalCbm, pricingTier, shippingCost, unitPrice, totalPrice }) {
        console.log('開始更新顯示結果...');
        
        document.getElementById('twd-cost').textContent = `NT$ ${Math.round(twdCost).toLocaleString()}`;
        document.getElementById('single-cbm').textContent = `${singleCbm.toFixed(3)} CBM`;
        document.getElementById('total-cbm').textContent = `${totalCbm.toFixed(2)} CBM`;
        document.getElementById('freight-cost').textContent = `NT$ ${Math.round(shippingCost)}`;
        document.getElementById('unit-price').textContent = `NT$ ${Math.round(unitPrice).toLocaleString()}`;
        document.getElementById('total-price').textContent = `NT$ ${Math.round(totalPrice).toLocaleString()}`;
        document.getElementById('quote-tier').textContent = `數量區間：${pricingTier.quantityRange}`;
        document.getElementById('quote-formula').textContent = 
            `計算公式：台幣成本 × ${pricingTier.multiplier} + 每個運費(${Math.round(shippingCost)}) = ${Math.round(unitPrice)}`;
        document.getElementById('final-quote').textContent = `NT$ ${Math.round(totalPrice).toLocaleString()}`;
        
        console.log('顯示結果更新完成');
    }

    // 清除結果
    clearResults() {
        console.log('清除結果顯示');
        document.getElementById('twd-cost').textContent = 'NT$ 0';
        document.getElementById('single-cbm').textContent = '0.000 CBM';
        document.getElementById('total-cbm').textContent = '0.00 CBM';
        document.getElementById('freight-cost').textContent = 'NT$ 0';
        document.getElementById('unit-price').textContent = 'NT$ 0';
        document.getElementById('total-price').textContent = 'NT$ 0';
        document.getElementById('quote-tier').textContent = '數量區間：待計算';
        document.getElementById('final-quote').textContent = 'NT$ 0';
        document.getElementById('quote-formula').textContent = '計算公式：待計算';
        
        // 清除輸入顯示
        document.getElementById('display-width').textContent = '0 mm';
        document.getElementById('display-depth').textContent = '0 mm';
        document.getElementById('display-height').textContent = '0 mm';
        document.getElementById('display-cost').textContent = '0 CNY';
        document.getElementById('display-rate').textContent = '0';
        document.getElementById('display-quantity').textContent = '0 個';
        document.getElementById('display-cbm-price').textContent = '0 元/CBM';
    }

    // 重置所有欄位
    reset() {
        console.log('重置所有欄位');
        Object.values(this.elements).forEach(element => {
            if (element && element.tagName === 'INPUT') {
                if (element.id === 'cbm-price') {
                    element.value = '2000';
                } else {
                    element.value = '';
                }
            }
        });
        
        this.clearResults();
        this.elements.resultSection.classList.remove('show');
        
        // 聚焦到第一個輸入欄位
        if (this.elements.width) this.elements.width.focus();
        
        this.showMessage('所有欄位已重置', 'success');
    }

    // 下載Excel報表
    downloadExcel() {
        console.log('開始下載Excel...');
        const values = this.getInputValues();
        
        if (!this.validateInputs(values)) {
            this.showError('請先完成所有欄位的計算後再下載報表');
            return;
        }

        // 從顯示元素讀取已計算的數值
        const displayedData = {
            width: document.getElementById('display-width').textContent.replace(' mm', ''),
            depth: document.getElementById('display-depth').textContent.replace(' mm', ''),
            height: document.getElementById('display-height').textContent.replace(' mm', ''),
            cost: document.getElementById('display-cost').textContent.replace(' CNY', ''),
            exchangeRate: document.getElementById('display-rate').textContent,
            quantity: document.getElementById('display-quantity').textContent.replace(' 個', ''),
            cbmPrice: document.getElementById('display-cbm-price').textContent.replace(' 元/CBM', ''),
            twdCost: document.getElementById('twd-cost').textContent.replace('NT$ ', '').replace(/,/g, ''),
            singleCbm: document.getElementById('single-cbm').textContent.replace(' CBM', ''),
            totalCbm: document.getElementById('total-cbm').textContent.replace(' CBM', ''),
            freightCost: document.getElementById('freight-cost').textContent.replace('NT$ ', '').replace(/,/g, ''),
            unitPrice: document.getElementById('unit-price').textContent.replace('NT$ ', '').replace(/,/g, ''),
            totalPrice: document.getElementById('total-price').textContent.replace('NT$ ', '').replace(/,/g, ''),
            quoteTier: document.getElementById('quote-tier').textContent,
            finalQuote: document.getElementById('final-quote').textContent.replace('NT$ ', '').replace(/,/g, ''),
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

    // 創建Excel文件
    createExcelFile(displayedData) {
        const wb = window.XLSX.utils.book_new();
        
        const data = [
            ['', '周轉箱及棧板報價計算結果', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '輸入數據', '', '數值', '', '', '計算結果', '', '數值', '', '', ''],
            ['', '寬度 (mm)', '', displayedData.width, '', '', 'CBM 單個', '', displayedData.singleCbm, '', '', ''],
            ['', '深度 (mm)', '', displayedData.depth, '', '', 'CBM 總計', '', displayedData.totalCbm, '', '', ''],
            ['', '高度 (mm)', '', displayedData.height, '', '', '台幣成本', '', displayedData.twdCost, '', '', ''],
            ['', '成本 (CNY)', '', displayedData.cost, '', '', '運費', '', displayedData.freightCost, '', '', ''],
            ['', '匯率', '', displayedData.exchangeRate, '', '', '單個報價', '', displayedData.unitPrice, '', '', ''],
            ['', '數量', '', displayedData.quantity, '', '', '總報價', '', displayedData.totalPrice, '', '', ''],
            ['', 'CBM價格', '', displayedData.cbmPrice, '', '', '數量區間', '', displayedData.quoteTier.replace('數量區間：', ''), '', '', '']
        ];

        const ws = window.XLSX.utils.aoa_to_sheet(data);
        ws['!cols'] = Array(12).fill({ wch: 15 });
        window.XLSX.utils.book_append_sheet(wb, ws, '報價結果');
        
        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
        const filename = `周轉箱報價_${timestamp}.xlsx`;
        
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
            success: 'background: #2e7d32;',
            error: 'background: #d32f2f;',
            info: 'background: #388e3c;'
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
    console.log('DOM載入完成，開始初始化周轉箱計算器...');
    try {
        window.turnoverCalculator = new TurnoverCalculator();
        console.log('周轉箱報價計算器已啟動');
        console.log('快捷鍵：Ctrl+Enter 計算，Ctrl+R 重置');
    } catch (error) {
        console.error('初始化計算器時發生錯誤:', error);
    }
});
/**
 * ============================================================================
 * 🐟 龍潭進銷存管理系統 - 最終完整修復版 (Part 1)
 * ============================================================================
 */

// ==========================================
// 1. 📋 Config (設定檔)
// ==========================================
var CONFIG = {
  SHEETS: {
    CALENDAR: "行事曆",
    PRODUCTS: "商品主檔",
    MATERIALS: "原料主檔", 
    CUSTOMERS: "客戶資料",
    SUPPLIERS: "供應商資料",
    SPECIAL_PRICE: "客戶特價表",
    PURCHASE: "進貨單",
    PURCHASE_DETAILS: "進貨明細",
    PAYABLE: "應付帳款",
    SALES: "銷貨單",
    SALES_DETAILS: "銷貨明細",
    RECEIVABLE: "應收帳款",
    INVENTORY: "庫存",
    INVENTORY_LOG: "庫存異動",
    STOCKTAKE: "盤點記錄",
    EXPENSES: "現金支出",
    PETTY_CASH: "零用金紀錄",
    PAYMENT_RECEIVED: "收款記錄",
    PAYMENT_MADE: "付款記錄",
    MONTHLY_REPORT: "月結報表",
    INVOICES_ISSUED: "銷項發票紀錄",
    INVOICES_RECEIVED: "進項發票紀錄",
    SUPPLIER_PRICES: "供應商進價記錄",
    EMPLOYEE_MEALS: "員工餐費"
  }
};



// ==========================================
// 2. 🛠️ Utils (工具函式)
// ==========================================

function getNowString() {
  return Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
}

function formatDate(date) {
  if (!date) return "";
  if (typeof date === 'string') return date.substring(0, 10);
  try {
    return Utilities.formatDate(new Date(date), "GMT+8", "yyyy-MM-dd");
  } catch (e) {
    return String(date);
  }
}

/**
 * 產生新的 ID (自動搜尋目前最大的編號並 +1)
 * 解決舊單號被移動到最後一行導致新單號重複的問題
 */
function generateId(prefix, sheetName, colIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  // 防呆：找不到工作表時，回傳日期編號
  if (!sheet) {
    return prefix + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMddHHmmss");
  }

  var lastRow = sheet.getLastRow();

  // 情況 1：如果是第一筆資料 (只有標題列)，從 00001 開始
  if (lastRow < 2) {
    return prefix + "00001";
  }

  // ==========================================
  // 🔥 修改核心：讀取整欄資料，找出最大值
  // ==========================================
  
  // 1. 一次讀取該欄位的所有 ID (從第2列讀到最後一列，只讀取 colIndex 這一欄)
  // getRange(起始列, 起始欄, 讀幾列, 讀幾欄)
  var data = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  
  var maxNum = 0; // 用來記錄目前找到的最大數字

  // 2. 跑迴圈檢查每一筆 ID
  for (var i = 0; i < data.length; i++) {
    var currentId = String(data[i][0]);
    
    // 簡單過濾：確保 ID 不為空，且包含我們指定的 prefix (例如 SO)
    // 這樣可以避免讀到奇怪的備註文字或日期格式導致報錯
    if (currentId && currentId.indexOf(prefix) !== -1) {
      
      // 取出數字部分 (移除 prefix 和非數字字元)
      // 例如 SO00056 -> 56
      var numPart = currentId.replace(/[^0-9]/g, '');
      var currentNum = parseInt(numPart, 10);

      // 如果是有效數字，且比目前記錄的最大值還大，就更新最大值
      if (!isNaN(currentNum) && currentNum > maxNum) {
        maxNum = currentNum;
      }
    }
  }

  // 3. 新的號碼 = 最大值 + 1
  // 即使最後一行是 SO00053，但因為中間有讀到 SO00056，maxNum 會是 56
  // 所以 nextNum 會變成 57，不會重複！
  var nextNum = maxNum + 1;

  // ==========================================

  // 補零 (補滿 5 位數)
  var paddedNum = ("00000" + nextNum).slice(-5);

  return prefix + paddedNum;
}

// 刪除列輔助函數
function deleteRowsById(sheet, colIndex, id) {
  var data = sheet.getDataRange().getValues();
  // 從後面往前刪
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][colIndex - 1]) === String(id)) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ==========================================
// 3. 🏠 Main Menu (主選單)
// ==========================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('❤️ 央廚進銷存管理系統')
    .addSubMenu(createDailyMenu(ui))
    .addSubMenu(createPaymentMenu(ui))
    .addSubMenu(createInventoryMenu(ui))
    .addSubMenu(createFinanceMenu(ui))
    .addSubMenu(createReportMenu(ui))
    .addSeparator()
    .addSubMenu(createSettingsMenu(ui))
    .addSeparator()
    .addItem('🔄 重新整理', 'refreshAll')
    .addToUi();
}

function createDailyMenu(ui) {
  return ui.createMenu('📝 開單作業')
    .addItem('📤 開新銷貨單', 'showSalesOrderPanel')
    .addItem('📥 開新進貨單', 'showPurchaseOrderPanel')
    .addSeparator()
    .addItem('🔍 查詢銷貨單', 'showSearchSalesPanel')
    .addItem('🔍 查詢進貨單', 'showSearchPurchasePanel');
}

function createPaymentMenu(ui) {
  return ui.createMenu('💰 收付款管理')
    .addItem('💰 客戶收款', 'showReceivePaymentPanel')
    .addItem('💸 供應商付款', 'showMakePaymentPanel')
    .addSeparator()
    .addItem('📊 應收帳款報表', 'showReceivableReport')
    .addItem('📊 應付帳款報表', 'showPayableReport')
    .addSeparator()
    .addItem('🇨🇳 大陸貨款管理', 'showChinaPaymentPanel');
}

function createInventoryMenu(ui) {
  return ui.createMenu('📦 庫存管理')
    .addItem('📋 查看庫存', 'showInventoryPanel')
    .addItem('⚠️ 庫存預警', 'showLowStockAlert')
    .addSeparator()
    .addItem('✅ 庫存盤點', 'showStocktakePanel')
    .addItem('📊 異動記錄', 'showInventoryLogPanel')
    .addItem('🔄 商品分裝轉換', 'showConvertPanel');
}

function createFinanceMenu(ui) {
  return ui.createMenu('💵 財務管理')
    .addItem('💵 記錄現金支出', 'showExpensePanel')
    .addItem('💰 零用金管理', 'showPettyCashPanel')
    .addItem('🍱 記錄員工餐費', 'showEmployeeMealPanel')
    .addItem('📋 支出明細', 'showExpenseList')
    .addItem('📊 收支統計', 'showCashFlowReport')
    .addSeparator()
    .addItem('🧾 紀錄銷項發票', 'showInvoiceIssuedForm')
    .addItem('🧾 紀錄進項發票', 'showInvoiceReceivedForm');
}

function createReportMenu(ui) {
  return ui.createMenu('📈 報表分析')
    .addItem('📅 行事曆', 'showCalendarPanel')
    .addItem('📅 詳細月結總表', 'showDetailedMonthlyReportPanel')
    .addItem('📅 簡易月結報表', 'showMonthlyReportPanel')
    .addItem('👥 客戶月結', 'showCustomerMonthlyReport')
    .addSeparator()
    .addItem('📦 商品分析', 'showProductAnalysisPanel')
    .addItem('💎 利潤分析', 'showProfitAnalysisPanel');
}

// ✅ 請用這段取代原本的 createSettingsMenu
function createSettingsMenu(ui) {
  return ui.createMenu('⚙️ 基礎設定')
    .addItem('➕ 新增商品', 'showProductForm')
    .addItem('🍱 原料管理', 'showMaterialManager')
    .addItem('👥 客戶管理', 'showCustomerManager')
    .addItem('🏢 供應商管理', 'showSupplierManager')
    .addSeparator()
    .addItem('🏷️ 設定客戶特價', 'showPriceForm')
    .addSeparator()
    .addItem('🛡️ 初始化/重設保護', 'initializeDataProtection');
}

function refreshAll() {
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('✅ 系統已重新整理完成!');
}

// ==========================================
// 4. 👥 Customer (客戶模組)
// ==========================================

function getCustomers() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CUSTOMERS);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    return data.map(function(row) {
      return {
        id: String(row[0] || ""),
        name: String(row[1] || ""),
        contact: String(row[2] || ""),
        phone: String(row[3] || ""),
        mobile: String(row[4] || ""),
        address: String(row[5] || ""),
        email: String(row[6] || ""),
        taxId: String(row[7] || ""),
        invoiceTitle: String(row[8] || ""),
        paymentTerm: String(row[9] || "現金"),
        creditLimit: row[10] || 0,
        taxType: String(row[11] || "免稅"),
        status: String(row[12] || "啟用"),
        note: String(row[14] || ""),
        priceGroup: String(row[15] || "一般").trim()
      };
    }).filter(function(c) { return c.id !== ""; });
  } catch (e) { return []; }
}

function addCustomer(data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CUSTOMERS);
    if (!sheet) throw new Error("找不到工作表：" + CONFIG.SHEETS.CUSTOMERS);
    var customerId = generateId("C", CONFIG.SHEETS.CUSTOMERS, 1);
    var nowStr = getNowString().split(' ')[0];
    var newRow = sheet.getLastRow() + 1;
    var rowData = [
      customerId, data.name, data.contact || '', data.phone ? "'" + data.phone : '',
      data.mobile ? "'" + data.mobile : '', data.address || '', data.email || '',
      data.taxId ? "'" + data.taxId : '', data.invoiceTitle || '', data.paymentTerm || '現金',
      data.creditLimit || 0, data.taxType || '免稅', '啟用', nowStr,
      data.note || '', data.priceGroup || '一般'
    ];
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    return { success: true, customerId: customerId };
  } catch (error) { return { success: false, error: error.toString() }; }
}

function searchCustomersByKeyword(keyword) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CUSTOMERS);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    var results = [];
    var k = keyword.toLowerCase();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var id = String(row[0] || "").toLowerCase();
      var name = String(row[1] || "").toLowerCase();
      var phone = String(row[3] || "").toLowerCase();
      var mobile = String(row[4] || "").toLowerCase();
      if (id.indexOf(k) > -1 || name.indexOf(k) > -1 || phone.indexOf(k) > -1 || mobile.indexOf(k) > -1) {
        results.push({ id: String(row[0]), name: String(row[1]), phone: String(row[3]), mobile: String(row[4]) });
      }
      if (results.length >= 20) break;
    }
    return results;
  } catch (e) { return []; }
}

function getCustomerById(customerId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CUSTOMERS);
    if (!sheet) return { success: false, error: "找不到客戶資料工作表" };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: "沒有客戶資料" };
    var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === customerId) {
        return {
          success: true,
          customer: {
            rowIndex: i + 2, id: String(data[i][0]), name: String(data[i][1]),
            contact: String(data[i][2]), phone: String(data[i][3]), mobile: String(data[i][4]),
            address: String(data[i][5]), email: String(data[i][6]), taxId: String(data[i][7]),
            invoiceTitle: String(data[i][8]), paymentTerm: String(data[i][9]),
            creditLimit: Number(data[i][10]) || 0, taxType: String(data[i][11]),
            status: String(data[i][12]), note: String(data[i][14]), priceGroup: String(data[i][15] || "一般")
          }
        };
      }
    }
    return { success: false, error: "找不到此客戶 ID：" + customerId };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function updateCustomer(data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CUSTOMERS);
    if (!sheet) throw new Error("找不到客戶資料工作表");
    var rowIndex = data.rowIndex;
    var currentId = sheet.getRange(rowIndex, 1).getValue();
    if (String(currentId) !== data.id) throw new Error("資料驗證失敗，請重新搜尋客戶");
    sheet.getRange(rowIndex, 2).setValue(data.name);
    sheet.getRange(rowIndex, 3).setValue(data.contact || '');
    sheet.getRange(rowIndex, 4).setValue(data.phone ? "'" + data.phone : '');
    sheet.getRange(rowIndex, 5).setValue(data.mobile ? "'" + data.mobile : '');
    sheet.getRange(rowIndex, 6).setValue(data.address || '');
    sheet.getRange(rowIndex, 7).setValue(data.email || '');
    sheet.getRange(rowIndex, 8).setValue(data.taxId ? "'" + data.taxId : '');
    sheet.getRange(rowIndex, 9).setValue(data.invoiceTitle || '');
    sheet.getRange(rowIndex, 10).setValue(data.paymentTerm || '現金');
    sheet.getRange(rowIndex, 11).setValue(data.creditLimit || 0);
    sheet.getRange(rowIndex, 12).setValue(data.taxType || '免稅');
    sheet.getRange(rowIndex, 13).setValue(data.status || '啟用');
    sheet.getRange(rowIndex, 15).setValue(data.note || '');
    sheet.getRange(rowIndex, 16).setValue(data.priceGroup || '一般');
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteCustomer(customerId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CUSTOMERS);
    if (!sheet) throw new Error("找不到客戶資料工作表");
    deleteRowsById(sheet, 1, customerId);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getCustomerFavoriteProducts(customerId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CUSTOMER_FAVORITES);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var results = [];
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][1]) === customerId) {
        results.push({
          productId: String(data[i][3]), productName: String(data[i][4]),
          sortOrder: Number(data[i][5]) || 999, purchaseCount: Number(data[i][7]) || 0
        });
      }
    }
    return results.sort(function(a, b) { return a.sortOrder - b.sortOrder; });
  } catch (e) { return []; }
}

// ==========================================
// 5. 🏢 Supplier (供應商模組) - 修改版
// ==========================================

function getSuppliers() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SUPPLIERS);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // 讀取範圍：A ~ O (共15欄)
    var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    
    var list = [];
    for (var i = 0; i < data.length; i++) {
      var id = String(data[i][0] || "").trim();
      var name = String(data[i][1] || "").trim();
      if (id !== "" && name !== "") {
        list.push({ 
          id: id, 
          name: name,
          // J欄 (Index 9) 付款方式
          paymentMethod: String(data[i][9] || "老闆付款"), 
          // ⭐ O欄 (Index 14) 稅別 - 修正為讀取第 15 欄
          taxType: String(data[i][14] || "免稅") 
        });
      }
    }
    return list;
  } catch (e) { return []; }
}

function addSupplier(data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SUPPLIERS);
    if (!sheet) throw new Error("找不到工作表：" + CONFIG.SHEETS.SUPPLIERS);
    var supplierId = generateId("S", CONFIG.SHEETS.SUPPLIERS, 1);
    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
    var newRow = sheet.getLastRow() + 1;
    
    var rowData = [
      supplierId,                        // A: 編號
      data.name,                         // B: 名稱
      data.contact || '',                // C: 聯絡人
      data.phone ? "'" + data.phone : '',// D: 電話
      data.mobile ? "'" + data.mobile : '',// E: 手機
      data.address || '',                // F: 地址
      data.email || '',                  // G: Email
      data.taxId ? "'" + data.taxId : '',// H: 統編
      data.paymentTerm || "貨到付現",      // I: 付款條件
      data.paymentMethod || "現金",        // J: 付款方式
      data.bankAccount || '',            // K: 銀行帳號
      "啟用",                             // L: 狀態
      nowStr,                            // M: 建立日期
      data.note || '',                   // N: 備註
      data.taxType || "免稅"              // O: 稅別 (第15欄)
    ];
    
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    return { success: true, supplierId: supplierId };
  } catch (error) { return { success: false, error: error.toString() }; }
}

function searchSuppliersByKeyword(keyword) {
  // 此函式原本只讀取前幾欄做搜尋，維持原樣即可，不需要改動
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SUPPLIERS);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    var results = [];
    var k = keyword.toLowerCase();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var id = String(row[0] || "").toLowerCase();
      var name = String(row[1] || "").toLowerCase();
      var phone = String(row[3] || "").toLowerCase();
      var mobile = String(row[4] || "").toLowerCase();
      if (id.indexOf(k) > -1 || name.indexOf(k) > -1 || phone.indexOf(k) > -1 || mobile.indexOf(k) > -1) {
        results.push({ id: String(row[0]), name: String(row[1]), phone: String(row[3]), mobile: String(row[4]) });
      }
      if (results.length >= 20) break;
    }
    return results;
  } catch (e) { return []; }
}

function getSupplierById(supplierId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SUPPLIERS);
    if (!sheet) return { success: false, error: "找不到供應商資料工作表" };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: "沒有供應商資料" };
    
    // ⚠️ 修改：範圍擴大到 16 欄
    var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === supplierId) {
        return {
          success: true,
          supplier: {
            rowIndex: i + 2, id: String(data[i][0]), name: String(data[i][1]),
            contact: String(data[i][2]), phone: String(data[i][3]), mobile: String(data[i][4]),
            address: String(data[i][5]), email: String(data[i][6]), taxId: String(data[i][7]),
            paymentTerm: String(data[i][8]), paymentMethod: String(data[i][9]),
            bankAccount: String(data[i][10]), status: String(data[i][11]),
            note: String(data[i][13]), supplierType: String(data[i][14] || '進貨供應商'),
            // ⚠️ 新增：回傳稅別
            taxType: String(data[i][15] || '免稅')
          }
        };
      }
    }
    return { success: false, error: "找不到此供應商 ID：" + supplierId };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function updateSupplier(data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SUPPLIERS);
    if (!sheet) throw new Error("找不到供應商資料工作表");
    var rowIndex = data.rowIndex;
    var currentId = sheet.getRange(rowIndex, 1).getValue();
    if (String(currentId) !== data.id) throw new Error("資料驗證失敗");
    
    // 更新各個欄位
    sheet.getRange(rowIndex, 2).setValue(data.name);
    sheet.getRange(rowIndex, 3).setValue(data.contact || '');
    sheet.getRange(rowIndex, 4).setValue(data.phone ? "'" + data.phone : '');
    sheet.getRange(rowIndex, 5).setValue(data.mobile ? "'" + data.mobile : '');
    sheet.getRange(rowIndex, 6).setValue(data.address || '');
    sheet.getRange(rowIndex, 7).setValue(data.email || '');
    sheet.getRange(rowIndex, 8).setValue(data.taxId ? "'" + data.taxId : '');
    sheet.getRange(rowIndex, 9).setValue(data.paymentTerm || '貨到付現');
    sheet.getRange(rowIndex, 10).setValue(data.paymentMethod || '現金');
    sheet.getRange(rowIndex, 11).setValue(data.bankAccount || '');
    sheet.getRange(rowIndex, 12).setValue(data.status || '啟用');
    sheet.getRange(rowIndex, 14).setValue(data.note || '');
    sheet.getRange(rowIndex, 15).setValue(data.taxType || '免稅');
    
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteSupplier(supplierId) {
  // 刪除邏輯不變
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SUPPLIERS);
    if (!sheet) throw new Error("找不到供應商資料工作表");
    deleteRowsById(sheet, 1, supplierId);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ==========================================
// 6. 📦 Product (商品模組)
// ==========================================

function getProducts() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PRODUCTS);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    return data.map(function(row) {
      return {
        id: String(row[0] || ""), category: String(row[1] || ""), name: String(row[2] || ""),
        spec: String(row[3] || ""), unit: String(row[4] || "個"), price: Number(row[5]) || 0,
        priceBranch: Number(row[6]) || 0, priceWholesale: Number(row[7]) || 0,
        priceGroupBuy: Number(row[8]) || 0, cost: Number(row[9]) || 0,
        supplier: String(row[10] || ""), alert: Number(row[11]) || 5,
        status: String(row[12] || "啟用"), displayName: String(row[0]) + " | " + String(row[2])
      };
    }).filter(function(p) { return p.id !== ""; });
  } catch (e) { return []; }
}

function getProductsLite() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PRODUCTS);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    var results = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var status = String(row[12] || "啟用");
      if (status.indexOf("停用") > -1) continue;
      results.push({
        id: String(row[0] || ""), category: String(row[1] || ""), name: String(row[2] || ""),
        spec: String(row[3] || ""), unit: String(row[4] || "個"), price: Number(row[5]) || 0,
        priceBranch: Number(row[6]) || 0, priceWholesale: Number(row[7]) || 0,
        priceGroupBuy: Number(row[8]) || 0, cost: Number(row[9]) || 0
      });
    }
    return results;
  } catch (e) { return []; }
}

function searchProductsByKeyword(keyword) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PRODUCTS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    var k = keyword.toString().trim();
    if (!k) return [];

    // 🔥 關鍵優化：使用 TextFinder 搜尋 A欄(ID)到 C欄(名稱)
    var finder = sheet.getRange("A2:C" + lastRow).createTextFinder(k).matchCase(false);
    var findResults = finder.findAll();
    
    var results = [];
    var rowIndices = new Set(); 

    for (var i = 0; i < findResults.length; i++) {
      var rowIdx = findResults[i].getRow();
      if (rowIndices.has(rowIdx)) continue; // 避免 ID 和名稱重複搜尋到同一列
      rowIndices.add(rowIdx);
      
      // 僅讀取該列的資料
      var row = sheet.getRange(rowIdx, 1, 1, 13).getValues()[0];
      var status = String(row[12] || "啟用");
      var displayName = String(row[2]);
      if (status.indexOf("停用") > -1) displayName = "💤 " + displayName;

      // ⭐ 關鍵修正：把主檔裡的所有價格等級都打包進去
      results.push({
        id: String(row[0]),                  // A: 編號
        category: String(row[1]),            // B: 類別
        name: displayName,                   // C: 名稱 (包含睡覺符號)
        spec: String(row[3]),                // D: 規格
        unit: String(row[4]),                // E: 單位
        price: Number(row[5]) || 0,          // F: 原價
        priceBranch: Number(row[6]) || 0,    // G: 龍潭價 (對應你的總倉價)
        priceStore: Number(row[7]) || 0,     // H: 店家價
        priceRest: Number(row[8]) || 0,      // I: 餐廳價
        cost: Number(row[9]) || 0,           // J: 成本
        isInactive: (status.indexOf("停用") > -1)
      });
      if (results.length >= 30) break; // 限制回傳數量，前端才不會卡
    }
    return results;
  } catch (e) { return []; }
}

// ✅ 這是新的，請貼在原本 saveProduct 的位置
function saveProduct(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
    if (!sheet) throw new Error("找不到工作表：" + CONFIG.SHEETS.PRODUCTS);
    
    var nowStr = getNowString().split(' ')[0];
    var newId = String(data.id || "").trim();
    
    // 1. 處理自動編號邏輯
    if (newId === "" || newId === "AUTO") {
      // 自動產生編號，使用 "P" 開頭
      newId = generateId("P", CONFIG.SHEETS.PRODUCTS, 1);
    } else {
      // 2. 如果是手動輸入，執行重複檢查防呆
      var dataValues = sheet.getDataRange().getValues();
      for (var i = 1; i < dataValues.length; i++) {
        if (String(dataValues[i][0]) === newId) {
          return { success: false, error: "⚠️ 編號 [" + newId + "] 已存在，請更換或使用自動產生。" };
        }
      }
    }

    var rowData = [
      newId,                            // A: 商品編號
      data.category || "未分類",        // B: 類別
      data.name,                        // C: 商品名稱
      data.spec || "",                  // D: 規格
      data.unit || "個",                // E: 單位
      Number(data.price) || 0,          // F: 售價
      Number(data.priceBranch) || 0,    // G: 總倉價 (✅ 更新)
      Number(data.priceStore) || 0,     // H: 店家價 (✅ 更新)
      Number(data.priceRest) || 0,      // I: 餐廳價 (✅ 更新)
      Number(data.cost) || 0,           // J: 成本
      data.supplier || "",              // K: 供應商
      Number(data.alert) || 5,          // L: 安全庫存
      "啟用",                           // M: 狀態
      nowStr,                           // N: 建立日期
      data.note || ""                   // O: 備註
    ];
    
    sheet.appendRow(rowData);
    
    // 🔥 強制刷新，避免批次新增時編號重複
    SpreadsheetApp.flush(); 
    
    return { success: true, productId: newId };
    
  } catch (error) { 
    return { success: false, error: error.toString() }; 
  }
}



// ==========================================
// ⭐ 修正版：批量儲存特價 (修復 maxSortMap 未定義錯誤)
// ==========================================
function saveSpecialPricesBatch(customerId, items) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: "系統忙碌" };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SPECIAL_PRICE);
    
    // 1. 取得輔助資料 (客戶名與商品清單)
    var customers = getCustomers();
    var products = getProducts(); // 為了抓取商品原價與名稱
    
    var customerObj = customers.find(function(c) { return c.id === customerId; });
    var custName = customerObj ? customerObj.name : "未知客戶";
    var nowStr = getNowString().split(' ')[0];
    
    // 2. 建立現有特價的索引 & 計算目前最大排序
    var existingData = sheet.getDataRange().getValues();
    var mapIndex = {}; 
    var maxSortMap = {}; // 🔥 修正點 1：在此宣告變數
    
    // 從第 2 列開始讀 (略過標題)
    for (var i = 1; i < existingData.length; i++) {
      var rowCustId = String(existingData[i][1]); // B欄 CustId
      var rowProdId = String(existingData[i][3]); // D欄 ProdId
      var rowSort = Number(existingData[i][10]) || 0; // K欄 排序
      
      // 建立索引：Key = 客戶ID_商品ID
      var key = rowCustId + "_" + rowProdId; 
      mapIndex[key] = i + 1; // 記錄列號 (Row Index)

      // 🔥 修正點 2：計算該客戶目前的最大排序
      if (!maxSortMap[rowCustId]) maxSortMap[rowCustId] = 0;
      if (rowSort > maxSortMap[rowCustId]) {
        maxSortMap[rowCustId] = rowSort;
      }
    }

    var updateCount = 0;
    var createCount = 0;

    // 3. 逐筆處理傳入的特價項目
    items.forEach(function(item) {
      var prodObj = products.find(function(p) { return p.id === item.id; });
      var prodName = prodObj ? prodObj.name : "未知商品";
      var prodOrigPrice = prodObj ? prodObj.price : 0;
      
      var uniqueKey = String(customerId) + "_" + String(item.id);
      var targetRow = mapIndex[uniqueKey];

      if (targetRow) {
        // --- 狀況 A: 已經設定過 -> 更新價格與排序 ---
        // F欄 (Index 6) = 特價
        sheet.getRange(targetRow, 6).setValue(Number(item.price));
        // K欄 (Index 11) = 排序 (如果有傳入就更新，沒傳入維持原樣或預設)
        if (item.sortOrder) {
           sheet.getRange(targetRow, 11).setValue(item.sortOrder);
        }
        // G欄 (Index 7) = 原價 (順便更新一下原價)
        sheet.getRange(targetRow, 7).setValue(prodOrigPrice);
        
        updateCount++;
      } else {
        // --- 狀況 B: 沒設定過 -> 新增一筆 ---
        
        // 🔥 修正點 3：正確取出目前最大排序並 +1
        var currentMax = maxSortMap[customerId] || 0;
        var nextSort = currentMax + 1;

        var spId = generateId("SP", CONFIG.SHEETS.SPECIAL_PRICE, 1);
        sheet.appendRow([
          spId,                 // A: 流水號
          String(customerId),   // B: 客戶ID
          custName,             // C: 客戶名稱
          String(item.id),      // D: 商品ID
          prodName,             // E: 商品名稱
          Number(item.price),   // F: 特價
          prodOrigPrice,        // G: 原價
          nowStr,               // H: 生效日
          "",                   // I: 到期日
          "批次設定",            // J: 備註
          nextSort              // K: 排序
        ]);
        
        // 更新記憶體中的最大值，以免連續新增時排序重複
        maxSortMap[customerId] = nextSort;
        createCount++;
      }
    });
    
    // 強制寫入
    SpreadsheetApp.flush();

    return { success: true, count: updateCount + createCount, message: "新增 " + createCount + " 筆，更新 " + updateCount + " 筆" };
    
  } catch (e) { 
    return { success: false, error: e.toString() }; 
  } finally { 
    lock.releaseLock(); 
  }
}

function getSpecialPrices() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SPECIAL_PRICE);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    var results = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[1] || !row[3]) continue;
      var fmtDate = function(d) { return d ? Utilities.formatDate(new Date(d), "GMT+8", "yyyy-MM-dd") : ''; };
      results.push({
        id: String(row[0] || ''), custId: String(row[1] || '').trim(), custName: String(row[2] || ''),
        prodId: String(row[3] || '').trim(), prodName: String(row[4] || ''), price: Number(row[5]) || 0,
        originalPrice: Number(row[6]) || 0, startDate: fmtDate(row[7]), endDate: fmtDate(row[8]),
        note: String(row[9] || ''), sortOrder: (row[10] !== "" && row[10] !== null) ? Number(row[10]) : 999
      });
    }
    return results;
  } catch (e) { return []; }
}
// ==========================================
// 7. 📦 Inventory (庫存模組)
// ==========================================

function batchUpdateInventory(items, docNo, type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
  var logSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY_LOG);
  var nowStr = getNowString();
  var invData = invSheet.getDataRange().getValues();
  var logRows = [];
  var productMap = {};
  for (var i = 1; i < invData.length; i++) productMap[String(invData[i][0])] = i;

  items.forEach(function(item) {
    var productId = String(item.productId);
    var qtyChange = -Math.abs(Number(item.qty) || 0);
    var rowIndex = productMap[productId];
    if (rowIndex !== undefined) {
      var oldQty = Number(invData[rowIndex][2]) || 0;
      var newQty = oldQty + qtyChange;
      invData[rowIndex][2] = newQty;
      invData[rowIndex][8] = nowStr.split(' ')[0];
      logRows.push([nowStr, productId, invData[rowIndex][1], type, qtyChange, invData[rowIndex][4], docNo, newQty, nowStr]);
    }
  });
  invSheet.getRange(1, 1, invData.length, invData[0].length).setValues(invData);
  if (logRows.length > 0) logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, 9).setValues(logRows);
}

/**
 * 修正後的進貨庫存更新 (最終版)
 * 功能：
 * 1. 更新原料主檔庫存
 * 2. 更新成品庫存 (若庫存表沒資料，會自動新增一行)
 * 3. 寫入異動日誌
 */
function batchUpdateInventoryForPurchase(items, docNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY_LOG);
  var prodInvSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY); // 成品庫存表
  var matMasterSheet = ss.getSheetByName(CONFIG.SHEETS.MATERIALS); // 原料主檔
  var nowStr = getNowString();

  // 1. 讀取現有資料
  var prodData = prodInvSheet.getDataRange().getValues();
  var matData = matMasterSheet ? matMasterSheet.getDataRange().getValues() : [];

  // 2. 建立索引 (ID -> Row Index) 方便快速查找
  var prodMap = {};
  for (var i = 1; i < prodData.length; i++) prodMap[String(prodData[i][0])] = i;

  var matMap = {};
  for (var j = 1; j < matData.length; j++) matMap[String(matData[j][0])] = j;

  var logRows = [];
  var newProdRows = []; // 🔥 用來暫存「庫存表還沒有的新商品」

  items.forEach(function(item) {
    var pid = String(item.productId);
    var addQty = Number(item.qty);
    var pName = item.productName;
    var unit = item.unit || "個";

    // ==========================================
    // 情況 A：如果是原料 (去原料主檔更新 I 欄)
    // ==========================================
    if (item.itemType === 'MATERIAL' && matMasterSheet) {
       var rowIdx = matMap[pid];
       if (rowIdx !== undefined) {
         var oldQty = Number(matData[rowIdx][8]) || 0; // I欄 (Index 8) 是目前庫存
         var newQty = oldQty + addQty;
         matData[rowIdx][8] = newQty;
         
         logRows.push([nowStr, pid, pName, "原料進貨", addQty, unit, docNo, newQty, nowStr]);
       }
    } 
    // ==========================================
    // 情況 B：如果是成品 (去庫存表更新 C 欄)
    // ==========================================
    else {
       var rowIdx = prodMap[pid];
       
       if (rowIdx !== undefined) {
         // 🟢 狀況 1: 庫存表已經有這個商品 -> 直接更新數量
         var oldQty = Number(prodData[rowIdx][2]) || 0; // C欄 (Index 2)
         var newQty = oldQty + addQty;
         prodData[rowIdx][2] = newQty; // 更新記憶體中的數據
         
         logRows.push([nowStr, pid, pName, "成品進貨", addQty, unit, docNo, newQty, nowStr]);
       } else {
         // 🔴 狀況 2: 庫存表還沒有這個商品 -> 準備新增一行
         // 庫存表結構預設: [ID, Name, Qty, Spec, Unit, Cost, Price, Safety, Date, Note]
         newProdRows.push([
           pid, 
           pName, 
           addQty, // 初始庫存 = 這次進貨量
           "",     // 規格 (留空)
           unit, 
           0,      // 售價 (暫補0)
           0,      // 成本 (暫補0)
           5,      // 安全庫存 (預設5)
           nowStr
         ]);
         
         logRows.push([nowStr, pid, pName, "成品進貨(新)", addQty, unit, docNo, addQty, nowStr]);
       }
    }
  });

  // 3. 寫回資料庫
  
  // A. 寫回原本已存在的成品庫存 (更新數量)
  if (prodData.length > 0) {
    prodInvSheet.getRange(1, 1, prodData.length, prodData[0].length).setValues(prodData);
  }
  
  // B. 🔥 寫入那些新商品到庫存表底部
  if (newProdRows.length > 0) {
    prodInvSheet.getRange(prodInvSheet.getLastRow() + 1, 1, newProdRows.length, newProdRows[0].length).setValues(newProdRows);
  }

  // C. 寫回原料庫存
  if (matMasterSheet && matData.length > 0) {
    matMasterSheet.getRange(1, 1, matData.length, matData[0].length).setValues(matData);
  }
  
  // D. 寫入異動日誌 (Log)
  if (logRows.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, 9).setValues(logRows);
  }
}

function batchUpdateInventoryForDelete(items, docNo, type) { updateInventoryGeneral(items, docNo, type, true); }
function batchUpdateInventoryForPurchaseDelete(items, docNo, type) { updateInventoryGeneral(items, docNo, type, false); }

function updateInventoryGeneral(items, docNo, type, isRestore) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
  var logSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY_LOG);
  var nowStr = getNowString();
  var invData = invSheet.getDataRange().getValues();
  var productMap = {};
  for (var i = 1; i < invData.length; i++) productMap[String(invData[i][0])] = i;
  items.forEach(function(item) {
    var rowIndex = productMap[String(item.productId)];
    if (rowIndex !== undefined) {
      var oldQty = Number(invData[rowIndex][2]) || 0;
      var qty = Math.abs(Number(item.qty));
      var change = isRestore ? qty : -qty;
      var newQty = oldQty + change;
      invData[rowIndex][2] = newQty;
      logSheet.appendRow([nowStr, item.productId, invData[rowIndex][1], type, change, invData[rowIndex][4], docNo, newQty, nowStr]);
    }
  });
  invSheet.getRange(1, 1, invData.length, invData[0].length).setValues(invData);
}

function convertProductStock(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    var logSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY_LOG);
    var nowStr = getNowString();
    var invData = invSheet.getDataRange().getValues();
    var productMap = {};
    for (var i = 1; i < invData.length; i++) productMap[String(invData[i][0])] = i;
    var rawIdx = productMap[data.rawId];
    if (rawIdx) {
      invData[rawIdx][2] = Number(invData[rawIdx][2]) - Number(data.rawQty);
      logSheet.appendRow([nowStr, data.rawId, invData[rawIdx][1], "分裝消耗", -data.rawQty, invData[rawIdx][4], "轉換單", invData[rawIdx][2], nowStr]);
    }
    var targetIdx = productMap[data.targetId];
    if (targetIdx) {
      invData[targetIdx][2] = Number(invData[targetIdx][2]) + Number(data.targetQty);
      logSheet.appendRow([nowStr, data.targetId, invData[targetIdx][1], "分裝產出", data.targetQty, invData[targetIdx][4], "轉換單", invData[targetIdx][2], nowStr]);
    }
    invSheet.getRange(1, 1, invData.length, invData[0].length).setValues(invData);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function saveStocktake(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    var logSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY_LOG);
    var stocktakeSheet = ss.getSheetByName(CONFIG.SHEETS.STOCKTAKE);
    var nowStr = getNowString();
    var stocktakeId = generateId("ST", CONFIG.SHEETS.STOCKTAKE, 1);
    if (stocktakeSheet) {
      stocktakeSheet.appendRow([
        stocktakeId, getNowString().split(' ')[0], data.productId, data.productName,
        data.bookQty, data.actualQty, (data.actualQty - data.bookQty), data.note, nowStr
      ]);
    }
    var invData = invSheet.getDataRange().getValues();
    for (var i = 1; i < invData.length; i++) {
      if (String(invData[i][0]) === data.productId) {
        invData[i][2] = data.actualQty;
        invSheet.getRange(i + 1, 3).setValue(data.actualQty);
        var diff = data.actualQty - data.bookQty;
        if (diff !== 0) {
          logSheet.appendRow([nowStr, data.productId, data.productName, "盤點調整", diff, invData[i][4], stocktakeId, data.actualQty, nowStr]);
        }
        break;
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getInventory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    var prodSheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
    if (!invSheet) return [];
    var invData = invSheet.getDataRange().getValues();
    var prodData = prodSheet ? prodSheet.getDataRange().getValues() : [];
    var alertMap = {};
    for (var i = 1; i < prodData.length; i++) alertMap[String(prodData[i][0])] = Number(prodData[i][11]) || 5;
    var results = [];
    for (var i = 1; i < invData.length; i++) {
      var row = invData[i];
      var id = String(row[0]);
      if (!id) continue;
      results.push({
        productId: id, productName: String(row[1]), currentStock: Number(row[2]) || 0,
        unit: String(row[4] || '個'), safetyStock: alertMap[id] || 5
      });
    }
    return results;
  } catch (e) { return []; }
}

function getLowStockItems() {
  try {
    var allItems = getInventory();
    return allItems.filter(function(item) { return item.currentStock <= item.safetyStock; });
  } catch (e) { return []; }
}

function getInventoryLog() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.INVENTORY_LOG);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var limit = 50;
    var startRow = Math.max(2, lastRow - limit + 1);
    var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 9).getValues();
    var results = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var qty = Number(row[4]);
      results.push({
        date: formatDate(row[0]), productName: String(row[2]), type: String(row[3]),
        qty: Math.abs(qty), unit: String(row[5]), docNo: String(row[6]),
        displayType: qty > 0 ? '入庫' : '出庫'
      });
    }
    return results.reverse();
  } catch (e) { return []; }
}

// ==========================================
// 8. 📤 Sales (銷貨模組)
// ==========================================

function getDataForSalesOrder() {
  try {
    var result = {
      customers: getCustomers(), 
      products: getProductsLite(), 
      specialPrices: getSpecialPrices(),
      copyData: null  // 👈 新增這個欄位
    };
    
    // ✅ 順便檢查是否有複製資料（不用前端再請求一次）
    var props = PropertiesService.getUserProperties();
    var json = props.getProperty('COPY_ORDER_DATA');
    if (json) {
      result.copyData = JSON.parse(json);
      props.deleteProperty('COPY_ORDER_DATA');
    }
    
    return result;
    
  } catch (error) { 
    throw new Error("初始化失敗: " + error); 
  }
}

function saveSalesOrder(data, forcedId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
    var detailSheet = ss.getSheetByName(CONFIG.SHEETS.SALES_DETAILS);
    var receivableSheet = ss.getSheetByName(CONFIG.SHEETS.RECEIVABLE);
    
    var orderId = forcedId ? forcedId : generateId("SO", CONFIG.SHEETS.SALES, 1);
    var nowStr = getNowString().split(' ')[0];

    var customers = getCustomers();
    var customer = customers.find(function(c) { return c.id === data.customer.id; });
    var taxType = customer ? (customer.taxType || '免稅') : '免稅';
    
    var subtotal = Number(data.totalAmount) || 0;
    var shipping = Number(data.shipping) || 0;
    var deduction = Number(data.deduction) || 0;
    var taxAmount = (taxType.indexOf('外加') > -1) ? Math.round(subtotal * 0.05) : 0;
    var grandTotal = subtotal + taxAmount + shipping - deduction;
    
    salesSheet.appendRow([
      orderId, data.date, data.customer.id, data.customer.name,
      subtotal, shipping, deduction, taxAmount, grandTotal,
      taxType, '未收款', nowStr, data.note || ''
      
    ]);

    var detailRows = [];
    data.items.forEach(function(item, index) {
      var qty = Number(item.qty) || 0;
      var unitPrice = Number(item.price) || 0;
      var unitCost = Number(item.cost) || 0;
      
      detailRows.push([
        orderId + "-" + (index + 1), orderId, item.productId, item.productName,
        qty, item.unit, unitPrice, qty * unitPrice, unitCost, (qty * unitPrice) - (qty * unitCost), item.note || ''
      ]);
    });

    if (detailRows.length > 0) {
      detailSheet.getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, 11).setValues(detailRows);
    }

    if (receivableSheet) {
      var arId = generateId("AR", CONFIG.SHEETS.RECEIVABLE, 1);
      receivableSheet.appendRow([
        arId, data.customer.id, data.customer.name, orderId, data.date,
        grandTotal, 0, grandTotal, formatDate(new Date(data.date)), 0, '未收款', ''
      ]);
    }

    batchUpdateInventory(data.items, orderId, '銷貨出庫');
    return { success: true, orderId: orderId };
  } catch (error) { 
    return { success: false, error: error.toString() }; 
  }
}

/**
 * 🔍 查詢銷貨單 (包含付款狀態篩選)
 * 修改日期：2026-02-11
 */
/**
 * 🔍 查詢銷貨單 (修正版：修復 customerId 未定義與 rowCustId 遺失導致查不到資料的問題)
 */
function searchSalesOrders(keyword, dateFrom, dateTo, statusFilter, customerId) { // 修正1: 補上 customerId 參數
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SALES);
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    var results = [];
    
    // 日期處理
    var from = dateFrom ? new Date(dateFrom) : null;
    if (from) from.setHours(0, 0, 0, 0); 
    
    var to = dateTo ? new Date(dateTo) : null;
    if (to) to.setHours(23, 59, 59, 999); 
    
    // 關鍵字轉小寫
    var k = keyword ? keyword.toLowerCase() : "";

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; // 沒有單號跳過
      
      var orderDate = row[1] instanceof Date ? row[1] : new Date(row[1]);
      
      // 1. 日期篩選
      if (from && orderDate < from) continue;
      if (to && orderDate > to) continue;
      
      var orderId = String(row[0]);
      var rowCustId = String(row[2]); // 修正2: 定義 rowCustId (C欄是客戶ID)
      var custName = String(row[3]);
      var payStatus = String(row[10] || "未收款"); // K欄 (Index 10) 是付款狀態
      
      // 修正3: 使用參數傳進來的 customerId 進行比對
      if (customerId && rowCustId !== customerId) continue;
      
      // 2. 付款狀態篩選
      if (statusFilter && statusFilter !== 'ALL') {
        if (statusFilter === 'PAID' && payStatus !== '已收款') continue;
        if (statusFilter === 'UNPAID' && payStatus === '已收款') continue;
      }

      // 3. 關鍵字篩選 (搜尋 單號 或 客戶名稱)
      if (k && orderId.toLowerCase().indexOf(k) === -1 && custName.toLowerCase().indexOf(k) === -1) {
        continue;
      }
      
      // 4. 組合回傳資料
      results.push({ 
        orderId: orderId, 
        date: formatDate(orderDate), 
        customerName: custName, 
        total: Number(row[8]) || 0,
        note: String(row[12] || ''),
        paymentStatus: payStatus 
      });
    }
    
    // 排序：日期新到舊
    results.sort(function(a, b) { 
      var dateA = new Date(a.date);
      var dateB = new Date(b.date);
      var dateDiff = dateB - dateA;
      
      if (dateDiff !== 0) {
        return dateDiff; 
      } else {
        return b.orderId.localeCompare(a.orderId); 
      }
    });
    
    return results;
  } catch (e) { 
    // 建議：將錯誤 log 出來，方便除錯，而不是只回傳空陣列
    console.error("searchSalesOrders Error: " + e.toString()); 
    return []; 
  }
}

function getSalesOrderDetail(orderId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
    var detailSheet = ss.getSheetByName(CONFIG.SHEETS.SALES_DETAILS);
    var salesData = salesSheet.getDataRange().getValues();
    var orderRow = null;
    for (var i = 1; i < salesData.length; i++) {
      if (salesData[i][0] === orderId) { orderRow = salesData[i]; break; }
    }
    if (!orderRow) return { success: false, error: "找不到訂單" };
    var detailData = detailSheet.getDataRange().getValues();
    var items = [];
    for (var i = 1; i < detailData.length; i++) {
      if (detailData[i][1] === orderId) {
        items.push({
          productId: String(detailData[i][2]), productName: String(detailData[i][3]),
          qty: Number(detailData[i][4]), unit: String(detailData[i][5]),
          price: Number(detailData[i][6]), cost: Number(detailData[i][8]),
          note: String(detailData[i][10] || '')
        });
      }
    }
    return {
      success: true,
      order: {
        orderId: orderId, date: orderRow[1] instanceof Date ? formatDate(orderRow[1]) : String(orderRow[1]),
        customerId: String(orderRow[2]), customerName: String(orderRow[3]),
        subtotal: Number(orderRow[4]), shipping: Number(orderRow[5]), deduction: Number(orderRow[6]),
        taxAmount: Number(orderRow[7]), grandTotal: Number(orderRow[8]), taxType: String(orderRow[9]), items: items
      }
    };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteSalesOrder(orderId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
  var detailSheet = ss.getSheetByName(CONFIG.SHEETS.SALES_DETAILS);
  var receivableSheet = ss.getSheetByName(CONFIG.SHEETS.RECEIVABLE);
  try {
    var details = detailSheet.getDataRange().getValues();
    var itemsToRestore = [];
    for (var i = 1; i < details.length; i++) {
      if (String(details[i][1]) === orderId) itemsToRestore.push({ productId: details[i][2], qty: details[i][4] });
    }
    if (itemsToRestore.length > 0) batchUpdateInventoryForDelete(itemsToRestore, orderId, "銷貨刪除回補");
    deleteRowsById(salesSheet, 1, orderId);
    deleteRowsById(detailSheet, 2, orderId);
    deleteRowsById(receivableSheet, 4, orderId);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function updateSalesOrder(oldOrderId, newData) {
  var delRes = deleteSalesOrder(oldOrderId);
  if (!delRes.success) return delRes;
  return saveSalesOrder(newData, oldOrderId);
}



/**
 * 搜尋商品：包含最新進價 + 客戶等級價 + 客戶特價
 * @param {string} keyword 關鍵字
 * @param {string} customerId 目前選取的客戶ID
 * @param {string} priceGroup 客戶所屬的價格群組 (龍潭價/店家價/餐廳價)
 */
function searchProductsWithSmartPrice(keyword, customerId, priceGroup) {
  if (!keyword) return [];

  // 1. 快速搜尋商品主檔 (僅回傳前 20-30 筆)
  var products = searchProductsByKeyword(keyword); 
  if (products.length === 0) return [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 2. 讀取進價記錄 (取最新成本)
  var priceSheet = ss.getSheetByName(CONFIG.SHEETS.SUPPLIER_PRICES);
  var costData = priceSheet ? priceSheet.getDataRange().getValues() : [];

  // 3. 讀取特價表 (過濾出該客戶的特價)
  var spSheet = ss.getSheetByName(CONFIG.SHEETS.SPECIAL_PRICE);
  var spData = spSheet ? spSheet.getDataRange().getValues() : [];

  // 4. 開始組合成員資料
  products.forEach(function(p) {
    // A. 決定群組價 (龍潭/店家/餐廳)
    // 假設你的商品主檔 F:原價, G:龍潭價, H:店家價, I:餐廳價
    if (priceGroup === "龍潭價") p.price = p.priceLongtan || p.price;
    else if (priceGroup === "店家價") p.price = p.priceStore || p.price;
    else if (priceGroup === "餐廳價") p.price = p.priceRest || p.price;

    // B. 檢查是否有「客戶專屬特價」 (最高優先權)
    var special = spData.find(row => String(row[1]) === customerId && String(row[3]) === p.id);
    if (special) {
      p.price = Number(special[5]); // 假設 F 欄是特價
      p.isSpecial = true; // 標記一下，讓前端可以變色提醒
    }

    // C. 抓取最新進成本
    for (var i = costData.length - 1; i >= 1; i--) {
      if (String(costData[i][2]) === p.id) {
        p.cost = Number(costData[i][4]) || 0;
        break;
      }
    }
  });

  return products;
}

// ==========================================
// 9. 📥 Purchase (進貨模組)
// ==========================================

/**
 * 🚀 極速版：一次讀取所有基礎資料 + 複製/編輯暫存
 */
function getDataForPurchaseOrder() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var results = {
      suppliers: [],
      products: [],
      materials: [],
      copyData: null
    };

    // 1. 同時獲取所有工作表
    var sheets = {
      suppliers: ss.getSheetByName(CONFIG.SHEETS.SUPPLIERS),
      products: ss.getSheetByName(CONFIG.SHEETS.PRODUCTS),
      materials: ss.getSheetByName(CONFIG.SHEETS.MATERIALS)
    };

    // 2. 批次處理供應商 (J欄是付款方式, O欄是稅別)
    if (sheets.suppliers) {
      var sData = sheets.suppliers.getDataRange().getValues();
      for (var i = 1; i < sData.length; i++) {
        if (sData[i][0]) {
          results.suppliers.push({
            id: String(sData[i][0]),
            name: String(sData[i][1]),
            paymentMethod: String(sData[i][9] || "老闆付款"),
            taxType: String(sData[i][14] || "免稅")
          });
        }
      }
    }

    // 3. 批次處理成品商品 (M欄是狀態, J欄是成本)
    if (sheets.products) {
      var pData = sheets.products.getDataRange().getValues();
      for (var i = 1; i < pData.length; i++) {
        var status = String(pData[i][12] || "啟用");
        if (status.indexOf("停用") > -1 || !pData[i][0]) continue;
        results.products.push({
          id: String(pData[i][0]),
          category: String(pData[i][1]),
          name: String(pData[i][2]),
          spec: String(pData[i][3]),
          unit: String(pData[i][4] || "個"),
          cost: Number(pData[i][9]) || 0
        });
      }
    }

    // 4. 批次處理原物料 (D欄是類別, F欄是成本)
    if (sheets.materials) {
      var mData = sheets.materials.getDataRange().getValues();
      for (var i = 1; i < mData.length; i++) {
        if (!mData[i][0]) continue;
        results.materials.push({
          id: String(mData[i][0]),
          name: String(mData[i][1]),
          spec: String(mData[i][2]),
          category: String(mData[i][3]),
          unit: String(mData[i][4] || "個"),
          cost: Number(mData[i][5]) || 0
        });
      }
    }

    // 5. 順便檢查是否有複製/編輯資料 (減少前端一次網路請求)
    var props = PropertiesService.getUserProperties();
    var json = props.getProperty('COPY_PURCHASE_ORDER');
    if (json) {
      results.copyData = JSON.parse(json);
      props.deleteProperty('COPY_PURCHASE_ORDER'); // 讀完即刪，保持乾淨
    }

    return results;
  } catch (error) {
    throw new Error("初始化提速讀取失敗: " + error);
  }
}

/**
 * 🔍 查詢進貨單 (修正版：支援付款狀態篩選)
 */
function searchPurchaseOrders(keyword, dateFrom, dateTo, supplierId, statusFilter) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PURCHASE);
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var results = [];
    
    var key = keyword ? keyword.toLowerCase() : "";
    var from = dateFrom ? new Date(dateFrom).getTime() : null;
    var to = dateTo ? new Date(dateTo).getTime() : null;
    if (to) { var toDate = new Date(dateTo); toDate.setHours(23, 59, 59, 999); to = toDate.getTime(); }
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var orderId = String(row[0]);
      if (!orderId) continue;
      
      // 1. 日期篩選
      var dTime = new Date(row[1]).getTime();
      if (from && dTime < from) continue;
      if (to && dTime > to) continue;
      
      // 2. 供應商篩選
      if (supplierId && String(row[2]) !== supplierId) continue;
      
      // 3. 關鍵字篩選
      var oid = orderId.toLowerCase();
      var supName = String(row[3]).toLowerCase();
      if (key && oid.indexOf(key) === -1 && supName.indexOf(key) === -1) continue;

      // 4. ⭐ 新增：付款狀態篩選 (第9欄 / Index 8 是付款狀態)
      var payStatus = String(row[8] || "未付款");
      if (statusFilter && statusFilter !== 'ALL') {
        if (statusFilter === 'PAID' && payStatus !== '已付款') continue;
        if (statusFilter === 'UNPAID' && payStatus === '已付款') continue;
      }

      results.push({
        orderId: row[0], 
        date: formatDate(row[1]), 
        supplierName: row[3], 
        total: Number(row[6]) || 0,
        status: payStatus // ⭐ 回傳狀態給前端顯示
      });
    }
    
    // 排序：單號新到舊
    results.sort(function(a, b) { return b.orderId.localeCompare(a.orderId); });
    return results;
  } catch (e) { return []; }
}

function getPurchaseOrderDetail(orderId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var purchaseSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE);
    var detailSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_DETAILS);
    var supplierSheet = ss.getSheetByName(CONFIG.SHEETS.SUPPLIERS);
    
    // 1. 抓取訂單主檔
    var purchaseData = purchaseSheet.getDataRange().getValues();
    var orderRow = null;
    for (var i = 1; i < purchaseData.length; i++) {
      if (String(purchaseData[i][0]) === String(orderId)) { 
        orderRow = purchaseData[i]; 
        break; 
      }
    }
    if (!orderRow) return { success: false, error: "找不到訂單" };

    // 2. ⭐ 解決 undefined：去供應商表補抓資料
    var supplierId = String(orderRow[2]);
    var pTerm = "貨到付現";   // 預設值
    var pMethod = "老闆付款"; // 預設值
    
    if (supplierSheet) {
      var suppData = supplierSheet.getDataRange().getValues();
      for (var k = 1; k < suppData.length; k++) {
        if (String(suppData[k][0]) === supplierId) {
          pTerm = String(suppData[k][8] || "貨到付現"); // I欄
          pMethod = String(suppData[k][9] || "老闆付款"); // J欄
          break;
        }
      }
    }

    // 3. 抓取明細
    var detailData = detailSheet.getDataRange().getValues();
    var items = [];
    for (var i = 1; i < detailData.length; i++) {
      if (String(detailData[i][1]) === String(orderId)) {
        items.push({
          productId: String(detailData[i][3]), 
          productName: String(detailData[i][4]),
          qty: Number(detailData[i][5]), 
          unit: String(detailData[i][6]), 
          cost: Number(detailData[i][7])
        });
      }
    }

    return {
      success: true,
      order: {
        orderId: orderId, 
        date: formatDate(orderRow[1]), 
        supplierId: String(orderRow[2]),
        supplierName: String(orderRow[3]), 
        totalAmount: Number(orderRow[4]),
        taxAmount: Number(orderRow[5]), 
        grandTotal: Number(orderRow[6]),
        // ⭐ 這裡確保回傳字串，不再是 undefined
        paymentTerm: pTerm,      
        paymentMethod: pMethod,   
        items: items,
        note: String(orderRow[11] || '')
      }
    };
  } catch (error) { 
    return { success: false, error: error.toString() }; 
  }
}

/**
 * 📥 儲存進貨單 (修復版：已補上「應付帳款」寫入功能)
 */
function savePurchaseOrder(data, forcedId) {
  var ss = SpreadsheetApp.getActive();
  
  // 1. 取得所有相關工作表
  var pur = ss.getSheetByName(CONFIG.SHEETS.PURCHASE);
  var det = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_DETAILS);
  var paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_MADE);
  var expSheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES);      // 現金支出
  var matSheet = ss.getSheetByName(CONFIG.SHEETS.MATERIALS);
  
  // 🔥【修復 1】這裡補上讀取應付帳款工作表
  var apSheet = ss.getSheetByName(CONFIG.SHEETS.PAYABLE);        
  
  // 防呆：確認關鍵工作表存在
  if (!pur || !det) return { success: false, error: "❌ 找不到進貨單或明細工作表，請檢查名稱" };

  var id = forcedId || generateId("P", CONFIG.SHEETS.PURCHASE, 1);
  var nowStr = getNowString().split(' ')[0]; // 取得 YYYY-MM-DD
  
  // 確保金額是數字
  var totalAmount = Number(data.totalAmount) || 0;
  var taxAmount = Number(data.taxAmount) || 0;
  var totalWithTax = totalAmount + taxAmount;

  // ==========================================
  // A. 準備明細資料 & 檢查原料價格波動
  // ==========================================
  var matCosts = {};
  if (matSheet) {
    var matData = matSheet.getDataRange().getValues();
    for (var m = 1; m < matData.length; m++) {
      // 假設 A欄是ID(0), F欄是成本(5)
      matCosts[String(matData[m][0])] = Number(matData[m][5]); 
    }
  }

  var rows = [];
  var priceAlerts = []; 

  // 🔥 改進：先在外部產生一次 ID 基準點，在記憶體內累加數字
  var pdPrefix = "PD";
  var pdBaseId = generateId(pdPrefix, CONFIG.SHEETS.PURCHASE_DETAILS, 1);
  var pdNum = parseInt(pdBaseId.replace(/[^0-9]/g, ''), 10);

  data.items.forEach(function(it, i) {
    var pid = String(it.productId);
    var inputCost = Number(it.cost); 

    // 檢查價格波動
    if (matCosts[pid] !== undefined && matCosts[pid] > 0 && inputCost !== matCosts[pid]) {
      var diff = inputCost - matCosts[pid];
      var diffPercent = ((diff / matCosts[pid]) * 100).toFixed(0); 
      var symbol = diff > 0 ? "🔺" : "🔻"; 
      priceAlerts.push(it.productName + symbol + diffPercent + "%");
    }
    var currentPdId = pdPrefix + ("00000" + (pdNum + i)).slice(-5); // 這行留著或刪掉都沒關係
    // 產生明細資料
    rows.push([
      id + "-" + (i + 1), // ✅ 改成這樣：用進貨單號(id)加上流水號 (例如 P00010-1)
      id, 
      data.supplier.id,
      it.productId, 
      it.productName, 
      Number(it.qty), 
      it.unit, 
      Number(it.cost), 
      Math.round(Number(it.qty) * Number(it.cost)), 
      nowStr, 
      ''
    ]);
  });

// ==========================================
  // 🔥【新增】同步寫入「供應商進價記錄」
  // ==========================================
  var priceSheet = ss.getSheetByName(CONFIG.SHEETS.SUPPLIER_PRICES);
  if (priceSheet) {
    var priceRows = [];
    data.items.forEach(function(it, idx) {
      // 產生 ID (或是用時間戳記)
      var priceId = generateId("PR", CONFIG.SHEETS.SUPPLIER_PRICES, 1); 
      // 注意：如果您不想每次都 generateId 拖慢速度，可以用 "PR" + nowStr + idx 
      
      priceRows.push([
        priceId,                // A: 流水號
        data.date,              // B: 日期
        it.productId,           // C: 商品ID
        it.productName,         // D: 商品名稱
        Number(it.cost),        // E: 成本
        it.unit,                // F: 單位
        data.supplier.id,       // G: 供應商ID
        data.supplier.name,     // H: 供應商名稱
        nowStr                  // I: 記錄時間
      ]);
    });
    
    if (priceRows.length > 0) {
      priceSheet.getRange(priceSheet.getLastRow() + 1, 1, priceRows.length, 9).setValues(priceRows);
    }
  }
  // ==========================================

  // ==========================================
  // B. 寫入資料庫 (主檔 & 明細)
  // ==========================================
  var finalNote = data.note || '';
  if (priceAlerts.length > 0) {
    finalNote += " [波動: " + priceAlerts.join(", ") + "]";
  }

  // 寫入主檔
  pur.appendRow([
    id, 
    data.date, 
    data.supplier.id, 
    data.supplier.name, 
    totalAmount, 
    data.taxAmount, 
    totalWithTax,
    data.handler || '', 
    data.paymentStatus || '未付款', 
    '', 
    nowStr, 
    finalNote
  ]);

  // 寫入明細
  if (rows.length > 0) {
    det.getRange(det.getLastRow() + 1, 1, rows.length, 11).setValues(rows);
  }

  // ==========================================
  // 🔥【修復 2】這裡補上寫入應付帳款 (PAYABLE) 的邏輯
  // ==========================================
  if (apSheet) {
    var apId = generateId("AP", CONFIG.SHEETS.PAYABLE, 1);
    
    // 計算初始已付與未付
    var initialPaid = (data.paymentStatus === '已付款') ? totalWithTax : 0;
    var initialUnpaid = totalWithTax - initialPaid;
    
    apSheet.appendRow([
      apId,               // A: 流水號
      data.supplier.id,   // B: 供應商編號
      data.supplier.name, // C: 供應商名稱
      id,                 // D: 進貨單號 (關聯)
      data.date,          // E: 進貨日期
      totalWithTax,       // F: 應付總額
      initialPaid,        // G: 已付金額
      initialUnpaid,      // H: 未付金額
      (initialPaid > 0) ? data.date : '', // I: 付款日期 (若已付)
      0,                  // J: 沖帳金額 (預設0)
      data.paymentStatus, // K: 狀態
      finalNote           // L: 備註
    ]);
  }

  // ==========================================
  // C. 連動邏輯：如果狀態是「已付款」
  // ==========================================
  if (data.paymentStatus === '已付款') {
    
    var payMethod = data.paymentMethod || '現金';
    var linkNote = '進貨單: ' + id + ' (' + data.supplier.name + ')';

    // 1. 寫入「付款紀錄」 (PAYMENT_MADE)
    if (paymentSheet) {
      paymentSheet.appendRow([
        generateId("PM", CONFIG.SHEETS.PAYMENT_MADE, 1), 
        data.date, 
        data.supplier.id, 
        data.supplier.name, 
        totalWithTax, 
        payMethod, 
        data.handler || '', 
        id, 
        nowStr, 
        data.note || ''
      ]);
    }

  // 2. 寫入「零用金紀錄」 (PETTY_CASH)
    if (payMethod.indexOf('零用金') > -1) {
      var pcSheet = ss.getSheetByName(CONFIG.SHEETS.PETTY_CASH);
      if (pcSheet) {
        var lastBal = 0;
        if (pcSheet.getLastRow() > 1) {
          lastBal = Number(pcSheet.getRange(pcSheet.getLastRow(), 8).getValue()) || 0;
        }
        var pcId = generateId("PC", CONFIG.SHEETS.PETTY_CASH, 1);
        pcSheet.appendRow([
          pcId,
          data.date,
          '出金',
          totalWithTax,
          '進貨付款-' + data.supplier.name,
          '',
          linkNote,
          lastBal - totalWithTax
        ]);
      }
    }

    // 3. 寫入「現金支出」 (EXPENSES)
    if (expSheet) {
      expSheet.appendRow([
        generateId("EXP", CONFIG.SHEETS.EXPENSES, 1),
        data.date,
        '進貨成本',            // 類別
        '支付貨款-' + data.supplier.name, // 項目
        totalWithTax,       // 金額
        payMethod,          // 支付方式
        data.supplier.name, // 對象
        '',                 // 發票
        '',
        '已支出',
        linkNote
      ]);
    }
  }

  // D. 更新庫存
  batchUpdateInventoryForPurchase(data.items, id);

  var returnMsg = "開單成功！單號：" + id;
  if (priceAlerts.length > 0) returnMsg += "\n⚠️ 注意：成本波動 " + priceAlerts.join("、");

  return { success: true, purchaseId: id, message: returnMsg };
}


function deletePurchaseOrder(orderId) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: "系統忙碌" };
  try {
    var ss = SpreadsheetApp.getActive();
    var pur = ss.getSheetByName(CONFIG.SHEETS.PURCHASE);
    var det = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_DETAILS);
    var ap = ss.getSheetByName(CONFIG.SHEETS.PAYABLE);
    var paySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_MADE);
    var expSheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES);
    var pcSheet = ss.getSheetByName(CONFIG.SHEETS.PETTY_CASH);

    var res = getPurchaseOrderDetail(orderId);
    if (!res.success) throw new Error("找不到進貨單");

    // 1. 回補庫存
    batchUpdateInventoryForPurchaseDelete(res.order.items, orderId, "刪除進貨");

    // 2. 刪除進貨主檔 & 明細 & 應付帳款
    deleteRowsById(pur, 1, orderId);
    deleteRowsById(det, 2, orderId);
    if (ap) deleteRowsById(ap, 4, orderId);

    // 3. 刪除付款紀錄 (H欄/第8欄 存的是進貨單號)
    if (paySheet) deleteRowsById(paySheet, 8, orderId);

    // 4. 刪除現金支出 (K欄/第11欄 備註裡包含進貨單號)
    if (expSheet) {
      var expData = expSheet.getDataRange().getValues();
      for (var i = expData.length - 1; i >= 1; i--) {
        if (String(expData[i][10] || '').indexOf(orderId) > -1) {
          expSheet.deleteRow(i + 1);
        }
      }
    }

    // 5. 刪除零用金紀錄 (G欄/第7欄 備註裡包含進貨單號)
    if (pcSheet) {
      var pcData = pcSheet.getDataRange().getValues();
      for (var i = pcData.length - 1; i >= 1; i--) {
        if (String(pcData[i][6] || '').indexOf(orderId) > -1) {
          pcSheet.deleteRow(i + 1);
        }
      }
    }

    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updatePurchaseOrder(oldOrderId, newData) {
  var del = deletePurchaseOrder(oldOrderId);
  if (!del.success) return del;
  return savePurchaseOrder(newData, oldOrderId);
}

function saveCopyPurchaseData(orderData) {
  try {
    PropertiesService.getUserProperties().setProperty('COPY_PURCHASE_ORDER', JSON.stringify(orderData));
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getCopyPurchaseData() {
  try {
    var props = PropertiesService.getUserProperties();
    var json = props.getProperty('COPY_PURCHASE_ORDER');
    if (json) {
      props.deleteProperty('COPY_PURCHASE_ORDER');
      return { success: true, data: JSON.parse(json) };
    }
    return { success: false };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function prepareAndOpenEditPurchaseForm(orderId) {
  try {
    var result = getPurchaseOrderDetail(orderId);
    if (!result.success) throw new Error(result.error);
    var orderData = result.order;
    orderData.isEditMode = true;
    orderData.oldOrderId = orderId;
    PropertiesService.getUserProperties().setProperty('COPY_PURCHASE_ORDER', JSON.stringify(orderData));
    showPurchaseOrderPanel();
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ==========================================
// 10. 💰 Payment & Expense (收付款與費用)
// ==========================================

function getAllCustomerOrders(customerId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var receivableSheet = ss.getSheetByName(CONFIG.SHEETS.RECEIVABLE);
    var paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_RECEIVED);
    if (!receivableSheet) return { success: false, error: "找不到『應收帳款』工作表" };
    var arData = receivableSheet.getDataRange().getValues();
    if (arData.length < 2) return { success: true, data: [] };
    var payData = paymentSheet ? paymentSheet.getDataRange().getValues() : [];
    var paidMap = {};
    for (var j = 1; j < payData.length; j++) {
      var oid = String(payData[j][1] || "").trim();
      var amt = Number(payData[j][4]) || 0;
      if (oid) paidMap[oid] = (paidMap[oid] || 0) + amt;
    }
    var results = [];
    for (var i = 1; i < arData.length; i++) {
      var rowCustomerId = String(arData[i][1] || "").trim();
      if (rowCustomerId.toLowerCase() === String(customerId).toLowerCase().trim()) {
        var oid = String(arData[i][3] || "");
        var total = Number(arData[i][5]) || 0;
        var paid = paidMap[oid] || 0;
        if (paid === 0 && (Number(arData[i][6]) || 0) > 0) paid = Number(arData[i][6]) || 0;
        var unpaid = total - paid;
        var status = unpaid <= 0 ? '已收款' : (paid > 0 ? '部分收款' : '未收款');
        results.push({ orderId: oid, orderDate: formatDate(arData[i][4]), totalAmount: total, paidAmount: paid, unpaid: unpaid, status: status });
      }
    }
    return { success: true, data: results };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function savePaymentReceived(paymentData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var receivableSheet = ss.getSheetByName(CONFIG.SHEETS.RECEIVABLE);
    var salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
    var paymentLogSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_RECEIVED);
    if (!paymentLogSheet) throw new Error("找不到『收款記錄』工作表");
    var paymentId = generateId("PY", CONFIG.SHEETS.PAYMENT_RECEIVED, 1);
    var nowStr = getNowString();
    paymentLogSheet.appendRow([
      paymentId, paymentData.orderId || "未指定", paymentData.customerId,
      paymentData.date, Number(paymentData.amount), paymentData.paymentMethod || "現金",
      paymentData.note || "", nowStr
    ]);
    if (paymentData.orderId && paymentData.orderId !== "未指定" && receivableSheet) {
      var arData = receivableSheet.getDataRange().getValues();
      for (var i = 1; i < arData.length; i++) {
        if (String(arData[i][3]).trim() === String(paymentData.orderId).trim()) {
          var currentPaid = Number(arData[i][6]) || 0;
          var total = Number(arData[i][5]) || 0;
          var newPaid = currentPaid + Number(paymentData.amount);
          var newUnpaid = total - newPaid;
          receivableSheet.getRange(i + 1, 7).setValue(newPaid);
          receivableSheet.getRange(i + 1, 8).setValue(newUnpaid);
          var newStatus = newUnpaid <= 0 ? '已收款' : (newPaid > 0 ? '部分收款' : '未收款');
          receivableSheet.getRange(i + 1, 11).setValue(newStatus);
          break;
        }
      }
      if (salesSheet) {
        var salesData = salesSheet.getDataRange().getValues();
        for (var j = 1; j < salesData.length; j++) {
          if (String(salesData[j][0]).trim() === String(paymentData.orderId).trim()) {
            salesSheet.getRange(j + 1, 11).setValue((newUnpaid <= 0) ? '已收款' : '部分收款');
            break;
          }
        }
      }
    }
    return { success: true, paymentId: paymentId };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getReceivableReport(dateFrom, dateTo, statusFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var arSheet = ss.getSheetByName(CONFIG.SHEETS.RECEIVABLE);
    var paySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_RECEIVED);
    if (!arSheet) return { summary: {}, list: [] };
    var arData = arSheet.getDataRange().getValues();
    var payData = paySheet ? paySheet.getDataRange().getValues() : [];
    var paidMap = {};
    for (var j = 1; j < payData.length; j++) {
      var oid = String(payData[j][1]); // B欄(1)是訂單編號
      paidMap[oid] = (paidMap[oid] || 0) + (Number(payData[j][4]) || 0);
    }
    var list = [], totalAR = 0, totalPaid = 0, totalUnpaid = 0;
    var from = dateFrom ? new Date(dateFrom) : null;
    var to = dateTo ? new Date(dateTo) : null;
    if (to) to.setHours(23, 59, 59);
    for (var i = 1; i < arData.length; i++) {
      var row = arData[i];
      if (!row[4]) continue;
      var d = new Date(row[4]);
      if (from && d < from) continue;
      if (to && d > to) continue;
      var oid = String(row[3]);
      var total = Number(row[5]) || 0;
      var actualPaid = paidMap[oid] || 0;
      if (actualPaid === 0 && (Number(row[6]) || 0) > 0) actualPaid = Number(row[6]);
      var unpaid = total - actualPaid;
      var status = unpaid <= 0 ? '已收款' : (actualPaid > 0 ? '部分收款' : '未收款');
      if (statusFilter && statusFilter !== '全部' && status !== statusFilter) continue;
      list.push({
        orderId: oid, date: formatDate(d), customerName: String(row[2]),
        total: total, paid: actualPaid, unpaid: unpaid, status: status
      });
      totalAR += total; totalPaid += actualPaid; totalUnpaid += unpaid;
    }
    return { summary: { total: totalAR, paid: totalPaid, unpaid: totalUnpaid }, list: list.sort(function(a, b) { return new Date(b.date) - new Date(a.date); }) };
  } catch (e) { return { summary: {}, list: [] }; }
}

function saveExpense(data) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: "系統忙碌" };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES);
    var expId = generateId("EXP", CONFIG.SHEETS.EXPENSES, 1);
    sheet.appendRow([
      expId, data.date, data.category, data.item, Number(data.amount) || 0,
      data.paymentMethod, data.recipient || "", data.invoiceNo || "", "", "已支出", data.note || ""
    ]);
    if (data.paymentMethod === '零用金') {
      addPettyCashLog({
        type: '出金', date: data.date, amount: data.amount,
        summary: data.item + " (" + data.category + ")", handler: Session.getActiveUser().getEmail(),
        note: "關聯支出單: " + expId
      });
    }
    return { success: true, expId: expId };
  } catch (e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getExpenseList(monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var dStr = formatDate(data[i][1]);
      if (dStr.indexOf(monthStr) > -1) {
        results.push({
          expId: String(data[i][0]), date: dStr, category: data[i][2],
          item: data[i][3], amount: data[i][4], paymentMethod: data[i][5],
          payee: data[i][6], invoiceNo: data[i][7]
        });
      }
    }
    return results.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });
  } catch (e) { return []; }
}

// 輔助函式：寫入零用金紀錄
function addPettyCashLog(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PETTY_CASH); // 請確認 CONFIG 裡有 PETTY_CASH: "零用金紀錄"
  
  if (!sheet) return; // 找不到表就不做

  var lastRow = sheet.getLastRow();
  // 取得目前結餘 (假設在第 H 欄 / 第 8 欄)
  var currentBalance = 0;
  if (lastRow > 1) {
    currentBalance = Number(sheet.getRange(lastRow, 8).getValue()) || 0;
  }

  var amount = Number(data.amount) || 0;
  var newBalance = (data.type === '入金') ? currentBalance + amount : currentBalance - amount;
  var pcId = generateId("PC", CONFIG.SHEETS.PETTY_CASH, 1);

  sheet.appendRow([
    pcId, 
    data.date, 
    data.type, 
    amount, 
    data.summary, 
    data.handler, 
    data.note, 
    newBalance
  ]);
  
  return { success: true };
}
function getPettyCashData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PETTY_CASH);
    if (!sheet) return { success: false, error: "無資料" };
    var data = sheet.getDataRange().getValues();
    var logs = [];
    var limit = 20;
    var start = Math.max(1, data.length - limit);
    for (var i = data.length - 1; i >= start; i--) {
      logs.push({
        id: data[i][0], date: formatDate(data[i][1]), type: data[i][2],
        amount: data[i][3], summary: data[i][4], handler: data[i][5],
        note: data[i][6], balance: data[i][7]
      });
    }
    var balance = data.length > 1 ? data[data.length - 1][7] : 0;
    return { success: true, currentBalance: balance, logs: logs };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function saveDeposit(data) {
  return addPettyCashLog({
    type: '入金', date: data.date, amount: data.amount,
    summary: data.source, handler: Session.getActiveUser().getEmail(), note: data.note
  });
}

function addEmployeeMeal(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEE_MEALS);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.EMPLOYEE_MEALS);
      sheet.appendRow(['編號', '日期', '用餐人數', '金額', '付款方式', '備註']);
    }
    var mealId = generateId("MF", CONFIG.SHEETS.EMPLOYEE_MEALS, 1);
    sheet.appendRow([mealId, data.date, Number(data.headcount) || 0, Number(data.amount) || 0, data.paymentMethod || '現金', data.note || '']);
    return { success: true, mealId: mealId };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getEmployeeMeals(monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEE_MEALS);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var d = formatDate(data[i][1]);
      if (d.indexOf(monthStr) > -1) {
        results.push({
          id: data[i][0], date: d, headcount: data[i][2], amount: data[i][3],
          paymentMethod: data[i][4], note: data[i][5]
        });
      }
    }
    return results.sort(function(a, b) { return a.date.localeCompare(b.date); });
  } catch (e) { return []; }
}

/**
 * 💸 儲存供應商付款 (全面連動版：包含應付帳款沖帳)
 */
function savePaymentMade(data) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: "系統忙碌" };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var paySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_MADE); // 付款紀錄
    var purSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE);     // 進貨單
    var apSheet = ss.getSheetByName(CONFIG.SHEETS.PAYABLE);       // 應付帳款 (應付帳款)
    var pcSheet = ss.getSheetByName(CONFIG.SHEETS.PETTY_CASH);    // 零用金
    var expSheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES);     // 現金支出

    var nowStr = getNowString();
    var amount = Number(data.amount);

    // 1. 寫入付款紀錄
    var suppliers = getSuppliers();
    var supp = suppliers.find(function(s) { return s.id === data.supplierId; });
    var payId = generateId("PM", CONFIG.SHEETS.PAYMENT_MADE, 1);
    
    paySheet.appendRow([
      payId, data.date, data.supplierId, (supp ? supp.name : ""),
      amount, data.paymentMethod, "", data.orderId || "未指定", 
      nowStr, data.note || ""
    ]);

    // 2. 更新「進貨單」狀態 (讓進貨單查詢畫面變更為已付款)
    if (purSheet && data.orderId) {
      var purData = purSheet.getDataRange().getValues();
      for (var i = 1; i < purData.length; i++) {
        if (String(purData[i][0]) === String(data.orderId)) {
          purSheet.getRange(i + 1, 9).setValue('已付款'); // I欄變更狀態
          break;
        }
      }
    }

    var amount = Math.round(Number(data.amount)); // 🔥 強制轉為整數

    // 3. 🔥🔥 核心修正：更新「應付帳款」表格 (沖帳) 🔥🔥
    if (apSheet && data.orderId) {
      var apData = apSheet.getDataRange().getValues();
      for (var j = 1; j < apData.length; j++) {
        // D欄是進貨單號 (Index 3)
        if (String(apData[j][3]) === String(data.orderId)) {
          var totalDue = Number(apData[j][5]) || 0; // F欄: 應付總額
          var oldPaid = Number(apData[j][6]) || 0;  // G欄: 已付金額
          
          var newPaid = oldPaid + amount;
          var newUnpaid = totalDue - newPaid;
          var newStatus = (newUnpaid <= 0) ? '已付款' : '部分付款';

          // 更新工作表對應儲存格
          apSheet.getRange(j + 1, 7).setValue(newPaid);      // G欄: 已付
          apSheet.getRange(j + 1, 8).setValue(newUnpaid);    // H欄: 未付
          apSheet.getRange(j + 1, 9).setValue(data.date);    // I欄: 最後付款日
          apSheet.getRange(j + 1, 11).setValue(newStatus);   // K欄: 狀態
          break;
        }
      }
    }

    // 4. 連動零用金與支出 (維持原邏輯)
    if (data.paymentMethod.indexOf('零用金') > -1 && pcSheet) {
        var lastBal = 0;
        if (pcSheet.getLastRow() > 1) lastBal = Number(pcSheet.getRange(pcSheet.getLastRow(), 8).getValue()) || 0;
        pcSheet.appendRow([generateId("PC", CONFIG.SHEETS.PETTY_CASH, 1), data.date, '出金', amount, '支付貨款-' + (supp ? supp.name : ""), '', '單號:'+data.orderId, lastBal - amount]);
    }
    if (expSheet) {
      expSheet.appendRow([generateId("EXP", CONFIG.SHEETS.EXPENSES, 1), data.date, '進貨成本', '支付貨款-' + (supp ? supp.name : ""), amount, data.paymentMethod, (supp ? supp.name : ""), '', '', '已支出', '單號:'+data.orderId]);
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getSupplierPayables(supplierId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var purSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE);
    var paySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_MADE);
    if (!purSheet) return [];
    var paidMap = {};
    if (paySheet) {
      var payData = paySheet.getDataRange().getValues();
      for (var j = 1; j < payData.length; j++) {
        var oid = String(payData[j][7]);
        var amt = Number(payData[j][4]) || 0;
        paidMap[oid] = (paidMap[oid] || 0) + amt;
      }
    }
    var results = [];
    var purData = purSheet.getDataRange().getValues();
    for (var i = 1; i < purData.length; i++) {
      if (String(purData[i][2]) === supplierId) {
        var oid = String(purData[i][0]);
        var total = Number(purData[i][6]) || 0;
        var paid = paidMap[oid] || 0;
        if (String(purData[i][8]) === '已付款' && paid === 0) paid = total;
        var remaining = total - paid;
        var status = '未付款';
        if (remaining <= 0) status = '已付款';
        else if (paid > 0) status = '部分付款';
        if (remaining > 0) {
          results.push({
            orderId: oid, date: formatDate(purData[i][1]), totalAmount: total,
            paidAmount: paid, remaining: remaining, status: status
          });
        }
      }
    }
    return results.sort(function(a, b) { return new Date(a.date) - new Date(b.date); });
  } catch (e) { return []; }
}

function getPayableReport(dateFrom, dateTo, statusFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var purSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE);
    var paySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_MADE);
    if (!purSheet) return [];
    var purData = purSheet.getDataRange().getValues();
    var payData = paySheet ? paySheet.getDataRange().getValues() : [];
    var paidMap = {};
    for (var j = 1; j < payData.length; j++) paidMap[String(payData[j][7])] = (paidMap[String(payData[j][7])] || 0) + Number(payData[j][4]);
    var results = [];
    var from = dateFrom ? new Date(dateFrom) : null;
    var to = dateTo ? new Date(dateTo) : null;
    if (to) to.setHours(23, 59, 59);
    for (var i = 1; i < purData.length; i++) {
      var d = new Date(purData[i][1]);
      if (from && d < from) continue;
      if (to && d > to) continue;
      var oid = String(purData[i][0]);
      var total = Number(purData[i][6]) || 0;
      var paid = paidMap[oid] || 0;
      var st = String(purData[i][8]);
      if (st === '已付款' && paid === 0) paid = total;
      var unpaid = total - paid;
      if (statusFilter && statusFilter !== st) continue;
      results.push({
        purchaseId: oid, date: formatDate(d), supplierName: String(purData[i][3]),
        totalAmount: total, paidAmount: paid, unpaidAmount: unpaid, status: st
      });
    }
    return results.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });
  } catch (e) { return []; }
}

function saveInvoiceData(data, type) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = (type === 'RECEIVED') ? CONFIG.SHEETS.INVOICES_RECEIVED : CONFIG.SHEETS.INVOICES_ISSUED;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("找不到工作表");
    sheet.appendRow([
      getNowString(), data.date, data.invoiceNo, data.targetName,
      data.taxId, data.taxType, Number(data.amount) || 0, Number(data.tax) || 0,
      Number(data.total) || 0, data.note, data.refNo
    ]);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ==========================================
// 📊 報表專用後端 (日期與金額強力修復版)
// ==========================================
function getDetailedMonthlyData(monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 🛠️ 日期比對工具
    var checkMonth = function(dateVal) {
      if (!dateVal) return false;
      var str = "";
      if (dateVal instanceof Date) {
        str = Utilities.formatDate(dateVal, "GMT+8", "yyyy-MM");
      } else {
        str = String(dateVal).trim().substring(0, 7).replace(/\//g, "-");
      }
      return str === monthStr;
    };

    // 🛠️ 日期格式化工具
    var fmtDate = function(dateVal) {
      if (dateVal instanceof Date) {
        return Utilities.formatDate(dateVal, "GMT+8", "MM-dd");
      }
      return String(dateVal).substring(5, 10);
    };

    // ==========================================
    // ✅ 修正：使用前端期待的資料結構
    // ==========================================
    var result = {
      success: true,
      customerList: [],           // ✅ 改這裡
      suppliersByType: {          // ✅ 改這裡
        '進貨供應商': [],
        '公用事業': [],
        '房東': [],
        '餐飲': [],
        '其他': []
      },
      mealList: [],               // ✅ 改這裡
      expenseList: [],            // ✅ 改這裡
      totals: {
        sales: 0,
        salesPaid: 0,
        salesUnpaid: 0,
        purchase: 0,
        purchasePaid: 0,
        purchaseUnpaid: 0,
        expense: 0,
        meals: 0
      }
    };

    // ==========================================
    // 1. 📗 銷貨收入
    // ==========================================
    var arSheet = ss.getSheetByName(CONFIG.SHEETS.RECEIVABLE);
    if (arSheet) {
      var arData = arSheet.getDataRange().getValues();
      var custMap = {};
      
      for (var i = 1; i < arData.length; i++) {
        if (checkMonth(arData[i][4])) { // E欄：訂單日期
          var custName = String(arData[i][2]);
          var total = Number(arData[i][5]) || 0;
          var paid = Number(arData[i][6]) || 0;
          var unpaid = Number(arData[i][7]) || 0;
          
          if (!custMap[custName]) {
            custMap[custName] = { name: custName, total: 0, paid: 0, unpaid: 0 };
          }
          custMap[custName].total += total;
          custMap[custName].paid += paid;
          custMap[custName].unpaid += unpaid;
          
          result.totals.sales += total;
          result.totals.salesPaid += paid;
          result.totals.salesUnpaid += unpaid;
        }
      }
      
      for (var name in custMap) {
        result.customerList.push(custMap[name]);
      }
      result.customerList.sort(function(a, b) { return b.total - a.total; });
    }

    // ==========================================
// 2. 📕 進貨與應付
// ==========================================
var purSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE);
var suppSheet = ss.getSheetByName(CONFIG.SHEETS.SUPPLIERS);
var payMadeSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_MADE); // ⭐ 新增：讀取付款記錄

// 建立供應商類型對照表
var suppTypeMap = {};
if (suppSheet) {
  var suppData = suppSheet.getDataRange().getValues();
  for (var i = 1; i < suppData.length; i++) {
    suppTypeMap[String(suppData[i][0])] = String(suppData[i][14] || '進貨供應商');
  }
}

// ⭐ 建立實際付款記錄對照表
var paidMap = {};
if (payMadeSheet) {
  var payData = payMadeSheet.getDataRange().getValues();
  for (var j = 1; j < payData.length; j++) {
    var oid = String(payData[j][7]); // H欄是進貨單號
    var amt = Number(payData[j][4]) || 0; // E欄是金額
    paidMap[oid] = (paidMap[oid] || 0) + amt;
  }
}

if (purSheet) {
  var purData = purSheet.getDataRange().getValues();
  var suppMap = {};
  
  for (var i = 1; i < purData.length; i++) {
    if (checkMonth(purData[i][1])) {
      var suppId = String(purData[i][2]);
      var suppName = String(purData[i][3]);
      var orderId = String(purData[i][0]); // ⭐ 新增：讀取進貨單號
      var total = Number(purData[i][6]) || 0;
      var status = String(purData[i][8]);
      
      // ⭐ 修正：從付款記錄查詢實際付款金額
      var paid = paidMap[orderId] || 0;
      
      // 如果狀態是「已付款」但沒有付款記錄，才用總額
      if (status === '已付款' && paid === 0) {
        paid = total;
      }
      
      var unpaid = total - paid;
      
      if (!suppMap[suppName]) {
        suppMap[suppName] = { 
          name: suppName, 
          type: suppTypeMap[suppId] || '進貨供應商',
          total: 0, 
          paid: 0, 
          unpaid: 0 
        };
      }
      suppMap[suppName].total += total;
      suppMap[suppName].paid += paid;      // ✅ 累加已付金額
      suppMap[suppName].unpaid += unpaid;
      
      result.totals.purchase += total;
      result.totals.purchasePaid += paid;  // ✅ 累加到總計
      result.totals.purchaseUnpaid += unpaid;
    }
  }
  
  // 將供應商分類
  for (var name in suppMap) {
    var s = suppMap[name];
    var type = s.type || '進貨供應商';
    if (!result.suppliersByType[type]) {
      result.suppliersByType[type] = [];
    }
    result.suppliersByType[type].push(s);
  }
}

    // ==========================================
    // 3. 🍱 員工餐費
    // ==========================================
    var mealSheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEE_MEALS);
    if (mealSheet) {
      var mealData = mealSheet.getDataRange().getValues();
      for (var i = 1; i < mealData.length; i++) {
        if (checkMonth(mealData[i][1])) { // B欄：日期
          var amt = Number(mealData[i][3]) || 0;
          result.mealList.push({
            date: fmtDate(mealData[i][1]),
            headcount: mealData[i][2],
            amount: amt,
            paymentMethod: String(mealData[i][4] || ''),
            note: String(mealData[i][5] || '')
          });
          result.totals.meals += amt;
        }
      }
    }

    // ==========================================
    // 4. 💸 現金支出
    // ==========================================
    var expSheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES);
    if (expSheet) {
      var expData = expSheet.getDataRange().getValues();
      for (var i = 1; i < expData.length; i++) {
        if (checkMonth(expData[i][1])) { // B欄：日期
          var amt = Number(expData[i][4]) || 0;
          result.expenseList.push({
            date: fmtDate(expData[i][1]),
            category: String(expData[i][2]),
            item: String(expData[i][3]),
            amount: amt,
            payee: String(expData[i][6] || '')
          });
          result.totals.expense += amt;
        }
      }
    }

    return result;

  } catch (e) {
    Logger.log("❌ getDetailedMonthlyData 錯誤: " + e.toString());
    return { 
      success: false, 
      error: "讀取失敗: " + e.toString(),
      customerList: [],
      suppliersByType: {},
      mealList: [],
      expenseList: [],
      totals: {
        sales: 0, salesPaid: 0, salesUnpaid: 0,
        purchase: 0, purchasePaid: 0, purchaseUnpaid: 0,
        expense: 0, meals: 0
      }
    };
  }
}

function getCustomerMonthlyReport(customerId, dateFrom, dateTo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
    var detailSheet = ss.getSheetByName(CONFIG.SHEETS.SALES_DETAILS);
    var customerSheet = ss.getSheetByName(CONFIG.SHEETS.CUSTOMERS);
    var custData = customerSheet.getDataRange().getValues();
    var custInfo = { name: customerId, phone: '', address: '' };
    for (var i = 1; i < custData.length; i++) {
      if (String(custData[i][0]) === customerId) {
        custInfo.name = custData[i][1];
        custInfo.phone = custData[i][4] || custData[i][3];
        custInfo.address = custData[i][5];
        break;
      }
    }
    var orders = [], summary = { totalOrders: 0, totalSales: 0, totalPaid: 0, totalUnpaid: 0, overdueAmount: 0 };
    var from = new Date(dateFrom), to = new Date(dateTo);
    to.setHours(23, 59, 59);
    var salesData = salesSheet.getDataRange().getValues();
    var detailData = detailSheet.getDataRange().getValues();
    for (var i = 1; i < salesData.length; i++) {
      if (String(salesData[i][2]) !== customerId) continue;
      var d = new Date(salesData[i][1]);
      if (d < from || d > to) continue;
      var oid = String(salesData[i][0]);
      var grandTotal = Number(salesData[i][8]) || 0;
      var status = String(salesData[i][10]);
      var myItems = [];
      for (var j = 1; j < detailData.length; j++) {
        if (String(detailData[j][1]) === oid) {
          myItems.push({ productName: detailData[j][3], qty: detailData[j][4], unit: detailData[j][5], price: detailData[j][6], subtotal: detailData[j][7] });
        }
      }
      orders.push({
        orderId: oid, date: formatDate(d), subtotal: Number(salesData[i][4]) || 0,
        shipping: Number(salesData[i][5]) || 0, deduction: Number(salesData[i][6]) || 0,
        taxAmount: Number(salesData[i][7]) || 0, grandTotal: grandTotal, paymentStatus: status, items: myItems
      });
      summary.totalOrders++; summary.totalSales += grandTotal;
      if (status === '已收款') summary.totalPaid += grandTotal;
      else { summary.totalUnpaid += grandTotal; if (d < new Date()) summary.overdueAmount += grandTotal; }
    }
    orders.sort(function(a, b) { return a.date.localeCompare(b.date); });
    return { success: true, customer: custInfo, dateRange: { from: dateFrom, to: dateTo }, orders: orders, summary: summary };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function analyzeCustomers(monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
    var detailSheet = ss.getSheetByName(CONFIG.SHEETS.SALES_DETAILS);
    if (!salesSheet || !detailSheet) return [];
    var stats = {}, validOrderIds = {};
    var salesData = salesSheet.getDataRange().getValues();
    for (var i = 1; i < salesData.length; i++) {
      if (formatDate(salesData[i][1]).substring(0, 7) === monthStr) {
        var oid = String(salesData[i][0]);
        var cname = String(salesData[i][3]);
        var total = Number(salesData[i][8]) || 0;
        if (!stats[cname]) stats[cname] = { revenue: 0, profit: 0, count: 0 };
        stats[cname].revenue += total; stats[cname].count += 1;
        validOrderIds[oid] = cname;
      }
    }
    var detailData = detailSheet.getDataRange().getValues();
    for (var j = 1; j < detailData.length; j++) {
      var cname = validOrderIds[String(detailData[j][1])];
      if (cname) stats[cname].profit += (Number(detailData[j][9]) || 0);
    }
    var result = [];
    for (var name in stats) {
      var s = stats[name];
      var rate = s.revenue > 0 ? ((s.profit / s.revenue) * 100).toFixed(1) : 0;
      result.push({ customerName: name, revenue: s.revenue, profit: s.profit, profitRate: rate, orderCount: s.count });
    }
    result.sort(function(a, b) { return b.revenue - a.revenue; });
    return result;
  } catch (e) { return []; }
}

function showCustomerAnalysisPanel() { createDialog('Customeranalysis', 600, 700, '👥 客戶分析'); }

function prepareInvoiceFromExpense(expId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES);
    var data = sheet.getDataRange().getValues();
    var targetRow = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === expId) { targetRow = data[i]; break; }
    }
    if (!targetRow) throw new Error("找不到支出單：" + expId);
    var invoiceData = {
      source: 'EXPENSE', refNo: expId, date: formatDate(targetRow[1]),
      note: String(targetRow[3]) + " (" + String(targetRow[2]) + ")", amount: Number(targetRow[4]),
      taxType: '內含5%', targetName: String(targetRow[6] || '')
    };
    PropertiesService.getUserProperties().setProperty('TEMP_INVOICE_DATA', JSON.stringify(invoiceData));
    return { success: true };
  } catch (e) { throw new Error(e.toString()); }
}

function getTempInvoiceData() {
  try {
    var props = PropertiesService.getUserProperties();
    var json = props.getProperty('TEMP_INVOICE_DATA');
    if (json) { props.deleteProperty('TEMP_INVOICE_DATA'); return JSON.parse(json); }
    return null;
  } catch (e) { return null; }
}

function getMonthlyReceivableSummary(monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var arSheet = ss.getSheetByName(CONFIG.SHEETS.RECEIVABLE);
    if (!arSheet) return { success: false, error: "找不到應收帳款表" };
    var arData = arSheet.getDataRange().getValues();
    var summaryMap = {};
    for (var i = 1; i < arData.length; i++) {
      if (formatDate(arData[i][4]).substring(0, 7) === monthStr) {
        var custName = String(arData[i][2]);
        var total = Number(arData[i][5]) || 0;
        var paid = Number(arData[i][6]) || 0;
        var unpaid = Number(arData[i][7]) || 0;
        if (!summaryMap[custName]) summaryMap[custName] = { sales: 0, paid: 0, unpaid: 0 };
        summaryMap[custName].sales += total; summaryMap[custName].paid += paid; summaryMap[custName].unpaid += unpaid;
      }
    }
    var resultList = [];
    for (var name in summaryMap) {
      resultList.push({ custName: name, sales: summaryMap[name].sales, paid: summaryMap[name].paid, unpaid: summaryMap[name].unpaid });
    }
    resultList.sort(function(a, b) { return b.unpaid - a.unpaid; });
    return { success: true, month: monthStr, data: resultList };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getDetailForPopup(orderId, type) {
  try { return (type === 'PURCHASE') ? getPurchaseOrderDetail(orderId) : getSalesOrderDetail(orderId); }
  catch (e) { return { success: false, error: e.toString() }; }
}

function showOrderDetailPopup(orderId, type) {
  var template = HtmlService.createTemplateFromFile('OrderDetailPopup');
  template.orderId = orderId;
  template.type = type;
  var html = template.evaluate().setWidth(600).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '📄 單據明細：' + orderId);
}

function getMonthlyPayableSummary(monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var purSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE);
    var paySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_MADE);
    if (!purSheet) return { success: false, error: "找不到進貨單" };
    var paidMap = {};
    if (paySheet) {
      var payData = paySheet.getDataRange().getValues();
      for (var j = 1; j < payData.length; j++) paidMap[String(payData[j][7])] = (paidMap[String(payData[j][7])] || 0) + (Number(payData[j][4]) || 0);
    }
    var summaryMap = {};
    var purData = purSheet.getDataRange().getValues();
    for (var i = 1; i < purData.length; i++) {
      if (formatDate(purData[i][1]).substring(0, 7) === monthStr) {
        var oid = String(purData[i][0]);
        var suppName = String(purData[i][3]);
        var total = Number(purData[i][6]) || 0;
        var paid = paidMap[oid] || 0;
        if (String(purData[i][8]) === '已付款' && paid === 0) paid = total;
        var unpaid = total - paid;
        if (!summaryMap[suppName]) summaryMap[suppName] = { purchase: 0, paid: 0, unpaid: 0 };
        summaryMap[suppName].purchase += total; summaryMap[suppName].paid += paid; summaryMap[suppName].unpaid += unpaid;
      }
    }
    var resultList = [];
    for (var name in summaryMap) {
      resultList.push({ supplierName: name, purchase: summaryMap[name].purchase, paid: summaryMap[name].paid, unpaid: summaryMap[name].unpaid });
    }
    resultList.sort(function(a, b) { return b.unpaid - a.unpaid; });
    return { success: true, month: monthStr, data: resultList };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function showPurchaseOrderDetailPopup(orderId) { showOrderDetailPopup(orderId, 'PURCHASE'); }

/**
 * 📊 商品分析 (支援日期範圍 + 類別)
 */
function analyzeProducts(dateFrom, dateTo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
    var detailSheet = ss.getSheetByName(CONFIG.SHEETS.SALES_DETAILS);
    var productSheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS); // 新增讀取商品主檔
    
    if (!salesSheet || !detailSheet) return [];

    // 1. 先建立商品類別對照表 (ID -> Category)
    var categoryMap = {};
    if (productSheet) {
      var pData = productSheet.getDataRange().getValues();
      // 從第 2 列開始讀
      for (var k = 1; k < pData.length; k++) {
        var pId = String(pData[k][0]); // A欄: ID
        var pCat = String(pData[k][1]); // B欄: 類別
        categoryMap[pId] = pCat;
      }
    }

    // 2. 處理日期範圍
    var fromDate = new Date(dateFrom);
    fromDate.setHours(0, 0, 0, 0);
    
    var toDate = new Date(dateTo);
    toDate.setHours(23, 59, 59, 999);

    // 3. 篩選符合日期的訂單 ID
    var validOrderIds = {};
    var salesData = salesSheet.getDataRange().getValues();
    
    for (var i = 1; i < salesData.length; i++) {
      var orderDate = new Date(salesData[i][1]); 
      if (orderDate >= fromDate && orderDate <= toDate) {
        validOrderIds[String(salesData[i][0])] = true; 
      }
    }

    // 4. 統計數據
    var stats = {};
    var detailData = detailSheet.getDataRange().getValues();
    
    for (var j = 1; j < detailData.length; j++) {
      var orderId = String(detailData[j][1]); 
      
      if (validOrderIds[orderId]) {
        var pid = String(detailData[j][2]);     
        var pname = String(detailData[j][3]);   
        var qty = Number(detailData[j][4]) || 0;       
        var unit = String(detailData[j][5]);           
        var revenue = Number(detailData[j][7]) || 0;   
        var profit = Number(detailData[j][9]) || 0;    

        if (!stats[pid]) {
          stats[pid] = { 
            id: pid, 
            category: categoryMap[pid] || "未分類", // 這裡加入類別
            name: pname, 
            qty: 0, 
            unit: unit, 
            revenue: 0, 
            profit: 0 
          };
        }
        stats[pid].qty += qty;
        stats[pid].revenue += revenue;
        stats[pid].profit += profit;
      }
    }

    // 5. 轉換為陣列
    var result = [];
    for (var pid in stats) {
      var s = stats[pid];
      var rate = s.revenue > 0 ? ((s.profit / s.revenue) * 100).toFixed(1) : 0;
      
      result.push({
        productId: s.id,
        category: s.category, // 回傳類別
        productName: s.name,
        qty: s.qty,
        unit: s.unit,
        revenue: s.revenue,
        profit: s.profit,
        profitRate: rate
      });
    }

    result.sort(function(a, b) { return b.revenue - a.revenue; });
    
    return result;

  } catch (e) {
    return []; 
  }
}

function showProductAnalysisPanel() { createDialog('Productanalysis', 600, 700, '📦 商品分析'); }


function getProfitData(dateFrom, dateTo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SALES_DETAILS);
    var salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
    if (!sheet || !salesSheet) return { kpi: {}, charts: {}, details: [] };
    
    var data = sheet.getDataRange().getValues();
    var salesData = salesSheet.getDataRange().getValues();
    var from = new Date(dateFrom).getTime();
    var toDateEnd = new Date(dateTo); toDateEnd.setHours(23, 59, 59, 999); var toEnd = toDateEnd.getTime();
    
    var orderDateMap = {};
    for (var s = 1; s < salesData.length; s++) {
      var d = new Date(salesData[s][1]);
      if (d.getTime() >= from && d.getTime() <= toEnd) orderDateMap[String(salesData[s][0])] = formatDate(salesData[s][1]);
    }
    
    var totalRevenue = 0, totalProfit = 0, dailyMap = {}, productMap = {}, details = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var orderId = String(row[1]);
      if (!orderDateMap[orderId]) continue;
      var date = orderDateMap[orderId];
      var prodName = String(row[3]);
      var qty = Number(row[4]);
      var subtotal = Number(row[7]);
      var profit = Number(row[9]);
      if (!profit && profit !== 0) { profit = (Number(row[6]) - Number(row[8])) * qty; }
      
      totalRevenue += subtotal; totalProfit += profit;
      if (!dailyMap[date]) dailyMap[date] = { r: 0, p: 0 };
      dailyMap[date].r += subtotal; dailyMap[date].p += profit;
      if (!productMap[prodName]) productMap[prodName] = 0;
      productMap[prodName] += profit;
      details.push({ date: date, product: prodName, qty: qty, revenue: subtotal, cost: subtotal - profit, profit: profit, margin: subtotal > 0 ? ((profit/subtotal)*100).toFixed(1) : 0 });
    }
    
    var trendChart = []; Object.keys(dailyMap).sort().forEach(function(d) { trendChart.push([d, dailyMap[d].r, dailyMap[d].p]); });
    var prodChart = []; Object.keys(productMap).map(function(p) { return [p, productMap[p]]; }).sort(function(a, b) { return b[1] - a[1]; }).slice(0, 5).forEach(function(item) { prodChart.push(item); });
    
    return {
      kpi: { revenue: totalRevenue, profit: totalProfit, margin: totalRevenue > 0 ? ((totalProfit/totalRevenue)*100).toFixed(1) : 0 },
      charts: { trend: trendChart, products: prodChart },
      details: details.sort(function(a, b) { return b.date.localeCompare(a.date); })
    };
  } catch (e) { return { error: e.toString() }; }
}

function showProfitAnalysisPanel() { createDialog('ProfitAnalysis', 1000, 800, '💰 利潤分析'); }

function showInvoiceIssuedForm() { createDialog('InvoiceForm', 650, 700, '📤 紀錄銷項發票'); }
function showInvoiceReceivedForm() { createDialog('InvoiceForm', 650, 700, '📥 紀錄進項發票'); }
function showCalendarPanel() { createDialog('Calendar', 900, 700, '📅 行事曆'); }
function showDetailedMonthlyReportPanel() { createDialog('DetailedReport', 1000, 800, '📅 詳細月結報表'); }
function showMonthlyReportPanel() { createDialog('MonthlyReport', 850, 800, '📅 月結報表'); }
function showCustomerMonthlyReport() { createDialog('CustomerMonthlyReport', 900, 700, '📊 客戶月結報表'); }
function showProductForm() { createDialog('ProductForm', 1000, 800, '➕ 新增商品'); }
function showCustomerManager() { createDialog('CustomerManager', 800, 750, '👥 客戶管理'); }
function showSupplierManager() { createDialog('SupplierManager', 800, 750, '🏢 供應商管理'); }
function showPriceForm() { createDialog('PriceForm', 850, 800, '🏷️ 設定客戶特價'); }
function processOrderAction(orderId, action) {
  var detail = getSalesOrderDetail(orderId);
  if (!detail.success) throw detail.error;
  var orderData = detail.order;
  if (action === 'EDIT') { orderData.isEditMode = true; orderData.oldOrderId = orderId; }
  else { orderData.isEditMode = false; delete orderData.orderId; }
  PropertiesService.getUserProperties().setProperty('COPY_ORDER_DATA', JSON.stringify(orderData));
  showSalesOrderPanel(); 
}
function getCopyOrderData() {
  var p = PropertiesService.getUserProperties();
  var d = p.getProperty('COPY_ORDER_DATA');
  if (d) { p.deleteProperty('COPY_ORDER_DATA'); return { success: true, data: JSON.parse(d) }; }
  return { success: false };
}
function processPurchaseOrderAction(orderId, action) {
  var res = getPurchaseOrderDetail(orderId);
  if (!res.success) throw new Error("找不到進貨單");
  var data = res.order;
  data.isEditMode = true;
  data.oldOrderId = orderId;
  PropertiesService.getUserProperties().setProperty("COPY_PURCHASE_ORDER", JSON.stringify(data));
  showPurchaseOrderPanel();
}
function protectAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var protectedCount = 0;
  protectedCount += protectSheet(ss, CONFIG.SHEETS.PRODUCTS, 12, '商品資料');
  protectedCount += protectSheet(ss, CONFIG.SHEETS.CUSTOMERS, 15, '客戶資料');
  protectedCount += protectSheet(ss, CONFIG.SHEETS.SUPPLIERS, 14, '供應商資料');
  protectedCount += protectSheet(ss, CONFIG.SHEETS.SALES, 15, '銷貨記錄');
  protectedCount += protectSheet(ss, CONFIG.SHEETS.SALES_DETAILS, 11, '銷貨明細');
  protectedCount += protectSheet(ss, CONFIG.SHEETS.PURCHASE, 10, '進貨記錄');
  protectedCount += protectSheet(ss, CONFIG.SHEETS.PURCHASE_DETAILS, 8, '進貨明細');
  protectedCount += protectSheet(ss, CONFIG.SHEETS.RECEIVABLE, 10, '應收帳款');
  protectedCount += protectSheet(ss, CONFIG.SHEETS.PAYABLE, 10, '應付帳款');
  protectedCount += protectSheet(ss, CONFIG.SHEETS.EXPENSES, 9, '支出記錄');
  return protectedCount;
}
function protectSheet(ss, sheetName, colCount, desc) {
  try {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return 0;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;
    var range = sheet.getRange(2, 1, lastRow - 1, colCount);
    var p = range.protect().setDescription('🔒 ' + desc + ' - 已鎖定');
    p.setWarningOnly(true);
    return 1;
  } catch (e) { return 0; }
}
function createBackup() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var name = '【備份】丸十庫存_' + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmmss");
    var backup = ss.copy(name);
    try {
      var iter = DriveApp.getFoldersByName('系統備份');
      var folder = iter.hasNext() ? iter.next() : DriveApp.createFolder('系統備份');
      DriveApp.getFileById(backup.getId()).moveTo(folder);
    } catch(e) {}
    return { success: true, name: name, url: backup.getUrl() };
  } catch(e) { return { success: false, error: e.toString() }; }
}
function initializeDataProtection() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('🛡️ 初始化保護', '將設定備份與鎖定歷史資料，確定嗎？', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;
  createBackup();
  protectAllData();
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){ if(t.getHandlerFunction()==='createBackup') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('createBackup').timeBased().atHour(2).everyDays(1).create();
  ui.alert('✅ 保護設定完成，已建立備份並鎖定歷史資料。');
}
function createDialog(filename, width, height, title) {
  var html = HtmlService.createHtmlOutputFromFile(filename).setWidth(width).setHeight(height);
  SpreadsheetApp.getUi().showModelessDialog(html, title);
}
// 1. 新增行事曆事件
function addCalendarEvent(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.CALENDAR);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.CALENDAR);
      sheet.appendRow(['EventID', 'Date', 'Time', 'Title', 'Type', 'Content', 'Status', 'CreatedAt']);
    }
    
    var eventId = generateId("EVT", CONFIG.SHEETS.CALENDAR, 1);
    var nowStr = getNowString();
    
    sheet.appendRow([
      eventId, 
      data.date,          // B欄: 日期
      data.time || '',    // C欄: 時間
      data.title,         // D欄: 標題
      data.type || '一般', // E欄: 類型
      data.content || '', // F欄: 內容
      '未完成',            // G欄: 狀態
      nowStr              // H欄: 建立時間
    ]);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 2. 取得指定月份的事件
function getMonthEvents(yearMonth) { // yearMonth 格式 "2026-02"
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CALENDAR);
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var results = [];
    
    for (var i = 1; i < data.length; i++) {
      var rowDate = formatDate(data[i][1]); // 確保轉為 YYYY-MM-DD 字串
      
      // 比對月份 (取前7碼)
      if (rowDate.substring(0, 7) === yearMonth) {
        results.push({
          id: String(data[i][0]),
          date: rowDate,
          time: String(data[i][2]),
          title: String(data[i][3]),
          type: String(data[i][4]),
          content: String(data[i][5]),
          status: String(data[i][6])
        });
      }
    }
    return results;
  } catch (e) {
    return [];
  }
}

// 3. 取得今日待辦 (包含行事曆事件 + 預期今日的進銷存單據)
function getTodayEvents() {
  var today = getNowString().split(' ')[0]; // YYYY-MM-DD
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CALENDAR);
    var results = [];
    
    // A. 讀取行事曆中的今日事項
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var rowDate = formatDate(data[i][1]);
        var status = String(data[i][6]);
        
        if (rowDate === today && status !== '已完成') {
          results.push({
            id: String(data[i][0]),
            date: rowDate,
            time: String(data[i][2]),
            title: String(data[i][3]),
            type: String(data[i][4]),
            source: 'calendar'
          });
        }
      }
    }
    
    return results;
  } catch (e) {
    return [];
  }
}

// 4. 更新事件狀態 (完成/刪除)
function updateEventStatus(eventId, newStatus) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CALENDAR);
    if (!sheet) return { success: false };
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === eventId) {
        sheet.getRange(i + 1, 7).setValue(newStatus); // G欄是狀態
        return { success: true };
      }
    }
    return { success: false, error: "找不到事件" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 5. 刪除事件
function deleteCalendarEvent(eventId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CALENDAR);
    if (!sheet) return { success: false };
    deleteRowsById(sheet, 1, eventId); // 1 = A欄是 ID
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 6. 取得付款/收款提醒 (Payment Reminders)
function getPaymentReminders() {
  try {
    var reminders = [];
    var today = new Date();
    today.setHours(0,0,0,0);
    
    // A. 檢查應收帳款 (Receivable)
    var arSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.RECEIVABLE);
    if (arSheet) {
      var arData = arSheet.getDataRange().getValues();
      for (var i = 1; i < arData.length; i++) {
        var status = String(arData[i][10]); // K欄: 狀態
        if (status === '已收款') continue;
        
        var dateStr = formatDate(arData[i][8]); // I欄: 預計收款日 (若無則抓 E欄 訂單日)
        if (!dateStr) dateStr = formatDate(arData[i][4]);
        
        var dueDate = new Date(dateStr);
        var diffTime = dueDate - today;
        var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        
        // 只顯示過期或未來 7 天內的
        if (diffDays <= 7) {
          reminders.push({
            type: '收款',
            title: String(arData[i][2]), // 客戶名稱
            content: '單號: ' + arData[i][3] + ' / 金額: $' + arData[i][7], // 未收金額
            dueDate: dateStr,
            daysLeft: diffDays < 0 ? 0 : diffDays // 過期顯示 0 或負數邏輯
          });
        }
      }
    }
    
    // B. 檢查進貨付款 (Purchase)
    var purSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PURCHASE);
    if (purSheet) {
      var purData = purSheet.getDataRange().getValues();
      for (var i = 1; i < purData.length; i++) {
        var status = String(purData[i][8]); // I欄: 付款狀態
        if (status === '已付款') continue;
        
        var dateStr = formatDate(purData[i][9]); // J欄: 預計付款日
        if (!dateStr) dateStr = formatDate(purData[i][1]); // 若無則抓 B欄 進貨日
        
        var dueDate = new Date(dateStr);
        var diffTime = dueDate - today;
        var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        
        if (diffDays <= 7) {
           reminders.push({
            type: '付款',
            title: String(purData[i][3]), // 供應商名稱
            content: '單號: ' + purData[i][0],
            dueDate: dateStr,
            daysLeft: diffDays < 0 ? 0 : diffDays
          });
        }
      }
    }
    
    // 排序：緊急的 (天數少) 在前
    reminders.sort(function(a, b) { return a.daysLeft - b.daysLeft; });
    return reminders;
    
  } catch (e) {
    return [];
  }
}

// ==========================================
// 15. 📊 Cash Flow Report Backend (收支報表補完)
// ==========================================

/**
 * 取得指定月份的收支統計數據
 * @param {string} monthStr - 格式 "YYYY-MM"
 */
function getCashFlowReport(monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 初始化數據
    var result = {
      success: true,
      salesRevenue: 0,    // 銷貨收入 (已收款)
      cashExpense: 0,     // 現金支出 (一般費用)
      supplierPayment: 0, // 進貨支出 (已付款)
      balance: 0,         // 淨利
      expenses: []        // 支出明細列表
    };

    // 2. 計算銷貨收入 (從收款記錄 PaymentReceived)
    var payRecSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_RECEIVED);
    if (payRecSheet) {
      var data = payRecSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var dateStr = formatDate(data[i][3]); // D欄: 收款日期
        if (dateStr.substring(0, 7) === monthStr) {
          result.salesRevenue += (Number(data[i][4]) || 0); // E欄: 金額
        }
      }
    }

    // 3. 計算進貨支出 (從付款記錄 PaymentMade)
    var payMadeSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_MADE);
    if (payMadeSheet) {
      var data = payMadeSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var dateStr = formatDate(data[i][1]); // B欄: 付款日期
        if (dateStr.substring(0, 7) === monthStr) {
          result.supplierPayment += (Number(data[i][4]) || 0); // E欄: 金額
        }
      }
    }

    // 4. 計算現金支出 (從費用記錄 Expenses)
    var expSheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES);
    if (expSheet) {
      var data = expSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var dateStr = formatDate(data[i][1]); // B欄: 日期
        if (dateStr.substring(0, 7) === monthStr) {
          var amt = (Number(data[i][4]) || 0); // E欄: 金額
          result.cashExpense += amt;
          
          // 加入明細列表
          result.expenses.push({
            date: dateStr,
            category: String(data[i][2]), // C欄: 類別
            item: String(data[i][3]),     // D欄: 項目
            amount: amt,
            payee: String(data[i][6]),    // G欄: 對象
            paymentMethod: String(data[i][5]) // F欄: 支付方式
          });
        }
      }
    }

    // 5. 計算員工餐費支出 (EmployeeMeals)
    var mealSheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEE_MEALS);
    if (mealSheet) {
      var data = mealSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var dateStr = formatDate(data[i][1]); // B欄
        if (dateStr.substring(0, 7) === monthStr) {
          var amt = (Number(data[i][3]) || 0); // D欄
          result.cashExpense += amt;
          
          result.expenses.push({
            date: dateStr,
            category: '員工餐費',
            item: '用餐人數: ' + data[i][2],
            amount: amt,
            payee: '員工',
            paymentMethod: String(data[i][4])
          });
        }
      }
    }

    // 6. 計算結餘
    // 淨利 = 銷貨收入 - (進貨支出 + 現金支出)
    result.balance = result.salesRevenue - (result.supplierPayment + result.cashExpense);

    // 7. 明細排序 (日期新到舊)
    result.expenses.sort(function(a, b) {
      return new Date(b.date) - new Date(a.date);
    });

    return result;

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}


// ==========================================
// 🪟 介面轉發函式 (請貼在 .gs 檔案的最下方)
// ==========================================

// 1. 銷貨與進貨
function showSalesOrderPanel() { createDialog('SalesOrder', 1200, 800, '📝 開立銷貨單'); }
function showSearchSalesPanel() { createDialog('SearchSales', 1200, 800, '🔍 查詢銷貨單'); }
function showPurchaseOrderPanel() { createDialog('PurchaseOrder', 1200, 800, '📝 開立進貨單'); }
function showSearchPurchasePanel() { createDialog('SearchPurchase', 1200, 800, '🔍 查詢進貨單'); }

// 2. 財務與收付款
function showReceivePaymentPanel() { createDialog('ReceivePayment', 850, 800, '💰 客戶收款'); }
function showMakePaymentPanel() { createDialog('MakePayment', 850, 800, '💸 供應商付款'); }
function showReceivableReport() { createDialog('Receivablereport', 850, 800, '📊 應收帳款報表'); }
function showPayableReport() { createDialog('PayableReport', 850, 800, '📊 應付帳款報表'); }

// 3. 庫存與盤點
function showInventoryPanel() { createDialog('InventoryView', 850, 800, '📋 庫存查詢'); }
function showInventoryLogPanel() { createDialog('InventoryLog', 850, 800, '📊 庫存異動記錄'); }
function showStocktakePanel() { createDialog('Stocktake', 850, 800, '✅ 庫存盤點'); }
function showConvertPanel() { createDialog('ConvertStock', 450, 500, '🔄 商品分裝轉換'); }
function showLowStockAlert() { createDialog('Lowstockalert', 500, 600, '⚠️ 庫存預警清單'); }

// 4. 財務紀錄
function showExpensePanel() { createDialog('AddExpense', 850, 800, '💵 記錄現金支出'); }
function showExpenseList() { createDialog('ExpenseList', 850, 800, '📋 支出明細'); }
function showCashFlowReport() { createDialog('CashFlowReport', 850, 800, '📊 收支統計'); }
function showPettyCashPanel() { createDialog('PettyCashPanel', 850, 600, '💰 零用金管理'); }
function showEmployeeMealPanel() { createDialog('EmployeeMeal', 600, 500, '🍱 記錄員工餐費'); }
function showInvoiceIssuedForm() { createDialog('InvoiceForm', 650, 700, '📤 紀錄銷項發票'); }
function showInvoiceReceivedForm() { createDialog('InvoiceForm', 650, 700, '📥 紀錄進項發票'); }

// 5. 報表與設定
function showCalendarPanel() { createDialog('Calendar', 900, 700, '📅 行事曆'); }
function showDetailedMonthlyReportPanel() { createDialog('DetailedReport', 1000, 800, '📅 詳細月結報表'); }
function showMonthlyReportPanel() { createDialog('MonthlyReport', 850, 800, '📅 簡易月結報表'); }
function showCustomerMonthlyReport() { createDialog('CustomerMonthlyReport', 900, 700, '📊 客戶月結報表'); }
function showProductAnalysisPanel() { createDialog('ProductAnalysis', 850, 800, '📦 商品分析'); }
function showProfitAnalysisPanel() { createDialog('ProfitAnalysis', 1000, 800, '💰 利潤分析'); }

// 6. 基礎資料管理
function showProductForm() { createDialog('ProductForm', 750, 700, '➕ 新增商品'); }
function showCustomerManager() { createDialog('CustomerManager', 800, 750, '👥 客戶管理'); }
function showSupplierManager() { createDialog('SupplierManager', 800, 750, '🏢 供應商管理'); }
function showPriceForm() { createDialog('PriceForm', 850, 800, '🏷️ 設定客戶特價'); }






// ========== 原料管理模組 ==========

// 開啟原料管理視窗
function showMaterialManager() { 
  createDialog('MaterialManager', 800, 750, '🍱 原料管理'); 
}

// 讀取「原料主檔」資料表
function getMaterials() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MATERIALS);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // 讀取 A 到 L 欄
    var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues(); 
    return data.map(function(row) {
      return {
        id: String(row[0] || ""),         // A: 原料編號
        name: String(row[1] || ""),       // B: 原料名稱
        spec: String(row[2] || ""),       // C: 規格
        category: String(row[3] || ""),   // D: 類別
        unit: String(row[4] || "個"),     // E: 單位
        cost: Number(row[5]) || 0,        // F: 預設成本
        displayName: "【原料】" + String(row[0]) + " | " + String(row[1])
      };
    }).filter(function(m) { return m.id !== ""; });
  } catch (e) { return []; }
}

// ==========================================
// 🍱 原料管理後端函式
// ==========================================

// 1. 提供畫面初始資料 (供應商清單 + 現有原料)
function getInitialDataForMaterialManager() {
  return {
    suppliers: getSuppliers(), // 沿用你系統裡原本的函式
    materials: getMaterialsFull() // 下面新寫的完整版讀取
  };
}

// 2. 讀取完整原料資料 (含 RowIndex 以便編輯)
function getMaterialsFull() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MATERIALS); // 確保 CONFIG 裡有定義 "原料主檔"
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // 讀取 A ~ L 欄 (共 12 欄)
    var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues(); 
    
    return data.map(function(row, index) {
      return {
        rowIndex: index + 2, // 記錄在 Excel 的第幾列 (方便更新)
        id: String(row[0] || ""),         // A
        name: String(row[1] || ""),       // B
        spec: String(row[2] || ""),       // C
        category: String(row[3] || ""),   // D
        unit: String(row[4] || ""),       // E
        cost: Number(row[5]) || 0,        // F
        supplier: String(row[6] || ""),   // G
        minStock: Number(row[7]) || 0,    // H
        currentStock: Number(row[8]) || 0,// I
        status: String(row[9] || "啟用"), // J
        date: formatDate(row[10]),        // K
        note: String(row[11] || "")       // L
      };
    }).reverse(); // 新的在上面
  } catch (e) { return []; }
}

// 3. 儲存原料 (新增或更新)
// ==========================================
// 🍱 原料管理 - 修正後的儲存函式
// (請替換掉原本檔案最後面的 saveMaterial)
// ==========================================

function saveMaterial(data) {
  var lock = LockService.getScriptLock();
  // 嘗試取得鎖定，避免多人同時按儲存導致編號重複
  if (!lock.tryLock(10000)) return { success: false, error: "系統忙碌中，請稍後再試" };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.MATERIALS);
    if (!sheet) return { success: false, error: "找不到『原料主檔』工作表" };

    var nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
    
    // === 判斷是新增還是修改 ===
    if (data.rowIndex) {
      // ----------------------------
      // 📝 修改模式 (Update)
      // ----------------------------
      var row = parseInt(data.rowIndex);
      
      // 比對價格波動 (如果這次修改了成本，可以在備註自動加註，這是一種提醒方式)
      var oldCost = sheet.getRange(row, 6).getValue(); // F欄是成本
      var newCost = Number(data.cost);
      var autoNote = "";
      
      if (oldCost != newCost) {
        autoNote = ` (成本調整: ${oldCost} -> ${newCost})`;
      }

      // 更新欄位 (跳過 A欄 ID 和 I欄 庫存)
      sheet.getRange(row, 2).setValue(data.name);       // B: 名稱
      sheet.getRange(row, 3).setValue(data.spec);       // C: 規格
      sheet.getRange(row, 4).setValue(data.category);   // D: 類別
      sheet.getRange(row, 5).setValue(data.unit);       // E: 單位
      sheet.getRange(row, 6).setValue(newCost);         // F: 成本
      sheet.getRange(row, 7).setValue(data.supplier);   // G: 供應商
      sheet.getRange(row, 8).setValue(Number(data.minStock)); // H: 最低庫存
      sheet.getRange(row, 10).setValue(data.status);    // J: 狀態
      
      // 更新備註 (保留使用者輸入的備註 + 價格變動紀錄)
      var currentNote = data.note + autoNote;
      sheet.getRange(row, 12).setValue(currentNote);    // L: 備註
      
    } else {
      // ----------------------------
      // 🆕 新增模式 (Insert)
      // ----------------------------
      
      // 1. 自動產生 ID (確保使用 Max+1 邏輯)
      var newId = data.id;
      if (!newId) {
        // 使用您現有的 generateId 函式，它會自動找最大值
        newId = generateId("M", CONFIG.SHEETS.MATERIALS, 1);
      } else {
        // 如果手動輸入 ID，檢查是否重複
        var ids = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
        if (ids.indexOf(newId) > -1) return { success: false, error: "編號 [" + newId + "] 已存在！" };
      }

      // 2. 準備寫入資料
      var rowData = [
        newId,               // A
        data.name,           // B
        data.spec,           // C
        data.category,       // D
        data.unit,           // E
        Number(data.cost),   // F
        data.supplier,       // G
        Number(data.minStock), // H
        0,                   // I (新原料庫存預設為 0)
        data.status || '啟用', // J
        nowStr,              // K
        data.note            // L
      ];
      
      sheet.appendRow(rowData);
    }
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 🗑️ 刪除原料函式
// ==========================================
function deleteMaterial(id) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: "系統忙碌中" };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.MATERIALS); // 確保您設定檔裡有 MATERIALS 對應 "原料主檔"
    
    if (!sheet) return { success: false, error: "找不到原料主檔" };
    
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // 尋找要刪除的 ID 在第幾行
    // data[i][0] 是 A欄 (ID)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        rowIndex = i + 1; // 陣列索引 + 1 = 實際列號
        break;
      }
    }
    
    if (rowIndex > 0) {
      sheet.deleteRow(rowIndex);
      return { success: true };
    } else {
      return { success: false, error: "找不到此編號：" + id };
    }
    
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 補強：從銷貨單準備發票資料
// ==========================================
function prepareInvoiceFromOrder(orderId) {
  try {
    // 1. 取得銷貨單詳細資料
    var result = getSalesOrderDetail(orderId);
    if (!result.success) throw new Error(result.error);
    var order = result.order;

    // 2. 嘗試取得客戶統編 (選填，為了更完整)
    var taxId = "";
    var custResult = getCustomerById(order.customerId);
    if (custResult.success) {
      taxId = custResult.customer.taxId || "";
    }

    // 3. 整理發票表單需要的資料格式
    var invoiceData = {
      source: 'SALES',              // 來源：銷貨
      refNo: order.orderId,         // 關聯單號
      date: order.date,             // 日期
      targetName: order.customerName, // 對象名稱 (客戶)
      taxId: taxId,                 // 統一編號
      taxType: order.taxType,       // 稅別
      amount: order.subtotal,       // 未稅金額
      tax: order.taxAmount,         // 稅額
      total: order.grandTotal,      // 含稅總額
      note: "銷貨單號: " + order.orderId // 預設備註
    };

    // 4. 存入暫存，讓發票視窗讀取
    PropertiesService.getUserProperties().setProperty('TEMP_INVOICE_DATA', JSON.stringify(invoiceData));

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 🗑️ 刪除支出 (連動修正零用金)
// ==========================================
function deleteExpense(expId) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: "系統忙碌中" };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var expSheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES); // 現金支出
    var pcSheet = ss.getSheetByName(CONFIG.SHEETS.PETTY_CASH); // 零用金紀錄
    
    if (!expSheet) throw new Error("找不到支出表");

    // 1. 先找到該筆支出，確認它是不是用「零用金」付的
    var expData = expSheet.getDataRange().getValues();
    var targetRowIndex = -1;
    var paymentMethod = "";
    
    for (var i = 1; i < expData.length; i++) {
      if (String(expData[i][0]) === expId) { // A欄是單號
        targetRowIndex = i + 1; // 轉為實際列號
        paymentMethod = String(expData[i][5]); // F欄是付款方式
        break;
      }
    }

    if (targetRowIndex === -1) throw new Error("找不到此單號：" + expId);

    // 2. 如果是用「零用金」付的，要去零用金表把那筆扣款刪掉 (錢補回來)
    if (paymentMethod.indexOf("零用金") > -1 && pcSheet) {
      var pcData = pcSheet.getDataRange().getValues();
      for (var j = pcData.length - 1; j >= 1; j--) {
        // 檢查備註欄 (G欄/Index 6) 是否包含此支出單號
        // 或者檢查摘要是否相符，這裡用單號最準 (前提是 saveExpense 有存入關聯單號)
        var note = String(pcData[j][6]); 
        if (note.indexOf(expId) > -1) {
          pcSheet.deleteRow(j + 1); // 刪除該筆扣款紀錄
          break; // 找到一筆就停，避免刪錯
        }
      }
    }

    // 3. 最後刪除支出表的那一行
    expSheet.deleteRow(targetRowIndex);

    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 簡易月結報表 - 返回完整明細
 */
function getSimpleMonthlyReport(monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 日期比對工具
    var checkMonth = function(dateVal) {
      if (!dateVal) return false;
      if (dateVal instanceof Date) {
        return Utilities.formatDate(dateVal, "GMT+8", "yyyy-MM") === monthStr;
      }
      return String(dateVal).substring(0, 7) === monthStr;
    };
    
    var result = {
      success: true,
      salesList: [],
      purchaseList: [],
      expenseList: [],
      mealList: [],
      totals: {
        sales: 0,
        purchase: 0,
        expense: 0,
        meals: 0
      }
    };
    
    // ==========================================
    // 1. 銷貨收入（從銷貨單 + 收款記錄）
    // ==========================================
    var salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
    var paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_RECEIVED); // 收款記錄
    
    // 建立收款對照表
    var paymentMap = {};
    if (paymentSheet) {
      var payData = paymentSheet.getDataRange().getValues();
      for (var j = 1; j < payData.length; j++) {
        var orderId = String(payData[j][1]); // B欄：銷貨單號
        var payAmt = Number(payData[j][4]) || 0; // E欄：收款金額
        var payMethod = String(payData[j][5] || ''); // F欄：收款方式
        
        if (!paymentMap[orderId]) {
          paymentMap[orderId] = { total: 0, methods: [] };
        }
        paymentMap[orderId].total += payAmt;
        if (payMethod && paymentMap[orderId].methods.indexOf(payMethod) === -1) {
          paymentMap[orderId].methods.push(payMethod);
        }
      }
    }
    
    if (salesSheet) {
      var data = salesSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (checkMonth(data[i][1])) {
          var orderId = String(data[i][0]);
          var grandTotal = Number(data[i][8]) || 0; // I欄：總金額
          var originalNote = String(data[i][12] || ''); // M欄：原備註
          
          // 計算收款狀態
          var payment = paymentMap[orderId] || { total: 0, methods: [] };
          var paidAmt = payment.total;
          var unpaidAmt = grandTotal - paidAmt;
          
          var status = '';
          var statusColor = '';
          if (unpaidAmt <= 0) {
            status = '✅ 已收款';
            statusColor = 'green';
          } else if (paidAmt > 0) {
            status = '⚠️ 部分收款';
            statusColor = 'orange';
          } else {
            status = '❌ 未收款';
            statusColor = 'red';
          }
          
          // 組合備註
          var noteText = status;
          if (payment.methods.length > 0) {
            noteText += ' (' + payment.methods.join(', ') + ')';
          }
          if (originalNote) {
            noteText += ' | ' + originalNote;
          }
          
          result.salesList.push({
            orderId: orderId,
            date: formatDate(data[i][1]),
            customerName: String(data[i][3]),
            amount: grandTotal,
            paidAmount: paidAmt,
            unpaidAmount: unpaidAmt,
            status: status,
            statusColor: statusColor,
            paymentMethods: payment.methods.join(', '),
            note: noteText,
            originalNote: originalNote
          });
          result.totals.sales += grandTotal;
        }
      }
    }
    
    // ==========================================
    // 2. 進貨成本（從進貨單）
    // ==========================================
    var purSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE);
    if (purSheet) {
      var data = purSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (checkMonth(data[i][1])) {
          var amt = Number(data[i][6]) || 0;
          result.purchaseList.push({
            orderId: String(data[i][0]),
            date: formatDate(data[i][1]),
            supplierName: String(data[i][3]),
            amount: amt,
            note: String(data[i][11] || '')
          });
          result.totals.purchase += amt;
        }
      }
    }
    
    // ==========================================
    // 3. 費用支出
    // ==========================================
    var expSheet = ss.getSheetByName(CONFIG.SHEETS.EXPENSES);
    if (expSheet) {
      var data = expSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (checkMonth(data[i][1])) {
          var amt = Number(data[i][4]) || 0;
          result.expenseList.push({
            date: formatDate(data[i][1]),
            category: String(data[i][2]),
            item: String(data[i][3]),
            amount: amt,
            paymentMethod: String(data[i][5] || ''),
            payee: String(data[i][6] || '')
          });
          result.totals.expense += amt;
        }
      }
    }
    
    // ==========================================
    // 4. 員工餐費
    // ==========================================
    var mealSheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEE_MEALS);
    if (mealSheet) {
      var data = mealSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (checkMonth(data[i][1])) {
          var amt = Number(data[i][3]) || 0;
          result.mealList.push({
            date: formatDate(data[i][1]),
            headcount: Number(data[i][2]) || 0,
            amount: amt,
            paymentMethod: String(data[i][4] || ''),
            note: String(data[i][5] || '')
          });
          result.totals.meals += amt;
        }
      }
    }

    
    return result;
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// ➕ 進階快速新增 (支援完整欄位與自訂類別)
// ==========================================
function quickAddItem(form, type) {
  try {
    var result;
    
    // 情況 A：如果是「成品商品」
    // 欄位對應：A編號, B類別, C名稱, D規格, E單位, F售價...
    if (type === 'PRODUCT') {
      var prodData = {
        id: "AUTO",          // 自動產生編號
        name: form.name,     // C: 商品名稱
        category: form.category || "未分類", // B: 類別 (可自訂)
        spec: form.spec || "", // D: 規格
        unit: form.unit || "個", // E: 單位
        
        // 價格設定
        price: Number(form.price) || 0,         // F: 售價
        priceBranch: Number(form.priceBranch) || 0, // G: 總倉價
        priceStore: Number(form.priceStore) || 0,   // H: 店家價
        priceRest: Number(form.priceRest) || 0,     // I: 餐廳價
        cost: Number(form.cost) || 0,           // J: 成本
        
        supplier: form.supplier || "",          // K: 供應商
        alert: Number(form.alert) || 5,         // L: 安全庫存
        status: "啟用",                         // M: 狀態
        note: form.note || "進貨單快速新增"       // O: 備註
      };
      
      // 呼叫原本的 saveProduct
      result = saveProduct(prodData);
      
    } 
    // 情況 B：如果是「原物料」
    // 欄位對應：A編號, B名稱, C規格, D類別, E單位, F成本, G供應商...
    else {
      var matData = {
        id: "",              // 自動產生編號
        name: form.name,     // B: 原料名稱
        spec: form.spec || "", // C: 規格
        category: form.category || "未分類", // D: 類別
        unit: form.unit || "個", // E: 單位
        cost: Number(form.cost) || 0, // F: 預設成本
        supplier: form.supplier || "", // G: 供應商
        minStock: Number(form.alert) || 0, // H: 最低庫存
        
        status: "啟用",      // J: 狀態
        note: form.note || "進貨單快速新增" // L: 備註
      };
      
      // 呼叫原本的 saveMaterial
      result = saveMaterial(matData);
    }

    return result;

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}


// ==========================================
// ➕ 快速新增供應商 (給進貨單專用)
// ==========================================
function quickAddSupplier(form) {
  try {
    // 呼叫既有的 addSupplier 函式
    // 我們在這裡幫忙補齊該函式需要的欄位格式
    var data = {
      name: form.name,
      taxId: form.taxId || "",
      phone: form.phone || "",
      paymentTerm: form.paymentTerm || "貨到付現",
      paymentMethod: form.paymentMethod || "現金",
      taxType: form.taxType || "免稅",  // 這是進貨單計算關鍵
      supplierType: "進貨供應商",       // 預設類型
      status: "啟用",
      note: "進貨單快速新增"
    };
    
    return addSupplier(data); // 呼叫原本的寫入函式

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function generatePrintHTML(orderId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ordersSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
    
    // 取得訂單資料
    var orderData = ordersSheet.getDataRange().getValues();
    var orderRow = null;
    
    for (var i = 1; i < orderData.length; i++) {
      if (String(orderData[i][0]) === String(orderId)) {
        orderRow = orderData[i];
        break;
      }
    }
    
    if (!orderRow) throw new Error('找不到此銷貨單');
    
    // 解析訂單資料
    var order = {
      orderId: orderRow[0],
      date: Utilities.formatDate(new Date(orderRow[1]), 'GMT+8', 'yyyy/MM/dd'),
      customerId: orderRow[2],
      customerName: orderRow[3],
      subtotal: orderRow[4] || 0,
      shipping: orderRow[5] || 0,
      deduction: orderRow[6] || 0,
      taxAmount: orderRow[7] || 0,
      grandTotal: orderRow[8] || 0,
      taxType: orderRow[9] || '免稅',
      note: orderRow[12] || ''
    };
    
    // 取得客戶詳細資料
    var customerSheet = ss.getSheetByName(CONFIG.SHEETS.CUSTOMERS);
    var customerData = customerSheet.getDataRange().getValues();
    var customerInfo = { phone: '', address: '', taxId: '' };
    
    for (var i = 1; i < customerData.length; i++) {
      if (String(customerData[i][0]) === String(order.customerId)) {
        customerInfo.phone = String(customerData[i][4] || customerData[i][3] || '');
        customerInfo.address = String(customerData[i][5] || '');
        customerInfo.taxId = String(customerData[i][7] || '');
        break;
      }
    }
    
    // 取得明細
    var detailSheet = ss.getSheetByName(CONFIG.SHEETS.SALES_DETAILS);
    var detailData = detailSheet.getDataRange().getValues();
    var items = [];

    for (var i = 1; i < detailData.length; i++) {
      if (String(detailData[i][1]) === String(orderId)) {
        items.push({
          productId: String(detailData[i][2]),
          productName: String(detailData[i][3]),
          qty: Number(detailData[i][4]),
          unit: String(detailData[i][5]),
          price: Number(detailData[i][6]),  // ✅ 這行很重要！
          note: String(detailData[i][10] || '')
        });
      }
    }
    
    // 組合明細行 HTML (含商品編號與金額)
    var itemsHtml = items.map(function(item, i) {
      var subtotal = item.qty * item.price;
      var nameDisplay = '<div>' + item.productName + '</div><div style="font-size:11px;color:#666;">' + item.productId + '</div>';
      if (item.note) nameDisplay += '<div class="note-text">(' + item.note + ')</div>';
      return '<tr><td>' + (i + 1) + '</td><td class="text-left">' + nameDisplay + '</td><td>' + item.unit + '</td><td>' + item.qty + '</td><td class="text-right">' + item.price.toLocaleString() + '</td><td class="text-right">' + subtotal.toLocaleString() + '</td></tr>';
    }).join('');
    
    // 組合合計區
    var summaryLines = [];
    summaryLines.push('<b>小計：</b>NT$ ' + order.subtotal.toLocaleString());
    if (order.shipping > 0) {
      summaryLines.push('<b>運費：</b>NT$ ' + order.shipping.toLocaleString());
    }
    if (order.deduction > 0) {
      summaryLines.push('<b>折扣：</b>-NT$ ' + order.deduction.toLocaleString());
    }
    if (order.taxAmount > 0) {
      summaryLines.push('<b>稅金：</b>NT$ ' + order.taxAmount.toLocaleString());
    }
    var summaryHtml = summaryLines.join('<br>');
    
    // 組合完整 HTML
    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>銷貨單 - ' + order.orderId + '</title>' +
      '<style>' +
      '*{margin:0;padding:0;box-sizing:border-box}' +
      '@page{size:A4;margin:0}' +
      'body{font-family:"Microsoft JhengHei",sans-serif;padding:8mm 10mm;font-size:13px;color:#000}' +
      '.header{text-align:center;margin-bottom:3mm;border-bottom:1.5px solid #000;padding-bottom:1.5mm}' +
      '.header h1{font-size:22px;margin-bottom:2mm}' +
      '.header p{font-size:14px;letter-spacing:12px;text-indent:12px}' +
      '.info-box{background:#f5f5f5!important;padding:3mm;border:1px solid #999;margin-bottom:3mm;display:grid;grid-template-columns:1fr 1fr;gap:1mm 4mm;-webkit-print-color-adjust:exact}' +
      '.info-item{line-height:1.4}' +
      'table{width:100%;border-collapse:collapse;border:1.5px solid #000;table-layout:fixed}' +
      'th{background:#e8e8e8!important;padding:1.5mm 1mm;border:1px solid #000;text-align:center;font-weight:bold;-webkit-print-color-adjust:exact}' +
      'td{padding:1.5mm 1mm;border:1px solid #000;vertical-align:middle;text-align:center;line-height:1.2}' +
      '.text-right{text-align:right!important;padding-right:2mm!important}' +
      '.text-left{text-align:left!important;padding-left:2mm!important}' +
      '.summary{text-align:right;margin-top:2mm;line-height:1.4}' +
      '.total-line{font-size:18px;font-weight:bold;border-top:1.8px solid #000;margin-top:1mm;padding-top:1mm}' +
      '.note-text{font-size:11px;color:#666;font-style:italic}' +
      '.remark{margin-top:3mm;padding:2mm;border:1px dashed #999;font-size:12px}' +
      '</style></head><body onload="setTimeout(function(){ window.print(); }, 300);">' +
      '<div class="header"><h1>丸十水產股份有限公司</h1><p>銷貨單</p></div>' +
      '<div class="info-box">' +
      '<div class="info-item"><b>單號：</b>' + order.orderId + '</div>' +
      '<div class="info-item"><b>日期：</b>' + order.date + '</div>' +
      '<div class="info-item"><b>客戶名稱：</b>' + order.customerName + '</div>' +
      '<div class="info-item"><b>電話：</b>' + customerInfo.phone + '</div>' +
      '<div class="info-item"><b>地址：</b>' + customerInfo.address + '</div>' +
      '<div class="info-item"><b>稅制：</b>' + order.taxType + '</div>' +
      '</div>' +
      '<table><thead><tr>' +
      '<th width="7%">序</th>' +
      '<th>品名/規格</th>' +
      '<th width="8%">單位</th>' +
      '<th width="10%">數量</th>' +
      '<th width="14%" class="text-right">單價</th>' +
      '<th width="16%" class="text-right">金額</th>' +
      '</tr></thead><tbody>' +
      itemsHtml +
      '</tbody></table>' +
      '<div class="summary">' + summaryHtml + '<div class="total-line">總計：NT$ ' + order.grandTotal.toLocaleString() + '</div></div>' +
      (order.note ? '<div class="remark"><b>備註：</b>' + order.note + '</div>' : '') +
      '</body></html>';
    
    return html;
    
  } catch (error) {
    throw new Error('列印失敗: ' + error.toString());
  }
}
// ==========================================
// 🖨️ 列印函式 (標準銷貨單 + 配送單)
// ==========================================

/**
 * 📝 標準銷貨單列印 (含金額)
 */
function generateSalesPrintHTML(orderId) {
  return generatePrintHTML(orderId);
}

/**
 * 📝 標準銷貨單列印 - 主函式
 */
function generatePrintHTML(orderId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ordersSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
    
    var orderData = ordersSheet.getDataRange().getValues();
    var orderRow = null;
    
    for (var i = 1; i < orderData.length; i++) {
      if (String(orderData[i][0]) === String(orderId)) {
        orderRow = orderData[i];
        break;
      }
    }
    
    if (!orderRow) throw new Error('找不到此銷貨單');
    
    var order = {
      orderId: orderRow[0],
      date: Utilities.formatDate(new Date(orderRow[1]), 'GMT+8', 'yyyy/MM/dd'),
      customerId: orderRow[2],
      customerName: orderRow[3],
      subtotal: orderRow[4] || 0,
      shipping: orderRow[5] || 0,
      deduction: orderRow[6] || 0,
      taxAmount: orderRow[7] || 0,
      grandTotal: orderRow[8] || 0,
      taxType: orderRow[9] || '免稅',
      note: orderRow[12] || ''
    };
    
    var customerSheet = ss.getSheetByName(CONFIG.SHEETS.CUSTOMERS);
    var customerData = customerSheet.getDataRange().getValues();
    var customerInfo = { phone: '', address: '', taxId: '' };
    
    for (var i = 1; i < customerData.length; i++) {
      if (String(customerData[i][0]) === String(order.customerId)) {
        customerInfo.phone = String(customerData[i][4] || customerData[i][3] || '');
        customerInfo.address = String(customerData[i][5] || '');
        customerInfo.taxId = String(customerData[i][7] || '');
        break;
      }
    }
    
    var detailSheet = ss.getSheetByName(CONFIG.SHEETS.SALES_DETAILS);
    var detailData = detailSheet.getDataRange().getValues();
    var items = [];

    for (var i = 1; i < detailData.length; i++) {
      if (String(detailData[i][1]) === String(orderId)) {
        items.push({
          productId: String(detailData[i][2]),
          productName: String(detailData[i][3]),
          qty: Number(detailData[i][4]),
          unit: String(detailData[i][5]),
          price: Number(detailData[i][6]),
          note: String(detailData[i][10] || '')
        });
      }
    }
    
    // --- 修改處 1：內容生成的 HTML 調整 ---
    var itemsHtml = items.map(function(item, i) {
      var subtotal = item.qty * item.price;
      // 這裡原本有 productId，現在移除，只保留品名
      var nameDisplay = '<div>' + item.productName + '</div>';
      if (item.note) nameDisplay += '<div class="note-text">(' + item.note + ')</div>';
      
      // 新增第二個 td 放 productId
      return '<tr>' +
             '<td>' + (i + 1) + '</td>' +
             '<td>' + item.productId + '</td>' + // 獨立的編號欄位
             '<td class="text-left">' + nameDisplay + '</td>' +
             '<td>' + item.unit + '</td>' +
             '<td>' + item.qty + '</td>' +
             '<td class="text-right">' + item.price.toLocaleString() + '</td>' +
             '<td class="text-right">' + subtotal.toLocaleString() + '</td>' +
             '</tr>';
    }).join('');
    
    var summaryLines = [];
    summaryLines.push('<b>小計：</b>NT$ ' + order.subtotal.toLocaleString());
    if (order.shipping > 0) summaryLines.push('<b>運費：</b>NT$ ' + order.shipping.toLocaleString());
    if (order.deduction > 0) summaryLines.push('<b>折扣：</b>-NT$ ' + order.deduction.toLocaleString());
    if (order.taxAmount > 0) summaryLines.push('<b>稅金：</b>NT$ ' + order.taxAmount.toLocaleString());
    var summaryHtml = summaryLines.join('<br>');
    
    // --- 修改處 2：表頭 (thead) 增加編號欄位並調整寬度 ---
    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>銷貨單 - ' + order.orderId + '</title>' +
      '<style>*{margin:0;padding:0;box-sizing:border-box}@page{size:A4;margin:0}body{font-family:"Microsoft JhengHei",sans-serif;padding:8mm 10mm;font-size:13px;color:#000}.header{text-align:center;margin-bottom:3mm;border-bottom:1.5px solid #000;padding-bottom:1.5mm}.header h1{font-size:22px;margin-bottom:2mm}.header p{font-size:14px;letter-spacing:12px;text-indent:12px}.info-box{background:#f5f5f5!important;padding:3mm;border:1px solid #999;margin-bottom:3mm;display:grid;grid-template-columns:1fr 1fr;gap:1mm 4mm;-webkit-print-color-adjust:exact}.info-item{line-height:1.4}table{width:100%;border-collapse:collapse;border:1.5px solid #000;table-layout:fixed}th{background:#e8e8e8!important;padding:1.5mm 1mm;border:1px solid #000;text-align:center;font-weight:bold;-webkit-print-color-adjust:exact}td{padding:1.5mm 1mm;border:1px solid #000;vertical-align:middle;text-align:center;line-height:1.2}.text-right{text-align:right!important;padding-right:2mm!important}.text-left{text-align:left!important;padding-left:2mm!important}.summary{text-align:right;margin-top:2mm;line-height:1.4}.total-line{font-size:18px;font-weight:bold;border-top:1.8px solid #000;margin-top:1mm;padding-top:1mm}.note-text{font-size:11px;color:#666;font-style:italic}.remark{margin-top:3mm;padding:2mm;border:1px dashed #999;font-size:12px}</style></head><body onload="setTimeout(function(){ window.print(); }, 300);">' +
      '<div class="header"><h1>丸十水產股份有限公司</h1><p>銷貨單</p></div>' +
      '<div class="info-box"><div class="info-item"><b>單號：</b>' + order.orderId + '</div><div class="info-item"><b>日期：</b>' + order.date + '</div><div class="info-item"><b>客戶名稱：</b>' + order.customerName + '</div><div class="info-item"><b>電話：</b>' + customerInfo.phone + '</div><div class="info-item"><b>地址：</b>' + customerInfo.address + '</div><div class="info-item"><b>稅制：</b>' + order.taxType + '</div></div>' +
      '<table><thead><tr>' + 
      '<th width="5%">序</th>' + 
      '<th width="15%">商品編號</th>' + // 新增這一欄
      '<th>品名/規格</th>' + 
      '<th width="8%">單位</th>' + 
      '<th width="8%">數量</th>' + 
      '<th width="12%" class="text-right">單價</th>' + 
      '<th width="14%" class="text-right">金額</th>' + 
      '</tr></thead><tbody>' + itemsHtml + '</tbody></table>' +
      '<div class="summary">' + summaryHtml + '<div class="total-line">總計：NT$ ' + order.grandTotal.toLocaleString() + '</div></div>' +
      (order.note ? '<div class="remark"><b>備註：</b>' + order.note + '</div>' : '') +
      '</body></html>';
    
    return html;
  } catch (error) {
    throw new Error('列印失敗: ' + error.toString());
  }
}

/**
 * 📦 配送單列印 (不含金額)
 */
function generateDeliveryPrintHTML(orderId, recipientName, recipientPhone, recipientAddress, recipientNote) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ordersSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
    
    var orderData = ordersSheet.getDataRange().getValues();
    var orderRow = null;
    
    for (var i = 1; i < orderData.length; i++) {
      if (String(orderData[i][0]) === String(orderId)) {
        orderRow = orderData[i];
        break;
      }
    }
    
    if (!orderRow) throw new Error('找不到此銷貨單');
    
    var order = {
      orderId: orderRow[0],
      date: Utilities.formatDate(new Date(orderRow[1]), 'GMT+8', 'yyyy/MM/dd'),
      customerName: orderRow[3],
      note: orderRow[12] || ''
    };
    
    var detailSheet = ss.getSheetByName(CONFIG.SHEETS.SALES_DETAILS);
    var detailData = detailSheet.getDataRange().getValues();
    var items = [];

    for (var i = 1; i < detailData.length; i++) {
      if (String(detailData[i][1]) === String(orderId)) {
        items.push({
          productId: String(detailData[i][2]),
          productName: String(detailData[i][3]),
          qty: Number(detailData[i][4]),
          unit: String(detailData[i][5]),
          note: String(detailData[i][10] || '')
        });
      }
    }
    
    // --- 修改處 3：配送單內容生成的 HTML 調整 ---
    var itemsHtml = items.map(function(item, i) {
      // 移除 productId 的堆疊顯示
      var nameDisplay = '<div>' + item.productName + '</div>';
      if (item.note) nameDisplay += '<div class="note-text">(' + item.note + ')</div>';
      
      return '<tr>' +
             '<td>' + (i + 1) + '</td>' +
             '<td>' + item.productId + '</td>' + // 獨立的編號欄位
             '<td class="text-left">' + nameDisplay + '</td>' +
             '<td>' + item.unit + '</td>' +
             '<td style="font-weight:bold;">' + item.qty + '</td>' +
             '</tr>';
    }).join('');
    
    var totalQty = items.reduce(function(sum, item) { return sum + item.qty; }, 0);
    
    // --- 修改處 4：配送單表頭 (thead) 增加編號欄位並調整寬度 ---
    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>配送單 - ' + order.orderId + '</title>' +
      '<style>*{margin:0;padding:0;box-sizing:border-box}@page{size:A4;margin:0}body{font-family:"Microsoft JhengHei",sans-serif;padding:8mm 10mm;font-size:13px;color:#000}.header{text-align:center;margin-bottom:3mm;border-bottom:1.5px solid #000;padding-bottom:1.5mm}.header h1{font-size:22px;margin-bottom:2mm}.header p{font-size:14px;letter-spacing:12px;text-indent:12px}.info-box{background:#f5f5f5!important;padding:3mm;border:1px solid #999;margin-bottom:3mm;display:grid;grid-template-columns:1fr 1fr;gap:1mm 4mm;-webkit-print-color-adjust:exact}.info-item{line-height:1.4}table{width:100%;border-collapse:collapse;border:1.5px solid #000;table-layout:fixed}th{background:#e8e8e8!important;padding:1.5mm 1mm;border:1px solid #000;text-align:center;font-weight:bold;-webkit-print-color-adjust:exact}td{padding:1.5mm 1mm;border:1px solid #000;vertical-align:middle;text-align:center;line-height:1.2}.text-left{text-align:left!important;padding-left:2mm!important}.note-text{font-size:11px;color:#666;font-style:italic}.summary{text-align:right;margin-top:2mm;line-height:1.4}.total-line{font-size:16px;font-weight:bold;border-top:1.8px solid #000;margin-top:1mm;padding-top:1mm}.remark{margin-top:3mm;padding:2mm;border:1px dashed #999;font-size:12px}</style></head><body onload="setTimeout(function(){ window.print(); }, 300);">' +
      '<div class="header"><h1>丸十水產股份有限公司</h1><p>配送單</p></div>' +
      '<div class="info-box"><div class="info-item"><b>單號：</b>' + order.orderId + '</div><div class="info-item"><b>日期：</b>' + order.date + '</div><div class="info-item"><b>收件人：</b>' + recipientName + '</div><div class="info-item"><b>電話：</b>' + recipientPhone + '</div><div class="info-item" style="grid-column:1/3;"><b>地址：</b>' + recipientAddress + '</div>' + (recipientNote ? '<div class="info-item" style="grid-column:1/3;"><b>配送備註：</b>' + recipientNote + '</div>' : '') + '</div>' +
      '<table><thead><tr>' + 
      '<th width="6%">序</th>' + 
      '<th width="18%">商品編號</th>' + // 新增這一欄
      '<th>品名/規格</th>' + 
      '<th width="10%">單位</th>' + 
      '<th width="10%">數量</th>' + 
      '</tr></thead><tbody>' + itemsHtml + '</tbody></table>' +
      '<div class="summary"><div class="total-line">合計數量：' + totalQty + ' 件</div></div>' +
      (order.note ? '<div class="remark"><b>訂單備註：</b>' + order.note + '</div>' : '') +
      '</body></html>';
    
    return html;
  } catch (error) {
    throw new Error('配送單列印失敗: ' + error.toString());
  }
}

// ==========================================
// 🖨️ 進貨單列印函式 (樣式同步更新)
// ==========================================

/**
 * 📝 進貨單列印
 */
function generatePurchasePrintHTML(orderId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var purchaseSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE);
    
    // 1. 取得進貨單主檔資料
    var purchaseData = purchaseSheet.getDataRange().getValues();
    var orderRow = null;
    
    for (var i = 1; i < purchaseData.length; i++) {
      if (String(purchaseData[i][0]) === String(orderId)) {
        orderRow = purchaseData[i];
        break;
      }
    }
    
    if (!orderRow) throw new Error('找不到此進貨單');
    
    var order = {
      orderId: orderRow[0],
      date: Utilities.formatDate(new Date(orderRow[1]), 'GMT+8', 'yyyy/MM/dd'),
      supplierId: orderRow[2],
      supplierName: orderRow[3],
      totalAmount: orderRow[4] || 0, // 未稅
      taxAmount: orderRow[5] || 0,   // 稅額
      grandTotal: orderRow[6] || 0,  // 含稅總額
      handler: orderRow[7] || '',
      note: orderRow[11] || ''
    };
    
    // 2. 取得供應商詳細資料 (電話、地址)
    var supplierSheet = ss.getSheetByName(CONFIG.SHEETS.SUPPLIERS);
    var supplierInfo = { phone: '', address: '' };
    
    if (supplierSheet) {
      var supplierData = supplierSheet.getDataRange().getValues();
      for (var i = 1; i < supplierData.length; i++) {
        if (String(supplierData[i][0]) === String(order.supplierId)) {
          // 假設供應商表：D欄(Index 3)是電話, F欄(Index 5)是地址
          supplierInfo.phone = String(supplierData[i][3] || '');
          supplierInfo.address = String(supplierData[i][5] || '');
          break;
        }
      }
    }
    
    // 3. 取得進貨明細
    var detailSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_DETAILS);
    var detailData = detailSheet.getDataRange().getValues();
    var items = [];

    for (var i = 1; i < detailData.length; i++) {
      if (String(detailData[i][1]) === String(orderId)) {
        items.push({
          productId: String(detailData[i][3]),  // D欄
          productName: String(detailData[i][4]),// E欄
          qty: Number(detailData[i][5]),        // F欄
          unit: String(detailData[i][6]),       // G欄
          cost: Number(detailData[i][7]),       // H欄 (進價)
          note: String(detailData[i][10] || '') // K欄
        });
      }
    }
    
    // 4. 產生明細 HTML (仿照銷貨單樣式，獨立編號欄位)
    var itemsHtml = items.map(function(item, i) {
      var subtotal = item.qty * item.cost;
      
      var nameDisplay = '<div>' + item.productName + '</div>';
      if (item.note) nameDisplay += '<div class="note-text">(' + item.note + ')</div>';
      
      return '<tr>' +
             '<td>' + (i + 1) + '</td>' +
             '<td>' + item.productId + '</td>' + // 獨立的編號欄位
             '<td class="text-left">' + nameDisplay + '</td>' +
             '<td>' + item.unit + '</td>' +
             '<td>' + item.qty + '</td>' +
             '<td class="text-right">' + item.cost.toLocaleString() + '</td>' +
             '<td class="text-right">' + subtotal.toLocaleString() + '</td>' +
             '</tr>';
    }).join('');
    
    // 5. 產生合計區 HTML
    var summaryLines = [];
    summaryLines.push('<b>未稅金額：</b>NT$ ' + order.totalAmount.toLocaleString());
    if (order.taxAmount > 0) {
      summaryLines.push('<b>稅金：</b>NT$ ' + order.taxAmount.toLocaleString());
    }
    var summaryHtml = summaryLines.join('<br>');
    
    // 6. 組合完整 HTML (樣式與銷貨單完全一致)
    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>進貨單 - ' + order.orderId + '</title>' +
      '<style>*{margin:0;padding:0;box-sizing:border-box}@page{size:A4;margin:0}body{font-family:"Microsoft JhengHei",sans-serif;padding:8mm 10mm;font-size:13px;color:#000}.header{text-align:center;margin-bottom:3mm;border-bottom:1.5px solid #000;padding-bottom:1.5mm}.header h1{font-size:22px;margin-bottom:2mm}.header p{font-size:14px;letter-spacing:12px;text-indent:12px}.info-box{background:#f5f5f5!important;padding:3mm;border:1px solid #999;margin-bottom:3mm;display:grid;grid-template-columns:1fr 1fr;gap:1mm 4mm;-webkit-print-color-adjust:exact}.info-item{line-height:1.4}table{width:100%;border-collapse:collapse;border:1.5px solid #000;table-layout:fixed}th{background:#e8e8e8!important;padding:1.5mm 1mm;border:1px solid #000;text-align:center;font-weight:bold;-webkit-print-color-adjust:exact}td{padding:1.5mm 1mm;border:1px solid #000;vertical-align:middle;text-align:center;line-height:1.2}.text-right{text-align:right!important;padding-right:2mm!important}.text-left{text-align:left!important;padding-left:2mm!important}.summary{text-align:right;margin-top:2mm;line-height:1.4}.total-line{font-size:18px;font-weight:bold;border-top:1.8px solid #000;margin-top:1mm;padding-top:1mm}.note-text{font-size:11px;color:#666;font-style:italic}.remark{margin-top:3mm;padding:2mm;border:1px dashed #999;font-size:12px}</style></head><body onload="setTimeout(function(){ window.print(); }, 300);">' +
      '<div class="header"><h1>丸十水產股份有限公司</h1><p>進貨單</p></div>' +
      '<div class="info-box">' +
      '<div class="info-item"><b>單號：</b>' + order.orderId + '</div>' +
      '<div class="info-item"><b>日期：</b>' + order.date + '</div>' +
      '<div class="info-item"><b>廠商名稱：</b>' + order.supplierName + '</div>' +
      '<div class="info-item"><b>電話：</b>' + supplierInfo.phone + '</div>' +
      '<div class="info-item"><b>地址：</b>' + supplierInfo.address + '</div>' +
      '<div class="info-item"><b>經辦人：</b>' + order.handler + '</div>' +
      '</div>' +
      '<table><thead><tr>' + 
      '<th width="5%">序</th>' + 
      '<th width="15%">商品編號</th>' + // 獨立欄位
      '<th>品名/規格</th>' + 
      '<th width="8%">單位</th>' + 
      '<th width="8%">數量</th>' + 
      '<th width="12%" class="text-right">單價</th>' + 
      '<th width="14%" class="text-right">金額</th>' + 
      '</tr></thead><tbody>' + itemsHtml + '</tbody></table>' +
      '<div class="summary">' + summaryHtml + '<div class="total-line">總計：NT$ ' + order.grandTotal.toLocaleString() + '</div></div>' +
      (order.note ? '<div class="remark"><b>備註：</b>' + order.note + '</div>' : '') +
      '</body></html>';
    
    return html;
  } catch (error) {
    throw new Error('列印失敗: ' + error.toString());
  }
}


// ==========================================
// 🧨 管理員專用：一次性系統重置
// ==========================================
function admin_RESET_SYSTEM_NOW() {
  var ui = SpreadsheetApp.getUi();
  
  // 1. 安全確認
  var result = ui.alert(
    '🛑 嚴重警告：手動重置系統',
    '您正在執行後台強制重置。\n\n' +
    '這將【永久刪除】所有資料：\n' +
    '1. 所有訂單 (進貨/銷貨)\n' +
    '2. 所有財務與庫存紀錄\n' +
    '3. 所有基本資料 (商品/客戶/供應商)\n\n' +
    '確定要執行嗎？',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  // 2. 二次確認密碼
  var prompt = ui.prompt('🔒 最終確認', '請輸入 "DELETE" (全大寫) 開始清除：', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK || prompt.getResponseText() !== 'DELETE') {
    ui.alert('❌ 密碼錯誤或取消，未執行任何動作。');
    return;
  }

  // 3. 開始執行
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { ui.alert('系統忙碌中'); return; }

  try {
    // 定義要清空的所有工作表
    var sheets = [
      CONFIG.SHEETS.CALENDAR,
      CONFIG.SHEETS.PRODUCTS,
      CONFIG.SHEETS.MATERIALS,
      CONFIG.SHEETS.CUSTOMERS,
      CONFIG.SHEETS.SUPPLIERS,
      CONFIG.SHEETS.SPECIAL_PRICE,
      CONFIG.SHEETS.PURCHASE,
      CONFIG.SHEETS.PURCHASE_DETAILS,
      CONFIG.SHEETS.PAYABLE,
      CONFIG.SHEETS.SALES,
      CONFIG.SHEETS.SALES_DETAILS,
      CONFIG.SHEETS.RECEIVABLE,
      CONFIG.SHEETS.INVENTORY,
      CONFIG.SHEETS.INVENTORY_LOG,
      CONFIG.SHEETS.STOCKTAKE,
      CONFIG.SHEETS.EXPENSES,
      CONFIG.SHEETS.PETTY_CASH,
      CONFIG.SHEETS.PAYMENT_RECEIVED,
      CONFIG.SHEETS.PAYMENT_MADE,
      CONFIG.SHEETS.MONTHLY_REPORT,
      CONFIG.SHEETS.INVOICES_ISSUED,
      CONFIG.SHEETS.INVOICES_RECEIVED,
      CONFIG.SHEETS.SUPPLIER_PRICES,
      CONFIG.SHEETS.EMPLOYEE_MEALS
    ];

    // 迴圈清除內容 (保留標題列)
    sheets.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (sheet) {
        var lastRow = sheet.getLastRow();
        var lastCol = sheet.getLastColumn();
        if (lastRow > 1 && lastCol > 0) {
          sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
        }
      }
    });

    // 清除系統暫存屬性
    PropertiesService.getUserProperties().deleteAllProperties();
    PropertiesService.getScriptProperties().deleteAllProperties();

    ui.alert('✅ 系統已重置完成！所有資料已清空。');

  } catch (e) {
    ui.alert('❌ 錯誤：' + e.toString());
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 🛠️ 補救工具：從現有的進貨單，重建「供應商進價記錄」
// ==========================================
function admin_Fix_SupplierPrices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var purSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE);
  var detSheet = ss.getSheetByName(CONFIG.SHEETS.PURCHASE_DETAILS);
  var priceSheet = ss.getSheetByName(CONFIG.SHEETS.SUPPLIER_PRICES);

  if (!purSheet || !detSheet || !priceSheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到必要的工作表 (進貨單/明細/進價記錄)");
    return;
  }

  // 1. 讀取現有資料
  var purData = purSheet.getDataRange().getValues();
  var detData = detSheet.getDataRange().getValues();
  
  // 建立進貨單索引 (Order ID -> {Date, SupplierID, SupplierName})
  var orderMap = {};
  for (var i = 1; i < purData.length; i++) {
    var oid = String(purData[i][0]);
    orderMap[oid] = {
      date: purData[i][1], // B欄 日期
      suppId: purData[i][2], // C欄 廠商ID
      suppName: purData[i][3] // D欄 廠商名稱
    };
  }

  // 2. 準備寫入進價記錄的資料
  var newRows = [];
  var nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");

  // 遍歷進貨明細
  for (var j = 1; j < detData.length; j++) {
    var oid = String(detData[j][1]); // B欄 進貨單號
    var info = orderMap[oid];
    
    if (info) {
      // 產生一筆進價記錄
      // 格式假設為: [流水號, 日期, 商品ID, 商品名稱, 成本, 單位, 供應商ID, 供應商名稱, 建立時間]
      // 請根據您實際的「供應商進價記錄」欄位順序調整，這裡以通用邏輯為主
      var priceId = "SP-" + oid + "-" + (j); // 暫時用組合ID
      
      newRows.push([
        priceId,                // A: 流水號
        info.date,              // B: 進貨日期
        detData[j][3],          // C: 商品ID
        detData[j][4],          // D: 商品名稱
        detData[j][7],          // E: 進貨單價(成本)
        detData[j][6],          // F: 單位
        info.suppId,            // G: 供應商ID
        info.suppName,          // H: 供應商名稱
        nowStr                  // I: 記錄時間
      ]);
    }
  }

  // 3. 寫入資料表 (先清空舊的比較保險，或者直接往下加)
  if (newRows.length > 0) {
    // 這裡採用「清空後重寫」的方式，確保資料不重複
    // 如果您不想清空，請註解掉下面這行，並改用 appendRow
    if (priceSheet.getLastRow() > 1) {
       priceSheet.getRange(2, 1, priceSheet.getLastRow()-1, 9).clearContent();
    }
    
    priceSheet.getRange(2, 1, newRows.length, 9).setValues(newRows);
    SpreadsheetApp.getUi().alert("✅ 已成功補救 " + newRows.length + " 筆進價記錄！");
  } else {
    SpreadsheetApp.getUi().alert("⚠️ 沒有找到任何進貨明細資料。");
  }
}


/* ========== 新增：大陸貨款專用功能 ========== */

// 取得大陸帳務列表與統計
function getCNList(monthStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("CN_Fund");
  if (!sheet) return { list: [], summary: { balance: 0, company: 0, personal: 0 } };

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { list: [], summary: { balance: 0, company: 0, personal: 0 } };

  var list = [];
  var totalRMB_Balance = 0; // 總餘額 (歷史累計)
  var monthCompany = 0;     // 本月公司支出
  var monthPersonal = 0;    // 本月個人支出

  // 格式化月份比較字串 (ex: "2026-02")
  var filterMonth = monthStr || ""; 

  // 從第二列開始讀取 (跳過標題)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rDate = new Date(row[1]);
    var rMonth = Utilities.formatDate(rDate, Session.getScriptTimeZone(), "yyyy-MM");
    
    var type = row[2]; // 入金, 公司貨款, 個人支出
    var amount = parseFloat(row[3]) || 0;

    // 計算總餘額 (所有歷史紀錄)
    if (type === "入金") {
      totalRMB_Balance += amount;
    } else {
      totalRMB_Balance -= amount; // 支出扣錢
    }

    // 篩選當月資料供列表與月統計使用
    if (rMonth === filterMonth) {
      if (type === "公司貨款") monthCompany += amount;
      if (type === "個人支出") monthPersonal += amount;

      list.push({
        id: row[0],
        date: Utilities.formatDate(rDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        type: type,
        rmb: amount,
        rate: row[4],
        ntd: row[5],
        item: row[6],
        note: row[7]
      });
    }
  }

  // 列表倒序排列 (新的在上面)
  list.sort(function(a, b) { return b.date.localeCompare(a.date); });

  return {
    list: list,
    summary: {
      balance: totalRMB_Balance, // 這是目前剩餘總額
      company: monthCompany,
      personal: monthPersonal
    }
  };
}

// 儲存大陸帳務
function saveCNTransaction(form) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("CN_Fund");
    if (!sheet) return { success: false, error: "找不到 'CN_Fund' 工作表" };

    var uuid = Utilities.getUuid();
    var ntdAmount = 0;

    // 如果是入金，計算台幣金額
    if (form.type === "入金" && form.rate) {
      ntdAmount = Math.round(form.rmb * form.rate);
    }

    sheet.appendRow([
      uuid,
      "'" + form.date, // 強制字串格式
      form.type,
      form.rmb,
      form.rate || "", // 支出通常沒有當下匯率，留空
      ntdAmount || "",
      form.item,
      form.note
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 刪除大陸帳務
function deleteCNTransaction(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("CN_Fund");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: "找不到該筆資料" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 🇨🇳 大陸貨款管理模組
// ==========================================

// 設定常數
var CHINA_CONFIG = {
  SHEETS: {
    DEPOSIT: "大陸入帳記錄",
    MONTHLY: "大陸月結記錄",
    PURCHASE: "大陸進貨"  // 改成這個
  }
};


/**
 * 取得指定月份的入帳總額
 */
function getChinaDepositTotal(monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.DEPOSIT);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, total: 0, count: 0 };
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    var total = 0;
    var count = 0;
    
    for (var i = 0; i < data.length; i++) {
      var dateStr = formatDate(data[i][1]).substring(0, 7); // B欄日期
      if (dateStr === monthStr) {
        total += Number(data[i][4]) || 0; // E欄台幣金額
        count++;
      }
    }
    
    return { success: true, total: total, count: count };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * 取得上月結餘
 */
function getLastMonthBalance(monthStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.MONTHLY);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return 0; // 沒有歷史記錄，從 0 開始
    }
    
    // 計算上個月
    var parts = monthStr.split('-');
    var year = parseInt(parts[0]);
    var month = parseInt(parts[1]);
    
    if (month === 1) {
      year -= 1;
      month = 12;
    } else {
      month -= 1;
    }
    
    var lastMonth = year + '-' + ('0' + month).slice(-2);
    
    // 查找上月結餘
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === lastMonth) {
        return Number(data[i][4]) || 0; // E欄：本月結餘
      }
    }
    
    return 0;
    
  } catch (e) {
    return 0;
  }
}

/**
 * 產生月結報表資料
 */
function generateChinaMonthlyReport(monthStr) {
  try {
    // 1. 取得上月結餘
    var carryOver = getLastMonthBalance(monthStr);
    
    // 2. 取得本月入帳
    var depositResult = getChinaDepositTotal(monthStr);
    var totalDeposit = depositResult.success ? depositResult.total : 0;
    
    // 3. 取得本月貨款
    var purchaseResult = getChinaPurchaseTotal(monthStr);
    var totalPurchase = purchaseResult.success ? purchaseResult.total : 0;
    
    // 4. 計算結餘
    var balance = carryOver + totalDeposit - totalPurchase;
    
    return {
      success: true,
      month: monthStr,
      carryOver: carryOver,
      totalDeposit: totalDeposit,
      depositCount: depositResult.count || 0,
      totalPurchase: totalPurchase,
      purchaseCount: purchaseResult.count || 0,
      purchaseItems: purchaseResult.items || [],
      balance: balance
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * 儲存月結記錄
 */
function saveChinaMonthlyRecord(data) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: "系統忙碌中" };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.MONTHLY);
    if (!sheet) {
      initChinaPaymentSheets();
      sheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.MONTHLY);
    }
    
    // 檢查是否已有該月份記錄
    var existingData = sheet.getDataRange().getValues();
    var existingRow = -1;
    
    for (var i = 1; i < existingData.length; i++) {
      if (String(existingData[i][0]) === data.month) {
        existingRow = i + 1;
        break;
      }
    }
    
    var nowStr = getNowString().split(' ')[0];
    var rowData = [
      data.month,
      data.carryOver,
      data.totalDeposit,
      data.totalPurchase,
      data.balance,
      nowStr,
      data.note || ''
    ];
    
    if (existingRow > 0) {
      // 更新現有記錄
      sheet.getRange(existingRow, 1, 1, 7).setValues([rowData]);
    } else {
      // 新增記錄
      sheet.appendRow(rowData);
    }
    
    return { success: true };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 取得所有月結歷史
 */
function getChinaMonthlyHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.MONTHLY);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, records: [] };
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    var records = data.map(function(row) {
      return {
        month: String(row[0]),
        carryOver: Number(row[1]) || 0,
        totalDeposit: Number(row[2]) || 0,
        totalPurchase: Number(row[3]) || 0,
        balance: Number(row[4]) || 0,
        settleDate: formatDate(row[5]),
        note: String(row[6] || '')
      };
    });
    
    // 按月份排序（新到舊）
    records.sort(function(a, b) {
      return b.month.localeCompare(a.month);
    });
    
    return { success: true, records: records };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
/**
 * 開啟大陸貨款管理視窗
 */
function showChinaPaymentPanel() {
  createDialog('ChinaPayment', 950, 800, '🇨🇳 大陸貨款管理');
}

/**
 * 初始化大陸貨款工作表（如果不存在就建立）
 */
function initChinaPaymentSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 建立入帳記錄表
  var depositSheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.DEPOSIT);
  if (!depositSheet) {
    depositSheet = ss.insertSheet(CHINA_CONFIG.SHEETS.DEPOSIT);
    depositSheet.appendRow(['編號', '日期', '人民幣金額', '匯率', '台幣金額', '備註', '建立時間']);
    depositSheet.getRange(1, 1, 1, 7).setBackground('#4a86e8').setFontColor('#fff').setFontWeight('bold');
  }
  
  // 建立月結記錄表
  var monthlySheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.MONTHLY);
  if (!monthlySheet) {
    monthlySheet = ss.insertSheet(CHINA_CONFIG.SHEETS.MONTHLY);
    monthlySheet.appendRow(['月份', '上月結轉', '本月入帳', '本月貨款', '本月結餘', '結算日期', '備註']);
    monthlySheet.getRange(1, 1, 1, 7).setBackground('#e06666').setFontColor('#fff').setFontWeight('bold');
  }
  
  return { success: true };
}

/**
 * 取得大陸貨款管理的初始資料
 */
function getChinaPaymentData() {
  try {
    initChinaPaymentSheets();
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 讀取入帳記錄
    var deposits = [];
    var depositSheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.DEPOSIT);
    if (depositSheet && depositSheet.getLastRow() > 1) {
      var data = depositSheet.getRange(2, 1, depositSheet.getLastRow() - 1, 7).getValues();
      deposits = data.map(function(row) {
        return {
          id: String(row[0]),
          date: formatDate(row[1]),
          rmbAmount: Number(row[2]) || 0,
          rate: Number(row[3]) || 0,
          twdAmount: Number(row[4]) || 0,
          note: String(row[5] || '')
        };
      }).reverse(); // 新的在上面
    }
    
    // 2. 讀取月結記錄
    var monthlyRecords = [];
    var monthlySheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.MONTHLY);
    if (monthlySheet && monthlySheet.getLastRow() > 1) {
      var data = monthlySheet.getRange(2, 1, monthlySheet.getLastRow() - 1, 7).getValues();
      monthlyRecords = data.map(function(row) {
        return {
          month: String(row[0]),
          carryOver: Number(row[1]) || 0,
          totalDeposit: Number(row[2]) || 0,
          totalPurchase: Number(row[3]) || 0,
          balance: Number(row[4]) || 0,
          settleDate: formatDate(row[5]),
          note: String(row[6] || '')
        };
      }).reverse();
    }
    
    // 3. 計算最新結餘（從月結記錄取最後一筆）
    var lastBalance = 0;
    if (monthlyRecords.length > 0) {
      lastBalance = monthlyRecords[0].balance;
    }
    
    return {
      success: true,
      deposits: deposits,
      monthlyRecords: monthlyRecords,
      lastBalance: lastBalance
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * 新增老闆入帳記錄
 */
function addChinaDeposit(data) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: "系統忙碌中" };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.DEPOSIT);
    if (!sheet) {
      initChinaPaymentSheets();
      sheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.DEPOSIT);
    }
    
    var id = generateId("CD", CHINA_CONFIG.SHEETS.DEPOSIT, 1);
    var rmbAmount = Number(data.rmbAmount) || 0;
    var rate = Number(data.rate) || 0;
    var twdAmount = Math.round(rmbAmount * rate);
    var nowStr = getNowString();
    
    sheet.appendRow([
      id,
      data.date,
      rmbAmount,
      rate,
      twdAmount,
      data.note || '',
      nowStr
    ]);
    
    return { success: true, id: id, twdAmount: twdAmount };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 刪除入帳記錄
 */
function deleteChinaDeposit(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CHINA_CONFIG.SHEETS.DEPOSIT);
    if (!sheet) return { success: false, error: "找不到工作表" };
    
    deleteRowsById(sheet, 1, id);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * 直接讀取外部團購購買表（不用 IMPORTRANGE）
 */
/**
 * 直接讀取外部團購購買表（修正版：支援 M/D 格式）
 */
function getChinaPurchaseTotal(monthStr) {
  try {
    var externalSS = SpreadsheetApp.openById("1Zk3YNuRJVIowEpIE95fxcp_D56KAz0l8LyBWD7pNF7E");
    var sheet = externalSS.getSheetByName("團購購買");
    
    if (!sheet) {
      return { success: false, error: "找不到團購購買工作表" };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, total: 0, count: 0, items: [] };
    }
    
    // 解析目標月份
    var targetParts = monthStr.split('-');
    var targetYear = parseInt(targetParts[0]);
    var targetMonth = parseInt(targetParts[1]);
    
    // 讀取 A欄(日期) 和 H欄(台幣)
    var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    var total = 0;
    var count = 0;
    var items = [];
    
    for (var i = 0; i < data.length; i++) {
      var dateVal = data[i][0];
      var rowMonth = -1;
      var rowDay = -1;
      
      // 解析日期
      if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
        // 如果是 Date 物件
        rowMonth = dateVal.getMonth() + 1;
        rowDay = dateVal.getDate();
      } else {
        // 如果是字串，嘗試解析 "M/D" 或 "M/D/Y" 格式
        var dateStr = String(dateVal).trim();
        if (!dateStr) continue;
        
        var parts = dateStr.split('/');
        if (parts.length >= 2) {
          rowMonth = parseInt(parts[0]);
          rowDay = parseInt(parts[1]);
        }
      }
      
      // 比對月份（假設都是同一年）
      if (rowMonth === targetMonth) {
        // H欄：台幣金額
        var twdVal = data[i][7];
        var twdAmount = 0;
        
        if (typeof twdVal === 'number') {
          twdAmount = twdVal;
        } else {
          var twdStr = String(twdVal).replace(/[,，$NT\s元]/g, '');
          twdAmount = parseFloat(twdStr) || 0;
        }
        
        if (twdAmount > 0) {
          total += twdAmount;
          count++;
          
          if (items.length < 50) {
            // E欄：人民幣金額
            var rmbVal = data[i][4];
            var rmbAmount = 0;
            if (typeof rmbVal === 'number') {
              rmbAmount = rmbVal;
            } else {
              rmbAmount = parseFloat(String(rmbVal).replace(/[,，]/g, '')) || 0;
            }
            
            items.push({
              date: targetYear + '/' + rowMonth + '/' + rowDay,
              orderNo: String(data[i][1] || ''),
              productName: String(data[i][3] || ''),
              rmbAmount: rmbAmount,
              twdAmount: twdAmount
            });
          }
        }
      }
    }
    
    return {
      success: true,
      total: Math.round(total),
      count: count,
      items: items
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
/** 
 * 合併後的單一選單：🚀 SQL 產生工具
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 產生 SQL 更新語法')
      .addItem('產生 Service 新增或修改語法 (新增優先)', 'generateSQLInsert')
      .addItem('產生 Fal 帳單 Insert 語法', 'generateFalExternalCostsSQL')
      .addItem('產生 Seedance 帳單 Insert 語法', 'generateSeedanceExternalCostsSQL')
      .addItem('產生 Runpod 帳單 Insert 語法 (New, 解析選取範圍)', 'generateRunpodExternalCostsSQL')
      .addItem('產生 Runpod 帳單 Insert 語法 (Legacy)', 'generateRunpodExternalCostsSQLLegacy')
      .addItem('產生 AWS SSM CLI for member limit 語法', 'generateAwsSsmCli')
      .addItem('產生 AWS SSM CLI for global limit 語法' (從 Finalized sheet), 'generateQueueGroupCli')      
      .addToUi();
}


function generateSQLInsert() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[2]; // 第 3 列 (Index 2)
  
  // 更新了欄位 Index
  var expectedHeaders = {
    0: "ServiceType", 1: "Provider", 2: "Queue Group", 3: "AP",
    4: "RLproductid", 5: "Dummy PID", 6: "Product Name",
    8: "ChargeType", 9: "Amount", 
    10: "Minlevel",   // 新增欄位 (對應 DB 的 MinMemberTier)
    11: "ToAdd",     // Index 順延
    12: "ToUpdate"   // Index 順延
  };

  var errors = [];
  for (var index in expectedHeaders) {
    var actual = (headers && headers[index]) ? headers[index].toString().trim() : "欄位不存在";
    if (actual !== expectedHeaders[index]) {
      var colLetter = String.fromCharCode(65 + parseInt(index));
      errors.push(`第 ${colLetter} 欄預期是 "${expectedHeaders[index]}"，但實際抓到 "${actual}"`);
    }
  }

  if (errors.length > 0) {
						  
    SpreadsheetApp.getUi().alert("❌ 欄位檢查失敗：\n\n" + errors.join("\n"));
    return; 
  }

  var sqlStatements = [];
  var lastValues = { st: "", prov: "", qg: "", ap: "", dpid: "", pn: "" };

  for (var i = 3; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length < 13) continue; // 確保有足夠的欄位長度

    // 依照新的 Index 抓取資料
    var toAddValue = row[11] ? row[11].toString().trim() : "";    
    var toUpdateValue = row[12] ? row[12].toString().trim() : ""; 

    if (row[0] !== "") lastValues.st = row[0];
    if (row[1] !== "") lastValues.prov = row[1];
    if (row[2] !== "") lastValues.qg = row[2];
    if (row[3] !== "") lastValues.ap = row[3];
    if (row[5] !== "") lastValues.dpid = row[5];
    if (row[6] !== "") lastValues.pn = row[6];

    if (toAddValue === "" && toUpdateValue === "") continue;

    var serviceType = (row[0] !== "" ? row[0] : lastValues.st).toString().trim();
    var provider    = (row[1] !== "" ? row[1] : lastValues.prov).toString().trim();
    var queueGroup  = (row[2] !== "" ? row[2] : lastValues.qg).toString().trim();
    var ap          = (row[3] !== "" ? row[3] : lastValues.ap).toString().trim();
    var rlProductId = (row[4] !== null && row[4] !== "") ? row[4].toString().trim() : "";
    var dummyPid    = (row[5] !== "" ? row[5] : lastValues.dpid).toString().trim();
    var productName = (row[6] !== "" ? row[6] : lastValues.pn).toString().trim();
    var chargeType  = (row[8] === "" || row[8] === null) ? 0 : row[8];
    var amount      = (row[9] === "" || row[9] === null) ? 0 : row[9];
    
    // 抓取 MinMemberTier，若為空值則預設為 1
    var minMemberTier = (row[10] === "" || row[10] === null) ? 1 : row[10];

    if (!rlProductId) continue;

    function clean(val) {
      if (!val) return "";
      return val.toString().replace(/'/g, "''");
    }

    var sql = "";
    if (toAddValue !== "") {
      // INSERT 語法更新：將原本寫死的 1 改為 minMemberTier 變數
      sql = "INSERT INTO [Reallusion].[dbo].[DA_External_Service] " +
            "([ServiceType], [QueueGroup], [Provider], [AP], [RLProductID], [DummyPID], [ProductName], [ChargeType], [Amount], [MinMemberTier], [HealthyStatus], [ServiceStatus]) " +
            "VALUES (" +
            "'" + clean(serviceType) + "', '" + clean(queueGroup) + "', '" + clean(provider) + "', '" + clean(ap) + "', " + 
            "'" + clean(rlProductId) + "', '" + clean(dummyPid) + "', '" + clean(productName) + "', " + 
            chargeType + ", " + amount + ", " + minMemberTier + ", 1, 1);";
    } 
    else if (toUpdateValue !== "") {
      // UPDATE 語法更新：加入 [MinMemberTier] 更新邏輯
      sql = "UPDATE [Reallusion].[dbo].[DA_External_Service] SET " +
            "[ServiceType] = '" + clean(serviceType) + "', " +
            "[QueueGroup] = '" + clean(queueGroup) + "', " +
            "[Provider] = '" + clean(provider) + "', " +
            "[AP] = '" + clean(ap) + "', " +
            "[DummyPID] = '" + clean(dummyPid) + "', " +
            "[ProductName] = '" + clean(productName) + "', " +
            "[ChargeType] = " + chargeType + ", " +
            "[Amount] = " + amount + ", " +
            "[MinMemberTier] = " + minMemberTier + " " +
            "WHERE [RLProductID] = '" + clean(rlProductId) + "';";
    }

    if (sql !== "") sqlStatements.push(sql);
  }

  if (sqlStatements.length > 0) {
    showOutputDialog(sqlStatements.join('\n\n'));
  } else {
    SpreadsheetApp.getUi().alert("未偵測到需要處理的資料。");
  }							
   																			   
								 
																		
}
	 
   
function generateFalExternalCostsSQL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Fal - source");
  const mappingSheet = ss.getSheetByName("Fal - mapping");
  
  if (!sourceSheet || !mappingSheet) {
    SpreadsheetApp.getUi().alert("找不到 'Fal - source' 或 'Fal - mapping' 工作表。");
    return;
  }

                           
  const mappingData = mappingSheet.getDataRange().getValues();
  let idToQueue = {};
  mappingData.slice(1).forEach(row => {
    let appId = String(row[0]).trim();
    if (appId) idToQueue[appId] = row[1];
  });
  
                         
  const sourceData = sourceSheet.getDataRange().getValues();
  let groupTotals = {};
  let unmatchedItems = [];

                                   
  sourceData.slice(1).forEach((row, index) => {
    const appId = String(row[0]).trim();
    const amount = parseFloat(row[1]);

                           
    if (!amount || isNaN(amount) || amount === 0) return;
        
    

                 
    const queueGroup = idToQueue[appId];
    if (queueGroup) {
      groupTotals[queueGroup] = (groupTotals[queueGroup] || 0) + amount;
    } else {
                                   
      unmatchedItems.push(`- ${appId}: $${amount} (列號: ${index + 2})`);
    }
  });

              
  if (unmatchedItems.length > 0) {
    // 修正 Alert 語法
    SpreadsheetApp.getUi().alert("⚠ 發現未對應項目\n\n" + unmatchedItems.join("\n"));
                                                             
    return;
  }

                  
                               
             
  const now = new Date();
              
  const lastDayOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);
                                    
                                       
                   
  const formattedDate = Utilities.formatDate(lastDayOfLastMonth, "GMT+8", "yyyy-MM-dd HH:mm:ss");


           
  let sqlStatements = [];
  let sqlRowsForSheet = [];  // 用來存完整 SQL 指令的二維陣列
  let previewRowsForSheet = []; // 新增：用來存資料預覽明細的二維陣列
  
  for (let queueGroup in groupTotals) {
    const totalAmount = parseFloat(groupTotals[queueGroup].toFixed(4));
    const sql = `INSERT INTO [Reallusion].[dbo].[DA_External_Costs] ([Time], [QueueGroup], [Amount], [Currency], [Provider]) VALUES ('${formattedDate}', '${queueGroup}', ${totalAmount.toFixed(4)}, 'USD', 'Fal.ai');`;
    
    sqlStatements.push(sql);
    sqlRowsForSheet.push([sql]); 
    
    // 儲存預覽明細：[Time, QueueGroup, Amount, Currency, Provider]
    previewRowsForSheet.push([formattedDate, queueGroup, totalAmount, 'USD', 'Fal.ai']);
  }

          
  if (sqlStatements.length > 0) {
    // 1. 保留原本功能：彈出視窗顯示 SQL
    showOutputDialog(sqlStatements.join('\n'));
    
    // ==== 寫入工作表下方 ====
    
    let currentLastRow = sourceSheet.getLastRow();
    
    // ----------------------------------------------------
    // 功能 A：建立「SQL 執行結果預覽表」(欄位拆開)
    // ----------------------------------------------------
    let previewStartRow = currentLastRow + 4; // 原資料下方空 3 行
    
    // 寫入大標題
    sourceSheet.getRange(previewStartRow, 1)
               .setValue("📊 SQL 執行結果預覽 (數據時間: " + formattedDate + ")")
               .setFontWeight("bold")
               .setBackground("#d9ead3"); // 綠色系大標題
               
    // 寫入欄位名稱 (第 1 欄到第 5 欄)
    const headers = [["[Time]", "[QueueGroup]", "[Amount]", "[Currency]", "[Provider]"]];
    sourceSheet.getRange(previewStartRow + 1, 1, 1, 5)
               .setValues(headers)
               .setFontWeight("bold")
               .setBackground("#f3f3f3");
               
    // 寫入預覽資料明細
    sourceSheet.getRange(previewStartRow + 2, 1, previewRowsForSheet.length, 5)
               .setValues(previewRowsForSheet);
    
    // ----------------------------------------------------
    // 功能 B：建立「SQL 完整指令列表」(方便整包複製)
    // ----------------------------------------------------
    // 重新計算目前的最後一行（因為剛剛填了預覽表）
    currentLastRow = sourceSheet.getLastRow();
    let sqlStartRow = currentLastRow + 3; // 預覽表下方空 2 行
    
    // 寫入 SQL 指令大標題
    sourceSheet.getRange(sqlStartRow, 1)
               .setValue("📜 產生的 SQL 原始指令列表")
               .setFontWeight("bold")
               .setBackground("#e6f2ff"); // 藍色系大標題
    
    // 寫入所有 SQL 指令
    sourceSheet.getRange(sqlStartRow + 1, 1, sqlRowsForSheet.length, 1)
               .setValues(sqlRowsForSheet);
  }
}

   
													
function generateSeedanceExternalCostsSQL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Seedance - source");
  const mappingSheet = ss.getSheetByName("Seedance - mapping");
  
  if (!sourceSheet || !mappingSheet) {
    SpreadsheetApp.getUi().alert("找不到 'Seedance - source' 或 'Seedance - mapping' 工作表。");
    return;
  }

  // 讀取 mapping 表：appId → queueGroup
  const mappingData = mappingSheet.getDataRange().getValues();
  let idToQueue = {};
  mappingData.slice(1).forEach(row => {
    let appId = String(row[0]).trim();
    if (appId) idToQueue[appId] = row[1];
  });

  // 讀取 source 表，按 queueGroup 加總金額
  const sourceData = sourceSheet.getDataRange().getValues();
  let groupTotals = {};
  let unmatchedItems = [];

  sourceData.slice(1).forEach((row, index) => {
    const appId = String(row[0]).trim();
    const amount = parseFloat(row[1]);

    if (!amount || isNaN(amount) || amount === 0) return;

    const queueGroup = idToQueue[appId];
    if (queueGroup) {
      groupTotals[queueGroup] = (groupTotals[queueGroup] || 0) + amount;
    } else {
      unmatchedItems.push(`- ${appId}: $${amount} (列號: ${index + 2})`);
    }
  });

  if (unmatchedItems.length > 0) {
    SpreadsheetApp.getUi().alert("⚠ 發現未對應項目\n\n" + unmatchedItems.join("\n"));
    return;
  }

  // 計算上個月最後一天 23:59:59
  const now = new Date();
  const lastDayOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);
  const formattedDate = Utilities.formatDate(lastDayOfLastMonth, "GMT+8", "yyyy-MM-dd HH:mm:ss");

  let sqlStatements = [];
  let sqlRowsForSheet = [];
  let previewRowsForSheet = [];

  for (let queueGroup in groupTotals) {
    const totalAmount = parseFloat(groupTotals[queueGroup].toFixed(4));
    const sql = `INSERT INTO [Reallusion].[dbo].[DA_External_Costs] ([Time], [QueueGroup], [Amount], [Currency], [Provider]) VALUES ('${formattedDate}', '${queueGroup}', ${totalAmount.toFixed(4)}, 'USD', 'Seedance');`;

    sqlStatements.push(sql);
    sqlRowsForSheet.push([sql]);

    // 儲存預覽明細：[Time, QueueGroup, Amount, Currency, Provider]
    previewRowsForSheet.push([formattedDate, queueGroup, totalAmount, 'USD', 'Seedance']);
  }

  if (sqlStatements.length > 0) {
    // 1. 彈出視窗顯示 SQL
    showOutputDialog(sqlStatements.join('\n'));

    // ==== 寫入 source 工作表下方 ====
    let currentLastRow = sourceSheet.getLastRow();

    // 功能 A：建立「SQL 執行結果預覽表」(欄位拆開)
    let previewStartRow = currentLastRow + 4; // 原資料下方空 3 行

    sourceSheet.getRange(previewStartRow, 1)
               .setValue("📊 SQL 執行結果預覽 (數據時間: " + formattedDate + ")")
               .setFontWeight("bold")
               .setBackground("#d9ead3");

    const headers = [["[Time]", "[QueueGroup]", "[Amount]", "[Currency]", "[Provider]"]];
    sourceSheet.getRange(previewStartRow + 1, 1, 1, 5)
               .setValues(headers)
               .setFontWeight("bold")
               .setBackground("#f3f3f3");

    sourceSheet.getRange(previewStartRow + 2, 1, previewRowsForSheet.length, 5)
               .setValues(previewRowsForSheet);

    // 功能 B：建立「SQL 完整指令列表」(方便整包複製)
    currentLastRow = sourceSheet.getLastRow();
    let sqlStartRow = currentLastRow + 3;

    sourceSheet.getRange(sqlStartRow, 1)
               .setValue("📜 產生的 SQL 原始指令列表")
               .setFontWeight("bold")
               .setBackground("#e6f2ff");

    sourceSheet.getRange(sqlStartRow + 1, 1, sqlRowsForSheet.length, 1)
               .setValues(sqlRowsForSheet);
  }
}

   
													
/** (Legacy) 用 end-point mapping */   
function generateRunpodExternalCostsSQLLegacy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Runpod- source");
  const mappingSheet = ss.getSheetByName("Runpod- mapping");
  
  if (!sourceSheet || !mappingSheet) {
    SpreadsheetApp.getUi().alert("找不到 'Runpod- source' 或 'Runpod- mapping' 工作表。");
    return;
  }

															   
  const mappingData = mappingSheet.getDataRange().getValues();
  let idToQueue = {};
  mappingData.slice(1).forEach(row => {
    let instanceId = String(row[0]).trim();
    if (instanceId) idToQueue[instanceId] = row[1];
  });

									
  const sourceData = sourceSheet.getDataRange().getValues();
																				  
  const instanceRow = sourceData[5]; 
  const amountRow = sourceData[6];
  
  let groupTotals = {};
  let unmatchedItems = [];

										   
  for (let col = 2; col < instanceRow.length; col++) {
    const instanceId = String(instanceRow[col]).trim();
    const amountStr = String(amountRow[col]).replace(/[$,]/g, ""); 
    const amount = parseFloat(amountStr);

    if (!instanceId || !amount || isNaN(amount) || amount === 0) continue;

    const queueGroup = idToQueue[instanceId];
    if (queueGroup) {
      groupTotals[queueGroup] = (groupTotals[queueGroup] || 0) + amount;
    } else {
      unmatchedItems.push(`- ${instanceId}: $${amount}`);
    }
  }

  if (unmatchedItems.length > 0) {
    // 修正 Alert 語法
    SpreadsheetApp.getUi().alert("⚠ 發現未對應項目\n\n請補上 Mapping：\n\n" + unmatchedItems.join("\n"));
    return;
  }

													  
															 
  const now = new Date();
  const lastDay = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);
  const formattedDate = Utilities.formatDate(lastDay, "GMT+8", "yyyy-MM-dd HH:mm:ss");

  let sqlStatements = [];
  for (let queueGroup in groupTotals) {
    const totalAmount = groupTotals[queueGroup].toFixed(4);
																									  
    sqlStatements.push(`INSERT INTO [Reallusion].[dbo].[DA_External_Costs] ([Time], [QueueGroup], [Amount], [Currency], [Provider]) VALUES ('${formattedDate}', '${queueGroup}', ${totalAmount}, 'USD', 'Runpod');`);
							
  }

  if (sqlStatements.length > 0) {
    showOutputDialog(sqlStatements.join('\n'));
  } else {
    SpreadsheetApp.getUi().alert("沒有符合條件的資料。");
  }
}

function showOutputDialog(text) {
  var html = HtmlService.createHtmlOutput(
    '<textarea style="width:100%; height:350px; font-family:Consolas, monospace; font-size:12px; padding:10px;">' + text + '</textarea>'
  ).setWidth(800).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '📦 SQL 產出結果');
}



/**
 * 功能 2：產生 Runpod 帳單 SQL (從目前選取的儲存格解析)
 * 修正：時間改為「上個月最後一天 23:59:59」
 */
function generateRunpodExternalCostsSQL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet(); // 新增：獲取當前正在操作的工作表
  const range = ss.getActiveRange();
  const data = range.getValues();
  let sqlStatements = [];

  // --- 計算上個月最後一天 23:59:59 ---
  const now = new Date();
  // 將日期設為本月第 0 天，自動會變成「上個月最後一天」
  const lastDayOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0);
  // 設定時間為 23:59:59
  lastDayOfLastMonth.setHours(23, 59, 59);
  
  // 格式化為 SQL 字串
  const timeString = Utilities.formatDate(lastDayOfLastMonth, "GMT+8", "yyyy-MM-dd HH:mm:ss");

  const queueRegex = /Queue\s+(\d+)/;
  const amountRegex = /^\$(\d+\.?\d*)$/;

  // 將所有儲存格的文字攤平成一個陣列，方便跨格讀取「下一行」
  const lines = [];
  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].length; j++) {
      const text = String(data[i][j]).trim();
      if (text) lines.push(text);
    }
  }

  let previewRowsForSheet = []; // 新增：用來存資料預覽明細的二維陣列
  let sqlRowsForSheet = [];      // 新增：用來存完整 SQL 指令的二維陣列

  for (let k = 0; k < lines.length; k++) {
    const line = lines[k];

    // 新格式：Queue/名稱 在這行，金額在下一行
    const matchQueue = line.match(queueRegex);
    const nextLine = lines[k + 1] ? lines[k + 1].trim() : "";
    const matchNextAmount = nextLine.match(amountRegex);

    if (matchNextAmount) {
      // 有對應金額的情況
      const amount = parseFloat(matchNextAmount[1]);
      const queueGroup = matchQueue ? matchQueue[1] : "999";
      
      const sql = `INSERT INTO [Reallusion].[dbo].[DA_External_Costs] ([Time], [QueueGroup], [Amount], [Currency], [Provider]) ` +
                  `VALUES ('${timeString}', '${queueGroup}', ${amount.toFixed(4)}, 'USD', 'Runpod');`;
      
      sqlStatements.push(sql);
      sqlRowsForSheet.push([sql]); // 儲存成試算表專用二維格式
      
      // 儲存預覽明細：[Time, QueueGroup, Amount, Currency, Provider]
      previewRowsForSheet.push([timeString, queueGroup, amount, 'USD', 'Runpod']);
      
      k++; // 跳過已處理的金額行
    }
  }

  if (sqlStatements.length > 0) {
    // 1. 保留原本功能：彈出視窗顯示 SQL (改呼叫你原始程式碼的 showOutputModal)
    showOutputModal('Runpod SQL 結果 (上個月底)', sqlStatements.join('\n'));
    
    // ==== 寫入目前操作的工作表最下方 ====
    
    let currentLastRow = currentSheet.getLastRow();
    
    // ----------------------------------------------------
    // 功能 A：建立「SQL 執行結果預覽表」(欄位拆開)
    // ----------------------------------------------------
    let previewStartRow = currentLastRow + 4; // 原始資料下方空 3 行
    
    // 寫入大標題
    currentSheet.getRange(previewStartRow, 1)
                .setValue("📊 SQL 執行結果預覽 (數據時間: " + timeString + ")")
                .setFontWeight("bold")
                .setBackground("#d9ead3"); // 綠色系大標題
               
    // 寫入欄位名稱 (第 1 欄到第 5 欄)
    const headers = [["[Time]", "[QueueGroup]", "[Amount]", "[Currency]", "[Provider]"]];
    currentSheet.getRange(previewStartRow + 1, 1, 1, 5)
                .setValues(headers)
                .setFontWeight("bold")
                .setBackground("#f3f3f3");
               
    // 寫入預覽資料明細
    currentSheet.getRange(previewStartRow + 2, 1, previewRowsForSheet.length, 5)
                .setValues(previewRowsForSheet);
               
    // ----------------------------------------------------
    // 功能 B：建立「SQL 完整指令列表」(方便整包複製)
    // ----------------------------------------------------
    // 重新計算目前的最後一行（因為剛剛填了預覽表）
    currentLastRow = currentSheet.getLastRow();
    let sqlStartRow = currentLastRow + 3; // 預覽表下方空 2 行
    
    // 寫入 SQL 指令大標題
    currentSheet.getRange(sqlStartRow, 1)
                .setValue("📜 產生的 SQL 原始指令列表")
                .setFontWeight("bold")
                .setBackground("#e6f2ff"); // 藍色系大標題
    
    // 寫入所有 SQL 指令
    currentSheet.getRange(sqlStartRow + 1, 1, sqlRowsForSheet.length, 1)
                .setValues(sqlRowsForSheet);

  } else {
    SpreadsheetApp.getUi().alert("未能在選取範圍內解析出資料。");
  }
}

/**
 * 統一顯示彈窗函式 (重要：確保此函式存在)
 */
function showOutputModal(title, content) {
  const html = HtmlService.createHtmlOutput(
    '<textarea style="width:100%; height:420px; font-family: Consolas, monospace; font-size: 11px; padding: 10px; background-color: #f4f4f4; border: 1px solid #ccc;">' + content + '</textarea>'
  )
  .setWidth(900)
  .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}


/**
 * 產生 AWS SSM Parameter Store 的 CLI 語法
 * 修正點：V 欄對應 Free, W 欄對應 Paid (修正先前放反的問題)
 */
function generateAwsSsmCli() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  let allCliCommands = [];
  const environments = ["dev", "stage", "live", "dev4"];

  // 從第 4 列開始 (Index 3)，截圖顯示資料從此開始
  for (let i = 3; i < data.length; i++) {
    const serviceType = String(data[i][0]).trim(); // A 欄 (Index 0)
    const freeValueRaw = data[i][22];              // W 欄 (Index 22) -> Free
    const paidValueRaw = data[i][23];              // X 欄 (Index 23) -> Paid

    // 只要 W 或 X 其中一個有值，且 ServiceType 不為空才處理
    if (!serviceType || (paidValueRaw === "" && freeValueRaw === "")) {
      continue;
    }

    // 強制轉數字處理，如果非數字則設為 null
    const paidValue = (paidValueRaw !== "" && !isNaN(parseFloat(paidValueRaw))) ? parseFloat(paidValueRaw) : null;
    const freeValue = (freeValueRaw !== "" && !isNaN(parseFloat(freeValueRaw))) ? parseFloat(freeValueRaw) : null;

    let serviceSection = [`# Service${serviceType}`];

    environments.forEach(env => {
      serviceSection.push(`# ========== ${env.toUpperCase()} 環境 ==========`);
      serviceSection.push(`echo "設定 ${env.toUpperCase()} 環境..."`);
      
      // 處理 Paid 參數 (對應原本 X 欄的值)
      if (paidValue !== null) {
        serviceSection.push(`aws ssm put-parameter --name "/RLAdaptor/${env}/ConcurrencyLimit/Service${serviceType}/Paid" --value "${paidValue}" --type "String" --overwrite`);
      }
      
      // 處理 Free 參數 (對應原本 W 欄的值)
      if (freeValue !== null) {
        serviceSection.push(`aws ssm put-parameter --name "/RLAdaptor/${env}/ConcurrencyLimit/Service${serviceType}/Free" --value "${freeValue}" --type "String" --overwrite`);
      }
      
      serviceSection.push(`echo "${env.toUpperCase()} 環境設定完成"`);
    });

    allCliCommands.push(serviceSection.join('\n'));
  }

  // 輸出結果
  if (allCliCommands.length > 0) {
    const finalOutput = allCliCommands.join('\n\n');
    showAwsCliModal(finalOutput);
  } else {
    SpreadsheetApp.getUi().alert("找不到符合條件的資料，請檢查 A 欄是否有 ServiceType，且 W 或 X 欄是否有數值。");
  }
}

/**
 * 專屬 AWS CLI 的顯示視窗
 */
function showAwsCliModal(content) {
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family: sans-serif;">產出的 AWS CLI 語法如下（包含所有環境）：</p>' +
    '<textarea style="width:100%; height:400px; font-family: Consolas, monospace; font-size: 11px; padding: 10px; background-color: #f4f4f4;">' + content + '</textarea>'
  )
  .setWidth(850)
  .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'AWS SSM CLI 語法產出');
}


/**
 * 產生 QueueGroup SSM Parameter Store 的 CLI 語法
 * 依據 M 欄 (Queue Group ID) 與 T 欄 (Global Value) 產出
 */
function generateQueueGroupCli() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  let allCliCommands = [];
  const environments = ["dev", "stage", "live", "dev4"];

  // 根據截圖，資料從第 4 列開始 (Index 3)
  for (let i = 3; i < data.length; i++) {
    const queueGroupId = String(data[i][14]).trim(); // O 欄 (Index 14)
    const globalValueRaw = data[i][21];              // V 欄 (Index 21)

    // 檢查 O 欄位是否有值且為數字，且 V 欄位不為空
    if (!queueGroupId || isNaN(queueGroupId) || globalValueRaw === "") {
      continue;
    }

    const globalValue = !isNaN(parseFloat(globalValueRaw)) ? parseFloat(globalValueRaw) : null;
    if (globalValue === null) continue;

    let section = [`# QueueGroup${queueGroupId}`];

    environments.forEach(env => {
      section.push(`# ========== ${env.toUpperCase()} 環境 ==========`);
      section.push(`echo "設定 ${env.toUpperCase()} 環境..."`);
      
      // 產生指令路徑：/RLAdaptor/[env]/ConcurrencyLimit/QueueGroup[ID]/Global
      section.push(`aws ssm put-parameter --name "/RLAdaptor/${env}/ConcurrencyLimit/QueueGroup${queueGroupId}/Global" --value "${globalValue}" --type "String" --overwrite`);
      
      section.push(`echo "${env.toUpperCase()} 環境設定完成"`);
    });

    allCliCommands.push(section.join('\n'));
  }

  // 輸出結果
  if (allCliCommands.length > 0) {
    const header = "# AWS Parameter Store 設定命令 - QueueGroup ConcurrencyLimit 配置\n# RLAdaptor QueueGroup 併發限制參數設定\n\n";
    showAwsCliModal(header + allCliCommands.join('\n\n'));
  } else {
    SpreadsheetApp.getUi().alert("找不到符合條件的資料，請檢查 O 欄是否為 QueueGroup ID 數字，且 V 欄是否有對應數值。");
  }
}


// Menu and Dropdown list
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Inventory')
    .addItem('Update Product', 'updateData')
    .addItem('Create New Product', 'createNewProduct')
    .addToUi();

  // Create dropdown list in Data Form Sheet - D4
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data-Form");
  var columns = [
    "موجودی اولیه",
    "مرجوعی",
    "اصلاح شده",
    "تولید فروردین",
    "فروش فروردین",
    "تولید اردیبهشت",
    "فروش اردیبهشت",
    "تولید خرداد",
    "فروش خرداد",
    "تولید تیر",
    "فروش تیر",
    "تولید مرداد",
    "فروش مرداد",
    "تولید شهریور",
    "فروش شهریور",
    "تولید مهر",
    "فروش مهر",
    "تولید آبان",
    "فروش آبان",
    "تولید آذر",
    "فروش آذر",
    "تولید دی",
    "فروش دی",
    "تولید بهمن",
    "فروش بهمن",
    "تولید اسفند",
    "فروش اسفند"
  ];

  var range = sheet.getRange("D4");
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(columns).build();
  range.setDataValidation(rule);
}

// Function to create new product
function createNewProduct() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName("Data-Form");
    var dataSheet = ss.getSheetByName("Product-Details");
    
    if (!formSheet || !dataSheet) {
      throw new Error("یکی از شیت‌ها یافت نشد");
    }
    
    // Get information from the form
    var newProductId = formSheet.getRange("I3").getValue();
    var newProductTitle = formSheet.getRange("I4:I5").getValue();
    
    // Input validation
    if (!newProductId) throw new Error("شناسه محصول جدید را وارد کنید");
    if (!newProductTitle) throw new Error("عنوان محصول جدید را وارد کنید");
    
    // Check for duplicate ID
    var existingIds = dataSheet.getRange("A:A").getValues().flat();
    if (existingIds.includes(newProductId)) {
      throw new Error("شناسه محصول تکراری است");
    }
    
    // Find the first empty row
    var lastRow = dataSheet.getLastRow();
    var emptyRow = lastRow + 1;
    for (var i = 1; i <= lastRow; i++) {
      if (!dataSheet.getRange(i, 1).getValue()) {
        emptyRow = i;
        break;
      }
    }
    
   // Add new product
    dataSheet.getRange(emptyRow, 1).setValue(newProductId);
    dataSheet.getRange(emptyRow, 2).setValue(newProductTitle);
    
    // Log report
    logActivity(
      "افزودن محصول", 
      newProductId, 
      newProductTitle
    );
    
    // Clear form
    formSheet.getRange("I3:I4").clearContent();
    
    SpreadsheetApp.getUi().alert("محصول جدید با موفقیت ثبت شد و گزارش ذخیره گردید ✅");
  } catch (error) {
    SpreadsheetApp.getUi().alert("خطا: ❌ " + error.message);
    console.error(error);
  }
}

// Update Product
function updateData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName("Data-Form");
    var dataSheet = ss.getSheetByName("Product-Details");
    
    if (!formSheet || !dataSheet) {
      throw new Error("یکی از شیت‌ها یافت نشد");
    }
    
    // Get information from the form
    var productId = formSheet.getRange("D3").getValue();
    var field = formSheet.getRange("D4").getValue();
    var valueToAdd = formSheet.getRange("D5").getValue();
    
    // Input validation
    if (!productId) throw new Error("شناسه محصول را وارد کنید");
    if (!field) throw new Error("فیلد مورد نظر را انتخاب کنید");
    if (!valueToAdd && valueToAdd !== 0) throw new Error("مقدار عددی را وارد کنید");
    
    // Find product
    var headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    var columnIndex = headers.indexOf(field) + 1;
    if (columnIndex === 0) throw new Error("ستون '" + field + "' یافت نشد");
    
    // Find product row
    var idData = dataSheet.getRange("A:A").getValues();
    var rowIndex = -1;
    for (var i = 0; i < idData.length; i++) {
      if (idData[i][0] == productId) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw new Error("محصول با شناسه '" + productId + "' یافت نشد");
    
    // Update value
    var cell = dataSheet.getRange(rowIndex, columnIndex);
    var currentValue = cell.getValue() || 0;
    cell.setValue(Number(currentValue) + Number(valueToAdd));
    
    // Log report
    var productTitle = dataSheet.getRange(rowIndex, 2).getValue();
    logActivity(
      "آپدیت محصول", 
      productId, 
      productTitle, 
      field, 
      valueToAdd
    );
    
    // Clear form
    formSheet.getRange("D3:D5").clearContent();
    
    SpreadsheetApp.getUi().alert("اطلاعات با موفقیت بروزرسانی شد و گزارش ثبت گردید ✅");
  } catch (error) {
    SpreadsheetApp.getUi().alert("خطا: ❌ " + error.message);
    console.error(error);
  }
}

// Add report to logs sheet
function logActivity(operationType, productId, productTitle, field, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("logs");
  
  // If logs sheet not found, create it with new column
  if (!logSheet) {
    logSheet = ss.insertSheet("logs");
    logSheet.appendRow([
      "نوع عملیات", 
      "تاریخ میلادی", 
      "تاریخ شمسی",
      "ساعت", 
      "شناسه محصول", 
      "عنوان محصول", 
      "موضوع عملیات", 
      "تعداد"
    ]);
  }
  
  // Current date and time
  var now = new Date();
  var date = Utilities.formatDate(now, "Asia/Tehran", "yyyy/MM/dd");
  var jalaliDate = JALALI(now); // Convert to Jalali
  var time = Utilities.formatDate(now, "Asia/Tehran", "HH:mm:ss");
  
  // Add log to new row
  logSheet.appendRow([
    operationType,
    date,
    jalaliDate,
    time,
    productId,
    productTitle || "",
    field || "",
    value || ""
  ]);
}

// Jalali date conversion function
function JALALI(date) {
  let jalaliFormat = new Date(date).toLocaleDateString('fa-IR').replace(/([۰-۹])/g, token => String.fromCharCode(token.charCodeAt(0) - 1728));
  return jalaliFormat;
}
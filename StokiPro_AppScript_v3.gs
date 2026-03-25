// ================================================================
// STOKIPRO v3 — GOOGLE APPSCRIPT BACKEND
// Tables: Tenants, Users, Products, Sales, SalesItems,
//         Purchases, PurchaseItems, Customers, Suppliers,
//         CustomerPayments, SupplierPayments, BizSettings, Activity
// ================================================================
// MAELEKEZO:
// 1. script.google.com → New Project → Paste code hii yote
// 2. Run setupSheets() KWANZA (mara moja tu)
// 3. Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Copy URL → weka kwenye const BACKEND kwenye HTML
// ================================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ================================================================
// ENTRY POINT
// ================================================================
function doGet(e) {
  var result;
  try {
    if (e.parameter && e.parameter.payload) {
      var body = JSON.parse(decodeURIComponent(e.parameter.payload));
      result = route(body);
    } else {
      result = { success: true, message: 'StokiPro v3 Backend Online ✅' };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    return ContentService.createTextOutput(JSON.stringify(route(body)))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ================================================================
// ROUTER
// ================================================================
function route(body) {
  var action = body.action || '';
  switch (action) {
    case 'ping':                return { success: true, message: 'StokiPro v3 Online ✅', version: '3.0' };
    // Auth
    case 'get_all':             return getAll(body.store);
    case 'get_by_tenant':       return getByTenant(body.store, body.tenantId);
    // Core saves — proper tables
    case 'save_tenant':         return saveRow('Tenants', body.data, tenantFields());
    case 'save_user':           return saveRow('Users', body.data, userFields());
    case 'save_product':        return saveRow('Products', body.data, productFields());
    case 'save_sale':           return saveSale(body.data);
    case 'save_purchase':       return savePurchase(body.data);
    case 'save_customer':       return saveRow('Customers', body.data, customerFields());
    case 'save_supplier':       return saveRow('Suppliers', body.data, supplierFields());
    case 'save_customer_payment': return saveCustomerPayment(body.data);
    case 'save_supplier_payment': return saveSupplierPayment(body.data);
    // Settings
    case 'save_biz_settings':   return saveBizSettings(body.data);
    case 'get_biz_settings':    return getBizSettings(body.tenantId);
    // Pull ALL data kwa tenant (ndio sync mkuu)
    case 'pull':                return pullAll(body.tenantId);
    // Bulk
    case 'bulk_save':           return bulkSave(body.store, body.data);
    // Delete
    case 'delete':              return deleteRow(body.store, body.id);
    // Email
    case 'send_alert':          return sendAlert(body);
    case 'forgot_password':     return forgotPassword(body.email);
    // Generic save (fallback)
    case 'save':                return genericSave(body.store, body.data);
    default:                    return { success: false, error: 'Unknown action: ' + action };
  }
}

// ================================================================
// PULL ALL — Rudisha kila kitu kwa tenant (ndio sync mkuu)
// ================================================================
function pullAll(tenantId) {
  if (!tenantId) return { success: false, error: 'tenantId required' };
  try {
    // Sales na items zao
    var salesRows = filterByTenant('Sales', tenantId);
    var saleIds = salesRows.map(function(s){ return s.id; });
    var allSalesItems = getSheet('SalesItems').getLastRow() > 1
      ? sheetToObjects('SalesItems').filter(function(it){ return saleIds.indexOf(it.saleId) >= 0; })
      : [];
    // Attach items kwa kila sale
    var sales = salesRows.map(function(s) {
      s.items = allSalesItems.filter(function(it){ return it.saleId === s.id; });
      return s;
    });

    // Purchases na items zao
    var purchRows = filterByTenant('Purchases', tenantId);
    var purchIds = purchRows.map(function(p){ return p.id; });
    var allPurchItems = getSheet('PurchaseItems').getLastRow() > 1
      ? sheetToObjects('PurchaseItems').filter(function(it){ return purchIds.indexOf(it.purchId) >= 0; })
      : [];
    var purchases = purchRows.map(function(p) {
      p.items = allPurchItems.filter(function(it){ return it.purchId === p.id; });
      return p;
    });

    // Customers na payment history
    var customerRows = filterByTenant('Customers', tenantId);
    var custIds = customerRows.map(function(c){ return c.id; });
    var allCustPay = getSheet('CustomerPayments').getLastRow() > 1
      ? sheetToObjects('CustomerPayments').filter(function(p){ return custIds.indexOf(p.custId) >= 0; })
      : [];
    var customers = customerRows.map(function(c) {
      c.payments = allCustPay.filter(function(p){ return p.custId === c.id; });
      return c;
    });

    // Suppliers na payment history
    var supplierRows = filterByTenant('Suppliers', tenantId);
    var supIds = supplierRows.map(function(s){ return s.id; });
    var allSupPay = getSheet('SupplierPayments').getLastRow() > 1
      ? sheetToObjects('SupplierPayments').filter(function(p){ return supIds.indexOf(p.supId) >= 0; })
      : [];
    var suppliers = supplierRows.map(function(s) {
      s.payments = allSupPay.filter(function(p){ return p.supId === s.id; });
      return s;
    });

    return {
      success: true,
      data: {
        products:  filterByTenant('Products', tenantId),
        sales:     sales,
        purchases: purchases,
        customers: customers,
        suppliers: suppliers,
        settings:  getBizSettingsRaw(tenantId),
      },
      timestamp: new Date().toISOString()
    };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ================================================================
// SAVE SALE — Sales table + SalesItems table
// ================================================================
function saveSale(data) {
  if (!data || !data.id) return { success: false, error: 'id required' };
  try {
    // Upsert Sales row
    var fields = saleFields();
    var existing = getRowById('Sales', data.id);
    if (existing) {
      updateRow('Sales', data.id, data, fields);
    } else {
      appendRow('Sales', data, fields);
    }
    // Upsert SalesItems — futa za zamani, ongeza mpya
    if (data.items && data.items.length > 0) {
      deleteItemsBySaleId('SalesItems', data.id);
      data.items.forEach(function(it) {
        appendToSheet('SalesItems', [
          data.id, data.tenantId||'',
          it.pid||'', it.pname||'',
          Number(it.qty)||0, Number(it.price)||0, Number(it.total)||0
        ]);
      });
    }
    return { success: true };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ================================================================
// SAVE PURCHASE — Purchases table + PurchaseItems table
// ================================================================
function savePurchase(data) {
  if (!data || !data.id) return { success: false, error: 'id required' };
  try {
    var fields = purchaseFields();
    var existing = getRowById('Purchases', data.id);
    if (existing) {
      updateRow('Purchases', data.id, data, fields);
    } else {
      appendRow('Purchases', data, fields);
    }
    if (data.items && data.items.length > 0) {
      deleteItemsByPurchId('PurchaseItems', data.id);
      data.items.forEach(function(it) {
        appendToSheet('PurchaseItems', [
          data.id, data.tenantId||'',
          it.pid||'', it.pname||'',
          Number(it.qty)||0, Number(it.price)||0, Number(it.total)||0
        ]);
      });
    }
    return { success: true };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ================================================================
// CUSTOMER PAYMENT — CustomerPayments table
// ================================================================
function saveCustomerPayment(data) {
  if (!data || !data.id) return { success: false, error: 'id required' };
  try {
    var existing = getRowById('CustomerPayments', data.id);
    if (!existing) {
      appendRow('CustomerPayments', data, customerPaymentFields());
    }
    // Also update customer currentDebt
    if (data.custId && data.debtAfter !== undefined) {
      var custSheet = getSheet('Customers');
      var rows = custSheet.getDataRange().getValues();
      var headers = rows[0];
      var idIdx = headers.indexOf('id');
      var debtIdx = headers.indexOf('currentDebt');
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][idIdx] === data.custId) {
          custSheet.getRange(i+1, debtIdx+1).setValue(Number(data.debtAfter)||0);
          break;
        }
      }
    }
    return { success: true };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ================================================================
// SUPPLIER PAYMENT — SupplierPayments table
// ================================================================
function saveSupplierPayment(data) {
  if (!data || !data.id) return { success: false, error: 'id required' };
  try {
    var existing = getRowById('SupplierPayments', data.id);
    if (!existing) {
      appendRow('SupplierPayments', data, supplierPaymentFields());
    }
    // Also update supplier currentDebt
    if (data.supId && data.debtAfter !== undefined) {
      var supSheet = getSheet('Suppliers');
      var rows = supSheet.getDataRange().getValues();
      var headers = rows[0];
      var idIdx = headers.indexOf('id');
      var debtIdx = headers.indexOf('currentDebt');
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][idIdx] === data.supId) {
          supSheet.getRange(i+1, debtIdx+1).setValue(Number(data.debtAfter)||0);
          break;
        }
      }
    }
    return { success: true };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ================================================================
// GENERIC SAVE (fallback kwa tenants, users, etc.)
// ================================================================
function genericSave(store, data) {
  if (!store || !data || !data.id) return { success: false, error: 'store/data/id required' };
  var storeMap = {
    'tenants': 'Tenants', 'users': 'Users', 'products': 'Products',
    'customers': 'Customers', 'suppliers': 'Suppliers',
    'sales': 'Sales', 'purchases': 'Purchases', 'activity': 'Activity'
  };
  var sheetName = storeMap[store] || store;
  try {
    // Route to specific save if available
    if (store === 'sales') return saveSale(data);
    if (store === 'purchases') return savePurchase(data);
    var fields = getFields(sheetName);
    var existing = getRowById(sheetName, data.id);
    if (existing) {
      updateRow(sheetName, data.id, data, fields);
    } else {
      appendRow(sheetName, data, fields);
    }
    return { success: true };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ================================================================
// GET ALL / GET BY TENANT
// ================================================================
function getAll(store) {
  if (!store) return { success: false, error: 'store required' };
  var storeMap = {
    'tenants':'Tenants','users':'Users','products':'Products',
    'sales':'Sales','purchases':'Purchases',
    'customers':'Customers','suppliers':'Suppliers','activity':'Activity'
  };
  var sheetName = storeMap[store] || store;
  try {
    return { success: true, data: sheetToObjects(sheetName) };
  } catch(err) {
    return { success: false, error: err.toString(), data: [] };
  }
}

function getByTenant(store, tenantId) {
  if (!store || !tenantId) return { success: false, error: 'store/tenantId required', data: [] };
  var storeMap = {
    'tenants':'Tenants','users':'Users','products':'Products',
    'sales':'Sales','purchases':'Purchases',
    'customers':'Customers','suppliers':'Suppliers','activity':'Activity'
  };
  var sheetName = storeMap[store] || store;
  try {
    return { success: true, data: filterByTenant(sheetName, tenantId) };
  } catch(err) {
    return { success: false, error: err.toString(), data: [] };
  }
}

// ================================================================
// BIZ SETTINGS
// ================================================================
function saveBizSettings(data) {
  if (!data || !data.tenantId) return { success: false, error: 'tenantId required' };
  try {
    var existing = getRowById('BizSettings', data.tenantId);
    var fields = bizSettingsFields();
    if (existing) {
      updateRow('BizSettings', data.tenantId, data, fields);
    } else {
      // BizSettings uses tenantId as primary key
      var row = fields.map(function(f){ return data[f] !== undefined ? data[f] : ''; });
      appendToSheet('BizSettings', row);
    }
    return { success: true };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

function getBizSettings(tenantId) {
  return { success: true, settings: getBizSettingsRaw(tenantId) };
}

function getBizSettingsRaw(tenantId) {
  try {
    var sheet = getSheet('BizSettings');
    if (!sheet || sheet.getLastRow() <= 1) return null;
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0];
    var tidIdx = headers.indexOf('tenantId');
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][tidIdx] === tenantId) {
        var obj = {};
        headers.forEach(function(h, j){ obj[h] = rows[i][j]; });
        return obj;
      }
    }
  } catch(e) {}
  return null;
}

// ================================================================
// BULK SAVE
// ================================================================
function bulkSave(store, data) {
  if (!Array.isArray(data)) return { success: false, error: 'data must be array' };
  var errors = [];
  data.forEach(function(item) {
    try { genericSave(store, item); }
    catch(e) { errors.push(e.toString()); }
  });
  return { success: errors.length === 0, errors: errors };
}

// ================================================================
// DELETE
// ================================================================
function deleteRow(store, id) {
  if (!store || !id) return { success: false, error: 'store/id required' };
  var storeMap = {
    'tenants':'Tenants','users':'Users','products':'Products',
    'sales':'Sales','purchases':'Purchases',
    'customers':'Customers','suppliers':'Suppliers','activity':'Activity'
  };
  var sheetName = storeMap[store] || store;
  try {
    var sheet = getSheet(sheetName);
    var rows = sheet.getDataRange().getValues();
    var idIdx = rows[0].indexOf('id');
    for (var i = rows.length - 1; i >= 1; i--) {
      if (rows[i][idIdx] === id) {
        sheet.deleteRow(i + 1);
        // Delete related items
        if (store === 'sales') deleteItemsBySaleId('SalesItems', id);
        if (store === 'purchases') deleteItemsByPurchId('PurchaseItems', id);
        return { success: true };
      }
    }
    return { success: true };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ================================================================
// EMAIL
// ================================================================
function sendAlert(body) {
  try {
    if (!body.email) return { success: false, error: 'email required' };
    var subject = body.subject || 'StokiPro Alert';
    var msg = 'StokiPro Notification\n\n';
    if (body.type === 'low_stock' && body.data) {
      msg += 'Bidhaa zinazokwisha:\n';
      body.data.forEach(function(p){ msg += '- ' + p.name + ': ' + p.stock + ' ' + (p.unit||'') + '\n'; });
    } else {
      msg += JSON.stringify(body.data || {});
    }
    MailApp.sendEmail(body.email, subject, msg);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function forgotPassword(email) {
  try {
    if (!email) return { success: false, error: 'email required' };
    var sheet = getSheet('Users');
    if (!sheet) return { success: false, error: 'Users sheet not found' };
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0];
    var emailIdx = headers.indexOf('email');
    var pwIdx = headers.indexOf('password');
    var nameIdx = headers.indexOf('fullname');
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][emailIdx] === email) {
        var pw = rows[i][pwIdx];
        var name = rows[i][nameIdx] || 'Mtumiaji';
        MailApp.sendEmail(email, 'StokiPro — Nywila Yako',
          'Habari ' + name + ',\n\nNywila yako ya StokiPro: ' + pw + '\n\nAsante.');
        return { success: true };
      }
    }
    return { success: false, error: 'Barua pepe haikupatikana' };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ================================================================
// SHEET HELPERS
// ================================================================
function getSheet(name) {
  var s = SS.getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" not found. Run setupSheets() first.');
  return s;
}

function sheetToObjects(sheetName) {
  try {
    var sheet = SS.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Skip empty rows
      var obj = {};
      headers.forEach(function(h, j){ obj[h] = data[i][j] !== undefined ? data[i][j] : ''; });
      result.push(obj);
    }
    return result;
  } catch(e) { return []; }
}

function filterByTenant(sheetName, tenantId) {
  return sheetToObjects(sheetName).filter(function(r){ return r.tenantId === tenantId; });
}

function getRowById(sheetName, id) {
  try {
    var sheet = SS.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return null;
    var rows = sheet.getDataRange().getValues();
    var idIdx = rows[0].indexOf('id');
    if (idIdx < 0) idIdx = rows[0].indexOf('tenantId'); // BizSettings uses tenantId as key
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][idIdx] === id) return rows[i];
    }
    return null;
  } catch(e) { return null; }
}

function appendRow(sheetName, data, fields) {
  var row = fields.map(function(f){ return data[f] !== undefined ? data[f] : ''; });
  appendToSheet(sheetName, row);
}

function appendToSheet(sheetName, row) {
  getSheet(sheetName).appendRow(row);
}

function updateRow(sheetName, id, data, fields) {
  var sheet = getSheet(sheetName);
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var idIdx = headers.indexOf('id');
  if (idIdx < 0) idIdx = headers.indexOf('tenantId');
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][idIdx] === id) {
      fields.forEach(function(f) {
        var colIdx = headers.indexOf(f);
        if (colIdx >= 0 && data[f] !== undefined) {
          sheet.getRange(i+1, colIdx+1).setValue(data[f]);
        }
      });
      return;
    }
  }
  // Not found — append
  appendRow(sheetName, data, fields);
}

function deleteItemsBySaleId(sheetName, saleId) {
  try {
    var sheet = SS.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    var rows = sheet.getDataRange().getValues();
    var saleIdIdx = rows[0].indexOf('saleId');
    for (var i = rows.length - 1; i >= 1; i--) {
      if (rows[i][saleIdIdx] === saleId) sheet.deleteRow(i+1);
    }
  } catch(e) {}
}

function deleteItemsByPurchId(sheetName, purchId) {
  try {
    var sheet = SS.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    var rows = sheet.getDataRange().getValues();
    var purchIdIdx = rows[0].indexOf('purchId');
    for (var i = rows.length - 1; i >= 1; i--) {
      if (rows[i][purchIdIdx] === purchId) sheet.deleteRow(i+1);
    }
  } catch(e) {}
}

function saveRow(sheetName, data, fields) {
  if (!data || !data.id) return { success: false, error: 'id required' };
  try {
    var existing = getRowById(sheetName, data.id);
    if (existing) { updateRow(sheetName, data.id, data, fields); }
    else { appendRow(sheetName, data, fields); }
    return { success: true };
  } catch(err) { return { success: false, error: err.toString() }; }
}

function getFields(sheetName) {
  try {
    var sheet = SS.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 1) return [];
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  } catch(e) { return []; }
}

// ================================================================
// FIELD DEFINITIONS
// ================================================================
function tenantFields()    { return ['id','name','owner','phone','email','type','status','created']; }
function userFields()      { return ['id','username','password','fullname','phone','email','role','tenantId','status']; }
function productFields()   { return ['id','tenantId','name','cat','unit','bp','sp','stock','minStock','desc','barcode']; }
function saleFields()      { return ['id','tenantId','ref','date','custId','custName','subtotal','discount','total','paid','change','payment','creditType','status','created']; }
function purchaseFields()  { return ['id','tenantId','ref','date','supId','supName','total','paid','payment','creditType','created']; }
function customerFields()  { return ['id','tenantId','name','phone','email','addr','group','creditLimit','currentDebt','salesCount','totalSpent','lastSale','notes','created']; }
function supplierFields()  { return ['id','tenantId','name','phone','email','prods','payDays','totalPurch','currentDebt','notes','created']; }
function customerPaymentFields() { return ['id','custId','tenantId','date','amount','method','debtBefore','debtAfter','note','created']; }
function supplierPaymentFields() { return ['id','supId','tenantId','date','amount','method','debtBefore','debtAfter','note','created']; }
function bizSettingsFields(){ return ['tenantId','bizName','bizPhone','bizEmail','bizAddress','tinNumber','currency','dailyReport','alertLow','alertDebt']; }
function activityFields()  { return ['id','uid2','uN','biz','action','type','created']; }

// ================================================================
// SETUP SHEETS — Run once to create all tables
// ================================================================
function setupSheets() {
  var sheets = {
    'Tenants':          tenantFields(),
    'Users':            userFields(),
    'Products':         productFields(),
    'Sales':            saleFields(),
    'SalesItems':       ['saleId','tenantId','pid','pname','qty','price','total'],
    'Purchases':        purchaseFields(),
    'PurchaseItems':    ['purchId','tenantId','pid','pname','qty','price','total'],
    'Customers':        customerFields(),
    'Suppliers':        supplierFields(),
    'CustomerPayments': customerPaymentFields(),
    'SupplierPayments': supplierPaymentFields(),
    'BizSettings':      bizSettingsFields(),
    'Activity':         activityFields(),
  };

  Object.keys(sheets).forEach(function(name) {
    var sheet = SS.getSheetByName(name);
    if (!sheet) {
      sheet = SS.insertSheet(name);
      Logger.log('Created: ' + name);
    }
    // Set headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(sheets[name]);
      sheet.getRange(1, 1, 1, sheets[name].length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      Logger.log('Headers set for: ' + name);
    }
  });

  Logger.log('✅ setupSheets() complete! All tables ready.');
  return 'Done — ' + Object.keys(sheets).length + ' tables created/verified';
}

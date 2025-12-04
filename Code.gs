/**************** CONFIG ****************/
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SH_PRODUCTS = 'MASTER_PRODUCT';
const SH_MACHINES = 'MASTER_MACHINE';
const SH_PROD_LOG = 'PRODUCTION_LOG';

/**************** UI ENTRY POINT ****************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Production Management System');
}

/**************** UTILS ****************/
function getSheet_(name) {
  return SS.getSheetByName(name);
}

function getDataAsObjects_(sheetName) {
  const sh = getSheet_(sheetName);
  const range = sh.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1);
  const out = [];
  rows.forEach(row => {
    if (row.join('') === '') return; // skip fully empty
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });
    out.push(obj);
  });
  return out;
}

function saveObject_(sheetName, obj) {
  const sh = getSheet_(sheetName);
  const range = sh.getDataRange();
  const values = range.getValues();
  const headers = values[0];

  const idIndex = headers.indexOf('ID');
  if (idIndex === -1) {
    throw new Error('No ID column found in ' + sheetName);
  }

  const id = obj.ID;
  let rowIndex = -1; // 1-based index for sheet

  // Update existing
  if (id && id !== '') {
    for (let i = 1; i < values.length; i++) {
      if (values[i][idIndex] === id) {
        rowIndex = i + 1;
        break;
      }
    }
  }

  // New record
  if (rowIndex === -1) {
    let maxId = 0;
    for (let i = 1; i < values.length; i++) {
      const val = values[i][idIndex];
      if (typeof val === 'number' && val > maxId) maxId = val;
    }
    const newId = maxId + 1;
    obj.ID = newId;

    const newRow = headers.map(h => obj[h] !== undefined ? obj[h] : '');
    sh.appendRow(newRow);
    return obj;
  }

  // Update existing row
  const rowValues = sh.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const updatedRow = rowValues.map((cell, i) => {
    const key = headers[i];
    return obj[key] !== undefined ? obj[key] : cell;
  });

  sh.getRange(rowIndex, 1, 1, headers.length).setValues([updatedRow]);
  return obj;
}

function deleteById_(sheetName, id) {
  const sh = getSheet_(sheetName);
  const range = sh.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const idIndex = headers.indexOf('ID');
  if (idIndex === -1) throw new Error('No ID column found in ' + sheetName);

  for (let i = 1; i < values.length; i++) {
    if (values[i][idIndex] === id) {
      sh.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

/**************** PRODUCTS API ****************/
function getProducts() {
  return getDataAsObjects_(SH_PRODUCTS);
}

function saveProduct(product) {
  // Normalize types
  product.ID = product.ID ? Number(product.ID) : '';
  if (product.StdCycleTime) product.StdCycleTime = Number(product.StdCycleTime);
  if (product.StdScrapPct) product.StdScrapPct = Number(product.StdScrapPct);
  if (product.SellingPrice) product.SellingPrice = Number(product.SellingPrice);
  return saveObject_(SH_PRODUCTS, product);
}

function deleteProduct(id) {
  return deleteById_(SH_PRODUCTS, Number(id));
}

/**************** MACHINES API ****************/
function getMachines() {
  return getDataAsObjects_(SH_MACHINES);
}

function saveMachine(machine) {
  machine.ID = machine.ID ? Number(machine.ID) : '';
  if (machine.Capacity) machine.Capacity = Number(machine.Capacity);
  return saveObject_(SH_MACHINES, machine);
}

function deleteMachine(id) {
  return deleteById_(SH_MACHINES, Number(id));
}

/**************** PRODUCTION LOG API ****************/
function getProductionLogs() {
  return getDataAsObjects_(SH_PROD_LOG);
}

function saveProductionLog(log) {
  log.ID = log.ID ? Number(log.ID) : '';
  if (log.PlannedQty) log.PlannedQty = Number(log.PlannedQty);
  if (log.ActualQty) log.ActualQty = Number(log.ActualQty);
  if (log.ScrapQty) log.ScrapQty = Number(log.ScrapQty);
  return saveObject_(SH_PROD_LOG, log);
}

function deleteProductionLog(id) {
  return deleteById_(SH_PROD_LOG, Number(id));
}

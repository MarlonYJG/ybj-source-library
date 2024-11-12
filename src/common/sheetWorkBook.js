/*
 * @Author: Marlon
 * @Date: 2024-06-12 10:31:55
 * @Description:
 */
import * as GC from '@grapecity/spread-sheets';

import { GeneratorTableStyle } from './generator';
/**
 * SetDataSource
 * @param {*} sheet
 * @param {*} dataSource
 */
export const SetDataSource = (sheet, dataSource) => {
  const data = new GC.Spread.Sheets.Bindings.CellBindingSource(dataSource);
  sheet.setDataSource(data);
};

/**
 * Lock the spreadsheet
 * @param {*} spread 
 * @param {*} domId 
 * @param {*} locked 
 */
export const SpreadLocked = (spread, domId = 'spreadsheet-quotation-view', locked = true) => {
  const workbook = GC.Spread.Sheets.findControl(document.getElementById(domId));
  const sheetCount = workbook.getSheetCount();
  spread.suspendPaint();
  for (let index = 0; index < sheetCount; index++) {
    const sheet = spread.getSheet(index);
    if (sheet) {
      sheet.options.isProtected = locked;
      sheet.getRange(0, 0, sheet.getRowCount() - 1, sheet.getColumnCount()).locked(locked);
    }
  }
  spread.resumePaint();
};

/**
 * Create table
 * @param {*} sheet
 * @param {*} id
 * @param {*} r
 * @param {*} c
 * @param {*} rc
 * @param {*} cc
 * @param {*} header
 * @param {*} bindDataPath
 */
export const CreateTable = (sheet, id, r, c, rc, cc, header, bindDataPath) => {
  console.log(sheet.name());

  console.log(id, r, c, rc, cc);
  
  console.log(sheet.tables);
  const table = sheet.tables.add(`table${id}`, r, c, rc, cc, GeneratorTableStyle());
  const tableColumns = [];
  for (const key in header) {
    if (Object.hasOwnProperty.call(header, key)) {
      const tableColumn = new GC.Spread.Sheets.Tables.TableColumn();
      tableColumn.name(header[key].name || key);
      tableColumn.dataField(header[key].bindPath);
      tableColumns.push(tableColumn);
    }
  }
  table.bindColumns(tableColumns);
  table.expandBoundRows(true);
  table.autoGenerateColumns(false);
  table.showFooter(false);
  table.showHeader(false);
  table.highlightFirstColumn(false);
  table.highlightLastColumn(false);

  if (bindDataPath) {
    table.bindingPath(bindDataPath);
  }
};

/**
 * Create a sheet object
 * @param {*} rootWorkBook 
 * @param {*} sheetName 
 * @param {*} index 
 * @param {*} sheetTemplateSource 
 */
export const CreateSheet = (rootWorkBook, sheetName, index = 1, sheetTemplateSource = null) => {
  const sheet = new GC.Spread.Sheets.Worksheet(sheetName);
  rootWorkBook.addSheet(index, sheet);
  if (sheetTemplateSource) {
    // sheet.fromJSON(sheetTemplateSource);
  }
};

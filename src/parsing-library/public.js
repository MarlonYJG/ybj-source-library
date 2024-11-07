/*
 * @Author: Marlon
 * @Date: 2024-04-10 20:46:55
 * @Description:General Purpose Processing
 */
import * as GC from '@grapecity/spread-sheets';
import * as ExcelIO from '@grapecity/spread-excelio';
import _ from '../lib/lodash/lodash.min.js';

import { exportErrorProxy } from '../common/proxyData'
// eslint-disable-next-line no-unused-vars
import { columnToNumber, PubGetRandomNumber } from '../common/public';
// eslint-disable-next-line no-unused-vars
import { setCellFormatter } from '../common/single-table';

/**
 * Get a view of resources
 * @param {*} spread
 * @param {*} quotation
 */
export const PubGetResourceViews = (spread, quotation) => {
  const sheet = spread.getActiveSheet();
  if (sheet.tag() == 'sheet') {
    return quotation.conferenceHall.resourceViews;
  } else {
    return quotation.parallelSessions[spread.getActiveSheetIndex() - 1].resourceViews;
  }
};

/**
 * Obtain the resource map
 * @param {*} spread
 * @param {*} quotation
 * @returns
 */
export const PubGetResourceViewsMap = (spread, quotation) => {
  const sheet = spread.getActiveSheet();
  let resourceViewsMap = {};
  if (sheet.tag() == 'sheet') {
    resourceViewsMap = quotation.conferenceHall.resourceViewsMap;
  } else {
    resourceViewsMap = quotation.parallelSessionsMap[sheet.tag()].resourceViewsMap;
  }
  return resourceViewsMap;
};

/**
 * Download feature
 * @param {*} fileType
 * @param {*} blob
 * @param {*} fileName
 */
const saveAs = (fileType = 'xlsx', blob, fileName) => {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName + `.${fileType}`;
  a.click();
  window.URL.revokeObjectURL(url);
};

/**
 * spread导出Excel
 * @param {*} spread
 * @param {*} fileName
 */
export const spreadExportExcel = (spread, fileName = '报价单') => {
  const option = {
    includeBindingSource: true, includeStyles: true, includeFormulas: true, saveAsView: false, rowHeadersAsFrozenColumns: false, columnHeadersAsFrozenRows: false, includeAutoMergedCells: false, includeCalcModelCache: false, includeUnusedNames: true, includeEmptyRegionCells: true, fileType: 0
  };
  const json = spread.toJSON({ includeBindingSource: true });
  const excelIo = new ExcelIO.IO();
  excelIo.save(json, (blob) => {
    saveAs('xlsx', blob, fileName);
  }, (err) => {
    console.log(err);
    exportErrorProxy.value = err;
  }, option);
};

/**
 * spread导出PDF
 * @param {*} spread
 * @param {*} fileName
 * @param {*} author
 */
export const spreadExportPDF = (spread, fileName = '报价单', author = 'yunbaojia') => {
  const option = {
    title: fileName,
    author: author,
    subject: fileName,
    keywords: 'yunbaojia',
    creator: 'yunbaojia'
  };
  spread.savePDF((blob) => {
    saveAs('pdf', blob, fileName);
  }, (error) => { console.log(error); }, option);
};

/**
 * Whether the cell is editable or not
 * @param {*} templateSource
 * @param {*} locked
 * @returns
 */
export const spreadStyleLocked = (templateSource, locked = false) => {
  const template = _.cloneDeep(templateSource);
  const sheets = template.sheets;

  for (const i in sheets) {
    if (Object.hasOwnProperty.call(sheets, i)) {
      const dataTable = sheets[i].data.dataTable;
      for (const j in dataTable) {
        if (Object.hasOwnProperty.call(dataTable, j)) {
          const row = dataTable[j];
          for (const k in row) {
            if (Object.hasOwnProperty.call(row, k)) {
              const column = row[k];
              if (column.style) {
                if (typeof (column.style) !== 'string') {
                  column.style.locked = locked;
                }
              } else {
                column.style = {
                  locked: locked
                };
              }
            }
          }
        }
      }
    }
  }

  return template;
};

/**
 * 排序
 * @param {*} spread
 */
export const spreadSort = (spread) => {
  const sheet = spread.getActiveSheet();
  console.log(sheet, '排序');
};

/**
 * Initialize the table
 * @param {*} spread
 * @param {*} tableId
 * @param {*} param2
 * @param {*} param3
 */
export const spreadInitTable = (spread, tableId, { insertRow, rowCount }, { startRow, columnCount }) => {
  console.log('初始化空分类 表格');
  const sheet = spread.getActiveSheet();
  if (!sheet.tables.findByName('table' + tableId)) {
    sheet.suspendPaint();
    sheet.addRows(insertRow, rowCount + 1);
    const tableStyle = new GC.Spread.Sheets.Tables.TableTheme();
    const tStyleInfo = spreadInitTableStyle();
    tableStyle.headerRowStyle(tStyleInfo);

    const table = sheet.tables.add('table' + tableId, startRow, 0, 2, columnCount, tableStyle);
    table.expandBoundRows(true);
    table.autoGenerateColumns(false);
    table.showFooter(false);
    table.showHeader(false);
    table.highlightFirstColumn(false);
    table.highlightLastColumn(false);
    sheet.resumePaint();
  }
};

/**
 * Initialize the table style
 */
export const spreadInitTableStyle = () => {
  const tStyleInfo = new GC.Spread.Sheets.Tables.TableStyle();
  tStyleInfo.backColor = 'white';
  return tStyleInfo;
};

/**
 * table 字段绑定
 * @param {*} spread
 * @param {*} tableId
 * @param {*} columnFields
 * @param {*} tablePath
 */
export const spreadTableBind = (spread, tableId, columnFields, tablePath) => {
  const sheet = spread.getActiveSheet();
  const table = sheet.tables.findByName('table' + tableId);
  if (table) {
    const tableColumns = [];
    for (const key in columnFields) {
      if (Object.hasOwnProperty.call(columnFields, key)) {
        const field = columnFields[key];
        const tableColumn = new GC.Spread.Sheets.Tables.TableColumn();
        tableColumn.name(field.name);
        tableColumn.dataField(field.bindPath);
        tableColumns.push(tableColumn);
      }
    }
    table.bindColumns(tableColumns);
    table.bindingPath(tablePath);
  }
};

/**
 * Calculate the formula field (subtotal) initialization
 * @param {*} sheet
 * @param {*} field
 * @param {*} fieldName
 * @param {*} row
 */
export const computedField = (sheet, field, fieldName, row) => {
  const computedFieldFormula = _.cloneDeep(field.bindPath[fieldName].formula);
  for (let i = 0; i < computedFieldFormula.size; i++) {
    computedFieldFormula.formula = computedFieldFormula.formula.replace('{{' + i + '}}', row + 1);
  }
  sheet.setFormula(row, computedFieldFormula.column, computedFieldFormula.formula);
};

/**
 * Total Initialization
 * @param {*} sheet
 * @param {*} totalField
 * @param {*} equipmentFormula
 * @param {*} totalStartRow
 * @param {*} equipmentStartRow
 */
export const computedTotalField = (sheet, totalField, equipmentFormula, totalStartRow, equipmentStartRow) => {
  const totalFormula = _.cloneDeep(totalField.bindPath.total.formula);
  for (let i = 0; i < equipmentFormula.size; i++) {
    totalFormula.formula = totalFormula.formula.replace('{{' + i + '}}', equipmentStartRow + 1);
  }
  sheet.setFormula(totalStartRow, totalFormula.column, totalFormula.formula);
};

/**
 * Insert an image into the table
 * @param {*} spread
 * @param {*} tableId
 * @param {*} image
 * @param {*} pictureName
 * @param {*} base64
 * @param {*} endRow
 * @param {*} startColumn
 * @param {*} endColumn
 */
export const addImageOfTable = (spread, tableId, image, pictureName, base64, endRow, startColumn, endColumn) => {
  const sheet = spread.getActiveSheet();
  const table = sheet.tables.findByName('table' + tableId);
  if (table) {
    const startRow = table.startRow();
    let imgX = 2;
    let imgY = 2;
    let curWidth = 100;
    let curHeight = 100;
    if (image) {
      imgX = image.imgX || 2;
      imgY = image.imgY || 2;
      curWidth = image.width || 100;
      curHeight = image.height || 100;
    }
    sheet.suspendPaint();
    const picture = sheet.pictures.add(pictureName, base64, imgX, imgY, 2, 2);
    picture.startRow(startRow);
    picture.endRow(startRow + endRow);
    picture.startColumn(startColumn);
    picture.endColumn(endColumn);
    picture.width(curWidth);
    picture.height(curHeight);
    picture.allowMove(true);
    picture.allowResize(true);
    picture.isLocked(false);

    sheet.resumePaint();
  }
};

export const resourcesDataFormat = (resources) => {
  resources.forEach(item => {
    if (item.discountUnitPrice === 0 || item.discountUnitPrice) {
      item.discountUnitPrice = parseFloat(item.discountUnitPrice);
    } else {
      item.discountUnitPrice = parseFloat(item.unitPrice);
    }

    if (item.quantity) {
      item.quantity = parseFloat(item.quantity);
    } else {
      item.quantity = 0;
    }
    if (!item.imageId) {
      item.imageId = PubGetRandomNumber();
    }
    if (item.numberOfDays) {
      item.numberOfDays = Number(item.numberOfDays);
    } else {
      item.numberOfDays = 1;
    }
  });
};

/**
 * set lastColumn Width
 * @param {*} spread
 * @param {*} template
 */
export const setLastColumnWidth = (spread, template, val = 10) => {
  const sheet = spread.getActiveSheet();
  sheet.suspendPaint();
  const column = template.cloudSheet.center.columnCount - 1;
  sheet.setColumnWidth(column, sheet.getColumnWidth(column) + val);
  sheet.resumePaint();
};

/**
 * Initialization configuration of the workbook
 * @param {*} workBook 
 */
export const initWorkBookConfig = (workBook) => {
  if (workBook) {
    workBook.options.tabEditable = false;
  }
};

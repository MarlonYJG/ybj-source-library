/*
 * @Author: Marlon
 * @Date: 2024-06-19 13:29:25
 * @Description:
 */
import * as GC from '@grapecity/spread-sheets';
import * as ExcelIO from '@grapecity/spread-excelio';
import _ from 'lodash';
import { Message } from 'element-ui';
import store from 'store';
import { SHOW_DELETE, UPDATE_QUOTATION_PATH, IGNORE_EVENT, FROZEN_HEAD_TEMPLATE } from 'store/quotation/mutation-types';

import { SetDataSource } from '../common/sheetWorkBook';

import { Reset } from './public';
import { MENU_TOTAL } from './config';

import { Render, insertField, removeAllTable, UpdateTotalBlock } from './single-table';

import { getTemplateTopRowCol } from '../common/parsing-template';

import { CheckCostPrice } from '../common/cost-price';

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
 * Clear product
 * @param {*} spread
 */
const clearProduct = (spread) => {
  store.commit(`quotationModule/${IGNORE_EVENT}`, true);
  const sheet = spread.getActiveSheet();
  const quotation = store.getters['quotationModule/GetterQuotationInit'];

  removeAllTable(sheet, quotation);

  const conferenceHall = _.cloneDeep(quotation.conferenceHall);
  conferenceHall.resourceViews = [];
  conferenceHall.resourceViewsMap = {};

  store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
    path: ['conferenceHall'],
    value: conferenceHall
  });
  SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);

  UpdateTotalBlock(sheet);

  store.commit(`quotationModule/${IGNORE_EVENT}`, false);
};

/**
 * Freeze header in spread
 * @param {*} spread 
 * @param {*} turnOn 
 */
const spreadFrozenHead = (spread, turnOn) => {
  const sheet = spread.getActiveSheet();
  if (turnOn) {
    const { rowCount, columnCount } = getTemplateTopRowCol()
    sheet.frozenRowCount(rowCount);
    sheet.frozenColumnCount(columnCount);
    sheet.options.frozenlineColor = "transparent";
  } else {
    sheet.frozenRowCount(0);
    sheet.frozenColumnCount(0);
  }
}

/**
 * spread type export Excel
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
    Message.error({
      message: '导出失败!'
    });
  }, option);
};

/**
 * spread type export PDF
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
 * Print
 * @param {*} spread
 * @param {*} sheetIndex
 */
export const spreadPrint = (spread, domId = 'spreadsheet-quotation') => {
  const workbook = GC.Spread.Sheets.findControl(document.getElementById(domId));
  const sheetCount = workbook.getSheetCount();
  for (let index = 0; index < sheetCount; index++) {
    // const sheet = spread.getSheet(index);
    // const printInfo = sheet.printInfo();
    // printInfo.pageHeaderFooter({
    //   normal: {
    //     header: {
    //       left: '&G',
    //       center: '&"Comic Sans MS"System Information',
    //       leftImage: 'images/GrapeCity_LOGO.jpg'
    //     },
    //     footer: {
    //       center: '&P/&N'
    //     }
    //   }
    // });
  }
  spread.print();
};

/**
 * zoom
 * @param {*} type
 */
export const zoom = (spread, type) => {
  const sheet = spread.getActiveSheet();
  const num = sheet.zoom();
  if (type) {
    if (num + 0.25 < 4) {
      sheet.zoom(num + 0.25);
    } else {
      sheet.zoom(4);
    }
  } else {
    if (num - 0.25 > 0) {
      sheet.zoom(num - 0.25);
    } else {
      sheet.zoom(0.25);
    }
  }
  return sheet.zoom();
};

/**
 * Create a modal for the TotalModel
 * @param {*} quotation
 * @param {*} template
 * @param {*} field
 * @param {*} cb
 */
export const FormComputedRowField = (field, cb) => {
  const data = {};
  switch (field.value) {
    case 'tax':
      data[field.value] = {
        label: field.label,
        prop: field.value,
        value: null
      };
      break;
    case 'rate':
      data.rate = {
        label: '服务费率',
        prop: 'rate',
        value: null
      };
      break;
    case 'concessional':
      data[field.value] = {
        label: field.label,
        prop: field.value,
        value: null
      };
      break;
    case 'freight':
      data[field.value] = {
        label: field.label,
        prop: field.value,
        value: null
      };
      break;
    case 'managementFee':
      data[field.value] = {
        label: '管理费率',
        prop: 'managementFee',
        value: null
      };
      break;
    case 'projectCost':
      data[field.value] = {
        label: field.label,
        prop: field.value,
        value: null
      };
      break;
    default:
      break;
  }

  cb(data);
};

/**
 * product sort
 * @param {*} spread
 * @param {*} resourceViews
 */
export const Sort = (spread, resourceViews) => {
  const sheet = spread.getActiveSheet();
  const resourceViewsMap = {};
  resourceViews.forEach(item => {
    resourceViewsMap[item.resourceLibraryId] = item;
  });

  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const conferenceHall = _.cloneDeep(quotation.conferenceHall);
  conferenceHall.resourceViews = resourceViews;
  conferenceHall.resourceViewsMap = resourceViewsMap;

  store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
    path: ['conferenceHall'],
    value: conferenceHall
  });

  Reset(spread);
  SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);
  Render(spread);
};

/**
 * Delete the product
 * @param {*} spread
 * @param {*} resourceViews
 */
export const DeleteProduct = (spread, resourceViews) => {
  const sheet = spread.getActiveSheet();
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const conferenceHall = _.cloneDeep(quotation.conferenceHall);
  conferenceHall.resourceViews = resourceViews;
  const resourceViewsMap = {};
  resourceViews.forEach(el1 => {
    resourceViewsMap[el1.resourceLibraryId] = el1;
  });
  conferenceHall.resourceViewsMap = resourceViewsMap;
  Reset(spread);

  store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
    path: ['conferenceHall'],
    value: conferenceHall
  });

  SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);
  Render(spread);
};

/**
 * HeadDelete
 * @param {*} spread
 * @param {*} fieldType
 * @returns
 */
export const HeadDelete = (spread, fieldType) => {
  if (!spread) {
    console.error('spreadJS初始化失败!');
    return;
  }
  const fiedls = MENU_TOTAL.map((item) => { return item.value; });
  if (fiedls.includes(fieldType.value)) {
    insertField(spread, fieldType.value, null);
  } else if (fieldType.value === 'clearProduct') {
    clearProduct(spread);
  } else if (fieldType.value === 'delProduct') {
    store.commit(`quotationModule/${SHOW_DELETE}`, true);
  }
};

/**
 * Redraw the product area and total area
 * @param {*} spread
 */
export const Repaint = (spread) => {
  const sheet = spread.getActiveSheet();
  Reset(spread);
  SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);
  Render(spread);
};

/**
 * Top action: Freeze the header
 * @param {*} spread 
 */
export const FrozenHead = (spread) => {
  const frozen = store.getters['quotationModule/GetterTemplateHFrozen'];
  store.commit(`quotationModule/${FROZEN_HEAD_TEMPLATE}`, !frozen);
  spreadFrozenHead(spread, !frozen)
}

/**
 * Show the cost price
 */
export const ShowCostPrice = (spread, locked = false) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const showCost = store.getters['quotationModule/GetterShowCostPrice'];
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  console.log(showCost, '是否显示成本价');
  console.log(template, 'template');
  if (spread && template) {
    const costPrice = new CheckCostPrice(spread, template, quotation);
    if (showCost) {
      costPrice.deleteCol()
      costPrice.render(locked)
    } else {
      costPrice.deleteCol()
    }
  }
}
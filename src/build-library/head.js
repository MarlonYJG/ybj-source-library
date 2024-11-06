/*
 * @Author: Marlon
 * @Date: 2024-06-19 13:29:25
 * @Description:
 */
import Decimal from '../lib/decimal/decimal.min.js';
import * as GC from '@grapecity/spread-sheets';
import * as ExcelIO from '@grapecity/spread-excelio';
import _ from '../lib/lodash/lodash.min.js';
import store from 'store';
import { SHOW_DELETE, UPDATE_QUOTATION_PATH, IGNORE_EVENT, FROZEN_HEAD_TEMPLATE } from 'store/quotation/mutation-types';

import { PRICE_SET_MAP } from '../common/constant'
import { SetDataSource } from '../common/sheetWorkBook';
import { exportErrorProxy } from '../common/proxyData';
import { getWorkBook } from '../common/store';

import {
  getTemplateTopRowCol, getDiscountField, showPriceSet,
  getImageConfig,
  getEquipmentConfig
} from '../common/parsing-template';
import { getConfig } from '../common/parsing-quotation';

import {
  getTableRowIndex,
  setAutoFitRow,
  defaultAutoFitRow,
} from '../common/single-table';

import { CheckCostPrice } from '../common/cost-price';

import { Reset } from './public';
import { MENU_TOTAL } from './config';

// eslint-disable-next-line no-unused-vars
import { Render, insertField, removeAllTable, UpdateTotalBlock, resetDiscountRatio, setRowImageHeight } from './single-table';


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
    exportErrorProxy.value = err;
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
 * @param {*} domId 
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
 * @param {*} spread 
 * @param {*} type 
 * @returns 
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
 * @param {*} spread 
 * @param {*} locked 
 */
export const ShowCostPrice = (spread, locked = false) => {
  const template = getWorkBook();
  const showCost = store.getters['quotationModule/GetterShowCostPrice'];
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
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

/**
 * Cost price status
 * @returns 
 */
export const ShowCostPriceStatus = () => {
  const costLocked = store.getters['GetterCostPrice'];
  return costLocked === 1 ? false : true;
}

/**
 * Update the discount value
 * @param {*} spread 
 * @param {*} percentage 
 */
export const UpdateDiscount = (spread, percentage) => {
  const discountField = getDiscountField();
  if (discountField) {
    const priceAdjustment = new Decimal(percentage).dividedBy(new Decimal(100)).toNumber();
    store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
      path: ['priceAdjustment'],
      value: priceAdjustment
    });

    const sheet = spread.getActiveSheet();
    const template = getWorkBook();
    console.log(template, 'template');

    const quotation = store.getters['quotationModule/GetterQuotationInit'];
    const conferenceHall = _.cloneDeep(quotation.conferenceHall);
    const resourceViews = _.cloneDeep(conferenceHall.resourceViews);
    const resourceViewsMap = {};
    // const PriceStatus = store.getters['quotationModule/GetterQuotationPriceStatus'];
    const PriceStatus = 0;
    resourceViews.forEach((item) => {
      if (item.resources.length) {
        item.resources.forEach((resource) => {
          if (Object.keys(resource).includes(discountField)) {
            const sourcePrice = resource[PRICE_SET_MAP[PriceStatus]];
            resource[discountField] = new Decimal(priceAdjustment).times(new Decimal(sourcePrice)).toNumber();
          } else {
            console.warn(`resource not include ${discountField}`);
          }
        });
      }
    });
    resourceViews.forEach(item => {
      resourceViewsMap[item.resourceLibraryId] = item;
    });
    conferenceHall.resourceViews = resourceViews;
    conferenceHall.resourceViewsMap = resourceViewsMap;

    store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
      path: ['conferenceHall'],
      value: conferenceHall
    });
    SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);

    ShowCostPrice(spread, ShowCostPriceStatus());
  }
}

/**
 * Update top price settings
 * @param {*} spread 
 * @param {*} priceSet 
 */
export const UpdatePriceSet = (spread, priceSet) => {
  if (showPriceSet()) {
    store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
      path: ['priceStatus'],
      value: priceSet
    });

    // resetDiscountRatio()

    const quotation = store.getters['quotationModule/GetterQuotationInit'];
    const conferenceHall = _.cloneDeep(quotation.conferenceHall);
    const resourceViews = _.cloneDeep(conferenceHall.resourceViews);
    const resourceViewsMap = {};

    resourceViews.forEach((item) => {
      if (item.resources.length) {
        item.resources.forEach((resource) => {
          if (Object.keys(resource).includes('discountUnitPrice')) {
            let sourcePrice = resource.unitPrice;
            if (priceSet === 0 || priceSet) {
              if (priceSet === 0) {
                if (_.has(resource, 'unitPrice')) {
                  sourcePrice = resource.unitPrice;
                } else {
                  console.error('resource not include unitPrice');
                }
              } else {
                if (_.has(resource, `unitPrice${priceSet}`)) {
                  sourcePrice = resource[`unitPrice${priceSet}`];
                } else {
                  console.error(`resource not include unitPrice${priceSet}`);
                }
              }
            }
            resource.discountUnitPrice = sourcePrice;
          } else {
            console.warn('resource not include discountUnitPrice');
          }
        });
      }
    });
    resourceViews.forEach(item => {
      resourceViewsMap[item.resourceLibraryId] = item;
    });
    conferenceHall.resourceViews = resourceViews;
    conferenceHall.resourceViewsMap = resourceViewsMap;

    store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
      path: ['conferenceHall'],
      value: conferenceHall
    });
    const sheet = spread.getActiveSheet();
    SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);

    ShowCostPrice(spread, ShowCostPriceStatus());
  }
}

/**
 * Start adaptive line height
 * @param {*} spread 
 */
export const StartAutoFitRow = (spread) => {
  const sheet = spread.getActiveSheet();

  const image = getImageConfig();
  const rowsField = getEquipmentConfig();
  const tableRows = getTableRowIndex(spread);
  const config = getConfig()

  console.log(config);

  console.log(image);


  sheet.suspendPaint();
  if (config && config.startAutoFitRow) {
    tableRows.forEach((row, i) => {
      ((row) => {
        sheet.autoFitRow(row);
        setAutoFitRow(sheet, row, rowsField, image, (height) => {
          setRowImageHeight(sheet, row, height, i);
        });
      })(row)
    });
  } else {
    tableRows.forEach(row => {
      ((row) => {
        defaultAutoFitRow(sheet, row, rowsField, image);
      })(row);
    });
  }
  sheet.resumePaint();
}



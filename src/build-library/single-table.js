/*
 * @Author: Marlon
 * @Date: 2024-05-16 14:20:35
 * @Description:single-build
 */
import _ from 'lodash';
import * as GC from '@grapecity/spread-sheets';
import store from 'store';
import { isNumber } from '../utils/index';
import { FormatDate } from '../utils/index';
import { regChineseCharacter } from '../utils/index';

import {
  UPDATE_QUOTATION_PATH, DELETE_QUOTATION_PATH, UPDATE_HEAD_TOTAL,
  SET_WORK_SHEET_QUOTATION, SET_SHEET_LEAVEL_POSITION, IGNORE_EVENT
} from 'store/quotation/mutation-types';

import { NEW_OLD_FIELD_MAP } from './config';

import { commandRegister, onOpenMenu } from './contextMenu';
import { ShowCostPrice } from './head';

import { CreateTable } from '../common/sheetWorkBook';
import IdentifierTemplate from '../common/identifier-template';

import { CombinationTypeBuild } from '../common/combination-type';
import { ASSOCIATED_FIELDS_FORMULA_MAP, DESCRIPTION_MAP, TOTAL_COMBINED_MAP } from '../common/constant';
import { GeneratorCellStyle, GeneratorLineBorder } from '../common/generator';
import { numberToColumn } from '../common/public'

import { getPositionBlock } from '../common/parsing-quotation';
import {
  PubGetTableStartColumnIndex,
  PubGetTableColumnCount,
  delTableHeaderRowCount,
  templateRenderFlag,
  getTableHeaderDataTable,
  mergeRow,
  PubSetCellHeight,
  setRowStyle,
  mergeColumn,
  showTotal,
  showSubTotal,
  getComputedColumnFormula,
  getFormulaFieldRowCol
} from '../common/parsing-template';

import {
  templateTotalMap, mergeSpan, setCellStyle, setTotalRowHeight, PubGetTableStartRowIndex,
  rowComputedFieldSort,
  GenerateFieldsRow,
  classificationAlgorithms,
  columnsTotal,
  columnTotalSumFormula,
  mixedDescriptionFields,
  PubGetTableRowCount,
  tableHeader,
  columnComputedValue,
  // setCellFormatter,
  SetComputedSubTotal,
  renderSheetImage,
  initShowCostPrice
} from '../common/single-table';

import { LayoutRowColBlock } from '../common/core';
import { CheckCostPrice } from '../common/cost-price';

import { UpdateSort } from './public';

const NzhCN = require('nzh/cn');

let RangeChangedTimer = null;

/**
 * Event bind
 * @param {*} spread
 */
const OnEventBind = (spread) => {
  const sheet = spread.getActiveSheet();
  // sheet.bind(GC.Spread.Sheets.Events.EditChange, (sender, args) => {
  //   console.log(0);
  // });
  // sheet.bind(GC.Spread.Sheets.Events.LeaveCell, (sender, args) => {
  //   console.log(1);
  // });
  // sheet.bind(GC.Spread.Sheets.Events.EditEnded, (sender, args) => {
  //   console.log(2);
  // });
  // sheet.bind(GC.Spread.Sheets.Events.EditEnding, (sender, args) => {
  //   console.log(3);
  // });

  sheet.bind(GC.Spread.Sheets.Events.RangeChanged, (sender, args) => {
    const ignoreEvent = store.getters['quotationModule/GetterIgnoreEvent'];

    if (!ignoreEvent) {
      RangeChangedTimer && clearTimeout(RangeChangedTimer);
      RangeChangedTimer = setTimeout(() => {
        const dataSource = sheet.getDataSource().getSource();
        console.log('RangeChanged 事件');
        if ([2, 0, 1, 3, 4].includes(args.action)) {
          store.commit(`quotationModule/${SET_WORK_SHEET_QUOTATION}`, dataSource);
          UpdateTotalBlock(sheet);
        }
      }, 0);
    }
    if (args.isUndo) {
      const dataSource = sheet.getDataSource().getSource();
      store.commit(`quotationModule/${SET_WORK_SHEET_QUOTATION}`, dataSource);
      UpdateTotalBlock(sheet);
    }
  });

  sheet.bind(GC.Spread.Sheets.Events.ValueChanged, (sender, args) => {
    const dataSource = sheet.getDataSource().getSource();
    console.log(dataSource, 'ValueChanged 事件', args);

    const { leavel1Area, leavel2Area } = getPositionBlock();
    // watch classification
    if (leavel2Area.includes(args.row)) {
      // TODO
    } else if (leavel1Area.includes(args.row)) {
      const index = leavel1Area.indexOf(args.row);
      if (index !== -1) {
        const conferenceHall = _.cloneDeep(dataSource.conferenceHall);
        conferenceHall.resourceViews[index].name = args.newValue;
        conferenceHall.resourceViewsMap[conferenceHall.resourceViews[index].resourceLibraryId].name = args.newValue;
        store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
          path: ['conferenceHall'],
          value: conferenceHall
        });
      }
    } else {
      store.commit(`quotationModule/${SET_WORK_SHEET_QUOTATION}`, dataSource);
    }
  });

  sheet.bind(GC.Spread.Sheets.Events.ClipboardPasted, () => {
    const dataSource = sheet.getDataSource().getSource();
    store.commit(`quotationModule/${SET_WORK_SHEET_QUOTATION}`, dataSource);
    UpdateTotalBlock(sheet);
  });
};
/**
 * menu event
 * @param {*} spread
 */
const OnEventMenu = (spread) => {
  commandRegister(spread);
  onOpenMenu(spread);
};

/**
 * config sheet
 * @param {*} spread
 */
const configSheet = (spread) => {
  const sheet = spread.getActiveSheet();
  sheet.options.clipBoardOptions = GC.Spread.Sheets.ClipboardPasteOptions.values;
};
/**
 * config workBook
 * @param {*} spread
 */
const ConfigWorkbook = (spread) => {
  spread.options.allowInvalidFormula = true;
  spread.options.allowUndo = true;
  spread.options.allowUserDragFill = true;
  spread.options.defaultDragFillType = 1;
  configSheet(spread);
};

/**
 * Synchronize Store in SumAmount
 * @param {*} sumAmount
 */
const synchronousStoreSumAmount = (sumAmount) => {
  if (sumAmount === 0 || sumAmount) {
    store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
      path: ['sumAmount'],
      value: sumAmount
    });
    store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
      path: ['DXzje'],
      value: (sumAmount === 0 || sumAmount) ? NzhCN.encodeB(sumAmount) : null
    });
  }
};

/**
 * Get the formula for the calculated column
 * @param {*} sheet
 * @returns
 */
export const GetColumnComputedTotal = (sheet) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const { type = null } = template.cloudSheet.center;
  const resourceViews = quotation.conferenceHall.resourceViews;

  const noClass = resourceViews.length === 1 && resourceViews[0].name === '无分类';

  // Obtain the index of the table
  let header = [];

  // table header
  if (!noClass) {
    if (type) {
      const headerTable = getTableHeaderDataTable(type);
      if (headerTable.length) {
        header = headerTable;
      }
    }
  }

  const { classRow, subTotal, tableHeaderRow } = classificationAlgorithms(quotation, header);

  let insertTableIndex = PubGetTableStartRowIndex();
  const columnTotal = [];
  for (let index = 0; index < resourceViews.length; index++) {
    if (index === 0) {
      const columnTotalMap = columnsTotal(sheet, insertTableIndex + classRow + tableHeaderRow + 1, index, true, null);
      columnTotal.push(columnTotalMap);
      insertTableIndex = insertTableIndex + classRow + tableHeaderRow + resourceViews[index].resources.length;
    } else {
      const columnTotalMap = columnsTotal(sheet, insertTableIndex + subTotal + classRow + tableHeaderRow + 1, index, true, null);
      columnTotal.push(columnTotalMap);
      insertTableIndex = insertTableIndex + subTotal + classRow + tableHeaderRow + resourceViews[index].resources.length;
    }
  }

  return columnTotal;
};

/**
 * Reset the total
 * @param {*} sheet
 * @param {*} quotation
 * @param {*} bottom
 * @param {*} total
 */
const resetTotal = (sheet, quotation, bottom, total, top, mark) => {
  if (mark) {
    const resourceViews = quotation.conferenceHall.resourceViews;
    const totalRowIndex = top.mixCount + resourceViews.length;
    const combined = TOTAL_COMBINED_MAP[CombinationTypeBuild(quotation)];
    sheet.deleteRows(totalRowIndex, mark[combined].rowCount);
  } else {
    const Total = total[CombinationTypeBuild(quotation)];

    const totalRowStartIndex = sheet.getRowCount() - bottom.rowCount - Total.rowCount;
    sheet.deleteRows(totalRowStartIndex, Total.rowCount);
  }
};

/**
 * Specific fields + freight + projectCost
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const totalBeforeAssignment = (fieldName = 'totalBeforeTax', fixedBindCellMap, columnTotalSum, totalBinds = {}) => {
  const { freight = null, projectCost = null } = fixedBindCellMap;
  const rowNames = rowComputedFieldSort(totalBinds);

  const currentField = rowNames.findIndex((name) => name === fieldName);
  const freightIndex = rowNames.findIndex((name) => name === 'freight');
  const projectCostIndex = rowNames.findIndex((name) => name === 'projectCost');
  const sum = [];
  if (columnTotalSum) {
    sum.push(columnTotalSum);
  }
  if (currentField > freightIndex && freight) {
    sum.push(freight);
  }
  if (currentField > projectCostIndex && projectCost) {
    sum.push(projectCost);
  }

  return sum.join('+');
};

/**
 * The sum of all the accumulations【value + 】
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const totalBeforeTaxAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds) => {
  // eslint-disable-next-line no-unused-vars
  const { managementExpense = null, serviceCharge = null, taxes = null } = fixedBindCellMap;
  const totalBeforeFormula = totalBeforeAssignment(fieldName = 'totalBeforeTax', fixedBindCellMap, columnTotalSum, totalBinds);

  const rowNames = rowComputedFieldSort(totalBinds);
  const currentField = rowNames.findIndex((name) => name === fieldName);
  const managementExpenseIndex = rowNames.findIndex((name) => name === 'managementExpense');
  const serviceChargeIndex = rowNames.findIndex((name) => name === 'serviceCharge');
  // const taxesIndex = rowNames.findIndex((name) => name === 'taxes');

  const sum = [];
  if (totalBeforeFormula) {
    sum.push(totalBeforeFormula);
  }

  if (currentField > managementExpenseIndex && managementExpense) {
    sum.push(managementExpense);
  }
  if (currentField > serviceChargeIndex && serviceCharge) {
    sum.push(serviceCharge);
  }
  // if (currentField > taxesIndex && taxes) {
  //   sum.push(taxes);
  // }

  return sum.join('+');
};

/**
 * managementExpense assignment
 * @param {*} fieldName
 * @param {*} fixedBindValueMap
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @returns
 */
const managementExpenseAssignment = (fieldName, fixedBindValueMap, fixedBindCellMap, columnTotalSum) => {
  const managementFeeCell = fixedBindCellMap.managementFee;
  const managementFee = managementFeeCell || fixedBindValueMap.managementFee;

  if (managementFee === 0 || managementFee) {
    if (columnTotalSum) {
      return `(${columnTotalSum}) * ${managementFee} / 100`;
    }
  } else {
    console.error('未检测到有埋点字段管理费率(managementFee)或管理费率字段(managementFee),无法计算管理费');
  }
  return '';
};
/**
 * serviceCharge assignment
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const serviceChargeAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap) => {
  const { rate = null, tax = null } = fixedBindCellMap;
  // eslint-disable-next-line no-unused-vars
  const { taxRate = null } = fixedBindValueMap;

  if (fieldName === 'serviceCharge') {
    const totalBeforeTaxFormula = totalBeforeTaxAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
    if (totalBeforeTaxFormula) {
      const rate_ = rate || fixedBindValueMap[NEW_OLD_FIELD_MAP.rate] || fixedBindValueMap.rate;
      if (rate_ === 0 || rate_) {
        return `(${totalBeforeTaxFormula}) * ${rate_} / 100`;
      } else {
        console.error('未检测到有埋点字段服务费率(rate)或服务费率字段(serviceCharge：旧；rate：新),无法计算服务费');
      }
    }
  } else if (fieldName === 'taxes') {
    const totalBeforeTaxFormula = totalBeforeTaxAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
    if (totalBeforeTaxFormula) {
      const tax_ = tax || fixedBindValueMap[NEW_OLD_FIELD_MAP.tax] || fixedBindValueMap.tax;
      if (tax_ === 0 || tax_) {
        return `(${totalBeforeTaxFormula}) * ${tax_} / 100`;
      } else {
        console.error('未检测到有埋点字段税率(tax)或税率字段(taxRate),无法计算税金');
      }
    }
  }

  return null;
};
/**
 * serviceChargeFee assignment
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 */
// eslint-disable-next-line no-unused-vars
const serviceChargeFeeAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds) => {
  // TODO 未开发功能
};
/**
 * totalServiceCharge assignment
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const totalServiceChargeAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds) => {
  const totalBeforeTaxFormula = totalBeforeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds);
  const { serviceCharge = null } = fixedBindCellMap;
  if (serviceCharge && totalBeforeTaxFormula) {
    return `(${totalBeforeTaxFormula}) + (${serviceCharge})`;
  } else {
    console.error('缺少服务费，无法计算!');
  }
  return null;
};
/**
 * taxes assignment
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const taxesAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap) => {
  const { tax = null, rate = null, totalServiceCharge = null } = fixedBindCellMap;
  if (!totalServiceCharge) {
    if (tax && rate) {
      const totalBeforeTaxFormula = totalBeforeTaxAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);

      if (totalBeforeTaxFormula) {
        const tax_ = tax || fixedBindValueMap[NEW_OLD_FIELD_MAP.tax] || fixedBindValueMap.tax;
        if (tax_ === 0 || tax_) {
          return `(${totalBeforeTaxFormula}) * ${tax_} / 100`;
        } else {
          console.error('未检测到有埋点字段税率(旧：rate；新:tax)或税率字段(旧：taxRate；新：tax),无法计算税金');
        }
      }
    } else if (tax) {
      return serviceChargeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
    }
  } else {
    const tax_ = tax || fixedBindValueMap.tax;
    if (tax_ === 0 || tax_) {
      return `(${totalServiceCharge}) * ${tax_} / 100`;
    } else {
      console.error('未检测到有埋点字段税率(tax)或税率字段(tax),无法计算税金');
    }
  }
  return null;
};
/**
 * addTaxRateBefore assignment
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const addTaxRateBeforeAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds) => {
  const { managementExpense = null, serviceCharge = null, serviceChargeFee = null } = fixedBindCellMap;
  const serviceChargeFormula = serviceChargeFee || serviceCharge;
  const sum1Formula = totalBeforeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds);

  const formula = [];
  if (serviceChargeFormula) {
    formula.push(serviceChargeFormula);
  }

  if (managementExpense) {
    formula.push(managementExpense);
  }
  if (formula.length) {
    return `(${sum1Formula}) + ${formula.join('+')}`;
  }
  return '';
};
/**
 * sumAmount assignment (运费 + 项目费用) + (管理费 + 服务费 + 税金)
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @param {*} fixedBindKeys
 * @param {*} fixedBindValueMap
 * @returns
 */
// eslint-disable-next-line no-unused-vars
const sumAmountAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindKeys, fixedBindValueMap) => {
  const { taxes = null, managementExpense = null, serviceCharge = null, serviceChargeFee = null, freight = null, projectCost = null } = fixedBindCellMap;
  const fieldFormulas = [];
  if (columnTotalSum) {
    fieldFormulas.push(`(${columnTotalSum})`);
  }
  if (freight) {
    fieldFormulas.push(`(${freight})`);
  }
  if (projectCost) {
    fieldFormulas.push(`(${projectCost})`);
  }
  if (serviceCharge) {
    fieldFormulas.push(`(${serviceCharge})`);
  }
  if (serviceChargeFee) {
    fieldFormulas.push(`(${serviceChargeFee})`);
  }
  if (managementExpense) {
    fieldFormulas.push(`(${managementExpense})`);
  }
  if (taxes) {
    fieldFormulas.push(`(${taxes})`);
  }
  return fieldFormulas.join('+');
};

/**
 * Update the sheet according to the new quotation
 * @param {*} sheet
 * @param {*} row
 * @param {*} totalField
 */
const updateTotalValue = (sheet, row, totalField) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  for (const key in totalField.bindPath) {
    if (Object.hasOwnProperty.call(totalField.bindPath, key)) {
      const rows = totalField.bindPath[key];
      if (rows.bindPath) {
        const path = rows.bindPath.split('.');
        if (_.has(quotation, path)) {
          const val = _.get(quotation, path);
          sheet.setValue(row + rows.row, rows.column, val);
        }
      }
    }
  }
};

/**
 * Calculate calculated field value in Total
 * @param {*} sheet
 * @param {*} field
 * @param {*} rowIndex
 * @param {*} fixedBindValueMap
 * @param {*} fixedBindCellMap
 * @param {*} key
 * @param {*} cb
 * @param {*} totalBinds
 * @param {*} cbVal
 * @param {*} columnTotal
 * @param {*} columnTotalSum
 */
const rowComputedField = (sheet, field, rowIndex, fixedBindValueMap, fixedBindCellMap, key, cb, totalBinds, cbVal, columnTotal, columnTotalSum) => {
  let fieldName = key;
  if (!regChineseCharacter.test(field.name)) {
    fieldName = field.name;
  }

  console.log(fieldName, '无bindpath字段');

  if (ASSOCIATED_FIELDS_FORMULA_MAP[fieldName]) {
    if (fieldName === 'managementExpense') {
      const fieldFormula = managementExpenseAssignment(fieldName, fixedBindValueMap, fixedBindCellMap, columnTotalSum);
      const fieldInfo = getFormulaFieldRowCol(field)
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (GenerateFieldsRow().includes(fieldName)) {
      const fieldFormula = totalBeforeTaxAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      const fieldInfo = getFormulaFieldRowCol(field)
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (fieldName === 'serviceCharge') {
      const fieldFormula = serviceChargeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      const fieldInfo = getFormulaFieldRowCol(field)
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (fieldName === 'totalServiceCharge') {
      const fieldFormula = totalServiceChargeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      const fieldInfo = getFormulaFieldRowCol(field)
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (fieldName === 'addTaxRateBefore') {
      const fieldFormula = addTaxRateBeforeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      const fieldInfo = getFormulaFieldRowCol(field)
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (fieldName === 'taxes') {
      const fieldFormula = taxesAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      const fieldInfo = getFormulaFieldRowCol(field)
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (fieldName === 'serviceChargeFee') {
      serviceChargeFeeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      console.error('该总计的组合类型字段(serviceChargeFee)已废弃，建议修改模板总计规则(不符合标准规则)!');
    } else {
      console.warn('在ASSOCIATED_FIELDS_FORMULA_MAP定义,但未存在过相关逻辑的字段', fieldName);
    }
  } else {
    console.warn('模板的总计block：识别出未在ASSOCIATED_FIELDS_FORMULA_MAP定义的字段', fieldName);
  }
};

/**
 * Clears the relevant calculated value for the totals area when no data is available
 * @param {*} sheet
 * @param {*} row
 * @param {*} totalFieldBind
 * @param {*} fixedBindValueMap
 */
const clearTotalNoData = (sheet, row, totalField, fixedBindValueMap, quotation, template) => {
  const totalFieldBind = totalField.bindPath;

  console.log('==================clearTotalNoData==========');
  // Set a fixed value
  for (const key in totalFieldBind) {
    if (Object.hasOwnProperty.call(totalFieldBind, key)) {
      const rows = totalFieldBind[key];
      if (template.truckage) {
        const TruckageIdentifier = new IdentifierTemplate(sheet, 'truckage');
        TruckageIdentifier.truckageFreight(totalField, row, fixedBindValueMap);
      } else {
        if (rows.bindPath) {
          if (Object.keys(DESCRIPTION_MAP).includes(rows.bindPath)) {
            mixedDescriptionFields(sheet, quotation, row, rows);
          }
          if (fixedBindValueMap[rows.bindPath] === 0 || fixedBindValueMap[rows.bindPath]) {
            sheet.setValue(row + rows.row, rows.column, fixedBindValueMap[rows.bindPath]);
          }
          if (rows.bindPath === 'sumAmount') {
            fixedBindValueMap[rows.bindPath] = null;
            sheet.setFormula(row + rows.row, rows.column, '');
            sheet.setValue(row + rows.row, rows.column, '');
            synchronousStoreSumAmount(null);
          }
        }
      }
    }
  }

  // Dynamic fields
  for (const key in totalFieldBind) {
    if (Object.hasOwnProperty.call(totalFieldBind, key)) {
      const rows = totalFieldBind[key];
      if (!rows.bindPath) {
        let fieldName = key;
        if (!regChineseCharacter.test(rows.name)) {
          fieldName = rows.name;
        }

        console.log(row + rows.row, rows.column, rows);

        sheet.setFormula(row + rows.row, rows.column, '');
        sheet.setValue(row + rows.row, rows.column, '');
        fixedBindValueMap[fieldName] = null;
        store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
          path: [fieldName],
          value: null
        });
      }
    }
  }
};

/**
 * Set dynamic field value for total
 * @param {*} sheet
 * @param {*} totalField
 * @param {*} row
 * @param {*} totalBinds
 */
const setTotalRowValue = (sheet, totalField, row, totalBinds, template) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const columnTotal = GetColumnComputedTotal(sheet);
  const columnTotalSum = columnTotalSumFormula(columnTotal);

  const fixedBindValueMap = {};
  const fixedBindCellMap = {};
  const fixedBindKeys = Object.keys(totalField.bindPath);

  // Get a fixed value
  for (const key in totalField.bindPath) {
    if (Object.hasOwnProperty.call(totalField.bindPath, key)) {
      const rows = totalField.bindPath[key];

      console.log(rows, key, 'rows');

      if (rows.bindPath) {
        const path = rows.bindPath.split('.');
        if (_.has(quotation, path)) {
          const val = _.get(quotation, path);
          if (val === 0 || val) {
            if (isNumber(Number(val))) {
              fixedBindValueMap[rows.bindPath] = Number(val);
            }
          }
        }
      }
    }
  }
  if (resourceViews && resourceViews.length) {
    if (!template.truckage) {
      // Set a fixed value
      for (const key in totalField.bindPath) {
        if (Object.hasOwnProperty.call(totalField.bindPath, key)) {
          const rows = totalField.bindPath[key];

          if (rows.bindPath) {
            // Binding value
            if (Object.keys(DESCRIPTION_MAP).includes(rows.bindPath)) {
              mixedDescriptionFields(sheet, quotation, row, rows);
            } else {
              // setCellFormatter(sheet, row + rows.row, rows.column);
              console.log(row, rows.row);

              fixedBindCellMap[key] = `${numberToColumn(rows.column + 1)}${row + rows.row + 1}`;

              if (fixedBindValueMap[rows.bindPath] === 0 || fixedBindValueMap[rows.bindPath]) {
                sheet.setValue(row + rows.row, rows.column, fixedBindValueMap[rows.bindPath]);
              }
              // sheet.autoFitColumn(rows.column);
            }
          }
        }
      }

      // Dynamic fields
      for (const key in totalField.bindPath) {
        if (Object.hasOwnProperty.call(totalField.bindPath, key)) {
          const rows = totalField.bindPath[key];
          let fieldName = key;
          if (!regChineseCharacter.test(rows.name)) {
            fieldName = rows.name;
          }
          if (!rows.bindPath) {
            rowComputedField(sheet, rows, row, fixedBindValueMap, fixedBindCellMap, key, (formula) => {
              fixedBindCellMap[fieldName] = formula;
            }, totalBinds, (value) => {
              fixedBindValueMap[fieldName] = value;
            }, columnTotal, columnTotalSum);
          }
        }
      }

      // Calculate the final total (sumAmount)
      for (const key in totalField.bindPath) {
        if (Object.hasOwnProperty.call(totalField.bindPath, key)) {
          const rows = totalField.bindPath[key];
          if (rows.bindPath === 'sumAmount') {
            const fieldFormula = sumAmountAssignment(key, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindKeys);

            console.log(fieldFormula, 'sumAmount');

            sheet.setFormula(row + rows.row, rows.column, fieldFormula);
            const val = sheet.getValue(row + rows.row, rows.column);
            if (val === 0 || val) {
              synchronousStoreSumAmount(val);
            }
          }
        }
      }
    } else {
      const TruckageIdentifier = new IdentifierTemplate(sheet, 'truckage');
      TruckageIdentifier.truckageFreight(totalField, row, fixedBindValueMap);
    }

    if (!template.truckage) {
      updateTotalValue(sheet, row, totalField);
    }
  } else {
    clearTotalNoData(sheet, row, totalField, fixedBindValueMap, quotation, template);
  }
  store.commit(`quotationModule/${IGNORE_EVENT}`, false);

  console.log(fixedBindValueMap, 'fixedBindValueMap');
  console.log(fixedBindCellMap, 'fixedBindCellMap');
};

/**
 * Update dynamic field value for total
 * @param {*} sheet
 * @param {*} totalRowIndex
 */
export const updateTotalRowValue = (sheet, totalRowIndex) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const { top, total, bottom, mixTopTotal = null } = template.cloudSheet;

  if (showTotal()) {
    if (mixTopTotal) {
      const initTotal = mixTopTotal.initTotal;
      const combined = TOTAL_COMBINED_MAP[CombinationTypeBuild(quotation)];
      const resourceViews = quotation.conferenceHall.resourceViews;
      const totalRowIndex = top.mixCount + resourceViews.length;
      sheet.suspendPaint();
      setTotalRowValue(sheet, mixTopTotal[combined], totalRowIndex, initTotal.bindPath, template);
      sheet.resumePaint();
    } else {
      const Total = total[CombinationTypeBuild(quotation)];
      console.log(Total, CombinationTypeBuild(quotation));
      let startRowIndex = sheet.getRowCount() - bottom.rowCount - Total.rowCount;
      if (totalRowIndex === 0 || totalRowIndex) {
        startRowIndex = totalRowIndex;
      }
      sheet.suspendPaint();
      setTotalRowValue(sheet, Total, startRowIndex, Total.bindPath, template);
      sheet.resumePaint();
    }
  }
};

/**
 * Subtotal assignment
 * @param {*} sheet
 */
export const updateSubTotalRowValue = (sheet) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const { type = null, total = null } = template.cloudSheet.center;

  if (!template.truckage) {
    if (resourceViews.length) {
      const noClass = resourceViews.length === 1 && resourceViews[0].name === '无分类';

      // Obtain the index of the table
      const tableStartRowIndex = PubGetTableStartRowIndex();

      let header = [];

      // table header
      if (!noClass) {
        if (type) {
          const headerTable = getTableHeaderDataTable(type, true);
          if (headerTable.length) {
            header = headerTable;
          }
        }
      }

      const { classRow, subTotal, tableHeaderRow } = classificationAlgorithms(quotation, header);

      const columnComputed = [];
      let insertTableIndex = tableStartRowIndex;
      const subTotalBindPath = total ? total.bindPath : null;
      sheet.suspendPaint();
      for (let index = 0; index < resourceViews.length; index++) {
        if (index === 0) {
          const columnTotalMap = columnsTotal(sheet, insertTableIndex + classRow + tableHeaderRow + 1, index, true, columnComputed, subTotalBindPath);
          !noClass && SetComputedSubTotal(sheet, columnTotalMap, subTotalBindPath);
          insertTableIndex = insertTableIndex + classRow + tableHeaderRow + resourceViews[index].resources.length;
        } else {
          const columnTotalMap = columnsTotal(sheet, insertTableIndex + subTotal + classRow + tableHeaderRow + 1, index, true, null, subTotalBindPath);
          !noClass && SetComputedSubTotal(sheet, columnTotalMap, subTotalBindPath);

          insertTableIndex = insertTableIndex + subTotal + classRow + tableHeaderRow + resourceViews[index].resources.length;
        }
      }
      sheet.resumePaint();
    }
  }
};

/**
 * Update the total block
 * @param {*} sheet
 */
export const UpdateTotalBlock = (sheet) => {
  console.log('UpdateTotalBlock');
  // Update the subtotal
  updateSubTotalRowValue(sheet);
  // Update totals
  const totalRowIndex = getTotalStartRowIndex(sheet);
  if (totalRowIndex !== null) {
    updateTotalRowValue(sheet, totalRowIndex);
  } else {
    console.error('更新总计计算出错(总计区域的开始行 索引获取失败)!');
  }
};

/**
 * Insert the specified combination type field
 * @param {*} spread
 * @param {*} fileName
 * @param {*} value
 */
export const insertField = (spread, fileName, value) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const { top, total, bottom, mixTopTotal = null } = template.cloudSheet;

  if (showTotal()) {
    const sheet = spread.getActiveSheet();

    const quotation = store.getters['quotationModule/GetterQuotationInit'];
    resetTotal(sheet, quotation, bottom, total, top, mixTopTotal);

    if (value) {
      switch (fileName) {
        case 'tax':
          store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
            path: ['taxRate'], // TODO 待优化 新旧版字段问题
            value: value
          });
          break;
        case 'rate':
          store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
            path: ['serviceCharge'], // TODO 待优化 新旧版字段问题
            value: value
          });
          break;
        case 'concessional':
          store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
            path: ['concessionalRate'], // TODO 待优化 新旧版字段问题
            value: value
          });
          break;
        case 'freight':
          store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
            path: ['freight'],
            value: value
          });
          break;
        case 'managementFee':
          store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
            path: ['managementFee'],
            value: value
          });
          break;
        case 'projectCost':
          store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
            path: ['projectCost'],
            value: value
          });
          break;
        default:
          break;
      }
    } else {
      switch (fileName) {
        case 'tax':
          store.commit(`quotationModule/${DELETE_QUOTATION_PATH}`, {
            path: ['taxRate'] // TODO 待优化 新旧版字段问题
          });
          break;
        case 'rate':
          store.commit(`quotationModule/${DELETE_QUOTATION_PATH}`, {
            path: ['serviceCharge'] // TODO 待优化 新旧版字段问题
          });
          break;
        case 'concessional':
          store.commit(`quotationModule/${DELETE_QUOTATION_PATH}`, {
            path: ['concessionalRate'] // TODO 待优化 新旧版字段问题
          });
          break;
        case 'freight':
          store.commit(`quotationModule/${DELETE_QUOTATION_PATH}`, {
            path: ['freight']
          });
          break;
        case 'managementExpense':
          store.commit(`quotationModule/${DELETE_QUOTATION_PATH}`, {
            path: ['managementFee'] // TODO 待优化 新旧版字段问题
          });
          break;
        case 'projectCost':
          store.commit(`quotationModule/${DELETE_QUOTATION_PATH}`, {
            path: ['projectCost']
          });
          break;
        default:
          break;
      }
    }

    rendering(spread, 'insert', template);

    store.commit(`quotationModule/${UPDATE_HEAD_TOTAL}`, {
      key: fileName,
      value: !!((value === 0 || value))
    });
  }

  const layout = new LayoutRowColBlock(spread);
  const { Tables, TotalMap } = layout.getLayout();

  const costPrice = new CheckCostPrice(spread, template, store.getters['quotationModule/GetterQuotationInit']);
  costPrice.updateTotalPosition(Tables, TotalMap)

};

/**
 * remove all tables
 * @param {*} sheet
 * @param {*} quotation
 */
export const removeAllTable = (sheet, quotation) => {
  const resourceViews = quotation.conferenceHall.resourceViews;
  const noClass = resourceViews.length === 1 && resourceViews[0].name === '无分类';

  let startIndex = null;
  let endIndex = null;
  const tables = sheet.tables.all();
  for (let index = 0; index < tables.length; index++) {
    const { row, rowCount } = tables[index].range();
    if (index === 0) {
      startIndex = row;
    }
    if (index === tables.length - 1) {
      endIndex = row + rowCount;
    }
  }

  if (!noClass) {
    const { classRow } = classificationAlgorithms(quotation);
    if (startIndex !== null) {
      startIndex = startIndex - classRow;
    }
    if (endIndex !== null) {
      endIndex = endIndex + 1;
    }
  }

  sheet.suspendPaint();
  sheet.deleteRows(startIndex, endIndex - startIndex);
  sheet.resumePaint();
};

const getTotalStartRowIndex = (sheet) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const topBottomIndex = template.cloudSheet.top.rowCount;
  const { subTotal } = classificationAlgorithms(quotation);
  const noClass = resourceViews.length === 1 && resourceViews[0].name === '无分类';

  const tables = sheet.tables.all();
  let startIndex = null;

  if (tables.length) {
    if (noClass) {
      for (let index = 0; index < tables.length; index++) {
        const table = tables[index];
        if (index === tables.length - 1) {
          const { row, rowCount } = table.range();
          startIndex = row + rowCount;
        }
      }
    } else {
      for (let index = 0; index < tables.length; index++) {
        const table = tables[index];
        if (index === tables.length - 1) {
          const { row, rowCount } = table.range();
          console.log(subTotal);
          startIndex = row + rowCount + subTotal;
        }
      }
    }
  } else {
    startIndex = topBottomIndex;
  }

  return startIndex;
};

/**
 * Directly manipulate the synchronization after SpreadJS(quotaion、subtotal、totals、BlockArea)
 * @param {*} sheet
 */
export const OperationWorkBookSync = (sheet) => {
  // Sync quotation data
  const dataSource = sheet.getDataSource().getSource();
  // update sort
  UpdateSort(dataSource.conferenceHall);
  store.commit(`quotationModule/${SET_WORK_SHEET_QUOTATION}`, dataSource);
  // Update the subtotal
  updateSubTotalRowValue(sheet);
  // Update totals
  const totalRowIndex = getTotalStartRowIndex(sheet);
  if (totalRowIndex !== null) {
    updateTotalRowValue(sheet, totalRowIndex);
  } else {
    console.error('更新总计计算出错(总计区域的开始行 索引获取失败)!');
  }
  // Update the location
  positionBlock(sheet);
};

/**
 * update Cell Value
 * @param {*} sheet
 * @param {*} value
 * @param {*} row
 * @param {*} col
 * @param {*} valueType
 */
export const updateCellValue = (sheet, value, row, col, valueType = 'string') => {
  if (valueType === 'string') {
    sheet.setValue(row, col, value);
  } else if (valueType === 'date') {
    sheet.setValue(row, col, FormatDate(value, 'YYYY/MM/DD'));
  }
};

/**
 * init sheet
 * @param {*} spread
 */
export const InitSheet = (spread) => {
  ConfigWorkbook(spread);
  OnEventBind(spread);
  OnEventMenu(spread);
};

/**
 * Insert total
 * @param {*} spread
 * @param {*} insertFieldType
 * @returns
 */
export const InsertTotal = (spread, insertFieldType) => {
  if (Object.keys(insertFieldType).length > 1) {
    console.warn('添加的类型未进行逻辑处理!');
  }
  if (!spread) {
    console.error('spreadJS初始化失败!');
    return;
  }

  for (const key in insertFieldType) {
    if (Object.hasOwnProperty.call(insertFieldType, key)) {
      const value = Number(insertFieldType[key].value);
      insertField(spread, key, value);
    }
  }
};

export const RenderHeaderClass = () => {
  // TODO 未开发
};

/**
 * Render total
 * @param {*} spread
 * @param {*} isInit
 */
export const RenderTotal = (spread, isInit) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];

  if (showTotal()) {
    const sheet = spread.getActiveSheet();
    if (!isInit) {
      const { total, bottom } = template.cloudSheet;
      const Total = total[CombinationTypeBuild(quotation)];

      const totalRowStartIndex = sheet.getRowCount() - bottom.rowCount - Total.rowCount;

      sheet.deleteRows(totalRowStartIndex, Total.rowCount);
      rendering(spread, 'insert', template);
    } else {
      rendering(spread, 'render', template);
    }
    sheet.bind(GC.Spread.Sheets.Events.ValueChanged, (e, info) => {
      const cell = [];
      const columnTotal = GetColumnComputedTotal(sheet);
      columnTotal.forEach(tableColumnsItem => {
        if (!Object.keys(tableColumnsItem).includes('total')) {
          console.error('缺少列合计字段(total),无法计算!');
          return;
        }
        const formulas = tableColumnsItem.total.formula.split('+');
        formulas.forEach(formula => {
          const column = tableColumnsItem.total.column;
          const row = formula.slice(1);
          cell.push({
            row: Number(row) - 1,
            column
          });
        });
      });
      for (let index = 0; index < cell.length; index++) {
        const row = [];
        const col = [];
        const getPrecedents = sheet.getPrecedents(cell[index].row, cell[index].column);
        for (let i = 0; i < getPrecedents.length; i++) {
          row.push(getPrecedents[i].row);
          col.push(getPrecedents[i].col);
        }
        for (let i = 0; i < row.length; i++) {
          if (info.row === row[i] && info.col === col[i]) {
            setTimeout(() => {
              updateSubTotalRowValue(sheet);
              updateTotalRowValue(sheet);
            }, 0);
          }
        }
      }
    });
  }
};

/**
 * start render
 * @param {*} spread
 * @param {*} type
 * @param {*} quotation
 * @param {*} template
 */
const rendering = (spread, type, template) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const { top, total, bottom, mixTopTotal = null } = template.cloudSheet;
  const sheet = spread.getActiveSheet();

  if (mixTopTotal) {
    const resourceViews = quotation.conferenceHall.resourceViews;
    const totalRowIndex = top.mixCount + resourceViews.length;
    let combined = TOTAL_COMBINED_MAP[CombinationTypeBuild(quotation)];

    sheet.suspendPaint();
    if (type === 'render') {
      combined = TOTAL_COMBINED_MAP[total.select];
      const initTotal = mixTopTotal.initTotal;
      sheet.deleteRows(totalRowIndex, initTotal.rowCount);
    }

    sheet.addRows(totalRowIndex, mixTopTotal[combined].rowCount);

    if (mixTopTotal[combined].spans) {
      mergeSpan(sheet, mixTopTotal[combined].spans, totalRowIndex);
    }
    setCellStyle(spread, mixTopTotal[combined], totalRowIndex, true);
    setTotalRowHeight(sheet, total, mixTopTotal[combined], totalRowIndex);
    sheet.resumePaint();
    updateTotalRowValue(sheet, totalRowIndex);

    RenderHeaderClass();
  } else {
    const bottomRowCount = bottom.rowCount;
    const rowCount = sheet.getRowCount();

    const totalRowIndex = rowCount - bottomRowCount;
    let Total = total[CombinationTypeBuild(quotation)];
    if (type === 'render') {
      Total = total[templateTotalMap(total.select)];
    }
    console.log(rowCount, bottomRowCount);

    sheet.suspendPaint();

    console.log(totalRowIndex, Total);

    sheet.addRows(totalRowIndex, Total.rowCount);

    mergeSpan(sheet, Total.spans, totalRowIndex);
    setCellStyle(spread, Total, totalRowIndex, true);
    setTotalRowHeight(sheet, total, Total, totalRowIndex);
    sheet.resumePaint();
    updateTotalRowValue(sheet, totalRowIndex);
  }

  if (template.truckage) {
    const TruckageIdentifier = new IdentifierTemplate(sheet, 'truckage');
    TruckageIdentifier.truckageRenderTotal(quotation);
  }
};

/**
 * Render
 * @param {*} spread
 * @param {*} isInit
 */
export const Render = (spread, isInit) => {
  store.commit(`quotationModule/${IGNORE_EVENT}`, true);
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  console.log(quotation, 'quotation');
  console.log(template, 'template');
  const sheet = spread.getActiveSheet();

  const { equipment, type = null, total = null, columnCount } = template.cloudSheet.center;
  const { mixTopTotal = null, image = null, top } = template.cloudSheet;
  const { mixRender, classType, isHaveChild } = templateRenderFlag();
  const resourceViews = quotation.conferenceHall.resourceViews;

  const noClass = resourceViews.length === 1 && resourceViews[0].name === '无分类';

  if (!noClass) {
    if (type) {
      delTableHeaderRowCount(sheet, type.dataTable, top);
    }
  }

  // Obtain the index of the table
  const tableStartRowIndex = PubGetTableStartRowIndex();
  const tableStartColumnIndex = PubGetTableStartColumnIndex();
  const tableColumnCount = PubGetTableColumnCount();

  sheet.suspendPaint();

  let rowClassIndex = tableStartRowIndex;
  const headerIndexs = [];
  let header = [];

  // table header
  if (!noClass) {
    if (type) {
      const headerTable = getTableHeaderDataTable(type, true);
      if (headerTable.length) {
        header = headerTable;
      }
    }
  }

  const { classRow, subTotal, classRow1, tableHeaderRow } = classificationAlgorithms(quotation, header);

  console.log(' classRow, subTotal, classRow1, tableHeaderRow ');
  console.log(classRow, subTotal, classRow1, tableHeaderRow);

  // 判断当前模板是否需要在分类下渲染表头
  const identifier = new IdentifierTemplate()
  console.log(identifier.builtInIdsIdentifier());
  if (identifier.builtInIdsIdentifier() === 'BH') {
    // 在分类下添加表头
    // if (type && type.dataTable) {
    //   // 分类下如果存在表头，则展示表头
    //   console.log('分类下的表头');
    //   console.log(rowClassIndex);

    //   // sheet.addRows(rowClassIndex + classRow + tableHeaderRow, 1, GC.Spread.Sheets.SheetArea.viewport);
    // }
  }

  // Render classification
  for (let i = 0; i < resourceViews.length; i++) {
    const headerRow = 1;

    headerIndexs.push(rowClassIndex + classRow + tableHeaderRow);
    const rowCount = classRow + tableHeaderRow + resourceViews[i].resources.length + headerRow + subTotal;

    // Classifications are displayed in rows
    if (!noClass) {
      if (classType === 'mergeClass') {
        // rowCount = resourceViews[i].resources.length + headerRow;
      } else {
        // Add a categorical row
        sheet.addRows(rowClassIndex, classRow, GC.Spread.Sheets.SheetArea.viewport);
        if (mixRender) {
          sheet.setValue(rowClassIndex, 0, `${numberToColumn(i + 1)} ${resourceViews[i].parentTypeName}`);
          if (!type) {
            const rows = [rowClassIndex];
            if (classRow1) {
              rows.push(rowClassIndex + classRow1);
            }

            const { style } = GeneratorCellStyle('className', { textIndent: 1 });
            mergeRow(sheet, rows, 0, 1, columnCount);

            for (let index = 0; index < rows.length; index++) {
              sheet.setStyle(rows[index], -1, style, GC.Spread.Sheets.SheetArea.viewport);
              sheet.getRange(rows[index], 0, 1, columnCount).setBorder(GeneratorLineBorder(), { all: true });
            }

            if (equipment.height === 0 || equipment.height) {
              for (let index = 0; index < rows.length; index++) {
                sheet.setRowHeight(rows[index], equipment.height);
              }
            }
          }
        }
        // Head classification name
        if (mixRender) {
          if (!resourceViews[i].parentTypeName) {
            sheet.setValue(rowClassIndex + classRow1, 0, `${numberToColumn(i + 1)} ${resourceViews[i].name}`);
          }
        }
        sheet.setValue(rowClassIndex + classRow1, 0, resourceViews[i].name);
      }

      // Merge categorical cells
      if (type) {
        mergeSpan(sheet, type.spans, rowClassIndex);
        tableHeader(sheet, header, rowClassIndex + classRow + tableHeaderRow, type.height);
        setCellStyle(spread, type, rowClassIndex, false);
        PubSetCellHeight(sheet, type, rowClassIndex);
      }
    }

    // Add a list of products
    sheet.addRows(rowClassIndex + classRow + tableHeaderRow, PubGetTableRowCount(i), GC.Spread.Sheets.SheetArea.viewport);
    // Create a table
    const tableId = resourceViews[i].resourceLibraryId;
    const table = sheet.tables.findByName('table' + tableId);
    if (!table) {
      const { bindPath } = equipment;
      CreateTable(sheet, tableId, rowClassIndex + classRow + tableHeaderRow, tableStartColumnIndex, PubGetTableRowCount(i), tableColumnCount, bindPath, `conferenceHall.resourceViewsMap.${tableId}.resources`);
    }

    // Initialize the product cell style
    const computedColumnFormula = getComputedColumnFormula(equipment.bindPath);
    const resources = resourceViews[i].resources;
    for (let index = 0; index < resources.length; index++) {

      console.log(rowClassIndex, classRow, tableHeaderRow, headerRow, index);

      const startRow = rowClassIndex + classRow + tableHeaderRow + headerRow + index;

      const { image = null } = template.cloudSheet;

      mergeSpan(sheet, equipment.spans, startRow);
      setRowStyle(sheet, equipment, startRow, image);
      columnComputedValue(sheet, equipment, startRow, computedColumnFormula);
      sheet.getCell(startRow, 0).locked(true);
    }

    // Classifications are displayed by columns
    if (classType === 'mergeClass') {
      // Merge column categorical cells
      mergeColumn(sheet, equipment.bindPath, rowClassIndex + classRow + tableHeaderRow + headerRow + subTotal, resourceViews[i].resources.length, 1, 'classname');
    } else if (isHaveChild) {
      mergeColumn(sheet, equipment.bindPath, rowClassIndex + classRow + tableHeaderRow + headerRow, resourceViews[i].resources.length, 1, 'pname');
    }

    // Add subtotal rows
    if (!noClass) {
      if (showSubTotal()) {
        const rowClassTotal = rowClassIndex + classRow + tableHeaderRow + PubGetTableRowCount(i) + headerRow;

        console.log(rowClassIndex, classRow, tableHeaderRow, PubGetTableRowCount(i), headerRow);
        console.log(rowClassTotal, subTotal);

        sheet.addRows(rowClassTotal, subTotal, GC.Spread.Sheets.SheetArea.viewport);
        if (total) {
          mergeSpan(sheet, total.spans, rowClassTotal);
          setCellStyle(spread, total, rowClassTotal, true);
          PubSetCellHeight(sheet, total, rowClassTotal);
          for (const key in total.bindPath) {
            if (Object.hasOwnProperty.call(total.bindPath, key)) {
              const field = total.bindPath[key];
              if (field.bindPath) {
                // setCellFormatter(sheet, rowClassTotal + field.row, field.column);
                sheet.setBindingPath(rowClassTotal + field.row, field.column, field.bindPath);
                // sheet.autoFitColumn(field.column);
              }
            }
          }
        }
      }
    }

    rowClassIndex = rowClassIndex + rowCount;
  }

  // Delete headers in batches
  for (let i = 0; i < headerIndexs.length; i++) {
    sheet.deleteRows(headerIndexs[i] - i, 1);
  }

  // Add product images
  if (image) {
    renderSheetImage(spread, tableStartRowIndex, true, false, false);
  }

  // Subtotal assignment
  const columnTotal = [];
  const columnComputed = [];
  let insertTableIndex = tableStartRowIndex;
  const subTotalBindPath = total ? total.bindPath : null;
  for (let index = 0; index < resourceViews.length; index++) {
    if (index === 0) {
      const columnTotalMap = columnsTotal(sheet, insertTableIndex + classRow + tableHeaderRow + 1, index, true, columnComputed, subTotalBindPath);
      !noClass && SetComputedSubTotal(sheet, columnTotalMap, subTotalBindPath);
      columnTotal.push(columnTotalMap);
      insertTableIndex = insertTableIndex + classRow + tableHeaderRow + resourceViews[index].resources.length;
    } else {
      const columnTotalMap = columnsTotal(sheet, insertTableIndex + subTotal + classRow + tableHeaderRow + 1, index, true, null, subTotalBindPath);
      !noClass && SetComputedSubTotal(sheet, columnTotalMap, subTotalBindPath);
      columnTotal.push(columnTotalMap);

      insertTableIndex = insertTableIndex + subTotal + classRow + tableHeaderRow + resourceViews[index].resources.length;
    }
  }

  sheet.resumePaint();

  if (mixTopTotal) {
    // TODO
  } else {
    RenderTotal(spread, !!isInit);
  }

  positionBlock(sheet);

  if (isInit) {
    initShowCostPrice(spread);
  } else {
    ShowCostPrice(spread);
  }
};

/**
 * Set the distribution locations for classifications, subtotals, and totals
 * @param {*} sheet
 */
export const positionBlock = (sheet) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  if (resourceViews.length) {
    const noClass = resourceViews.length === 1 && resourceViews[0].name === '无分类';

    const { bottom = null, top } = template.cloudSheet;
    const tables = sheet.tables.all();
    const sheetRowCount = sheet.getRowCount();
    const leavel1 = [];
    const leavel2 = [];
    const subTotal = [];
    const total = [];

    let startTotal = 0;
    let endTotal = sheetRowCount - 1;
    if (tables.length) {
      for (let index = 0; index < tables.length; index++) {
        const tableRange = tables[index].range();
        const { row, col, rowCount, colCount } = tableRange;
        startTotal = row + rowCount;
        if (!noClass) {
          // TODO 是否存在二级分类

          // if ('存在二级分类') {
          //   leavel2.push({
          //     row: row - 1, col, rowCount: 1, colCount
          //   });
          //   leavel1.push({
          //     row: row - 2, col, rowCount: 1, colCount
          //   });
          // } else {
          leavel1.push({
            row: row - 1, col, rowCount: 1, colCount
          });
          // }
          subTotal.push({
            row: row + rowCount, col, rowCount: 1, colCount
          });
          if (index === tables.length - 1) {
            startTotal = row + rowCount + 1;
          }
        } else {
          if (bottom) {
            endTotal = sheetRowCount - bottom.rowCount - 1;
          }
        }
      }

      total.push({
        row: startTotal,
        rowCount: endTotal - startTotal
      });
    } else {
      if (bottom) {
        endTotal = sheetRowCount - bottom.rowCount - 1;
      }
    }
    // TODO 是否存在顶部总计
    startTotal = sheetRowCount - top.rowCount;

    store.commit(`quotationModule/${SET_SHEET_LEAVEL_POSITION}`, {
      sheetLeavel_1_position: leavel1,
      sheetLeavel_2_position: leavel2,
      sheetTotal_position: subTotal
    });
  } else {
    store.commit(`quotationModule/${SET_SHEET_LEAVEL_POSITION}`, {
      sheetLeavel_1_position: [],
      sheetLeavel_2_position: [],
      sheetTotal_position: []
    });
  }
};

/**
 * Set the project name in the spread
 * @param {*} spread
 * @param {*} projectName
 */
export const setProjectName = (spread, projectName, projectNameField) => {
  const sheet = spread.getActiveSheet();
  sheet.suspendPaint();
  sheet.setValue(projectNameField.row, projectNameField.column, projectName);
  sheet.resumePaint()
};

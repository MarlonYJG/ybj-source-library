/*
 * @Author: Marlon
 * @Date: 2024-04-07 18:29:42
 * @Description: 单表 - 逻辑处理
 */
import Decimal from 'decimal.js';
import * as GC from '@grapecity/spread-sheets';
import _ from 'lodash';
import store from 'store';
import { GetUserInfoDetail, GetUserCompany, imgUrlToBase64 } from 'utils';
import { getSystemDate } from 'utils/date';
import { ResDatas } from 'utils/res-format';
import { regChineseCharacter } from 'utils/regular-expression';
import API from 'api';

import { CreateTable } from '../common/sheetWorkBook';
import { GeneratorStyle, GeneratorLineBorder } from '../common/generator';
import { TOTAL_COMBINED_MAP, ASSOCIATED_FIELDS_FORMULA_MAP, DESCRIPTION_MAP, REGULAR } from '../common/constant';

import { numberToColumn } from '../common/public'

import IdentifierTemplate from '../common/identifier-template'

import {
  PubGetTableStartColumnIndex,
  PubGetTableColumnCount,
  templateRenderFlag,
  delTableHeaderRowCount,
  getTableHeaderDataTable,
  mergeRow,
  PubSetCellHeight,
  setRowStyle,
  mergeColumn,
  showTotal,
  getComputedColumnFormula,
  getFormulaFieldRowCol
} from '../common/parsing-template';
import {
  templateTotalMap, GenerateFieldsRow, mergeSpan, setCellStyle, setTotalRowHeight, PubGetTableStartRowIndex,
  PubGetTableRowCount,
  classificationAlgorithms,
  rowComputedFieldSort,
  plusColumnTotalSum,
  columnsTotal,
  mixedDescriptionFields,
  tableHeader,
  columnComputedValue,
  setCellFormatter,
  SetComputedSubTotal,
  renderSheetImage,
  translateSheet
} from '../common/single-table';

import {
  PubGetResourceViews,
  // eslint-disable-next-line no-unused-vars
  setLastColumnWidth
} from './public';

/**
 * Get an index of the first-level classification
 * Gets an index of the subtotal
 * @returns
 */
// eslint-disable-next-line no-unused-vars
const getOneClassAndSubRowsIndex = () => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const tableStartRowIndex = PubGetTableStartRowIndex();

  const classIndexs = [];
  const classSubs = [];

  let insertClassIndex = tableStartRowIndex;
  let insertClassSub = tableStartRowIndex;

  for (let index = 0; index < resourceViews.length; index++) {
    const classRow = 1;
    const subTotal = 1;

    if (index === 0) {
      insertClassSub = insertClassSub + classRow + resourceViews[index].resources.length;
    } else {
      insertClassIndex = insertClassIndex + classRow + resourceViews[index].resources.length + subTotal;

      insertClassSub = insertClassSub + subTotal + classRow + resourceViews[index].resources.length;
    }

    classIndexs.push(insertClassIndex);
    classSubs.push(insertClassSub);
  }

  return {
    oneClassIndex: classIndexs,
    subClassIndex: classSubs
  };
};

/**
 * Delete the bottom
 * @param {*} spread
 * @param {*} template
 */
// eslint-disable-next-line no-unused-vars
const deleteBottomRow = (spread, template) => {
  const sheet = spread.getActiveSheet();
  const bottom = template.cloudSheet.bottom;

  sheet.suspendPaint();

  sheet.deleteRows(PubGetTableStartRowIndex(), bottom.rowCount);

  sheet.resumePaint();
};

/**
 * Get the index of the start position of the bottom
 * @param {*} spread
 * @param {*} template
 * @returns
 */
// eslint-disable-next-line no-unused-vars
const getBottomStartRowIndex = (spread) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const { rowCount } = template.cloudSheet.top;
  const { center, total } = template.cloudSheet;
  const resourceViews = PubGetResourceViews(spread, quotation);

  // 产品分类及产品之和的总个数
  let startRowIndex = rowCount;
  if (resourceViews.length === 1 && resourceViews[0].name === '无分类') {
    for (let i = 0; i < resourceViews.length; i++) {
      startRowIndex = startRowIndex + resourceViews[i].resources.length;
    }
  } else {
    for (let i = 0; i < resourceViews.length; i++) {
      startRowIndex = startRowIndex + resourceViews[i].resources.length + center.rowCount - center.equipment.rowCount;
    }
  }

  // 总计所占行的个数
  if (showTotal()) {
    const totalRowCount = total[templateTotalMap(total.select)].rowCount;
    if (totalRowCount) {
      startRowIndex = startRowIndex + totalRowCount;
    }
  }

  return startRowIndex;
};

// 家具建材默认总量计算方式
const setDefaultCalMethod = (sheet, template, row) => {
  // 1、宽高同时存在或者仅存在高时，总量 = 宽*高*数量   2、仅存在宽时，总量 = 宽*数量   3、宽高都没有时，总量 = 数量
  const eqBindPathList = Object.keys(template.cloudSheet.center.equipment.bindPath);
  const curReWidthIndex = eqBindPathList.indexOf('resourceWidth');
  const curReHeightIndex = eqBindPathList.indexOf('resourceHeight');
  const curAreaTotalIndex = eqBindPathList.indexOf('areaTotal');
  const curAreaTotal = template.cloudSheet.center.equipment.bindPath.areaTotal;
  // 监听宽的单元格值的变化
  if (curAreaTotalIndex != -1 && curReWidthIndex != -1 && curReHeightIndex != -1) {
    const curReWidthVal = sheet.getValue(row, template.cloudSheet.center.equipment.bindPath.resourceWidth.column);
    const curReHeightVal = sheet.getValue(row, template.cloudSheet.center.equipment.bindPath.resourceHeight.column);
    const eqFmAreaT = Object.assign({}, curAreaTotal.formula);
    if (curReWidthVal && !curReHeightVal) {
      for (let i = 0; i < eqFmAreaT.wSize; i++) {
        eqFmAreaT.formulaW = eqFmAreaT.formulaW.replace('{{' + i + '}}', row + 1);
      }
      sheet.setFormula(row, eqFmAreaT.column, eqFmAreaT.formulaW);
    } else if (!curReWidthVal && !curReHeightVal) {
      for (let i = 0; i < eqFmAreaT.quaSize; i++) {
        eqFmAreaT.formulaQua = eqFmAreaT.formulaQua.replace('{{' + i + '}}', row + 1);
      }
      sheet.setFormula(row, eqFmAreaT.column, eqFmAreaT.formulaQua);
    } else {
      for (let i = 0; i < eqFmAreaT.size; i++) {
        eqFmAreaT.formula = eqFmAreaT.formula.replace('{{' + i + '}}', row + 1);
      }
      sheet.setFormula(row, eqFmAreaT.column, eqFmAreaT.formula);
    }
  }
};

/**
 * Bind company's logo
 * @param {*} spread
 * @param {*} template
 * @param {*} base64
 * @returns
 */
export const InsertLogo = (spread, template, base64) => {
  if (!spread) return;
  if (!base64) return;
  if (template.cloudSheet.logo) {
    const sheet = spread.getActiveSheet();
    sheet.suspendPaint();
    let picture = sheet.pictures.get(template.cloudSheet.logo.name);
    if (picture) {
      picture.src(base64);
      sheet.resumePaint();
      return;
    }
    picture = sheet.pictures.add(template.cloudSheet.logo.name, base64, template.cloudSheet.logo.x, template.cloudSheet.logo.y, template.cloudSheet.logo.width, template.cloudSheet.logo.height);
    picture.allowMove(false);
    picture.allowResize(true);
    picture.isLocked(true);
    sheet.resumePaint();
  }
};

/**
 * Initialize the company seal
 * @param {*} spread
 * @param {*} template
 * @param {*} base64
 */
const InsertSeal = (spread, template, base64) => {
  const sheet = spread.getActiveSheet();
  sheet.suspendPaint();
  const row = sheet.getRowCount();// 表格所有行
  const sealName = template.cloudSheet.seal.name;// 公司印章名称
  const sealH = template.cloudSheet.seal.height;// 公司印章高度
  const sealW = template.cloudSheet.seal.width;// 公司印章宽度
  const rowOffsetLeft = template.cloudSheet.seal.rowOffsetLeft;// 公司印章距离表格左侧的列数
  const bottomR = template.cloudSheet.seal.columntBottm;// 公司印章距离底部行数
  let offsetLeft = 0;// 公司印章距离表格左侧距离
  let offsetTop = 0;// 公司印章距离顶部距离
  let bottomH = sealH;// 公司印章距离底部高度

  // 计算距离左侧距离
  for (let i = 0; i < rowOffsetLeft; i++) {
    offsetLeft = offsetLeft + sheet.getColumnWidth(i);
  }
  // 计算距离顶部距离
  for (let i = 0; i < row; i++) {
    offsetTop = offsetTop + sheet.getRowHeight(i);
  }
  // 计算距离底部距离
  for (let i = row - bottomR; i < row; i++) {
    bottomH = bottomH + sheet.getRowHeight(i);
  }
  offsetTop = offsetTop - bottomH;
  let picture = sheet.pictures.get(sealName);

  if (picture) {
    sheet.pictures.remove(sealName);
    picture = sheet.pictures.add(sealName, base64, offsetLeft, offsetTop, sealW, sealH);
  } else {
    picture = sheet.pictures.add(sealName, base64, offsetLeft, offsetTop, sealW, sealH);
  }
  picture.allowMove(false);
  picture.allowResize(true);
  picture.isLocked(true);
  sheet.resumePaint();
};

/**
 * Add an image to the quote
 * @param {*} spread
 * @param {*} quotationImgs
 * @param {*} imgs
 */
const InsertImages = (spread, quotationImgs, imgs) => {
  if (imgs && imgs.length) {
    const sheet = spread.getActiveSheet();
    sheet.suspendPaint();
    quotationImgs.forEach((item, index) => {
      ((item, index) => {
        if (!imgs[index]) return;
        const row = sheet.getRowCount();
        const sealH = item.height;
        const rowOffsetLeft = item.columnOffsetLeft;
        const bottomR = item.rows;// 距离底部行数
        let bottomH = sealH;// 距离底部高度

        let offsetLeft = 0;// 距离表格左侧距离
        let offsetTop = 0;// 距离顶部距离
        // 计算距离左侧距离
        for (let i = 0; i < rowOffsetLeft; i++) {
          offsetLeft = offsetLeft + sheet.getColumnWidth(i);
        }
        if (item.position == 'top') {
          // 计算距离底部距离
          for (let i = 0; i < bottomR; i++) {
            offsetTop = offsetTop + sheet.getRowHeight(i);
          }
        } else if (item.position == 'bottom') {
          // 计算距离顶部距离
          for (let i = 0; i < row; i++) {
            offsetTop = offsetTop + sheet.getRowHeight(i);
          }
          // 计算距离底部距离
          for (let i = row - bottomR; i < row; i++) {
            bottomH = bottomH + sheet.getRowHeight(i);
          }
          offsetTop = offsetTop - bottomH;
        } else {
          console.warn('报价单多Logo字段(position)配置出错');
        }
        let picture = sheet.pictures.get(item.name);
        if (picture) {
          picture.src(imgs[index].url);
          sheet.resumePaint();
          return;
        }
        picture = sheet.pictures.add(item.name, imgs[index].url, offsetLeft, offsetTop, item.width, item.height);
        picture.isLocked(true);
      })(item, index);
    });
    sheet.resumePaint();
  }
};

/**
 * Add an image to the quote
 * @param {*} spread
 * @param {*} template
 * @param {*} base64
 */
const InsertImage = (spread, template, base64) => {
  const sheet = spread.getActiveSheet();
  sheet.suspendPaint();
  const row = sheet.getRowCount();// 表格所有行
  const quotationImageName = template.cloudSheet.quotationImage.name;// 公司印章名称
  const quotationImageH = template.cloudSheet.quotationImage.height;// 公司印章高度
  const quotationImageW = template.cloudSheet.quotationImage.width;// 公司印章宽度
  const rowOffsetLeft = template.cloudSheet.quotationImage.rowOffsetLeft;// 公司印章距离表格左侧的列数
  const bottomR = template.cloudSheet.quotationImage.columntBottm;// 公司印章距离底部行数
  let offsetLeft = 0;// 公司印章距离表格左侧距离
  let offsetTop = 0;// 公司印章距离顶部距离
  let bottomH = quotationImageH;// 公司印章距离底部高度
  // 计算距离左侧距离
  for (let i = 0; i < rowOffsetLeft; i++) {
    offsetLeft = offsetLeft + sheet.getColumnWidth(i);
  }
  // 计算距离顶部距离
  for (let i = 0; i < row; i++) {
    offsetTop = offsetTop + sheet.getRowHeight(i);
  }
  // 计算距离底部距离
  for (let i = row - bottomR; i < row; i++) {
    bottomH = bottomH + sheet.getRowHeight(i);
  }

  offsetTop = offsetTop - bottomH;
  let picture = sheet.pictures.get(quotationImageName);
  if (picture) {
    sheet.pictures.remove(quotationImageName);
    picture = sheet.pictures.add(quotationImageName, base64, offsetLeft, offsetTop, quotationImageW, quotationImageH);
  } else {
    picture = sheet.pictures.add(quotationImageName, base64, offsetLeft, offsetTop, quotationImageW, quotationImageH);
  }
  picture.allowMove(false);
  picture.allowResize(true);
  picture.isLocked(true);
  sheet.resumePaint();
};

/**
 * Images are added asynchronously at the end of rendering
 * @param {*} spread
 * @param {*} template
 * @param {*} quotation
 */
const renderFinishedAddImage = (spread, template, quotation) => {
  const { quaLogos = [], seal = null } = template.cloudSheet;

  if (_.has(template, ['cloudSheet', 'quotationImage'])) {
    if (quotation.image) {
      let imgUrl = '';
      if (REGULAR.chineseCharacters.test(quotation.image)) {
        imgUrl = encodeURI(quotation.image);
      } else {
        imgUrl = quotation.image;
      }
      imgUrlToBase64(imgUrl, (base64) => {
        InsertImage(spread, template, base64);
      });
    }
  }
  if (quaLogos && quaLogos.length) {
    API.qtLogoImages().then(res => {
      const Res = new ResDatas({ res }).init();
      if (Res) {
        Res.forEach(item => {
          ((item) => {
            if (item.url) {
              imgUrlToBase64(item.url, (base64) => {
                item.url = base64;
                InsertImages(spread, quaLogos, Res);
              });
            }
          })(item);
        });
      }
    });
  }

  if (seal && quotation.seal) {
    imgUrlToBase64(quotation.seal, (base64) => {
      InsertSeal(spread, template, base64);
    });
  }
};

const getHeaderRowCount = () => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const { top, bottom } = template.cloudSheet;
  return top.rowCount + bottom.rowCount;
};

/**
 * Initialize the initData field value
 * @param {*} sheet
 * @param {*} field
 * @param {*} source
 * @param {*} path
 * @param {*} type
 */
const initDataSetValue = (sheet, field, source, path, type) => {
  const lastRow = sheet.getRowCount() - 1;

  let row;

  if (field.row === 0 || field.row) {
    row = field.row;
  } else {
    row = lastRow - Number(field.lastRow);
  }

  if (type === 'date') {
    if (path && _.has(source, path)) {
      const value = _.get(source, path);
      if (value) {
        sheet.setValue(row, field.column, getSystemDate(new Date(value), 'yyyy/MM/dd'));
      }
    }
  } else if (type === 'company') {
    const value = _.get(source, path);
    if (value === 0 || value) {
      sheet.setValue(row, field.column, value);
    }
  } else if (type === 'information') {
    if (_.has(source, path)) {
      const value = _.get(source, path);
      if (value === 0 || value) {
        sheet.setValue(row, field.column, value);
      }
    }
  }
};

/**
 * Gets the value of the calculated row
 * @param {*} fixedBindValueMap
 * @param {*} computedFieldMap
 * @returns
 */
const getComputedRowDefaultVal = (fixedBindValueMap, computedFieldMap) => {
  const managementExpense = fixedBindValueMap.managementExpense || computedFieldMap.managementExpense;
  const serviceCharge = fixedBindValueMap.serviceChargeFee || fixedBindValueMap.serviceCharge || computedFieldMap.serviceChargeFee || computedFieldMap.serviceCharge;
  const taxes = fixedBindValueMap.taxes || computedFieldMap.taxes;

  return {
    managementExpense: Number(managementExpense),
    serviceCharge: Number(serviceCharge),
    taxes: Number(taxes)
  };
};

/**
 * Excludes service charge and management fee
 * columnTotalSum + freight + projectCost + other
 * @param {*} totalBinds
 * @param {*} columnTotalSum
 * @param {*} fixedBindValueMap
 * @returns
 */
const totalBeforeAssignment = (fieldName = 'totalBeforeTax', totalBinds = {}, columnTotalSum, fixedBindValueMap) => {
  const { freight = 0, projectCost = 0 } = fixedBindValueMap;
  const rowNames = rowComputedFieldSort(totalBinds);

  const currentField = rowNames.findIndex((name) => name === fieldName);
  const freightIndex = rowNames.findIndex((name) => name === 'freight');
  const projectCostIndex = rowNames.findIndex((name) => name === 'projectCost');

  let freightSum = 0;
  let projectCostSum = 0;
  if (currentField > freightIndex) {
    freightSum = freight;
  }
  if (currentField > projectCostIndex) {
    projectCostSum = projectCost;
  }

  return new Decimal(columnTotalSum).plus(new Decimal(freightSum)).plus(new Decimal(projectCostSum)).toNumber();
};

/**
 * Dynamically interpolated summary fields
 * @param {*} fieldName
 * @param {*} totalBinds
 * @param {*} columnTotalSum
 * @param {*} fixedBindValueMap
 * @param {*} computedFieldMap
 * @returns
 */
const totalBeforeTaxAssignment = (fieldName, totalBinds = {}, columnTotalSum, fixedBindValueMap, computedFieldMap) => {
  // eslint-disable-next-line no-unused-vars
  const { managementExpense = 0, serviceCharge = 0, taxes = 0 } = getComputedRowDefaultVal(fixedBindValueMap, computedFieldMap);
  const totalBefore = totalBeforeAssignment(fieldName, totalBinds, columnTotalSum, fixedBindValueMap);
  const sums = [];

  const sumMap = {
    managementExpense: 0,
    serviceCharge: 0,
    taxes: 0
  };

  const rowNames = rowComputedFieldSort(totalBinds);
  const currentField = rowNames.findIndex((name) => name === fieldName);
  const managementExpenseIndex = rowNames.findIndex((name) => name === 'managementExpense');
  const serviceChargeIndex = rowNames.findIndex((name) => name === 'serviceCharge');
  // const taxesIndex = rowNames.findIndex((name) => name === 'taxes');

  if (currentField > managementExpenseIndex) {
    sumMap.managementExpense = managementExpense || 0;
  }
  if (currentField > serviceChargeIndex) {
    sumMap.serviceCharge = serviceCharge || 0;
  }

  // if (currentField > taxesIndex) {
  //   sumMap.taxes = taxes || 0;
  // }

  for (const key in sumMap) {
    if (Object.hasOwnProperty.call(sumMap, key)) {
      sums.push(new Decimal(sumMap[key]));
    }
  }

  return Decimal.add(new Decimal(totalBefore), new Decimal(sums[0]), new Decimal(sums[1]), new Decimal(sums[2])).toNumber();
};
/**
 * serviceChargeSum + managementExpenseSum + taxesSum
 * @param {*} fieldName
 * @param {*} totalBinds
 * @param {*} fixedBindValueMap
 * @param {*} computedFieldMap
 * @returns
 */
const serviceAndManageAndTaxSum = (fieldName, totalBinds = {}, fixedBindValueMap, computedFieldMap) => {
  // eslint-disable-next-line no-unused-vars
  const { managementExpense = 0, serviceCharge = 0, taxes = 0 } = getComputedRowDefaultVal(fixedBindValueMap, computedFieldMap);
  const sums = [];

  const sumMap = {
    managementExpense: 0,
    serviceCharge: 0,
    taxes: 0
  };

  const rowNames = rowComputedFieldSort(totalBinds);
  const currentField = rowNames.findIndex((name) => name === fieldName);
  const managementExpenseIndex = rowNames.findIndex((name) => name === 'managementExpense');
  const serviceChargeIndex = rowNames.findIndex((name) => name === 'serviceCharge');
  // const taxesIndex = rowNames.findIndex((name) => name === 'taxes');

  if (currentField > managementExpenseIndex) {
    sumMap.managementExpense = managementExpense || 0;
  }
  if (currentField > serviceChargeIndex) {
    sumMap.serviceCharge = serviceCharge || 0;
  }

  // if (currentField > taxesIndex) {
  //   sumMap.taxes = taxes || 0;
  // }

  for (const key in sumMap) {
    if (Object.hasOwnProperty.call(sumMap, key)) {
      sums.push(new Decimal(sumMap[key]));
    }
  }
  return Decimal.add(new Decimal(sums[0]), new Decimal(sums[1]), new Decimal(sums[2])).toNumber();
};

/**
 * The sum of the totals of each column
 * @param {*} columnTotal
 * @returns
 */
const totalAfterTaxsColumnAssignment = (columnTotal) => {
  const mapsum = {};
  columnTotal.forEach(tableColumnComputedItem => {
    for (const key in tableColumnComputedItem) {
      if (Object.hasOwnProperty.call(tableColumnComputedItem, key)) {
        if (Object.keys(mapsum).includes(key)) {
          mapsum[tableColumnComputedItem[key].columnHeader] = Decimal.add(new Decimal(mapsum[key]), new Decimal(tableColumnComputedItem[key].sum)).toNumber();
        } else {
          mapsum[tableColumnComputedItem[key].columnHeader] = tableColumnComputedItem[key].sum;
        }
      }
    }
  });
  return mapsum;
};

/**
 * totalAfterTaxs assignment
 * @param {*} columnTotal
 * @param {*} field
 * @param {*} columnSum
 * @param {*} totalBinds
 * @param {*} columnTotalSum
 * @param {*} fixedBindValueMap
 * @param {*} setUnitCb
 * @param {*} cb
 */
const totalAfterTaxsAssignment = (fieldName, columnTotal, field, columnSum, totalBinds, columnTotalSum, fixedBindValueMap, setUnitCb, cb) => {
  let totalHeader = '';
  for (let index = 0; index < columnTotal.length; index++) {
    if (index === 0) {
      if (Object.keys(columnTotal[index]).includes('total')) {
        totalHeader = columnTotal[index].total.columnHeader;
        setUnitCb && setUnitCb();
      } else {
        console.error('缺少列总价字段(total)对应的行小计字段(totalBeforeTax)');
      }
    }
  }

  if (Object.keys(columnSum).includes(field.columnHeader)) {
    if (!field.bindPath) {
      if (field.columnHeader === totalHeader) {
        cb(field.column, totalBeforeAssignment(fieldName, totalBinds, columnTotalSum, fixedBindValueMap));
      } else if (fieldName === 'totalBeforeTax') {
        cb(field.column, totalBeforeAssignment(fieldName, totalBinds, columnTotalSum, fixedBindValueMap));
      } else {
        cb(field.column, columnSum[field.columnHeader]);
      }
    }
  } else {
    console.warn('column Total 与统计埋点字段未在同一列');
  }
};
/**
 * managementExpense assignment
 * @param {*} columnTotal
 * @param {*} fixedBindValueMap
 * @param {*} computedFieldMap
 * @returns
 */
const managementExpenseAssignment = (columnTotal, fixedBindValueMap, computedFieldMap) => {
  const managementFee = fixedBindValueMap.managementFee || computedFieldMap.managementFee;

  let sum = 0;
  columnTotal.forEach(tableColumnComputedItem => {
    const columns = Object.keys(tableColumnComputedItem);
    if (columns.includes('total')) {
      sum = new Decimal(sum).plus(new Decimal(tableColumnComputedItem.total.sum)).toNumber();
    } else {
      console.warn('管理费缺少列字段(total)合计,造成无法计算!');
    }
  });
  return new Decimal(sum).times(new Decimal(managementFee)).dividedBy(new Decimal(100)).toNumber();
};
/**
 * serviceCharge assignment
 * @param {*} totalBinds
 * @param {*} columnTotalSum
 * @param {*} fixedBindValueMap
 * @param {*} computedFieldMap
 * @returns
 */
const serviceChargeAssignment = (fieldName, totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap) => {
  const { rate } = fixedBindValueMap;
  const totalBeforeTax = totalBeforeTaxAssignment(fieldName, totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap);

  return new Decimal(totalBeforeTax).times(new Decimal(rate)).dividedBy(new Decimal(100)).toNumber();
};
/**
 * serviceChargeFee assignment
 * @param {*} totalBinds
 * @param {*} columnTotalSum
 * @param {*} fixedBindValueMap
 * @param {*} computedFieldMap
 * @returns
 */
const serviceChargeFeeAssignment = (fieldName, totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap) => {
  const totalBeforeTax = totalBeforeTaxAssignment(fieldName, totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap);
  const { serviceCharge } = getComputedRowDefaultVal(fixedBindValueMap, computedFieldMap);

  return new Decimal(totalBeforeTax).times(new Decimal(serviceCharge)).dividedBy(new Decimal(100)).toNumber();
};
/**
 * totalServiceCharge assignment
 * @param {*} totalBinds
 * @param {*} columnTotalSum
 * @param {*} fixedBindValueMap
 * @param {*} computedFieldMap
 */
const totalServiceChargeAssignment = (totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap) => {
  const totalBeforeTax = totalBeforeAssignment('totalBeforeTax', totalBinds, columnTotalSum, fixedBindValueMap);
  const { serviceCharge } = getComputedRowDefaultVal(fixedBindValueMap, computedFieldMap);

  return new Decimal(totalBeforeTax).plus(new Decimal(serviceCharge)).toNumber();
};
/**
 * addTaxRateBefore assignment
 * @param {*} totalBinds
 * @param {*} columnTotalSum
 * @param {*} fixedBindValueMap
 * @param {*} computedFieldMap
 * @returns
 */
const addTaxRateBeforeAssignment = (fieldName, totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap) => {
  const totalBeforeTax = totalBeforeTaxAssignment(fieldName, totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap);
  return totalBeforeTax;
};
/**
 * taxes assignment
 * @param {*} totalBinds
 * @param {*} columnTotalSum
 * @param {*} fixedBindValueMap
 * @param {*} computedFieldMap
 * @returns
 */
const taxesAssignment = (fieldName, totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap) => {
  const { tax = null, totalServiceCharge = null } = fixedBindValueMap;
  const totalBefore = totalBeforeAssignment(fieldName, totalBinds, columnTotalSum, fixedBindValueMap);
  let includeServiceSum = totalServiceCharge || computedFieldMap.totalServiceCharge;
  if (!includeServiceSum) {
    const smtSum = serviceAndManageAndTaxSum(fieldName, totalBinds, fixedBindValueMap, computedFieldMap);
    let smtSum_ = 0;
    let totalBefore_ = 0;
    if (smtSum) {
      smtSum_ = smtSum;
    }
    if (totalBefore) {
      totalBefore_ = totalBefore;
    }
    includeServiceSum = new Decimal(new Decimal(totalBefore_)).plus(new Decimal(smtSum_)).toNumber();
  }
  return new Decimal(includeServiceSum).times(new Decimal(tax)).dividedBy(new Decimal(100)).toNumber();
};

/**
 *rowComputed set value
 * @param {*} sheet
 * @param {*} field
 * @param {*} fixedBindValueMap
 * @param {*} computedFieldMap
 * @param {*} columnTotal
 * @param {*} rowIndex
 * @param {*} key
 * @param {*} cb
 * @param {*} totalBinds
 * @param {*} columnSum
 * @param {*} columnTotalSum
 */
const rowComputedField = (sheet, field, fixedBindValueMap, computedFieldMap, columnTotal, rowIndex, key, cb, totalBinds, columnSum, columnTotalSum) => {
  let fieldName = key;
  if (!regChineseCharacter.test(field.name)) {
    fieldName = field.name;
  }
  const fieldInfo = getFormulaFieldRowCol(field);
  console.log(fieldName, 'fieldName 无bindpath字段');

  if (ASSOCIATED_FIELDS_FORMULA_MAP[fieldName]) {
    setCellFormatter(sheet, rowIndex + fieldInfo.row, fieldInfo.column);
    if (fieldName === 'managementExpense') {
      const fieldVal = managementExpenseAssignment(columnTotal, fixedBindValueMap, computedFieldMap);
      sheet.setValue(rowIndex + fieldInfo.row, fieldInfo.column, fieldVal);
      fieldVal && cb(fieldVal);
    } else if (GenerateFieldsRow().includes(fieldName)) {
      totalAfterTaxsAssignment(fieldName, columnTotal, fieldInfo, columnSum, totalBinds, columnTotalSum, fixedBindValueMap, () => {
        setCellFormatter(sheet, rowIndex + fieldInfo.row, fieldInfo.column);
      }, (column, fieldVal) => {
        sheet.setValue(rowIndex + fieldInfo.row, column, fieldVal);
        fieldVal && cb(fieldVal);
      });
    } else if (fieldName === 'serviceCharge') {
      const fieldVal = serviceChargeAssignment(fieldName, totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap);
      sheet.setValue(rowIndex + fieldInfo.row, fieldInfo.column, fieldVal);
      fieldVal && cb(fieldVal);
    } else if (fieldName === 'totalServiceCharge') {
      const fieldVal = totalServiceChargeAssignment(totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap);
      sheet.setValue(rowIndex + fieldInfo.row, fieldInfo.column, fieldVal);
      fieldVal && cb(fieldVal);
    } else if (fieldName === 'addTaxRateBefore') {
      const fieldVal = addTaxRateBeforeAssignment(fieldName, totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap);
      sheet.setValue(rowIndex + fieldInfo.row, fieldInfo.column, fieldVal);
      fieldVal && cb(fieldVal);
    } else if (fieldName === 'taxes') {
      const fieldVal = taxesAssignment(fieldName, totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap);
      sheet.setValue(rowIndex + fieldInfo.row, fieldInfo.column, fieldVal);
      fieldVal && cb(fieldVal);
    } else if (fieldName === 'serviceChargeFee') {
      const fieldVal = serviceChargeFeeAssignment(fieldName, totalBinds, columnTotalSum, fixedBindValueMap, computedFieldMap);
      sheet.setValue(rowIndex + fieldInfo.row, fieldInfo.column, fieldVal);
      fieldVal && cb(fieldVal, 'serviceCharge');
    } else {
      console.warn('在ASSOCIATED_FIELDS_FORMULA_MAP定义,但未存在过相关逻辑的字段', fieldName);
    }
    sheet.autoFitColumn(fieldInfo.column);
  } else {
    console.warn('模板的总计block：识别出未在ASSOCIATED_FIELDS_FORMULA_MAP定义的字段', fieldName);
  }
};

/**
 * Total set value
 * @param {*} sheet
 * @param {*} totalField
 * @param {*} row
 * @param {*} columnTotal
 * @param {*} position
 * @param {*} columnComputed
 */
// eslint-disable-next-line no-unused-vars
const setTotalRowValue = (sheet, totalField, row, columnTotal, position, columnComputed) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const { truckage = null, cloudSheet: { total } } = store.getters['quotationModule/GetterQuotationWorkBook'];
  const totalBinds = total ? total[templateTotalMap(total.select)].bindPath : null;

  const columnTotalSum = plusColumnTotalSum(columnTotal);
  const columnSum = totalAfterTaxsColumnAssignment(columnTotal);

  const fixedBindValueMap = {};
  const computedFieldMap = {};
  const fixedBindCellMap = {};

  let columnHeader = null;
  if (position) {
    for (const key in position) {
      if (Object.hasOwnProperty.call(position, key)) {
        columnHeader = position[key].columnHeader;
      }
    }
  }

  // Get a fixed value
  for (const key in totalField.bindPath) {
    if (Object.hasOwnProperty.call(totalField.bindPath, key)) {
      const rows = _.cloneDeep(totalField.bindPath[key]);
      if (rows.bindPath) {
        if (columnHeader) {
          if (rows.columnHeader !== columnHeader) {
            const path = rows.bindPath.split('.');
            fixedBindValueMap[key] = Number(_.get(quotation, path));
          }
        } else {
          if (!Object.keys(DESCRIPTION_MAP).includes(rows.bindPath)) {
            const path = rows.bindPath.split('.');
            fixedBindValueMap[key] = Number(_.get(quotation, path));
          }
        }
      }
    }
  }

  if (!truckage) {
    // Set a fixed value
    for (const key in totalField.bindPath) {
      if (Object.hasOwnProperty.call(totalField.bindPath, key)) {
        const rows = _.cloneDeep(totalField.bindPath[key]);

        console.log(rows, 'rows');

        if (rows.bindPath) {
          // Binding value
          if (columnHeader) {
            if (rows.columnHeader === columnHeader) {
              fixedBindCellMap[key] = {
                row: row + rows.row,
                column: rows.column
              };

              setCellFormatter(sheet, row + rows.row, rows.column);
              sheet.setBindingPath(row + rows.row, rows.column, rows.bindPath);
              sheet.autoFitColumn(rows.column);
            }
          } else {
            if (Object.keys(DESCRIPTION_MAP).includes(rows.bindPath)) {
              mixedDescriptionFields(sheet, quotation, row, rows);
            } else {
              setCellFormatter(sheet, row + rows.row, rows.column);

              fixedBindCellMap[key] = {
                row: row + rows.row,
                column: rows.column
              };

              sheet.setBindingPath(row + rows.row, rows.column, rows.bindPath);
              sheet.autoFitColumn(rows.column);
            }
          }
        }
      }
    }

    // Dynamic fields
    for (const key in totalField.bindPath) {
      if (Object.hasOwnProperty.call(totalField.bindPath, key)) {
        const rows = _.cloneDeep(totalField.bindPath[key]);
        if (!rows.bindPath) {
          rowComputedField(sheet, rows, fixedBindValueMap, computedFieldMap, columnTotal, row, key, (value, name) => {
            let fieldName = key;
            if (!regChineseCharacter.test(rows.name)) {
              fieldName = rows.name;
            }
            if (name) {
              computedFieldMap[name] = value;
            } else {
              computedFieldMap[fieldName] = value;
            }
          }, totalBinds, columnSum, columnTotalSum);
        }
      }
    }

    console.log(fixedBindCellMap, 'fixedBindCellMap');

    for (const key in totalField.bindPath) {
      if (Object.hasOwnProperty.call(totalField.bindPath, key)) {
        const rows = _.cloneDeep(totalField.bindPath[key]);
        if (rows.bindPath) {
          const path = rows.bindPath.split('.');
          if (_.has(quotation, path)) {
            if (_.get(quotation, path) !== 0 && !_.get(quotation, path)) {
              let fieldName = key;
              let value = '';
              if (!regChineseCharacter.test(rows.name)) {
                fieldName = rows.name;
              }

              if (fieldName === 'totalAfterTax') {
                const total = computedFieldMap.totalServiceCharge || computedFieldMap.totalBeforeTax;
                value = new Decimal(total).plus(new Decimal(computedFieldMap.taxes)).toNumber();
              } else if (GenerateFieldsRow().includes(fieldName)) {
                totalAfterTaxsAssignment(fieldName, columnTotal, rows, columnSum, totalBinds, columnTotalSum, fixedBindValueMap, null, (column, fieldVal) => {
                  value = fieldVal;
                });
              }

              if (value === 0 || value) {
                computedFieldMap[fieldName] = value;
              }

              setCellFormatter(sheet, row + rows.row, rows.column);
              sheet.setValue(row + rows.row, rows.column, value);
              sheet.autoFitColumn(rows.column);
            }
          }
        }
      }
    }
  } else {
    const TruckageIdentifier = new IdentifierTemplate(sheet, 'truckage');
    TruckageIdentifier.truckageFreight(totalField, row, fixedBindValueMap);
  }

};

/**
 * A subtotal of the classification layer
 * @param {*} spread
 * @param {*} quotation
 * @param {*} columnTotal
 */
const RenderHeaderClass = (spread, quotation, columnTotal) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const { top, center: { equipment }, mixTopTotal: { initTotal } } = template.cloudSheet;
  const rows = resourceViews.map((item, index) => { return top.mixCount + index; });

  const sheet = spread.getActiveSheet();
  sheet.suspendPaint();

  sheet.addRows(top.mixCount, resourceViews.length);

  const { style } = GeneratorStyle('classNameHeader', { fontWeight: 'normal', textIndent: 1 });
  const classNameHeaderCenter = GeneratorStyle('classNameHeaderCenter', { fontWeight: 'normal', textIndent: 0, hAlign: 1, borderTop: GeneratorLineBorder() });

  let columnHeader = null;
  for (const key in initTotal.bindPath) {
    if (Object.hasOwnProperty.call(initTotal.bindPath, key)) {
      columnHeader = initTotal.bindPath[key].column;
    }
  }
  mergeRow(sheet, rows, 1, 1, columnHeader - 1);

  sheet.getRange(top.mixCount, 0, resourceViews.length, 1).setStyle(classNameHeaderCenter.style);
  for (let index = 0; index < rows.length; index++) {
    sheet.setStyle(rows[index], -1, style, GC.Spread.Sheets.SheetArea.viewport);
    sheet.getRange(rows[index], 0, 1, equipment.columnCount).setBorder(GeneratorLineBorder(), { all: true });
    sheet.setValue(rows[index], 0, numberToColumn(index + 1));
    sheet.setValue(rows[index], 1, resourceViews[index].parentTypeName || resourceViews[index].name);

    const column = columnTotal[index];
    for (const key in column) {
      if (Object.hasOwnProperty.call(column, key)) {
        setCellFormatter(sheet, rows[index], column[key].column);
        sheet.setValue(rows[index], column[key].column, column[key].sum);
        sheet.autoFitColumn(column[key].column);
        sheet.getCell(rows[index], column[key].column).setStyle(classNameHeaderCenter.style);
      }
    }
  }
  if (equipment.height === 0 || equipment.height) {
    for (let index = 0; index < rows.length; index++) {
      sheet.setRowHeight(rows[index], equipment.height);
    }
  }

  sheet.resumePaint();
};
/**
 * Tota rendering
 * @param {*} spread
 * @param {*} quotation
 * @param {*} columnTotal
 */
const RenderHeaderTotal = (spread, quotation, columnTotal) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const { top, total, mixTopTotal } = template.cloudSheet;
  const initTotal = mixTopTotal.initTotal;
  const { resourceViews } = quotation.conferenceHall;

  const totalRowIndex = top.mixCount + resourceViews.length;
  const combined = TOTAL_COMBINED_MAP[total.select];

  const sheet = spread.getActiveSheet();
  sheet.suspendPaint();

  sheet.deleteRows(totalRowIndex, initTotal.rowCount);
  sheet.addRows(totalRowIndex, mixTopTotal[combined].rowCount);

  if (mixTopTotal[combined].spans) {
    mergeSpan(sheet, mixTopTotal[combined].spans, totalRowIndex);
  }
  setCellStyle(spread, mixTopTotal[combined], totalRowIndex, true);
  setTotalRowHeight(sheet, total, mixTopTotal[combined], totalRowIndex);
  setTotalRowValue(sheet, mixTopTotal[combined], totalRowIndex, columnTotal, initTotal.bindPath, null);

  sheet.resumePaint();
};

/**
 * 判断模板是否存在总量多种计算方式
 * @param {*} spread
 * @param {*} template
 */
export const LogicalTotalCalculationType = (spread) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  if (_.has(template, ['cloudSheet', 'center', 'equipment', 'bindPath', 'areaTotal'])) {
    const sheet = spread.getActiveSheet();
    const allRow = sheet.getRowCount();// 表格所有行
    const allNum = allRow - template.cloudSheet.bottom.rowCount - 1;
    for (let index = template.cloudSheet.top.rowCount; index < allNum; index++) {
      ((index) => {
        setDefaultCalMethod(sheet, index);
      })(index);
    }
  }
};

/**
 * Header Initialize the top assignment
 * @param {*} spread
 * @param {*} template
 * @param {*} quotation
 */
export const InitBindValueTop = (spread, template, quotation) => {
  const sheet = spread.getActiveSheet();
  const initDataPath = ['cloudSheet', 'initData'];

  sheet.suspendPaint();
  if (_.has(template, initDataPath)) {
    const initData = _.get(template, initDataPath);

    // 初始化个人信息
    if (initData.userName || initData.userTel || initData.belongsEmail) {
      const userInfo = GetUserInfoDetail();

      if (initData.userName && userInfo) {
        initDataSetValue(sheet, initData.userName, userInfo, ['name'], 'information');
      }
      if (initData.userTel && userInfo) {
        initDataSetValue(sheet, initData.userTel, userInfo, ['phone'], 'information');
      }
      if (initData.belongsEmail && userInfo) {
        initDataSetValue(sheet, initData.belongsEmail, userInfo, ['email'], 'information');
      }
    }

    // 初始化公司信息
    if (initData) {
      const companyInfo = GetUserCompany();

      if (initData.userCompanyName && companyInfo && companyInfo.name) {
        initDataSetValue(sheet, initData.userCompanyName, companyInfo, ['name'], 'company');
      }
      if (initData.userCompanyEnglishName && companyInfo && companyInfo.companyEnglishName) {
        initDataSetValue(sheet, initData.userCompanyEnglishName, companyInfo, ['companyEnglishName'], 'company');
      }
      if (initData.userCompanyFax && companyInfo && companyInfo.fax) {
        initDataSetValue(sheet, initData.userCompanyFax, companyInfo, ['fax'], 'company');
      }
      if (initData.userCompanyWebsite && companyInfo && companyInfo.website) {
        initDataSetValue(sheet, initData.userCompanyWebsite, companyInfo, ['website'], 'company');
      }
      if (initData.userCompanyLocation && companyInfo && companyInfo.address) {
        initDataSetValue(sheet, initData.userCompanyLocation, companyInfo, ['address'], 'company');
      }
      if (initData.userCompanyTel && companyInfo && companyInfo.tel) {
        initDataSetValue(sheet, initData.userCompanyTel, companyInfo, ['tel'], 'company');
      }

      if (template.cloudSheet.logo && quotation.logo) {
        imgUrlToBase64(quotation.logo, (base64) => {
          InsertLogo(spread, template, base64);
        });
      } else {
        let userBindCompany = sessionStorage.getItem('userBindCompany');
        if (userBindCompany) {
          userBindCompany = JSON.parse(userBindCompany);
          if (userBindCompany.logoURL) {
            imgUrlToBase64(userBindCompany.logoURL, (base64) => {
              InsertLogo(spread, template, base64);
            });
          }
        }
      }
    }

    // 初始化时间
    if (initData) {
      if (initData.approachDate) {
        initDataSetValue(sheet, initData.approachDate, quotation, ['conferenceHall', 'approachDate'], 'date');
      }
      if (initData.startDate) {
        initDataSetValue(sheet, initData.startDate, quotation, ['conferenceHall', 'startDate'], 'date');
      }
      if (initData.fieldWithdrawalDate) {
        initDataSetValue(sheet, initData.fieldWithdrawalDate, quotation, ['conferenceHall', 'fieldWithdrawalDate'], 'date');
      }
      if (initData.createTime) {
        initDataSetValue(sheet, initData.createTime, quotation, ['createTime'], 'date');
      }
    }
  }
  sheet.resumePaint();
};

/**
 * Fields are bound and assigned
 * @param {*} spread
 * @param {*} template
 * @param {*} bindingPath
 */
export const FieldBindPath = (spread, template, bindingPath) => {
  const sheet = spread.getActiveSheet();
  if (_.has(template, bindingPath)) {
    const bindPath = _.get(template, bindingPath);
    for (const key in bindPath) {
      if (Object.hasOwnProperty.call(bindPath, key)) {
        const ele = bindPath[key];
        if (ele.row === 0 || ele.row) {
          sheet.setBindingPath(ele.row, ele.column, ele.bindPath);
        } else {
          const lastRow = getHeaderRowCount() - 1;
          const row = lastRow - Number(ele.lastRow);
          sheet.setBindingPath(row, ele.column, ele.bindPath);
        }
      }
    }
  }
};

/**
 * Initialize Total and assign a value
 * @param {*} spread
 */
export const InitTotal = (spread) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const sheet = spread.getActiveSheet();
  const { total } = template.cloudSheet;

  if (showTotal()) {
    const tableStartRowIndex = PubGetTableStartRowIndex();
    const tableRowCount = PubGetTableRowCount();
    const templateTotal = total[templateTotalMap(total.select)];
    const totalStartRowIndex = tableStartRowIndex + tableRowCount;

    sheet.suspendPaint();
    sheet.addRows(totalStartRowIndex, templateTotal.rowCount);
    mergeSpan(sheet, templateTotal.spans, totalStartRowIndex);
    setCellStyle(spread, templateTotal, totalStartRowIndex, true);
    setTotalRowHeight(sheet, total, templateTotal, totalStartRowIndex);
    sheet.resumePaint();
  }
};
/**
 * Tota rendering
 * @param {*} spread
 * @param {*} columnTotal
 */
const RenderTotal = (spread, columnTotal, columnComputed) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const { total, bottom } = template.cloudSheet;
  if (showTotal()) {
    const sheet = spread.getActiveSheet();
    sheet.suspendPaint();

    // index
    const bottomRowCount = bottom.rowCount;
    const totalRowIndex = sheet.getRowCount() - bottomRowCount;

    const Total = total[templateTotalMap(total.select)];
    sheet.addRows(totalRowIndex, Total.rowCount);

    console.log(Total, 'Total');

    mergeSpan(sheet, Total.spans, totalRowIndex);
    setCellStyle(spread, Total, totalRowIndex, true);
    setTotalRowHeight(sheet, total, Total, totalRowIndex);
    setTotalRowValue(sheet, Total, totalRowIndex, columnTotal, null, columnComputed);

    sheet.resumePaint();

    if (template.truckage) {
      const TruckageIdentifier = new IdentifierTemplate(sheet, 'truckage');
      TruckageIdentifier.truckageRenderTotal(quotation);
    }
  }
};

/**
 * Product Rendering
 * @param {*} spread
 */
const renderSheet = (spread) => {
  const sheet = spread.getActiveSheet();
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
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

            const { style } = GeneratorStyle('className', { textIndent: 1 });
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
        setCellStyle(spread, type, rowClassIndex, true);
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
      const startRow = rowClassIndex + classRow + tableHeaderRow + headerRow + index;

      const { image = null } = template.cloudSheet;

      mergeSpan(sheet, equipment.spans, startRow);
      setRowStyle(sheet, equipment, startRow, image);
      columnComputedValue(sheet, equipment, startRow, computedColumnFormula);
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
      const rowClassTotal = rowClassIndex + classRow + tableHeaderRow + PubGetTableRowCount(i) + headerRow;
      sheet.addRows(rowClassTotal, subTotal, GC.Spread.Sheets.SheetArea.viewport);
      if (total) {
        mergeSpan(sheet, total.spans, rowClassTotal);
        setCellStyle(spread, total, rowClassTotal, true);
        PubSetCellHeight(sheet, total, rowClassTotal);
        for (const key in total.bindPath) {
          if (Object.hasOwnProperty.call(total.bindPath, key)) {
            const field = total.bindPath[key];
            if (field.bindPath) {
              setCellFormatter(sheet, rowClassTotal + field.row, field.column);
              sheet.setBindingPath(rowClassTotal + field.row, field.column, field.bindPath);
              sheet.autoFitColumn(field.column);
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
    renderSheetImage(spread, tableStartRowIndex, false, true, true);
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
    // TODO header class 有新的标识符
    RenderHeaderClass(spread, quotation, columnTotal);
    // TODO 顶部组合未同步修改
    RenderHeaderTotal(spread, quotation, columnTotal);
  } else {
    RenderTotal(spread, columnTotal, columnComputed);
  }
};

/**
 * Render styles by template type
 * @param {*} spread
 */
export const Render = (spread) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = _.cloneDeep(store.getters['quotationModule/GetterQuotationInit']);
  console.log(quotation, 'quotation');
  console.log(template, 'template');
  renderSheet(spread);
  renderFinishedAddImage(spread, template, quotation);
  // setLastColumnWidth(spread, template);
  translateSheet(spread);
};

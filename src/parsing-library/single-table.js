/*
 * @Author: Marlon
 * @Date: 2024-04-07 18:29:42
 * @Description: 单表 - 逻辑处理
 */
import * as GC from '@grapecity/spread-sheets';
import _ from 'lodash';
import store from 'store';
import { getSystemDate, isNumber, regChineseCharacter, GetUserInfoDetail, GetUserCompany, imgUrlToBase64 } from '../utils/index';

import { CreateTable, SetDataSource } from '../common/sheetWorkBook';
import { GeneratorCellStyle, GeneratorLineBorder } from '../common/generator';
import { TOTAL_COMBINED_MAP, DESCRIPTION_MAP, REGULAR } from '../common/constant';

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
  getPaths
} from '../common/parsing-template';
import {
  templateTotalMap, mergeSpan, setCellStyle, setTotalRowHeight, PubGetTableStartRowIndex,
  PubGetTableRowCount,
  classificationAlgorithms,
  columnsTotal,
  mixedDescriptionFields,
  tableHeader,
  columnComputedValue,
  setCellFormatter,
  SetComputedSubTotal,
  renderSheetImage,
  translateSheet,
  initShowCostPrice,
  columnTotalSumFormula,
  GetColumnComputedTotal,
  clearTotalNoData,
  rowComputedField,
  sumAmountFormula,
} from '../common/single-table';

import {
  PubGetResourceViews,
  // eslint-disable-next-line no-unused-vars
  setLastColumnWidth
} from './public';

const NzhCN = require('nzh/cn');

let SumAmount = null;

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
const getBottomStartRowIndex = (spread, GetterQuotationWorkBook, GetterQuotationInit) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = GetterQuotationInit || store.getters['quotationModule/GetterQuotationInit'];
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
  if (showTotal(template)) {
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
const InsertImages = (spread, quotationImgs) => {
  const sheet = spread.getActiveSheet();
  sheet.suspendPaint();
  quotationImgs.forEach((item, index) => {
    ((item) => {
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
        picture.src(item.url);
        sheet.resumePaint();
        return;
      }
      picture = sheet.pictures.add(item.name, item.url, offsetLeft, offsetTop, item.width, item.height);
      picture.isLocked(true);
    })(item, index);
  });
  sheet.resumePaint();
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
    quaLogos.forEach(item => {
      ((item) => {
        if (item.url) {
          imgUrlToBase64(item.url, (base64) => {
            item.url = base64;
            InsertImages(spread, quaLogos, quaLogos);
          });
        }
      })(item);
    });
  }

  if (seal && quotation.seal) {
    imgUrlToBase64(quotation.seal, (base64) => {
      InsertSeal(spread, template, base64);
    });
  }
};

const getHeaderRowCount = (GetterQuotationWorkBook) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
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
 * Set dynamic field value for total
 * @param {*} sheet 
 * @param {*} totalField 
 * @param {*} row 
 * @param {*} totalBinds 
 * @param {*} template 
 * @param {*} GetterQuotationInit 
 */
const setTotalRowValue = (sheet, totalField, row, totalBinds, template, GetterQuotationInit = null) => {
  console.log(totalField, row, totalBinds);

  const quotation = GetterQuotationInit || store.getters['quotationModule/GetterQuotationInit'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const columnTotal = GetColumnComputedTotal(sheet, template, quotation);
  const columnTotalSum = columnTotalSumFormula(columnTotal);

  const fixedBindValueMap = {};
  const fixedBindCellMap = {};
  // const fixedBindKeys = Object.keys(totalField.bindPath);

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
            // TODO 最终价以优惠价为主
            const fieldFormula = sumAmountFormula(key, fixedBindCellMap, columnTotalSum);

            console.log(fieldFormula, 'sumAmount');

            sheet.setFormula(row + rows.row, rows.column, fieldFormula);
            const val = sheet.getValue(row + rows.row, rows.column);
            if (val === 0 || val) {
              SumAmount = val;
            }
          }
        }
      }
    } else {
      const TruckageIdentifier = new IdentifierTemplate(sheet, 'truckage');
      TruckageIdentifier.truckageFreight(totalField, row, fixedBindValueMap);
    }

    if (!template.truckage) {
      updateUpperCase(sheet, row, totalField, quotation);
    }
  } else {
    clearTotalNoData(sheet, row, totalField, fixedBindValueMap, quotation, template, 'parsing');
  }

  console.log(fixedBindValueMap, 'fixedBindValueMap');
  console.log(fixedBindCellMap, 'fixedBindCellMap');
};

/**
 * Update the uppercase value
 * @param {*} sheet 
 * @param {*} row 
 * @param {*} totalField 
 * @param {*} quotation 
 */
const updateUpperCase = (sheet, row, totalField, quotation) => {
  for (const key in totalField.bindPath) {
    if (Object.prototype.hasOwnProperty.call(totalField.bindPath, key)) {
      const rows = totalField.bindPath[key];
      if (rows.bindPath && rows.bindPath === 'DXzje') {
        const sumAmount = SumAmount || _.get(quotation, rows.bindPath)
        sheet.setValue(row + rows.row, rows.column, NzhCN.encodeB(sumAmount));
      }
    }
  }
}


/**
 * A subtotal of the classification layer
 * @param {*} spread
 * @param {*} quotation
 * @param {*} columnTotal
 */
const RenderHeaderClass = (spread, quotation, columnTotal, GetterQuotationWorkBook = null) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const { top, center: { equipment }, mixTopTotal: { initTotal } } = template.cloudSheet;
  const rows = resourceViews.map((item, index) => { return top.mixCount + index; });

  const sheet = spread.getActiveSheet();
  sheet.suspendPaint();

  sheet.addRows(top.mixCount, resourceViews.length);

  const { style } = GeneratorCellStyle('classNameHeader', { fontWeight: 'normal', textIndent: 1 });
  const classNameHeaderCenter = GeneratorCellStyle('classNameHeaderCenter', { fontWeight: 'normal', textIndent: 0, hAlign: 1, borderTop: GeneratorLineBorder() });

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
        setCellFormatter(sheet, rows[index], column[key].column, quotation);
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
// eslint-disable-next-line no-unused-vars
const RenderHeaderTotal = (spread, quotation, columnTotal = null, GetterQuotationWorkBook = null) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
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
  // 1setTotalRowValue(sheet, mixTopTotal[combined], totalRowIndex, columnTotal, initTotal.bindPath, null);
  setTotalRowValue(sheet, mixTopTotal[combined], totalRowIndex, initTotal.bindPath, template, quotation);

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
          const lastRow = getHeaderRowCount(template) - 1;
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
export const InitTotal = (spread, GetterQuotationWorkBook = null, quotation = null) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
  const sheet = spread.getActiveSheet();
  const { total } = template.cloudSheet;

  if (showTotal(template)) {
    const tableStartRowIndex = PubGetTableStartRowIndex(template);
    const tableRowCount = PubGetTableRowCount(0, quotation);
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
// eslint-disable-next-line no-unused-vars
const RenderTotal = (spread, columnTotal = null, columnComputed = null, GetterQuotationWorkBook = null, GetterQuotationInit = null) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = GetterQuotationInit || store.getters['quotationModule/GetterQuotationInit'];

  console.log(template, '===== RenderTotal =========');


  const { total, bottom } = template.cloudSheet;
  if (showTotal(template)) {
    const sheet = spread.getActiveSheet();
    sheet.suspendPaint();

    // index
    const bottomRowCount = bottom.rowCount;
    const totalRowIndex = sheet.getRowCount() - bottomRowCount;

    console.log(total);


    const Total = total[templateTotalMap(total.select)];
    sheet.addRows(totalRowIndex, Total.rowCount);

    console.log(Total, 'Total');

    mergeSpan(sheet, Total.spans, totalRowIndex);
    setCellStyle(spread, Total, totalRowIndex, true);
    setTotalRowHeight(sheet, total, Total, totalRowIndex);
    // 1setTotalRowValue(sheet, Total, totalRowIndex, columnTotal, null, columnComputed);
    setTotalRowValue(sheet, Total, totalRowIndex, Total.bindPath, template, quotation);

    sheet.resumePaint();

    if (template.truckage) {
      const TruckageIdentifier = new IdentifierTemplate(sheet, 'truckage', template, quotation);
      TruckageIdentifier.truckageRenderTotal(quotation);
    }
  }
};

/**
 * Product Rendering
 * @param {*} spread
 */
const renderSheet = (spread, GetterQuotationWorkBook, GetterQuotationInit, isCompress = false) => {
  const sheet = spread.getActiveSheet();
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = GetterQuotationInit || store.getters['quotationModule/GetterQuotationInit'];
  const { equipment, type = null, total = null, columnCount } = template.cloudSheet.center;
  const { mixTopTotal = null, image = null, top } = template.cloudSheet;
  const { mixRender, classType, isHaveChild } = templateRenderFlag(GetterQuotationWorkBook);
  const resourceViews = quotation.conferenceHall.resourceViews;

  const noClass = resourceViews.length === 1 && resourceViews[0].name === '无分类';

  if (!noClass) {
    if (type) {
      delTableHeaderRowCount(sheet, type.dataTable, top);
    }
  }

  // Obtain the index of the table
  const tableStartRowIndex = PubGetTableStartRowIndex(template);
  const tableStartColumnIndex = PubGetTableStartColumnIndex();
  const tableColumnCount = PubGetTableColumnCount(template);

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

  console.log(quotation, '============');
  console.log(template, '============');

  const { classRow, subTotal, classRow1, tableHeaderRow } = classificationAlgorithms(quotation, header, template);

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
        setCellStyle(spread, type, rowClassIndex, true);
        PubSetCellHeight(sheet, type, rowClassIndex);
      }
    }

    // Add a list of products
    sheet.addRows(rowClassIndex + classRow + tableHeaderRow, PubGetTableRowCount(i, quotation), GC.Spread.Sheets.SheetArea.viewport);
    // Create a table
    const tableId = resourceViews[i].resourceLibraryId;
    const table = sheet.tables.findByName('table' + tableId);
    if (!table) {
      const { bindPath } = equipment;
      CreateTable(sheet, tableId, rowClassIndex + classRow + tableHeaderRow, tableStartColumnIndex, PubGetTableRowCount(i, quotation), tableColumnCount, bindPath, `conferenceHall.resourceViewsMap.${tableId}.resources`);
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
      const rowClassTotal = rowClassIndex + classRow + tableHeaderRow + PubGetTableRowCount(i, quotation) + headerRow;
      sheet.addRows(rowClassTotal, subTotal, GC.Spread.Sheets.SheetArea.viewport);
      if (total) {
        mergeSpan(sheet, total.spans, rowClassTotal);
        setCellStyle(spread, total, rowClassTotal, true);
        PubSetCellHeight(sheet, total, rowClassTotal);
        for (const key in total.bindPath) {
          if (Object.hasOwnProperty.call(total.bindPath, key)) {
            const field = total.bindPath[key];
            if (field.bindPath) {
              setCellFormatter(sheet, rowClassTotal + field.row, field.column, quotation);
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
    renderSheetImage(spread, tableStartRowIndex, false, true, true, quotation, template, isCompress);
  }

  // Subtotal assignment
  const columnTotal = [];
  const columnComputed = [];
  let insertTableIndex = tableStartRowIndex;
  const subTotalBindPath = total ? total.bindPath : null;
  for (let index = 0; index < resourceViews.length; index++) {
    if (index === 0) {
      const columnTotalMap = columnsTotal(sheet, insertTableIndex + classRow + tableHeaderRow + 1, index, true, columnComputed, subTotalBindPath, template, quotation);
      !noClass && SetComputedSubTotal(sheet, columnTotalMap, subTotalBindPath, quotation);
      columnTotal.push(columnTotalMap);
      insertTableIndex = insertTableIndex + classRow + tableHeaderRow + resourceViews[index].resources.length;
    } else {
      const columnTotalMap = columnsTotal(sheet, insertTableIndex + subTotal + classRow + tableHeaderRow + 1, index, true, null, subTotalBindPath, template, quotation);
      !noClass && SetComputedSubTotal(sheet, columnTotalMap, subTotalBindPath, quotation);
      columnTotal.push(columnTotalMap);

      insertTableIndex = insertTableIndex + subTotal + classRow + tableHeaderRow + resourceViews[index].resources.length;
    }
  }

  sheet.resumePaint();

  if (mixTopTotal) {
    // TODO header class 有新的标识符
    RenderHeaderClass(spread, quotation, columnTotal, template);
    // TODO 顶部组合未同步修改
    RenderHeaderTotal(spread, quotation, columnTotal, template);
  } else {
    RenderTotal(spread, columnTotal, columnComputed, template, quotation);
  }
};

/**
 * Render styles by template type
 * @param {*} spread
 */
export const Render = (spread, GetterQuotationWorkBook, GetterQuotationInit, isCompress = false) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = GetterQuotationInit || _.cloneDeep(store.getters['quotationModule/GetterQuotationInit']);
  console.log('------------解析库--------------');

  console.log(quotation, 'quotation');
  console.log(template, 'template');
  renderSheet(spread, GetterQuotationWorkBook, GetterQuotationInit, isCompress);
  // setLastColumnWidth(spread, template);
  translateSheet(spread);
  initShowCostPrice(spread);
};

const InitWorksheet = (sheet, dataSource) => {
  if (!sheet) return;
  sheet.name('sheet');
  sheet.tag('sheet');
  SetDataSource(sheet, dataSource);
};

const InitBindPath = (spread, template, quotation) => {
  InitBindValueTop(spread, template, quotation);
  const { topPath, conferenceHallTopPath, conferenceHallBottomPath, bottomPath } = getPaths();

  // top
  FieldBindPath(spread, template, topPath);
  // 主会场
  FieldBindPath(spread, template, conferenceHallTopPath);
  FieldBindPath(spread, template, conferenceHallBottomPath);
  // bottom
  FieldBindPath(spread, template, bottomPath);
};

const InitSheetRender = (spread, template, quotation, isCompress = false) => {
  // 逻辑处理
  // LogicalTotalCalculationType(this.spread);
  // render center
  const { conferenceHall } = quotation;
  const resourceViews = conferenceHall.resourceViews;
  if (resourceViews.length) {
    Render(spread, template, quotation, isCompress);
  } else {
    InitTotal(spread, template, quotation);
  }
  renderFinishedAddImage(spread, template, quotation);
};

/**
 * Initialization of a single table
 * @param {*} spread 
 * @param {*} template 
 * @param {*} dataSource 
 * @param {*} isCompress 
 * @returns 
 */
export const initSingleTable = (spread, template, dataSource, isCompress = false) => {
  if (!spread) {
    console.error('spread is null');
    return
  }
  const sheet = spread.getActiveSheet();
  InitWorksheet(sheet, dataSource);
  InitBindPath(spread, template, dataSource)
  InitSheetRender(spread, template, dataSource, isCompress)
};
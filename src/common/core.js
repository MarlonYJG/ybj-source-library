/*
 * @Author: Marlon
 * @Date: 2024-07-26 13:13:18
 * @Description: The corresponding row and column indexes are calculated based on the classification
 */
import store from 'store';
import { IGNORE_EVENT } from 'store/quotation/mutation-types';
import { sortObjectByRow } from '../utils/index'

import { LOG_STYLE_1 } from './log-style';
import { CombinationTypeBuild } from './combination-type';
import { getTemplateClassType } from './single-table';
import { showTotal, getPriceColumn } from './parsing-template';

const NzhCN = require('../lib/nzh/cn.min.js');

/**
 * Get quote data 
 * @returns 
 */
const getQuotationResource = (GetterQuotationInit = null) => {
  const quotation = GetterQuotationInit || store.getters['quotationModule/GetterQuotationInit'];
  return quotation.conferenceHall.resourceViews;
}

/**
 * Get template data
 * @param {*} GetterQuotationWorkBook 
 * @returns 
 */
const getTemplateCloudSheet = (GetterQuotationWorkBook = null) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
  return template.cloudSheet;
}


/**
 * A mapping table that calculates the row and column indexes based on the classification type
 */
export class LayoutRowColBlock {
  static ClassType = null;
  static Levels = [];
  static Tables = [];
  static SubTotals = [];
  static Summations = [];
  static TotalMap = null;
  static LevelsRowMap = null;

  constructor(spread, template, quotation) {
    this.Template = template;
    this.Quotation = quotation;
    this.Spread = spread;
    if (!template && store && store.getters) {
      this.Template = store.getters['quotationModule/GetterQuotationWorkBook'];
    }
    if (!quotation && store && store.getters) {
      this.Quotation = store.getters['quotationModule/GetterQuotationInit'];
    }
    this._init();
  }

  _init() {
    const classType = getTemplateClassType(this.Quotation, this.Template);
    console.log(`%c 报价单分类类型：${classType}`, LOG_STYLE_1);

    LayoutRowColBlock.ClassType = classType;
    this._initLayout(classType)
  }
  /**
   * Get the layout
   * @param {*} classType 
   */
  _initLayout(classType) {
    if (classType) {
      switch (classType) {
        case 'noLevel':
          this._noLevel();
          break;
        case 'Level_1_row':
          this._Level_1_row();
          break;
        case 'Level_1_col':
          this._Level_1_col();
          break;
        default:
          break;
      }
    } else {
      console.error('The template classification type does not exist');
    }
  }

  _noLevel() {
    const sheet = this.Spread.getActiveSheet();
    let tableMap = {};
    let totalMap = null;
    let levelsRowMap = {};

    const { top, bottom, total } = getTemplateCloudSheet(this.Template);

    const resourceViews = getQuotationResource(this.Quotation);
    for (let i = 0; i < resourceViews.length; i++) {
      const item = resourceViews[i];
      const resourceLibraryId = item.resourceLibraryId;
      if (i === 0) {
        levelsRowMap[resourceLibraryId] = top.rowCount;
        tableMap[resourceLibraryId] = {
          row: top.rowCount,
          rowCount: item.resources.length,
        }
      }
    }

    tableMap = sortObjectByRow(tableMap)

    if (showTotal(this.Template)) {
      const quotation = this.Quotation || store.getters['quotationModule/GetterQuotationInit'];
      const Total = total[CombinationTypeBuild(quotation)];
      if (Total) {
        const rowCount = sheet.getRowCount();
        const bottomCount = bottom.rowCount;
        totalMap = {
          row: rowCount - bottomCount - Total.rowCount,
          rowCount: Total.rowCount,
        }
      } else {
        console.error('The total does not exist [the corresponding value of the permutation and combination has not been calculated]');
      }
    }

    LayoutRowColBlock.Levels = null;
    LayoutRowColBlock.Tables = tableMap;
    LayoutRowColBlock.SubTotals = null;
    LayoutRowColBlock.Summations = null;
    LayoutRowColBlock.TotalMap = totalMap;
    LayoutRowColBlock.LevelsRowMap = levelsRowMap;
  }

  _Level_1_row() {
    const sheet = this.Spread.getActiveSheet();
    let level_1 = {};
    let tableMap = {};
    let subTotalMap = {};
    let totalMap = null;
    let levelsRowMap = {};

    const { top, bottom, total } = getTemplateCloudSheet(this.Template);

    const resourceViews = getQuotationResource(this.Quotation);
    for (let i = 0; i < resourceViews.length; i++) {
      const item = resourceViews[i];
      const resourceLibraryId = item.resourceLibraryId;

      if (i === 0) {

        levelsRowMap[resourceLibraryId] = top.rowCount;
        level_1[resourceLibraryId] = {
          row: top.rowCount,
          rowCount: 1,
        }

        tableMap[resourceLibraryId] = {
          row: level_1[resourceLibraryId].row + level_1[resourceLibraryId].rowCount,
          rowCount: item.resources.length,
        }

        subTotalMap[resourceLibraryId] = {
          row: tableMap[resourceLibraryId].row + tableMap[resourceLibraryId].rowCount,
          rowCount: 1,
        }
      } else {
        const row = subTotalMap[resourceViews[i - 1].resourceLibraryId].row + subTotalMap[resourceViews[i - 1].resourceLibraryId].rowCount;
        levelsRowMap[resourceLibraryId] = row;
        level_1[resourceLibraryId] = {
          row: row,
          rowCount: 1,
        }

        tableMap[resourceLibraryId] = {
          row: level_1[resourceLibraryId].row + level_1[resourceLibraryId].rowCount,
          rowCount: item.resources.length,
        }

        subTotalMap[resourceLibraryId] = {
          row: tableMap[resourceLibraryId].row + tableMap[resourceLibraryId].rowCount,
          rowCount: 1,
        }

      }
    }

    level_1 = sortObjectByRow(level_1)
    tableMap = sortObjectByRow(tableMap)
    subTotalMap = sortObjectByRow(subTotalMap)

    if (showTotal(this.Template)) {
      const quotation = this.Quotation || store.getters['quotationModule/GetterQuotationInit'];
      const Total = total[CombinationTypeBuild(quotation)];
      if (Total) {
        const rowCount = sheet.getRowCount();
        const bottomCount = bottom.rowCount;
        totalMap = {
          row: rowCount - bottomCount - Total.rowCount,
          rowCount: Total.rowCount,
        }
      } else {
        console.error('The total does not exist [the corresponding value of the permutation and combination has not been calculated]');
      }
    }

    LayoutRowColBlock.Levels = [level_1];
    LayoutRowColBlock.Tables = tableMap;
    LayoutRowColBlock.SubTotals = subTotalMap;
    LayoutRowColBlock.Summations = null;
    LayoutRowColBlock.TotalMap = totalMap;
    LayoutRowColBlock.LevelsRowMap = levelsRowMap;
  }

  _Level_1_col() {
    const sheet = this.Spread.getActiveSheet();
    let level_1 = {};
    let tableMap = {};
    let subTotalMap = {};
    let totalMap = null;
    let levelsRowMap = {};

    const { top, bottom, total } = getTemplateCloudSheet(this.Template);

    const resourceViews = getQuotationResource(this.Quotation);
    for (let i = 0; i < resourceViews.length; i++) {
      const item = resourceViews[i];
      const resourceLibraryId = item.resourceLibraryId;

      if (i === 0) {

        levelsRowMap[resourceLibraryId] = top.rowCount;

        level_1[resourceLibraryId] = {
          row: top.rowCount,
          rowCount: item.resources.length,
        }

        tableMap[resourceLibraryId] = {
          row: top.rowCount,
          rowCount: item.resources.length,
        }

        subTotalMap[resourceLibraryId] = {
          row: item.resources.length + 1,
          rowCount: 1,
        }
      } else {
        const row = subTotalMap[resourceViews[i - 1].resourceLibraryId].row + subTotalMap[resourceViews[i - 1].resourceLibraryId].rowCount;

        levelsRowMap[resourceLibraryId] = row;

        level_1[resourceLibraryId] = {
          row: row,
          rowCount: item.resources.length,
        }

        tableMap[resourceLibraryId] = {
          row: row,
          rowCount: item.resources.length,
        }

        subTotalMap[resourceLibraryId] = {
          row: tableMap[resourceLibraryId].row + tableMap[resourceLibraryId].rowCount,
          rowCount: 1,
        }

      }
    }

    level_1 = sortObjectByRow(level_1)
    tableMap = sortObjectByRow(tableMap)
    subTotalMap = sortObjectByRow(subTotalMap)

    if (showTotal(this.Template)) {
      const quotation = this.Quotation || store.getters['quotationModule/GetterQuotationInit'];
      const Total = total[CombinationTypeBuild(quotation)];
      if (Total) {
        const rowCount = sheet.getRowCount();
        const bottomCount = bottom.rowCount;
        totalMap = {
          row: rowCount - bottomCount - Total.rowCount,
          rowCount: Total.rowCount,
        }
      } else {
        console.error('The total does not exist [the corresponding value of the permutation and combination has not been calculated]');
      }
    }

    LayoutRowColBlock.Levels = [level_1];
    LayoutRowColBlock.Tables = tableMap;
    LayoutRowColBlock.SubTotals = subTotalMap;
    LayoutRowColBlock.Summations = null;
    LayoutRowColBlock.TotalMap = totalMap;
    LayoutRowColBlock.LevelsRowMap = levelsRowMap;
  }

  /**
   * Update the cell value
   * @param {*} row 
   * @param {*} col 
   * @param {*} value 
   */
  _updateCellValue(row, col, value) {
    this.Spread.getActiveSheet().setValue(row, col, value);
  }

  /**
   * Obtain the template classification type
   * @returns 
   */
  getClassType() {
    console.log(LayoutRowColBlock.ClassType, '模板分类类型');
    return LayoutRowColBlock.ClassType;
  }

  getLayout() {
    console.log(LayoutRowColBlock.Levels, '层级');
    console.log(LayoutRowColBlock.Tables, '布局');
    console.log(LayoutRowColBlock.SubTotals, '子表');
    console.log(LayoutRowColBlock.Summations, '子表合计');
    console.log(LayoutRowColBlock.TotalMap, '合计');

    return {
      Levels: LayoutRowColBlock.Levels,
      Tables: LayoutRowColBlock.Tables,
      SubTotals: LayoutRowColBlock.SubTotals,
      Summations: LayoutRowColBlock.Summations,
      TotalMap: LayoutRowColBlock.TotalMap,
    }
  }

  getLevel() {
    const levels = LayoutRowColBlock.Levels;
    const levelCell = [];
    if (levels && levels.length) {
      for (let i = 0; i < levels.length; i++) {
        for (let j = 0; j < levels[i].length; j++) {
          for (const key in levels[i][j]) {
            if (Object.hasOwnProperty.call(levels[i][j], key)) {
              levelCell.push(levels[i][j][key]);
            }
          }
        }
      }
    }
    return {
      rows: levelCell.map((item) => item.row),
      cells: levelCell
    }
  }

  /**
   * Get the row data in the table
   * @param {*} tables 
   * @returns 
   */
  getTableMapRows(tables = []) {
    const tableMapRows = {};
    tables.forEach(item => {
      for (const key in item) {
        tableMapRows[key] = [];
        if (Object.prototype.hasOwnProperty.call(item, key)) {
          for (let i = item[key].row; i < (item[key].rowCount + item[key].row); i++) {
            tableMapRows[key].push(i);
          }
        }
      }
    });
    if (!Object.keys(tableMapRows).length) {
      return null;
    }
    return tableMapRows;
  }

  /**
   * Gets and sets all the uppercase amounts for the totals area
   * @param {*} quotation 
   */
  setTotalUppercaseAmounts(quotation) {
    if (quotation && showTotal(this.Template)) {
      const totalMap = LayoutRowColBlock.TotalMap;
      const total = getTemplateCloudSheet().total;
      const Total = total[CombinationTypeBuild(quotation)];
      console.log('------------Update capitalization--------');

      if (Total) {
        if (Object.keys(Total.bindPath).includes("DXzje")) {
          const bindFields = {};
          for (let i = 0; i < totalMap.rowCount; i++) {
            for (const key in Total.bindPath) {
              if (Object.prototype.hasOwnProperty.call(Total.bindPath, key)) {
                bindFields[key] = {
                  row: totalMap.row + Total.bindPath[key].row,
                  col: Total.bindPath[key].column
                }
              }
            }
          }

          const sheet = this.Spread.getActiveSheet();
          // Get the uppercase amount (concessional > totalAfterTax > totalBeforeTax > sumAmount )
          let value = quotation.sumAmount || 0;
          if (Object.keys(Total.bindPath).includes("concessional")) {
            const { row, col } = bindFields.concessional;
            value = sheet.getValue(row, col);
          } else if (Object.keys(Total.bindPath).includes("totalAfterTax")) {
            const { row, col } = bindFields.totalAfterTax;
            value = sheet.getValue(row, col);
          } else if (Object.keys(Total.bindPath).includes("totalBeforeTax")) {
            const { row, col } = bindFields.totalBeforeTax;
            value = sheet.getValue(row, col);
          }
          if (value === 0 || value) {
            // Update the uppercase value
            const { row, col } = bindFields.DXzje;
            store.commit(`quotationModule/${IGNORE_EVENT}`, true);
            sheet.suspendPaint();
            this._updateCellValue(row, col, NzhCN.encodeB(value));
            sheet.resumePaint();
          }
        }
      }
    }

  }

  /**
   * Obtain products based on classification IDs and indexes
   * @param {*} resourceLibraryId 
   * @param {*} index 
   * @returns 
   */
  getProductByIndex(resourceLibraryId, index) {
    const resourceViews = getQuotationResource();
    for (let i = 0; i < resourceViews.length; i++) {
      if (resourceViews[i].resourceLibraryId === resourceLibraryId) {
        return resourceViews[i].resources[index];
      }
    }
    return null;
  }

  /**
   * Get the product it belongs to based on the active cell
   * @param {*} activeRow 
   * @param {*} activeCol 
   * @param {*} tableId 
   * @returns 
   */
  getProductByActiveCell(activeRow, activeCol, tableId) {
    console.log(tableId);
    const leMap = LayoutRowColBlock.LevelsRowMap;
    const classType = LayoutRowColBlock.ClassType;
    if (['noLevel', 'Level_1_col'].includes(classType)) {
      const sortIndex = activeRow - (leMap[tableId]);
      if (sortIndex >= 0) {
        return this.getProductByIndex(tableId, sortIndex);
      } else {
        console.error('获取产品失败');
      }
    } else if (['Level_1_row'].includes(classType)) {
      const sortIndex = activeRow - (leMap[tableId] + 1);
      if (sortIndex >= 0) {
        return this.getProductByIndex(tableId, sortIndex);
      } else {
        console.error('获取产品失败');
      }
    }
    console.error('获取产品失败');
    return null;
  }

  /**
 * Get Configuration Information (First Data)
 */
  getConfigFirstData() {
    const resourceViews = getQuotationResource();
    if (resourceViews.length) {
      const resources = resourceViews[0].resources;
      if (resources && resources.length) {
        const firstTable = LayoutRowColBlock.Tables[0];
        const sheet = this.Spread.getActiveSheet();
        let tableId = null;

        for (const key in firstTable) {
          if (Object.prototype.hasOwnProperty.call(firstTable, key)) {
            tableId = key;
          }
        }
        const table = sheet.tables.findByName(`table${tableId}`);
        const priceCol = getPriceColumn();
        if (table && (priceCol === 0 || priceCol)) {
          return sheet.getFormatter(firstTable[tableId].row, priceCol);
        }
      }
    }
    return null
  }
}

export default LayoutRowColBlock;
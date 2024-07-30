/*
 * @Author: Marlon
 * @Date: 2024-07-26 13:13:18
 * @Description: The corresponding row and column indexes are calculated based on the classification
 */
import store from 'store';
import { sortObjectByRow } from '../utils/index'

import { CombinationTypeBuild } from './combination-type'
import { getTemplateClassType } from './single-table';
import { showTotal } from './parsing-template';

/**
 * Get quote data 
 * @returns 
 */
const getQuotationResource = () => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  return quotation.conferenceHall.resourceViews;
}

/**
 * Get template data
 * @returns 
 */
const getTemplateCloudSheet = () => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
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

  constructor(spread) {
    this.Spread = spread;
    this.init();
  }

  init() {
    const classType = getTemplateClassType();
    LayoutRowColBlock.ClassType = classType;
    this.initLayout(classType)
  }
  /**
   * Get the layout
   * @param {*} classType 
   */
  initLayout(classType) {
    if (classType) {
      switch (classType) {
        case 'noLevel':
          this.noLevel();
          break;
        case 'Level_1_row':
          this.Level_1_row();
          break;
        case '':
          // TODO
          break;
        default:
          break;
      }
    } else {
      console.error('The template classification type does not exist');
    }
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

  // 无分类
  noLevel() {
    const sheet = this.Spread.getActiveSheet();
    let tableMap = {};
    let totalMap = null;

    const { top, bottom, total } = getTemplateCloudSheet();

    const resourceViews = getQuotationResource();
    for (let i = 0; i < resourceViews.length; i++) {
      const item = resourceViews[i];
      const resourceLibraryId = item.resourceLibraryId;
      if (i === 0) {
        tableMap[resourceLibraryId] = {
          row: top.rowCount,
          rowCount: item.resources.length,
        }
      }
    }

    tableMap = sortObjectByRow(tableMap)

    if (showTotal()) {
      const quotation = store.getters['quotationModule/GetterQuotationInit'];
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

  }

  Level_1_row() {
    const sheet = this.Spread.getActiveSheet();
    let level_1 = {};
    let tableMap = {};
    let subTotalMap = {};
    let totalMap = null;

    const { top, bottom, total } = getTemplateCloudSheet();

    const resourceViews = getQuotationResource();
    for (let i = 0; i < resourceViews.length; i++) {
      const item = resourceViews[i];
      const resourceLibraryId = item.resourceLibraryId;

      if (i === 0) {

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
        level_1[resourceLibraryId] = {
          row: subTotalMap[resourceViews[i - 1].resourceLibraryId].row + subTotalMap[resourceViews[i - 1].resourceLibraryId].rowCount,
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

    if (showTotal()) {
      const quotation = store.getters['quotationModule/GetterQuotationInit'];
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
  }
}

export default LayoutRowColBlock;
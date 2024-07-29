/*
 * @Author: Marlon
 * @Date: 2024-07-26 13:13:18
 * @Description: The corresponding row and column indexes are calculated based on the classification
 */
import store from 'store';
import { sortObjectByRow } from '../utils/index'
import { getTemplateClassType } from './single-table';

/**
 * A mapping table that calculates the row and column indexes based on the classification type
 */
export const LayoutRowColBlock = () => {
  const classType = getTemplateClassType();
  console.log(classType, '模板分类类型');
  if (classType) {
    switch (classType) {
      case 'noLevel':

        break;
      case 'Level_1_row':
        {
          return Level_1_row();
        }
      default:
        break;
    }
  }
}

// 获取报价单数据
const getQuotationResource = () => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  return quotation.conferenceHall.resourceViews;
}

// 获取模板数据
const getTemplateCloudSheet = () => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  return template.cloudSheet;
}

const Level_1_row = () => {
  let level_1 = {};
  let tableMap = {};
  let subTotalMap = {};

  const { top } = getTemplateCloudSheet();

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


  console.log(level_1, tableMap, subTotalMap, '一级分类');

  return {
    levels: [level_1],
    tables: tableMap,
    subTotals: subTotalMap,
  }
}
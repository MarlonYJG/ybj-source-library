/*
 * @Author: Marlon
 * @Date: 2024-06-06 09:55:19
 * @Description:
 */
import { v4 as uuidv4 } from 'uuid';
import _ from '../lib/lodash/lodash.min.js';
import store from 'store';

import { getWorkBook, getInitData } from './store';
import { flattenArray } from '../utils/index'

/**
 * Price format
 * @returns
 */
export const formatterPrice = (quotationInit) => {
  const quotation = getInitData(quotationInit);
  let symbol = '';
  let type = '';
  let typeInt = '0.00';

  if (quotation.isInt) {
    typeInt = '0';
  }
  switch (quotation.priceType) {
    case '1':
      type = '¥';
      break;
    case '2':
      type = '$';
      break;
    case '3':
      type = '';
      break;
    default:
      type = '';
      break;
  }
  if (type) {
    symbol = `${type}#${typeInt}`;
  }
  return symbol;
};

/**
 * Get distribution locations for classifications, subtotals, and totals
 * @returns
 */
export const getPositionBlock = () => {
  const { leavel1, leavel2, subTotal } = store.getters['quotationModule/GetterClassLeavel'];
  const leavel1Area = leavel1.map((item) => item.row);
  const leavel2Area = leavel2.map((item) => item.row);
  const subTotalArea = subTotal.map((item) => item.row);

  return {
    leavel1Area,
    leavel2Area,
    subTotalArea
  };
};

/**
 * Get classification name
 * @param {*} classificationId
 * @returns
 */
const getClassificationName = (classificationId) => {
  const classificationData = getQuotationAllClassification();
  const classification = flattenArray(classificationData, 'children');
  let classname = classification.filter((item) => item.classificationId == classificationId)
  return classname;
};

/**
 * build datas
 * @param {*} tableId
 * @param {*} insertIndex
 * @param {*} count
 * @returns
 */
export const buildData = (tableId, insertIndex, count, classType) => {
  const quotation = getInitData();
  const conferenceHall = _.cloneDeep(quotation.conferenceHall);
  const template = getWorkBook();
  const { equipment } = template.cloudSheet.center;

  const resource = {};
  const fileds = equipment.bindPath;
  for (const key in fileds) {
    if (Object.hasOwnProperty.call(fileds, key)) {
      resource[key] = null;
    }
  }
  const resourceCount = [];
  const cname = getClassificationName(tableId);
  for (let index = 0; index < count; index++) {
    const item = {
      quantity: null,
      classification: tableId,
      classificationName: cname.length ? cname[0].classificationName : '',
      parentClassification: tableId,
      parentClassificationName: cname.length ? cname[0].classificationName : ''
    }
    if (!['noLevel', 'Level_1_row'].includes(classType)) {
      // TODO 如果是二级分类,修改二级分类名称以及parentClassification
    }
    const initDataItem = {
      id: uuidv4(),
      imageId: uuidv4()
    }
    resourceCount.push(Object.assign(item, resource, initDataItem));
  }

  const resourceViews = conferenceHall.resourceViews;
  for (let index = 0; index < resourceViews.length; index++) {
    if (resourceViews[index].resourceLibraryId == tableId) {
      resourceViews[index].resources.splice(insertIndex, 0, ...resourceCount);
    }
  }

  resourceViews.forEach(item => {
    if (item.resources.length) {
      item.resources.forEach((item, i) => {
        item.sort = i + 1;
      });
    }
  });
  const resourceViewsMap = {};
  resourceViews.forEach(item => {
    resourceViewsMap[item.resourceLibraryId] = item;
  });
  return conferenceHall;
};

/**
 * Get all categories in a quote
 */
export const getQuotationAllClassification = () => {
  const quotation = getInitData();
  const conferenceHall = quotation.conferenceHall;
  const resourceViews = conferenceHall.resourceViews;
  const classification = [];
  resourceViews.forEach(item => {
    let classp = {};
    if (item.resourceLibraryId) {
      classp = {
        classificationId: item.resourceLibraryId,
        classificationName: item.name,
        children: [] // TODO 二级分类
      }
      classification.push(classp);
    }
  })
  return classification;
}

/**
 * Get the cost price display status
 * @returns 
 */
export const getShowCostPrice = () => {
  return getInitData().showCost;
}

/**
 * Initialize the percentage of the discount
 * @param {*} quotation 
 * @returns 
 */
export const initDiscountPercentage = (quo) => {
  const quotation = getInitData(quo);
  if (quotation) {
    return quotation.priceAdjustment || 1;
  }
  return 1
}

/**
 * Initialize the Price Setup field
 * @param {*} quotation 
 * @returns 
 */
export const initPriceSetField = (quo) => {
  const quotation = getInitData(quo);
  if (quotation && quotation.priceStatus) {
    return Number(quotation.priceStatus);
  }
  return 0;
}

/**
 * Obtain configuration information
 * @param {*} quotation 
 * @returns 
 */
export const getConfig = (quotation) => {
  quotation = getInitData(quotation);
  if (quotation.config) {
    return quotation.config;
  }
  console.warn('quotation.config is null');
  return null;
}

/**
 * Determine whether the current quotation is a single table or a total score table
 * @param {*} quotation 
 * @returns 
 */
export const isSingleTable = (quotation) => {
  quotation = getInitData(quotation);
  if (quotation.resources && quotation.resources.length) {
    return false;
  }
  return true;
}

/**
 * 获取分表对应模板的索引
 * @param {*} quotation 
 * @returns 
 */
export const getSheetTemplateIndexs = (quotation) => {
  quotation = getInitData(quotation);
  if (quotation.title && quotation.title.indexOf('@') !== -1) {
    const templateIndex = quotation.title.split('@')[1];
    return templateIndex.split('-');
  }
  return [];
}

/**
 * Gets all shards and returns
 * @param {*} quotation 
 * @returns 
 */
export const getAllSheet = (quotation) => {
  let trunks = quotation.resources || [];
  if (!(trunks.length === 1 && trunks[0].name === 'noProject')) {
    return trunks
  } else {
    return trunks = []
  }
}

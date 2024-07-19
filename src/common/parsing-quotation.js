/*
 * @Author: Marlon
 * @Date: 2024-06-06 09:55:19
 * @Description:
 */
import { v4 as uuidv4 } from 'uuid';
import _ from 'lodash';
import store from 'store';

import { flattenArray } from '../utils/index'

/**
 * Price format
 * @returns
 */
export const formatterPrice = () => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
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
export const buildData = (tableId, insertIndex, count) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const conferenceHall = _.cloneDeep(quotation.conferenceHall);
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
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
    resourceCount.push(Object.assign({
      quantity: null,
      classification: '',
      classificationName: '',
      parentClassification: tableId,
      parentClassificationName: cname.length ? cname[0].classificationName : '',
      classname: cname.length ? cname[0].classificationName : '' // TODO
    }, resource, { id: uuidv4() }));
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
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const conferenceHall = quotation.conferenceHall;
  const resourceViews = conferenceHall.resourceViews;
  const classification = [];
  resourceViews.forEach(item => {
    const classp = {};
    if (item.resourceLibraryId) {
      classp = {
        classificationId: item.resourceLibraryId,
        classificationName: item.name,
        children: []
      }

      if (item.resources.length) {
        item.resources.forEach(item => {
          classp.children.push({
            classificationId: item.classification,
            classificationName: item.classificationName
          });
        })
      }

      classification.push(classp);
    }
  })
}
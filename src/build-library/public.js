/*
 * @Author: Marlon
 * @Date: 2024-05-16 14:08:41
 * @Description:
 */
import store from 'store';

import { getTableHeaderDataTable } from '../common/parsing-template';

import { PubGetTableStartRowIndex, classificationAlgorithms } from '../common/single-table';

/**
 * reset
 * @param {*} spread
 */
export const Reset = (spread) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const noClass = resourceViews.length === 1 && resourceViews[0].name === '无分类';

  const { type = null } = template.cloudSheet.center;
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

  let centerCount = tableHeaderRow;
  resourceViews.forEach((item) => {
    centerCount = centerCount + item.resources.length + classRow + subTotal;
  });

  const sheet = spread.getActiveSheet();
  sheet.suspendPaint();
  sheet.deleteRows(tableStartRowIndex, centerCount);
  sheet.resumePaint();
};

/**
 * Update sort of conferenceHall
 * @param {*} conferenceHall
 * @returns
 */
export const UpdateSort = (conferenceHall) => {
  const resourceViews = conferenceHall.resourceViews;
  const resourceViewsMap = {};
  if (resourceViews.length) {
    resourceViews.forEach((el1, i) => {
      el1.sort = i + 1;
      if (el1.resources) {
        el1.resources.forEach((el2, index) => {
          el2.sort = index + 1;
        });
      }
      resourceViewsMap[el1.resourceLibraryId] = el1;
    });
  }
  conferenceHall.resourceViewsMap = resourceViewsMap;
  return conferenceHall;
};

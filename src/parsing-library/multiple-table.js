/*
 * @Author: Marlon
 * @Date: 2024-05-14 21:52:13
 * @Description:multipleTable
 */
import { rootWorkBook } from '../common/core';
import { CreateSheet } from '../common/sheetWorkBook';

import { getAllSheet } from '../common/parsing-quotation';

import { SetDataSource } from '../common/sheetWorkBook';
import { InitBindPath } from '../common/single-table';
import { templateMap, OnEventBind } from '../common/multiple-table';

import { initWorkBookConfig } from './public';
import { InitSheetRender } from './single-table';

/**
 * Initialize multiple sheet render
 * @param {*} spread 
 * @param {*} dataSource 
 * @param {*} templateSource 
 * @param {*} isCompress 
 * @returns 
 */
const InitWorksheet = (spread, dataSource, templateSource, isCompress) => {
  if (!spread) return;
  const trunks = getAllSheet(dataSource);

  console.log(trunks);

  if (trunks.length) {
    for (let index = 0; index < trunks.length; index++) {
      const tMap = templateMap(dataSource, templateSource);
      CreateSheet(spread, trunks[index].name, index + 1, tMap[trunks[index].name].sheets[trunks[index].name]);
    }
    spread.setSheetCount(trunks.length + 1);
  }
  spread.setActiveSheetIndex(0);

  InitMainSheetRender(spread, templateSource, dataSource, isCompress);

};

/**
 * Draw a summary table
 * @param {*} spread 
 * @param {*} templateSource 
 * @param {*} dataSource 
 * @param {*} isCompress 
 * @param {*} type 
 */
const InitMainSheetRender = (spread, templateSource, dataSource, isCompress) => {
  const templateMapData = templateMap(dataSource, templateSource);
  console.log(templateMapData, 'templateMapData');
  rootWorkBook._setTemplateMap(templateMapData);
  rootWorkBook._setActiveTemplate(templateSource);
  console.log(templateMapData);

  OnEventBind(spread, dataSource, templateSource, isCompress, 'parsing', templateMapData);

  const sheet = spread.getActiveSheet();
  SetDataSource(sheet, dataSource);
  InitBindPath(spread, templateSource, dataSource);
}

/**
 * Initialize the total score table
 * @param {*} spread 
 * @param {*} templateSource 
 * @param {*} dataSource 
 * @param {*} isCompress 
 * @returns 
 */
export const initMultipleTable = (spread, templateSource, dataSource, isCompress = false) => {
  if (!spread) {
    return console.error('spread is null');
  }
  console.log(templateSource);
  rootWorkBook._setWorkBook(spread);
  rootWorkBook._setActiveQuotation(dataSource);
  initWorkBookConfig(spread);
  InitWorksheet(spread, dataSource, templateSource, isCompress);
};

/**
 * center sheet Render
 * @param {*} spread 
 * @param {*} template 
 * @param {*} quotation 
 * @param {*} isCompress 
 */
export const CenterSheetRender = (spread, template, quotation, isCompress) => {
  InitSheetRender(spread, template, quotation, isCompress);
}

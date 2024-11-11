/*
 * @Author: Marlon
 * @Date: 2024-05-14 21:52:13
 * @Description:multipleTable
 */
import { rootWorkBook } from '../common/core';
import { getWorkBook, getInitData } from '../common/store';
import { CreateSheet } from '../common/sheetWorkBook';

import { getAllSheet } from '../common/parsing-quotation';

import { InitMainSheetRender, templateMap } from '../common/multiple-table';

import { initWorkBookConfig } from './public';
import { InitSheetRender } from './single-table';

/**
 * Initialize multiple sheet render
 * @param {*} spread 
 * @param {*} dataSource 
 * @param {*} template 
 * @param {*} isCompress 
 * @returns 
 */
const InitWorksheet = (spread, dataSource, template, isCompress) => {
  if (!spread) return;
  const trunks = getAllSheet(dataSource);

  console.log(trunks);

  if (trunks.length) {
    for (let index = 0; index < trunks.length; index++) {
      const tMap = templateMap(dataSource, template);
      CreateSheet(spread, trunks[index].name, index + 1, tMap[trunks[index].name].sheets[trunks[index].name]);
    }
    spread.setSheetCount(trunks.length + 1);
  }
  spread.setActiveSheetIndex(0);

  InitMainSheetRender(spread, template, dataSource, isCompress, 'parsing');

};

/**
 * Initialize the total score table
 * @param {*} spread 
 * @param {*} template 
 * @param {*} dataSource 
 * @param {*} isCompress 
 * @returns 
 */
export const initMultipleTable = (spread, template, dataSource, isCompress = false) => {
  if (!spread) {
    return console.error('spread is null');
  }
  console.log(template);
  rootWorkBook._setWorkBook(spread);
  rootWorkBook._setActiveQuotation(dataSource); 
  initWorkBookConfig(spread);
  InitWorksheet(spread, dataSource, template, isCompress);
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

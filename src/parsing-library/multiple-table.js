/*
 * @Author: Marlon
 * @Date: 2024-05-14 21:52:13
 * @Description:multipleTable
 */
import { getWorkBook, getInitData } from '../common/store';
import { CreateSheet } from '../common/sheetWorkBook';

import { getAllSheet } from '../common/parsing-quotation';

import { InitMainSheetRender, templateMap } from '../common/multiple-table';

import { initWorkBookConfig } from './public';

const InitWorksheet = (spread, dataSource, template, isCompress) => {
  if (!spread) return;
  if (!dataSource) {
    dataSource = getInitData();
  }
  if (!template) {
    template = getWorkBook();
  }

  const trunks = getAllSheet(dataSource);

  console.log(trunks);

  if (trunks.length) {
    for (let index = 0; index < trunks.length; index++) {
      const tMap = templateMap(dataSource, template);
      CreateSheet(spread, trunks[index].name, index + 1, tMap[trunks[index].name].sheets.sheet);
    }
    spread.setSheetCount(trunks.length + 1);
  }
  spread.setActiveSheetIndex(0);

  InitMainSheetRender(spread, template, dataSource, isCompress);

};

export const initMultipleTable = (spread, template, dataSource, isCompress = false) => {
  if (!spread) {
    console.error('spread is null');
    return
  }
  initWorkBookConfig(spread);

  console.log(template, 'template');
  console.log(dataSource, 'dataSource');

  InitWorksheet(spread, dataSource, template, isCompress);
};




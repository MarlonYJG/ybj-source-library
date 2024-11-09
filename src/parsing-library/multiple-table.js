/*
 * @Author: Marlon
 * @Date: 2024-05-14 21:52:13
 * @Description:multipleTable
 */
import _ from '../lib/lodash/lodash.min.js';

import { getWorkBook, getInitData } from '../common/store';
import { CreateSheet } from '../common/sheetWorkBook';

import { getTrunkTemplate } from '../common/parsing-template';
import { getAllSheet } from '../common/parsing-quotation';

import { SetDataSource } from '../common/sheetWorkBook';
import { InitMainSheetRender, getSheetTemplateIndexs } from '../common/multiple-table';

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
    const trunkTemplate = getTrunkTemplate(template);
    const trunkIndex = getSheetTemplateIndexs(dataSource, trunkTemplate);

    console.log('trunkIndex', trunkIndex);
    console.log('trunkTemplate', trunkTemplate);

    let sheetTemplateData = trunkTemplate[0];
    for (let index = 0; index < trunks.length; index++) {
      if (trunkIndex.length && trunkIndex[index]) {
        sheetTemplateData = trunkTemplate[trunkIndex[index]];
      }
      sheetTemplateData.sheets['分表'].name = trunks[index].name;
      const sheetTemplate = _.cloneDeep(sheetTemplateData.sheets['分表']);
      CreateSheet(spread, trunks[index].name, index + 1, sheetTemplate);
    }
    spread.setSheetCount(trunks.length + 1);
  }
  spread.setActiveSheetIndex(0);

  const sheet = spread.getActiveSheet();
  SetDataSource(sheet, dataSource);

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

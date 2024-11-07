/*
 * @Author: Marlon
 * @Date: 2024-05-14 21:52:13
 * @Description:multipleTable
 */
import _ from '../lib/lodash/lodash.min.js';

import { getWorkBook, getInitData } from '../common/store';
import { CreateSheet } from '../common/sheetWorkBook';

import { getTrunkTemplate } from '../common/parsing-template';
import { getSheetTemplateIndexs } from '../common/parsing-quotation';

import { initWorkBookConfig } from './public';

/**
 * Multiple render
 * @param {*} spread 
 * @param {*} GetterMultipleWorkBook 
 * @param {*} GetterMultipleInitData 
 * @param {*} isCompress 
 * @returns 
 */
export const MultipleRender = (spread, GetterMultipleWorkBook = null, GetterMultipleInitData = null, isCompress = false) => {
  if (!spread) return;
  const template = getWorkBook(GetterMultipleWorkBook);
  const quotation = getInitData(GetterMultipleInitData);

  console.log('quotation', quotation);
  console.log('template', template);




};

const InitWorksheet = (spread, dataSource, template) => {
  if (!spread) return;
  const trunks = dataSource.resources || [];

  console.log(trunks);

  if (!(trunks.length === 1 && trunks[0].name === 'noProject')) {
    const trunkTemplate = getTrunkTemplate(template);
    const trunkIndex = getSheetTemplateIndexs(dataSource);

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
  }







};

export const initMultipleTable = (spread, template, dataSource, isCompress = false) => {
  if (!spread) {
    console.error('spread is null');
    return
  }
  initWorkBookConfig(spread);
  console.log(template, 'template');

  console.log(dataSource, 'dataSource');

  // const sheet = spread.getActiveSheet();
  InitWorksheet(spread, dataSource, template);
  // InitBindPath(spread, template, dataSource)
  // InitSheetRender(spread, template, dataSource, isCompress)
};

/*
 * @Author: Marlon
 * @Date: 2024-05-14 21:52:13
 * @Description:multipleTable
 */
import { getWorkBook, getInitData } from '../common/store';
import { CreateTable, SetDataSource } from '../common/sheetWorkBook';

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

const InitWorksheet = (sheet, dataSource) => {
  if (!sheet) return;
  // sheet.name('sheet');
  // sheet.tag('sheet');
  SetDataSource(sheet, dataSource);
};

export const initMultipleTable = (spread, template, dataSource, isCompress = false) => {
  if (!spread) {
    console.error('spread is null');
    return
  }
  console.log(template,'template');
  
  console.log(dataSource,'dataSource');
  
  const sheet = spread.getActiveSheet();
  InitWorksheet(sheet, dataSource);
  // InitBindPath(spread, template, dataSource)
  // InitSheetRender(spread, template, dataSource, isCompress)
};
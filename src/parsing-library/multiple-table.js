/*
 * @Author: Marlon
 * @Date: 2024-05-14 21:52:13
 * @Description:multipleTable
 */
import { getWorkBook, getInitData } from '../common/store';

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


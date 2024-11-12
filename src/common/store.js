/*
 * @Author: Marlon
 * @Date: 2024-10-26 10:10:06
 * @Description: store getter
 */
import store from 'store';
import { rootWorkBook } from './core';
export const getWorkBook = (GetterQuotationWorkBook) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
  const spread = rootWorkBook._getWorkBook();
  const templateMap = rootWorkBook._getTemplateMap();
  if (templateMap) {
    const sheet = spread.getActiveSheet();
    const activeName = sheet.name();
    if (Object.keys(templateMap).includes(activeName)) {
      return templateMap[activeName];
    }
  }

  if (typeof template === 'object') {
    return template;
  } else {
    return JSON.parse(template);
  }
}

export const getInitData = (GetterQuotationInit) => {
  const quotation = GetterQuotationInit || store.getters['quotationModule/GetterQuotationInit'];
  const active = rootWorkBook._getActiveQuotation();
  if (active) {
    return active;
  }
  return quotation

}

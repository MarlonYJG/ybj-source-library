/*
 * @Author: Marlon
 * @Date: 2024-10-26 10:10:06
 * @Description: store getter
 */
import store from 'store';

export const getWorkBook = (GetterQuotationWorkBook) => {
  return GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
}

export const getInitData = (GetterQuotationInit) => {
  return GetterQuotationInit || store.getters['quotationModule/GetterQuotationInit'];
}

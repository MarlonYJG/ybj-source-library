/*
 * @Author: Marlon
 * @Date: 2024-10-26 10:10:06
 * @Description: store getter
 */
import store from 'store';

export const getSingleWorkBook = (GetterQuotationWorkBook) => {
  return GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
}

export const getSingleInitData = (GetterQuotationInit) => {
  return GetterQuotationInit || store.getters['quotationModule/GetterQuotationInit'];
}

export const getMultipleWorkBook = (GetterMultipleWorkBook) => {
  return GetterMultipleWorkBook || store.getters['quotationModule/GetterMultipleWorkBook'];
}

export const getMultipleInitData = (GetterMultipleInitData) => {
  return GetterMultipleInitData || store.getters['quotationModule/GetterMultipleInitData'];
}
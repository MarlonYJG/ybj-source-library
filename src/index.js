/*
 * @Author: Marlon
 * @Date: 2024-06-28 10:13:41
 * @Description: 
 */
import Template from './template/index';
import { initMobileData, startMobileMode } from './parsing-library/mobile'

import { addSubLimitDiscountInput, addSubLimitDiscountInputType, addSubSortObject, addSubExportError, addSubSelectEquipmentImage } from './common/dep';
import { PROJECT_INIT_DATA, QUOTATION_INIT_DATA } from './common/constant';
import { singleTableSyncStore, resourceSort, LogicalProcessing, LogicalAmount, translateSheet } from './common/single-table';
import { PubGetRandomNumber, GetAllTableRange } from './common/public';
import { CombinationTypeBuild, CombinationType } from './common/combination-type';
import { formatterPrice, getShowCostPrice, initDiscountPercentage, initPriceSetField } from './common/parsing-quotation';
import { getTemplateClassType, getProjectNameField, showDiscount, showPriceSet, initTemplateData } from './common/parsing-template';
import { DEFINE_IDENTIFIER_MAP } from './common/identifier-template'
import { SetDataSource, SpreadLocked } from './common/sheetWorkBook';

import { spreadExportExcel, spreadExportPDF, spreadStyleLocked } from './parsing-library/public';
import { FieldBindPath, InitBindValueTop, LogicalTotalCalculationType, Render as ParsingRender, InitTotal, initSingleTable } from './parsing-library/single-table';

import { MENU_DELETE } from './build-library/config';
import { menuTotal } from './build-library/menu';
import { Reset } from './build-library/public';
import { InsertTotal, RenderTotal, Render as BuildRender, InitSheet, updateCellValue, setProjectName } from './build-library/single-table';
import {
  Sort, DeleteProduct, HeadDelete, spreadPrint,
  spreadExportExcel as headSpreadExportExcel,
  spreadExportPDF as headSpreadExportPDF,
  zoom, FormComputedRowField, Repaint, FrozenHead, ShowCostPrice, ShowCostPriceStatus, UpdateDiscount, UpdatePriceSet, StartAutoFitRow
} from './build-library/head';

export let store = null;
export let vue = null;

const YBJSourceLibrary = {
  install(Vue, options) {
    vue = Vue;
    store = options.store;
  }
}

export {
  PROJECT_INIT_DATA,
  MENU_DELETE,
  QUOTATION_INIT_DATA,
  DEFINE_IDENTIFIER_MAP,
  Template,
  initTemplateData,
  initMobileData,
  startMobileMode,
  initSingleTable,
  initPriceSetField,
  initDiscountPercentage,
  showPriceSet,
  showDiscount,
  UpdateDiscount,
  UpdatePriceSet,
  StartAutoFitRow,
  getShowCostPrice,
  GetAllTableRange,
  ShowCostPrice,
  ShowCostPriceStatus,
  setProjectName,
  getProjectNameField,
  singleTableSyncStore,
  resourceSort,
  PubGetRandomNumber,
  LogicalProcessing,
  LogicalAmount,
  menuTotal,
  CombinationTypeBuild,
  CombinationType,
  formatterPrice,
  SetDataSource,
  Reset,
  InsertTotal,
  RenderTotal,
  BuildRender,
  ParsingRender,
  InitSheet,
  updateCellValue,
  SpreadLocked,
  spreadExportExcel,
  spreadExportPDF,
  spreadStyleLocked,
  FieldBindPath,
  InitBindValueTop,
  LogicalTotalCalculationType,
  InitTotal,
  translateSheet,
  Sort,
  DeleteProduct,
  headSpreadExportExcel,
  headSpreadExportPDF,
  zoom,
  FormComputedRowField,
  Repaint,
  HeadDelete,
  spreadPrint,
  FrozenHead,
  getTemplateClassType,
  addSubLimitDiscountInput,
  addSubLimitDiscountInputType,
  addSubSortObject,
  addSubExportError,
  addSubSelectEquipmentImage
}

export default YBJSourceLibrary

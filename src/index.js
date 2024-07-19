/*
 * @Author: Marlon
 * @Date: 2024-06-28 10:13:41
 * @Description: 
 */

import { PROJECT_INIT_DATA } from './common/constant';
import { singleTableSyncStore, resourceSort } from './common/single-table';
import { MENU_DELETE } from './build-library/config';
import { QUOTATION_INIT_DATA } from './common/constant';
import { PubGetRandomNumber } from './common/public';
import { LogicalProcessing, LogicalAmount } from './common/single-table';
import { menuTotal } from './build-library/menu';
import { CombinationTypeBuild, CombinationType } from './common/combination-type';
import { formatterPrice } from './common/parsing-quotation';
import { getTemplateClassType } from './common/parsing-template';
import { getProjectCfg, getProjectNumberRuler } from './common/single-table';
import { SetDataSource } from './common/sheetWorkBook';
import { Reset } from './build-library/public';
import { InsertTotal, RenderTotal, Render as BuildRender, InitSheet, updateCellValue } from './build-library/single-table';
import { SpreadLocked } from './common/sheetWorkBook';
import { spreadExportExcel, spreadExportPDF, spreadStyleLocked } from './parsing-library/public';
import { FieldBindPath, InitBindValueTop, LogicalTotalCalculationType, Render as ParsingRender, InitTotal } from './parsing-library/single-table';
import { translateSheet } from './common/single-table';
import { DEFINE_IDENTIFIER_MAP } from './common/identifier-template'

import {
  Sort, DeleteProduct, HeadDelete, spreadPrint,
  spreadExportExcel as headSpreadExportExcel,
  spreadExportPDF as headSpreadExportPDF,
  zoom, FormComputedRowField, Repaint, FrozenHead
} from './build-library/head';
// import { GetAllTableRange } from './common/public';


const YBJSourceLibrary = {
  install(Vue, options) {
    console.log('---------------------------');
    console.log(Vue, options);
    console.log('----------------------------');
    // options.store.commit(`quotationModule/${options.modules.quotationModule.mutationType.TEST}`, 1231231);
  }
}

export {
  PROJECT_INIT_DATA,
  MENU_DELETE,
  QUOTATION_INIT_DATA,
  DEFINE_IDENTIFIER_MAP,
  singleTableSyncStore,
  resourceSort,
  PubGetRandomNumber,
  LogicalProcessing,
  LogicalAmount,
  menuTotal,
  CombinationTypeBuild,
  CombinationType,
  formatterPrice,
  getProjectCfg,
  getProjectNumberRuler,
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
  getTemplateClassType
}

export default YBJSourceLibrary

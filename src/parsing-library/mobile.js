/*
 * @Author: Marlon
 * @Date: 2024-09-05 10:49:19
 * @Description: mobile
 */
import _ from '../lib/lodash/lodash.min.js'
import * as GC from '@grapecity/spread-sheets'
import { CombinationType } from '../common/combination-type'
import { singleTableSyncStore } from '../common/single-table';

/**
 * Initialize data source for mobile mode
 * @param {*} dataSource 
 * @param {*} template 
 * @returns 
 */
export const initMobileData = (dataSource, template) => {
  if (!template) return;
  if (!dataSource) return;
  // 初始化报价单数据
  if (dataSource) {
    singleTableSyncStore(dataSource, 'mobile');
  }

  // 初始化模板
  if (template) {
    CombinationType(dataSource, template)
    template.sheets = { sheet: _.cloneDeep(template.cloudSheet.sheets['会场']) }
  }

  return {
    quotation: dataSource,
    template
  }
}

/**
 * Start mobile mode
 * @param {*} spread 
 * @returns 
 */
export const startMobileMode = (spread) => {
  if (!spread) return
  spread.options.useTouchLayout = true;
  spread.options.protectionOptions = {
    allowSelectLockedCells: false,
    allowSelectUnlockedCells: false,
  };
  spread.options.allowInvalidFormula = true;
  spread.options.scrollbarAppearance = GC.Spread.Sheets.ScrollbarAppearance.mobile;
  spread.options.allowInvalidFormula = true;
  spread.options.allowUndo = true;
  spread.options.allowContextMenu = false;
  spread.options.newTabVisible = false;
  spread.options.scrollByPixel = true;
};


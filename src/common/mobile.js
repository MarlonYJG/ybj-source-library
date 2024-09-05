/*
 * @Author: Marlon
 * @Date: 2024-09-05 10:49:19
 * @Description: mobile
 */
import * as GC from '@grapecity/spread-sheets'

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

/*
 * @Author: Marlon
 * @Date: 2024-06-12 10:18:07
 * @Description:
 */
import * as GC from '@grapecity/spread-sheets';
import store from 'store';
import _ from '../lib/lodash/lodash.min.js';
import { imgUrlToBase64, BlobToBase64 } from '../utils/index';
import { SHOW_DRAWER, UPDATE_QUOTATION_PATH, SHOW_SORT, SORT_TYPE_ID, SHOW_DELETE, IGNORE_EVENT } from 'store/quotation/mutation-types';
import { sortObjectProxy } from '../common/proxyData'
import { LayoutRowColBlock } from '../common/core';
import { REGULAR } from '../common/constant';
import { SetDataSource } from '../common/sheetWorkBook';
import { uniqAndSortBy, GetAllTableRange } from '../common/public';

import { buildData, getPositionBlock } from '../common/parsing-quotation';
import { setRowStyle, renderAutoFitRow, AddEquipmentImage, getComputedColumnFormula, getImageField } from '../common/parsing-template';
import { updateEquipmentImage } from '../common/update-quotation'

import { mergeSpan, columnComputedValue } from '../common/single-table';

import { Reset, UpdateSort } from './public';

// eslint-disable-next-line no-unused-vars
import { Render, positionBlock, OperationWorkBookSync, UpdateTotalBlock } from './single-table';

let pickRow = null;
let pickCol = null;

/**
 * repaint table style
 * @param {*} sheet
 * @param {*} classId
 * @param {*} row
 * @param {*} rowH
 */
const repaintTableStyle = (sheet, classId, row, rowH) => {
  const spread = sheet.getParent();
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const { equipment } = template.cloudSheet.center;
  const { image = null } = template.cloudSheet;
  let imageObj = {};
  if (image === null) {
    imageObj.height = rowH;
  } else {
    imageObj = image;
  }

  console.log(resourceViews);


  sheet.suspendPaint();
  const computedColumnFormula = getComputedColumnFormula(equipment.bindPath);
  for (let index = 0; index < resourceViews.length; index++) {
    if (resourceViews[index].resourceLibraryId == classId) {
      const table = resourceViews[index].resources;
      for (let i = 0; i < table.length; i++) {
        mergeSpan(sheet, equipment.spans, row + i);
        setRowStyle(sheet, equipment, row + i, imageObj, false);
        columnComputedValue(sheet, equipment, row + i, computedColumnFormula);
        sheet.getCell(row + i, 0).locked(true);
        ((item, startRow) => {
          repaintImage(spread, item, startRow);
        })(table[i], row + i);
      }
    }
  }

  // Adaptive rendering
  if (resourceViews.length) {
    renderAutoFitRow(sheet, equipment, image, quotation, template);
  }
  sheet.resumePaint();
};

/**
 * repaint image
 * @param {*} spread
 * @param {*} rowData
 * @param {*} rowIndex
 */
const repaintImage = (spread, rowData, rowIndex) => {
  const sheet = spread.getActiveSheet();
  sheet.pictures.remove(rowData.imageId);

  if (rowData.imageFile) {
    BlobToBase64(rowData.imageFile, (base64) => {
      AddEquipmentImage(spread, rowData.imageId, base64, rowIndex, true, false, false);
    })
  } else if (rowData.images) {
    let imgUrl = rowData.images;
    if (REGULAR.chineseCharacters.test(imgUrl)) {
      imgUrl = encodeURI(imgUrl);
    }
    imgUrlToBase64(imgUrl, (base64) => {
      AddEquipmentImage(spread, rowData.imageId, base64, rowIndex, true, false, false);
    });
    sheet.repaint();
  } else {
    console.warn('no image field(imageFile/images)', rowData);
  }
};

/**
 * Dynamically insert data (table area)
 * @param {*} context
 * @param {*} options
 * @param {*} buildDataConfig
 */
const tableInsert = (context, options, buildDataConfig, classType) => {
  const sheet = context.getSheetFromName(options.sheetName);
  const table = sheet.tables.findByName(options.tableName);
  const activeRow = options.activeRow;
  const rowH = sheet.getRowHeight(activeRow);

  let tableRange = null;
  let insertIndex = null;
  if (table) {
    tableRange = table.range();
    insertIndex = activeRow - tableRange.row;
  }
  const tableId = options.tableName.split('table')[1];
  if (insertIndex !== null) {
    const { index, count } = buildDataConfig;
    const conferenceHall = buildData(tableId, insertIndex + index, count, classType);
    store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
      path: ['conferenceHall'],
      value: conferenceHall
    });
    SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);
  }

  console.log('00000000000000');

  console.log(store.getters['quotationModule/GetterQuotationInit']);

  if (tableRange) {
    repaintTableStyle(sheet, tableId, tableRange.row, rowH);
  }
  positionBlock(sheet);
};
/**
 * Dynamically insert data (classification area)
 * @param {*} context
 * @param {*} options
 * @param {*} param2
 */
// eslint-disable-next-line no-unused-vars
const leavelInsert = (context, options, { index, count }) => {
  // eslint-disable-next-line no-unused-vars
  const sheet = context.getSheetFromName(options.sheetName);
};

/**
 * Determine whether the number of products selected under the current category is equal to the number of original products
 * @param {*} tableId
 * @param {*} count
 * @returns
 */
const isRemoveLeavel = (tableId, count) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  if (resourceViews && resourceViews.length) {
    const tables = resourceViews.filter((item) => item.resourceLibraryId == tableId);
    if (tables.length) {
      const resources = tables[0].resources;
      if (count === resources.length) {
        return true;
      } else {
        return false;
      }
    }
  }
};

/**
 * Whether it is on the same floor
 * @param {*} sourceArr
 * @param {*} targetArr
 * @returns
 */
const isInnerLeavel1 = (sourceArr, targetArr) => {
  // Convert array sourceArr to a Map for efficient lookup
  const mapA = new Map();
  sourceArr.forEach(item => {
    const key = item.row;
    if (!mapA.has(key) || item.rowCount < mapA.get(key).rowCount) {
      mapA.set(key, item);
    }
  });

  // Check if each item in targetArr is present in mapA
  for (let i = 0; i < targetArr.length; i++) {
    const itemB = targetArr[i];
    const keyB = itemB.row;
    if (!mapA.has(keyB) || itemB.rowCount > mapA.get(keyB).rowCount) {
      return false;
    }
  }

  return true;
};

/**
 * Determine whether the selected intervals are all in the table
 * @param {*} sheet
 * @param {*} selectRanges
 * @returns
 */
const isInnerTable = (sheet, selectRanges) => {
  const tablesRange = Object.values(GetAllTableRange(sheet));
  console.log(tablesRange);
  const isInner = [];

  for (let i = 0; i < selectRanges.length; i++) {
    const range = selectRanges[i];
    let proxy = null;
    if (range.col === -1) {
      // 代理区间
      proxy = { row: range.row, col: 1, rowCount: range.rowCount, colCount: 1 };
    } else {
      proxy = { row: range.row, col: 1, rowCount: range.rowCount, colCount: 1 };
    }
    const proxyRange = new GC.Spread.Sheets.Range(proxy.row, proxy.col, proxy.rowCount, proxy.colCount);

    isInner.push(isInnerTableList(tablesRange, proxyRange));
  }

  console.log(isInner, ';isInner');

  if (isInner.includes(false)) {
    return false;
  } else {
    return true;
  }
};

/**
 * Determine whether the current interval belongs to the table set
 * @param {*} tablesRange
 * @param {*} targetRange
 * @returns
 */
const isInnerTableList = (tablesRange, targetRange) => {
  for (let j = 0; j < tablesRange.length; j++) {
    if (tablesRange[j].containsRange(targetRange) || tablesRange[j].equals(targetRange)) {
      return true;
    }
  }
  return false;
};

/**
 * Menu Name Modification (Image)
 * @param {*} spread 
 * @param {*} row 
 * @param {*} col 
 * @param {*} table 
 * @param {*} insertImg 
 */
const menuItemImgName = (spread, row, col, table, insertImg) => {
  const tableId = table.name().split('table')[1];
  const layout = new LayoutRowColBlock(spread);
  const proItem = layout.getProductByActiveCell(row, col, tableId);
  if (proItem && proItem.images) {
    insertImg[1].text = '替换图片';
  }
};

/**
 * on open menu
 * @param {*} spread
 */
export const onOpenMenu = (spread) => {
  const designerOnOpenMenu = spread.contextMenu.onOpenMenu;
  spread.contextMenu.onOpenMenu = function (menuData, itemsDataForShown, hitInfo, workbook) {
    itemsDataForShown.splice(0, itemsDataForShown.length);
    const sheet = workbook.getActiveSheet();
    const pubMenu = [
      {
        text: '复制',
        name: 'gc.spread.copy',
        command: 'gc.spread.contextMenu.copy',
        iconClass: 'gc-spread-copy',
        workArea: 'viewportcolHeaderrowHeaderslicercornerpivotTabletimeline'
      },
      {
        command: 'gc.spread.contextMenu.pasteValues',
        name: 'gc.spread.pasteValues',
        iconClass: 'gc-spread-pasteOptions',
        text: '粘贴',
        workArea: 'viewportcolHeaderrowHeadercorner'
      }
    ];
    const insertRow = [
      {
        type: 'separator'
      },
      {
        text: '增加行',
        name: 'addRow',
        iconClass: 'gc-spread-tableInsertRowsBelow',
        command: 'addRow',
        workArea: 'viewportcolHeaderrowHeaderslicercornertable'
      },
      {
        text: '增加多行',
        name: 'addRowMore',
        iconClass: 'gc-spread-tableInsertRowsBelow',
        command: 'addRowMore',
        workArea: 'viewportcolHeaderrowHeaderslicercornertable'
      },
      {
        text: '插入行',
        name: 'insertRow',
        iconClass: 'gc-spread-tableInsertRowsAbove',
        command: 'insertRow',
        workArea: 'viewportcolHeaderrowHeaderslicercornertable'
      },
      {
        text: '插入多行',
        name: 'insertRowMore',
        iconClass: 'gc-spread-tableInsertRowsAbove',
        command: 'insertRowMore',
        workArea: 'viewportcolHeaderrowHeaderslicercornertable'
      }
    ];
    const delRow = [
      {
        text: '删除行',
        name: 'delRow',
        iconClass: 'gc-spread-tableDeleteRows',
        command: 'delRow',
        workArea: 'viewportcolHeaderrowHeaderslicercornertable'
      }
    ];
    const removeRowImg = [
      {
        text: '删除图片',
        name: 'removeRowImg',
        iconClass: '',
        command: 'removeRowImg',
        workArea: 'viewportcolHeaderrowHeaderslicercornertable'
      }
    ];
    const addRowPro = [
      {
        type: 'separator'
      },
      {
        text: '添加产品',
        name: 'addpro',
        iconClass: 'gc-spread-insertComment',
        command: 'addpro',
        workArea: 'viewportcolHeaderrowHeaderslicercornertable'
      }
    ];
    const insertRowPro = [
      // {
      //   text: '插入产品',
      //   name: 'insertPro',
      //   iconClass: 'gc-spread-insertComment',
      //   command: 'insertPro',
      //   workArea: 'viewportcolHeaderrowHeaderslicercornertable'
      // }
    ];
    const sortPro = [
      {
        type: 'separator'
      },
      {
        text: '排序',
        name: 'sortPro',
        iconClass: 'gc-spread-sortAscend',
        command: 'sortPro',
        workArea: 'viewportcolHeaderrowHeaderslicercornertable'
      }
    ];
    const insertImg = [
      {
        type: 'separator'
      },
      {
        text: '插入图片',
        name: 'insertImg',
        iconClass: '',
        command: 'insertImg',
        workArea: 'viewportcolHeaderrowHeaderslicercornertable'
      }
    ]

    if (hitInfo && hitInfo.worksheetHitInfo) {
      const { leavel1Area, leavel2Area, subTotalArea } = getPositionBlock();
      const { row, col } = hitInfo.worksheetHitInfo;
      pickRow = row;
      pickCol = col;
      const rowCount = sheet.getRowCount();
      const colCount = sheet.getColumnCount();
      const ranges = sheet.getSelections();

      if (row <= rowCount - 1 && col <= colCount - 1) {
        if (ranges.length) {
          if (ranges.length === 1) {
            const range = ranges[0];

            const rangeIndexs = [];
            for (let index = 0; index < range.rowCount; index++) {
              rangeIndexs.push(range.row + index);
            }

            console.log(rangeIndexs, '框选的行号', row);

            if (rangeIndexs.includes(row)) {
              itemsDataForShown.push(...pubMenu);
              const table = sheet.tables.find(row, col);
              console.log(table, 'table');

              if (table) {
                menuItemImgName(spread, row, col, table, insertImg);
                if (table.range().containsRange(range)) {
                  itemsDataForShown.push(...delRow, ...insertRow, ...addRowPro, ...insertRowPro, ...sortPro, ...insertImg, ...removeRowImg);
                  console.log('点击区域：table');
                } else {
                  itemsDataForShown.push(...delRow, ...addRowPro, ...sortPro, ...insertImg, ...removeRowImg);
                }
              } else if (leavel1Area.includes(row)) {
                itemsDataForShown.push(...delRow, ...addRowPro, ...sortPro);
                console.log('点击区域：leavel 1');
              } else if (leavel2Area.includes(row)) {
                itemsDataForShown.push(...delRow, ...addRowPro, ...sortPro);
                console.log('点击区域：leavel 2');
              } else if (subTotalArea.includes(row)) {
                itemsDataForShown.push(...addRowPro, ...sortPro);
                console.log('点击区域：小计');
              }
            } else {
              const table = sheet.tables.find(row, col);
              menuItemImgName(spread, row, col, table, insertImg);
              itemsDataForShown.push(...removeRowImg, ...insertImg);
            }
          } else {
            itemsDataForShown.push(...pubMenu, ...delRow);
          }
        } else {
          const table = sheet.tables.find(row, col);
          if (table) {
            menuItemImgName(spread, row, col, table, insertImg);
            itemsDataForShown.push(...removeRowImg, ...insertImg);
          }
        }
      }
    }
    designerOnOpenMenu.apply(this, arguments);
  };
};

/**
 * Device picture editing
 * @param {*} spread 
 * @param {*} context 
 * @param {*} options 
 * @param {*} type 
 */
const equipmentImageEdit = (spread, context, options, type) => {
  const sheet = context.getSheetFromName(options.sheetName);
  const table = sheet.tables.find(pickRow, pickCol);
  if (table) {
    const tableId = table.name().split('table')[1];
    const layout = new LayoutRowColBlock(spread);
    const proItem = layout.getProductByActiveCell(pickRow, pickCol, tableId);
    if (proItem) {
      const imgField = getImageField();
      if (imgField) {
        updateEquipmentImage(spread, tableId, proItem, imgField.fieldName, type, pickRow);
      }
    } else {
      console.error('Product not found');
    }
  }
};

/**
 * command register
 * @param {*} spread
 */
export const commandRegister = (spread) => {
  const commandManager = spread.commandManager();
  const commandMap = {
    addpro: {
      canUndo: true,
      name: 'addpro',
      execute: () => {
        store.commit(`quotationModule/${SHOW_DRAWER}`, true);
        return true;
      }
    },
    insertPro: {
      canUndo: true,
      name: 'insertPro',
      execute: () => {
        console.log('插入产品');
        return true;
      }
    },
    sortPro: {
      canUndo: true,
      name: 'sortPro',
      execute: (context, options) => {
        console.log(options);
        const activeRow = options.activeRow;
        const { leavel1Area = [], leavel2Area = [] } = getPositionBlock();
        const sheet = context.getSheetFromName(options.sheetName);

        const table = sheet.tables.find(activeRow, options.activeCol);
        if (leavel2Area.includes(activeRow)) {
          const t = sheet.tables.find(activeRow + 1, options.activeCol);
          if (t) {
            // TODO
            const tableId = t.name().split('table')[1];
            store.commit(`quotationModule/${SORT_TYPE_ID}`, tableId);
            store.commit(`quotationModule/${SHOW_SORT}`, true);
          }
          return true;
        } else if (leavel1Area.includes(activeRow)) {
          const t = sheet.tables.find(activeRow + 1, options.activeCol);
          if (t) {
            const tableId = t.name().split('table')[1];
            store.commit(`quotationModule/${SORT_TYPE_ID}`, tableId);
            store.commit(`quotationModule/${SHOW_SORT}`, true);
          }
          return true;
        } else if (table) {
          const tableId = options.tableName.split('table')[1];
          store.commit(`quotationModule/${SORT_TYPE_ID}`, tableId);
          store.commit(`quotationModule/${SHOW_SORT}`, true);
          return true;
        } else {
          sortObjectProxy.value = null;
        }
        return true;
      }
    },
    addRow: {
      canUndo: true,
      name: 'addRow',
      execute: (context, options, isUndo) => {
        const Commands = GC.Spread.Sheets.Commands;
        const layout = new LayoutRowColBlock(spread);
        const { rows } = layout.getLevel();
        const classType = layout.getClassType();

        if (isUndo) {
          // Commands.undoTransaction(context, options);
        } else {
          Commands.startTransaction(context, options);
          if (rows.includes(options.activeRow)) {
            leavelInsert(context, options, { index: 1, count: 1 });
          } else {
            tableInsert(context, options, { index: 1, count: 1 }, classType);
          }
          Commands.endTransaction(context, options);
        }
        return true;
      }
    },
    addRowMore: {
      canUndo: true,
      name: 'addRowMore',
      execute: (context, options, isUndo) => {
        const Commands = GC.Spread.Sheets.Commands;
        if (isUndo) {
          Commands.undoTransaction(context, options);
        } else {
          Commands.startTransaction(context, options);
          tableInsert(context, options, { index: 1, count: 5 });
          Commands.endTransaction(context, options);
        }
        return true;
      }
    },
    insertRow: {
      canUndo: true,
      name: 'insertRow',
      execute: (context, options, isUndo) => {
        const Commands = GC.Spread.Sheets.Commands;
        if (isUndo) {
          Commands.undoTransaction(context, options);
        } else {
          Commands.startTransaction(context, options);
          tableInsert(context, options, { index: 0, count: 1 });
          Commands.endTransaction(context, options);
        }
        return true;
      }
    },
    insertRowMore: {
      canUndo: true,
      name: 'insertRowMore',
      execute: (context, options, isUndo) => {
        const Commands = GC.Spread.Sheets.Commands;
        if (isUndo) {
          Commands.undoTransaction(context, options);
        } else {
          Commands.startTransaction(context, options);
          tableInsert(context, options, { index: 0, count: 5 });
          Commands.endTransaction(context, options);
        }
        return true;
      }
    },
    delRow: {
      canUndo: true,
      name: 'delRow',
      execute: (context, options, isUndo) => {
        store.commit(`quotationModule/${IGNORE_EVENT}`, true);
        const { leavel1Area, leavel2Area } = getPositionBlock();
        // const Commands = GC.Spread.Sheets.Commands;
        if (isUndo) {
          // Commands.undoTransaction(context, options);
        } else {
          const sheet = context.getSheetFromName(options.sheetName);
          const ranges = sheet.getSelections();
          const rangesMores = [];

          // Single selection (click + box selection)
          if (ranges.length) {
            if (ranges.length === 1) {
              const range = ranges[0];
              let table = sheet.tables.find(range.row, range.col);
              if (range.col === -1) {
                table = sheet.tables.find(range.row, 1);
              }

              if (table) {
                const rangeR = new GC.Spread.Sheets.Range(range.row, 0, range.rowCount, 2);
                if (table.range().containsRange(rangeR)) {
                  const table = sheet.tables.find(range.row, 1);
                  if (table) {
                    const id = table.name().split('table')[1];
                    if (isRemoveLeavel(id, range.rowCount)) {
                      deleteClassRowSingle(sheet, table, leavel1Area, leavel2Area, null);
                    } else {
                      deleteTableRowSingle(sheet, range);
                    }
                  }
                }
              } else if (leavel2Area.length) {
                // TODO
                console.log('le2 内区域');
              } else if (leavel1Area.length) {
                if (leavel1Area.includes(range.row) && range.rowCount === 1) {
                  const table = sheet.tables.find(range.row + 1, 2);
                  if (table) {
                    // 删除分类及其产品
                    deleteClassRowSingle(sheet, table, leavel1Area, leavel2Area, range);
                  } else {
                    store.commit(`quotationModule/${SHOW_DELETE}`, true);
                  }
                } else {
                  store.commit(`quotationModule/${SHOW_DELETE}`, true);
                }
              } else {
                store.commit(`quotationModule/${SHOW_DELETE}`, true);
              }
            } else {
              for (let index = 0; index < ranges.length; index++) {
                const range = ranges[index];
                rangesMores.push({
                  row: range.row,
                  rowCount: range.rowCount
                });
              }
            }
          }

          // Check (click + box select)
          if (rangesMores.length) {
            if (leavel2Area.length) {
              // TODO 如果存在二级分类，则必须复选的所有行都必须属于二级分类
            } else if (leavel1Area.length) {
              // 如果复选的所有值 都属于 分类集合中，则删除对应的分类
              const leavel1 = [];
              for (let index = 0; index < leavel1Area.length; index++) {
                leavel1.push({
                  row: leavel1Area[index],
                  rowCount: 1
                });
              }

              if (isInnerLeavel1(leavel1, rangesMores)) {
                deleteClassRowMulti(sheet, rangesMores);
              } else if (isInnerTable(sheet, rangesMores)) {
                deleteTableRowMulti(sheet, rangesMores);
              } else {
                store.commit(`quotationModule/${SHOW_DELETE}`, true);
              }
            } else {
              // 无分类 如果复选的所有值 都属于 table集合中，则删除对应的table 行
              if (isInnerTable(sheet, rangesMores)) {
                deleteTableRowMulti(sheet, rangesMores);
              } else {
                store.commit(`quotationModule/${SHOW_DELETE}`, true);
              }
            }
          }
        }
        store.commit(`quotationModule/${IGNORE_EVENT}`, false);
        return true;
      }
    },
    insertImg: {
      canUndo: true,
      name: 'insertImg',
      execute: (context, options, isUndo) => {
        const Commands = GC.Spread.Sheets.Commands;
        if (isUndo) {
          Commands.undoTransaction(context, options);
        } else {
          Commands.startTransaction(context, options);
          equipmentImageEdit(spread, context, options, 'insert')
          Commands.endTransaction(context, options);
        }
      }
    },
    removeRowImg: {
      canUndo: true,
      name: 'removeRowImg',
      execute: (context, options, isUndo) => {
        const Commands = GC.Spread.Sheets.Commands;
        if (isUndo) {
          Commands.undoTransaction(context, options);
        } else {
          Commands.startTransaction(context, options);
          equipmentImageEdit(spread, context, options, 'del');
          Commands.endTransaction(context, options);
        }
      }
    }
  };
  for (const key in commandMap) {
    if (Object.hasOwnProperty.call(commandMap, key)) {
      commandManager.register(key, commandMap[key]);
    }
  }
};

/**
 * Delete the corresponding product category in the quotation and update the sheet data source
 * @param {*} sheet
 * @param {*} table
 * @param {*} leavel1Area
 * @param {*} leavel2Area
 * @param {*} srange
 */
const deleteClassRowSingle = (sheet, table, leavel1Area, leavel2Area, srange) => {
  const id = table.name().split('table')[1];
  const { rowCount, row } = table.range();

  sheet.suspendPaint();
  if (!srange) {
    if (leavel2Area.length) {
      sheet.deleteRows(row - 2, rowCount + 3);
    } else if (leavel1Area.length) {
      sheet.deleteRows(row - 1, rowCount + 2);
    } else {
      sheet.deleteRows(row, rowCount);
    }
  } else {
    sheet.deleteRows(srange.row, rowCount + 2);
  }
  sheet.resumePaint();

  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const conferenceHall = _.cloneDeep(quotation.conferenceHall);
  const resourceViews = conferenceHall.resourceViews;
  const resourceViewsMap = conferenceHall.resourceViewsMap;

  resourceViews.forEach((le1, i) => {
    if (le1.resourceLibraryId == id) {
      resourceViews.splice(i, 1);
    }
  });

  for (const key in resourceViewsMap) {
    if (Object.hasOwnProperty.call(resourceViewsMap, key)) {
      if (key == id) {
        delete resourceViewsMap[key];
      }
    }
  }

  UpdateSort(conferenceHall);

  store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
    path: ['conferenceHall'],
    value: conferenceHall
  });
  SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);

  UpdateTotalBlock(sheet);
  positionBlock(sheet);
};

/**
 * Delete the corresponding product in the quotation and update the sheet data source
 * @param {*} sheet
 * @param {*} srange
 */
const deleteTableRowSingle = (sheet, srange) => {
  const se = [{ row: srange.row, rowCount: srange.rowCount }];
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const conferenceHall = _.cloneDeep(quotation.conferenceHall);

  const selectedResource = getSelectionsRowData(sheet, se);
  deleteSelectionsRowData(selectedResource, conferenceHall);

  sheet.suspendPaint();
  sheet.deleteRows(srange.row, srange.rowCount);
  sheet.resumePaint();

  store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
    path: ['conferenceHall'],
    value: conferenceHall
  });
  SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);

  UpdateTotalBlock(sheet);
  positionBlock(sheet);
};

/**
 * Delete multiple classifications
 * @param {*} sheet
 * @param {*} srange
 */
const deleteClassRowMulti = (sheet, srange) => {
  const spread = sheet.getParent();
  const removeRange = uniqAndSortBy(srange, item => item.row, 'row');
  const tableIds = [];
  for (let i = 0; i < removeRange.length; i++) {
    const table = sheet.tables.find(removeRange[i].row + 1, 1);
    if (table) {
      const id = table.name().split('table')[1];
      tableIds.push(id);
    }
  }

  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const conferenceHall = _.cloneDeep(quotation.conferenceHall);
  const sourceData = conferenceHall.resourceViews;
  const resourceViews = [];
  const resourceViewsMap = {};

  for (let i = 0; i < sourceData.length; i++) {
    if (!tableIds.includes(sourceData[i].resourceLibraryId)) {
      resourceViews.push(sourceData[i]);
    }
  }
  resourceViews.forEach((item) => {
    resourceViewsMap[item.resourceLibraryId] = item;
  });

  conferenceHall.resourceViews = resourceViews;
  conferenceHall.resourceViewsMap = resourceViewsMap;

  UpdateSort(conferenceHall);

  Reset(spread);

  store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
    path: ['conferenceHall'],
    value: conferenceHall
  });
  SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);

  Render(spread);
};

/**
 * Deleting products in multiple tables
 * @param {*} sheet
 * @param {*} srange
 */
const deleteTableRowMulti = (sheet, srange) => {
  const spread = sheet.getParent();
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const conferenceHall = _.cloneDeep(quotation.conferenceHall);

  const selectedResource = getSelectionsRowData(sheet, srange);
  deleteSelectionsRowData(selectedResource, conferenceHall);

  Reset(spread);

  store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
    path: ['conferenceHall'],
    value: conferenceHall
  });
  SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);

  Render(spread);
};

/**
 * Get selected row data
 * @param {*} sheet
 * @param {*} srange
 * @returns
 */
const getSelectionsRowData = (sheet, srange) => {
  const selectedRange = uniqAndSortBy(srange, item => item.row, 'row');
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const conferenceHall = quotation.conferenceHall;
  const resourceViews = conferenceHall.resourceViews;
  const selectedResource = [];

  for (let i = 0; i < selectedRange.length; i++) {
    const table = sheet.tables.find(selectedRange[i].row, 1);
    if (table) {
      const id = table.name().split('table')[1];
      const { row } = table.range();

      resourceViews.forEach((el1) => {
        if (el1.resourceLibraryId == id) {
          const resources = el1.resources;
          if (resources && resources.length) {
            const start = selectedRange[i].row - row;
            const n = selectedRange[i].rowCount;
            selectedResource.push(...resources.slice(start, start + n));
          }
        }
      });
    }
  }
  return selectedResource;
};

/**
 * Delete selected product rows
 * @param {*} delResource
 * @param {*} conferenceHall
 */
const deleteSelectionsRowData = (delResource, conferenceHall) => {
  const resourceViews = conferenceHall.resourceViews;
  const resourceViewsMap = {};

  const delIds = delResource.map((item) => item.id);
  for (let i = 0; i < resourceViews.length; i++) {
    const resources = resourceViews[i].resources;
    const newResource = [];
    for (let j = 0; j < resources.length; j++) {
      if (!delIds.includes(resources[j].id)) {
        newResource.push(resources[j]);
      }
    }
    resourceViews[i].resources = newResource;
  }

  for (let i = 0; i < resourceViews.length; i++) {
    if (!resourceViews[i].resources || !resourceViews[i].resources.length) {
      resourceViews.splice(i, 1);
    }
  }

  resourceViews.forEach(item => {
    resourceViewsMap[item.resourceLibraryId] = item;
  });
  conferenceHall.resourceViewsMap = resourceViewsMap;

  UpdateSort(conferenceHall);
};

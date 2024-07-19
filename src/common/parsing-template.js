/*
 * @Author: Marlon
 * @Date: 2024-05-17 14:31:48
 * @Description:parsing template
 */
import * as GC from '@grapecity/spread-sheets';
import _ from 'lodash';
import store from 'store';

import { updateFormula } from './public'
import { DEFINE_IDENTIFIER_MAP } from './identifier-template'

/**
 * Insert the start of the column of the table
 */
export const PubGetTableStartColumnIndex = () => {
  return 0;
};
/**
 * The total number of columns inserted into the table
 * @returns
 */
export const PubGetTableColumnCount = () => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  return template.cloudSheet.center.columnCount;
};

/**
 * Get the template top area of the row and column
 * @returns 
 */
export const getTemplateTopRowCol = () => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  if (template.cloudSheet.top) {
    return {
      rowCount: template.cloudSheet.top.rowCount,
      columnCount: template.cloudSheet.top.columnCount
    };
  }
  return {
    rowCount: 0,
    columnCount: 0
  }
}

/**
 * Template render markers
 */
export const templateRenderFlag = () => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const { mixRender = false, classType = null, isHaveChild = false } = template;

  return {
    template,
    mixRender,
    classType,
    isHaveChild
  };
};

/**
 * Delete table header
 * @param {*} sheet
 * @param {*} dataTable
 * @param {*} top
 */
export const delTableHeaderRowCount = (sheet, dataTable, top) => {
  let delRowCount = 0;
  if (Object.keys(dataTable).length > 1) {
    delRowCount = top.rowCount - (Object.keys(dataTable).length - 1);
  }
  if (delRowCount) {
    sheet.deleteRows(top.rowCount, delRowCount);
  }
};

/**
 * 获取表头配置信息
 * @param {*} type 
 * @param {*} isDel 
 * @returns 
 */
export const getTableHeaderDataTable = (type, isDel) => {
  const header = [];
  if (Object.keys(type.dataTable).length > 1) {
    for (const key in type.dataTable) {
      if (Object.hasOwnProperty.call(type.dataTable, key)) {
        if (Number(key) >= 1) {
          header.push(_.cloneDeep(type.dataTable[key]));
          isDel && delete type.dataTable[key];
        }
      }
    }
  }
  return header;
};

/**
 * Merge Row
 * @param {*} sheet
 * @param {*} rows
 * @param {*} column
 * @param {*} rowCount
 * @param {*} columnCount
 */
export const mergeRow = (sheet, rows, column = 0, rowCount = 1, columnCount = 1) => {
  if (rows.length) {
    for (let index = 0; index < rows.length; index++) {
      sheet.addSpan(rows[index], column, rowCount, columnCount);
    }
  }
};

/**
 * Style the cell
 * @param {*} sheet
 * @param {*} row
 * @param {*} startRowIndex
 * @param {*} locked
 */
export const setCell = (sheet, row, startRowIndex, locked = false) => {
  for (const j in row) {
    if (Object.hasOwnProperty.call(row, j)) {
      const column = _.cloneDeep(row[j]);
      column.style || (column.style = {});
      column.style.locked = locked;
      const style = new GC.Spread.Sheets.Style();
      style.fromJSON(column.style);
      sheet.setStyle(startRowIndex, Number(j), style, GC.Spread.Sheets.SheetArea.viewport);
      if (column.value === 0 || column.value) {
        sheet.setValue(startRowIndex, Number(j), column.value);
      }
    }
  }
};

/**
 * Set all cell heights
 * @param {*} sheet
 * @param {*} field
 * @param {*} row
 */
export const PubSetCellHeight = (sheet, field, row) => {
  for (const i in field.dataTable) {
    if (Object.hasOwnProperty.call(field.dataTable, i)) {
      if (field.height) {
        sheet.setRowHeight(row + Number(i), field.height);
      }
    }
  }
};

/**
 * Style the rows
 * @param {*} sheet
 * @param {*} rowsField
 * @param {*} startRow
 * @param {*} image
 * @param {*} locked
 */
export const setRowStyle = (sheet, rowsField, startRow, image, locked = false) => {
  if (rowsField.dataTable) {
    for (const i in rowsField.dataTable) {
      if (Object.hasOwnProperty.call(rowsField.dataTable, i)) {
        const row = rowsField.dataTable[i];
        for (const j in row) {
          if (Object.hasOwnProperty.call(row, j)) {
            const column = _.cloneDeep(row[j]);
            if (!column.style) {
              column.style = {};
            }
            column.style.locked = locked;
            const style = new GC.Spread.Sheets.Style();
            style.fromJSON(column.style);
            sheet.setStyle(startRow + Number(i), Number(j), style, GC.Spread.Sheets.SheetArea.viewport);
            if (column.value === 0 || column.value) {
              sheet.setValue(startRow + Number(i), Number(j), column.value);
            }
          }
        }

        if (image && image.height && rowsField.height) {
          const maxH = Math.max(image.height, rowsField.height);
          sheet.setRowHeight(startRow + Number(i), maxH);
        } else if (image) {
          image.height && sheet.setRowHeight(startRow + Number(i), image.height);
        } else {
          rowsField.height && sheet.setRowHeight(startRow + Number(i), rowsField.height);
        }
      }
    }
  }
};

/**
 * Merge Column
 * @param {*} sheet
 * @param {*} field
 * @param {*} row
 * @param {*} rowCount
 * @param {*} columnCount
 * @param {*} columnName
 */
export const mergeColumn = (sheet, field, row, rowCount, columnCount = 1, columnName) => {
  for (const key in field) {
    if (Object.hasOwnProperty.call(field, key)) {
      const bindPath = field[key].bindPath;
      if (bindPath) {
        if (columnName) {
          if (bindPath === columnName) {
            sheet.addSpan(row, field[key].column, rowCount, columnCount);
          }
        } else {
          if (['classname', 'pname'].includes(bindPath)) {
            sheet.addSpan(row, field[key].column, rowCount, columnCount);
          }
        }
      }
    }
  }
};

/**
 * Add an image (equipment)
 * @param {*} spread
 * @param {*} pictureName
 * @param {*} base64
 * @param {*} startRow
 * @param {*} allowMove
 * @param {*} allowResize
 * @param {*} isLocked
 */
export const AddEquipmentImage = (spread, pictureName, base64, startRow, allowMove = false, allowResize = true, isLocked = true) => {
  const sheet = spread.getActiveSheet();
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  let x = 2;
  let y = 2;
  let curWidth = 100;
  let curHeight = 100;
  const { image = null } = template.cloudSheet;
  if (image) {
    const { imgX = 2, imgY = 2, width = 100, height = 100 } = image;
    x = imgX;
    y = imgY;
    curWidth = width;
    curHeight = height;

    const picture = sheet.pictures.add(pictureName, base64, x, y, curWidth, curHeight);

    picture.startRow(startRow);
    picture.startColumn(image.column);
    picture.allowMove(allowMove);
    picture.allowResize(allowResize);
    picture.isLocked(isLocked);
  }
};

/**
 * Show total
 * @param {*} total
 */
export const showTotal = () => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const total = template.cloudSheet.total;
  let show = true;
  if (total) {
    if (!total.showHide) {
      if (_.has(total, ['show'])) {
        show = total.show;// new flag for showHide
      }
    } else {
      if (total.showHide === 'hide') {// old flag for showHide
        show = false;
      } else {
        show = true;
      }
    }
  } else {
    show = false;
  }
  return show;
};

/**
 * Show subTotal
 * @returns 
 */
export const showSubTotal = () => {
  let show = true;
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const subTotal = template.cloudSheet.center.total;
  if (showTotal()) {
    if (subTotal) {
      if (_.has(subTotal, ['show'])) {
        show = subTotal.show;
      }
    } else {
      show = false;
    }
  } else {
    show = false;
  }
  return show;
}

/**
 * Get computed column formula
 * @param {*} bindPath 
 * @returns 
 */
export const getComputedColumnFormula = (bindPath) => {
  const computedColumnFormula = {};
  const computedColFormula = {}
  for (const key in bindPath) {
    if (Object.hasOwnProperty.call(bindPath, key)) {
      if (bindPath[key].formula) {
        computedColumnFormula[key] = bindPath[key];
      }
    }
  }

  for (const key in computedColumnFormula) {
    if (Object.hasOwnProperty.call(computedColumnFormula, key)) {
      computedColFormula[key] = {
        ...getFormulaFieldRowCol(computedColumnFormula[key]),
        formula: updateFormula(computedColumnFormula[key].formula.formula)
      }
    }
  }

  return computedColFormula
}

/**
 * Get the row and column information of a formula field
 * @param {*} field 
 * @returns 
 */
export const getFormulaFieldRowCol = (field) => {
  let formulaField = {}
  if (field.formula) {
    formulaField = {
      ...field,
      ...field.formula,
    }
    if (field.column !== field.formula.column) {
      console.error('当前字段行列数据不统一【模板错误】', field);
    }
  } else {
    console.warn('当前字段不是公式字段', field);
  }

  return formulaField
}

/**
 * Obtain the classification type of the current template
 * @returns 
 */
export const getTemplateClassType = () => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  if (template.templateClassIdentifier) {
    if (Object.keys(DEFINE_IDENTIFIER_MAP).includes(template.templateClassIdentifier)) {
      return DEFINE_IDENTIFIER_MAP[template.templateClassIdentifier].identifier;
    } else {
      return null;
    }
  } else {
    return null
  }

}
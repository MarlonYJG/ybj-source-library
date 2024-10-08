/*
 * @Author: Marlon
 * @Date: 2024-05-17 14:31:48
 * @Description:parsing template
 */
import * as GC from '@grapecity/spread-sheets';
import _ from '../lib/lodash/lodash.min.js';
import store from 'store';

import { updateFormula } from './public';
import { DEFINE_IDENTIFIER_MAP } from './identifier-template';
import { PRICE_SET_MAP } from "./constant";
import { LogicalProcessing, setAutoFitRow, defaultAutoFitRow } from './single-table';
import { CombinationType } from './combination-type';
import { getConfig } from './parsing-quotation';

/**
 * Get the template
 * @returns 
 */
const getTemplate = () => {
  return store.getters['quotationModule/GetterQuotationWorkBook'];
}

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
export const PubGetTableColumnCount = (GetterQuotationWorkBook) => {
  const template = GetterQuotationWorkBook || getTemplate();
  return template.cloudSheet.center.columnCount;
};

/**
 * Get the template top area of the row and column
 * @returns 
 */
export const getTemplateTopRowCol = () => {
  const template = getTemplate();
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
export const templateRenderFlag = (GetterQuotationWorkBook) => {
  const template = GetterQuotationWorkBook || getTemplate();
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
 * @param {*} quotation 
 */
export const setRowStyle = (sheet, rowsField, startRow, image, locked = false, quotation = null) => {
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
        const config = getConfig(quotation);
        if (config && config.startAutoFitRow) {
          sheet.getCell(startRow + Number(i), -1).wordWrap(true);
          sheet.autoFitRow(startRow + Number(i))
          setAutoFitRow(sheet, startRow + Number(i), rowsField, image)
        } else {
          defaultAutoFitRow(sheet, startRow + Number(i), rowsField, image);
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
export const AddEquipmentImage = (spread, pictureName, base64, startRow, allowMove = false, allowResize = true, isLocked = true, GetterQuotationWorkBook) => {
  const sheet = spread.getActiveSheet();
  const template = GetterQuotationWorkBook || getTemplate();
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
export const showTotal = (GetterQuotationWorkBook) => {
  const template = GetterQuotationWorkBook || getTemplate();
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
export const showSubTotal = (GetterQuotationWorkBook) => {
  let show = true;
  const template = GetterQuotationWorkBook || getTemplate();
  const subTotal = template.cloudSheet.center.total;
  if (showTotal(GetterQuotationWorkBook)) {
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
    if (field.columnHeader !== field.formula.columnHeader) {
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
  const template = getTemplate().cloudSheet;
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

/**
 * Determines whether the template has multiple headers or a single header
 * @returns 
 */
export const isMultiHead = () => {
  const template = getTemplate().cloudSheet;
  const regex = /title@/;
  if (template.templateClassIdentifier && regex.test(template.templateClassIdentifier)) {
    return true;
  }
  return false;
}

/**
 * Whether it is a Level_1_row@Level_2_row type
 * @returns 
 */
export const isL1L2RowMerge = () => {
  const template = getTemplate().cloudSheet;
  if (template.templateClassIdentifier === 'Level_1_row@Level_2_row') {
    return true;
  }
  return false;
}


/**
 * Obtain the project name field in the template
 * @returns 
 */
export const getProjectNameField = () => {
  const template = getTemplate();
  const projectNameField = template.cloudSheet.top.bindPath.quotation.name;
  return projectNameField;
}

/**
 * Get the field where the discount is calculated
 * @returns 
 */
export const getDiscountField = () => {
  const template = getTemplate();
  const discountField = template.cloudSheet.center.equipment.bindPath;
  if (Object.keys(discountField).includes('discountUnitPrice') || Object.keys(discountField).includes('unitPrice')) {
    if (Object.keys(discountField).includes('discountUnitPrice') && Object.keys(discountField).includes('unitPrice')) {
      console.warn('The discount field is ambiguous');
    }
    if (Object.keys(discountField).includes('discountUnitPrice')) {
      return discountField.discountUnitPrice.bindPath || 'discountUnitPrice';
    } else if (Object.keys(discountField).includes('unitPrice')) {
      return discountField.unitPrice.bindPath || 'unitPrice';
    }
  } else {
    console.error('The discount field does not exist');
  }
  return null;
}

/**
 * Whether or not to display discounts
 * @param {*} template 
 * @returns 
 */
export const showDiscount = (temp) => {
  const template = temp || getTemplate();
  if (template) {
    const discountField = template.cloudSheet.center.equipment.bindPath;
    if (Object.keys(discountField).includes('discountUnitPrice')) {
      return discountField.discountUnitPrice || false;
    }
  }
  return false;
}

/**
 * Whether or not to display the price settings
 * @param {*} template 
 * @returns 
 */
export const showPriceSet = (temp) => {
  const template = temp || getTemplate();
  if (template) {
    return !!showDiscount(template);
  }
  console.warn('The price setup field does not exist');
  return false;
}

/**
 * Get the Unit Price column in the template
 * @returns 
 */
export const getPriceColumn = () => {
  const template = getTemplate();
  let priceFieldCol = null;
  const binds = template.cloudSheet.center.equipment.bindPath;
  for (const key in binds) {
    if (Object.hasOwnProperty.call(binds, key)) {
      if ([].concat(Object.keys(PRICE_SET_MAP), ['discountUnitPrice']).includes(key)) {
        priceFieldCol = binds[key].column;
      }
    }
  }
  return priceFieldCol;
}

/**
 * Get paths in the template
 * @returns 
 */
export const getPaths = () => {
  const topPath = ['cloudSheet', 'top', 'bindPath', 'quotation'];
  const bottomPath = ['cloudSheet', 'bottom', 'bindPath', 'quotation'];
  const conferenceHallTopPath = ['cloudSheet', 'top', 'bindPath', 'conferenceHall'];
  const conferenceHallBottomPath = ['cloudSheet', 'bottom', 'bindPath', 'conferenceHall'];
  return {
    topPath,
    bottomPath,
    conferenceHallTopPath,
    conferenceHallBottomPath
  };
}

/**
 * Initialize the template data
 * @param {*} templateJSON 
 * @param {*} quotation 
 * @param {*} quaLogos 
 * @returns 
 */
export const initTemplateData = (templateJSON, quotation = null, quaLogos = []) => {
  const template = templateJSON.excelJson;

  if (templateJSON.excelJson.cloudSheet.quaLogos && templateJSON.excelJson.cloudSheet.quaLogos.length) {
    const logos = templateJSON.excelJson.cloudSheet.quaLogos;
    logos.forEach((item, i) => {
      item.url = quaLogos[i].url;
      item.id = quaLogos[i].id;
    })
    templateJSON.excelJson.cloudSheet.quaLogos = logos;
  }

  if (quotation && quotation.templateId == templateJSON.id) {
    LogicalProcessing(quotation, templateJSON.excelJson);
    if (template.mixRender && !quotation.remark) {
      quotation.remark = '本表为预估报价，最终结算以实际产生的费用为准。';
    }
  }

  CombinationType(quotation, template);

  return {
    templateJSON,
    quotation
  }
}

/**
 * Obtain the field name of the corresponding image in the template
 * @param {*} template 
 * @returns 
 */
export const getImageField = (template) => {
  if (!template) {
    template = getTemplate();
  }
  const imgField = template.cloudSheet.center.equipment.bindPath.img;
  const imageField = template.cloudSheet.center.equipment.bindPath.imageId;
  if (imgField) {
    return {
      ...imgField,
      fieldName: 'img'
    }
  }
  if (imageField) {
    return {
      ...imageField,
      fieldName: 'imageId'
    }
  }
  console.warn('The image field(img||imageId) does not exist');
  return null;
}

/**
 * Obtain the image configuration in the template
 * @param {*} template 
 * @returns 
 */
export const getImageConfig = (template) => {
  if (!template) {
    template = getTemplate();
  }
  if (getImageField(template)) {
    const imageConfig = template.cloudSheet.image;
    if (imageConfig) {
      return imageConfig;
    }
  }
  console.warn('The image config does not exist');

  return null;
}

/**
 * Obtain device configuration information
 * @param {*} template 
 * @returns 
 */
export const getEquipmentConfig = (template) => {
  if (!template) {
    template = getTemplate();
  }
  const equipmentConfig = template.cloudSheet.center.equipment;
  return equipmentConfig;
}
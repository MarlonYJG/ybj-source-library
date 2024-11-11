/*
 * @Author: Marlon
 * @Date: 2024-05-17 14:53:45
 * @Description:
 */

import { imgUrlToBase64, GetUserCompany } from '../utils/index';
import { monitorInstance } from './monitor';

import { getWorkBook, getInitData } from './store';
import { getTableHeaderDataTable } from './parsing-template';
import { PubGetTableStartRowIndex, classificationAlgorithms } from './single-table';
import { TOTAL_COMBINED_MAP } from './constant';
import { CombinationTypeBuild } from './combination-type';

/**
 *Index letter algorithm
 * @param {*} number
 * @returns
 */
export const numberToColumn = (number) => {
  let result = '';
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
};

/**
 * Alphabetical index algorithm
 * @param {*} column
 * @returns
 */
export const columnToNumber = (column) => {
  let result = 0;
  const length = column.length;
  for (let i = 0; i < length; i++) {
    result *= 26;
    result += column.charCodeAt(i) - 64;
  }
  return Number(result);
};

/**
 * Get a random number
 * @returns
 */
export const PubGetRandomNumber = () => {
  const chars = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
  let res = '';
  for (let i = 0; i < 8; i++) {
    const id = Math.ceil(Math.random() * 35);
    res += chars[id];
  }
  const timestamp = new Date().getTime();
  return timestamp + res;
};

/**
 * Array objects are deduplicated and sorted by the specified field
 * @param {*} array
 * @param {*} key
 * @returns
 */
export const uniqAndSortBy = (array, key, sortKey) => {
  const map = new Map();
  array.forEach(item => map.set(key(item), item));
  const sortedArray = Array.from(map.values()).sort((a, b) => a[sortKey] - b[sortKey]);

  return sortedArray;
};

/**
 * Get all tables and their ranges
 * @param {*} sheet
 * @returns
 */
export const GetAllTableRange = (sheet) => {
  const tables = sheet.tables.all();
  const tablesRange = {};
  for (let index = 0; index < tables.length; index++) {
    const table = tables[index];
    tablesRange[table.name()] = table.range();
  }
  return tablesRange;
};

/**
 * Update formula in the formula field
 * @param {*} obj 
 * @returns 
 */
export const updateFormula = (str) => {
  const regex = /{{\d+}}/g;
  return str.replace(regex, '{{row}}');
}

/**
 * Replace multiple placeholders in a string with specified dynamic field names and values
 * @param {*} formula 
 * @param {*} variables 
 * @returns 
 */
export const replacePlaceholders = (formula, variables) => {
  let updatedFormula = formula;
  for (const [fieldName, value] of Object.entries(variables)) {
    const regex = new RegExp(`{{${fieldName}}}`, 'g');
    updatedFormula = updatedFormula.replace(regex, value);
  }
  return updatedFormula;
}

/**
 * Cell pop-up box
 * @param {*} spread 
 * @param {*} args 
 * @param {*} template 
 */
export const cellDialog = (spread, args, template) => {
  const sheet = spread.getActiveSheet();
  if (args?.newSelections?.length === 1) {
    const target = args.newSelections[0];
    const templateSelection = template.cloudSheet.dateSelection;
    if (templateSelection) {
      for (const key in templateSelection) {
        if (Object.prototype.hasOwnProperty.call(templateSelection, key)) {
          const value = templateSelection[key];
          if (!value.row) {
            value.row = sheet.getRowCount() - value.lastRow - 1;
          }

          const hitTarget = value.row === target.row && value.rowCount == target.rowCount && value.column == target.col && value.columnCount == target.colCount;
          if (hitTarget) {
            return monitorInstance.callFunction('cellDialog', {
              cellBindValue: value
            });;
          }
        }
      }
    }
  }
}

/**
 * Bind company's logo
 * @param {*} spread
 * @param {*} template
 * @param {*} base64
 * @returns
 */
const insertLogo = (spread, template, base64) => {
  if (!spread) return;
  if (!base64) return;
  if (template.cloudSheet.logo) {
    const sheet = spread.getActiveSheet();
    sheet.suspendPaint();
    let picture = sheet.pictures.get(template.cloudSheet.logo.name);
    if (picture) {
      picture.src(base64);
      sheet.resumePaint();
      return;
    }
    picture = sheet.pictures.add(template.cloudSheet.logo.name, base64, template.cloudSheet.logo.x, template.cloudSheet.logo.y, template.cloudSheet.logo.width, template.cloudSheet.logo.height);
    picture.allowMove(false);
    picture.allowResize(true);
    picture.isLocked(true);
    sheet.resumePaint();
  }
};

/**
 * Render logo
 * @param {*} spread 
 * @param {*} template 
 * @param {*} base64 
 */
export const RenderLogo = (spread, template, quotation) => {
  template = getWorkBook(template);
  if (template.cloudSheet.logo && quotation.logo) {
    imgUrlToBase64(quotation.logo, (base64) => {
      insertLogo(spread, template, base64);
    });
  } else {
    let userBindCompany = GetUserCompany();
    if (userBindCompany && userBindCompany.logoURL) {
      imgUrlToBase64(userBindCompany.logoURL, (base64) => {
        insertLogo(spread, template, base64);
      });
    }
  }
};

/**
 * reset center sheet data
 * @param {*} spread
 */
export const Reset = (spread) => {
  const template = getWorkBook();
  const quotation = getInitData();
  const resourceViews = quotation.conferenceHall.resourceViews;
  const noClass = resourceViews.length === 1 && resourceViews[0].name === '无分类';

  const { type = null } = template.cloudSheet.center;
  const tableStartRowIndex = PubGetTableStartRowIndex();

  let header = [];
  // table header
  if (!noClass) {
    if (type) {
      const headerTable = getTableHeaderDataTable(type, true);
      if (headerTable.length) {
        header = headerTable;
      }
    }
  }
  const { classRow, subTotal, tableHeaderRow } = classificationAlgorithms(quotation, header, null);

  let centerCount = tableHeaderRow;
  resourceViews.forEach((item) => {
    centerCount = centerCount + item.resources.length + classRow + subTotal;
  });

  const sheet = spread.getActiveSheet();
  sheet.suspendPaint();
  sheet.deleteRows(tableStartRowIndex, centerCount);
  sheet.resumePaint();
};

/**
 * Reset the total
 * @param {*} sheet
 * @param {*} quotation
 * @param {*} bottom
 * @param {*} total
 */
export const ResetTotal = (sheet, quotation, bottom, total, top, mark) => {
  if (mark) {
    const resourceViews = quotation.conferenceHall.resourceViews;
    const totalRowIndex = top.mixCount + resourceViews.length;
    const combined = TOTAL_COMBINED_MAP[CombinationTypeBuild(quotation)];
    sheet.deleteRows(totalRowIndex, mark[combined].rowCount);
  } else {
    const Total = total[CombinationTypeBuild(quotation)];

    const totalRowStartIndex = sheet.getRowCount() - bottom.rowCount - Total.rowCount;
    sheet.deleteRows(totalRowStartIndex, Total.rowCount);
  }
};

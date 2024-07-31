/*
 * @Author: Marlon
 * @Date: 2024-05-17 14:53:45
 * @Description:
 */

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
/*
 * @Author: Marlon
 * @Date: 2024-05-15 21:16:41
 * @Description:
 */
import _ from '../lib/lodash/lodash.min.js';
import { isNumber } from '../utils/index';

import { MENU_TOTAL } from './config';
import { showTotal } from '../common/parsing-template'

/**
 *
 * @param {*} total
 * @param {*} arr
 * @param {*} name
 * @param {*} key
 */
export const delField = (keys, arr, name, key) => {
  let list = _.cloneDeep(arr);
  if (!keys.includes(key)) {
    list = arr.filter((item) => {
      return item.value !== name;
    });
  }
  return list;
};

/**
 * A list of totals
 * @param {*} total
 * @returns
 */
export const menuTotal = (total, selectTotal) => {
  let list = _.cloneDeep(MENU_TOTAL);

  console.log(total, 'total',showTotal());
  if (showTotal()) {
    const keys = Object.keys(total);
    const keyBin = [];
    let bindPaths = [];
    keys.forEach(key => {
      if (isNumber(Number(key))) {
        keyBin.push(key);
      }
    });

    list = delField(keyBin, list, 'tax', '1');
    list = delField(keyBin, list, 'concessional', '2');
    list = delField(keyBin, list, 'rate', '4');
    list = delField(keyBin, list, 'freight', '8');
    list = delField(keyBin, list, 'managementExpense', '16');
    list = delField(keyBin, list, 'projectCost', '32');

    keyBin.forEach(key => {
      const bindP = total[key].bindPath;
      if (bindP) {
        bindPaths = bindPaths.concat(Object.keys(bindP));
      }
    });
    bindPaths = Array.from(new Set(bindPaths));
    for (let index = 0; index < list.length; index++) {
      if (!bindPaths.includes(list[index].value)) {
        list.splice(index, 1);
      }
    }

    // disable
    const usedTotalKey = [];
    const activeTotal = total[selectTotal].bindPath;
    for (const key in activeTotal) {
      if (Object.hasOwnProperty.call(activeTotal, key)) {
        usedTotalKey.push(key);
      }
    }
    list.forEach(item => {
      if (usedTotalKey.includes(item.value)) {
        item.disable = true;
      }
    });
  } else {
    list = [];
  }
  return list;
};

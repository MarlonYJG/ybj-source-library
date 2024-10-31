/*
 * @Author: Marlon
 * @Date: 2024-07-11 17:37:38
 * @Description: Utils
 */
import moment from '../lib/dayjs/dayjs.min.js';

/**
 * Chinese
 */
export const regChineseCharacter = /[\u4E00-\u9FFF]/;

/**
 * The image address is changed to base64
 * base64格式：data:image/png;base64,iVBORw0KGgoAAAANSU...
 * @param {*} url 
 * @param {*} callback 
 */
export const imgUrlToBase64 = async (url, callback, isCompress = false) => {
  if (isCompress) {
    const imgUrl = `${url}?x-oss-process=image/resize,w_100,h_100/quality,q_50`
    await fetch(imgUrl).then(res => {
      return res.arrayBuffer()
    }).then(arrayBuffer => {
      const base64 = 'data:image/png;base64,' +
        btoa(
          new Uint8Array(arrayBuffer).reduce(function (data, byte) {
            return data + String.fromCharCode(byte)
          }, '')
        );
      callback(base64);
    }).catch(error => {
      console.error(error, '图片转base64失败');
    });
  } else {
    const img = new Image();
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');

    img.crossOrigin = 'anonymous';// 也需要后端的支持
    img.onload = function () {
      canvas.width = img.width;
      canvas.height = img.height;

      ctx.drawImage(img, 0, 0);
      callback(canvas.toDataURL());
    };
    img.src = url;
  }
};

/**
* file或blob转base64
* @param {*} blob file或者blob
* @param {*} callback function (data)通过参数获得base64
*/
export const BlobToBase64 = (blob, callback) => {
  const reader = new FileReader();
  reader.addEventListener('load', () => {
    callback(reader.result);
  });
  reader.readAsDataURL(blob);
};

/**
 * Whether the value is a number or not
 * @param {*} value 
 * @returns 
 */
export const isNumber = (value) => {
  return typeof value === 'number' && !Number.isNaN(value);
};

/**
 * Time Format Conversion and Get Current Time
 * @param {*} date
 * @param {*} format
 * @returns
 */
export const FormatDate = (date, format = 'YYYY-MM-DD HH:mm:ss') => {
  if (date) {
    return moment(date, format).format(format);
  }
  return moment().format(format);
};

/**
 * Obtain the company to which the user is bound
 * @returns
 */
export const GetUserCompany = () => {
  const company = sessionStorage.getItem('userBindCompany');
  if (company) {
    return JSON.parse(company);
  }
  return null;
};

/**
 * Get user information details
 */
export const GetUserInfoDetail = () => {
  const user = sessionStorage.getItem('userInfoDetail');
  if (user) {
    return JSON.parse(user);
  }
  return null;
};

/**
 * Obtain the current user member information
 * @returns 
 */
export const GetUserEmployee = () => {
  const userEmployee = sessionStorage.getItem('userEmployeeInfo');
  if (userEmployee) {
    return JSON.parse(userEmployee);
  }
  return null;
};

/**
 * Get the system time
 * @param {*} date 
 * @param {*} fmt 
 * @returns 
 */
export const getSystemDate = (date, fmt) => {
  try {
    const o = {
      'M+': date.getMonth() + 1, // 月份
      'd+': date.getDate(), // 日
      'h+': date.getHours() % 12 == 0 ? 12 : date.getHours() % 12, // 小时
      'H+': date.getHours(), // 小时
      'm+': date.getMinutes(), // 分
      's+': date.getSeconds(), // 秒
      'q+': Math.floor((date.getMonth() + 3) / 3), // 季度
      S: date.getMilliseconds() // 毫秒
    };
    const week = {
      0: '/u65e5',
      1: '/u4e00',
      2: '/u4e8c',
      3: '/u4e09',
      4: '/u56db',
      5: '/u4e94',
      6: '/u516d'
    };
    if (/(y+)/.test(fmt)) {
      fmt = fmt.replace(RegExp.$1, (date.getFullYear() + '').substr(4 - RegExp.$1.length));
    }
    if (/(E+)/.test(fmt)) {
      fmt = fmt.replace(RegExp.$1, ((RegExp.$1.length > 1) ? (RegExp.$1.length > 2 ? '/u661f/u671f' : '/u5468') : '') + week[date.getDay() + '']);
    }
    for (const k in o) {
      if (new RegExp('(' + k + ')').test(fmt)) {
        fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (('00' + o[k]).substr(('' + o[k]).length)));
      }
    }
    return fmt;
  } catch (error) {
    console.log(error);
    return date;
  }
};

/**
 * Flattens nested array data structures into one-dimensional arrays according to the specified fields and maintains the original order
 * @param {*} nestedArray 
 * @param {*} field 
 * @returns 
 */
export const flattenArray = (nestedArray, field) => {
  const result = [];
  function traverse(node) {
    if (Array.isArray(node)) {
      node.forEach(item => traverse(item));
    } else {
      const newNode = { ...node };
      if (newNode[field] && Array.isArray(newNode[field])) {
        traverse(newNode[field]);
      }
      delete newNode[field];
      result.push(newNode);
    }
  }

  traverse(nestedArray);
  return result;
};

/**
 * Pass in an object map and sort it according to the specified field row.
 * @param {*} map 
 * @returns 
 */
export const sortObjectByRow = (map) => {
  const entries = Object.entries(map);
  entries.sort((a, b) => a[1].row - b[1].row);
  const sortedArray = entries.map(entry => ({ [entry[0]]: entry[1] }));
  return sortedArray;
}

/**
 * Whether it is in percentage format
 * @param {*} value 
 * @returns 
 */
export const isPercentage = (value) => {
  const pattern = /^\d+(\.\d+)?%$/;
  return pattern.test(value);
};

/**
 * 对数字进行四舍五入到指定的小数位数
 * @param {number} num - 需要四舍五入的数字
 * @param {number} decimalPlaces - 保留的小数位数
 * @returns {number} - 四舍五入后的结果
 */
export const roundToDecimal = (num, decimalPlaces = 2) => {
  const factor = Math.pow(10, decimalPlaces);
  return Math.round(num * factor) / factor;
};
/*
 * @Author: Marlon
 * @Date: 2024-07-11 17:37:38
 * @Description: Utils
 */
import moment from 'moment';
import { Message } from 'element-ui';

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
export const imgUrlToBase64 = (url, callback) => {
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
 * @name: ResDatas
 * @param {type}
 * @return:
 * @description: 响应处理类
 */
export class ResDatas {
  constructor({
    res = null, code = 200, msg = null, error = null,
    fieldKey = {
      data: 'data',
      title: 'title',
      total: 'total',
      pageSize: 'pageSize'
    }
  }, load = null, vm = null) {
    this.res = res;
    this.code = Number(code);
    this.vm = vm;
    this.load = load;
    this.msg = msg;
    this.error = error;
    this.serverError = '服务异常';
    this.fieldKey = fieldKey;
  }

  init(cb) {
    if (this.vm && this.load) {
      this.vm[this.load] = false;
    }
    try {
      if (this.res && this.res.data && (Number(this.res.data.code) === this.code)) {
        cb && cb();
        const resData = this.res.data.data;
        if (resData === 0 || resData) {
          return resData;
        } else {
          console.warn('暂无数据(缺少data字段)');
          return null;
        }
      } else {
        // Message({
        //     showClose: true,
        //     message: this.error || this.res.data.message,
        //     type: "error",
        //     dangerouslyUseHTMLString: true
        // });
        return null;
      }
    } catch (error) {
      // Message({
      //     showClose: true,
      //     message: this.serverError,
      //     type: "error",
      //     dangerouslyUseHTMLString: true
      // });
      console.error(error, '服务异常——响应解析类');
      return null;
    }
  }

  // 状态(增/删/改) 格式化
  state(cb) {
    try {
      if (this.res && this.res.data && Number(this.res.data.code) === this.code) {
        cb && cb(this.res.data.data);
        if (this.msg) {
          Message({
            showClose: true,
            message: this.msg,
            type: 'success',
            dangerouslyUseHTMLString: true
          });
          return true;
        } else {
          Message({
            showClose: true,
            message: this.res.data.message || this.res.data.msg || '',
            type: 'success',
            dangerouslyUseHTMLString: true
          });
          return true;
        }
      } else {
        // if (this.error) {
        //     Message({
        //         showClose: true,
        //         message: this.error,
        //         type: "error",
        //         dangerouslyUseHTMLString: true
        //     });
        //     return false
        // } else {
        //     Message({
        //         showClose: true,
        //         message: this.res.data.message || '',
        //         type: "error",
        //         dangerouslyUseHTMLString: true
        //     });
        //     return false
        // }
        return false;
      }
    } catch (error) {
      console.error(error, '服务异常——响应解析类')
      // if (this.error) {
      //     Message({
      //         showClose: true,
      //         message: this.serverError,
      //         type: "error",
      //         dangerouslyUseHTMLString: true
      //     });
      //     return false
      // } else {
      //     Message({
      //         showClose: true,
      //         message: this.res.data.message || '',
      //         type: "error",
      //         dangerouslyUseHTMLString: true
      //     });
      //     return false
      // }
      return false;
    }
  }
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

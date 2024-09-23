/*
 * @Author: Marlon
 * @Date: 2024-03-27 22:35:21
 * @Description:single - public
 */
import Decimal from '../lib/decimal/decimal.min.js';
import * as GC from '@grapecity/spread-sheets';
import _ from '../lib/lodash/lodash.min.js';
import store from 'store';
import { GetUserCompany, imgUrlToBase64, regChineseCharacter } from '../utils/index';

import { LayoutRowColBlock } from './core';

import { GENERATE_FIELDS_NUMBER, DESCRIPTION_MAP, REGULAR, ASSOCIATED_FIELDS_FORMULA_MAP } from './constant';
import { columnToNumber, PubGetRandomNumber, replacePlaceholders } from './public';
import { GeneratorUpperCaseFormatter } from './generator';

import { NEW_OLD_FIELD_MAP } from '../build-library/config'
import { ShowCostPrice, ShowCostPriceStatus } from '../build-library/head'
import { synchronousStoreSumAmount } from '../build-library/single-table'

import {
  templateRenderFlag,
  setCell,
  AddEquipmentImage,
  showSubTotal,
  getTableHeaderDataTable,
  getFormulaFieldRowCol
} from './parsing-template';
import { IdentifierTemplate, DEFINE_IDENTIFIER_MAP } from './identifier-template'
import { formatterPrice, getShowCostPrice } from './parsing-quotation';

import { SHOW_COST_PRICE_HEAD, UPDATE_QUOTATION_PATH } from "store/quotation/mutation-types";

const NzhCN = require('../lib/nzh/cn.min.js');

/**
 * Row computed fields
 * @returns
 */
const GenerateFieldsRow = () => {
  const totalBeforeTax = ['totalBeforeTax'];
  for (let index = 0; index < GENERATE_FIELDS_NUMBER.totalBeforeTax; index++) {
    totalBeforeTax.push(`totalBeforeTax${index}`);
  }
  return totalBeforeTax;
};

/**
 * Column total 字段名(全量)
 * @returns
 */
export const AssociatedFieldsColumn = () => {
  const fields = ['total'];
  for (let index = 0; index <= 15; index++) {
    fields.push(`total${index}`);
  }
  return fields;
};

/**
 * Data processing - Synchronize the local store according to the quotation information interface
 * @param {*} Res
 * @returns
 */
export const singleTableSyncStore = (Res) => {
  const customer = Res.customer;
  const company = GetUserCompany();
  const { projectNumber = '', belongs = '', name } = store.getters['projectModule/GetterProjectInit'];
  const quotationDefault = {
    id: Res.id, // 报价单id
    title: Res.title, // 报价单标题
    logo: Res.logo, // 绑定公司的logo
    seal: Res.seal, // 绑定公司的印章
    quotationImage: Res.quotationImage, // 报价单的图片(旧版字段)
    quaLogos: Res.quaLogos || [], // 报价单多logo
    state: Res.state,
    leaderId: belongs, // 负责人id
    belongs: Res.belongs,
    belongsEmail: Res.belongsEmail,
    phone: Res.phone,
    createTime: Res.createTime, // 创建时间
    updateTime: Res.updateTime, // 更新时间
    needApproval: '', // 是否需要审批
    resources: Res.resources, // 存储总表的数据
    storePhone: Res.storePhone,

    // 项目
    name: Res.name || name, // 项目名称
    projectName: Res.projectName, // 项目名称
    projectType: Res.projectType,
    projectNumber: Res.projectNumber || projectNumber,
    // 客户
    projectManager: Res.projectManager,
    projectManagerPhone: Res.projectManagerPhone,

    // 客户
    customer,
    // 公司信息
    companyAddress: Res.companyAddress,
    companyEmail: Res.companyEmail,
    companyFax: Res.companyFax,
    companyName: Res.companyName || company.name,
    companyPhone: Res.companyPhone,
    companyWebsite: Res.companyWebsite,

    company: {
      companyName: Res.companyName || company.name,
      createUserId: company.createUserId, // 用户名
      companyAddress: Res.companyAddress || company.address, // 公司地址
      companyEmail: Res.companyEmail, // 公司邮箱，
      logoURL: company.logoURL,
      companyPhone: Res.companyPhone || company.tel, // 公司电话
      companyWebsite: Res.companyWebsite || company.website, // 公司网址
      companyFax: Res.companyFax || company.fax, // 公司传真
      industry: company.industry, // 行业
      state: company.state
    },

    conferenceHall: Res.conferenceHall, // 报价单的会场信息
    parallelSessions: [], // 分会场信息
    preferentialWay: Res.preferentialWay,
    noImgTemplate: Res.noImgTemplate,
    designerPhone: Res.designerPhone,
    designer: Res.designer,
    engineerPhone: Res.engineerPhone,
    engineer: Res.engineer,

    /** 总计相关 */
    willPay: Res.willPay,
    DXdeposit: Res.DXdeposit, // 定金大写
    deposit: Res.deposit, // 定金
    DXwillPay: Res.DXwillPay, // 合计大写
    DXzje: Res.DXzje,
    capitalizeTotalAmount: Res.capitalizeTotalAmount,
    showCost: Res.showCost,
    totalAmount: Res.totalAmount,
    sumAmount: Res.sumAmount, // 合计
    dxsumAmount: Res.dxsumAmount, // 总金额大写
    freight: Res.freight, // 运费
    projectCost: Res.projectCost, // 项目费用
    managementFee: Res.managementFee, // 管理费率
    managementExpense: Res.managementExpense, // 管理费
    rate: Res.rate, // 服务费率(新)
    serviceCharge: Res.serviceCharge, // 服务费
    taxRate: Res.taxRate, // 税率
    tax: Res.tax, // 税率(新)
    discount: Res.discount, // 折扣
    concessionalRate: Res.concessionalRate, // 优惠价

    excelJson: Res.excelJson,
    templateType: Res.templateType,
    remark: Res.remark, // 备注
    exportName: Res.exportName, // 导出名称
    subheadingOne: Res.subheadingOne, // 副标题一
    subheadingTwo: Res.subheadingTwo, // 副标题二
    priceAdjustment: Res.priceAdjustment, // 单价调整比例
    priceStatus: Res.priceStatus, // 价格设置
    extFields: Res.extFields, // 预置字段
    image: Res.image, // 报价单图片
    templateId: Res.templateId,
    isInt: Res.isInt,
    priceType: Res.priceType,

    // 特殊标识
    haveExport: Res.haveExport, // 决定能否导出:未审核通过
    quotationExcel: Res.quotationExcel, // 导出excel报价单的URL
    quotationPdf: Res.quotationPdf,// 导出pdf报价单的URL

    // 配置相关信息
    config: Res.config || {},

  };

  console.log(quotationDefault, '====响应数据处理-并同步至store');
  return quotationDefault;
};

/**
 * quotation与template的逻辑关联处理
 * @param {*} quotation
 * @param {*} template
 */
export const LogicalProcessing = (quotation, template) => {
  const { titleHAlign = '' } = template;
  // 产品字段关联性 逻辑处理
  const resourceViews = quotation.conferenceHall.resourceViews;
  if (resourceViews.length) {
    resourceViews.forEach(classItem => {
      if (classItem.resources.length) {
        classItem.resources.forEach(item => {
          if (titleHAlign) {
            item.quantity = 1;
            item.batch = 1;
          }

          if (item.discountUnitPrice === 0 || item.discountUnitPrice) {
            item.discountUnitPrice = parseFloat(item.discountUnitPrice);
          } else {
            item.discountUnitPrice = parseFloat(Number(item.unitPrice));
          }
          if (item.quantity === 0 || item.quantity) {
            item.quantity = parseFloat(item.quantity);
          } else {
            item.quantity = 0;
          }
          if (!item.imageId) {
            item.imageId = PubGetRandomNumber();
          }
          if (item.numberOfDays === 0 || item.numberOfDays) {
            item.numberOfDays = Number(item.numberOfDays);
          } else {
            item.numberOfDays = 1;
          }
        });
      }
    });
  }
};

/**
 * Amount-related logical processing
 * @param {*} quotation
 */
export const LogicalAmount = (quotation) => {
  if (quotation.sumAmount === 0 || quotation.sumAmount) {
    quotation.DXwillPay = NzhCN.encodeB(quotation.sumAmount);
  } else {
    quotation.DXwillPay = '';
  }
  if (quotation.deposit === 0 || quotation.deposit) {
    quotation.DXdeposit = NzhCN.encodeB(quotation.deposit);
  } else {
    quotation.DXdeposit = '';
  }
  if (quotation.sumAmount === 0 || quotation.sumAmount) {
    quotation.dxsumAmount = NzhCN.encodeB(quotation.sumAmount);
  } else {
    quotation.dxsumAmount = '';
  }
  if (quotation.concessionalRate === 0 || quotation.concessionalRate) {
    quotation.totalAmount = quotation.concessionalRate;
  } else {
    quotation.totalAmount = quotation.sumAmount || '';
  }
  if (!quotation.capitalizeTotalAmount) {
    quotation.capitalizeTotalAmount = NzhCN.encodeB(quotation.totalAmount);
  }
  if (quotation.deposit > 0) {
    quotation.willPay = quotation.totalAmount - quotation.deposit;
    quotation.DXzje = quotation.willPay ? NzhCN.encodeB(quotation.willPay) : '';
  } else {
    quotation.willPay = quotation.totalAmount;
    quotation.DXzje = quotation.capitalizeTotalAmount || NzhCN.encodeB(quotation.totalAmount);
  }
  console.log(quotation, '金额转换');
};

/**
 * total中的组合类型
 * @param {*} key
 * @returns
 */
export const templateTotalMap = (key) => {
  // const map = {
  //   16: 3,
  //   17: 2,
  //   18: 1,
  //   19: 0,
  //   20: 4,
  //   21: 5,
  //   22: 6,
  //   23: 7,
  //   24: 8,
  //   25: 9,
  //   26: 10,
  //   27: 11,
  //   28: 12,
  //   29: 13,
  //   30: 14,
  //   31: 15
  // };
  // if (map[key] === 0 || map[key]) {
  //   return map[key];
  // }
  return key;
};

/**
 * Product list added sorting
 * @param {*} quotation
 */
export const resourceSort = (quotation) => {
  const resourceViews = quotation.conferenceHall.resourceViews;
  for (let index = 0; index < resourceViews.length; index++) {
    resourceViews[index].sort = index + 1;
    if (resourceViews[index].resources) {
      resourceViews[index].resources.forEach((resource, i) => {
        resource.sort = i + 1;
      });
    }
  }
};

/**
 * Merge cells
 * @param {*} sheet
 * @param {*} spans
 * @param {*} row
 */
export const mergeSpan = (sheet, spans, row) => {
  if (spans) {
    for (let i = 0; i < spans.length; i++) {
      const ele = spans[i];
      sheet.addSpan(row + ele.row, ele.col, ele.rowCount, ele.colCount);
    }
  }
};

/**
 * Style the cells
 * @param {*} spread
 * @param {*} field
 * @param {*} startRowIndex
 * @param {*} locked
 */
export const setCellStyle = (spread, field, startRowIndex, locked = false) => {
  const sheet = spread.getActiveSheet();
  const { dataTable } = field;
  for (const i in dataTable) {
    if (Object.hasOwnProperty.call(dataTable, i)) {
      const row = dataTable[i];
      for (const j in row) {
        if (Object.hasOwnProperty.call(row, j)) {
          const column = _.cloneDeep(row[j]);
          column.style || (column.style = {});
          column.style.locked = locked;
          const style = new GC.Spread.Sheets.Style();
          style.fromJSON(column.style);
          sheet.setStyle(startRowIndex + Number(i), Number(j), style, GC.Spread.Sheets.SheetArea.viewport);
          if (column.value === 0 || column.value) {
            sheet.setValue(startRowIndex + Number(i), Number(j), column.value);
          }
        }
      }
    }
  }
};

/**
 * Total 行高
 * @param {*} sheet
 * @param {*} total
 * @param {*} totalField
 * @param {*} row
 */
export const setTotalRowHeight = (sheet, total, totalField, row) => {
  for (const key in totalField.dataTable) {
    if (Object.hasOwnProperty.call(totalField.dataTable, key)) {
      if (totalField.height === 0 || totalField.height) {
        sheet.setRowHeight(row + Number(key), totalField.height);
      } else if (total.style && total.style.height) {
        sheet.setRowHeight(row + Number(key), total.style.height);
      }
    }
  }
  for (const key in totalField.rows) {
    if (Object.hasOwnProperty.call(totalField.rows, key)) {
      sheet.setRowHeight(row + Number(key), totalField.rows[key].size);
    }
  }
};

/**
 * Insert the start of the row of the table
 * @returns
 */
export const PubGetTableStartRowIndex = (GetterQuotationWorkBook) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
  return template.cloudSheet.top.rowCount;
};
/**
 * The total number of rows inserted into the table
 */
export const PubGetTableRowCount = (index = 0, GetterQuotationInit) => {
  const quotation = GetterQuotationInit || store.getters['quotationModule/GetterQuotationInit'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  if (resourceViews.length === 1 && resourceViews[0].name === '无分类') {
    return resourceViews[0].resources.length;
  } else if (resourceViews.length) {
    return resourceViews[index].resources.length;
  } else {
    return 0;
  }
};

const getClassRowCount = (type, subtotal) => {
  const classRow = type ? type.rowCount : 0;
  const subTotal = subtotal ? subtotal.rowCount : 0;
  return {
    classRowCount: classRow,
    subTotalCount: subTotal
  };
};

/**
 * classification Algorithms
 * @param {*} quotation
 * @param {*} headers
 * @returns
 */
export const classificationAlgorithms = (quotation, headers = [], GetterQuotationWorkBook) => {
  const { template, mixRender, classType } = templateRenderFlag(GetterQuotationWorkBook);
  const { total = null, type = null } = template.cloudSheet.center;
  const { classRowCount, subTotalCount } = getClassRowCount(type, total);
  const { resourceViews } = quotation.conferenceHall;

  let classRCount;
  let classRow1 = 0;
  let subTotal;
  let tableHeaderRow = 0;

  if (resourceViews.length === 1 && resourceViews[0].name === '无分类') {
    classRCount = 0;
    classRow1 = 0;
    subTotal = 0;
    tableHeaderRow = 0;
  } else {
    if (classType === 'mergeClass') {
      classRCount = 0;
      classRow1 = 0;
      subTotal = 0;
    } else {
      if (mixRender) {
        if (resourceViews.length) {
          if (resourceViews[0].parentTypeName && resourceViews[0].name) {
            classRCount = 2;
            classRow1 = 1;
          } else if (!resourceViews[0].parentTypeName) {
            classRCount = 1;
            classRow1 = 0;
          } else if (!resourceViews[0].name) {
            classRCount = 1;
            classRow1 = 0;
          } else {
            console.warn('缺少分类标题：parentTypeName或name');
          }
        }

        subTotal = subTotalCount;
      } else {
        classRCount = classRowCount;
        subTotal = showSubTotal(GetterQuotationWorkBook) ? subTotalCount : 0;
      }
    }
  }

  if (headers.length) {
    tableHeaderRow = 1;
  }

  return {
    classRow: classRCount,
    classRow1,
    tableHeaderRow,
    subTotal
  };
};

/**
 * Get the subtotals for each column
 * @param {*} sheet
 * @param {*} tableStartRowIndex
 * @param {*} tableIndex
 * @param {*} showRowTotal
 * @param {*} columnComputed
 * @param {*} subBindPath
 * @returns
 */
export const columnsTotal = (sheet, tableStartRowIndex, tableIndex, showRowTotal = false, columnComputed, subBindPath, GetterQuotationWorkBook = null, quotation = null) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
  const equipment = template.cloudSheet.center.equipment.bindPath;

  const columnTotalMap = {};

  // Get all columns of Colentotal
  for (const column in equipment) {
    if (Object.hasOwnProperty.call(equipment, column)) {
      const columnComputedField = AssociatedFieldsColumn();
      columnComputedField.push('quantity');
      if (columnComputedField.includes(column)) {
        columnTotalMap[column] = _.cloneDeep(equipment[column]);
        if (subBindPath && subBindPath[column]) {
          columnTotalMap[column].rowHeader = subBindPath[column].columnHeader;
        } else if (column === 'quantity') {
          columnTotalMap[column].rowHeader = equipment[column].columnHeader;
        }
        if (columnComputed) {
          columnComputed.push(column);
        }
      }
    }
  }
  // Sets the formula for calculating the subtotals for each column
  const resources = PubGetTableRowCount(tableIndex, quotation);
  for (const key in columnTotalMap) {
    const formula = [];
    if (Object.hasOwnProperty.call(columnTotalMap, key)) {
      for (let index = 0; index < resources; index++) {
        if (showRowTotal) {
          // Subtotals exist
          formula.push(`${columnTotalMap[key].columnHeader}${tableStartRowIndex + index}`);
        } else {
          if (tableIndex === 0) {
            // No classification
            formula.push(`${columnTotalMap[key].columnHeader}${tableStartRowIndex + 1 + index}`);
          }
        }
      }
    }
    columnTotalMap[key].formula = `${formula.join('+')}`;
  }

  // Sets the sum of the subtotals for each column
  for (const totalField in columnTotalMap) {
    if (Object.hasOwnProperty.call(columnTotalMap, totalField)) {
      const arr = [];
      const cellArr = PubGetFormulaValue(columnTotalMap[totalField].formula);
      cellArr.forEach(item => {
        const val = sheet.getValue(item[0], item[1]);
        if (typeof (val) === 'number') {
          arr.push(val);
        } else {
          arr.push(0);
        }
      });
      columnTotalMap[totalField].value = arr;
      if (arr.length) {
        columnTotalMap[totalField].sum = arr.reduce((prev, curr) => { return new Decimal(prev).plus(new Decimal(curr)).toNumber(); });
      } else {
        columnTotalMap[totalField].sum = 0;
      }
    }
  }

  return columnTotalMap;
};

export const PubGetFormulaValue = (formulas) => {
  const formulaArr = formulas.split('+');
  const rowColumn = [];
  formulaArr.forEach(item => {
    rowColumn.push([Number(item.substring(1)) - 1, columnToNumber(item.substring(0, 1)) - 1]);
  });
  return rowColumn;
};

/**
 * The sum of the totals of all calculated columns
 * @param {*} columnTotal
 * @returns
 */
export const plusColumnTotalSum = (columnTotal) => {
  let sum = 0;
  columnTotal.forEach(tableColumnComputedItem => {
    const columns = Object.keys(tableColumnComputedItem);
    if (columns.length >= 1) {
      if (columns.includes('total')) {
        sum = new Decimal(sum).plus(new Decimal(tableColumnComputedItem.total.sum)).toNumber();
      } else {
        console.warn('计算列 中缺少固定计算字段(total：金额)');
      }
    }
  });
  return sum;
};

export const columnTotalSumFormula = (columnTotal) => {
  let formula = [];
  columnTotal.forEach((tableColumnComputedItem) => {
    const columns = Object.keys(tableColumnComputedItem);
    if (columns.length >= 1) {
      if (columns.includes('total')) {
        formula = formula.concat(tableColumnComputedItem.total.formula.split('+'));
      } else {
        console.warn('计算列 中缺少固定计算字段(total：金额)');
      }
    }
  });
  if (formula.length) {
    return formula.join('+');
  }
  return '';
};

/**
 * Calculated field sorting for totals rows
 * @param {*} totalBinds
 * @returns
 */
const rowComputedFieldSort = (totalBinds = {}) => {
  const totalBindsArr = [];
  for (const key in totalBinds) {
    if (Object.hasOwnProperty.call(totalBinds, key)) {
      totalBindsArr.push({
        ...totalBinds[key],
        fieldName: key
      });
    }
  }
  totalBindsArr.sort((a, b) => a.row - b.row);
  return totalBindsArr.map((item) => item.fieldName);
};

/**
 * Mixed description fields
 * @param {*} sheet
 * @param {*} quotation
 * @param {*} row
 * @param {*} rows
 */
export const mixedDescriptionFields = (sheet, quotation, row, rows) => {
  const pathString = DESCRIPTION_MAP[rows.bindPath].percentage;
  const paths = pathString.split('|');
  const path1 = paths[0];
  if (paths.length > 1) {
    const path2 = paths[1];
    if (_.has(quotation, [path2]) && _.get(quotation, [path2])) {
      let rate = _.get(quotation, [path2]);
      if (rate === 0 || rate) {
        rate = Number(rate);
      }
      sheet.setValue(row + rows.row, rows.column, `${rows.description2}${rate}%${rows.description}`);
      return;
    }
  }

  if (_.has(quotation, [path1])) {
    let rate = _.get(quotation, [path1]);
    if (rate === 0 || rate) {
      rate = Number(rate);
    }
    sheet.setValue(row + rows.row, rows.column, `${rows.description2}${rate}%${rows.description}`);
  }
};

/**
 * Multi table header
 * @param {*} sheet
 * @param {*} headers
 * @param {*} rowIndex
 * @param {*} height
 */
export const tableHeader = (sheet, headers, rowIndex, height) => {
  if (headers.length) {
    for (let index = 0; index < headers.length; index++) {
      sheet.addRows(rowIndex + index, 1);
      setCell(sheet, headers[index], rowIndex + index - 1);
      if (height) {
        sheet.setRowHeight(rowIndex + index - 1, height);
      }
    }
  }
};

/**
 * Set the calculation formula for cell binding
 * @param {*} sheet
 * @param {*} formulaObj
 * @param {*} startRowIndex
 * @param {*} index
 */
const PubGetFormula = (sheet, formulaObj, startRowIndex, watchErr) => {
  // TODO 待优化
  const columnFormula = _.cloneDeep(formulaObj);

  let forms = columnFormula.formula.substring(columnFormula.formula.indexOf('(') + 1, columnFormula.formula.lastIndexOf(')'));

  if (forms.indexOf('{{') !== -1) {
    for (let i = 0; i < columnFormula.size; i++) {
      forms = forms.replace('{{' + i + '}}', startRowIndex + 1);
    }
  } else if (forms.indexOf('*') !== -1) {
    const columns = [];
    const cells = forms.split('*');
    cells.forEach(cell => {
      const column = cell.substring(0, 1);
      columns.push(`${column}${startRowIndex + 1}`);
    });
    forms = columns.join('*');
  } else {
    console.warn(`无法识别的计算公式：${columnFormula.formula}`);
  }
  if (watchErr) {
    // TODO
    // columnFormula.formula = 'if(iserror(' + forms + '),"",' + forms + ')';
    // IFERROR(A3/A5,"dogs")
    // forms = `IFERROR(${forms},'')`;
  }
  sheet.getCell(startRowIndex, columnFormula.column).locked(true);
  setCellFormatter(sheet, startRowIndex, columnFormula.column);
  sheet.setFormula(startRowIndex, columnFormula.column, forms);
  // sheet.autoFitColumn(columnFormula.column);

  if (sheet.getValue(startRowIndex, columnFormula.column) instanceof GC.Spread.CalcEngine.CalcError) {
    console.warn('无效公式');
  }
};

/**
 * Set the calculation formula for cell binding
 * @param {*} sheet 
 * @param {*} startRowIndex 
 * @param {*} computedColumn 
 */
const SetComputedColumnFormula = (sheet, startRowIndex, computedColumn) => {
  sheet.getCell(startRowIndex, computedColumn.column).locked(true);
  // setCellFormatter(sheet, startRowIndex, computedColumn.column);
  const formula = replacePlaceholders(computedColumn.formula, { row: startRowIndex + 1 })
  sheet.setFormula(startRowIndex, computedColumn.column, formula);
  // sheet.autoFitColumn(computedColumn.column);
};

/**
 * Dynamically calculated field values for columns
 * @param {*} sheet
 * @param {*} equipment
 * @param {*} startRow
 * @param {*} computedColumnFormula
 */
export const columnComputedValue = (sheet, equipment, startRow, computedColumnFormula) => {
  for (const key in equipment.bindPath) {
    if (Object.hasOwnProperty.call(equipment.bindPath, key)) {
      // total ~ 15
      if (AssociatedFieldsColumn().includes(key)) {
        PubGetFormula(sheet, equipment.bindPath[key].formula, startRow, true);
      } else if (Object.keys(computedColumnFormula).includes(key)) {
        SetComputedColumnFormula(sheet, startRow, computedColumnFormula[key])
      }
    }
  }
};

/**
 * Format cells
 * @param {*} sheet
 * @param {*} row
 * @param {*} column
 */
export const setCellFormatter = (sheet, row, column, quotation) => {
  const format = formatterPrice(quotation);
  if (format) {
    sheet.setFormatter(row, column, format);
  }
};

/**
 * Set the calculated value for the row Total (subtotal).
 * @param {*} sheet
 * @param {*} columnTotalMap
 * @param {*} bindPath
 */
export const SetComputedSubTotal = (sheet, columnTotalMap, bindPath, quotation = null) => {
  for (const key in columnTotalMap) {
    if (Object.hasOwnProperty.call(columnTotalMap, key)) {
      if (columnTotalMap[key] && columnTotalMap[key].formula) {
        const formula = columnTotalMap[key].formula.split('+');
        const subTotalRowIndex = formula[formula.length - 1].substring(1);
        if (bindPath && bindPath[key]) {
          const columnIndex = bindPath[key].columnHeader;
          setCellFormatter(sheet, Number(subTotalRowIndex), columnToNumber(columnIndex) - 1, quotation);
          sheet.setFormula(Number(subTotalRowIndex), columnToNumber(columnIndex) - 1, columnTotalMap[key].formula);
          // sheet.autoFitColumn(columnToNumber(columnIndex) - 1);
        }
      }
    }
  }
};

/**
 * Add an image to a table
 * @param {*} spread
 * @param {*} table
 * @param {*} classIndex
 * @param {*} insertTableIndex
 * @param {*} classRow
 * @param {*} subTotal
 * @param {*} allowMove
 * @param {*} allowResize
 * @param {*} isLocked
 */
const tableAddImage = (spread, table, classIndex, insertTableIndex, classRow, subTotal, allowMove, allowResize, isLocked, template, isCompress = false) => {
  table.forEach((item, index) => {
    let startRow = insertTableIndex + classRow + index;
    if (classIndex !== 0) {
      startRow = insertTableIndex + subTotal + classRow + index;
    }
    if (item.images) {
      // Product images
      let imgUrl = item.images;
      if (REGULAR.chineseCharacters.test(item.images)) {
        imgUrl = encodeURI(item.images);
      }
      ((item, imgUrl, startRow) => {
        imgUrlToBase64(imgUrl, (base64) => {
          AddEquipmentImage(spread, item.id, base64, startRow, allowMove, allowResize, isLocked, template);
        }, isCompress);
      })(item, imgUrl, startRow);
    }
  });
};

/**
 * Render a picture of the product (asynchronous)
 * @param {*} spread
 * @param {*} tableStartRowIndex
 * @param {*} allowMove
 * @param {*} allowResize
 * @param {*} isLocked
 */
export const renderSheetImage = (spread, tableStartRowIndex, allowMove, allowResize, isLocked, GetterQuotationInit = null, template = null, isCompress = false) => {
  const quotation = GetterQuotationInit || store.getters['quotationModule/GetterQuotationInit'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const { classRow, subTotal } = classificationAlgorithms(quotation, [], template);

  let insertTableIndex = tableStartRowIndex;
  for (let index = 0; index < resourceViews.length; index++) {
    if (resourceViews[index].resources) {
      if (index === 0) {
        tableAddImage(spread, resourceViews[index].resources, index, insertTableIndex, classRow, subTotal, allowMove, allowResize, isLocked, template, isCompress);
        insertTableIndex = insertTableIndex + classRow + resourceViews[index].resources.length;
      } else {
        tableAddImage(spread, resourceViews[index].resources, index, insertTableIndex, classRow, subTotal, allowMove, allowResize, isLocked, template, isCompress);
        insertTableIndex = insertTableIndex + subTotal + classRow + resourceViews[index].resources.length;
      }
    }
  }
};

/**
 * table to sheet
 * @param {*} spread
 */
export const translateSheet = (spread) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const sheet = spread.getActiveSheet();
  sheet.suspendPaint();
  resourceViews.forEach(table => {
    ((tableId) => {
      const sheetTable = sheet.tables.findByName(`table${tableId}`);
      sheetTable && sheet.tables.remove(sheetTable, GC.Spread.Sheets.Tables.TableRemoveOptions.keepData | GC.Spread.Sheets.Tables.TableRemoveOptions.keepStyle);
    })(table.resourceLibraryId);
  });
  sheet.resumePaint();
};

/**
 * sheet to table
 * @param {*} spread
 * @param {*} tableMap
 */
export const translateTable = (spread, tableMap) => {
  const tables = Object.values(tableMap);
  if (tables && tables.length) {
    const template = store.getters['quotationModule/GetterQuotationWorkBook'];
    const header = template.cloudSheet.center.equipment.bindPath;
    const sheet = spread.getActiveSheet();
    for (const key in tableMap) {
      if (Object.hasOwnProperty.call(tableMap, key)) {
        const { row, rowCount, col, colCount } = tableMap[key];
        const id = key.split('table')[1];

        sheet.addRows(row, 1);

        // CreateTable(sheet, id, row, col, rowCount + 1, colCount, header);
        const table = sheet.tables.add(`table${id}`, row, col, rowCount, colCount);
        const tableColumns = [];
        for (const key in header) {
          if (Object.hasOwnProperty.call(header, key)) {
            const tableColumn = new GC.Spread.Sheets.Tables.TableColumn();
            tableColumn.name(header[key].name || key);
            tableColumn.dataField(header[key].bindPath);
            tableColumns.push(tableColumn);
          }
        }
        table.bindColumns(tableColumns);
        table.expandBoundRows(true);
        table.autoGenerateColumns(false);
        table.showFooter(false);
        table.showHeader(false);
        table.highlightFirstColumn(false);
        table.highlightLastColumn(false);

        // table.bindingPath(bindDataPath);
      }
    }
    console.log(sheet.tables.all());
  }
};

/**
 * The status of the initialization cost price is displayed
 * @returns 
 */
export const initShowCostPrice = (spread) => {
  const isShow = getShowCostPrice();
  store.commit(`quotationModule/${SHOW_COST_PRICE_HEAD}`, isShow);
  ShowCostPrice(spread, ShowCostPriceStatus());
}

/**
 * Obtain the classification type of the template
 * @returns 
 */
export const getTemplateClassType = () => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const templateClassIdentifier = template.cloudSheet.templateClassIdentifier;
  const resourceViews = quotation.conferenceHall.resourceViews;

  if (templateClassIdentifier) {
    if (Object.keys(DEFINE_IDENTIFIER_MAP).includes(templateClassIdentifier)) {
      return templateClassIdentifier
    } else {
      console.error('模板分类标识符不存在,无法确定模板分类类型【模板错误】');
    }
  } else {
    // TODO 兼容旧版标识符
    if (template.classType) {
      if (template.classType === 'mergeClass') {
        return 'Level_1_col'
      }
    }

    if (resourceViews.length) {
      if (resourceViews.length === 1 && resourceViews[0].name === '无分类') {
        return 'noLevel'
      } else {
        return 'Level_1_row'
      }
    } else {
      return 'Level_1_row'
    }
  }

  return null
}

/**
 * Get the formula for the calculated column
 * @param {*} sheet
 * @returns
 */
export const GetColumnComputedTotal = (sheet, GetterQuotationWorkBook = null, GetterQuotationInit = null) => {
  const template = GetterQuotationWorkBook || store.getters['quotationModule/GetterQuotationWorkBook'];
  const quotation = GetterQuotationInit || store.getters['quotationModule/GetterQuotationInit'];
  const { type = null } = template.cloudSheet.center;
  const resourceViews = quotation.conferenceHall.resourceViews;

  const noClass = resourceViews.length === 1 && resourceViews[0].name === '无分类';

  // Obtain the index of the table
  let header = [];

  // table header
  if (!noClass) {
    if (type) {
      const headerTable = getTableHeaderDataTable(type);
      if (headerTable.length) {
        header = headerTable;
      }
    }
  }

  const { classRow, subTotal, tableHeaderRow } = classificationAlgorithms(quotation, header, template);

  let insertTableIndex = PubGetTableStartRowIndex(template);
  const columnTotal = [];
  for (let index = 0; index < resourceViews.length; index++) {
    if (index === 0) {
      const columnTotalMap = columnsTotal(sheet, insertTableIndex + classRow + tableHeaderRow + 1, index, true, null, null, template, quotation);
      columnTotal.push(columnTotalMap);
      insertTableIndex = insertTableIndex + classRow + tableHeaderRow + resourceViews[index].resources.length;
    } else {
      const columnTotalMap = columnsTotal(sheet, insertTableIndex + subTotal + classRow + tableHeaderRow + 1, index, true, null, null, template, quotation);
      columnTotal.push(columnTotalMap);
      insertTableIndex = insertTableIndex + subTotal + classRow + tableHeaderRow + resourceViews[index].resources.length;
    }
  }

  return columnTotal;
};

/**
 * Clears the relevant calculated value for the totals area when no data is available
 * @param {*} sheet 
 * @param {*} row 
 * @param {*} totalField 
 * @param {*} fixedBindValueMap 
 * @param {*} quotation 
 * @param {*} template 
 * @param {*} libraryType 
 */
export const clearTotalNoData = (sheet, row, totalField, fixedBindValueMap, quotation, template, libraryType) => {
  const totalFieldBind = totalField.bindPath;

  console.log('==================clearTotalNoData==========');
  // Set a fixed value
  for (const key in totalFieldBind) {
    if (Object.hasOwnProperty.call(totalFieldBind, key)) {
      const rows = totalFieldBind[key];
      if (template.truckage) {
        const TruckageIdentifier = new IdentifierTemplate(sheet, 'truckage');
        TruckageIdentifier.truckageFreight(totalField, row, fixedBindValueMap);
      } else {
        if (rows.bindPath) {
          if (Object.keys(DESCRIPTION_MAP).includes(rows.bindPath)) {
            mixedDescriptionFields(sheet, quotation, row, rows);
          }
          if (fixedBindValueMap[rows.bindPath] === 0 || fixedBindValueMap[rows.bindPath]) {
            sheet.setValue(row + rows.row, rows.column, fixedBindValueMap[rows.bindPath]);
          }
          if (rows.bindPath === 'sumAmount') {
            fixedBindValueMap[rows.bindPath] = null;
            sheet.setFormula(row + rows.row, rows.column, '');
            sheet.setValue(row + rows.row, rows.column, '');
            libraryType === 'build' && synchronousStoreSumAmount(null);
          }
        }
      }
    }
  }

  // Dynamic fields
  for (const key in totalFieldBind) {
    if (Object.hasOwnProperty.call(totalFieldBind, key)) {
      const rows = totalFieldBind[key];
      if (!rows.bindPath) {
        let fieldName = key;
        if (!regChineseCharacter.test(rows.name)) {
          fieldName = rows.name;
        }

        console.log(row + rows.row, rows.column, rows);

        sheet.setFormula(row + rows.row, rows.column, '');
        sheet.setValue(row + rows.row, rows.column, '');
        fixedBindValueMap[fieldName] = null;
        if (libraryType === 'build') {
          store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
            path: [fieldName],
            value: null
          });
        }
      }
    }
  }
};

/**
 * Calculate calculated field value in Total
 * @param {*} sheet
 * @param {*} field
 * @param {*} rowIndex
 * @param {*} fixedBindValueMap
 * @param {*} fixedBindCellMap
 * @param {*} key
 * @param {*} cb
 * @param {*} totalBinds
 * @param {*} cbVal
 * @param {*} columnTotal
 * @param {*} columnTotalSum
 */
export const rowComputedField = (sheet, field, rowIndex, fixedBindValueMap, fixedBindCellMap, key, cb, totalBinds, cbVal, columnTotal, columnTotalSum) => {
  let fieldName = key;
  if (!regChineseCharacter.test(field.name)) {
    fieldName = field.name;
  }

  console.log(fieldName, '无bindpath字段');
  const fieldInfo = getFormulaFieldRowCol(field)
  if (ASSOCIATED_FIELDS_FORMULA_MAP[fieldName]) {
    if (fieldName === 'managementExpense') {
      const fieldFormula = managementExpenseAssignment(fieldName, fixedBindValueMap, fixedBindCellMap, columnTotalSum);
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (GenerateFieldsRow().includes(fieldName)) {
      const fieldFormula = totalBeforeTaxAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (fieldName === 'serviceCharge') {
      const fieldFormula = serviceChargeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (fieldName === 'totalServiceCharge') {
      const fieldFormula = totalServiceChargeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (fieldName === 'addTaxRateBefore') {
      const fieldFormula = addTaxRateBeforeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (fieldName === 'taxes') {
      const fieldFormula = taxesAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      const val = sheet.getValue(rowIndex + fieldInfo.row, fieldInfo.column);
      fieldFormula && cb(fieldFormula);
      if (val === 0 || val) {
        cbVal(val);
      }
    } else if (fieldName === 'serviceChargeFee') {
      serviceChargeFeeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
      console.error('该总计的组合类型字段(serviceChargeFee)已废弃，建议修改模板总计规则(不符合标准规则)!');
    } else if (fieldName === 'DXzje') {
      sheet.setFormatter(rowIndex + fieldInfo.row, fieldInfo.column, GeneratorUpperCaseFormatter());
      if (fixedBindCellMap.concessional) {
        sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fixedBindCellMap.concessional);
      } else {
        const fieldFormula = sumAmountFormula(fieldName, fixedBindCellMap, columnTotalSum);
        sheet.setFormula(rowIndex + fieldInfo.row, fieldInfo.column, fieldFormula);
      }
    } else {
      console.warn('在ASSOCIATED_FIELDS_FORMULA_MAP定义,但未存在过相关逻辑的字段', fieldName);
    }
  } else {
    console.warn('模板的总计block：识别出未在ASSOCIATED_FIELDS_FORMULA_MAP定义的字段', fieldName);
  }
};


// -----------------
/**
 * sumAmount assignment (运费 + 项目费用) + (管理费 + 服务费 + 税金)
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @returns
 */
// eslint-disable-next-line no-unused-vars
export const sumAmountFormula = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindKeys, fixedBindValueMap) => {
  const { taxes = null, managementExpense = null, serviceCharge = null, serviceChargeFee = null, freight = null, projectCost = null } = fixedBindCellMap;
  const fieldFormulas = [];
  if (columnTotalSum) {
    fieldFormulas.push(`(${columnTotalSum})`);
  }
  if (freight) {
    fieldFormulas.push(`(${freight})`);
  }
  if (projectCost) {
    fieldFormulas.push(`(${projectCost})`);
  }
  if (serviceCharge) {
    fieldFormulas.push(`(${serviceCharge})`);
  }
  if (serviceChargeFee) {
    fieldFormulas.push(`(${serviceChargeFee})`);
  }
  if (managementExpense) {
    fieldFormulas.push(`(${managementExpense})`);
  }
  if (taxes) {
    fieldFormulas.push(`(${taxes})`);
  }
  return fieldFormulas.join('+');
};

/**
 * managementExpense assignment
 * @param {*} fieldName
 * @param {*} fixedBindValueMap
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @returns
 */
const managementExpenseAssignment = (fieldName, fixedBindValueMap, fixedBindCellMap, columnTotalSum) => {
  const managementFeeCell = fixedBindCellMap.managementFee;
  const managementFee = managementFeeCell || fixedBindValueMap.managementFee;

  if (managementFee === 0 || managementFee) {
    if (columnTotalSum) {
      return `(${columnTotalSum}) * ${managementFee} / 100`;
    }
  } else {
    console.error('未检测到有埋点字段管理费率(managementFee)或管理费率字段(managementFee),无法计算管理费');
  }
  return '';
};

/**
 * The sum of all the accumulations【value + 】
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const totalBeforeTaxAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds) => {
  // eslint-disable-next-line no-unused-vars
  const { managementExpense = null, serviceCharge = null, taxes = null } = fixedBindCellMap;
  const totalBeforeFormula = totalBeforeAssignment(fieldName = 'totalBeforeTax', fixedBindCellMap, columnTotalSum, totalBinds);

  const rowNames = rowComputedFieldSort(totalBinds);
  const currentField = rowNames.findIndex((name) => name === fieldName);
  const managementExpenseIndex = rowNames.findIndex((name) => name === 'managementExpense');
  const serviceChargeIndex = rowNames.findIndex((name) => name === 'serviceCharge');
  // const taxesIndex = rowNames.findIndex((name) => name === 'taxes');

  const sum = [];
  if (totalBeforeFormula) {
    sum.push(totalBeforeFormula);
  }

  if (currentField > managementExpenseIndex && managementExpense) {
    sum.push(managementExpense);
  }
  if (currentField > serviceChargeIndex && serviceCharge) {
    sum.push(serviceCharge);
  }
  // if (currentField > taxesIndex && taxes) {
  //   sum.push(taxes);
  // }

  return sum.join('+');
};

/**
 * Specific fields + freight + projectCost
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const totalBeforeAssignment = (fieldName = 'totalBeforeTax', fixedBindCellMap, columnTotalSum, totalBinds = {}) => {
  const { freight = null, projectCost = null } = fixedBindCellMap;
  const rowNames = rowComputedFieldSort(totalBinds);

  const currentField = rowNames.findIndex((name) => name === fieldName);
  const freightIndex = rowNames.findIndex((name) => name === 'freight');
  const projectCostIndex = rowNames.findIndex((name) => name === 'projectCost');
  const sum = [];
  if (columnTotalSum) {
    sum.push(columnTotalSum);
  }
  if (currentField > freightIndex && freight) {
    sum.push(freight);
  }
  if (currentField > projectCostIndex && projectCost) {
    sum.push(projectCost);
  }

  return sum.join('+');
};

/**
 * serviceCharge assignment
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const serviceChargeAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap) => {
  const { rate = null, tax = null } = fixedBindCellMap;
  // eslint-disable-next-line no-unused-vars
  const { taxRate = null } = fixedBindValueMap;

  if (fieldName === 'serviceCharge') {
    const totalBeforeTaxFormula = totalBeforeTaxAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
    if (totalBeforeTaxFormula) {
      const rate_ = rate || fixedBindValueMap[NEW_OLD_FIELD_MAP.rate] || fixedBindValueMap.rate;
      if (rate_ === 0 || rate_) {
        return `(${totalBeforeTaxFormula}) * ${rate_} / 100`;
      } else {
        console.error('未检测到有埋点字段服务费率(rate)或服务费率字段(serviceCharge：旧；rate：新),无法计算服务费');
      }
    }
  } else if (fieldName === 'taxes') {
    const totalBeforeTaxFormula = totalBeforeTaxAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
    if (totalBeforeTaxFormula) {
      const tax_ = tax || fixedBindValueMap[NEW_OLD_FIELD_MAP.tax] || fixedBindValueMap.tax;
      if (tax_ === 0 || tax_) {
        return `(${totalBeforeTaxFormula}) * ${tax_} / 100`;
      } else {
        console.error('未检测到有埋点字段税率(tax)或税率字段(taxRate),无法计算税金');
      }
    }
  }

  return null;
};

/**
 * serviceChargeFee assignment
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 */
// eslint-disable-next-line no-unused-vars
const serviceChargeFeeAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds) => {
  // TODO 未开发功能
};

/**
 * totalServiceCharge assignment
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const totalServiceChargeAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds) => {
  const totalBeforeTaxFormula = totalBeforeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds);
  const { serviceCharge = null } = fixedBindCellMap;
  if (serviceCharge && totalBeforeTaxFormula) {
    return `(${totalBeforeTaxFormula}) + (${serviceCharge})`;
  } else {
    console.error('缺少服务费，无法计算!');
  }
  return null;
};

/**
 * taxes assignment
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const taxesAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap) => {
  const { tax = null, rate = null, totalServiceCharge = null } = fixedBindCellMap;
  if (!totalServiceCharge) {
    if (tax && rate) {
      const totalBeforeTaxFormula = totalBeforeTaxAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);

      if (totalBeforeTaxFormula) {
        const tax_ = tax || fixedBindValueMap[NEW_OLD_FIELD_MAP.tax] || fixedBindValueMap.tax;
        if (tax_ === 0 || tax_) {
          return `(${totalBeforeTaxFormula}) * ${tax_} / 100`;
        } else {
          console.error('未检测到有埋点字段税率(旧：rate；新:tax)或税率字段(旧：taxRate；新：tax),无法计算税金');
        }
      }
    } else if (tax) {
      return serviceChargeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds, fixedBindValueMap);
    }
  } else {
    const tax_ = tax || fixedBindValueMap.tax;
    if (tax_ === 0 || tax_) {
      return `(${totalServiceCharge}) * ${tax_} / 100`;
    } else {
      console.error('未检测到有埋点字段税率(tax)或税率字段(tax),无法计算税金');
    }
  }
  return null;
};

/**
 * addTaxRateBefore assignment
 * @param {*} fieldName
 * @param {*} fixedBindCellMap
 * @param {*} columnTotalSum
 * @param {*} totalBinds
 * @returns
 */
const addTaxRateBeforeAssignment = (fieldName, fixedBindCellMap, columnTotalSum, totalBinds) => {
  const { managementExpense = null, serviceCharge = null, serviceChargeFee = null } = fixedBindCellMap;
  const serviceChargeFormula = serviceChargeFee || serviceCharge;
  const sum1Formula = totalBeforeAssignment(fieldName, fixedBindCellMap, columnTotalSum, totalBinds);

  const formula = [];
  if (serviceChargeFormula) {
    formula.push(serviceChargeFormula);
  }

  if (managementExpense) {
    formula.push(managementExpense);
  }
  if (formula.length) {
    return `(${sum1Formula}) + ${formula.join('+')}`;
  }
  return '';
};

/**
 * The final price is based on the preferential price
 * @param {*} fixedBindCellMap 
 * @param {*} fixedBindValueMap 
 * @returns 
 */
export const finalPriceByConcessional = (fixedBindCellMap, fixedBindValueMap) => {
  if (fixedBindCellMap.concessional) {
    return fixedBindValueMap.concessionalRate;
  }
  return null;
};

/**
 * Adaptive row height settings
 * @param {*} sheet 
 * @param {*} row 
 * @param {*} rowsField 
 * @param {*} image 
 */
export const setAutoFitRow = (sheet, row, rowsField, image) => {
  const h = sheet.getRowHeight(row);
  if (image && image.height) {
    const maxH = Math.max(image.height, rowsField.height || 0);
    if (h < maxH) {
      sheet.setRowHeight(row, maxH);
    }
  } else {
    if (h < (rowsField.height || 0)) {
      sheet.setRowHeight(row, rowsField.height);
    }
  }
};

/**
 * Default row height settings
 * @param {*} sheet 
 * @param {*} row 
 * @param {*} rowsField 
 * @param {*} image 
 */
export const defaultAutoFitRow = (sheet, row, rowsField, image) => {
  if (image && image.height) {
    const maxH = Math.max(image.height, rowsField.height || 0);
    sheet.setRowHeight(row, maxH);
  } else if (rowsField && rowsField.height) {
    sheet.setRowHeight(row, rowsField.height);
  } else {
    sheet.getCell(row, -1).wordWrap(true);
    sheet.autoFitRow(row)
  }
};

/**
 * Obtain the row index in the table
 * @param {*} spread 
 * @returns 
 */
export const getTableRowIndex = (spread) => {
  const layout = new LayoutRowColBlock(spread);
  const { Tables } = layout.getLayout();
  const tableMapRows = layout.getTableMapRows(Tables);
  const rows = [];
  if (tableMapRows) {
    for (const key in tableMapRows) {
      if (Object.prototype.hasOwnProperty.call(tableMapRows, key)) {
        rows.push(...tableMapRows[key]);
      }
    }
  }
  return rows;
};

/*
 * @Author: Marlon
 * @Date: 2024-03-27 22:35:21
 * @Description:single - public
 */
import Decimal from 'decimal.js';
import * as GC from '@grapecity/spread-sheets';
import _ from 'lodash';
import API from 'api';
import { ResDatas } from '../utils/index';
import { GetUserCompany, imgUrlToBase64 } from 'utils';
import store from 'store';

import { GENERATE_FIELDS_NUMBER, DESCRIPTION_MAP, REGULAR } from './constant';
import { columnToNumber, PubGetRandomNumber, replacePlaceholders } from './public';

import { ShowCostPrice } from '../build-library/head'

import {
  templateRenderFlag,
  setCell,
  AddEquipmentImage,
  showSubTotal,
} from './parsing-template';
import { DEFINE_IDENTIFIER_MAP } from './identifier-template'
import { formatterPrice, getShowCostPrice } from './parsing-quotation';

import { SHOW_COST_PRICE_HEAD } from "store/quotation/mutation-types";

const NzhCN = require('nzh/cn');

/**
 * Row computed fields
 * @returns
 */
export const GenerateFieldsRow = () => {
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
 * Get the element configuration
 * @returns
 */
export const getProjectCfg = async () => {
  let elecfg = null;
  await API.qtElementType().then(res => {
    const Res = new ResDatas({ res }).init();
    elecfg = Res;
  });
  return elecfg;
};

/**
 * Get Item Number - Generate Rule
 * @param {*} eleId
 * @returns
 */
export const getProjectNumberRuler = async (eleId) => {
  let rules = [];
  if (eleId) {
    await API.qtProjectNameRule(eleId).then(res => {
      const Res = new ResDatas({ res }).init();
      if (Res) {
        rules = Res.filter((item) => item.flag);
      }
    });
  }
  return rules;
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
    quotationPdf: Res.quotationPdf// 导出pdf报价单的URL

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
export const PubGetTableStartRowIndex = () => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  return template.cloudSheet.top.rowCount;
};
/**
 * The total number of rows inserted into the table
 */
export const PubGetTableRowCount = (index = 0) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
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
export const classificationAlgorithms = (quotation, headers = []) => {
  const { template, mixRender, classType } = templateRenderFlag();
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
        subTotal = showSubTotal() ? subTotalCount : 0;
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
export const columnsTotal = (sheet, tableStartRowIndex, tableIndex, showRowTotal = false, columnComputed, subBindPath) => {
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
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
  const resources = PubGetTableRowCount(tableIndex);
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
export const rowComputedFieldSort = (totalBinds = {}) => {
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
export const setCellFormatter = (sheet, row, column) => {
  const format = formatterPrice();
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
export const SetComputedSubTotal = (sheet, columnTotalMap, bindPath) => {
  for (const key in columnTotalMap) {
    if (Object.hasOwnProperty.call(columnTotalMap, key)) {
      if (columnTotalMap[key] && columnTotalMap[key].formula) {
        const formula = columnTotalMap[key].formula.split('+');
        const subTotalRowIndex = formula[formula.length - 1].substring(1);
        if (bindPath && bindPath[key]) {
          const columnIndex = bindPath[key].columnHeader;
          setCellFormatter(sheet, Number(subTotalRowIndex), columnToNumber(columnIndex) - 1);
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
const tableAddImage = (spread, table, classIndex, insertTableIndex, classRow, subTotal, allowMove, allowResize, isLocked) => {
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
          AddEquipmentImage(spread, item.id, base64, startRow, allowMove, allowResize, isLocked);
        });
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
export const renderSheetImage = (spread, tableStartRowIndex, allowMove, allowResize, isLocked) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const resourceViews = quotation.conferenceHall.resourceViews;
  const { classRow, subTotal } = classificationAlgorithms(quotation);

  let insertTableIndex = tableStartRowIndex;
  for (let index = 0; index < resourceViews.length; index++) {
    if (resourceViews[index].resources) {
      if (index === 0) {
        tableAddImage(spread, resourceViews[index].resources, index, insertTableIndex, classRow, subTotal, allowMove, allowResize, isLocked);
        insertTableIndex = insertTableIndex + classRow + resourceViews[index].resources.length;
      } else {
        tableAddImage(spread, resourceViews[index].resources, index, insertTableIndex, classRow, subTotal, allowMove, allowResize, isLocked);
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
  ShowCostPrice(spread)
}

/**
 * Obtain the classification type of the template
 * @returns 
 */
export const getTemplateClassType = () => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const template = store.getters['quotationModule/GetterQuotationWorkBook'];
  const resourceViews = quotation.conferenceHall.resourceViews;

  if (template.templateClassIdentifier) {
    if (Object.keys(DEFINE_IDENTIFIER_MAP).includes(template.templateClassIdentifier)) {
      return template.templateClassIdentifier
    } else {
      console.error('模板分类标识符不存在,无法确定模板分类类型【模板错误】');
    }
  } else {
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

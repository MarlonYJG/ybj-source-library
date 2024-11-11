/*
 * @Author: Marlon
 * @Date: 2024-03-27 22:36:48
 * @Description:multiple - public
 */
import { v4 as uuidv4 } from 'uuid';
import store from 'store';
// import Decimal from '../lib/decimal/decimal.min.js';
import * as GC from '@grapecity/spread-sheets';
import _ from '../lib/lodash/lodash.min.js';
import { GetUserCompany, FormatDate } from '../utils/index';

import { QuotationInitData } from './constant';

import { rootWorkBook } from './core';
import { SetDataSource, SpreadLocked } from './sheetWorkBook';
import { InitBindPath, renderFinishedAddImage } from './single-table';

import { getAllSheet } from './parsing-quotation';
import { getTrunkTemplate } from './parsing-template';

import { CenterSheetRender } from '../parsing-library/multiple-table';
import { CenterSheetBuild } from "../build-library/multiple-table";


/**
 * Initialize the total score table data
 * @param {*} Res 
 * @param {*} type 
 * @returns 
 */
export const multipleTableSyncStore = (Res, type) => {
  let company = {
    name: '',
    createUserId: '',
    address: '',
    website: '',
    fax: '',
    tel: '',
    industry: '',
    state: '',
    logoURL: ''
  }
  let projectNumber = '';
  let belongs = '';
  let projectName = '';
  let projectId = '';
  if (type !== 'mobile') {
    if (GetUserCompany()) {
      company = GetUserCompany();
    }
    if (store && store.getters && store.getters['projectModule/GetterProjectInit']) {
      const projectInit = store.getters['projectModule/GetterProjectInit'];
      projectNumber = projectInit.projectNumber;
      projectName = projectInit.name;
      projectId = projectInit.id;
      belongs = projectInit.belongs;
    }
  }
  const customer = Res.customer;
  const quotationDefault = {
    ...Res,
    id: Res.id,
    title: Res.title,
    logo: Res.logo,
    seal: Res.seal,
    quotationImage: Res.quotationImage,
    quaLogos: Res.quaLogos || [],
    state: Res.state,
    leaderId: belongs || Res.belongs,
    belongs: Res.belongs,
    belongsEmail: Res.belongsEmail,
    phone: Res.phone,
    createTime: Res.createTime,
    updateTime: Res.updateTime,
    needApproval: Res.needApproval || 'NO',
    storePhone: Res.storePhone,


    projectId: Res.projectId || projectId,
    name: Res.name || projectName,
    projectName: Res.projectName,
    projectType: Res.projectType,
    projectNumber: Res.projectNumber || projectNumber,

    projectManager: Res.projectManager,
    projectManagerPhone: Res.projectManagerPhone,


    customer,

    companyAddress: Res.companyAddress,
    companyEmail: Res.companyEmail,
    companyFax: Res.companyFax,
    companyName: Res.companyName || company.name,
    companyPhone: Res.companyPhone,
    companyWebsite: Res.companyWebsite,

    company: {
      companyName: Res.companyName || company.name,
      createUserId: company.createUserId,
      companyAddress: Res.companyAddress || company.address,
      companyEmail: Res.companyEmail,
      logoURL: company.logoURL,
      companyPhone: Res.companyPhone || company.tel,
      companyWebsite: Res.companyWebsite || company.website,
      companyFax: Res.companyFax || company.fax,
      industry: company.industry,
      state: company.state
    },

    conferenceHall: Res.conferenceHall,
    resources: Res.resources || [],

    parallelSessions: [],
    preferentialWay: Res.preferentialWay,
    noImgTemplate: Res.noImgTemplate,
    designerPhone: Res.designerPhone,
    designer: Res.designer,
    engineerPhone: Res.engineerPhone,
    engineer: Res.engineer,


    willPay: Res.willPay,
    DXdeposit: Res.DXdeposit,
    deposit: Res.deposit,
    DXwillPay: Res.DXwillPay,
    DXzje: Res.DXzje,
    capitalizeTotalAmount: Res.capitalizeTotalAmount,
    showCost: Res.showCost,
    totalAmount: Res.totalAmount,
    sumAmount: Res.sumAmount,
    dxsumAmount: Res.dxsumAmount,
    freight: Res.freight,
    projectCost: Res.projectCost,
    managementFee: Res.managementFee,
    managementExpense: Res.managementExpense,
    rate: Res.rate,
    serviceCharge: Res.serviceCharge,
    taxRate: Res.taxRate,
    tax: Res.tax,
    discount: Res.discount,
    concessionalRate: Res.concessionalRate,
    concessionalType: (Res.concessionalType === 0 || Res.concessionalType) ? Res.concessionalType.toString() : '0',
    concessionalDiscount: Res.concessionalDiscount,

    excelJson: Res.excelJson,
    templateType: Res.templateType,
    remark: Res.remark,
    exportName: Res.exportName,
    subheadingOne: Res.subheadingOne,
    subheadingTwo: Res.subheadingTwo,
    priceAdjustment: Res.priceAdjustment,
    priceStatus: Res.priceStatus,
    extFields: Res.extFields,
    image: Res.image,
    templateId: Res.templateId,
    isInt: Res.isInt,
    priceType: Res.priceType,


    haveExport: Res.haveExport,
    quotationExcel: Res.quotationExcel,
    quotationPdf: Res.quotationPdf,


    config: typeof Res.config === 'object' ? Res.config : ((Res.config && typeof Res.config === 'string') ? JSON.parse(Res.config) : {
      startAutoFitRow: false
    })
  }
  // 时间格式化
  const { approachDate = '', approachTime = '', fieldWithdrawalDate = '', fieldWithdrawalTime = '', startDate = '', startTime = '' } = quotationDefault.conferenceHall;
  if (approachDate) {
    quotationDefault.conferenceHall.approachDate = FormatDate(approachDate, 'YYYY-MM-DD');
  }
  if (approachTime) {
    quotationDefault.conferenceHall.approachTime = FormatDate(approachTime, 'YYYY-MM-DD');
  }
  if (fieldWithdrawalDate) {
    quotationDefault.conferenceHall.fieldWithdrawalDate = FormatDate(fieldWithdrawalDate, 'YYYY-MM-DD');
  }
  if (fieldWithdrawalTime) {
    quotationDefault.conferenceHall.fieldWithdrawalTime = FormatDate(fieldWithdrawalTime, 'YYYY-MM-DD');
  }
  if (startDate) {
    quotationDefault.conferenceHall.startDate = FormatDate(startDate, 'YYYY-MM-DD');
  }
  if (startTime) {
    quotationDefault.conferenceHall.startTime = FormatDate(startTime, 'YYYY-MM-DD');
  }

  quotationDefault.name = quotationDefault.name || quotationDefault.projectName;
  // 扩展字段特殊处理
  if (quotationDefault.extFields) {
    if (typeof quotationDefault.extFields === 'string') {
      quotationDefault.extFields = JSON.parse(quotationDefault.extFields);
    }
  }

  let resourceViews = _.cloneDeep(quotationDefault.conferenceHall.resourceViews);
  resourceViews.forEach(sheet => {
    if (sheet.resourceViews && sheet.resourceViews.length) {
      sheet.resourceViews.forEach(classItem => {
        if (!classItem.resources || !classItem.resources.length) {
          sheet.resourceViews.splice(sheet.resourceViews.indexOf(classItem), 1);
        }
        classItem.resources.forEach((item, i) => {
          item.sort = i + 1;
          if (!item.imageId) {
            item.imageId = uuidv4();
          }
        });
      });
    }
  });

  quotationDefault.conferenceHall.resourceViews = resourceViews;

  // resourceSort(quotationDefault)

  console.log(quotationDefault, '====响应数据处理-并同步至store');
  return quotationDefault;
};

// /**
//  * Product list added sorting and init resourceViewsMap
//  * @param {*} quotation
//  */
// const resourceSort = (quotation) => {
//   const resourceViewsMap = {};
//   const resourceViews = quotation.conferenceHall.resourceViews;
//   for (let index = 0; index < resourceViews.length; index++) {
//     resourceViews[index].sort = index + 1;
//     if (resourceViews[index].resources) {
//       resourceViews[index].resources.forEach((resource, i) => {
//         resource.sort = i + 1;
//       });
//     }
//   }
//   resourceViews.forEach(item => {
//     resourceViewsMap[item.resourceLibraryId] = item;
//   });
//   quotation.conferenceHall.resourceViewsMap = resourceViewsMap;
// };

// 分表排序
// const branchSort = (quotation) => {
//   //  TODO 底部分表排序时，需要同步总表中的分表数据
// };

/**
 * The template index corresponding to the table sharding
 * @param {*} quotation 
 * @param {*} templates 
 * @returns 
 */
export const getSheetTemplateIndexs = (quotation, templates) => {
  const trunkTempIndex = [];
  const trunks = quotation.resources || [];

  if (templates && templates.length === 1) {
    for (let i = 0; i < trunks.length; i++) {
      trunkTempIndex.push(0);
    }
  } else if (templates && templates.length > 1) {
    for (let i = 0; i < trunks.length; i++) {
      trunkTempIndex.push(trunks[i].templateIndex || 0);
    }
  } else {
    console.error('The data of the table sharding template is abnormal');
  }
  return trunkTempIndex;
}

/**
 * Build a mapping table of quotes to templates
 * @param {*} dataSource 
 * @param {*} template 
 * @returns 
 */
export const templateMap = (dataSource, template) => {
  const trunks = getAllSheet(dataSource);
  if (trunks.length) {
    const trunkTemplate = getTrunkTemplate(template);
    const templateIndexs = getSheetTemplateIndexs(dataSource, trunkTemplate);

    const templateMap = {};
    trunks.forEach((trunk, i) => {
      const templateItem = _.cloneDeep(trunkTemplate[templateIndexs[i]]);
      templateItem.sheets['分表'].name = trunk.name;
      templateMap[trunk.name] = templateItem;
    });

    for (const key in templateMap) {
      if (Object.prototype.hasOwnProperty.call(templateMap, key)) {
        const sourceBuild = _.cloneDeep(template);
        sourceBuild.sheets[key] = templateMap[key].sheets['分表'];
        sourceBuild.name = key;
        sourceBuild.cloudSheet = templateMap[key];
        templateMap[key] = sourceBuild;
      }
    }

    return templateMap;
  }
  return null;
}

/**
 * Create a single sharded table data
 * @param {*} spread 
 * @param {*} dataSource 
 */
const buildSheetData = (spread, dataSource) => {
  const sheet = spread.getActiveSheet();
  const sheetName = sheet.name();
  const sheets = dataSource.conferenceHall.resourceViews;
  const sheetQuotation = sheets.filter(item => item.name === sheetName);
  console.log(sheetQuotation, 'sheetQuotation');
  const quotation = QuotationInitData();
  if (sheetQuotation.length) {
    const { extFields = {}, resourceViews = [] } = sheetQuotation[0];
    for (const key in extFields) {
      if (Object.prototype.hasOwnProperty.call(extFields, key)) {
        quotation[key] = extFields[key];
      }
    }
    quotation.conferenceHall.resourceViews = resourceViews;
    quotation.conferenceHall.resourceViewsMap = resourceViews.reduce((acc, cur) => {
      acc[cur.resourceLibraryId] = cur;
      return acc;
    }, {});
  }

  console.log(quotation, 'initData quotation');
  return quotation;
};

/**
 * bingn event
 * @param {*} spread 
 * @param {*} quotation 
 * @param {*} template 
 * @param {*} isCompress 
 * @param {*} type 
 */
const OnEventBind = (spread, quotation, template, isCompress, type, templateMapData) => {
  spread.bind(GC.Spread.Sheets.Events.ActiveSheetChanged, function (sender, args) {
    console.log(sender, args);
    const trunks = getAllSheet(quotation);
    const sheet = spread.getActiveSheet();
    const sheetName = sheet.name();
    const sheetNames = trunks.map((item) => item.name);

    if (sheetNames.includes(sheetName)) {
      // TODO 分表
      InitSheetRender(spread, quotation, templateMapData[sheetName], isCompress, type);
    } else {
      // SetDataSource(sheet, quotation);
      // InitBindPath(spread, template, quotation);
    }
  })
}

/**
 * Draw a summary table
 * @param {*} spread 
 * @param {*} template 
 * @param {*} quotation 
 * @param {*} isCompress 
 * @param {*} type 
 */
export const InitMainSheetRender = (spread, template, quotation, isCompress, type) => {
  const templateMapData = templateMap(quotation, template);
  console.log(templateMapData, 'templateMapData');
  rootWorkBook._setTemplateMap(templateMapData);
  OnEventBind(spread, quotation, template, isCompress, type, templateMapData);
  const sheet = spread.getActiveSheet();
  SetDataSource(sheet, quotation);
  InitBindPath(spread, template, quotation);
}

/**
 * Draw a table split
 * @param {*} spread 
 * @param {*} dataSource 
 * @param {*} template 
 * @param {*} isCompress 
 * @param {*} type 
 */
const InitSheetRender = (spread, dataSource, template, isCompress, type) => {
  const initData = buildSheetData(spread, dataSource);
  rootWorkBook._setActiveQuotation(initData);

  if (type === 'parsing') {
    const sheet = spread.getActiveSheet();
    SetDataSource(sheet, initData);
    InitBindPath(spread, template, initData);
    renderFinishedAddImage(spread, template, initData);


    CenterSheetRender(spread, template, initData, isCompress);
    SpreadLocked(spread);
  } else if (type === 'build') {
    CenterSheetBuild(spread, template, initData, isCompress);
  }
};


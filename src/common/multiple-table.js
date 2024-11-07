/*
 * @Author: Marlon
 * @Date: 2024-03-27 22:36:48
 * @Description:multiple - public
 */
import { v4 as uuidv4 } from 'uuid';
// import Decimal from '../lib/decimal/decimal.min.js';
// import * as GC from '@grapecity/spread-sheets';
import _ from '../lib/lodash/lodash.min.js';
import store from 'store';
import { GetUserCompany, FormatDate } from '../utils/index';
export const multipleTableSyncStore = (Res, type) => {
  console.log(Res);
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

  // 报价单中的项目名称字段不统一问题
  quotationDefault.name = quotationDefault.name || quotationDefault.projectName;
  // 扩展字段特殊处理
  if (quotationDefault.extFields) {
    if (typeof quotationDefault.extFields === 'string') {
      quotationDefault.extFields = JSON.parse(quotationDefault.extFields);
    }
  }
  // 去除空数据
  let resourceViews = _.cloneDeep(quotationDefault.conferenceHall.resourceViews);
  resourceViews = resourceViews.filter((item) => { return item.resources && item.resources.length; });
  resourceViews.forEach((item, i) => {
    item.sort = i + 1;
    item.resources.forEach((resource) => {
      if (!resource.imageId) {
        resource.imageId = uuidv4();
      }
    });
  });
  quotationDefault.conferenceHall.resourceViews = resourceViews;

  resourceSort(quotationDefault)

  console.log(quotationDefault, '====响应数据处理-并同步至store');
  return quotationDefault;
};

/**
 * Product list added sorting and init resourceViewsMap
 * @param {*} quotation
 */
export const resourceSort = (quotation) => {
  const resourceViewsMap = {};
  const resourceViews = quotation.conferenceHall.resourceViews;
  for (let index = 0; index < resourceViews.length; index++) {
    resourceViews[index].sort = index + 1;
    if (resourceViews[index].resources) {
      resourceViews[index].resources.forEach((resource, i) => {
        resource.sort = i + 1;
      });
    }
  }
  resourceViews.forEach(item => {
    resourceViewsMap[item.resourceLibraryId] = item;
  });
  quotation.conferenceHall.resourceViewsMap = resourceViewsMap;
};

// 分表排序
export const branchSort = (quotation) => {
  //  TODO 底部分表排序时，需要同步总表中的分表数据
};

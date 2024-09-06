/*
 * @Author: Marlon
 * @Date: 2024-09-05 10:49:19
 * @Description: mobile
 */
import _ from 'lodash'
import * as GC from '@grapecity/spread-sheets'
import { FormatDate } from '../utils/index'
import { CombinationType } from '../common/combination-type'

// 数据处理并同步到 store
const FormaterQuotationInfo = (conferenceHall, map = {}) => {
  let quotation = {};

  if (conferenceHall) {
    const { approachDate = '', approachTime = '', fieldWithdrawalDate = '', fieldWithdrawalTime = '', startDate = '', startTime = '' } = conferenceHall;
    if (approachDate) {
      conferenceHall.approachDate = FormatDate(approachDate, 'YYYY-MM-DD');
    }
    if (approachTime) {
      conferenceHall.approachTime = FormatDate(approachTime, 'YYYY-MM-DD');
    }
    if (fieldWithdrawalDate) {
      conferenceHall.fieldWithdrawalDate = FormatDate(fieldWithdrawalDate, 'YYYY-MM-DD');
    }
    if (fieldWithdrawalTime) {
      conferenceHall.fieldWithdrawalTime = FormatDate(fieldWithdrawalTime, 'YYYY-MM-DD');
    }
    if (startDate) {
      conferenceHall.startDate = FormatDate(startDate, 'YYYY-MM-DD');
    }
    if (startTime) {
      conferenceHall.startTime = FormatDate(startTime, 'YYYY-MM-DD');
    }

    if (Object.keys(map).length) {
      for (const key in map) {
        if (Object.hasOwnProperty.call(map, key)) {
          quotation[map[key]] = conferenceHall[key];
        }
      }
    } else {
      quotation = conferenceHall;
    }
  }
  return quotation;
};

// 数据格式处理
const DataProcessing = (quotation) => {
  // 报价单中的项目名称字段不统一问题
  quotation.name = quotation.name || quotation.projectName;

  // 扩展字段特殊处理
  if (quotation.extFields) {
    if (typeof quotation.extFields === 'string') {
      quotation.extFields = JSON.parse(quotation.extFields);
    }
  }
  // 去除空数据
  let resourceViews = _.cloneDeep(quotation.conferenceHall.resourceViews);
  resourceViews = resourceViews.filter((item) => { return item.resources && item.resources.length; });
  resourceViews.forEach((item, i) => {
    item.sort = i + 1;
  });

  quotation.conferenceHall.resourceViews = resourceViews;

  return quotation;
}

const singleTableSyncStore = (Res) => {
  const customer = Res.customer;
  const quotationDefault = {
    id: Res.id, // 报价单id
    title: Res.title, // 报价单标题
    logo: Res.logo, // 绑定公司的logo
    seal: Res.seal, // 绑定公司的印章
    quotationImage: Res.quotationImage, // 报价单的图片(旧版字段)
    quaLogos: Res.quaLogos || [], // 报价单多logo
    state: Res.state,
    leaderId: Res.belongs, // 负责人id
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
    projectNumber: Res.projectNumber,
    // 客户
    projectManager: Res.projectManager,
    projectManagerPhone: Res.projectManagerPhone,

    // 客户
    customer,
    // 公司信息
    companyAddress: Res.companyAddress,
    companyEmail: Res.companyEmail,
    companyFax: Res.companyFax,
    companyName: Res.companyName,
    companyPhone: Res.companyPhone,
    companyWebsite: Res.companyWebsite,

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

}

/**
 * Product list added sorting
 * @param {*} quotation
 */
const resourceSort = (quotation) => {
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
 * Initialize data source for mobile mode
 * @param {*} dataSource 
 * @param {*} template 
 * @returns 
 */
export const initMobileData = (dataSource, template) => {
  if (!template) return;
  let quotation = _.cloneDeep(dataSource);
  if (dataSource) {
    if (dataSource.conferenceHall) {
      FormaterQuotationInfo(quotation);
    }

    DataProcessing(quotation)
    quotation = singleTableSyncStore(quotation);
    resourceSort(quotation)

    // 构建resourceViewsMap
    const { resourceViews = [] } = quotation.conferenceHall;
    quotation.conferenceHall.resourceViewsMap = {};
    resourceViews.forEach(item => {
      quotation.conferenceHall.resourceViewsMap[item.resourceLibraryId] = item;
    });

    CombinationType(quotation, template)
  }

  return {
    quotation,
    template
  }
}

/**
 * Start mobile mode
 * @param {*} spread 
 * @returns 
 */
export const startMobileMode = (spread) => {
  if (!spread) return
  spread.options.useTouchLayout = true;
  spread.options.protectionOptions = {
    allowSelectLockedCells: false,
    allowSelectUnlockedCells: false,
  };
  spread.options.allowInvalidFormula = true;
  spread.options.scrollbarAppearance = GC.Spread.Sheets.ScrollbarAppearance.mobile;
  spread.options.allowInvalidFormula = true;
  spread.options.allowUndo = true;
  spread.options.allowContextMenu = false;
  spread.options.newTabVisible = false;
  spread.options.scrollByPixel = true;
};


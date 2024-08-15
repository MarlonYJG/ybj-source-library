/*
 * @Author: Marlon
 * @Date: 2024-05-07 14:03:25
 * @Description: constant rule
 */

export const REGULAR = {

  chineseCharacters: new RegExp('[\\u4E00-\\u9FFF]+', 'g')
};

/**
 * Project Initialization field
 */
export const PROJECT_INIT_DATA = {
  id: '', // 项目Id
  name: '', // 项目名称
  belongs: '', // 负责人id
  projectNumber: '', // 项目编号
  startDate: new Date(),
  approachDate: new Date(),
  projectRequirements: ''
};

/**
 * Quote initialization field
 */
export const QUOTATION_INIT_DATA = {
  projectNumber: '',
  id: null, // 报价单id
  title: '', // 报价单标题
  state: 'NEWLY_BUILD', // 报价单状态 ['NEWLY_BUILD', 'HAVE_IN_HAND', 'SUSPEND', 'COMPLETE', 'GIVE_UP', 'END']
  name: '', // 项目名称
  projectName: '', // 项目名称
  belongs: '', // 用户
  belongsEmail: '',
  leaderId: '', // 负责人id
  phone: '', // 用户
  createTime: null, // 创建时间
  updateTime: null, // 更新时间
  needApproval: 'NO', // 是否需要审批
  resources: [], // 存储总分表的数据
  haveExport: false,
  customer: { // 客户信息
    id: null, // 客户 id
    name: null, // 客户 名称
    address: '', // 客户 地址
    contactInformation: null, // 联系方式
    personInCharge: null // 客户 公司负责人名称
  },
  conferenceHall: { // 主会场信息
    address: null, // 场地地址
    resourceViews: [], // 会场资源
    approachDate: new Date(), // 进场日期
    approachTime: new Date(), // 进场时间
    fieldWithdrawalDate: new Date(), // 撤场日期
    fieldWithdrawalTime: new Date(), // 撤场时间
    startDate: new Date(), // 开始日期
    startTime: new Date() // 开始时间
  },
  parallelSessions: [], // 分会场信息
  preferentialWay: 'CONCESSIONAL_RATE', // 优惠方式 ['DISCOUNT', 'CONCESSIONAL_RATE']
  taxRate: null, // 税率
  discount: null, // 折扣
  serviceCharge: null, // 服务费
  freight: null, // 运费
  concessionalRate: null, // 优惠价
  excelJson: null,
  templateId: null,
  templateType: 'PLATFORM',
  remark: '', // 备注
  exportName: '', // 导出名称
  subheadingOne: null, // 副标题一
  subheadingTwo: null, // 副标题二
  priceAdjustment: 1, // 单价调整比例
  calculateMix: 3, // 计算组合
  priceStatus: null, // 价格设置
  extFields: {}, // 预置字段
  image: '', // 报价单图片
  company: {
    companyName: '', // 公司名
    belongs: '', // 用户名
    phone: '', // 负责人电话
    companyAddress: '', // 公司地址
    companyEmail: '', // 公司邮箱，
    companyPhone: '', // 公司电话
    companyWebsite: '', // 公司网址
    companyFax: '' // 公司传真
  }
};

/**
 * Define the fields
 */
export const GENERATE_FIELDS_NUMBER = {
  total: 15,
  totalBeforeTax: 7
};

/**
 * Total Combined
 */
export const TOTAL_COMBINED_MAP = {
  0: 'taxRate_concessionalRate',
  1: 'taxRate',
  2: 'concessionalRate',
  3: 'initTotal',
  4: 'serviceCharge',
  5: 'serviceCharge_concessionalRate',
  6: 'serviceCharge_taxRate',
  7: 'serviceCharge_taxRate-concessionalRate',
  8: 'freight',
  9: 'freight_concessionalRate',
  10: 'taxRate_freight',
  11: 'taxRate_freight_concessionalRate',
  12: 'serviceCharge_freight',
  13: 'serviceCharge_freight_concessionalRate',
  14: 'taxRate_serviceCharge_freight',
  15: 'taxRate_serviceCharge_freight_concessionalRate'
};

/**
 * Total block Map(All)
 */
export const ASSOCIATED_FIELDS_FORMULA_MAP = {
  columnTotalSum: {
    label: '所有产品的合计累计之和(金额)',
    formula: '{{columnTotalSum}}'
  },

  freight: {
    label: '运费',
    formula: '{{freight}}'
  },
  projectCost: {
    label: '项目费用',
    formula: '{{projectCost}}'
  },

  totalBeforeTax1: {
    label: '无意义字段',
    formula: ''
  },
  totalBeforeTax2: {
    label: '无意义字段',
    formula: ''
  },
  totalBeforeTax3: {
    label: '无意义字段',
    formula: ''
  },
  totalBeforeTax4: {
    label: '无意义字段',
    formula: ''
  },
  totalBeforeTax5: {
    label: '无意义字段',
    formula: ''
  },
  totalBeforeTax6: {
    label: '无意义字段',
    formula: ''
  },

  managementFee: {
    label: '管理费率',
    formula: '{{managementFee}}'
  },
  managementExpense: {
    label: '管理费',
    formula: '{{columnTotalSum}} * {{managementFee}} / 100'
  },
  managementTotalAfter: {
    label: '含管理费的总计',
    formula: '{{totalBeforeTax}} + {{managementExpense}}'
  },

  rate: {
    label: '服务费利率',
    formula: '{{rate}}'
  },
  serviceCharge: {
    label: '服务费',
    formula: '{{totalBeforeTax}} * {{rate}} /100'
  },
  serviceChargeFee: {
    label: '服务费',
    formula: '{{totalBeforeTax}} * {{rate}} /100'
  },
  totalServiceCharge: {
    label: '合计(含服务费)',
    formula: '{{totalBeforeTax}} + {{serviceCharge}}'
  },

  totalBeforeTax: {
    label: '合计',
    formula: '{{columnTotalSum}} + [{{freight}}] + [{{projectCost}}] + [{{managementExpense}}] + [{{serviceCharge}}]'
  },

  addTaxRateBefore: {
    label: '未含税之前的总计',
    formula: '{{totalBeforeTax}}'
  },
  tax: {
    label: '税率',
    formula: '{{tax}}'
  },
  taxes: {
    label: '税金',
    formula: '{{addTaxRateBefore}} * {{tax}}/100'
  },
  totalAfterTax: {
    label: '税后合计',
    formula: '{{addTaxRateBefore}} + {{taxes}}'
  },

  concessional: {
    label: '优惠后合计',
    formula: '{{concessional}}'
  }
};

/**
 * field description
 */
export const DESCRIPTION_MAP = {
  managementRateDescription: {
    label: '用于将管理费与管理费率放在一起显示，支持描述信息。',
    key: 'managementRateDescription',
    percentage: 'managementFee'
  },
  serviceChargeDescription: {
    label: '用于将服务费与服务费率放在一起显示，支持描述信息。',
    key: 'serviceChargeDescription',
    percentage: 'serviceCharge|rate' // 新：rate
  },
  taxRateDescription: {
    label: '用于将税金与税率放在一起显示，支持描述信息。',
    key: 'taxRateDescription',
    percentage: 'taxRate|tax'// 新：tax
  }
};

/**
 * Price setup field mapping table
 */
export const PRICE_SET_MAP = {
  0: 'unitPrice',
  1: 'unitPrice1',
  2: 'unitPrice2',
  'unitPrice': 0,
  'unitPrice1': 1,
  'unitPrice2': 2
};
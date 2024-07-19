/*
 * @Author: Marlon
 * @Date: 2024-06-21 23:31:27
 * @Description:
 */
export const MENU_TOTAL = [
  {
    label: '税率',
    value: 'tax',
    disable: false
  },
  {
    label: '服务费',
    value: 'rate',
    disable: false
  },
  {
    label: '优惠价',
    value: 'concessional',
    disable: false
  },
  {
    label: '运费',
    value: 'freight',
    disable: false
  },
  {
    label: '管理费',
    value: 'managementFee',
    disable: false
  },
  {
    label: '项目费用',
    value: 'projectCost',
    disable: false
  }
];
export const MENU_DELETE = [
  {
    label: '清空产品',
    value: 'clearProduct',
    disable: false
  },
  // {
  //   label: '分类',
  //   value: 'delClass',
  //   disable: false
  // },
  {
    label: '产品',
    value: 'delProduct',
    disable: false
  }
];

export const NEW_OLD_FIELD_MAP = {
  tax: 'taxRate',
  rate: 'serviceCharge',
  concessional: 'concessionalRate',
  freight: 'freight',
  managementFee: 'managementFee',
  projectCost: 'projectCost'
};

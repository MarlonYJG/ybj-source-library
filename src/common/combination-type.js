
/*
 * @Author: Marlon
 * @Date: 2020-03-14 11:40:08
 * @Description: Combination type
 */

/**
 * Combination type
 * @param {*} quotation
 * @param {*} template
 */
export const CombinationType = (quotation, template) => {
  Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 0);
  Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 1);

  Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 6);
  Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 7);

  Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 10);
  Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 11);

  Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 14);
  Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 15);

  Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 18); // 1/18
  Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 19); // 0/19

  Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 22);// 6/22
  Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 23);// 7/23

  Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 26);// 10/26
  Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 27);// 11/27

  Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 30);// 14/30
  Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 31);// 15/31

  !Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 2);
  !Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 3);
  !Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 4);
  !Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 5);

  !Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 8);
  !Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 9);

  !Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 12);
  !Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 13);

  !Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 16); //  3/16
  !Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 17); // 2/17

  !Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 20);// 4/20
  !Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 21);// 5/21

  !Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 24);// 8/24
  !Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 25);// 9/25

  !Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 28);// 12/28
  !Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost) && (template.cloudSheet.total.select = 29);// 13/29

  !Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && Number(quotation.projectCost) && (template.cloudSheet.total.select = 32);// 3/32
  !Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && Number(quotation.projectCost) && (template.cloudSheet.total.select = 48); // 16/48
};

export const CombinationTypeBuild = (quotation) => {
  if (Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 0;
  } else if (Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 1;
  } else if (Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 6;
  } else if (Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 7;
  } else if (Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 10;
  } else if (Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 11;
  } else if (Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 14;
  } else if (Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 15;
  } else if (Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 18;
  } else if (Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 19;
  } else if (Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 22;
  } else if (Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 23;
  } else if (Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 26;
  } else if (Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 27;
  } else if (Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 30;
  } else if (Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 31;
  } else if (!Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 2;
  } else if (!Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 3;
  } else if (!Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 4;
  } else if (!Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 5;
  } else if (!Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 8;
  } else if (!Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 9;
  } else if (!Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 12;
  } else if (!Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && !Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 13;
  } else if (!Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 16;
  } else if (!Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 17;
  } else if (!Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 20;
  } else if (!Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 21;
  } else if (!Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 24;
  } else if (!Number(quotation.taxRate) && Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 25;
  } else if (!Number(quotation.taxRate) && !Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 28;
  } else if (!Number(quotation.taxRate) && Number(quotation.concessionalRate) && Number(quotation.serviceCharge) && Number(quotation.freight) && Number(quotation.managementFee) && !Number(quotation.projectCost)) {
    return 29;
  } else if (!Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && !Number(quotation.managementFee) && Number(quotation.projectCost)) {
    return 32;
  } else if (!Number(quotation.taxRate) && !Number(quotation.concessionalRate) && !Number(quotation.serviceCharge) && !Number(quotation.freight) && Number(quotation.managementFee) && Number(quotation.projectCost)) {
    return 48;
  }
};

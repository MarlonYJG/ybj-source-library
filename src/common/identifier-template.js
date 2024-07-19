/*
 * @Author: Marlon
 * @Date: 2024-07-09 17:01:22
 * @Description: Template identifier
 */
import _ from 'lodash';
import store from 'store';

import { GetColumnComputedTotal } from '../build-library/single-table';
import { PubGetTableStartRowIndex, PubGetTableRowCount } from './single-table';

export class IdentifierTemplate {
  constructor(sheet, name) {
    this.sheet = sheet;
    this.instanceName = name;
    this.builtInIdMap = {
      '2c91e6548b931749018b93bac6fb001e': 'BH'
    }
  }

  /**
   * Identifier:truckage
   * @param {*} totalField
   * @param {*} row
   * @param {*} fixedBindValueMap
   */
  truckageFreight(totalField, row, fixedBindValueMap) {
    for (const key in totalField.bindPath) {
      if (Object.hasOwnProperty.call(totalField.bindPath, key)) {
        const rows = totalField.bindPath[key];
        if (rows.bindPath === 'freight') {
          this.sheet.setValue(row + rows.row + 1, rows.column, fixedBindValueMap[key]);
        }
      }
    }
  }

  truckageRenderTotal(quotation) {
    const columnTotal = GetColumnComputedTotal(this.sheet);
    if (columnTotal.length) {
      const columnTotalMap = columnTotal[0];
      const subRow = PubGetTableStartRowIndex() + PubGetTableRowCount();
      this.sheet.suspendPaint();
      if (Object.keys(columnTotalMap).includes('quantity') && columnTotalMap.quantity.formula) {
        this.sheet.setFormula(subRow, columnTotalMap.quantity.column, columnTotalMap.quantity.formula);
        // this.sheet.autoFitColumn(columnTotalMap.quantity.column);
      } else {
        console.warn('truckage模板中缺少数量字段，无法计算搬运费所在列索引!');
      }
      if (Object.keys(columnTotalMap).includes('total1')) {
        if (Object.keys(columnTotalMap).includes('quantity') && columnTotalMap.total1.formula) {
          // setCellFormatter(this.sheet, subRow + 1, columnTotalMap.quantity.column);
          this.sheet.setFormula(subRow + 1, columnTotalMap.quantity.column, columnTotalMap.total1.formula);
          // this.sheet.autoFitColumn(columnTotalMap.quantity.column);
        }
      }
      if (Object.keys(columnTotalMap).includes('total')) {
        if (columnTotalMap.total.formula) {
          // setCellFormatter(this.sheet, subRow, columnTotalMap.total.column);
          this.sheet.setFormula(subRow, columnTotalMap.total.column, columnTotalMap.total.formula);
          // this.sheet.autoFitColumn(columnTotalMap.total.column);
        }

        let porterageCosts = '';
        let freight = '';
        if (Object.keys(columnTotalMap).includes('total1')) {
          porterageCosts = columnTotalMap.total1.formula;
        }
        if (_.has(quotation, ['freight'])) {
          if (_.get(quotation, ['freight'])) {
            freight = Number(_.get(quotation, ['freight']));
          }
        }

        let total = columnTotalMap.total.formula;
        if (freight) {
          total = `(${total}) + ${freight}`;
        }
        if (porterageCosts) {
          total = `(${total}) + ${porterageCosts}`;
        }
        if (total) {
          // setCellFormatter(this.sheet, subRow + 1, columnTotalMap.total.column);
          this.sheet.setFormula(subRow + 1, columnTotalMap.total.column, total);
          // this.sheet.autoFitColumn(columnTotalMap.total.column);
        }
      }
      this.sheet.resumePaint();
    }
  }

  /**
   * 内置的模板标识符(id)，目的是为了解决兼容旧的模板
   */
  builtInIdsIdentifier() {
    const templateId = store.getters['quotationModule/GetterTemplateId'];
    if (this.builtInIdMap[templateId]) {
      return this.builtInIdMap[templateId];
    }
  }
}

/**
 * Define a template classification identifier
 */
export const DEFINE_IDENTIFIER_MAP = {
  'noLevel': {
    label: '无分类',
    identifier: 'noLevel',
  },
  'Level_1_row': {
    label: '一级分类 + 按 行 合并 + top层配置表头',
    identifier: 'Level_1_row',
  },
  'Level_1_col': {
    label: '一级分类 + 按 列 合并 + top层配置表头',
    identifier: 'Level_1_col',
  },
  'title@Level_1_row': {
    label: '一级分类 + 按 行 合并 +  center中的表头',
    identifier: 'title@Level_1_row',
  },
  'title@Level_1_col': {
    label: '一级分类 + 按 列 合并 +  center中的表头',
    identifier: 'title@Level_1_col',
  },
  'Level_1_row@Level_2_row': {
    label: '',
    identifier: 'Level_1_row@Level_2_row'
  },
  'Level_1_row@Level_2_col': {
    label: '',
    identifier: 'Level_1_row@Level_2_col'
  },
  'Level_1_col@Level_2_row': {
    label: '',
    identifier: 'Level_1_col@Level_2_row'
  },
  'Level_1_col@Level_2_col': {
    label: '',
    identifier: 'Level_1_col@Level_2_col'
  },
  'title@Level_1_row@Level_2_row': {
    label: '',
    identifier: 'title@Level_1_row@Level_2_row'
  },
  'title@Level_1_row@Level_2_col': {
    label: '',
    identifier: 'title@Level_1_row@Level_2_col'
  },
  'title@Level_1_col@Level_2_row': {
    label: '',
    identifier: 'title@Level_1_col@Level_2_row'
  },
  'title@Level_1_col@Level_2_col': {
    label: '',
    identifier: 'title@Level_1_col@Level_2_col'
  }
}


export default IdentifierTemplate;

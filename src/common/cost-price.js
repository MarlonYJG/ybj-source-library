/*
 * @Author: Marlon
 * @Date: 2024-07-25 17:34:49
 * @Description: check cost price
 */
import * as GC from '@grapecity/spread-sheets';
import { GeneratorTableStyle } from './generator';

import { isMultiHead } from './parsing-template'

const TITLE = [
  { name: '成本单价', bindPath: 'costPrice' },
  { name: '成本总价', bindPath: 'costTotalPrice' },
  { name: '毛利润', bindPath: 'grossProfit' },
  { name: '毛利率', bindPath: 'grossMargin' }
]

export class CheckCostPrice {
  constructor(spread, template, quotation) {
    this.spread = spread;
    this.sheet = null;
    this.colCount = 4;
    this.template = template;
    this.quotation = quotation;
    this.HeadStyle = null;
    this.sourceSheetColCount = 0;
    this.headStartIndex = [];
    this.columnIndex = [];
    this.init();
  }
  init() {
    const sheet = this.spread.getActiveSheet();
    this.sheet = sheet;
    this.sourceSheetColCount = this.template.sheets[sheet.name()].columnCount;
  }

  /**
   * Delete the added columns
   */
  deleteCol() {
    this.sheet.suspendPaint();
    const colCount = this.sheet.getColumnCount();
    if (this.sourceSheetColCount < colCount) {
      this.sheet.deleteColumns(colCount - this.colCount, this.colCount);
    }
    this.sheet.resumePaint()
  }
  /**
   * Initialize the rendering
   */
  render() {
    const startColIndex = this.sheet.getColumnCount();
    this.sheet.suspendPaint()
    this.sheet.addColumns(startColIndex, this.colCount);
    this.sheet.resumePaint()
    this.getColIndex();
    this.getHeaderStyle();
  }

  // 获取列索引
  getColIndex() {
    for (let c = this.sourceSheetColCount; c < this.colCount + this.sourceSheetColCount; c++) {
      this.columnIndex.push(c);
    }
  }

  // 获取表头样式
  getHeaderStyle() {
    if (!isMultiHead()) {
      const row = this.template.cloudSheet.top.rowCount - 1;
      const col = 1;
      const hstyle = this.sheet.getStyle(row, col);
      this.HeadStyle = hstyle;
      this.headStartIndex = [row]
    } else {
      // TODO 多表头
      // this.HeadStyle = hstyle;
      // this.headStartIndex = [row]
    }
  }

  /**--------------------- */

  // 绘制表头
  drawTitles() {
    this.sheet.suspendPaint()
    this.setHeaderStyle();
    if (this.headStartIndex.length) {
      for (let r = 0; r < this.headStartIndex.length; r++) {
        for (let c = 0; c < this.columnIndex.length; c++) {
          this.sheet.setValue(this.headStartIndex[r], this.columnIndex[c], TITLE[c].name);
        }
      }
    }
    this.sheet.resumePaint()
  }

  // 设置表头样式
  setHeaderStyle() {
    for (let r = 0; r < this.headStartIndex.length; r++) {
      for (let c = 0; c < this.columnIndex.length; c++) {
        this.sheet.setStyle(this.headStartIndex[r], this.columnIndex[c], this.HeadStyle);
      }
    }
  }

  // 绘制表格
  drawTables() {
    const tableColumns = [];
    TITLE.forEach(h => {
      const col = new GC.Spread.Sheets.Tables.TableColumn();
      col.name(h.name);
      col.dataField(h.bindPath)
      tableColumns.push(col);
    });
    // const tableCostMap = {
    //   '1': {
    //     r: 10,
    //     rc: 1
    //   }
    // }
    // for (const key in tableCostMap) {
    //   if (Object.hasOwnProperty.call(tableCostMap, key)) {
    //     const table = this.sheet.tables.add(`tableCost_${key}`, tableCostMap[key].r, this.sourceSheetColCount, tableCostMap[key].rc, this.colCount, GeneratorTableStyle());
    //     table.bindColumns(tableColumns);
    //     table.expandBoundRows(true);
    //     table.autoGenerateColumns(false);
    //     table.showFooter(true);
    //     table.showHeader(true);
    //     table.highlightFirstColumn(false);
    //     table.highlightLastColumn(false);
    //     for (let i = 0; i < this.colCount; i++) {
    //       table.filterButtonVisible(i, false)
    //     }



    //     // 绑定数据
    //     // table.bindingPath('costPrice');

    //   }
    // }
  }
  // 绘制小计
  drawSubTotal() {


  }
  // 绘制总计
  drawTotal() {


  }
}

export default CheckCostPrice;
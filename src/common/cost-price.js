/*
 * @Author: Marlon
 * @Date: 2024-07-25 17:34:49
 * @Description: check cost price
 */


export class CheckCostPrice {
  constructor(spread, template) {
    this.spread = spread;
    this.colCount = 4;
    this.template = template;
    this.init();
  }
  init() {
    const sheet = this.spread.getActiveSheet();
    this.sheet = sheet;
    this.sourceSheetColCount = this.template.sheets[sheet.name()].columnCount;
  }

  // 删除列
  deleteCol() {
    const colCount = this.sheet.getColumnCount();
    if (this.sourceSheetColCount < colCount) {
      this.sheet.deleteColumns(colCount - this.colCount, this.colCount);
    }
  }

  render() {
    const startColIndex = this.sheet.getColumnCount();
    this.sheet.addColumns(startColIndex, this.colCount);
  }
  // 绘制表头
  drawHeads(sheet) {


  }
  // 绘制表体
  drawTables(sheet) {


  }
  // 绘制小计
  drawSubTotal(sheet) {


  }
  // 绘制总计
  drawTotal(sheet) {


  }
}

export default CheckCostPrice;
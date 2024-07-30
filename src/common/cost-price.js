/*
 * @Author: Marlon
 * @Date: 2024-07-25 17:34:49
 * @Description: check cost price
 */
import * as GC from '@grapecity/spread-sheets';
import { GeneratorCostTableStyle, GeneratorCellStyle, GeneratorLineBorder } from './generator';

import { isMultiHead } from './parsing-template'
import { LayoutRowColBlock } from './core';
import { numberToColumn } from './public'
import { SetDataSource } from './sheetWorkBook'

const TITLE = [
  { name: '成本单价', bindPath: 'costPrice' },
  { name: '成本总价', bindPath: 'costTotalPrice' },
  { name: '毛利润', bindPath: 'grossProfit' },
  { name: '毛利率', bindPath: 'grossMargin' }
]

const KEYS = ['numberOfDays', 'quantity', 'total']

export class CheckCostPrice {
  static CellEdit = [];
  constructor(spread, template, quotation) {
    this.spread = spread;
    this.sheet = null;
    this.colCount = 4;
    this.template = template;
    this.quotation = quotation;
    this.HeadStyle = null;
    this.sourceSheetColCount = 0;
    this.headStartIndex = [];
    this.ColumnIndex = [];
    this.TableSubMap = null;
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
    this.check();
    const startColIndex = this.sheet.getColumnCount();
    this.sheet.suspendPaint()
    this.sheet.addColumns(startColIndex, this.colCount);
    for (let c = startColIndex; c < startColIndex + this.colCount; c++) {
      this.sheet.setColumnWidth(c, 100);
    }
    this.sheet.resumePaint()
    this.getColIndex();
    this.getHeaderStyle();
  }

  /**
   * Get the column index
   */
  getColIndex() {
    for (let c = this.sourceSheetColCount; c < this.colCount + this.sourceSheetColCount; c++) {
      this.ColumnIndex.push(c);
    }
  }

  /**
   * Get the header style
   */
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

  /**
   * Draw the header
   */
  drawTitles() {
    this.sheet.suspendPaint()
    this.setHeaderStyle();
    if (this.headStartIndex.length) {
      for (let r = 0; r < this.headStartIndex.length; r++) {
        for (let c = 0; c < this.ColumnIndex.length; c++) {
          this.sheet.setValue(this.headStartIndex[r], this.ColumnIndex[c], TITLE[c].name);
        }
      }
    }
    this.sheet.resumePaint()
  }

  /**
   * Style the headers
   */
  setHeaderStyle() {
    for (let r = 0; r < this.headStartIndex.length; r++) {
      for (let c = 0; c < this.ColumnIndex.length; c++) {
        this.sheet.setStyle(this.headStartIndex[r], this.ColumnIndex[c], this.HeadStyle);
      }
    }
  }

  /**
   * Draw a table
   */
  drawTables() {
    SetDataSource(this.sheet, this.quotation)
    const tableColumns = [];
    TITLE.forEach(h => {
      const col = new GC.Spread.Sheets.Tables.TableColumn();
      col.name(h.name);
      col.dataField(h.bindPath)
      tableColumns.push(col);
    });

    const { tables, subTotals, totalMap } = LayoutRowColBlock(this.spread);
    this.sheet.suspendPaint();
    for (let i = 0; i < tables.length; i++) {
      for (const key in tables[i]) {
        if (Object.hasOwnProperty.call(tables[i], key)) {
          const table = this.sheet.tables.add(`tableCost_${key}`, tables[i][key].row - 1, this.sourceSheetColCount, tables[i][key].rowCount + 1, this.colCount, GeneratorCostTableStyle());
          table.bindColumns(tableColumns);
          table.expandBoundRows(true);
          table.autoGenerateColumns(false);
          table.highlightFirstColumn(false);
          table.highlightLastColumn(false);
          table.showHeader(false);
          table.showFooter(false);
          table.bindingPath(`conferenceHall.resourceViewsMap.${key}.resources`);
        }
      }
    }

    this.setStyle(tables, subTotals, totalMap);
    this.enableCellEditable(tables);
    this.setFormula(tables);
    this.drawSubTotal(tables)
    this.setSubTotal(subTotals);
    this.drawTotal(totalMap);

    this.sheet.resumePaint();
  }


  /**
   * Set the calculation formula
   * @param {*} tables 
   */
  setFormula(tables) {
    const title = this.template.cloudSheet.center.equipment.bindPath;
    const colMap = {};
    for (const key in title) {
      if (Object.hasOwnProperty.call(title, key)) {
        if (KEYS.includes(key)) {
          colMap[key] = title[key].columnHeader
        }
      }
    }

    for (let i = 0; i < tables.length; i++) {
      for (const key in tables[i]) {
        if (Object.hasOwnProperty.call(tables[i], key)) {
          for (let r = tables[i][key].row; r < tables[i][key].row + tables[i][key].rowCount; r++) {

            const costTotalPrice = [];
            if (colMap.quantity) {
              costTotalPrice.push(`${colMap.quantity}${r + 1}`)
            }
            if (colMap.numberOfDays) {
              costTotalPrice.push(`${colMap.numberOfDays}${r + 1}`)
            }
            costTotalPrice.push(`${numberToColumn(this.ColumnIndex[0] + 1)}${r + 1}`);

            const costTotalPriceFormula = costTotalPrice.join('*');
            const grossProfitFormula = `${colMap.total}${r + 1} - ${numberToColumn(this.ColumnIndex[1] + 1)}${r + 1}`;
            const grossMarginFormula = `${numberToColumn(this.ColumnIndex[2] + 1)}${r + 1} / ${colMap.total}${r + 1}`;

            this.sheet.setFormula(r, this.ColumnIndex[1], `IFERROR(${costTotalPriceFormula},"")`);
            this.sheet.setFormula(r, this.ColumnIndex[2], `IFERROR(${grossProfitFormula},"")`);
            this.sheet.setFormula(r, this.ColumnIndex[3], `IFERROR(${grossMarginFormula},"")`);
          }
        }
      }
    }
  }

  /**
   * Draw subtotals
   * @param {*} tables 
   */
  drawSubTotal(tables) {
    const tableSubMap = {};
    for (let i = 0; i < tables.length; i++) {
      for (const key in tables[i]) {
        if (Object.hasOwnProperty.call(tables[i], key)) {

          const sub = [];
          const rows = []
          for (let r = tables[i][key].row; r < tables[i][key].row + tables[i][key].rowCount; r++) {
            rows.push(r);
          }

          for (let c = 0; c < this.ColumnIndex.length; c++) {
            const formulaSum = [];
            for (let r = 0; r < rows.length; r++) {
              formulaSum.push(`${numberToColumn(this.ColumnIndex[c] + 1)}${rows[r] + 1}`)
            }
            sub.push(formulaSum.join('+'))
          }

          tableSubMap[key] = sub;
        }
      }
    }
    this.TableSubMap = tableSubMap;
    console.log(tableSubMap, 'tableSubMap');
  }


  /**
   * Set subtotal
   * @param {*} subTotals 
   */
  setSubTotal(subTotals) {
    for (let j = 0; j < subTotals.length; j++) {
      for (const key in subTotals[j]) {
        if (Object.hasOwnProperty.call(subTotals[j], key)) {
          for (let c = 0; c < this.ColumnIndex.length; c++) {
            this.sheet.setFormula(subTotals[j][key].row, this.ColumnIndex[c], `IFERROR(${this.TableSubMap[key][c]},"")`);
          }
        }
      }
    }
  }


  /**
   * Draw total
   * @param {*} totalMap 
   */
  drawTotal(totalMap) {
    if (!totalMap || !this.TableSubMap) return;

    const formulas = [[], [], [], []];

    for (const key in this.TableSubMap) {
      if (Object.hasOwnProperty.call(this.TableSubMap, key)) {
        this.TableSubMap[key].forEach((value, index) => {
          formulas[index].push(value);
        });
      }
    }

    formulas.forEach((formula, index) => {
      const formulaStr = formula.join('+');
      this.sheet.setFormula(totalMap.row, this.ColumnIndex[index], `IFERROR(${formulaStr},"")`);
    });
  }

  /**
   * Turn on the specified cell to edit
   */
  enableCellEditable(tables) {
    const cells = [];
    for (let i = 0; i < tables.length; i++) {
      for (const key in tables[i]) {
        if (Object.hasOwnProperty.call(tables[i], key)) {
          for (let r = tables[i][key].row; r < tables[i][key].row + tables[i][key].rowCount; r++) {
            cells.push({ row: r, col: this.ColumnIndex[0] });
            this.sheet.getRange(r, this.ColumnIndex[0], 1, 1).locked(false);
          }
        }
      }
    }
    CheckCostPrice.CellEdit = cells;
  }

  /**
   * Verify that the necessary conditions for calculating the cost price are met
   * @returns 
   */
  check() {
    const title = this.template.cloudSheet.center.equipment.bindPath;
    const missingKeys = KEYS.filter(key => !Object.prototype.hasOwnProperty.call(title, key));
    if (missingKeys.length === 0) {
      return true;
    } else {
      console.error(`Missing keys: ${missingKeys.join(', ')}`);
      return false;
    }
  }

  /**
   * Style the body
   * @param {*} tables 
   * @param {*} subTotals 
   * @param {*} totalMap 
   */
  setStyle(tables, subTotals, totalMap) {
    const lineBorder = GeneratorLineBorder();
    const { style } = GeneratorCellStyle('costStyle', { hAlign: 1, vAlign: 1, borderBottom: lineBorder, borderLeft: lineBorder, borderRight: lineBorder, borderTop: lineBorder, });
    for (let i = 0; i < tables.length; i++) {
      for (const key in tables[i]) {
        if (Object.hasOwnProperty.call(tables[i], key)) {
          this.sheet.getRange(tables[i][key].row, this.ColumnIndex[0], tables[i][key].rowCount, this.colCount).setStyle(style)
        }
      }
    }

    for (let i = 0; i < subTotals.length; i++) {
      for (const key in subTotals[i]) {
        if (Object.hasOwnProperty.call(subTotals[i], key)) {
          this.sheet.getRange(subTotals[i][key].row, this.ColumnIndex[0], 1, this.colCount).setStyle(style)
        }
      }

    }

    if (totalMap) {
      this.sheet.getRange(totalMap.row, this.ColumnIndex[0], 1, this.colCount).setStyle(style)
    }
  }

  /**
   * Get editable cells
   * @param {*} tables 
   * @returns 
   */
  getCellEditable() {
    return CheckCostPrice.CellEdit;
  }
}

export default CheckCostPrice;
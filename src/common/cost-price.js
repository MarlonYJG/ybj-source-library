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
  static ColumnIndex = [];

  constructor(spread, template, quotation) {
    this.spread = spread;
    this.sheet = null;
    this.colCount = 4;
    this.template = template;
    this.quotation = quotation;
    this.HeadStyle = null;
    this.sourceSheetColCount = 0;
    this.headStartIndex = [];
    this.TableSubMap = null;
    this.init();
  }
  init() {
    const sheet = this.spread.getActiveSheet();
    this.sheet = sheet;
    this.sourceSheetColCount = this.template.sheets[sheet.name()].columnCount;
    this.getColIndex();
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
    this.getHeaderStyle();
  }

  /**
   * Get the column index
   */
  getColIndex() {
    const columnIndex = [];
    for (let c = this.sourceSheetColCount; c < this.colCount + this.sourceSheetColCount; c++) {
      columnIndex.push(c);
    }
    CheckCostPrice.ColumnIndex = columnIndex;
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
    const columnIndex = CheckCostPrice.ColumnIndex;
    this.sheet.suspendPaint()
    this.setHeaderStyle();
    if (this.headStartIndex.length) {
      for (let r = 0; r < this.headStartIndex.length; r++) {
        for (let c = 0; c < columnIndex.length; c++) {
          this.sheet.setValue(this.headStartIndex[r], columnIndex[c], TITLE[c].name);
        }
      }
    }
    this.sheet.resumePaint()
  }

  /**
   * Style the headers
   */
  setHeaderStyle() {
    const columnIndex = CheckCostPrice.ColumnIndex;
    for (let r = 0; r < this.headStartIndex.length; r++) {
      for (let c = 0; c < columnIndex.length; c++) {
        this.sheet.setStyle(this.headStartIndex[r], columnIndex[c], this.HeadStyle);
      }
    }
  }

  /**
   * Draw a table
   */
  drawTables() {
    SetDataSource(this.sheet, this.quotation)

    const layout = new LayoutRowColBlock(this.spread);
    const { Tables, SubTotals, TotalMap, Levels } = layout.getLayout();
    console.log(Tables, SubTotals, TotalMap);

    const classType = layout.getClassType();

    if (classType === 'noLevel') {
      // 
    } else if (classType === 'Level_1_row') {
      this.sheet.suspendPaint();
      this.createTable(Tables, -1, 1)
      this.setStyle(Tables, SubTotals, TotalMap, Levels);
      this.enableCellEditable(Tables);
      this.setFormula(Tables);
      this.drawSubTotal(Tables)
      this.setSubTotal(SubTotals);
      this.drawTotal(TotalMap);

      this.sheet.resumePaint();
    }
  }


  /**
   * Set the calculation formula
   * @param {*} tables 
   */
  setFormula(tables) {
    const columnIndex = CheckCostPrice.ColumnIndex;
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
            costTotalPrice.push(`${numberToColumn(columnIndex[0] + 1)}${r + 1}`);

            const costTotalPriceFormula = costTotalPrice.join('*');
            const grossProfitFormula = `${colMap.total}${r + 1} - ${numberToColumn(columnIndex[1] + 1)}${r + 1}`;
            const grossMarginFormula = `${numberToColumn(columnIndex[2] + 1)}${r + 1} / ${colMap.total}${r + 1}`;

            this.sheet.setFormula(r, columnIndex[1], `IFERROR(${costTotalPriceFormula},"")`);
            this.sheet.setFormula(r, columnIndex[2], `IFERROR(${grossProfitFormula},"")`);
            this.sheet.setFormula(r, columnIndex[3], `IFERROR(${grossMarginFormula},"")`);
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
    const columnIndex = CheckCostPrice.ColumnIndex;
    const tableSubMap = {};
    for (let i = 0; i < tables.length; i++) {
      for (const key in tables[i]) {
        if (Object.hasOwnProperty.call(tables[i], key)) {

          const sub = [];
          const rows = []
          for (let r = tables[i][key].row; r < tables[i][key].row + tables[i][key].rowCount; r++) {
            rows.push(r);
          }

          for (let c = 0; c < columnIndex.length; c++) {
            const formulaSum = [];
            for (let r = 0; r < rows.length; r++) {
              formulaSum.push(`${numberToColumn(columnIndex[c] + 1)}${rows[r] + 1}`)
            }
            sub.push(formulaSum.join('+'))
          }

          tableSubMap[key] = sub;
        }
      }
    }
    this.TableSubMap = tableSubMap;
  }


  /**
   * Set subtotal
   * @param {*} subTotals 
   */
  setSubTotal(subTotals) {
    if (subTotals) {
      const columnIndex = CheckCostPrice.ColumnIndex;
      for (let j = 0; j < subTotals.length; j++) {
        for (const key in subTotals[j]) {
          if (Object.hasOwnProperty.call(subTotals[j], key)) {
            for (let c = 0; c < columnIndex.length; c++) {
              this.sheet.setFormula(subTotals[j][key].row, columnIndex[c], `IFERROR(${this.TableSubMap[key][c]},"")`);
            }
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

    const columnIndex = CheckCostPrice.ColumnIndex;
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
      this.sheet.setFormula(totalMap.row + (totalMap.rowCount - 1), columnIndex[index], `IFERROR(${formulaStr},"")`);
    });
  }

  /**
   * Update the location of the total
   * @param {*} tables 
   * @param {*} totalMap 
   */
  updateTotalPosition(tables, totalMap) {
    this.setTotalStyle(totalMap);
    this.drawSubTotal(tables)
    this.drawTotal(totalMap);
  }

  /**
   * Turn on the specified cell to edit
   */
  enableCellEditable(tables) {
    const cells = [];
    const columnIndex = CheckCostPrice.ColumnIndex;
    for (let i = 0; i < tables.length; i++) {
      for (const key in tables[i]) {
        if (Object.hasOwnProperty.call(tables[i], key)) {
          for (let r = tables[i][key].row; r < tables[i][key].row + tables[i][key].rowCount; r++) {
            cells.push({ row: r, col: columnIndex[0] });
            this.sheet.getRange(r, columnIndex[0], 1, 1).locked(false);
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
   * @param {*} Levels 
   */
  setStyle(tables, subTotals, totalMap, Levels) {
    const columnIndex = CheckCostPrice.ColumnIndex;
    for (let i = 0; i < tables.length; i++) {
      for (const key in tables[i]) {
        if (Object.hasOwnProperty.call(tables[i], key)) {
          this.sheet.getRange(tables[i][key].row, columnIndex[0], tables[i][key].rowCount, this.colCount).setStyle(this.style())
        }
      }
    }

    if (subTotals) {
      for (let i = 0; i < subTotals.length; i++) {
        for (const key in subTotals[i]) {
          if (Object.hasOwnProperty.call(subTotals[i], key)) {
            this.sheet.getRange(subTotals[i][key].row, columnIndex[0], 1, this.colCount).setStyle(this.style())
          }
        }
      }
    }

    if (Levels) {
      if (Levels.length === 1) {
        for (let i = 0; i < Levels[0].length; i++) {
          for (const key in Levels[0][i]) {
            if (Object.hasOwnProperty.call(Levels[0][i], key)) {
              this.sheet.getRange(Levels[0][i][key].row, columnIndex[0], 1, this.colCount).setStyle(this.style())
            }
          }

        }
      } else if (Levels.length === 2) {
        // TODO
      }
    }

    this.setTotalStyle(totalMap);
  }

  /**
   * Set the grand total style
   * @param {*} totalMap 
   */
  setTotalStyle(totalMap) {
    if (totalMap) {
      const columnIndex = CheckCostPrice.ColumnIndex;
      this.sheet.getRange(totalMap.row, columnIndex[0], totalMap.rowCount, this.colCount).setStyle(this.style());
    }
  }

  /**
   * Generate style
   * @returns 
   */
  style() {
    const lineBorder = GeneratorLineBorder();
    const { style } = GeneratorCellStyle('costStyle', { hAlign: 1, vAlign: 1, borderBottom: lineBorder, borderLeft: lineBorder, borderRight: lineBorder, borderTop: lineBorder, });
    return style;
  }

  /**
   * create table and binding data
   * @param {*} tables 
   * @param {*} r 
   * @param {*} rc 
   */
  createTable(tables, r, rc) {
    const tableColumns = [];
    TITLE.forEach(h => {
      const col = new GC.Spread.Sheets.Tables.TableColumn();
      col.name(h.name);
      col.dataField(h.bindPath)
      tableColumns.push(col);
    });

    for (let i = 0; i < tables.length; i++) {
      for (const key in tables[i]) {
        if (Object.hasOwnProperty.call(tables[i], key)) {
          const table = this.sheet.tables.add(`tableCost_${key}`, tables[i][key].row + r, this.sourceSheetColCount, tables[i][key].rowCount + rc, this.colCount, GeneratorCostTableStyle());

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
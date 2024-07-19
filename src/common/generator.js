/*
 * @Author: Marlon
 * @Date: 2024-05-04 11:38:22
 * @Description:Spread generator
 */
import * as GC from '@grapecity/spread-sheets';

/**
 * Style Generator
 * @param {*} name
 * @param {*} cfg
 * @returns
 */
export const GeneratorStyle = (name = 'style', cfg = {}) => {
  const style = new GC.Spread.Sheets.Style();
  const defaultCfg = {
    hAlign: 'left',
    vAlign: 1,
    fontWeight: 'bolder',
    fontSize: '12px',
    textIndent: 2,
    wordWrap: false,
    name: name
  };

  const config = Object.assign({}, defaultCfg, cfg);

  for (const key in config) {
    if (Object.hasOwnProperty.call(config, key)) {
      style[key] = config[key];
    }
  }

  return {
    style,
    name
  };
};

/**
 * Border Generator
 * @param {*} cfg
 * @returns
 */
export const GeneratorLineBorder = (cfg = {}) => {
  const border = new GC.Spread.Sheets.LineBorder();
  const defaultCfg = {
    style: GC.Spread.Sheets.LineStyle.thin,
    color: '#000000'
  };
  const config = Object.assign({}, defaultCfg, cfg);
  for (const key in config) {
    if (Object.hasOwnProperty.call(config, key)) {
      border[key] = config[key];
    }
  }
  return border;
};

/**
 * Table style Generator
 * @returns
 */
export const GeneratorTableStyle = () => {
  const tStyleInfo = new GC.Spread.Sheets.Tables.TableStyle();
  tStyleInfo.backColor = 'white';
  const tableStyle = new GC.Spread.Sheets.Tables.TableTheme();
  tableStyle.headerRowStyle(tStyleInfo);
  return tableStyle;
};

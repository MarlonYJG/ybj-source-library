/*
 * @Author: Marlon
 * @Date: 2024-08-01 13:11:10
 * @Description: template
 */
import { sheet as common } from './public'
import { sheet as noLevel } from "./noLevel/index";
import { sheet as Level_1_row } from './Level_1_row/index'

const defValue = {
  id: 'public',
  excelJson: common
}

export default {
  default: defValue,
  noLevel: {
    id: 'noLevel',
    excelJson: noLevel
  },
  Level_1_row: {
    id: 'Level_1_row',
    excelJson: Level_1_row
  }
}
/*
 * @Author: Marlon
 * @Date: 2024-06-28 10:33:22
 * @Description: 
 */

declare namespace SourceLibraryYbj {
  /**
   * 生成数字范围内的随机数
   * @param min 最小数字
   * @param max 最大数字
   * @returns number类型
   */
  export function random(min: number, max: number): number
}

declare module 'source-library-ybj' {
  export = SourceLibraryYbj
}

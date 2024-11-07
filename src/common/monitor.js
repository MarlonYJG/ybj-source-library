/*
 * @Author: Marlon
 * @Date: 2024-11-07 09:53:35
 * @Description: Monitor class, which allows multiple instance functions to be stored and called when needed
 */

class Monitor {
  constructor() {
    this.functions = {
      cellDialog: () => {
        console.log('Default cellDialog function');
      }
    };
  }

  /**
   * Add a function
   * @param {*} name 
   * @param {*} func 
   */
  addFunction(name, func) {
    if (typeof func === 'function') {
      this.functions[name] = func;
    } else {
      throw new Error('Only functions can be added.');
    }
  }

  /**
   * Call a function
   * @param {*} name 
   * @param  {...any} args 
   * @returns 
   */
  callFunction(name, ...args) {
    if (this.functions[name]) {
      return this.functions[name](...args);
    } else {
      throw new Error(`Function "${name}" not found.`);
    }
  }

  /**
   * Call all functions
   * @param  {...any} args 
   */
  callAll(...args) {
    Object.keys(this.functions).forEach(name => {
      this.functions[name](...args);
    });
  }
}

export const monitorInstance = new Monitor();

export default Monitor;

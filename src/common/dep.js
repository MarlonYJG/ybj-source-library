/*
 * @Author: Marlon
 * @Date: 2024-09-14 10:00:16
 * @Description: Define subscriber classes and instances
 */

export class Dep {
  constructor() {
    this.subscribers = new Set();
  }

  /**
   * Add subscribers
   * @param {*} subscriber 
   */
  addSub(subscriber) {
    this.subscribers.add(subscriber);
  }

  /**
   * Notify all subscribers
   */
  notify(value) {
    this.subscribers.forEach((subscriber) => subscriber(value));
  }
}

// Limit discount input
export const depLimitDiscountInput = new Dep();
export const addSubLimitDiscountInput = function (subscriber) {
  depLimitDiscountInput.addSub(subscriber);
};

// Limit discount input type
export const depLimitDiscountInputType = new Dep();
export const addSubLimitDiscountInputType = function (subscriber) {
  depLimitDiscountInputType.addSub(subscriber);
};

// Sort object
export const depSortObject = new Dep();
export const addSubSortObject = function (subscriber) {
  depSortObject.addSub(subscriber);
};

// Export error
export const depExportError = new Dep();
export const addSubExportError = function (subscriber) {
  depExportError.addSub(subscriber);
};

// Select equipment image
export const depSelectEquipmentImage = new Dep();
export const addSubSelectEquipmentImage = function (subscriber) {
  depSelectEquipmentImage.addSub(subscriber);
};




export default Dep;

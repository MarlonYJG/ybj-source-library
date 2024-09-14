/*
 * @Author: Marlon
 * @Date: 2024-09-14 10:04:18
 * @Description: Define the proxy object
 */
import { depLimitDiscountInput, depLimitDiscountInputType, depSortObject, depExportError } from './dep';

// The object to be proxied
const limitDiscountInput = {
  value: null,
};
const limitDiscountInputType = {
  value: null,
};
const sortObject = {
  value: null
};
const exportError = {
  value: null
}


// Use Proxy objects
export const limitDiscountInputProxy = new Proxy(limitDiscountInput, {
  set(target, key, value) {
    target[key] = value;
    depLimitDiscountInput.notify(value);
    return true;
  },
});
export const limitDiscountInputTypeProxy = new Proxy(limitDiscountInputType, {
  set(target, key, value) {
    target[key] = value;
    depLimitDiscountInputType.notify(value);
    return true;
  },
});
export const sortObjectProxy = new Proxy(sortObject, {
  set(target, key, value) {
    target[key] = value;
    depSortObject.notify(value);
    return true;
  },
});
export const exportErrorProxy = new Proxy(exportError, {
  set(target, key, value) {
    target[key] = value;
    depExportError.notify(value);
    return true;
  },
});


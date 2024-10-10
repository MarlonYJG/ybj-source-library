/*
 * @Author: Marlon
 * @Date: 2024-09-19 11:48:52
 * @Description: 
 */
import _ from '../lib/lodash/lodash.min.js';
import store from 'store';
import {
  UPDATE_QUOTATION_PATH
} from 'store/quotation/mutation-types';

import { SetDataSource } from './sheetWorkBook';
import { selectEquipmentImageProxy } from './proxyData';
import { AddEquipmentImage } from './parsing-template';

let isProcessing = null;// 加锁防止重复执行

/**
 * Delete the device picture
 * @param {*} spread 
 * @param {*} pictureId 
 */
const delEquipmentImage = (spread, pictureId) => {
  const sheet = spread.getActiveSheet();
  const allPictures = sheet.pictures.all()
  allPictures.forEach(item => {
    console.log(item.name());

  });
  const picture = sheet.pictures.get(pictureId);
  if (picture) {
    sheet.resumePaint();
    sheet.pictures.remove(pictureId);
    sheet.repaint();
  }
}

/**
 * Add a picture of your device
 * @param {*} spread 
 * @param {*} pictureId 
 * @param {*} row 
 * @param {*} id 
 */
const insertEquipmentImage = (spread, pictureId, row, id) => {
  const sheet = spread.getActiveSheet();
  selectEquipmentImageProxy.value = (base64, file) => {
    if (base64) {
      isProcessing && clearTimeout(isProcessing);
      isProcessing = setTimeout(() => {
        const picture = sheet.pictures.get(pictureId);
        if (picture) {
          delEquipmentImage(spread, pictureId);
        }
        AddEquipmentImage(spread, pictureId, base64, row);
        uploadEquipmentImageByPath(sheet, id, file)
        isProcessing = null;
      }, 0);
    }
  };
}

/**
 * Upload the images in the quotation to the latest image resources
 * @param {*} sheet 
 * @param {*} id 
 * @param {*} file 
 */
const uploadEquipmentImageByPath = (sheet, id, file) => {
  const quotation = store.getters['quotationModule/GetterQuotationInit'];
  const conferenceHall = _.cloneDeep(quotation.conferenceHall);
  const resourceViews = conferenceHall.resourceViews;
  const resourceViewsMap = {};
  resourceViews.forEach(item => {
    item.resources.forEach(resource => {
      if (resource.id === id) {
        resource.imageFile = file
      }
    });
  });
  resourceViews.forEach(item => {
    resourceViewsMap[item.resourceLibraryId] = item;
  });
  conferenceHall.resourceViews = resourceViews;
  conferenceHall.resourceViewsMap = resourceViewsMap;

  store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
    path: ['conferenceHall'],
    value: conferenceHall
  });
  SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);
}

/**
 * Update the equipment image in the quote
 * @param {*} spread 
 * @param {*} tableId 
 * @param {*} data 
 * @param {*} imageName 
 * @param {*} type 
 * @param {*} row 
 * @param {*} quotation 
 */
export const updateEquipmentImage = (spread, tableId, data, imageName, type, row, quotation) => {
  if (!quotation) {
    quotation = store.getters['quotationModule/GetterQuotationInit'];
  }

  const sheet = spread.getActiveSheet();
  const conferenceHall = quotation.conferenceHall;
  const resourceViews = conferenceHall.resourceViews;
  const resourceViewsMap = {};
  const findMap = {
    i: null,
    j: null,
    id: null,
    imageId: null
  }

  if (resourceViews && resourceViews.length) {
    for (let i = 0; i < resourceViews.length; i++) {
      if (resourceViews[i].resourceLibraryId === tableId) {
        const resources = resourceViews[i].resources;
        for (let j = 0; j < resources.length; j++) {
          if (resources[j].id === data.id) {
            findMap.imageId = resources[j].imageId;
            findMap.i = i;
            findMap.j = j;
            findMap.id = resources[j].id;
          }
        }
      }
    }
  }

  console.log('待替换的产品id', findMap);

  if (!findMap.id) {
    console.error('找不到对应资源id');
    return;
  }
  if (!findMap.imageId) {
    console.error('找不到对应图片名称字段(imageId)');
    return;
  }
  const resources = resourceViews[findMap.i].resources;
  if (type === 'del') {
    resources[findMap.j].images = null;
    delEquipmentImage(spread, resources[findMap.j].imageId)
  } else if (type === 'insert') {
    insertEquipmentImage(spread, resources[findMap.j].imageId, row, resources[findMap.j].id);
  }

  resourceViews.forEach(item => {
    resourceViewsMap[item.resourceLibraryId] = item;
  });
  conferenceHall.resourceViewsMap = resourceViewsMap;

  store.commit(`quotationModule/${UPDATE_QUOTATION_PATH}`, {
    path: ['conferenceHall'],
    value: conferenceHall
  });
  SetDataSource(sheet, store.getters['quotationModule/GetterQuotationInit']);
}

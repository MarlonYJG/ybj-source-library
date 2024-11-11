/*
 * @Author: Marlon
 * @Date: 2024-05-16 14:08:41
 * @Description:
 */

/**
 * Update sort of conferenceHall
 * @param {*} conferenceHall
 * @returns
 */
export const UpdateSort = (conferenceHall) => {
  const resourceViews = conferenceHall.resourceViews;
  const resourceViewsMap = {};
  if (resourceViews.length) {
    resourceViews.forEach((el1, i) => {
      el1.sort = i + 1;
      if (el1.resources) {
        el1.resources.forEach((el2, index) => {
          el2.sort = index + 1;
        });
      }
      resourceViewsMap[el1.resourceLibraryId] = el1;
    });
  }
  conferenceHall.resourceViewsMap = resourceViewsMap;
  return conferenceHall;
};

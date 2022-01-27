var exceljs;
if (typeof document !== "undefined") {
  exceljs = require("exceljs/dist/es5/exceljs.browser");
} else if (
  typeof navigator != "undefined" &&
  navigator.product == "ReactNative"
) {
  exceljs = require("exceljs/dist/es5/exceljs.bare");
} else {
  exceljs = require("exceljs/dist/es5/exceljs.nodejs");
}
module.exports = exceljs;

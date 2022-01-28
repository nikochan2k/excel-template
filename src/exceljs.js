var exceljs;
if (typeof document !== "undefined") {
  exceljs = require("exceljs/dist/exceljs");
} else if (
  typeof navigator !== "undefined" &&
  navigator.product === "ReactNative"
) {
  exceljs = require("exceljs/dist/exceljs.bare");
} else {
  exceljs = require("exceljs");
}
module.exports = exceljs;

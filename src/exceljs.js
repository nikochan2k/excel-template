var exceljs;
if (typeof document !== "undefined") {
  exceljs = require("exceljs/lib/exceljs.browser.js");
} else if (
  typeof navigator != "undefined" &&
  navigator.product == "ReactNative"
) {
  exceljs = require("exceljs/lib/exceljs.bare.js");
} else {
  exceljs = require("exceljs/lib/exceljs.nodejs.js");
}
module.exports = exceljs;

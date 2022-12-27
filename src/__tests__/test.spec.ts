import type { Workbook } from "exceljs";
import { writeFileSync } from "fs";
import sizeOf from "image-size";
import { join } from "path";
import { pathToFileURL } from "url";
import { deserialize, fit, serialize } from "../ExcelTemplator";
import { ExcelTemplator } from "../index";
import { NodeFetcher } from "../NodeFetcher";
const exceljs = require("../exceljs");

test("test1", async () => {
  const path = join(__dirname, "test1.xlsx");
  const url = pathToFileURL(path);
  const template = new ExcelTemplator(url.href, new NodeFetcher());
  const ab = await template.generate({
    data1: "fuga",
    data2: "hoge",
    data3: "foo",
    data4: "bar",
  });
  const buffer = Buffer.from(ab);
  const outpath = "tmp/test1_out.xlsx";
  writeFileSync(outpath, buffer);
});

test("test2", async () => {
  const path = join(__dirname, "test1.xlsx");
  const url = pathToFileURL(path);
  const template = new ExcelTemplator(url.href, new NodeFetcher());
  const imgPath = join(__dirname, "test.jpg");
  const imgUrl = pathToFileURL(imgPath);
  imgUrl.hash = "#embed";

  const sheetMap = await template.parse();
  for (const targetMap of Object.values(sheetMap)) {
    for (const target of Object.values(targetMap)) {
      if (target.val === "data4") {
        const size = sizeOf(imgPath);
        fit(target, size.width!, size.height!, false);
      }
    }
  }

  const ab = await template.generate(
    {
      data1: "fuga",
      data2: "hoge",
      data3: "foo",
      data4: imgUrl.href,
    },
    sheetMap
  );
  const buffer = Buffer.from(ab);
  const outpath = "tmp/test2_out.xlsx";
  writeFileSync(outpath, buffer);
});

test("test2.1", async () => {
  const path = join(__dirname, "test1.xlsx");
  let workbook: Workbook = new exceljs.Workbook();
  await workbook.xlsx.readFile(path);
  const json = serialize(workbook);
  workbook = deserialize(json);
  const template = new ExcelTemplator(workbook, new NodeFetcher());
  const imgPath = join(__dirname, "test.jpg");
  const imgUrl = pathToFileURL(imgPath);
  imgUrl.hash = "#embed";

  const sheetMap = await template.parse();
  for (const targetMap of Object.values(sheetMap)) {
    for (const target of Object.values(targetMap)) {
      if (target.val === "data4") {
        const size = sizeOf(imgPath);
        fit(target, size.width!, size.height!, false);
      }
    }
  }

  const ab = await template.generate(
    {
      data1: "fuga",
      data2: "hoge",
      data3: "foo",
      data4: imgUrl.href,
    },
    sheetMap
  );
  const buffer = Buffer.from(ab);
  const outpath = "tmp/test2_1_out.xlsx";
  writeFileSync(outpath, buffer);
});

test("test3", async () => {
  const path = join(__dirname, "test1.xlsx");
  const url = pathToFileURL(path);
  const template = new ExcelTemplator(url.href, new NodeFetcher());
  const imgPath = join(__dirname, "test.jpg");
  const imgUrl = pathToFileURL(imgPath);
  const ab = await template.generate({
    data1: "fuga",
    data2: "hoge",
    data3: "foo",
    data4: imgUrl.href,
  });
  const buffer = Buffer.from(ab);
  const outpath = "tmp/test3_out.xlsx";
  writeFileSync(outpath, buffer);
});

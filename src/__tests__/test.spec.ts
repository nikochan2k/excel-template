import { writeFileSync } from "fs";
import sizeOf from "image-size";
import { join } from "path";
import { pathToFileURL } from "url";
import { Workbook, WorkbookModel } from "../exceljs";
import { fit } from "../ExcelTemplator";
import { ExcelTemplator } from "../index";
import { NodeFetcher } from "../NodeFetcher";

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
        target.ext = fit(target, size.width!, size.height!);
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
  const url = pathToFileURL(path);
  let workbook = new Workbook();
  await workbook.xlsx.readFile(path);
  const json = JSON.stringify(workbook.model);
  const model: WorkbookModel = JSON.parse(json, (_, value) =>
    /\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/.test(value)
      ? new Date(value)
      : value
  );
  workbook = new Workbook();
  workbook.model = model;
  const template = new ExcelTemplator(url.href, new NodeFetcher());
  const imgPath = join(__dirname, "test.jpg");
  const imgUrl = pathToFileURL(imgPath);
  imgUrl.hash = "#embed";

  const sheetMap = await template.parse();
  for (const targetMap of Object.values(sheetMap)) {
    for (const target of Object.values(targetMap)) {
      if (target.val === "data4") {
        const size = sizeOf(imgPath);
        target.ext = fit(target, size.width!, size.height!);
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

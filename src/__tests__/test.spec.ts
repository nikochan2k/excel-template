import { writeFileSync } from "fs";
import sizeOf from "image-size";
import { join } from "path";
import { pathToFileURL } from "url";
import { fit } from "../ExcelTemplator";
import { ExcelTemplator } from "../index";

test("test1", async () => {
  const path = join(__dirname, "test1.xlsx");
  const url = pathToFileURL(path);
  const template = new ExcelTemplator(url.href);
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
  const template = new ExcelTemplator(url.href);
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
  const template = new ExcelTemplator(url.href);
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

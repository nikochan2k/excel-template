import { ExcelTemplate } from "../ExcelTemplate";
import { pathToFileURL } from "url";
import { join } from "path";

test("readFile", async () => {
  const path = join(__dirname, "test1.xlsx");
  const url = pathToFileURL(path);
  const template = new ExcelTemplate(url.href);
  await template.parse({
    data1: "fuga",
    data2: "hoge",
    data3: "foo",
    data4: "bar",
  });
});

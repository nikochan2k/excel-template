module.exports = {
  mode: "development",
  devtool: "source-map",
  entry: ["./src/index-browser.ts"],
  output: {
    filename: "excel-templator.js",
    path: __dirname + "/dist",
    libraryTarget: "umd",
    globalObject: "this",
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        use: "ts-loader",
      },
    ],
  },
  resolve: {
    extensions: [".ts", ".js"],
  },
};

module.exports = {
  mode: "production",
  entry: {
    index: "./src/index-browser.ts",
  },
  output: {
    filename: "excel-template.js",
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

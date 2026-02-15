const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = (env, argv) => {
  const isProduction = argv.mode === "production";

  return {
    entry: "./src/taskpane/taskpane.js",
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "taskpane.js",
      clean: true,
    },
    module: {
      rules: [
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/taskpane/taskpane.html",
        filename: "taskpane.html",
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "src/assets", to: "assets" },
        ],
      }),
    ],
    devServer: {
      port: 3000,
      server: "https",
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    },
    devtool: isProduction ? false : "source-map",
  };
};

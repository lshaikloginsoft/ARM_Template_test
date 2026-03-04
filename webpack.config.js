/* eslint-disable no-undef */

const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const nodeExternals = require("webpack-node-externals");
const path = require("path");

const urlDev = "localhost:3000/";
const urlProd = process.env.APP_DOMAIN; 

module.exports = (env, options) => {
  const dev = options.mode === "development";
  const config = [
    {
      devtool: dev ? "source-map" : false,
      entry: {
        polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
        taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
        fallbackauthdialog: "./src/helpers/fallbackauthdialog.js",
      },
      output: {
        filename: "[name].js",
        path: path.resolve(__dirname, "dist"),
        publicPath: "/"
      },
      resolve: {
        extensions: [".html", ".js"],
        fallback: {
          buffer: require.resolve("buffer/")
        },
      },
      module: {
        rules: [
          {
            test: /\.js$/,
            exclude: /node_modules/,
            use: {
              loader: "babel-loader",
            },
          },
          {
            test: /\.html$/,
            exclude: /node_modules/,
            use: "html-loader",
          },
          {
            test: /\.(png|jpg|jpeg|gif|ico)$/,
            type: "asset/resource",
            generator: {
              filename: "assets/[name][ext]",
            },
          },
        ],
      },
      plugins: [
        new HtmlWebpackPlugin({
          filename: "taskpane.html",
          template: "./src/taskpane/taskpane.html",
          chunks: ["polyfill", "taskpane"],
        }),
        new HtmlWebpackPlugin({
          filename: "fallbackauthdialog.html",
          template: "./src/helpers/fallbackauthdialog.html",
          chunks: ["polyfill", "fallbackauthdialog"],
        }),
        new CopyWebpackPlugin({
          patterns: [
            {
              from: "assets/*",
              to: "assets/[name][ext][query]",
            },
            {
              from: "package.json",
              to: "package.json",
            },
            {
              from: "manifest*.xml",
              to: "[name]" + "[ext]",
              transform(content) {
                if (dev) {
                  return content;
                } else {
                  return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
                }
              },
            },
          ],
        }),
      ],
      optimization: {
        minimize: !dev
      }
    },
    {
      devtool: dev ? "source-map" : false,
      target: "node",

      entry: {
        middletier: "./src/middle-tier/app.js",
      },

      output: {
        filename: "[name].js",
        path: path.resolve(__dirname, "dist")
      },

      externals: [nodeExternals()],

      resolve: {
        extensions: [".js"],
      },

      module: {
        rules: [
          {
            test: /\.js$/,
            exclude: /node_modules/,
            use: {
              loader: "babel-loader",
            },
          },
        ],
      }
    },
  ];

  return config;
};

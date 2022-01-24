import Dotenv from 'dotenv-webpack'
import { getHttpsServerOptions } from "office-addin-dev-certs";
import { CleanWebpackPlugin } from "clean-webpack-plugin";
import CopyWebpackPlugin from "copy-webpack-plugin";
import HtmlWebpackPlugin from "html-webpack-plugin";
import HtmlWebpackInlineSVGPlugin from "html-webpack-inline-svg-plugin";
// eslint-disable-next-line no-unused-vars
import fs from "fs";
import ProvidePlugin from "webpack/lib/ProvidePlugin.js";

const urlDev = "https://localhost:3000/";
const urlProd = "https://2yy7ugzn26vanfsfnu8njaml.z22.web.core.windows.net/";
export default async (env, options) => {
  const dev = options.mode === "development";
  const buildType = dev ? "dev" : "prod";
  const config = {
    devtool: "source-map",
    entry: {
      // polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.tsx",
      commands: "./src/commands/commands.ts"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    module: {
      rules: [
        /*         {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env", "@babel/preset-react"]
            }
          }
        }, */
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: "ts-loader"
        },
        /*         {
          test: /\.(jsx?|tsx?)$/,
          include: require("@fluentui/webpack-utilities/lib/fabricAsyncLoaderInclude"),
          loader: "@fluentui/webpack-utilities/lib/fabricAsyncLoader.js"
        }, */
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          loader: "file-loader",
          options: {
            name: "[path][name].[ext]"
          }
        },
        {
          test: /\.svg$/i,
          use: [
            {
              loader: "url-loader",
              options: {
                esModule: false
              }
            }
          ]
        }
      ]
    },
    plugins: [
      new Dotenv(),
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "vendor"],
        minify: true
      }),
      new HtmlWebpackInlineSVGPlugin(),
      new CopyWebpackPlugin({
        patterns: [
          {
            to: "taskpane.css",
            from: "./src/taskpane/taskpane.css"
          },
          {
            to: "[name]." + buildType + ".[ext]",
            from: "manifest*.xml",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            }
          }
        ]
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
        minify: true
      }),
      new ProvidePlugin({
        React: "react"
      })
    ],
    optimization: {
      splitChunks: {
        chunks: "all"
      }
    },
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      https: options.https !== undefined ? options.https : await getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000
    }
  };

  return config;
};

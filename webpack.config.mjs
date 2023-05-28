import {CleanWebpackPlugin} from "clean-webpack-plugin";
import CopyWebpackPlugin from "copy-webpack-plugin";
import HtmlWebpackInlineSVGPlugin from "html-webpack-inline-svg-plugin";
import HtmlWebpackPlugin from "html-webpack-plugin";
import {getHttpsServerOptions} from "office-addin-dev-certs";
import process from "process";
import ProvidePlugin from "webpack/lib/ProvidePlugin.js";
import * as url from 'url';
const __dirname = url.fileURLToPath(new URL('.', import.meta.url));
import path from "path";

const urlDev = "https://localhost:3000/";
const urlProd = "https://2yy7ugzn26vanfsfnu8njaml.z22.web.core.windows.net/";
export default async (env, options) => {
  const dev = options.mode === "development";
  const buildType = dev ? "dev" : "prod";
  const config = {
    output: {
      path: path.resolve(__dirname, "dist"),
    },
    devtool: "source-map",
    entry: {
      taskpane: "./src/taskpane/taskpane.tsx",
      // commands: "./src/commands/commands.ts"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: [/node_modules/, /otherSrc/],
          use: "ts-loader"
        },
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
      /*       new HtmlWebpackPlugin({
              filename: "commands.html",
              template: "./src/commands/commands.html",
              chunks: ["polyfill", "commands"],
              minify: true
            }), */
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
      // https: options.https !== undefined ? options.https : await getHttpsServerOptions(),
      server: {
        type: "https",
        options: {
          ...await getHttpsServerOptions()
        }

      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      host: process.env.npm_package_config_dev_server_host || "127.0.0.1"
    }
  }

  return config;
}

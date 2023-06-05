import {CleanWebpackPlugin} from "clean-webpack-plugin";
import CopyWebpackPlugin from "copy-webpack-plugin";
import HtmlWebpackPlugin from "html-webpack-plugin";
import {getHttpsServerOptions} from "office-addin-dev-certs";
import process from "process";
// import ProvidePlugin from "webpack/lib/ProvidePlugin.js"; <-- makes react ambient - probably don't need it
import * as url from 'url';
const __dirname = url.fileURLToPath(new URL('.', import.meta.url));
import path from "path";
// import {BundleAnalyzerPlugin} from 'webpack-bundle-analyzer';
// const BundleAnalyzerPlugin = require('webpack-bundle-analyzer').BundleAnalyzerPlugin;
const fabricAsyncLoaderInclude = (await import('@fluentui/webpack-utilities/lib/fabricAsyncLoaderInclude.js')).default

const urlDev = "https://localhost:3000/";
const urlProd = "https://2yy7ugzn26vanfsfnu8njaml.z22.web.core.windows.net/";
export default async (env, options) => {
  const dev = options.mode === "development";
  const buildType = dev ? "dev" : "prod";
  const config = {
    output: {
      path: path.resolve(__dirname, "dist"),
      assetModuleFilename: 'assets/[hash][ext][query]'
    },
    devtool: "source-map",
    entry: {
      taskpane: "./src/taskpane/taskpane.tsx",
      commands: "./src/commands/commands.ts"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    stats: {
      optimizationBailout: true
    },
    module: {
      rules: [
        {
          test: /\.(jsx?|tsx?)$/,
          exclude: [/node_modules/, /otherSrc/],
          include: fabricAsyncLoaderInclude,
          use: [
            '@fluentui/webpack-utilities/lib/fabricAsyncLoader.js',
            "ts-loader"
          ]
        },
        {
          test: /\.(jsx?|tsx?)$/,
          exclude: [/node_modules/, /otherSrc/, fabricAsyncLoaderInclude],
          use: [
            "ts-loader"
          ]
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
        },
        {
          test: /\.(png|jpg|jpeg|gif|svg)$/,
          type: 'asset/resource',
        },
      ]
    },
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        //chunks: ["polyfill", "taskpane", "vendor"], <-- doesn't appear to be needed to include all chunks
        chunks: ["taskpane"],
        minify: true
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
        minify: true
      }),
      // new HtmlWebpackInlineSVGPlugin(), <-- exporting as an asset. no-longer inlining.
      new CopyWebpackPlugin({
        patterns: [
          {
            to: "taskpane.css",
            from: "./src/taskpane/taskpane.css"
          },
          {
            to: "[name]." + buildType + "[ext]",
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

      /*       new ProvidePlugin({ <-- doesn't appear to be needed
              React: "react"
            }), */
      // new BundleAnalyzerPlugin()
      // new webpack.optimize.ModuleConcatenationPlugin(), <-- doesn't appear to be needed
    ],
    optimization: {
      splitChunks: {
        chunks: "all",
        usedExports: true,
      }
    },
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
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

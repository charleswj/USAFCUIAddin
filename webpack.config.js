/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const fs = require("fs");
const webpack = require("webpack");

const urlDev="https://cep2.mail.us.af.mil/CUI_ADDIn/";
const urlProd="https://cep2.mail.us.af.mil/CUI_ADDIn/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const buildType = dev ? "dev" : "prod";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      //taskpane: "./src/taskpane/taskpane.js",
      //commands: "./src/commands/commands.js",
      functionfile: "./FunctionFile/Functions.js",
      inserttext:"./InsertTextPane/InsertText.js"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader", 
            options: {
              presets: ["@babel/preset-env"]
            }
          }
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
            name: '[path][name].[ext]',          
          }
        }
      ]
    },
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
          template: "./InsertTextPane/InsertText.html",
        chunks: ["polyfill", "taskpane"]
      }),
      //new CopyWebpackPlugin({
      //  patterns: [
      //  {
      //    to: "taskpane.css",
      //    from: "./src/taskpane/taskpane.css"
      //  },
      //  {
      //    to: "[name]." + buildType + ".[ext]",
      //    from: "manifest*.xml",
      //    transform(content) {
      //      if (dev) {
      //        return content;
      //      } else {
      //        return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
      //      }
      //    }
      //  }
      //]}),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./FunctionFile/Functions.html",
        chunks: ["polyfill", "commands"]
      })
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },      
      https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000
    }
  };

  return config;
};

/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const webpack = require("webpack");
const dotenv = require('dotenv');

// Load environment variables from .env file
const envResult = dotenv.config();
let envKeys = {};

if (envResult.parsed) {
  console.log("ENV vars loaded from .env file:", Object.keys(envResult.parsed));
  // Create an object with properly formatted environment variables
  Object.keys(envResult.parsed).forEach(key => {
    envKeys[`process.env.${key}`] = JSON.stringify(envResult.parsed[key]);
  });
} else {
  console.warn("No .env file found or error parsing it:", envResult.error);
}

// Add process.env polyfill
envKeys["process.env"] = JSON.stringify({});

// Add debug logging to check what's being defined
console.log("Environment variables being defined:", Object.keys(envKeys).map(k => k.replace('process.env.', '')));

const urlDev = "https://localhost:3002/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

/* global require, module, process, __dirname */

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
      functions: "./src/functions/functions.js",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
      fallback: {
        // Provide node module polyfills
        "process": require.resolve("process/browser"),
        "os": require.resolve("os-browserify"),
        "path": false, // No polyfill needed
        "fs": false // No polyfill needed
      }
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
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.js",
      }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"],
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new webpack.DefinePlugin(envKeys),
      new webpack.ProvidePlugin({
        process: 'process/browser',
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "assets/prompts/*",
            to: "assets/prompts/[name][ext][query]",
          },
          {
            from: "src/prompts/*",
            to: "prompts/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content
                  .toString()
                  .replace(new RegExp(urlProd, "g"), urlDev);
              } else {
                return content
                  .toString()
                  .replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
          {
            from: "config.js",
            to: "config.js",
          },
        ],
      }),
    ],
    devServer: {
      static: [
        {
          directory: path.join(__dirname, "dist"),
          publicPath: "/"
        },
        {
          directory: path.join(__dirname, "assets"),
          publicPath: "/assets"
        }
      ],
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3002,
    },
  };

  return config;
};

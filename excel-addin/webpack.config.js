/* eslint-disable no-undef */

const path = require("path");
const webpack = require("webpack");
require("dotenv").config();
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

// AUTH_MODE=whitelist → auth.whitelist.js (por defecto)
// AUTH_MODE=sso       → auth.sso.js (requiere app en Azure AD)
const authMode = process.env.AUTH_MODE || "whitelist";

// Whitelist leída de .env — nunca van al repositorio
const allowedDomains = (process.env.ALLOWED_DOMAINS || "").split(",").map(s => s.trim()).filter(Boolean);
const allowedEmails  = (process.env.ALLOWED_EMAILS  || "").split(",").map(s => s.trim()).filter(Boolean);
// Nombre de la empresa inyectado en el bundle — configure COMPANY_NAME antes de compilar
const companyName    = process.env.COMPANY_NAME || "Empresa";

const urlDev  = "https://localhost:3000/";
// ADDIN_URL se define en Railway como la URL pública del despliegue
// Ejemplo: https://asistente-excel.railway.app/
const urlProd = (process.env.ADDIN_URL || "https://localhost:3000/").replace(/\/?$/, "/");

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
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
      alias: {
        // Inyecta la implementación de auth según AUTH_MODE
        "./auth.js": path.resolve(__dirname, `src/taskpane/auth.${authMode}.js`),
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
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        __ALLOWED_DOMAINS__: JSON.stringify(allowedDomains),
        __ALLOWED_EMAILS__:  JSON.stringify(allowedEmails),
        // API_KEY inyectada en build time — nunca hardcodeada en el código
        __API_KEY__: JSON.stringify(process.env.API_KEY || ""),
        // Nombre de la empresa para el tema "Empresa" (botón selector y subtítulo)
        __COMPANY_NAME__: JSON.stringify(companyName),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
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
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      proxy: [
        {
          context: [
            "/ask", "/edit", "/analizar", "/health",
            "/addin-config", "/tiene-vinculo", "/enviar-al-bot",
            "/vincular-addin",
          ],
          target: "http://localhost:8000",
        },
      ],
    },
  };

  return config;
};

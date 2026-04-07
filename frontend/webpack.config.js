const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const devServerPort = 3000;

// The public URL where the add-in frontend will be served.
// In local dev this is always https://localhost:3000 (webpack-dev-server).
// In Docker/OpenShift builds, pass FRONTEND_URL=https://your-app.example.com
// so the correct URL gets baked into manifest.xml at build time.
const FRONTEND_URL = process.env.FRONTEND_URL || `https://localhost:${devServerPort}`;

// Office.js source URL. Defaults to Microsoft's CDN.
// For enclosed/air-gapped networks, set OFFICE_JS_SRC=/assets/office.js and
// drop a downloaded copy of office.js into frontend/public/assets/office.js
// before building. See AIRGAP.md.
const OFFICE_JS_SRC =
  process.env.OFFICE_JS_SRC ||
  "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";

module.exports = (env, argv) => {
  const isDev = argv.mode === "development";

  return {
    entry: {
      taskpane: "./src/index.tsx",
      commands: "./src/commands/commands.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx"],
      alias: {
        "@engine": path.resolve(__dirname, "src/engine"),
        "@services": path.resolve(__dirname, "src/services"),
        "@shared": path.resolve(__dirname, "src/shared"),
        "@components": path.resolve(__dirname, "src/taskpane/components"),
        "@hooks": path.resolve(__dirname, "src/taskpane/hooks"),
      },
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./public/taskpane.html",
        filename: "taskpane.html",
        chunks: ["taskpane"],
        officeJsSrc: OFFICE_JS_SRC,
      }),
      new HtmlWebpackPlugin({
        template: "./public/commands.html",
        filename: "commands.html",
        chunks: ["commands"],
        officeJsSrc: OFFICE_JS_SRC,
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "manifest.xml",
            to: "manifest.xml",
            // Replace all localhost:3000 references with the target URL.
            // In dev builds FRONTEND_URL === "https://localhost:3000" so nothing changes.
            // In Docker builds FRONTEND_URL is the real OpenShift/production URL.
            transform(content) {
              return content
                .toString()
                .replace(/https:\/\/localhost:3000/g, FRONTEND_URL);
            },
          },
          { from: "public/assets", to: "assets" },
        ],
      }),
    ],
    devServer: {
      port: devServerPort,
      server: "https",
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      static: {
        directory: path.resolve(__dirname, "dist"),
      },
      // Proxy /api and /health to the FastAPI backend.
      // This keeps all requests on the same HTTPS origin so the browser
      // never makes an HTTP fetch from an HTTPS page (mixed-content block).
      proxy: [
        {
          context: ["/api", "/health"],
          target: "http://localhost:8000",
          changeOrigin: true,
          secure: false,
        },
      ],
    },
    devtool: isDev ? "source-map" : false,
  };
};

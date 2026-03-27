const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const devServerPort = 3000;

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
      }),
      new HtmlWebpackPlugin({
        template: "./public/commands.html",
        filename: "commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [{ from: "manifest.xml", to: "manifest.xml" },
        {from: "public/assets", to: "assets" },]
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

const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = async (env, argv) => {
  const dev = argv.mode === "development";
  const httpsOptions = dev ? await devCerts.getHttpsServerOptions() : {};

  return {
    entry: {},
    output: {
      path: path.resolve(__dirname, "dist"),
      clean: true
    },
    devServer: {
      static: [
        { directory: path.resolve(__dirname, "src"), publicPath: "/src" },
        { directory: path.resolve(__dirname, "templates"), publicPath: "/templates" },
        { directory: path.resolve(__dirname, "assets"), publicPath: "/assets" },
        { directory: path.resolve(__dirname, "addons"), publicPath: "/addons" }
      ],
      server: {
        type: "https",
        options: httpsOptions
      },
      port: 3001,
      headers: {
        "Access-Control-Allow-Origin": "*"
      }
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          { from: "src", to: "src" },
          { from: "templates", to: "templates" },
          { from: "assets", to: "assets" },
          { from: "addons", to: "addons" }
        ]
      })
    ]
  };
};

// Note this only includes basic configuration for development mode.
// For a more comprehensive configuration check:
// https://github.com/fable-compiler/webpack-config-template

var path = require("path");

module.exports = {
    mode: "development",
    entry: "./src/App.fsproj",
    output: {
        path: path.join(__dirname, "./public"),
        filename: "bundle.js",
    },
    devServer: {
        static: "./public",
        port: 5000,
        https: true
    },
    module: {
        rules: [{
            test: /\.fs(x|proj)?$/,
            enforce: "pre",
            use: "fable-loader"
        },
        {
            test: /\.(sa|c)ss$/,
            use: [
                "style-loader",
                "css-loader"
            ]
        }]
    }
}

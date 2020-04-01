const fs = require('fs');
const path = require('path');
const webpack = require('webpack');
const webpackMerge = require('webpack-merge');
const commonConfig = require('./webpack.common.js');

module.exports = async (env, options) => {
    const config = webpackMerge(commonConfig, {
    devtool: 'eval-source-map',
    devServer: {
        publicPath: '/',
        contentBase: path.resolve('dist'),
        hot: true,
        compress: true,
        overlay: {
            warnings: false,
            errors: true
        },
        historyApiFallback: true,
	headers: {
            "Access-Control-Allow-Origin": "*"
	},
	https: {
            key: fs.readFileSync('C:/git/etc/live/excelint-addin.westus2.cloudapp.azure.com-0002/privkey.pem'),
            cert: fs.readFileSync('C:/git/etc/letsencrypt/live/excelint-addin.westus2.cloudapp.azure.com-0002/fullchain.pem'),
        },
	allowedHosts: [
            '.amazonaws.com', '.azure.com'
        ]
    },
    plugins: [
        new webpack.HotModuleReplacementPlugin()
    ]});
    return config;
}

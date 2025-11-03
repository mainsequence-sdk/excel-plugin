const path = require('path');
const Dotenv = require('dotenv-webpack');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
  entry: './src/functions/index.ts',
  output: {
    filename: 'index.js',
    path: path.resolve(__dirname, 'dist'),
    library: 'CustomFunctions',
    libraryTarget: 'var'
  },
  resolve: {
    extensions: ['.ts', '.js']
  },
  module: {
    rules: [
      { test: /\.ts$/, use: 'ts-loader', exclude: /node_modules/ }
    ]
  },
  plugins: [
    new Dotenv({ path: path.resolve(__dirname, '..', '.env') }),
    new HtmlWebpackPlugin({
      template: path.resolve(__dirname, 'src/ui/index.html'),
      filename: 'index.html',
      inject: false
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: path.resolve(__dirname, 'functions.json'), to: path.resolve(__dirname, 'dist', 'functions.json') }
      ]
    })
  ],
  devServer: {
    static: { directory: path.join(__dirname, 'dist') },
    compress: true,
    port: 3000,
    https: true
  }
};

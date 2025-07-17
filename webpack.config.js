const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
  entry: {
    taskpane: './src/taskpane/taskpane.ts',
    commands: './src/commands/commands.ts'
  },
  resolve: {
    extensions: ['.ts', '.tsx', '.html', '.js']
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        exclude: /node_modules/,
        use: 'ts-loader'
      },
      {
        test: /\.html$/,
        exclude: /node_modules/,
        use: 'html-loader'
      },
      {
        test: /\.(png|jpg|jpeg|gif)$/,
        use: 'file-loader'
      }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: 'taskpane.html',
      template: './src/taskpane/taskpane.html',
      chunks: ['taskpane'],
      inject: 'body',
      scriptLoading: 'blocking'
    }),
    new HtmlWebpackPlugin({
      filename: 'commands.html',
      template: './src/commands/commands.html',
      chunks: ['commands'],
      inject: 'body',
      scriptLoading: 'blocking'
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: './manifest.xml',
          to: 'manifest.xml'
        },
        {
          from: './assets',
          to: 'assets'
        }
      ]
    })
  ],
  devServer: {
    static: [
      {
        directory: path.join(__dirname, 'dist'),
        publicPath: '/'
      },
      {
        directory: path.join(__dirname, 'assets'),
        publicPath: '/assets'
      }
    ],
    port: 3000,
    server: 'https',
    open: false,
    hot: true,
    compress: true,
    allowedHosts: 'all',
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, PATCH, OPTIONS',
      'Access-Control-Allow-Headers': 'X-Requested-With, content-type, Authorization'
    }
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].js'
  }
};

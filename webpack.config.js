const webpack=require("webpack");
module.exports = {
    entry: {
        //indexedDB: __dirname + '/dbindex.js',
        "ExcelSpirit": __dirname + '/ExcelSpirit.js',
    },
    output: {
        path: __dirname + '/dist/js',
        filename: '[name].js'
    },
     module: {
        loaders: [
            { test: /\.js$/, exclude: /node_modules/, loader: "babel-loader" }
        ]
    }, 
    plugins:[
        /* new webpack.optimize.UglifyJsPlugin({
            output: {
                comments: false,  // remove all comments
              },
            compress: {
              warnings: false
            }
          }) */
      ]
}
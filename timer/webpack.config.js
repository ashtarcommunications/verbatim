const path = require('path');

module.exports = {
    entry: './dist/timer.js',
    output: {
        filename: 'timer.min.js',
        path: path.resolve(__dirname, 'dist'),
    },
    devServer: {
        static: {
            directory: path.join(__dirname, 'dist'),
        },
    },
};

/// <binding BeforeBuild='default' />


var gulp = require('gulp');
var concat = require('gulp-concat');
var del = require('del');
var filelog = require('gulp-filelog');
var fs = require('fs');
var path = require('path');
var insert = require('gulp-insert');
var merge = require('merge-stream');

var config = {
    obj: './obj',
    bin: './bin/Debug',
    reservedDirPattern: /\.vs|bin|obj|node_modules/i
}


function getFolders(dir) {
    return fs.readdirSync(dir)
        .filter(function (file) {
            return fs.statSync(path.join(dir, file)).isDirectory();
        })
        .filter(function (file) {
            return !file.match(config.reservedDirPattern);
        });
}


gulp.task("clean-obj", function () {
    return del(config.obj + '/**/*')
});

gulp.task("compile-regions", ['clean-obj'], function () {
    var project = path.basename(__dirname);
    var folders = getFolders('.');
    var regions = folders
        .map(function (folder) {
            return gulp.src(path.join(folder, '/*.ps1'))
                .pipe(filelog('ps1 brick'))
                .pipe(concat(folder + '.obj.psm1', { newLine: '\r\n' }))
                .pipe(insert.wrap('#region ' + folder + '\r\n', '\r\n#endregion'))
                .pipe(gulp.dest(config.obj))
                .pipe(filelog('obj.psm1 region assembly'));
        });
    return merge(regions)
        .pipe(filelog(project))
        .pipe(concat(project + '.psm1', { newLine: '\r\n\r\n\r\n' }))
        .pipe(gulp.dest(config.bin))
        .pipe(filelog('psm1'));
});




//Set a default tasks
gulp.task('default', ['compile-regions'], function () { });

/// <vs BeforeBuild='default' />
var gulp = require('gulp');
var concat = require('gulp-concat');

var config = {
    componentSrc: [
        'Components/deg.sharepoint.module.user.js',
        'Components/deg.sharepoint.module.file.js',
    ]
}

gulp.task('generate-components', function () {
    return gulp.src(config.componentSrc)
        .pipe(concat('test.js'))
        .pipe(gulp.dest('./'));

});

gulp.task('scripts', ['generate-components'], function () { });

//Set a default tasks
gulp.task('default', ['scripts'], function () { });

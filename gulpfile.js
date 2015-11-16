/// <vs BeforeBuild='default' />
var gulp = require('gulp');
var concat = require('gulp-concat');
var uglify = require('gulp-uglify');
var rename = require('gulp-rename');

var config = {
    distribution_dir: "dist",
    componentSrc: [
        'src/components/deg.sharepoint.module.init.js',
        'src/components/deg.sharepoint.module.common.js',
        'src/components/deg.sharepoint.module.column.js',
        'src/components/deg.sharepoint.module.contenttype.js',
        'src/components/deg.sharepoint.module.file.js',
        'src/components/deg.sharepoint.module.group.js',
        'src/components/deg.sharepoint.module.item.js',
        'src/components/deg.sharepoint.module.list.js',
        'src/components/deg.sharepoint.module.propertybag.js',
        'src/components/deg.sharepoint.module.user.js',
        'src/components/deg.sharepoint.module.taxonomy.js',
        'src/components/deg.sharepoint.module.service.js'
    ]
}

gulp.task('generate-components', function () {
    return gulp.src(config.componentSrc)
        .pipe(concat('deg.sharepoint.module.js'))
        .pipe(gulp.dest('./src/'));

});

gulp.task('generate-dist', function () {
    return gulp.src(config.componentSrc)
        .pipe(concat('deg.sharepoint.module.js'))
        .pipe(gulp.dest(config.distribution_dir))
        .pipe(rename('deg.sharepoint.module.min.js'))
        .pipe(uglify())
        .pipe(gulp.dest(config.distribution_dir));

});

gulp.task('scripts', ['generate-components', 'generate-dist'], function () { });

//Set a default tasks
gulp.task('default', ['scripts'], function () { });

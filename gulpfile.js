/// <vs BeforeBuild='default' />
var gulp = require('gulp');
var concat = require('gulp-concat');

var config = {
    componentSrc: [
        'Components/deg.sharepoint.module.init.js',
        'Components/deg.sharepoint.module.common.js',
        'Components/deg.sharepoint.module.column.js',
        'Components/deg.sharepoint.module.contenttype.js',
        'Components/deg.sharepoint.module.file.js',
        'Components/deg.sharepoint.module.group.js',
        'Components/deg.sharepoint.module.item.js',
        'Components/deg.sharepoint.module.list.js',
        'Components/deg.sharepoint.module.propertybag.js',
        'Components/deg.sharepoint.module.user.js',
        'Components/deg.sharepoint.module.service.js'        
    ]
}

gulp.task('generate-components', function () {
    return gulp.src(config.componentSrc)
        .pipe(concat('deg.sharepoint.module.js'))
        .pipe(gulp.dest('./'));

});

gulp.task('scripts', ['generate-components'], function () { });

//Set a default tasks
gulp.task('default', ['scripts'], function () { });

'use strict';

var gulp = require('gulp');

var webserver = require('gulp-webserver');

/**
 * Default task
 */
console.log('✩ Mr. Universe ✩');
gulp.task('default', () => {
});

/**
 * Server Tasks
 */
gulp.task('server', () => {
  gulp.src('./appcompose/home/home.html')
    .pipe(webserver({
      open: true
    }));
});

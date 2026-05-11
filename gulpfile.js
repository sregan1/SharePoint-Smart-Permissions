'use strict';

const build = require('@microsoft/sp-build-web');
const gulp = require('gulp');

build.addSuppression(/Warning - sass/);

build.initialize(gulp);

gulp.task('serve', gulp.series('serve-deprecated'));

#!/bin/bash
# Gradle build script to bypass shell configuration issues
cd "$(dirname "$0")"
java -classpath gradle/wrapper/gradle-wrapper.jar org.gradle.wrapper.GradleWrapperMain "$@"

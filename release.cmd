@ECHO OFF
@SETLOCAL

SET "POM_DIR=%~dp0"

REM go to mvn package
CD %POM_DIR%
mvn clean deploy -Dmaven.test.skip=true -Prelease

PAUSE &
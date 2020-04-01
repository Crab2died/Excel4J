@ECHO OFF
@SETLOCAL

SET "POM_DIR=%~dp0"

REM go to mvn package
CD %POM_DIR%
mvn clean deploy -Dmaven.test.skip=true -Prelease ^
-Dmaven.multiModuleProjectDirectory=%MAVEN_HOME% -Dmaven.wagon.http.ssl.insecure=true ^
-Dmaven.wagon.http.ssl.allowall=true -Dmaven.wagon.http.ssl.ignore.validity.dates=true

PAUSE &
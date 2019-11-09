@echo off
rem If you want to redefine variables GROOVY_HOME and JAVA_HOME you can simply create env.bat file
rem and redefine variables in this file.

if exist env.bat @call env.bat
set SCRIPT_CLASS_PATH=../lib/jxl.jar
set SCRIPT_PATH=../script/daily-report-export.groovy
set SCRIPT_DATA_SOURCE=vis-skoly

%GROOVY_HOME%\bin\groovy -cp %SCRIPT_CLASS_PATH% %SCRIPT_PATH% -s %SCRIPT_DATA_SOURCE% %*

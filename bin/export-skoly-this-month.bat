@echo off
rem If you want to redefine variables GROOVY_HOME and JAVA_HOME you can simply create env.bat file
rem and redefine variables in this file.

if exist env.bat @call env.bat
set SCRIPT_CLASS_PATH=../lib/jxl.jar
set SCRIPT_PATH=../script/export.groovy
set SCRIPT_DATA_SOURCE=vis-skoly
set SCRIPT_MONTH_SHIFT=0

%GROOVY_HOME%\bin\groovy -cp %SCRIPT_CLASS_PATH% %SCRIPT_PATH% -s %SCRIPT_DATA_SOURCE% -m %SCRIPT_MONTH_SHIFT%

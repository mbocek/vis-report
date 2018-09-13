@echo off
rem If you want to redefine variables GROOVY_HOME and JAVA_HOME you can simply create env.bat file
rem and redefine variables in this file.

IF "%GROOVY_HOME%"=="" echo GROOVY_HOME environment variable should point to groovy home directory
IF "%JAVA_HOME%"=="" echo JAVA_HOME environment variable should point to java 32 bit version home directory

if exist env.bat @call env.bat

%GROOVY_HOME%\bin\groovy -cp lib/jxl.jar daily-report-export.groovy %*
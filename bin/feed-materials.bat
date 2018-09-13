@echo off
rem If you want to redefine variables GROOVY_HOME and JAVA_HOME you can simply create env.bat file
rem and redefine variables in this file.

if exist env.bat @call env.bat

%GROOVY_HOME%\bin\groovy feed-materials.groovy %*
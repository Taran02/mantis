set projectLocation=%~dp0
cd %projectLocation%
set classpath=%projectLocation%\bin;%projectLocation%\Jars\*
java Mantis
pause
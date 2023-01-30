@echo off

echo Setting path to python
@REM Add python to the path
set PATH=%PATH%;C:\Users\%USERNAME%\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0

echo Setting path to geckodriver
@REM Add geckodriver.exe to path
set PATH=%PATH%;%CD%\assets\geckodriver.exe

echo Done! If the program does not start, please make sure that python was installed correctly

python da.py

pause
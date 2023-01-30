@echo off

echo Setting path to python
@REM Add python to the path
set PATH=%PATH%;C:\Users\%USERNAME%\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0

echo Installing python modules
pip install -r assets\requirements.txt

pause
"%PYTHON%" setup.py install --single-version-externally-managed --record=record.txt
if errorlevel 1 exit 1
copy %SRC_DIR%\bin\* %PREFIX%
if errorlevel 1 exit 1
mkdir %PREFIX%\Tools\axl
if errorlevel 1 exit 1
copy %SRC_DIR%\tools\* %PREFIX%\Tools\axl
if errorlevel 1 exit 1

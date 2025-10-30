@echo off
echo ========================================
echo Excel Splitter - Windows Run Script
echo ========================================

echo.
echo [1/3] Installing Python packages...
pip install pandas numpy openpyxl python-dotenv
if %ERRORLEVEL% neq 0 (
    echo ERROR: Failed to install Python packages
    pause
    exit /b 1
)
echo Python packages installed successfully.

echo.
echo [2/3] Executing excel_splitter.py...
python excel_splitter.py
if %ERRORLEVEL% neq 0 (
    echo ERROR: Failed to execute excel_splitter.py
    pause
    exit /b 1
)
echo excel_splitter.py executed successfully.

echo.
echo [3/3] Zipping output folders...
for /d %%d in (output\*) do (
    echo Zipping %%d...
    powershell -command "Compress-Archive -Path '%%d' -DestinationPath '%%d.zip' -Force"
    if %ERRORLEVEL% neq 0 (
        echo WARNING: Failed to zip %%d
    ) else (
        echo Successfully zipped %%d to %%d.zip
    )
)

echo.
echo ========================================
echo Process completed successfully!
echo ========================================
echo.
echo Check the output folder for zipped results:
dir /b output\*.zip

pause

#!/bin/bash

echo "========================================"
echo "Excel Splitter - Linux Run Script"
echo "========================================"

echo
echo "[1/3] Installing Python packages..."
pip install pandas numpy openpyxl python-dotenv
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install Python packages"
    exit 1
fi
echo "Python packages installed successfully."

echo
echo "[2/3] Executing excel_splitter.py..."
python excel_splitter.py
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to execute excel_splitter.py"
    exit 1
fi
echo "excel_splitter.py executed successfully."

echo
echo "[3/3] Zipping output folders..."
if [ -d "output" ]; then
    for folder in output/*/; do
        if [ -d "$folder" ]; then
            folder_name=$(basename "$folder")
            echo "Zipping $folder..."
            zip -r "output/${folder_name}.zip" "$folder"
            if [ $? -eq 0 ]; then
                echo "Successfully zipped $folder to output/${folder_name}.zip"
            else
                echo "WARNING: Failed to zip $folder"
            fi
        fi
    done
else
    echo "WARNING: Output folder not found"
fi

echo
echo "========================================"
echo "Process completed successfully!"
echo "========================================"
echo
echo "Check the output folder for zipped results:"
ls -la output/*.zip 2>/dev/null || echo "No zip files found in output folder"

@echo off
echo.
echo ===================================================
echo   JANVI ORDERS TRACKER CONSOLIDATED AUTO-UPDATE
echo ===================================================
echo.
echo Loading new Shopify and Shiprocket files from Data/ ...
python generate_sheet.py
if %ERRORLEVEL% EQU 0 (
    echo.
    echo Syncing data to Google Sheets...
    python sync_via_web_app.py
)
echo.
echo ===================================================
echo   Update completed successfully!
echo   Excel file generated: Janvi_Consolidated_Orders_2026_2027.xlsx
echo ===================================================
echo.
pause

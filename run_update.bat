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
echo Syncing updates to GitHub for Vercel web app...
git add data.js Janvi_Consolidated_Orders_2026_2027.xlsx
git commit -m "Auto-update consolidated order data and Excel sheet"
if %ERRORLEVEL% EQU 0 (
    echo Pushing latest dataset to Vercel...
    git push
) else (
    echo No new changes to push.
)
echo.
echo ===================================================
echo   Update completed successfully!
echo   Excel file generated: Janvi_Consolidated_Orders_2026_2027.xlsx
echo ===================================================
echo.
pause

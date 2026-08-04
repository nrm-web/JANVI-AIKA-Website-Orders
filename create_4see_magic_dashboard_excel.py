import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

def generate_4see_magic_excel():
    print("Loading master consolidated data from Janvi_Consolidated_Orders_2026_2027.xlsx...")
    master_file = "Janvi_Consolidated_Orders_2026_2027.xlsx"
    
    if not os.path.exists(master_file):
        print(f"Error: {master_file} not found. Running generate_sheet.py first...")
        import generate_sheet
        generate_sheet.process_and_create_excel()
        
    df_master = pd.read_excel(master_file, sheet_name="Master Sheet")
    # Filter out footer summary row if present
    df_master = df_master[df_master['Order No'].astype(str).str.startswith('#')].copy()
    print(f"Loaded {len(df_master)} actual master orders starting with '#' from {master_file}.")
    
    # ----------------------------------------------------
    # CALCULATE EXACT 1:1 METRICS FROM MASTER SHEET
    # ----------------------------------------------------
    total_orders = len(df_master)
    total_revenue = float(df_master['Total Price'].sum())
    
    canceled_df = df_master[df_master['Fulfillment Status'].astype(str).str.upper().str.contains("CANCELED|CANCELLED", na=False)]
    canceled_count = len(canceled_df)
    canceled_val = float(canceled_df['Total Price'].sum())
    
    denied_df = df_master[((df_master['Fulfillment Status'].astype(str).str.upper().str.contains("RTO|DENIED", na=False)) | (df_master['COD Denies (Yes/No)'] == "Yes")) & (~df_master['Fulfillment Status'].astype(str).str.upper().str.contains("CANCELED|CANCELLED", na=False))]
    denied_count = len(denied_df)
    denied_val = float(denied_df['Total Price'].sum())
    
    returned_df = df_master[(df_master['Returned (True/False)'] == True) & (~df_master['Fulfillment Status'].astype(str).str.upper().str.contains("CANCELED|CANCELLED|RTO|DENIED", na=False))]
    returned_count = len(returned_df)
    returned_val = float(returned_df['Total Price'].sum())
    
    successful_df = df_master[df_master['Fulfillment Status'].astype(str).str.upper().str.strip().isin(["DELIVERED", "SELF FULFILED", "FULFILLED"]) & (df_master['Returned (True/False)'] == False) & (df_master['COD Denies (Yes/No)'] == "No") & (~df_master['Fulfillment Status'].astype(str).str.upper().str.contains("CANCELED|CANCELLED", na=False))]
    successful_count = len(successful_df)
    successful_val = float(successful_df['Total Price'].sum())
    
    net_revenue_val = total_revenue - canceled_val - denied_val - returned_val
    net_orders_count = total_orders - canceled_count - denied_count - returned_count
    
    aov_val = net_revenue_val / successful_count if successful_count > 0 else 0.0
    
    # COD vs Prepaid Breakdown
    cod_delivered = df_master[(df_master['COD (Yes/No)'] == "Yes") & df_master['Fulfillment Status'].astype(str).str.upper().str.strip().isin(["DELIVERED", "SELF FULFILED", "FULFILLED"]) & (df_master['Returned (True/False)'] == False) & (df_master['COD Denies (Yes/No)'] == "No")]
    cod_net_revenue = float(cod_delivered['Total Price'].sum())
    cod_delivered_count = len(cod_delivered)
    cod_aov = cod_net_revenue / cod_delivered_count if cod_delivered_count > 0 else 0.0
    
    prepaid_delivered = df_master[(df_master['Prepaid (Yes/No)'] == "Yes") & df_master['Fulfillment Status'].astype(str).str.upper().str.strip().isin(["DELIVERED", "SELF FULFILED", "FULFILLED"]) & (df_master['Returned (True/False)'] == False)]
    prepaid_net_revenue = float(prepaid_delivered['Total Price'].sum())
    prepaid_delivered_count = len(prepaid_delivered)
    prepaid_aov = prepaid_net_revenue / prepaid_delivered_count if prepaid_delivered_count > 0 else 0.0
    
    # Active Pipeline Breakdown
    in_transit_df = df_master[df_master['Fulfillment Status'].astype(str).str.upper().str.contains("TRANSIT|SHIPPED|OUT FOR DELIVERY|REACHED DESTINATION HUB", na=False)]
    pickup_df = df_master[df_master['Fulfillment Status'].astype(str).str.upper().str.contains("PICKUP|MANIFESTED", na=False)]
    unfulfilled_df = df_master[df_master['Fulfillment Status'].astype(str).str.upper().str.contains("NEW ORDER|UNFULFILLED|PENDING", na=False)]
    
    in_progress_count = len(in_transit_df) + len(pickup_df) + len(unfulfilled_df)
    in_progress_val = float(in_transit_df['Total Price'].sum() + pickup_df['Total Price'].sum() + unfulfilled_df['Total Price'].sum())
    
    # Output file
    output_path = "Janvi_4see_Magic_Dashboard_Master.xlsx"
    print(f"Generating 4see Magic Dashboard Master file: {output_path}...")
    
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet
    
    font_family = "Inter"
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    font_regular = Font(name=font_family, size=10)
    
    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )
    
    def write_df_clean(sheet, df_data):
        sheet.views.sheetView[0].showGridLines = True
        
        # Headers directly at Row 1 (No title offset so 4see reads column names perfectly)
        for col_idx, col_name in enumerate(df_data.columns, 1):
            cell = sheet.cell(row=1, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        sheet.row_dimensions[1].height = 28
        
        # Data starting at Row 2
        for row_idx, r in df_data.iterrows():
            curr_row = row_idx + 2
            for col_idx, val in enumerate(r, 1):
                cell = sheet.cell(row=curr_row, column=col_idx, value=val)
                cell.font = font_regular
                cell.border = thin_border
                
                # Format numbers / currency / percentage
                col_name_str = df_data.columns[col_idx-1]
                if isinstance(val, (int, float)):
                    if any(kw in col_name_str for kw in ["Price", "Revenue", "Amount", "Value", "AOV"]):
                        cell.number_format = '₹#,##0.00'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif any(kw in col_name_str for kw in ["Rate", "Pct", "Share"]):
                        cell.number_format = '0.00%'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.number_format = '#,##0'
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    
        # Auto-fit column widths
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 15)

    # 1. Executive KPIs Sheet
    df_exec = pd.DataFrame([
        {"KPI_ID": "KPI-01", "Metric_Title": "Total Orders", "Primary_Value": total_orders, "Unit": "Orders", "Subtext": "Active Shopify Orders", "Progress_Pct": 1.0, "Theme_Color": "Blue"},
        {"KPI_ID": "KPI-02", "Metric_Title": "Total Revenue", "Primary_Value": total_revenue, "Unit": "INR ₹", "Subtext": f"From {total_orders} active orders", "Progress_Pct": 1.0, "Theme_Color": "Green"},
        {"KPI_ID": "KPI-03", "Metric_Title": "Net Revenue", "Primary_Value": net_revenue_val, "Unit": "INR ₹", "Subtext": f"{net_orders_count} net orders ({successful_count} delivered + {in_progress_count} in progress)", "Progress_Pct": 1.0, "Theme_Color": "Green"},
        {"KPI_ID": "KPI-04", "Metric_Title": "Average Order Value (AOV)", "Primary_Value": aov_val, "Unit": "INR ₹", "Subtext": f"Net revenue / {successful_count} successful", "Progress_Pct": 1.0, "Theme_Color": "Blue"}
    ])
    ws1 = wb.create_sheet("01_Executive_KPIs")
    write_df_clean(ws1, df_exec)

    # 2. Payment Performance KPIs Sheet (Split Full COD vs Partial COD vs Prepaid)
    full_cod_total = df_master[df_master['Payment Method'] == 'COD']
    partial_cod_total = df_master[df_master['Payment Method'] == 'Partial COD']
    
    df_pay_kpi = pd.DataFrame([
        {"KPI_ID": "PAY-01", "Metric_Title": "Full COD Orders", "Primary_Value": len(full_cod_total), "Payment_Type": "Full COD", "Subtext": f"55 full COD orders (₹{full_cod_total['Total Price'].sum():,.2f})", "Theme_Color": "Orange"},
        {"KPI_ID": "PAY-02", "Metric_Title": "Partial COD Orders", "Primary_Value": len(partial_cod_total), "Payment_Type": "Partial COD", "Subtext": f"21 Razorpay Partial COD orders (₹{partial_cod_total['Total Price'].sum():,.2f})", "Theme_Color": "Orange"},
        {"KPI_ID": "PAY-03", "Metric_Title": "COD Net Revenue", "Primary_Value": cod_net_revenue, "Payment_Type": "COD", "Subtext": f"From {cod_delivered_count} delivered COD orders", "Theme_Color": "Orange"},
        {"KPI_ID": "PAY-04", "Metric_Title": "Prepaid Net Revenue", "Primary_Value": prepaid_net_revenue, "Payment_Type": "Prepaid", "Subtext": f"From {prepaid_delivered_count} delivered Prepaid orders", "Theme_Color": "Green"},
        {"KPI_ID": "PAY-05", "Metric_Title": "Avg Prepaid Order Value", "Primary_Value": prepaid_aov, "Payment_Type": "Prepaid", "Subtext": f"Prepaid net revenue / {prepaid_delivered_count} successful", "Theme_Color": "Green"}
    ])
    ws2 = wb.create_sheet("02_Payment_KPIs")
    write_df_clean(ws2, df_pay_kpi)

    # 3. Master Order Lifecycle Breakdown Sheet
    df_lifecycle = pd.DataFrame([
        {"Lifecycle_Stage": "Successful Orders", "Order_Count": successful_count, "Total_Amount_INR": successful_val, "Stage_Rate_Pct": successful_count/total_orders if total_orders > 0 else 0.0, "Subtext": f"{successful_count} delivered of {total_orders} total orders", "Description": "Delivered & self fulfilled", "Status_Color": "Green"},
        {"Lifecycle_Stage": "Doorstep RTO Denied", "Order_Count": denied_count, "Total_Amount_INR": denied_val, "Stage_Rate_Pct": denied_count/total_orders if total_orders > 0 else 0.0, "Subtext": f"{denied_count} doorstep delivery refusals (out of 76 COD orders)", "Description": "Doorstep delivery denials", "Status_Color": "Red"},
        {"Lifecycle_Stage": "Post-Delivery Customer Returns", "Order_Count": returned_count, "Total_Amount_INR": returned_val, "Stage_Rate_Pct": returned_count/successful_count if successful_count > 0 else 0.0, "Subtext": f"{returned_count} customer post-delivery returns (out of {successful_count} delivered)", "Description": "Delivered then sent back", "Status_Color": "Orange"},
        {"Lifecycle_Stage": "Canceled Orders", "Order_Count": canceled_count, "Total_Amount_INR": canceled_val, "Stage_Rate_Pct": canceled_count/total_orders if total_orders > 0 else 0.0, "Subtext": f"{canceled_count} canceled of {total_orders} total orders", "Description": "Order not accepted", "Status_Color": "Orange"}
    ])
    ws3 = wb.create_sheet("03_Lifecycle_Metrics")
    write_df_clean(ws3, df_lifecycle)

    # 4. In-Progress Logistics Pipeline Sheet
    df_pipeline = pd.DataFrame([
        {"Pipeline_Stage": "Total In-Progress", "Order_Count": in_progress_count, "Total_Amount_INR": in_progress_val, "Active_Rate_Pct": in_progress_count/total_orders if total_orders > 0 else 0.0, "Theme_Color": "Blue"},
        {"Pipeline_Stage": "In Transit", "Order_Count": len(in_transit_df), "Total_Amount_INR": float(in_transit_df['Total Price'].sum()), "Active_Rate_Pct": len(in_transit_df)/total_orders if total_orders > 0 else 0.0, "Theme_Color": "Blue"},
        {"Pipeline_Stage": "Pickup Scheduled", "Order_Count": len(pickup_df), "Total_Amount_INR": float(pickup_df['Total Price'].sum()), "Active_Rate_Pct": len(pickup_df)/total_orders if total_orders > 0 else 0.0, "Theme_Color": "Purple"},
        {"Pipeline_Stage": "Unfulfilled", "Order_Count": len(unfulfilled_df), "Total_Amount_INR": float(unfulfilled_df['Total Price'].sum()), "Active_Rate_Pct": len(unfulfilled_df)/total_orders if total_orders > 0 else 0.0, "Theme_Color": "Amber"}
    ])
    ws4 = wb.create_sheet("04_In_Progress_Pipeline")
    write_df_clean(ws4, df_pipeline)

    # 5. Product Category Summary Sheet
    cat_rows = []
    raw_cats = df_master['Category'].astype(str).str.split(',').explode().str.strip().unique()
    clean_cats = [str(c) for c in raw_cats if str(c) not in ['nan', 'None', '-', '', 'OTHER']]
    for cat in sorted(clean_cats):
        if not cat or cat in ['nan', 'None', '-']:
            continue
        cat_df = df_master[df_master['Category'].astype(str).str.contains(cat, na=False)]
        rev = float(cat_df['Total Price'].sum())
        cnt = len(cat_df)
        share = rev / total_revenue if total_revenue > 0 else 0.0
        
        deliv = len(cat_df[cat_df['Fulfillment Status'].astype(str).str.upper().str.strip().isin(["DELIVERED", "SELF FULFILED", "FULFILLED"])])
        trans = len(cat_df[cat_df['Fulfillment Status'].astype(str).str.upper().str.contains("TRANSIT|SHIPPED", na=False)])
        unful = len(cat_df[cat_df['Fulfillment Status'].astype(str).str.upper().str.contains("NEW ORDER|UNFULFILLED", na=False)])
        deni = len(cat_df[(cat_df['COD Denies (Yes/No)'] == "Yes") | (cat_df['Fulfillment Status'].astype(str).str.upper().str.contains("DENIED|RTO", na=False))])
        retu = len(cat_df[cat_df['Returned (True/False)'] == True])
        canc = len(cat_df[cat_df['Fulfillment Status'].astype(str).str.upper().str.contains("CANCELED|CANCELLED", na=False)])
        
        cat_rows.append({
            "Category_Name": cat,
            "Total_Revenue_INR": rev,
            "Order_Count": cnt,
            "Revenue_Share_Pct": share,
            "Delivered_Count": deliv,
            "In_Transit_Count": trans,
            "Unfulfilled_Count": unful,
            "Denied_Count": deni,
            "Returned_Count": retu,
            "Canceled_Count": canc
        })
    df_cat = pd.DataFrame(cat_rows).sort_values(by="Total_Revenue_INR", ascending=False)
    ws5 = wb.create_sheet("05_Category_Performance")
    write_df_clean(ws5, df_cat)

    # 6. Payment Mode Performance Sheet
    pay_rows = []
    for mode in ["COD", "Prepaid (Razorpay)", "Partial COD"]:
        pm_df = df_master[df_master['Payment Method'] == mode]
        cnt = len(pm_df)
        if cnt == 0:
            continue
        rev = float(pm_df['Total Price'].sum())
        deliv = len(pm_df[pm_df['Fulfillment Status'].astype(str).str.upper().str.strip().isin(["DELIVERED", "SELF FULFILED", "FULFILLED"])])
        trans = len(pm_df[pm_df['Fulfillment Status'].astype(str).str.upper().str.contains("TRANSIT|SHIPPED", na=False)])
        pick = len(pm_df[pm_df['Fulfillment Status'].astype(str).str.upper().str.contains("PICKUP", na=False)])
        unful = len(pm_df[pm_df['Fulfillment Status'].astype(str).str.upper().str.contains("NEW ORDER|UNFULFILLED", na=False)])
        deni = len(pm_df[(pm_df['COD Denies (Yes/No)'] == "Yes") | (pm_df['Fulfillment Status'].astype(str).str.upper().str.contains("DENIED|RTO", na=False))])
        retu = len(pm_df[pm_df['Returned (True/False)'] == True])
        canc = len(pm_df[pm_df['Fulfillment Status'].astype(str).str.upper().str.contains("CANCELED|CANCELLED", na=False)])
        
        pay_rows.append({
            "Payment_Mode": mode,
            "Total_Orders": cnt,
            "Total_Revenue_INR": rev,
            "Delivered_Count": deliv,
            "In_Transit_Count": trans,
            "Pickup_Count": pick,
            "Unfulfilled_Count": unful,
            "Denied_Count": deni,
            "Returned_Count": retu,
            "Canceled_Count": canc,
            "Denial_Return_Rate_Pct": (deni + retu) / cnt if cnt > 0 else 0.0
        })
    df_pay_mode = pd.DataFrame(pay_rows)
    ws6 = wb.create_sheet("06_Payment_Mode_Breakdown")
    write_df_clean(ws6, df_pay_mode)

    # 7. Monthly Sales Trends Sheet
    df_master['Date_Obj'] = pd.to_datetime(df_master['Date of Order'], errors='coerce')
    df_master['Month_Year'] = df_master['Date_Obj'].dt.strftime('%b %Y')
    
    monthly_rows = []
    for month, m_df in df_master.groupby('Month_Year', sort=False):
        monthly_rows.append({
            "Month_Year": month,
            "Total_Orders": len(m_df),
            "Gross_Revenue_INR": float(m_df['Total Price'].sum()),
            "Delivered_Orders": len(m_df[m_df['Fulfillment Status'].astype(str).str.upper().str.strip().isin(["DELIVERED", "SELF FULFILED", "FULFILLED"])]),
            "Denied_Orders": len(m_df[(m_df['COD Denies (Yes/No)'] == "Yes") | (m_df['Fulfillment Status'].astype(str).str.upper().str.contains("DENIED|RTO", na=False))]),
            "Returned_Orders": len(m_df[m_df['Returned (True/False)'] == True]),
            "Canceled_Orders": len(m_df[m_df['Fulfillment Status'].astype(str).str.upper().str.contains("CANCELED|CANCELLED", na=False)])
        })
    df_monthly = pd.DataFrame(monthly_rows)
    ws7 = wb.create_sheet("07_Monthly_Sales_Trends")
    write_df_clean(ws7, df_monthly)

    # 8. Master Orders Dataset Sheet
    df_master_clean = df_master.drop(columns=['Date_Obj', 'Month_Year'])
    ws8 = wb.create_sheet("Master_Orders_Dataset")
    write_df_clean(ws8, df_master_clean)

    wb.save(output_path)
    print(f"Successfully generated clean 4see Magic Dashboard Master Excel: {output_path}")

if __name__ == "__main__":
    generate_4see_magic_excel()

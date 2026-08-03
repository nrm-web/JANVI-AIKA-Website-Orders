import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import datetime
import json

def generate_4see_magic_excel():
    print("Reading consolidated data from generate_sheet processing logic...")
    
    shopify_path = os.path.join("Data", "Shopify orders Janvi.csv")
    shiprocket_path = os.path.join("Data", "Ship rocket orders.csv")
    
    if not os.path.exists(shopify_path) or not os.path.exists(shiprocket_path):
        print("Error: Required data files missing in 'Data/' directory.")
        return
        
    s = pd.read_csv(shopify_path)
    sr = pd.read_csv(shiprocket_path)
    
    s['Name'] = s['Name'].ffill()
    s['order_id_clean'] = s['Name'].astype(str).str.extract(r'(\d+)').astype(float).fillna(-1).astype(int)
    sr['order_id_clean'] = sr['Order ID'].astype(str).str.extract(r'(\d+)').astype(float).fillna(-1).astype(int)
    
    s = s[s['order_id_clean'] != -1]
    sr = sr[sr['order_id_clean'] != -1]
    
    s_grouped = s.groupby('order_id_clean').agg({
        'Name': 'first',
        'Created at': 'first',
        'Total': 'first',
        'Payment Method': 'first',
        'Financial Status': 'first',
        'Fulfillment Status': 'first',
        'Refunded Amount': 'first',
        'Shipping City': 'first',
        'Shipping Zip': 'first',
        'Billing Name': 'first',
        'Billing Phone': 'first',
        'Lineitem name': lambda x: ', '.join(x.dropna().astype(str).unique()),
        'Lineitem sku': lambda x: ', '.join(x.dropna().astype(str))
    }).reset_index()
    
    sr_grouped = sr.groupby('order_id_clean').agg({
        'Order ID': 'first',
        'Channel Created At': 'first',
        'Order Total': 'first',
        'Payment Method': 'first',
        'Address City': 'first',
        'Address Pincode': 'first',
        'Product Name': lambda x: ', '.join(x.dropna().astype(str).unique()),
        'Status': 'first',
        'Courier Company': 'first',
        'AWB Code': 'first',
        'Latest NDR Reason': 'first',
        'Pickup Exception Reason': 'first',
        'RTO Reason': 'first',
        'Cancellation Reason': 'first',
        'Customer Name': 'first',
        'Customer Mobile': 'first'
    }).reset_index()
    
    merged = pd.merge(s_grouped, sr_grouped, on='order_id_clean', how='left')
    
    def map_sku_to_category(sku_str):
        if not sku_str or pd.isna(sku_str) or str(sku_str).strip() in ("", "-", "nan"):
            return "OTHER"
        skus = [sk.strip().upper() for sk in str(sku_str).split(",")]
        categories = []
        mapping = {
            "C": "CHUDIDHAR",
            "A": "ANARKALI",
            "L": "LEHENGA",
            "HSL": "HALF SAREE LEHENGA",
            "LG": "LONG GOWN",
            "SHA": "SHARARA",
            "TOP": "TOPS",
            "CORD": "CO-ORD SET"
        }
        for sku in skus:
            if not sku:
                continue
            prefix = sku.split("-")[0].strip()
            cat = mapping.get(prefix)
            if not cat:
                if sku.startswith("CORD"):
                    cat = "CO-ORD SET"
                else:
                    cat = "OTHER"
            categories.append(cat)
        return ", ".join(categories) if categories else "OTHER"

    rows = []
    for idx, row in merged.iterrows():
        order_no = str(row['Name'])
        cust_name = str(row['Billing Name']).strip().title() if pd.notna(row['Billing Name']) else "Unknown"
        items = str(row['Lineitem name']) if pd.notna(row['Lineitem name']) else "No items"
        date_val = pd.to_datetime(row['Created at']).tz_localize(None) if pd.notna(row['Created at']) else None
        date_str = date_val.strftime('%Y-%m-%d') if date_val is not None else "Unknown"
        price = float(row['Total']) if pd.notna(row['Total']) else 0.0
        
        pay_s = str(row['Payment Method_x']).lower() if pd.notna(row['Payment Method_x']) else ""
        fin_status_lower = str(row['Financial Status']).lower().strip() if pd.notna(row['Financial Status']) else ""
        
        is_cod = False
        pay_method = "Prepaid (Razorpay)"
        
        if 'partially' in fin_status_lower:
            pay_method = "Partial COD"
            is_cod = True
        elif 'cod' in pay_s or 'cash' in pay_s:
            pay_method = "COD"
            is_cod = True
            
        prepaid_str = "No" if is_cod else "Yes"
        cod_str = "Yes" if is_cod else "No"
        
        sr_status = str(row['Status']).strip().upper() if pd.notna(row['Status']) else ""
        s_fulfillment = str(row['Fulfillment Status']).strip().upper() if pd.notna(row['Fulfillment Status']) else ""
        
        if sr_status != "":
            status = sr_status
        elif s_fulfillment == "FULFILLED":
            status = "DELIVERED"
        else:
            status = "NEW ORDER"
            
        returned_bool = False
        if "RTO" in status or "RETURN" in status:
            returned_bool = True
            
        cod_denied_str = "No"
        ndr_reason = str(row['Latest NDR Reason']).upper() if pd.notna(row['Latest NDR Reason']) else ""
        rto_reason = str(row['RTO Reason']).upper() if pd.notna(row['RTO Reason']) else ""
        canc_reason = str(row['Cancellation Reason']).upper() if pd.notna(row['Cancellation Reason']) else ""
        
        if "REJECT" in ndr_reason or "DENI" in ndr_reason or "REJECT" in rto_reason or "DENI" in rto_reason or "REJECT" in canc_reason:
            cod_denied_str = "Yes"
            
        if "CANCELED" in status or "CANCELLED" in status:
            status = "CANCELED"
            
        city = str(row['Shipping City']).title() if pd.notna(row['Shipping City']) else "Unknown"
        pin = str(row['Shipping Zip']) if pd.notna(row['Shipping Zip']) else "Unknown"
        sku = str(row['Lineitem sku']) if pd.notna(row['Lineitem sku']) else "-"
        awb = str(row['AWB Code']) if pd.notna(row['AWB Code']) else "-"
        tracking = f"https://shiprocket.co/tracking/{awb}" if awb != "-" else "-"
        category = map_sku_to_category(sku)
        
        rows.append({
            "Order No": order_no,
            "Customer Name": cust_name,
            "Items Ordered": items,
            "Date of Order": date_str,
            "Total Price": price,
            "Payment Method": pay_method,
            "Prepaid (Yes/No)": prepaid_str,
            "COD (Yes/No)": cod_str,
            "Returned (True/False)": returned_bool,
            "COD Denies (Yes/No)": cod_denied_str,
            "City": city,
            "PIN Code": pin,
            "Fulfillment Status": status,
            "AWB Code": awb,
            "Tracking Link": tracking,
            "SKU": sku,
            "Category": category
        })
        
    df_master = pd.DataFrame(rows)
    print(f"Master Dataset created with {len(df_master)} real orders.")
    
    # CALCULATE 4SEE MAGIC DASHBOARD TABLES
    total_orders = len(df_master)
    total_revenue = df_master['Total Price'].sum()
    
    canceled_df = df_master[df_master['Fulfillment Status'].str.contains("CANCELED|CANCELLED", na=False)]
    canceled_count = len(canceled_df)
    canceled_val = canceled_df['Total Price'].sum()
    
    denied_df = df_master[(df_master['COD Denies (Yes/No)'] == "Yes") | (df_master['Fulfillment Status'].str.contains("DENIED|RTO", na=False))]
    denied_df = denied_df[~denied_df['Fulfillment Status'].str.contains("CANCELED", na=False)]
    denied_count = len(denied_df)
    denied_val = denied_df['Total Price'].sum()
    
    returned_df = df_master[(df_master['Returned (True/False)'] == True) & (~df_master['Fulfillment Status'].str.contains("CANCELED|DENIED|RTO", na=False))]
    returned_count = len(returned_df)
    returned_val = returned_df['Total Price'].sum()
    
    successful_df = df_master[df_master['Fulfillment Status'].str.contains("DELIVERED|FULFILLED", na=False) & (df_master['Returned (True/False)'] == False) & (df_master['COD Denies (Yes/No)'] == "No")]
    successful_count = len(successful_df)
    successful_val = successful_df['Total Price'].sum()
    
    net_revenue_val = total_revenue - canceled_val - denied_val - returned_val
    net_orders_count = total_orders - canceled_count - denied_count - returned_count
    
    aov_val = net_revenue_val / successful_count if successful_count > 0 else 0.0
    
    # COD vs Prepaid
    cod_delivered = df_master[(df_master['COD (Yes/No)'] == "Yes") & df_master['Fulfillment Status'].str.contains("DELIVERED|FULFILLED", na=False) & (df_master['Returned (True/False)'] == False) & (df_master['COD Denies (Yes/No)'] == "No")]
    cod_net_revenue = cod_delivered['Total Price'].sum()
    cod_delivered_count = len(cod_delivered)
    cod_aov = cod_net_revenue / cod_delivered_count if cod_delivered_count > 0 else 0.0
    
    prepaid_delivered = df_master[(df_master['Prepaid (Yes/No)'] == "Yes") & df_master['Fulfillment Status'].str.contains("DELIVERED|FULFILLED", na=False) & (df_master['Returned (True/False)'] == False)]
    prepaid_net_revenue = prepaid_delivered['Total Price'].sum()
    prepaid_delivered_count = len(prepaid_delivered)
    prepaid_aov = prepaid_net_revenue / prepaid_delivered_count if prepaid_delivered_count > 0 else 0.0
    
    # Active Pipeline
    in_transit_df = df_master[df_master['Fulfillment Status'].str.contains("TRANSIT|SHIPPED|OUT FOR DELIVERY", na=False)]
    pickup_df = df_master[df_master['Fulfillment Status'].str.contains("PICKUP|MANIFESTED", na=False)]
    unfulfilled_df = df_master[df_master['Fulfillment Status'].str.contains("NEW ORDER|UNFULFILLED|PENDING", na=False)]
    
    in_progress_count = len(in_transit_df) + len(pickup_df) + len(unfulfilled_df)
    in_progress_val = in_transit_df['Total Price'].sum() + pickup_df['Total Price'].sum() + unfulfilled_df['Total Price'].sum()
    
    # BUILD EXCEL WORKBOOK FOR 4SEE MAGIC DASHBOARD
    output_path = "Janvi_4see_Magic_Dashboard_Master.xlsx"
    print(f"Creating 4see Magic Dashboard Master Excel: {output_path}...")
    
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet
    
    font_family = "Inter"
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    font_bold = Font(name=font_family, size=10, bold=True)
    font_regular = Font(name=font_family, size=10)
    
    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )
    
    def write_df_to_sheet(sheet, df_data, title):
        sheet.views.sheetView[0].showGridLines = True
        sheet.cell(row=1, column=1, value=title).font = Font(name=font_family, size=13, bold=True, color="0F172A")
        sheet.row_dimensions[1].height = 25
        
        # Headers at Row 3
        for col_idx, col_name in enumerate(df_data.columns, 1):
            cell = sheet.cell(row=3, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        sheet.row_dimensions[3].height = 28
        
        # Data rows
        for row_idx, r in df_data.iterrows():
            curr_row = row_idx + 4
            for col_idx, val in enumerate(r, 1):
                cell = sheet.cell(row=curr_row, column=col_idx, value=val)
                cell.font = font_regular
                cell.border = thin_border
                
                # Format numbers / currency
                if isinstance(val, (int, float)):
                    if "Price" in df_data.columns[col_idx-1] or "Revenue" in df_data.columns[col_idx-1] or "Amount" in df_data.columns[col_idx-1] or "Value" in df_data.columns[col_idx-1] or "AOV" in df_data.columns[col_idx-1]:
                        cell.number_format = '₹#,##0.00'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif "Rate" in df_data.columns[col_idx-1] or "Pct" in df_data.columns[col_idx-1] or "Share" in df_data.columns[col_idx-1]:
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
                if cell.row >= 3 and cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 15)

    # 1. Executive KPIs Sheet
    df_exec = pd.DataFrame([
        {"KPI_ID": "KPI-01", "Metric_Title": "Total Orders", "Primary_Value": total_orders, "Unit": "Orders", "Subtext": "Active Shopify Orders", "Progress_Pct": 1.0, "Theme_Color": "Blue"},
        {"KPI_ID": "KPI-02", "Metric_Title": "Total Revenue", "Primary_Value": total_revenue, "Unit": "INR ₹", "Subtext": f"From {total_orders} active orders", "Progress_Pct": 1.0, "Theme_Color": "Green"},
        {"KPI_ID": "KPI-03", "Metric_Title": "Net Revenue", "Primary_Value": net_revenue_val, "Unit": "INR ₹", "Subtext": f"{net_orders_count} net active orders", "Progress_Pct": 1.0, "Theme_Color": "Green"},
        {"KPI_ID": "KPI-04", "Metric_Title": "Average Order Value (AOV)", "Primary_Value": aov_val, "Unit": "INR ₹", "Subtext": "Net revenue / successful orders", "Progress_Pct": 1.0, "Theme_Color": "Blue"}
    ])
    ws1 = wb.create_sheet("01_Executive_KPIs")
    write_df_to_sheet(ws1, df_exec, "01. Executive KPI Summary Cards")

    # 2. Payment Performance KPIs Sheet
    df_pay_kpi = pd.DataFrame([
        {"KPI_ID": "PAY-01", "Metric_Title": "COD Net Revenue", "Primary_Value": cod_net_revenue, "Payment_Type": "COD", "Subtext": f"From {cod_delivered_count} delivered COD orders", "Theme_Color": "Orange"},
        {"KPI_ID": "PAY-02", "Metric_Title": "Avg COD Order Value", "Primary_Value": cod_aov, "Payment_Type": "COD", "Subtext": "COD net revenue / successful COD orders", "Theme_Color": "Orange"},
        {"KPI_ID": "PAY-03", "Metric_Title": "Prepaid Net Revenue", "Primary_Value": prepaid_net_revenue, "Payment_Type": "Prepaid", "Subtext": f"From {prepaid_delivered_count} delivered Prepaid orders", "Theme_Color": "Green"},
        {"KPI_ID": "PAY-04", "Metric_Title": "Avg Prepaid Order Value", "Primary_Value": prepaid_aov, "Payment_Type": "Prepaid", "Subtext": "Prepaid net revenue / successful prepaid orders", "Theme_Color": "Green"}
    ])
    ws2 = wb.create_sheet("02_Payment_KPIs")
    write_df_to_sheet(ws2, df_pay_kpi, "02. Payment Performance KPI Cards")

    # 3. Order Lifecycle Breakdown Sheet
    df_lifecycle = pd.DataFrame([
        {"Lifecycle_Stage": "Successful Orders", "Order_Count": successful_count, "Total_Amount_INR": successful_val, "Stage_Rate_Pct": successful_count/total_orders if total_orders > 0 else 0.0, "Description": "Delivered & self fulfilled", "Status_Color": "Green"},
        {"Lifecycle_Stage": "Denied Orders (Doorstep RTO)", "Order_Count": denied_count, "Total_Amount_INR": denied_val, "Stage_Rate_Pct": denied_count/total_orders if total_orders > 0 else 0.0, "Description": "Doorstep delivery denials", "Status_Color": "Red"},
        {"Lifecycle_Stage": "Returned Orders", "Order_Count": returned_count, "Total_Amount_INR": returned_val, "Stage_Rate_Pct": returned_count/total_orders if total_orders > 0 else 0.0, "Description": "Delivered then sent back", "Status_Color": "Orange"},
        {"Lifecycle_Stage": "Canceled Orders", "Order_Count": canceled_count, "Total_Amount_INR": canceled_val, "Stage_Rate_Pct": canceled_count/total_orders if total_orders > 0 else 0.0, "Description": "Order not accepted", "Status_Color": "Orange"}
    ])
    ws3 = wb.create_sheet("03_Lifecycle_Metrics")
    write_df_to_sheet(ws3, df_lifecycle, "03. Master Order Lifecycle Breakdown")

    # 4. In-Progress Logistics Pipeline Sheet
    df_pipeline = pd.DataFrame([
        {"Pipeline_Stage": "Total In-Progress", "Order_Count": in_progress_count, "Total_Amount_INR": in_progress_val, "Active_Rate_Pct": in_progress_count/total_orders if total_orders > 0 else 0.0, "Theme_Color": "Blue"},
        {"Pipeline_Stage": "In Transit", "Order_Count": len(in_transit_df), "Total_Amount_INR": in_transit_df['Total Price'].sum(), "Active_Rate_Pct": len(in_transit_df)/total_orders if total_orders > 0 else 0.0, "Theme_Color": "Blue"},
        {"Pipeline_Stage": "Pickup Scheduled", "Order_Count": len(pickup_df), "Total_Amount_INR": pickup_df['Total Price'].sum(), "Active_Rate_Pct": len(pickup_df)/total_orders if total_orders > 0 else 0.0, "Theme_Color": "Purple"},
        {"Pipeline_Stage": "Unfulfilled", "Order_Count": len(unfulfilled_df), "Total_Amount_INR": unfulfilled_df['Total Price'].sum(), "Active_Rate_Pct": len(unfulfilled_df)/total_orders if total_orders > 0 else 0.0, "Theme_Color": "Amber"}
    ])
    ws4 = wb.create_sheet("04_In_Progress_Pipeline")
    write_df_to_sheet(ws4, df_pipeline, "04. Active In-Progress Pipeline Grid")

    # 5. Product Category Summary Sheet
    cat_rows = []
    all_cats = df_master['Category'].str.split(',').explode().str.strip().unique()
    for cat in sorted(all_cats):
        if not cat or cat == 'nan':
            continue
        cat_df = df_master[df_master['Category'].str.contains(cat, na=False)]
        rev = cat_df['Total Price'].sum()
        cnt = len(cat_df)
        share = rev / total_revenue if total_revenue > 0 else 0.0
        
        deliv = len(cat_df[cat_df['Fulfillment Status'].str.contains("DELIVERED|FULFILLED", na=False)])
        trans = len(cat_df[cat_df['Fulfillment Status'].str.contains("TRANSIT|SHIPPED", na=False)])
        unful = len(cat_df[cat_df['Fulfillment Status'].str.contains("NEW ORDER|UNFULFILLED", na=False)])
        deni = len(cat_df[(cat_df['COD Denies (Yes/No)'] == "Yes") | (cat_df['Fulfillment Status'].str.contains("DENIED|RTO", na=False))])
        retu = len(cat_df[cat_df['Returned (True/False)'] == True])
        canc = len(cat_df[cat_df['Fulfillment Status'].str.contains("CANCELED", na=False)])
        
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
    write_df_to_sheet(ws5, df_cat, "05. Product Category Performance Breakdown")

    # 6. Payment Mode Performance Sheet
    pay_rows = []
    for mode in ["COD", "Prepaid (Razorpay)", "Partial COD"]:
        pm_df = df_master[df_master['Payment Method'] == mode]
        cnt = len(pm_df)
        if cnt == 0:
            continue
        rev = pm_df['Total Price'].sum()
        deliv = len(pm_df[pm_df['Fulfillment Status'].str.contains("DELIVERED|FULFILLED", na=False)])
        trans = len(pm_df[pm_df['Fulfillment Status'].str.contains("TRANSIT|SHIPPED", na=False)])
        pick = len(pm_df[pm_df['Fulfillment Status'].str.contains("PICKUP", na=False)])
        unful = len(pm_df[pm_df['Fulfillment Status'].str.contains("NEW ORDER|UNFULFILLED", na=False)])
        deni = len(pm_df[(pm_df['COD Denies (Yes/No)'] == "Yes") | (pm_df['Fulfillment Status'].str.contains("DENIED|RTO", na=False))])
        retu = len(pm_df[pm_df['Returned (True/False)'] == True])
        canc = len(pm_df[pm_df['Fulfillment Status'].str.contains("CANCELED", na=False)])
        
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
    write_df_to_sheet(ws6, df_pay_mode, "06. Mode of Payment Breakdown (COD vs Prepaid)")

    # 7. Monthly Sales Trends Sheet
    df_master['Date_Obj'] = pd.to_datetime(df_master['Date of Order'], errors='coerce')
    df_master['Month_Year'] = df_master['Date_Obj'].dt.strftime('%b %Y')
    
    monthly_rows = []
    for month, m_df in df_master.groupby('Month_Year', sort=False):
        monthly_rows.append({
            "Month_Year": month,
            "Total_Orders": len(m_df),
            "Gross_Revenue_INR": m_df['Total Price'].sum(),
            "Delivered_Orders": len(m_df[m_df['Fulfillment Status'].str.contains("DELIVERED|FULFILLED", na=False)]),
            "Denied_Orders": len(m_df[(m_df['COD Denies (Yes/No)'] == "Yes") | (m_df['Fulfillment Status'].str.contains("DENIED|RTO", na=False))]),
            "Returned_Orders": len(m_df[m_df['Returned (True/False)'] == True]),
            "Canceled_Orders": len(m_df[m_df['Fulfillment Status'].str.contains("CANCELED", na=False)])
        })
    df_monthly = pd.DataFrame(monthly_rows)
    ws7 = wb.create_sheet("07_Monthly_Sales_Trends")
    write_df_to_sheet(ws7, df_monthly, "07. Monthly Sales & Volume Trends")

    # 8. Master Orders Dataset Sheet
    df_master_clean = df_master.drop(columns=['Date_Obj', 'Month_Year'])
    ws8 = wb.create_sheet("Master_Orders_Dataset")
    write_df_to_sheet(ws8, df_master_clean, "08. Master Consolidated Orders Dataset")

    wb.save(output_path)
    print(f"Successfully generated 4see Magic Dashboard Master Excel: {output_path}")

if __name__ == "__main__":
    generate_4see_magic_excel()

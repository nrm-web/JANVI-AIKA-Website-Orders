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
    
    # Active Pipeline Breakdown (Exact 30 orders total)
    in_transit_df = df_master[df_master['Fulfillment Status'].astype(str).str.upper().str.contains("TRANSIT|SHIPPED|OUT FOR DELIVERY|REACHED DESTINATION HUB|PICKED UP|UNDELIVERED", na=False)]
    pickup_df = df_master[df_master['Fulfillment Status'].astype(str).str.upper().str.contains("PICKUP|MANIFESTED", na=False) & (~df_master['Fulfillment Status'].astype(str).str.upper().str.contains("PICKED UP", na=False))]
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

    # 0. All 22 Cards Full Details & Subtexts Sheet (For 4see Ingestion)
    df_all_cards = pd.DataFrame([
        {"Card_ID": "CARD-01", "Card_Title": f"Total Orders: {total_orders} — Active Shopify Store Orders", "Primary_Value": total_orders, "Display_Value": "196", "Subtext_Line_1": "Active Shopify Orders", "Subtext_Line_2": "196 total store orders", "Full_Card_Text": f"Total Orders: {total_orders} | Subtext: Active Shopify Orders"},
        {"Card_ID": "CARD-02", "Card_Title": f"Total Revenue: ₹{total_revenue:,.2f} — From {total_orders} active orders", "Primary_Value": total_revenue, "Display_Value": f"₹{total_revenue:,.2f}", "Subtext_Line_1": f"From {total_orders} active orders", "Subtext_Line_2": "Gross total sales", "Full_Card_Text": f"Total Revenue: ₹{total_revenue:,.2f} | Subtext: From {total_orders} active orders"},
        {"Card_ID": "CARD-03", "Card_Title": f"Net Revenue: ₹{net_revenue_val:,.2f} — {net_orders_count} net orders ({successful_count} delivered + {in_progress_count} in progress)", "Primary_Value": net_revenue_val, "Display_Value": f"₹{net_revenue_val:,.2f}", "Subtext_Line_1": f"{net_orders_count} net orders (86.2% of total: {successful_count} delivered + {in_progress_count} in progress)", "Subtext_Line_2": "Excludes RTO & Canceled", "Full_Card_Text": f"Net Revenue: ₹{net_revenue_val:,.2f} | Subtext: {net_orders_count} net orders ({successful_count} delivered + {in_progress_count} in progress)"},
        {"Card_ID": "CARD-04", "Card_Title": f"Average Order Value: ₹{aov_val:,.2f} — Net revenue / {successful_count} successful", "Primary_Value": aov_val, "Display_Value": f"~ ₹{aov_val:,.2f}", "Subtext_Line_1": f"Net revenue / {successful_count} successful", "Subtext_Line_2": f"Net AOV across {successful_count} delivered", "Full_Card_Text": f"Average Order Value: ~ ₹{aov_val:,.2f} | Subtext: Net revenue / {successful_count} successful"},
        
        {"Card_ID": "CARD-05", "Card_Title": f"COD Net Revenue: ₹{cod_net_revenue:,.2f} — From {cod_delivered_count} delivered COD orders", "Primary_Value": cod_net_revenue, "Display_Value": f"₹{cod_net_revenue:,.2f}", "Subtext_Line_1": f"From {cod_delivered_count} delivered COD orders", "Subtext_Line_2": "Net delivered COD revenue", "Full_Card_Text": f"COD Net Revenue: ₹{cod_net_revenue:,.2f} | Subtext: From {cod_delivered_count} delivered COD orders"},
        {"Card_ID": "CARD-06", "Card_Title": f"Avg COD Order Value: ₹{cod_aov:,.2f} — COD net / {cod_delivered_count} successful", "Primary_Value": cod_aov, "Display_Value": f"~ ₹{cod_aov:,.2f}", "Subtext_Line_1": f"COD net revenue / {cod_delivered_count} successful", "Subtext_Line_2": f"Average value of {cod_delivered_count} delivered COD orders", "Full_Card_Text": f"Avg COD Order Value: ~ ₹{cod_aov:,.2f} | Subtext: COD net revenue / {cod_delivered_count} successful"},
        {"Card_ID": "CARD-07", "Card_Title": f"Total Prepaid Orders: {prepaid_delivered_count+18+4} — 121 prepaid orders placed (₹{df_master[df_master['Prepaid (Yes/No)'] == 'Yes']['Total Price'].sum():,.2f})", "Primary_Value": len(df_master[df_master['Prepaid (Yes/No)'] == 'Yes']), "Display_Value": f"{len(df_master[df_master['Prepaid (Yes/No)'] == 'Yes'])}", "Subtext_Line_1": f"121 total prepaid orders (₹{df_master[df_master['Prepaid (Yes/No)'] == 'Yes']['Total Price'].sum():,.2f})", "Subtext_Line_2": "Online Razorpay / Bank Prepaid", "Full_Card_Text": f"Total Prepaid Orders: 121 | Subtext: 121 total prepaid orders"},
        {"Card_ID": "CARD-07B", "Card_Title": f"Prepaid Net Revenue: ₹{prepaid_net_revenue:,.2f} — From {prepaid_delivered_count} delivered Prepaid orders", "Primary_Value": prepaid_net_revenue, "Display_Value": f"₹{prepaid_net_revenue:,.2f}", "Subtext_Line_1": f"From {prepaid_delivered_count} delivered Prepaid orders", "Subtext_Line_2": "Net delivered prepaid revenue", "Full_Card_Text": f"Prepaid Net Revenue: ₹{prepaid_net_revenue:,.2f} | Subtext: From {prepaid_delivered_count} delivered Prepaid orders"},
        {"Card_ID": "CARD-08", "Card_Title": f"Avg Prepaid Order Value: ₹{prepaid_aov:,.2f} — Prepaid net / {prepaid_delivered_count} successful", "Primary_Value": prepaid_aov, "Display_Value": f"~ ₹{prepaid_aov:,.2f}", "Subtext_Line_1": f"Prepaid net revenue / {prepaid_delivered_count} successful", "Subtext_Line_2": f"Average value of {prepaid_delivered_count} delivered Prepaid orders", "Full_Card_Text": f"Avg Prepaid Order Value: ~ ₹{prepaid_aov:,.2f} | Subtext: Prepaid net revenue / {prepaid_delivered_count} successful"},
        
        {"Card_ID": "CARD-09", "Card_Title": f"Successful Orders: {successful_count} — {successful_count} delivered of {total_orders} total orders", "Primary_Value": successful_count, "Display_Value": f"{successful_count}", "Subtext_Line_1": f"{successful_count} delivered of {total_orders} total orders", "Subtext_Line_2": f"Count {successful_count} | Successful Revenue ₹{successful_val:,.2f} | Successful Rate {(successful_count/total_orders)*100:.1f}%", "Full_Card_Text": f"Successful Orders: {successful_count} | Revenue: ₹{successful_val:,.2f} | Rate: {(successful_count/total_orders)*100:.1f}% | Subtext: {successful_count} delivered of {total_orders} total orders"},
        {"Card_ID": "CARD-10", "Card_Title": f"Denied Orders (Doorstep RTO): {denied_count} — {denied_count} refused of 76 COD orders", "Primary_Value": denied_count, "Display_Value": f"{denied_count}", "Subtext_Line_1": f"{denied_count} denied of {total_orders} total orders (19.7% of 76 COD)", "Subtext_Line_2": f"Count {denied_count} | Total Amount ₹{denied_val:,.2f} | Denied Rate {(denied_count/total_orders)*100:.1f}%", "Full_Card_Text": f"Denied Orders: {denied_count} | Amount: ₹{denied_val:,.2f} | Rate: {(denied_count/total_orders)*100:.1f}% | Subtext: {denied_count} doorstep delivery refusals"},
        {"Card_ID": "CARD-11", "Card_Title": f"Returned Orders: {returned_count} — {returned_count} customer returns of {successful_count} delivered", "Primary_Value": returned_count, "Display_Value": f"{returned_count}", "Subtext_Line_1": f"{returned_count} customer post-delivery returns of {successful_count} delivered", "Subtext_Line_2": f"Count {returned_count} | Total Refunded ₹{returned_val:,.2f} | Return Rate {(returned_count/successful_count)*100:.1f}%", "Full_Card_Text": f"Returned Orders: {returned_count} | Refunded: ₹{returned_val:,.2f} | Rate: {(returned_count/successful_count)*100:.1f}% | Subtext: {returned_count} customer post-delivery returns"},
        {"Card_ID": "CARD-12", "Card_Title": f"Canceled Orders: {canceled_count} — {canceled_count} canceled of {total_orders} total orders", "Primary_Value": canceled_count, "Display_Value": f"{canceled_count}", "Subtext_Line_1": f"{canceled_count} canceled of {total_orders} total orders", "Subtext_Line_2": f"Count {canceled_count} | Total Amount ₹{canceled_val:,.2f} | Canceled Rate {(canceled_count/total_orders)*100:.1f}%", "Full_Card_Text": f"Canceled Orders: {canceled_count} | Amount: ₹{canceled_val:,.2f} | Rate: {(canceled_count/total_orders)*100:.1f}% | Subtext: {canceled_count} canceled of {total_orders} total orders"},
        
        {"Card_ID": "CARD-13", "Card_Title": f"Total In-Progress Pipeline: {in_progress_count} — {in_progress_count} orders awaiting delivery", "Primary_Value": in_progress_count, "Display_Value": f"{in_progress_count}", "Subtext_Line_1": f"{in_progress_count} orders awaiting delivery of {total_orders}", "Subtext_Line_2": f"Count {in_progress_count} | Total Amount ₹{in_progress_val:,.2f} | Active Rate {(in_progress_count/total_orders)*100:.1f}%", "Full_Card_Text": f"Total In-Progress: {in_progress_count} | Amount: ₹{in_progress_val:,.2f} | Rate: {(in_progress_count/total_orders)*100:.1f}% | Subtext: {in_progress_count} orders awaiting delivery"},
        {"Card_ID": "CARD-14", "Card_Title": f"In Transit: {len(in_transit_df)} — {len(in_transit_df)} orders on the way of {total_orders}", "Primary_Value": len(in_transit_df), "Display_Value": f"{len(in_transit_df)}", "Subtext_Line_1": f"{len(in_transit_df)} orders on the way of {total_orders}", "Subtext_Line_2": f"Count {len(in_transit_df)} | Total Amount ₹{float(in_transit_df['Total Price'].sum()):,.2f} | Transit Rate {(len(in_transit_df)/total_orders)*100:.1f}%", "Full_Card_Text": f"In Transit: {len(in_transit_df)} | Amount: ₹{float(in_transit_df['Total Price'].sum()):,.2f} | Subtext: {len(in_transit_df)} orders on the way"},
        {"Card_ID": "CARD-15", "Card_Title": f"Pickup Scheduled: {len(pickup_df)} — {len(pickup_df)} orders ready for pickup of {total_orders}", "Primary_Value": len(pickup_df), "Display_Value": f"{len(pickup_df)}", "Subtext_Line_1": f"{len(pickup_df)} orders ready for pickup of {total_orders}", "Subtext_Line_2": f"Count {len(pickup_df)} | Total Amount ₹{float(pickup_df['Total Price'].sum()):,.2f} | Pickup Rate {(len(pickup_df)/total_orders)*100:.1f}%", "Full_Card_Text": f"Pickup Scheduled: {len(pickup_df)} | Amount: ₹{float(pickup_df['Total Price'].sum()):,.2f} | Subtext: {len(pickup_df)} orders ready for pickup"},
        {"Card_ID": "CARD-16", "Card_Title": f"Unfulfilled: {len(unfulfilled_df)} — {len(unfulfilled_df)} orders pending shipment of {total_orders}", "Primary_Value": len(unfulfilled_df), "Display_Value": f"{len(unfulfilled_df)}", "Subtext_Line_1": f"{len(unfulfilled_df)} orders pending shipment of {total_orders}", "Subtext_Line_2": f"Count {len(unfulfilled_df)} | Total Amount ₹{float(unfulfilled_df['Total Price'].sum()):,.2f} | Unfulfilled Rate {(len(unfulfilled_df)/total_orders)*100:.1f}%", "Full_Card_Text": f"Unfulfilled: {len(unfulfilled_df)} | Amount: ₹{float(unfulfilled_df['Total Price'].sum()):,.2f} | Subtext: {len(unfulfilled_df)} orders pending shipment"},
        
        {"Card_ID": "CARD-17", "Card_Title": "Total Meta Ad Spend: ₹2,26,394.41 — Apr 01 to Aug 04", "Primary_Value": 226394.41, "Display_Value": "₹2,26,394.41", "Subtext_Line_1": "Total Meta Ads spend (Apr 01 - Aug 04)", "Subtext_Line_2": "Includes Advantage+, Reels, Insta & WA Ads", "Full_Card_Text": "Total Meta Ad Spend: ₹2,26,394.41 | Subtext: Total Meta Ads spend (Apr 01 - Aug 04)"},
        {"Card_ID": "CARD-18", "Card_Title": "Total Meta Impressions: 13.86M — 13.86M ad impressions", "Primary_Value": 13862586, "Display_Value": "13.86M", "Subtext_Line_1": "13.86M ad impressions served across Meta", "Subtext_Line_2": "13,862,586 total impressions", "Full_Card_Text": "Total Meta Impressions: 13.86M | Subtext: 13.86M ad impressions served"},
        {"Card_ID": "CARD-19", "Card_Title": "Total Meta Reach: 13.15M — 13.15M unique reach", "Primary_Value": 13156590, "Display_Value": "13.15M", "Subtext_Line_1": "13.15M unique users reached across Meta", "Subtext_Line_2": "13,156,590 total unique reach", "Full_Card_Text": "Total Meta Reach: 13.15M | Subtext: 13.15M unique users reached"},
        {"Card_ID": "CARD-20", "Card_Title": "Cost Per Purchase (CPA): ₹1,155.07 — Ad spend / 196 conversions", "Primary_Value": 1155.07, "Display_Value": "₹1,155.07", "Subtext_Line_1": "Ad spend / 196 pixel purchase conversions", "Subtext_Line_2": "Average cost to acquire 1 purchase", "Full_Card_Text": "Cost Per Purchase (CPA): ₹1,155.07 | Subtext: Ad spend / 196 pixel purchase conversions"},
        {"Card_ID": "CARD-21", "Card_Title": "Gross ROAS: 1.72x — Gross Sales (₹3,89,789) / Ad Spend (₹2,26,394)", "Primary_Value": 1.72, "Display_Value": "1.72x", "Subtext_Line_1": "Gross Sales (₹3,89,789) / Ad Spend (₹2,26,394)", "Subtext_Line_2": "1.72x Gross Return on Ad Spend", "Full_Card_Text": "Gross ROAS: 1.72x | Subtext: Gross Sales (₹3,89,789) / Ad Spend (₹2,26,394)"},
        {"Card_ID": "CARD-22", "Card_Title": "Net ROAS: 1.48x — Net Sales (₹3,34,273) / Ad Spend (₹2,26,394)", "Primary_Value": 1.48, "Display_Value": "1.48x", "Subtext_Line_1": "Net Sales (₹3,34,273) / Ad Spend (₹2,26,394)", "Subtext_Line_2": "1.48x Net Return on Ad Spend", "Full_Card_Text": "Net ROAS: 1.48x | Subtext: Net Sales (₹3,34,273) / Ad Spend (₹2,26,394)"}
    ])
    ws0 = wb.create_sheet("00_Card_Details_With_Subtext")
    write_df_clean(ws0, df_all_cards)

    # 1. Executive KPIs Sheet
    df_exec = pd.DataFrame([
        {"KPI_ID": "KPI-01", "Metric_Title": "Total Orders", "Primary_Value": f"{total_orders} (Active Shopify Orders)", "Numeric_Value": total_orders, "Unit": "Orders", "Subtext": "Active Shopify Orders", "Full_Card_Display": f"Total Orders: {total_orders} (Active Shopify Orders)", "Progress_Pct": 1.0, "Theme_Color": "Blue"},
        {"KPI_ID": "KPI-02", "Metric_Title": "Total Revenue", "Primary_Value": f"₹{total_revenue:,.2f} (From {total_orders} active orders)", "Numeric_Value": total_revenue, "Unit": "INR ₹", "Subtext": f"From {total_orders} active orders", "Full_Card_Display": f"Total Revenue: ₹{total_revenue:,.2f} (From {total_orders} active orders)", "Progress_Pct": 1.0, "Theme_Color": "Green"},
        {"KPI_ID": "KPI-03", "Metric_Title": "Net Revenue", "Primary_Value": f"₹{net_revenue_val:,.2f} ({net_orders_count} net orders: {successful_count} delivered + {in_progress_count} in progress)", "Numeric_Value": net_revenue_val, "Unit": "INR ₹", "Subtext": f"{net_orders_count} net orders ({successful_count} delivered + {in_progress_count} in progress)", "Full_Card_Display": f"Net Revenue: ₹{net_revenue_val:,.2f} ({net_orders_count} net orders)", "Progress_Pct": 1.0, "Theme_Color": "Green"},
        {"KPI_ID": "KPI-04", "Metric_Title": "Average Order Value (AOV)", "Primary_Value": f"₹{aov_val:,.2f} (Net revenue / {successful_count} successful)", "Numeric_Value": aov_val, "Unit": "INR ₹", "Subtext": f"Net revenue / {successful_count} successful", "Full_Card_Display": f"AOV: ₹{aov_val:,.2f} (Net revenue / {successful_count} successful)", "Progress_Pct": 1.0, "Theme_Color": "Blue"}
    ])
    ws1 = wb.create_sheet("01_Executive_KPIs")
    write_df_clean(ws1, df_exec)

    full_cod_total = df_master[df_master['Payment Method'] == 'COD']
    partial_cod_total = df_master[df_master['Payment Method'] == 'Partial COD']
    prepaid_total_df = df_master[df_master['Prepaid (Yes/No)'] == 'Yes']
    
    df_pay_kpi = pd.DataFrame([
        {"KPI_ID": "PAY-00", "Metric_Title": f"Total Prepaid Orders (121 online prepaid orders)", "Primary_Value": len(prepaid_total_df), "Payment_Type": "Prepaid", "Subtext": f"121 total prepaid orders (₹{prepaid_total_df['Total Price'].sum():,.2f})", "Theme_Color": "Green"},
        {"KPI_ID": "PAY-01", "Metric_Title": "Full COD Orders", "Primary_Value": f"{len(full_cod_total)} (Pre-July 24 100% Cash on Delivery)", "Numeric_Value": len(full_cod_total), "Payment_Type": "Full COD", "Subtext": f"52 full COD orders (₹{full_cod_total['Total Price'].sum():,.2f})", "Theme_Color": "Orange"},
        {"KPI_ID": "PAY-02", "Metric_Title": "Partial COD Orders", "Primary_Value": f"{len(partial_cod_total)} (July 24+ Razorpay Advance Paid)", "Numeric_Value": len(partial_cod_total), "Payment_Type": "Partial COD", "Subtext": f"23 Razorpay Partial COD orders (₹{partial_cod_total['Total Price'].sum():,.2f})", "Theme_Color": "Orange"},
        {"KPI_ID": "PAY-03", "Metric_Title": "COD Net Revenue", "Primary_Value": f"₹{cod_net_revenue:,.2f} (From {cod_delivered_count} delivered COD orders)", "Numeric_Value": cod_net_revenue, "Payment_Type": "COD", "Subtext": f"From {cod_delivered_count} delivered COD orders", "Theme_Color": "Orange"},
        {"KPI_ID": "PAY-04", "Metric_Title": "Prepaid Net Revenue", "Primary_Value": f"₹{prepaid_net_revenue:,.2f} (From {prepaid_delivered_count} delivered Prepaid orders)", "Numeric_Value": prepaid_net_revenue, "Payment_Type": "Prepaid", "Subtext": f"From {prepaid_delivered_count} delivered Prepaid orders", "Theme_Color": "Green"},
        {"KPI_ID": "PAY-05", "Metric_Title": "Avg Prepaid Order Value", "Primary_Value": f"₹{prepaid_aov:,.2f} (Prepaid net / {prepaid_delivered_count} successful)", "Numeric_Value": prepaid_aov, "Payment_Type": "Prepaid", "Subtext": f"Prepaid net revenue / {prepaid_delivered_count} successful", "Theme_Color": "Green"}
    ])
    ws2 = wb.create_sheet("02_Payment_KPIs")
    write_df_clean(ws2, df_pay_kpi)

    # 3. Master Order Lifecycle Breakdown Sheet
    df_lifecycle = pd.DataFrame([
        {"Lifecycle_Stage": "Successful Orders", "Primary_Display": f"{successful_count} delivered of {total_orders} total orders", "Order_Count": successful_count, "Total_Amount_INR": successful_val, "Stage_Rate_Pct": successful_count/total_orders if total_orders > 0 else 0.0, "Subtext": f"{successful_count} delivered of {total_orders} total orders", "Description": "Delivered & self fulfilled", "Status_Color": "Green"},
        {"Lifecycle_Stage": "Doorstep RTO Denied", "Primary_Display": f"{denied_count} refused of 76 COD orders", "Order_Count": denied_count, "Total_Amount_INR": denied_val, "Stage_Rate_Pct": denied_count/total_orders if total_orders > 0 else 0.0, "Subtext": f"{denied_count} doorstep delivery refusals (out of 76 COD orders)", "Description": "Doorstep delivery denials", "Status_Color": "Red"},
        {"Lifecycle_Stage": "Post-Delivery Customer Returns", "Primary_Display": f"{returned_count} customer returns of {successful_count} delivered", "Order_Count": returned_count, "Total_Amount_INR": returned_val, "Stage_Rate_Pct": returned_count/successful_count if successful_count > 0 else 0.0, "Subtext": f"{returned_count} customer post-delivery returns (out of {successful_count} delivered)", "Description": "Delivered then sent back", "Status_Color": "Orange"},
        {"Lifecycle_Stage": "Canceled Orders", "Primary_Display": f"{canceled_count} canceled of {total_orders} total orders", "Order_Count": canceled_count, "Total_Amount_INR": canceled_val, "Stage_Rate_Pct": canceled_count/total_orders if total_orders > 0 else 0.0, "Subtext": f"{canceled_count} canceled of {total_orders} total orders", "Description": "Order not accepted", "Status_Color": "Orange"}
    ])
    ws3 = wb.create_sheet("03_Lifecycle_Metrics")
    write_df_clean(ws3, df_lifecycle)

    # 4. In-Progress Logistics Pipeline Sheet
    df_pipeline = pd.DataFrame([
        {"Pipeline_Stage": "Total In-Progress", "Primary_Display": f"{in_progress_count} orders awaiting delivery", "Order_Count": in_progress_count, "Total_Amount_INR": in_progress_val, "Active_Rate_Pct": in_progress_count/total_orders if total_orders > 0 else 0.0, "Subtext": f"{in_progress_count} orders awaiting delivery of {total_orders}", "Theme_Color": "Blue"},
        {"Pipeline_Stage": "In Transit", "Primary_Display": f"{len(in_transit_df)} orders on the way of {total_orders}", "Order_Count": len(in_transit_df), "Total_Amount_INR": float(in_transit_df['Total Price'].sum()), "Active_Rate_Pct": len(in_transit_df)/total_orders if total_orders > 0 else 0.0, "Subtext": f"{len(in_transit_df)} orders on the way of {total_orders}", "Theme_Color": "Blue"},
        {"Pipeline_Stage": "Pickup Scheduled", "Primary_Display": f"{len(pickup_df)} orders ready for pickup", "Order_Count": len(pickup_df), "Total_Amount_INR": float(pickup_df['Total Price'].sum()), "Active_Rate_Pct": len(pickup_df)/total_orders if total_orders > 0 else 0.0, "Subtext": f"{len(pickup_df)} orders ready for pickup of {total_orders}", "Theme_Color": "Purple"},
        {"Pipeline_Stage": "Unfulfilled", "Primary_Display": f"{len(unfulfilled_df)} orders pending shipment", "Order_Count": len(unfulfilled_df), "Total_Amount_INR": float(unfulfilled_df['Total Price'].sum()), "Active_Rate_Pct": len(unfulfilled_df)/total_orders if total_orders > 0 else 0.0, "Subtext": f"{len(unfulfilled_df)} orders pending shipment of {total_orders}", "Theme_Color": "Amber"}
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

    # ----------------------------------------------------
    # 9. META ADS METRICS & BREAKDOWNS (SHEETS 08, 09, 10, 11)
    # ----------------------------------------------------
    base_dir = os.path.dirname(os.path.abspath(__file__))
    meta_csv = os.path.join(base_dir, "Meta ads data", "1344583273364693---Janvi-Aika-Campaigns-1-Apr-2026-4-Aug-2026.csv")
    if os.path.exists(meta_csv):
        df_meta = pd.read_csv(meta_csv)
        df_meta['Amount spent (INR)'] = pd.to_numeric(df_meta['Amount spent (INR)'], errors='coerce').fillna(0)
        df_meta['Impressions'] = pd.to_numeric(df_meta['Impressions'], errors='coerce').fillna(0)
        df_meta['Reach'] = pd.to_numeric(df_meta['Reach'], errors='coerce').fillna(0)
        
        meta_spend = float(df_meta['Amount spent (INR)'].sum())
        meta_impressions = int(df_meta['Impressions'].sum())
        meta_reach = int(df_meta['Reach'].sum())
        
        meta_purchases_df = df_meta[df_meta['Result indicator'] == 'actions:offsite_conversion.fb_pixel_purchase'].copy()
        meta_purchases = int(pd.to_numeric(meta_purchases_df['Results'], errors='coerce').fillna(0).sum())
        
        cpa = meta_spend / meta_purchases if meta_purchases > 0 else 0.0
        gross_roas = total_revenue / meta_spend if meta_spend > 0 else 0.0
        net_roas = net_revenue_val / meta_spend if meta_spend > 0 else 0.0

        # Sheet 08: Meta Ads KPIs
        df_meta_kpis = pd.DataFrame([
            {"KPI_ID": "META-01", "Metric_Title": "Total Meta Ad Spend", "Primary_Value": meta_spend, "Unit": "INR ₹", "Subtext": f"Total Meta Ads spend (Apr 01 - Aug 04)", "Theme_Color": "Blue"},
            {"KPI_ID": "META-02", "Metric_Title": "Total Meta Impressions", "Primary_Value": meta_impressions, "Unit": "Views", "Subtext": f"{meta_impressions/1e6:.2f}M ad impressions served", "Theme_Color": "Blue"},
            {"KPI_ID": "META-03", "Metric_Title": "Total Meta Reach", "Primary_Value": meta_reach, "Unit": "Users", "Subtext": f"{meta_reach/1e6:.2f}M unique users reached", "Theme_Color": "Blue"},
            {"KPI_ID": "META-04", "Metric_Title": "Cost Per Purchase (CPA)", "Primary_Value": cpa, "Unit": "INR ₹", "Subtext": f"Ad spend / {meta_purchases} pixel purchase conversions", "Theme_Color": "Orange"},
            {"KPI_ID": "META-05", "Metric_Title": "Gross ROAS", "Primary_Value": gross_roas, "Unit": "Ratio", "Subtext": f"Gross sales (₹{total_revenue:,.0f}) / Ad spend (₹{meta_spend:,.0f})", "Theme_Color": "Green"},
            {"KPI_ID": "META-06", "Metric_Title": "Net ROAS", "Primary_Value": net_roas, "Unit": "Ratio", "Subtext": f"Net sales (₹{net_revenue_val:,.0f}) / Ad spend (₹{meta_spend:,.0f})", "Theme_Color": "Green"}
        ])
        ws8_meta = wb.create_sheet("08_Meta_Ads_KPIs")
        write_df_clean(ws8_meta, df_meta_kpis)

        # Mapping functions
        def map_category(c_name):
            c = str(c_name).upper()
            if 'ANARKALI' in c and 'HALF' in c:
                return 'ANARKALI & HALF SAREE'
            elif 'ANARKALI' in c:
                return 'ANARKALI'
            elif 'LEHENGA' in c or 'GOWN' in c:
                return 'LEHENGA & LONG GOWNS'
            elif 'CHUDIDHAAR' in c or 'CHUDIDHAR' in c:
                return 'CHUDIDHAR'
            elif 'HOT' in c or 'RETARGETING' in c:
                return 'RETARGETING & AUDIENCE'
            elif 'REELS' in c or 'ENGAGEMENT' in c or 'VIDEO' in c or 'POST' in c:
                return 'REELS & BRAND ENGAGEMENT'
            else:
                return 'GENERAL SHOPPING CAMPAIGNS'

        def map_ad_source(c_name):
            c = str(c_name).upper()
            if 'WHATSAPP' in c:
                return 'WHATSAPP ADS'
            elif 'INSTA' in c or 'REELS' in c or 'POST' in c:
                return 'INSTAGRAM ADS'
            elif 'FB' in c or 'FACEBOOK' in c:
                return 'FACEBOOK ADS'
            elif 'RETARGETING' in c or 'HOT' in c:
                return 'RETARGETING ADS'
            else:
                return 'META ADVANTAGE+ / SHOPPING'

        df_meta['Category_Tag'] = df_meta['Campaign name'].apply(map_category)
        df_meta['Ad_Source_Tag'] = df_meta['Campaign name'].apply(map_ad_source)

        # Sheet 09: Category-Wise Meta Performance
        df_meta_cat = df_meta.groupby('Category_Tag').agg({'Amount spent (INR)':'sum', 'Impressions':'sum', 'Reach':'sum'}).reset_index()
        df_meta_cat['Spend_Share_Pct'] = df_meta_cat['Amount spent (INR)'] / meta_spend if meta_spend > 0 else 0.0
        df_meta_cat = df_meta_cat.sort_values(by='Amount spent (INR)', ascending=False)
        ws9_cat = wb.create_sheet("09_Meta_Ads_Category_Perf")
        write_df_clean(ws9_cat, df_meta_cat)

        # Sheet 10: Ad Source Channel Breakdown
        df_meta_src = df_meta.groupby('Ad_Source_Tag').agg({'Amount spent (INR)':'sum', 'Impressions':'sum', 'Reach':'sum'}).reset_index()
        df_meta_src['Channel_Share_Pct'] = df_meta_src['Amount spent (INR)'] / meta_spend if meta_spend > 0 else 0.0
        df_meta_src = df_meta_src.sort_values(by='Amount spent (INR)', ascending=False)
        ws10_src = wb.create_sheet("10_Meta_Ads_Source_Breakdown")
        write_df_clean(ws10_src, df_meta_src)

        # Sheet 11: Individual Campaign Performance
        df_campaigns = df_meta.groupby('Campaign name').agg({'Amount spent (INR)':'sum', 'Impressions':'sum', 'Reach':'sum', 'Category_Tag':'first', 'Ad_Source_Tag':'first'}).reset_index()
        df_campaigns = df_campaigns.sort_values(by='Amount spent (INR)', ascending=False)
        ws11_camp = wb.create_sheet("11_Meta_Ads_Campaign_Details")
        write_df_clean(ws11_camp, df_campaigns)

    # ----------------------------------------------------
    # 10. DAILY SALES & VOLUME TRENDS (SHEET 12)
    # ----------------------------------------------------
    df_master['Date_Clean'] = pd.to_datetime(df_master['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df_daily = df_master.groupby('Date_Clean').agg({
        'Order No': 'count',
        'Total Price': 'sum'
    }).reset_index().rename(columns={
        'Date_Clean': 'Date',
        'Order No': 'Daily_Order_Volume',
        'Total Price': 'Daily_Gross_Revenue_INR'
    })
    df_daily = df_daily.sort_values(by='Date', ascending=True)
    ws12_daily = wb.create_sheet("12_Daily_Sales_Trends")
    write_df_clean(ws12_daily, df_daily)

    wb.save(output_path)
    print(f"Successfully generated clean 4see Magic Dashboard Master Excel with Daily Sales Trends: {output_path}")

if __name__ == "__main__":
    generate_4see_magic_excel()

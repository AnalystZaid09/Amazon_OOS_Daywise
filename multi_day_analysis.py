import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from datetime import datetime

# Page configuration
st.set_page_config(page_title="Multi-Day Business Report Tool", layout="wide")

st.markdown("""
<style>
    .main {
        background-color: #f5f7f9;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #007bff;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Multi-Day Business Report Tool")
st.markdown("Consolidate business reports for 7, 15, 30, 45, 60, 75, and 90-day intervals.")

# Sidebar for file uploads
with st.sidebar:
    st.header("📁 Upload Reports")
    
    intervals = [7, 15, 30, 45, 60, 75, 90]
    br_files = {}
    for day in intervals:
        br_files[day] = st.file_uploader(f"Business Report - {day} Days", type=["csv", "xlsx"], key=f"br_{day}")
    
    st.divider()
    pm_file = st.file_uploader("Purchase Master (Excel)", type=["xlsx", "xls"])
    inv_file = st.file_uploader("Inventory Report (CSV)", type=["csv"])
    res_file = st.file_uploader("Reserved Inventory (CSV)", type=["csv"], key="res_file")
    list_file = st.file_uploader("Listing Report (CSV/Excel)", type=["csv", "xlsx"], key="list_file")
    
    st.divider()
    st.markdown("### 📊 DOC Color Legend")
    st.markdown("""
    🔴 **Red (0-7)**: Critical  
    🟠 **Orange (7-15)**: Low  
    🟢 **Green (15-30)**: Optimal  
    🟡 **Yellow (30-45)**: Monitor  
    🔵 **Sky Blue (45-60)**: High  
    🟤 **Brown (60-90)**: Excess  
    ⬛ **Black (>90)**: Overstocked
    """)
    
    process_btn = st.button("🚀 Process Multi-Day Reports")

def clean_numeric_col(df, col):
    if col in df.columns:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("₹", "")
            .str.replace(",", "")
            .str.replace(" ", "")
            .str.replace("%", "")
            .fillna("0")
        )
        return pd.to_numeric(df[col], errors="coerce").fillna(0)
    return pd.Series(0, index=df.index)

def apply_doc_styling(df):
    """
    Apply background color to DOC columns for Streamlit display.
    """
    def style_doc(val):
        try:
            val = float(val)
            if val <= 7:
                return "background-color: #FF0000; color: white" # Red
            elif val <= 15:
                return "background-color: #FFA500; color: black" # Orange
            elif val <= 30:
                return "background-color: #008000; color: white" # Green
            elif val <= 45:
                return "background-color: #FFFF00; color: black" # Yellow
            elif val <= 60:
                return "background-color: #87CEEB; color: black" # Sky Blue
            elif val <= 90:
                return "background-color: #A52A2A; color: white" # Brown
            else:
                return "background-color: #000000; color: white" # Black
        except:
            return ""

    doc_cols = [c for c in df.columns if "DOC" in c]
    if not doc_cols:
        return df
        
    return df.style.map(style_doc, subset=doc_cols)

def create_stock_pivot(df):
    if df.empty: return pd.DataFrame()
    df_copy = df.copy()
    
    # Map Max columns to base names for pivot if they exist
    rename_map = {}
    if "DOC (Max)" in df_copy.columns: 
        rename_map["DOC (Max)"] = "DOC"
    elif "DOC" in df_copy.columns:
        pass # Already named correctly
    else:
        # Check if there are any other DOC columns we can use as a fallback
        any_doc = next((c for c in df_copy.columns if "DOC" in c), None)
        if any_doc: rename_map[any_doc] = "DOC"

    if "Max DRR" in df_copy.columns: 
        rename_map["Max DRR"] = "DRR"
    elif "DRR" in df_copy.columns:
        pass
    else:
        any_drr = next((c for c in df_copy.columns if "DRR" in c), None)
        if any_drr: rename_map[any_drr] = "DRR"

    if rename_map:
        df_copy = df_copy.rename(columns=rename_map)
        
    # Identify which of our target values actually exist now
    target_values = ["DOC", "DRR", "CP"]
    available_values = [v for v in target_values if v in df_copy.columns]
    
    if not available_values:
        return pd.DataFrame()

    for col in available_values:
        df_copy[col] = pd.to_numeric(df_copy[col], errors="coerce").fillna(0)
    
    # Ensure index columns exist
    index_cols = [c for c in ["Brand", "SKU"] if c in df_copy.columns]
    if not index_cols:
        return pd.DataFrame()

    pivot = pd.pivot_table(
        df_copy,
        index=index_cols,
        values=available_values,
        aggfunc="sum",
        margins=True,
        margins_name="Grand Total"
    ).reset_index()
    
    # Precise rename for final display
    pivot_rename = {v: f"Sum of {v}" for v in available_values}
    pivot.rename(columns=pivot_rename, inplace=True)
    return pivot

def process_br(file, days):
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    
    # Cleaning
    # cl.py uses Total Order Items + B2B as per snippet, but also has Units Ordered
    df["Units Ordered"] = clean_numeric_col(df, "Units Ordered")
    df["Units Ordered - B2B"] = clean_numeric_col(df, "Units Ordered - B2B")
    df["Total Order Items"] = clean_numeric_col(df, "Total Order Items")
    df["Total Order Items - B2B"] = clean_numeric_col(df, "Total Order Items - B2B")
    df["Page Views - Total"] = clean_numeric_col(df, "Page Views - Total")
    df["Page Views - Total - B2B"] = clean_numeric_col(df, "Page Views - Total - B2B")
    df["Sessions - Total"] = clean_numeric_col(df, "Sessions - Total")
    df["Buy Box Percentage"] = clean_numeric_col(df, "Buy Box Percentage")
    df["Unit Session Percentage"] = clean_numeric_col(df, "Unit Session Percentage")
    
    # cl.py logic for Sales Qty
    df["Period Sales Qty"] = df["Total Order Items"] + df["Total Order Items - B2B"]
    if df["Period Sales Qty"].sum() == 0:
        # Fallback to Units Ordered if Total Order Items is missing
        df["Period Sales Qty"] = df["Units Ordered"] + df["Units Ordered - B2B"]

    # Find SKU column in BR
    possible_sku_cols = ["SKU", "sku", "Seller SKU"]
    sku_col_br = next((c for c in possible_sku_cols if c in df.columns), None)
    
    if not sku_col_br:
        st.error(f"Could not find a SKU column in Business Report for {days} days.")
        st.stop()
    
    # Normalize SKU
    df[sku_col_br] = df[sku_col_br].astype(str).str.strip().str.upper()

    # Find ASIN in BR
    asin_col_br = next((c for c in ["(Parent) ASIN", "ASIN", "asin"] if c in df.columns), None)
    
    # Total Page Views (B2C + B2B)
    df["Period Page Views"] = df["Page Views - Total"] + df["Page Views - Total - B2B"]

    # Pivot
    agg_dict = {
        "Period Sales Qty": "sum",
        "Period Page Views": "sum",
        "Sessions - Total": "sum",
        "Buy Box Percentage": "mean",
        "Unit Session Percentage": "mean"
    }
    if asin_col_br:
        agg_dict[asin_col_br] = "first"

    pivot = df.groupby(sku_col_br).agg(agg_dict).reset_index()
    
    cols = [
        "SKU", 
        f"{days} Days Sales Qty", 
        f"{days} Days Page Views",
        f"{days} Days Sessions",
        f"{days} Days Buy Box %",
        f"{days} Days Unit Session %"
    ]
    if asin_col_br:
        cols.append(f"(Parent) ASIN_{days}")

    pivot.columns = cols
    return pivot

def create_excel(df_dict, intervals):
    """
    df_dict: Dictionary of {SheetName: DataFrame}
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Ranges and Colors (cl.py 7-color logic)
            colors = {
                "critical": PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"), # Red 0-7
                "low": PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid"),      # Orange 7-15
                "optimal": PatternFill(start_color="008000", end_color="008000", fill_type="solid"),  # Green 15-30
                "monitor": PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"),  # Yellow 30-45
                "high": PatternFill(start_color="87CEEB", end_color="87CEEB", fill_type="solid"),     # Sky Blue 45-60
                "excess": PatternFill(start_color="A52A2A", end_color="A52A2A", fill_type="solid"),   # Brown 60-90
                "overstock": PatternFill(start_color="000000", end_color="000000", fill_type="solid") # Black >90
            }
            font_white = Font(color="FFFFFF", bold=True)
            font_black = Font(color="000000", bold=True)

            # Apply styling to DOC columns
            doc_cols = [c for c in df.columns if "DOC" in c]
            col_indices = [df.columns.get_loc(c) + 1 for c in doc_cols]
            
            for r in range(2, worksheet.max_row + 1):
                for c_idx in col_indices:
                    cell = worksheet.cell(row=r, column=c_idx)
                    try:
                        val = float(cell.value)
                        if val <= 7:
                            cell.fill = colors["critical"]
                            cell.font = font_white
                        elif val <= 15:
                            cell.fill = colors["low"]
                            cell.font = font_black
                        elif val <= 30:
                            cell.fill = colors["optimal"]
                            cell.font = font_white
                        elif val <= 45:
                            cell.fill = colors["monitor"]
                            cell.font = font_black
                        elif val <= 60:
                            cell.fill = colors["high"]
                            cell.font = font_black
                        elif val <= 90:
                            cell.fill = colors["excess"]
                            cell.font = font_white
                        else:
                            cell.fill = colors["overstock"]
                            cell.font = font_white
                    except:
                        pass
                        
            # Header styling
            for cell in worksheet[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
            
    return output.getvalue()

if process_btn:
    if not (pm_file and inv_file):
        st.error("Please upload Purchase Master and Inventory Report.")
    else:
        with st.spinner("Processing reports..."):
            # 1. Process Inventory
            inventory_df = pd.read_csv(inv_file)
            
            # Find SKU column in Inventory
            sku_col_inv = next((c for c in ["sku", "Seller SKU", "SKU"] if c in inventory_df.columns), None)
            if not sku_col_inv:
                st.error("Could not find a SKU column in Inventory report. Expected 'sku', 'Seller SKU', or 'SKU'.")
                st.stop()
            
            # Normalize SKU
            inventory_df[sku_col_inv] = inventory_df[sku_col_inv].astype(str).str.strip().str.upper()
            
            # Find ASIN column in Inventory
            asin_col_inv = next((c for c in ["asin", "ASIN"] if c in inventory_df.columns), None)

            # Process Reserved Report if provided
            reserved_df = pd.DataFrame()
            if res_file:
                reserved_df = pd.read_csv(res_file)
                sku_col_res = next((c for c in ["sku", "SKU", "Seller SKU"] if c in reserved_df.columns), None)
                if sku_col_res:
                    reserved_df = reserved_df.rename(columns={sku_col_res: "SKU"})
                    # Ensure numeric
                    for col in ["reserved_customerorders", "reserved_fc-transfers", "reserved_fc-processing"]:
                        reserved_df[col] = clean_numeric_col(reserved_df, col)
                    reserved_df = reserved_df.groupby("SKU")[["reserved_customerorders", "reserved_fc-transfers", "reserved_fc-processing"]].sum().reset_index()

            # Merge Reserved into Inventory
            inventory_df = inventory_df.rename(columns={sku_col_inv: "SKU"})
            if not reserved_df.empty:
                inventory_df = inventory_df.merge(reserved_df, on="SKU", how="left")
                # Fill NaNs from reserved
                for col in ["reserved_customerorders", "reserved_fc-transfers", "reserved_fc-processing"]:
                    if col in inventory_df.columns:
                        inventory_df[col] = inventory_df[col].fillna(0)
                    else:
                        inventory_df[col] = 0
            else:
                inventory_df["reserved_customerorders"] = 0
                inventory_df["reserved_fc-transfers"] = 0
                inventory_df["reserved_fc-processing"] = 0

            # Calculate Total Stock using formula from previous integration: 
            # afn-fulfillable + fc-transfers + fc-processing
            inventory_df["afn-fulfillable-qty"] = clean_numeric_col(inventory_df, "afn-fulfillable-quantity")
            inventory_df["afn-reserved-qty"] = clean_numeric_col(inventory_df, "afn-reserved-quantity")
            
            # Formula refinement: Total Stock is Fulfillable + non-order reserved
            inventory_df["Total Stock"] = (
                inventory_df["afn-fulfillable-qty"] - 
                inventory_df["reserved_customerorders"] +
                inventory_df["reserved_fc-transfers"] + 
                inventory_df["reserved_fc-processing"]
            )
            
            inventory_df["Transfer Stock"] = (
                clean_numeric_col(inventory_df, "afn-inbound-working-quantity") + 
                clean_numeric_col(inventory_df, "afn-inbound-shipped-quantity") +
                clean_numeric_col(inventory_df, "afn-inbound-receiving-quantity")
            )
            
            inv_cols = ["SKU", "Total Stock", "Transfer Stock", "afn-fulfillable-qty", "afn-reserved-qty", 
                        "reserved_customerorders", "reserved_fc-transfers", "reserved_fc-processing"]
            if asin_col_inv:
                inv_cols.append(asin_col_inv)
            
            inv_subset = inventory_df[inv_cols].copy()
            if asin_col_inv:
                inv_subset = inv_subset.rename(columns={asin_col_inv: "(Parent) ASIN"})
            
            inv_subset = inv_subset.groupby("SKU").agg({
                "Total Stock": "sum",
                "Transfer Stock": "sum",
                "afn-fulfillable-qty": "sum",
                "afn-reserved-qty": "sum",
                "reserved_customerorders": "sum",
                "reserved_fc-transfers": "sum",
                "reserved_fc-processing": "sum",
                **({"(Parent) ASIN": "first"} if asin_col_inv else {})
            }).reset_index()

            # 2. Process Purchase Master
            pm_df = pd.read_excel(pm_file)
            
            # Find SKU column in PM
            possible_sku_cols = ["Seller SKU", "Amazon Sku Name", "SKU", "EasycomSKU", "sku", "Amazon Sku"]
            sku_col_pm = next((c for c in possible_sku_cols if c in pm_df.columns), None)
            
            if not sku_col_pm:
                st.error(f"Could not find a SKU column in Purchase Master. Expected one of: {possible_sku_cols}")
                st.stop()
            
            # Normalize SKU
            pm_df[sku_col_pm] = pm_df[sku_col_pm].astype(str).str.strip().str.upper()
                
            # Columns to keep (with safety checks)
            meta_cols = {
                "Brand": "Brand",
                "Product Name": "Product Name",
                "Brand Manager": "Brand Manager",
                "CP": "CP",
                "MRP": "MRP",
                "Vendor SKU Codes": "Vendor SKU Codes",
                "ASIN": "(Parent) ASIN",
                "(Parent) ASIN": "(Parent) ASIN",
                "asin": "(Parent) ASIN"
            }
            cols_to_keep = [sku_col_pm]
            for orig, target in meta_cols.items():
                if orig in pm_df.columns: cols_to_keep.append(orig)
            
            pm_subset = pm_df[cols_to_keep].copy()
            pm_subset = pm_subset.rename(columns={sku_col_pm: "SKU", **{k:v for k,v in meta_cols.items() if k in pm_df.columns}})
            pm_subset = pm_subset.drop_duplicates(subset=["SKU"])
            
            # Additional Processing for PM (CP and MRP cleaning)
            for col in ["CP", "MRP"]:
                if col in pm_subset.columns:
                    pm_subset[col] = pd.to_numeric(
                        pm_subset[col].astype(str).str.replace(",", "", regex=False).str.strip(), 
                        errors="coerce"
                    ).fillna(0)

            # Listing Status and seller-sku lookup will be performed at the end to cover all merged SKUs

            # 3. Process Business Reports
            consolidated_df = pm_subset.merge(inv_subset, on="SKU", how="outer", suffixes=('', '_inv'))
            
            # Consolidate (Parent) ASIN if it exists in both
            if "(Parent) ASIN_inv" in consolidated_df.columns:
                if "(Parent) ASIN" not in consolidated_df.columns:
                    consolidated_df["(Parent) ASIN"] = consolidated_df["(Parent) ASIN_inv"]
                else:
                    consolidated_df["(Parent) ASIN"] = consolidated_df["(Parent) ASIN"].fillna(consolidated_df["(Parent) ASIN_inv"])
                consolidated_df.drop(columns=["(Parent) ASIN_inv"], inplace=True)
            
            consolidated_df["Total Stock"] = consolidated_df["Total Stock"].fillna(0)
            consolidated_df["Transfer Stock"] = consolidated_df["Transfer Stock"].fillna(0)
            
            drr_cols = []
            
            for days in intervals:
                if br_files[days]:
                    br_pivot = process_br(br_files[days], days)
                    consolidated_df = consolidated_df.merge(br_pivot, on="SKU", how="outer")
                    
                    # Consolidate ASIN from BR if missing
                    br_asin_col = f"(Parent) ASIN_{days}"
                    if br_asin_col in consolidated_df.columns:
                        if "(Parent) ASIN" not in consolidated_df.columns:
                            consolidated_df["(Parent) ASIN"] = consolidated_df[br_asin_col]
                        else:
                            consolidated_df["(Parent) ASIN"] = consolidated_df["(Parent) ASIN"].fillna(consolidated_df[br_asin_col])
                        consolidated_df.drop(columns=[br_asin_col], inplace=True)

                    # Safety check for numeric columns that might have NaNs after outer join
                    num_cols_to_fill = ["Total Stock", "Transfer Stock", "afn-fulfillable-qty", "afn-reserved-qty", 
                                        "reserved_customerorders", "reserved_fc-transfers", "reserved_fc-processing", "CP", "MRP"]
                    for col in num_cols_to_fill:
                        if col in consolidated_df.columns:
                            consolidated_df[col] = consolidated_df[col].fillna(0)

                    # Fill NaN with 0 for current interval Sales and Page Views
                    consolidated_df[f"{days} Days Sales Qty"] = consolidated_df[f"{days} Days Sales Qty"].fillna(0)
                    consolidated_df[f"{days} Days Page Views"] = consolidated_df[f"{days} Days Page Views"].fillna(0)
                    
                    # Calculations
                    consolidated_df[f"{days} Days DRR"] = (consolidated_df[f"{days} Days Sales Qty"] / days).round(2)
                    
                    # DOC calculation: avoid division by zero
                    consolidated_df[f"{days} Days DOC"] = consolidated_df.apply(
                        lambda row: round(row["Total Stock"] / row[f"{days} Days DRR"], 0) if row[f"{days} Days DRR"] > 0 else 0,
                        axis=1
                    )
                    drr_cols.append(f"{days} Days DRR")

            # 4. Final Data Enrichment (Listing Status & seller-sku)
            consolidated_df["Listing Status"] = ""
            consolidated_df["seller-sku"] = ""
            if list_file:
                if list_file.name.endswith(".csv"):
                    listing_df = pd.read_csv(list_file)
                else:
                    listing_df = pd.read_excel(list_file)
                
                # Use 4th column for SKU as per cl.py logic
                if len(listing_df.columns) >= 4:
                    sku_series = listing_df.iloc[:, 3].astype(str).str.strip().str.upper()
                    listing_skus_dict = dict(zip(sku_series, sku_series))
                    
                    # Apply to all SKUs in consolidated_df
                    temp_sku = consolidated_df["SKU"].astype(str).str.strip().str.upper()
                    consolidated_df["seller-sku"] = temp_sku.map(listing_skus_dict).fillna("")
                    consolidated_df["Listing Status"] = temp_sku.apply(lambda x: "Closed" if x in listing_skus_dict else "")

            # 5. Aggregate Metrics and Specialized Reports
            if drr_cols:
                consolidated_df["Max DRR"] = consolidated_df[drr_cols].max(axis=1)
                consolidated_df["Avg DRR"] = consolidated_df[drr_cols].mean(axis=1).round(2)
                
                # Base DOC for Inventory/OOS/Overstock based on Max DRR
                consolidated_df["DOC (Max)"] = consolidated_df.apply(
                    lambda row: round(row["Total Stock"] / row["Max DRR"], 0) if row["Max DRR"] > 0 else 0,
                    axis=1
                ).fillna(0).astype(int)

            # Calculated Financial Columns
            if "CP" in consolidated_df.columns:
                consolidated_df["CP As Per Total Stock Qty"] = (consolidated_df["CP"] * consolidated_df["Total Stock"]).round(2)
                for days in intervals:
                    if f"{days} Days Sales Qty" in consolidated_df.columns:
                        consolidated_df[f"{days} CP As Per Total Sale Qty"] = (consolidated_df["CP"] * consolidated_df[f"{days} Days Sales Qty"]).round(2)

            # Define Layout order precisely as requested by the user
            # SKU, (Parent) ASIN, Vendor SKU Codes, Brand, Brand Manager, Product Name, CP, 
            # afn-fulfillable-qty, afn-reserved-qty, reserved_customerorders, reserved_fc-transfers, reserved_fc-processing, 
            # Total Stock, CP As Per Total Stock Qty, ... multi-day ...
            # seller-sku, Listing Status
            layout_meta = ["SKU", "(Parent) ASIN", "Vendor SKU Codes", "Brand", "Brand Manager", "Product Name", "CP"]
            
            layout_stock = ["afn-fulfillable-qty", "afn-reserved-qty", "reserved_customerorders", 
                           "reserved_fc-transfers", "reserved_fc-processing", "Total Stock", "CP As Per Total Stock Qty"]
            
            layout_metrics = []
            for days in intervals:
                # {day} Days Sales Qty, {day} CP As Per Total Sale Qty, {day} Days DRR, {day} Days DOC, {day} Days Page Views
                prefix = f"{days} Days"
                cp_prefix = f"{days} CP As Per Total Sale Qty"
                if f"{prefix} Sales Qty" in consolidated_df.columns:
                    layout_metrics.extend([
                        f"{prefix} Sales Qty",
                        cp_prefix,
                        f"{prefix} DRR", 
                        f"{prefix} DOC", 
                        f"{prefix} Page Views"
                    ])
            
            ordered_cols = [c for c in (layout_meta + layout_stock + layout_metrics + ["seller-sku", "Listing Status"]) if c in consolidated_df.columns]
            final_df = consolidated_df[ordered_cols].copy()
            
            # Create Specialized Reports (All tabs use the uniform layout `final_df`)
            inv_view = final_df.copy()
            
            # OOS View (Fulfillable == 0 or Total Stock == 0)
            oos_view = final_df[final_df["Total Stock"] == 0].copy()
            
            # Overstock View (Any interval DOC > 90 or DOC (Max) > 90)
            # Find all DOC columns
            doc_cols = [c for c in final_df.columns if "DOC" in c]
            if doc_cols:
                overstock_view = final_df[final_df[doc_cols].gt(90).any(axis=1)].copy()
            else:
                overstock_view = final_df[final_df["Total Stock"] > 500].copy() # Fallback
            
            # Create Pivot Tables for OOS and Overstock
            oos_pivot = create_stock_pivot(oos_view)
            overstock_pivot = create_stock_pivot(overstock_view)
            
            st.success("Analysis Complete!")
            
            tab1, tab2, tab3, tab4 = st.tabs(["📊 Consolidated View", "📦 Inventory View", "⚠️ OOS Analysis", "📈 Overstock"])
            
            with tab1:
                st.subheader("Multi-Day Consolidated Analysis")
                st.dataframe(apply_doc_styling(final_df))
            
            with tab2:
                st.subheader("Inventory Breakdown (Full Layout)")
                st.dataframe(apply_doc_styling(inv_view))
                
            with tab3:
                st.subheader("Out of Stock (OOS) Items")
                st.dataframe(apply_doc_styling(oos_view))
                st.divider()
                st.subheader("OOS Pivot Table")
                st.dataframe(apply_doc_styling(oos_pivot))
                
            with tab4:
                st.subheader("Overstock Analysis (DOC > 90)")
                st.dataframe(apply_doc_styling(overstock_view))
                st.divider()
                st.subheader("Overstock Pivot Table")
                st.dataframe(apply_doc_styling(overstock_pivot))
            
            # Excel Download
            excel_sheets = {
                "Consolidated": final_df,
                "Inventory": inv_view,
                "OOS_Report": oos_view,
                "OOS_Pivot": oos_pivot,
                "Overstock_Report": overstock_view,
                "Overstock_Pivot": overstock_pivot
            }
            excel_data = create_excel(excel_sheets, intervals)
            st.download_button(
                label="📥 Download All Reports (Excel)",
                data=excel_data,
                file_name=f"Multi_Day_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

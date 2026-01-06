import streamlit as st
import pandas as pd
import itertools
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="BC Posting Setup Pro", layout="wide")

# --- SIDEBAR SETTINGS (User Friendly Options) ---
st.sidebar.header("⚙️ Matrix Settings")
include_blank = st.sidebar.checkbox("Include Blank Bus. Group", value=True)
blank_desc_text = st.sidebar.text_input("Description for Blank Group", value="Standard")
default_blocked = st.sidebar.checkbox("Default Blocked Status", value=False)
default_lookup = st.sidebar.checkbox("View All Accounts on Lookup", value=True)

st.title("📊 BC General Posting Setup Matrix")
st.markdown("Generate a full Cartesian product of your Business and Product groups for Business Central.")

# --- DEFINING ALL COLUMNS ---
STANDARD_COLUMNS = [
    "Gen. Bus. Posting Group", "Gen. Prod. Posting Group", "Description",
    "Sales Account", "Sales Credit Memo Account", "Sales Line Disc. Account",
    "Sales Inv. Disc. Account", "Sales Pmt. Disc. Debit Acc.", "Sales Pmt. Disc. Credit Acc.",
    "Purch. Account", "Purch. Credit Memo Account", "Purch. Line Disc. Account",
    "Purch. Inv. Disc. Account", "Purch. Pmt. Disc. Debit Acc.", "Purch. Pmt. Disc. Credit Acc.",
    "COGS Account", "Inventory Adjmt. Account", "Direct Cost Applied Account",
    "Overhead Applied Account", "Purchase Variance Account",
    "View All Accounts on Lookup", "Blocked"
]

# --- FILE UPLOADERS ---
col1, col2 = st.columns(2)
with col1:
    bus_file = st.file_uploader("1. Upload Gen. Bus. Groups (.xlsx)", type=['xlsx'])
with col2:
    prod_file = st.file_uploader("2. Upload Gen. Prod. Groups (.xlsx)", type=['xlsx'])

if bus_file and prod_file:
    try:
        # Load Data
        df_bus = pd.read_excel(bus_file)
        df_prod = pd.read_excel(prod_file)

        # Fool-proof Validation: Check for 'Code' column
        if 'Code' not in df_bus.columns or 'Code' not in df_prod.columns:
            st.error("❌ Error: Both files must have a column named 'Code'. Please check your BC exports.")
            st.stop()

        # Data Cleaning: Remove spaces and drop empty rows
        df_bus['Code'] = df_bus['Code'].astype(str).str.strip()
        df_prod['Code'] = df_prod['Code'].astype(str).str.strip()
        df_bus = df_bus[df_bus['Code'] != 'nan']
        df_prod = df_prod[df_prod['Code'] != 'nan']

        bus_codes = df_bus['Code'].unique().tolist()
        prod_codes = df_prod['Code'].unique().tolist()

        if include_blank:
            bus_codes.insert(0, "")

        # Generate Cartesian Product
        combinations = list(itertools.product(bus_codes, prod_codes))
        output_df = pd.DataFrame(combinations, columns=['Gen. Bus. Posting Group', 'Gen. Prod. Posting Group'])

        # Show Metrics for confidence
        m1, m2, m3 = st.columns(3)
        m1.metric("Business Groups", len(bus_codes))
        m2.metric("Product Groups", len(prod_codes))
        m3.metric("Total Rows to Generate", len(output_df))

        # Merge Descriptions
        output_df = output_df.merge(
            df_bus[['Code', 'Description']], 
            left_on='Gen. Bus. Posting Group', 
            right_on='Code', 
            how='left'
        ).rename(columns={'Description': 'Bus_Desc'})
        
        # Set description for blank group
        output_df.loc[output_df['Gen. Bus. Posting Group'] == "", 'Bus_Desc'] = blank_desc_text

        output_df = output_df.merge(
            df_prod[['Code', 'Description']], 
            left_on='Gen. Prod. Posting Group', 
            right_on='Code', 
            how='left'
        ).rename(columns={'Description': 'Prod_Desc'})

        output_df['Description'] = output_df['Bus_Desc'].fillna('') + ' - ' + output_df['Prod_Desc'].fillna('')

        # Add all missing columns from the standard list
        for col in STANDARD_COLUMNS:
            if col not in output_df.columns:
                output_df[col] = None

        # Apply User Settings from Sidebar
        output_df['View All Accounts on Lookup'] = default_lookup
        output_df['Blocked'] = default_blocked

        # Final reorder to match STANDARD_COLUMNS exactly
        output_df = output_df[STANDARD_COLUMNS]

        # --- PREVIEW & DOWNLOAD ---
        st.divider()
        st.subheader("Data Preview (First 15 rows)")
        st.dataframe(output_df.head(15), use_container_width=True)

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False, sheet_name='Gen. Posting Setup')
        
        st.download_button(
            label="📥 Download Excel for Configuration Package",
            data=buffer.getvalue(),
            file_name="BC_General_Posting_Setup_Matrix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"⚠️ An unexpected error occurred: {e}")
else:
    st.info("Upload the Excel files exported from Business Central to begin.")
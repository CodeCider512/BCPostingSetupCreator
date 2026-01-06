import streamlit as st
import pandas as pd
import itertools
from io import BytesIO

# --- PAGE CONFIG ---
st.set_page_config(page_title="BC Implementation Toolkit", layout="wide")

# --- TOOL NAVIGATION ---
st.sidebar.title("🛠️ BC Toolkit")
tool_selection = st.sidebar.radio(
    "Select Matrix Generator:",
    ["General Posting Setup", "Inventory Posting Setup"]
)

st.sidebar.divider()
st.sidebar.header("⚙️ Settings")
include_blank = st.sidebar.checkbox("Include Blank Group/Location", value=True)
blank_desc_text = st.sidebar.text_input("Description for Blank", value="Standard / All Locations")

# --- COLUMN DEFINITIONS ---
GEN_POSTING_COLUMNS = [
    "Gen. Bus. Posting Group", "Gen. Prod. Posting Group", "Description",
    "Sales Account", "Sales Credit Memo Account", "Sales Line Disc. Account",
    "Sales Inv. Disc. Account", "Sales Pmt. Disc. Debit Acc.", "Sales Pmt. Disc. Credit Acc.",
    "Purch. Account", "Purch. Credit Memo Account", "Purch. Line Disc. Account",
    "Purch. Inv. Disc. Account", "Purch. Pmt. Disc. Debit Acc.", "Purch. Pmt. Disc. Credit Acc.",
    "COGS Account", "Inventory Adjmt. Account", "Direct Cost Applied Account",
    "Overhead Applied Account", "Purchase Variance Account",
    "View All Accounts on Lookup", "Blocked"
]

INV_POSTING_COLUMNS = [
    "Location Code", "Invt. Posting Group Code", "Description",
    "View All Accounts on Lookup", "Inventory Account", "Inventory Account (Interim)",
    "WIP Account", "Material Variance Account", "Capacity Variance Account",
    "Subcontracted Variance Account", "Cap. Overhead Variance Account",
    "Mfg. Overhead Variance Account", "Material Non-Inventory Variance Account"
]

# --- SHARED HELPER FUNCTIONS ---
def generate_matrix(df_a, df_b, col_a_name, col_b_name, final_columns):
    # Clean data
    df_a['Code'] = df_a['Code'].astype(str).str.strip()
    df_b['Code'] = df_b['Code'].astype(str).str.strip()
    
    codes_a = df_a['Code'].unique().tolist()
    codes_b = df_b['Code'].unique().tolist()
    
    if include_blank:
        codes_a.insert(0, "")

    # Generate Combinations
    combinations = list(itertools.product(codes_a, codes_b))
    output_df = pd.DataFrame(combinations, columns=[col_a_name, col_b_name])

    # Merge Descriptions
    output_df = output_df.merge(df_a[['Code', 'Description']], left_on=col_a_name, right_on='Code', how='left').rename(columns={'Description': 'Desc_A'})
    output_df.loc[output_df[col_a_name] == "", 'Desc_A'] = blank_desc_text
    output_df = output_df.merge(df_b[['Code', 'Description']], left_on=col_b_name, right_on='Code', how='left').rename(columns={'Description': 'Desc_B'})
    
    output_df['Description'] = output_df['Desc_A'].fillna('') + ' - ' + output_df['Desc_B'].fillna('')

    # Format Columns
    for col in final_columns:
        if col not in output_df.columns:
            output_df[col] = None
    
    output_df['View All Accounts on Lookup'] = True
    if "Blocked" in final_columns:
        output_df['Blocked'] = False
        
    return output_df[final_columns]

# --- UI LOGIC ---
if tool_selection == "General Posting Setup":
    st.title("📊 General Posting Setup Generator")
    col1, col2 = st.columns(2)
    file_a = col1.file_uploader("Upload Gen. Bus. Groups", type=['xlsx'])
    file_b = col2.file_uploader("Upload Gen. Prod. Groups", type=['xlsx'])
    
    col_a_title, col_b_title = "Gen. Bus. Posting Group", "Gen. Prod. Posting Group"
    target_columns = GEN_POSTING_COLUMNS

else:
    st.title("📦 Inventory Posting Setup Generator")
    col1, col2 = st.columns(2)
    file_a = col1.file_uploader("Upload Locations", type=['xlsx'])
    file_b = col2.file_uploader("Upload Inventory Posting Groups", type=['xlsx'])
    
    col_a_title, col_b_title = "Location Code", "Invt. Posting Group Code"
    target_columns = INV_POSTING_COLUMNS

# --- EXECUTION ---
if file_a and file_b:
    try:
        df_a_in = pd.read_excel(file_a)
        df_b_in = pd.read_excel(file_b)

        if 'Code' not in df_a_in.columns or 'Code' not in df_b_in.columns:
            st.error("❌ Error: Both files must have a 'Code' column.")
        else:
            final_df = generate_matrix(df_a_in, df_b_in, col_a_title, col_b_title, target_columns)
            
            st.divider()
            st.metric("Total Rows Generated", len(final_df))
            st.dataframe(final_df, use_container_width=True)

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)
            
            st.download_button(
                label=f"📥 Download {tool_selection} Matrix",
                data=buffer.getvalue(),
                file_name=f"BC_{tool_selection.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info(f"Please upload the required files for {tool_selection}.")

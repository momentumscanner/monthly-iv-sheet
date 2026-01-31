import streamlit as st
import os
import tempfile
import pandas as pd
from processor import Processor
import shutil

# Set page configuration
st.set_page_config(
    page_title="Monthly IV Sheet Generator", 
    page_icon="📈",
    layout="wide"
)

def save_uploaded_file(uploaded_file):
    """Save uploaded file to a temporary directory and return the path."""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{uploaded_file.name}") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            return tmp_file.name
    except Exception as e:
        st.error(f"Error saving file: {e}")
        return None

def main():
    st.title("Monthly IV Sheet Generator 📈")
    st.markdown("### Using NSE Bhav Copy and icharts Daywise Reports")
    st.divider()

    # Initialize session state for expiries
    if 'expiries' not in st.session_state:
        st.session_state.expiries = []
    
    # --- 1. Upload Files ---
    st.subheader("📂 1. Upload Files")
    
    col1, col2 = st.columns(2)
    
    with col1:
        zip_file = st.file_uploader("Upload Bhav Copy ZIP", type="zip", help="The daily Bhav Copy ZIP file from NSE")
        
    with col2:
        daywise_file = st.file_uploader("Upload Daywise Report", type="xlsx", help="The Daywise Price and OI Summary Excel file")
    
    st.write("") # Spacer
    
    # --- Logic ---
    if zip_file and daywise_file:
        # Save files temporarily
        zip_path = save_uploaded_file(zip_file)
        daywise_path = save_uploaded_file(daywise_file)
        
        if zip_path and daywise_path:
            # Scan for expiries automatically
            if not st.session_state.expiries:
                try:
                    with st.spinner("Scanning file for available expiries..."):
                        p = Processor(zip_path)
                        expiries = p.get_expiries()
                        if expiries:
                            st.session_state.expiries = list(expiries)
                            st.rerun()
                except Exception as e:
                    st.warning(f"Could not scan expiries: {e}")

            st.divider()
            
            # --- 2. Configuration & Process ---
            st.subheader("⚙️ 2. Configuration & Process")
            
            # Configuration Row
            config_col1, config_col2, config_col3 = st.columns(3)
            
            with config_col1:
                instrument = st.selectbox("Select Instrument", ["STOCK FUTURE", "INDEX FUTURE (Coming Soon)"])
            
            with config_col2:
                expiry_options = ["All Expiries"] + st.session_state.expiries
                target_expiry = st.selectbox("Select Expiry", expiry_options)
            
            with config_col3:
                st.write("Export Format")
                col_ex_1, col_ex_2 = st.columns(2)
                with col_ex_1:
                    export_excel = st.checkbox("Excel", value=True)
                with col_ex_2:
                    export_csv = st.checkbox("CSV", value=True)

            st.write("") # Spacer
            
            # Process Button
            process_btn = st.button("Start Processing", type="primary", use_container_width=True)
            
            if process_btn:
                if instrument != "STOCK FUTURE":
                    st.error("Only STOCK FUTURE is currently supported.")
                else:
                    progress_bar = st.progress(0)
                    status_area = st.empty()
                    
                    # Log container
                    with st.expander("View Processing Logs", expanded=True):
                        log_container = st.container()
                        log_text = log_container.empty()
                        logs = []

                        def log_callback(msg):
                            logs.append(msg)
                            # Keep only last 10 logs to avoid UI lag if very verbose
                            log_text.code("\n".join(logs[-10:]))
                        
                        try:
                            processor = Processor(zip_path, daywise_path)
                            processor.set_log_callback(log_callback)
                            
                            selected_expiry = None if target_expiry == "All Expiries" else target_expiry
                            
                            with st.spinner("Processing... please wait"):
                                files = processor.process(
                                    export_excel=export_excel,
                                    export_csv=export_csv,
                                    target_expiry=selected_expiry
                                )
                            
                            progress_bar.progress(100)
                            status_area.success("✅ Processing Complete!")
                            
                            st.divider()
                            st.subheader("📥 3. Download Results")
                            
                            # Display Download Buttons
                            dl_cols = st.columns(len(files))
                            for idx, file_path in enumerate(files):
                                original_name = os.path.basename(file_path)
                                extension = os.path.splitext(original_name)[1]
                                
                                # Custom Filename Logic: "Monthly IV {Selected Expiry}"
                                if selected_expiry:
                                    # Format: Monthly IV 29Jan2026.xlsx
                                    # Convert expiry date to friendly format if possible, or just use as is
                                    try:
                                        dt = pd.to_datetime(selected_expiry)
                                        exp_str = dt.strftime("%d%b%Y").upper()
                                        custom_name = f"Monthly IV {exp_str}{extension}"
                                    except:
                                        custom_name = f"Monthly IV {selected_expiry}{extension}"
                                else:
                                    custom_name = f"Monthly IV All Expiries{extension}"

                                with dl_cols[idx]:
                                    with open(file_path, "rb") as f:
                                        st.download_button(
                                            label=f"Download {custom_name}",
                                            data=f,
                                            file_name=custom_name,
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if extension == '.xlsx' else "text/csv",
                                            use_container_width=True
                                        )
                                        
                        except Exception as e:
                            status_area.error(f"Processing failed: {e}")
    
    else:
        st.info("Waiting for file uploads...")

if __name__ == "__main__":
    main()

import streamlit as st
import pandas as pd
import numpy as np

def calculate_percent_change(initial_value, final_value):
    """Calculate the percentage change from initial to final value."""
    if initial_value == 0:
        return None  # Cannot calculate percent change when initial value is zero
    
    percent_change = ((final_value - initial_value) / abs(initial_value)) * 100
    return percent_change

# Set page config
st.set_page_config(
    page_title="Percent Change Calculator",
    page_icon="📊",
    layout="centered"
)

# App title and description
st.title("Percent Change Calculator")
st.markdown("""
This app calculates the percentage change between two values.
Enter the initial and final values to see the percentage increase or decrease.
""")

# Create two columns for input fields
col1, col2 = st.columns(2)

with col1:
    initial_value = st.number_input("Initial Value", value=100.0, step=1.0)

with col2:
    final_value = st.number_input("Final Value", value=150.0, step=1.0)

# Calculate percent change when button is clicked
if st.button("Calculate Percent Change", type="primary"):
    if initial_value == 0:
        st.error("Initial value cannot be zero when calculating percent change.")
    else:
        percent_change = calculate_percent_change(initial_value, final_value)
        
        # Display result with appropriate formatting and color
        st.subheader("Result:")
        
        if percent_change > 0:
            st.success(f"**Increase of {percent_change:.2f}%**")
            change_direction = "increased"
        elif percent_change < 0:
            st.error(f"**Decrease of {abs(percent_change):.2f}%**")
            change_direction = "decreased"
        else:
            st.info("**No change (0%)**")
            change_direction = "remained the same"
        
        if percent_change != 0:
            st.markdown(f"The value {change_direction} from {initial_value} to {final_value}, "
                      f"a {abs(percent_change):.2f}% {change_direction}.")

# Add advanced options in an expandable section
with st.expander("Advanced Options"):
    st.subheader("Batch Calculation")
    st.markdown("Upload a CSV file with two columns: 'Initial' and 'Final' to calculate percent changes for multiple values.")
    
    uploaded_file = st.file_uploader("Choose a CSV file", type="csv")
    
    if uploaded_file is not None:
        # Read the CSV file
        try:
            df = pd.read_csv(uploaded_file)
            
            # Check if the required columns exist
            if 'Initial' in df.columns and 'Final' in df.columns:
                # Calculate percent changes
                df['Percent Change'] = df.apply(
                    lambda row: calculate_percent_change(row['Initial'], row['Final']), 
                    axis=1
                )
                
                # Display the results
                st.dataframe(df.style.format({'Percent Change': '{:.2f}%'}))
                
                # Add download button for the results
                csv = df.to_csv(index=False)
                st.download_button(
                    label="Download Results as CSV",
                    data=csv,
                    file_name="percent_change_results.csv",
                    mime="text/csv"
                )
            else:
                st.error("CSV file must contain 'Initial' and 'Final' columns.")
        except Exception as e:
            st.error(f"Error processing file: {e}")

# Add information section
st.sidebar.title("Information")
st.sidebar.markdown("""
### Formula Used:
Percent Change = ((Final Value - Initial Value) / |Initial Value|) × 100

### Examples:
- Initial: 100, Final: 150 → 50% increase
- Initial: 200, Final: 150 → 25% decrease
- Initial: 100, Final: 100 → 0% (no change)

### Notes:
- Initial value cannot be zero
- A positive result indicates an increase
- A negative result indicates a decrease
""")

# Footer
st.markdown("---")
st.markdown("Created with Streamlit and Python")

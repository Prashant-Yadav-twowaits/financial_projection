import io
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import calendar
from dateutil.relativedelta import relativedelta
import os

def load_projections(uploaded_file):
    """
    Load the financial projections from the uploaded Excel file.
    """
    try:
        if uploaded_file is not None:
            # Load the Excel file with all sheets
            xl = pd.ExcelFile(uploaded_file)

            # Get all sheet names
            all_sheets = xl.sheet_names
            # st.info(f"Available sheets in the Excel file: {all_sheets}")

            # Create a dictionary to store projections for each sport
            projections = {}

            # Look for sheets for each sport
            sports = ['badminton', 'football', 'pickleball']

            found_any_sport = False

            for sport in sports:
                if sport in all_sheets:
                    try:
                        raw_data = pd.read_excel(xl, sheet_name=sport)
                        if not raw_data.empty and len(raw_data.columns) > 1:
                            metrics = raw_data.iloc[:, 0].astype(str).tolist()
                            months = raw_data.columns[1:].astype(str).tolist()
                            data_values = raw_data.iloc[:, 1:].values

                            sport_data = pd.DataFrame(data_values, index=metrics, columns=months)
                            sport_data.reset_index(inplace=True)
                            sport_data.rename(columns={'index': 'Metric'}, inplace=True)

                            # st.info(f"Columns loaded for {sport}: {sport_data.columns.tolist()}")
                            projections[sport] = sport_data
                            st.success(f"Loaded projection data for {sport} from sheet '{sport}'")
                            found_any_sport = True
                        else:
                            st.warning(f"Sheet '{sport}' is empty or has insufficient columns.")
                    except Exception as e:
                        st.error(f"Error loading data for {sport}: {e}")

            if found_any_sport:
                return projections
            else:
                st.warning("No projection data found for any sport in the provided Excel file.")
                return None
        else:
            st.info("Please upload an Excel file.")
            return None

    except Exception as e:
        st.error(f"Error loading projections: {e}")
        return None

def generate_venue_projection(start_date, base_projection, num_months):
    """
    Generate projection for a single venue starting from a specific date.
    The projection data is transposed with months as columns and metrics as rows.
    """
    if 'Metric' not in base_projection.columns:
        st.error("Error: 'Metric' column not found in the base projection data.")
        return pd.DataFrame()

    # Extract metric values and transpose the data
    metric_values = base_projection.set_index('Metric').drop(columns=['Month-Year'], errors='ignore')
    transposed_data = pd.DataFrame()

    month_year_list = []  # Track month-year values for the new column

    for i in range(1, num_months + 1):
        month_col_name = f"Month {i}"
        month_year = (start_date + relativedelta(months=i-1)).strftime('%b-%Y')
        month_year_list.append(month_year)

        if month_col_name in metric_values.columns:
            transposed_data[month_year] = metric_values[month_col_name]
        else:
            st.warning(f"Warning: Could not find column '{month_col_name}' in base projection.")
            transposed_data[month_year] = np.nan

    # Add metrics as the first column
    transposed_data.insert(0, 'Metric', metric_values.index)

    # Add 'Month-Year' column for tracking
    transposed_data['Month-Year'] = month_year_list

    return transposed_data.reset_index(drop=True)

def consolidate_projections(venues_data, all_dates):
    """
    Consolidate projections from multiple venues over a common date range.
    The consolidated data is transposed with months as columns and metrics as rows.
    """
    if not venues_data:
        return pd.DataFrame()

    # Initialize a DataFrame for consolidated data
    consolidated = pd.DataFrame()

    # Collect all unique metrics
    all_metrics = set()
    for venue_data in venues_data:
        if not venue_data.empty and 'Metric' in venue_data.columns:
            all_metrics.update(venue_data['Metric'].unique())

    # Create a DataFrame with metrics as rows and months as columns
    consolidated['Metric'] = list(all_metrics)
    for date in all_dates:
        month_year = date.strftime('%b-%Y')
        consolidated[month_year] = 0.0

    # Sum up data from all venues
    for venue_data in venues_data:
        if not venue_data.empty:
            for metric in all_metrics:
                if metric in venue_data['Metric'].values:
                    metric_row = venue_data.loc[venue_data['Metric'] == metric]
                    for col in metric_row.columns[1:]:  # Skip 'Metric' column
                        if col in consolidated.columns:
                            consolidated.loc[consolidated['Metric'] == metric, col] += metric_row[col].values[0]

    return consolidated

def save_to_excel_for_download(venue_projections, consolidated_projection, sport):
    """
    Generates the Excel file in memory and returns it as bytes for download.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_filename = f"{sport}_projection_{timestamp}.xlsx"
    excel_buffer = io.BytesIO()

    try:
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            for i, (venue_data, period) in enumerate(venue_projections):
                sheet_name = f"Venue {i+1} ({period} Months)"
                venue_data.to_excel(writer, sheet_name=sheet_name, index=False)

            if consolidated_projection is not None and not consolidated_projection.empty:
                consolidated_projection.to_excel(writer, sheet_name='Consolidated', index=False)
            else:
                st.warning("Consolidated projection is empty and will not be included in the download.")

        excel_buffer.seek(0)
        return excel_buffer, base_filename
    except Exception as e:
        st.error(f"Error creating Excel file for download: {e}")
        return None, None

def main():
    st.title("Financial Projection Tool")

    uploaded_file = st.file_uploader("Upload your Excel file with projection data", type=["xlsx", "xls"])
    projections = load_projections(uploaded_file)

    if projections:
        sport = st.selectbox("Select the sport", options=projections.keys())

        num_venues = st.number_input("Enter the number of venues", min_value=1, step=1, value=1)

        venues_data_with_period = []
        all_start_dates = []
        max_projection_period = 0

        st.subheader("Venue Configurations")
        for i in range(num_venues):
            st.markdown(f"### Venue {i+1}")
            col1, col2 = st.columns(2)
            with col1:
                start_date = st.date_input(f"Start date for venue {i+1}", key=f"start_date_{i}")
            with col2:
                projection_period = st.number_input(f"Projection time period (in months) for venue {i+1}", min_value=1, step=1, value=12, key=f"period_{i}")

            try:
                venue_projection = generate_venue_projection(
                    start_date,
                    projections[sport],
                    projection_period
                )
                if not venue_projection.empty:
                    venues_data_with_period.append((venue_projection, projection_period))
                    all_start_dates.append(start_date)
                    max_projection_period = max(max_projection_period, projection_period)
                    st.info(f"Projection generated for Venue {i+1}")
                    st.dataframe(venue_projection.head()) # Display a preview
                else:
                    st.warning(f"No projection generated for venue {i+1} due to data issues.")
            except Exception as e:
                st.error(f"Error generating projection for venue {i+1}: {e}")
                return

        if venues_data_with_period:
            st.subheader("Overall Consolidation")
            overall_projection_period = st.number_input("Enter the overall consolidation time period (in months)", min_value=1, step=1, value=24)

            # Determine the common date range for consolidation
            if all_start_dates:
                earliest_start = min(all_start_dates)
                consolidated_end_date = earliest_start + relativedelta(months=overall_projection_period - 1)
                consolidated_date_range = pd.date_range(start=earliest_start, end=consolidated_end_date, freq='MS')
                consolidated_projection = consolidate_projections([data for data, period in venues_data_with_period], consolidated_date_range)

                st.subheader("Consolidated Projection Preview")
                st.dataframe(consolidated_projection.head())

                excel_file, filename = save_to_excel_for_download(venues_data_with_period, consolidated_projection, sport)

                if excel_file:
                    st.download_button(
                        label="Download Projections as Excel",
                        data=excel_file.getvalue(),
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                st.subheader("Projection Summary")
                st.write(f"**Sport:** {sport}")
                st.write(f"**Number of venues:** {num_venues}")
                for i, (data, period) in enumerate(venues_data_with_period):
                    st.write(f"  - **Venue {i+1}:** Projection period - {period} months, Start Date - {data['Month-Year'].iloc[0] if not data.empty else 'N/A'}")
                st.write(f"**Overall consolidation period:** {overall_projection_period} months")

            else:
                st.warning("No start dates available for consolidation.")
        else:
            st.warning("No venue projections were generated. Please check the input and data.")

if __name__ == "__main__":
    main()
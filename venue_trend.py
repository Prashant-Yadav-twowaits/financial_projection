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
    The projection data always starts from 'Month 1' of the base projection.
    """
    end_date = start_date + relativedelta(months=num_months - 1)
    date_range = pd.date_range(start=start_date, end=end_date, freq='MS')
    venue_projection = pd.DataFrame({'Date': date_range})
    venue_projection['Month-Year'] = venue_projection['Date'].dt.strftime('%b-%Y')

    if 'Metric' not in base_projection.columns:
        st.error("Error: 'Metric' column not found in the base projection data.")
        return pd.DataFrame()

    # Ensure no duplicate column names in base_projection
    base_projection = base_projection.loc[:, ~base_projection.columns.duplicated()]

    metric_values = base_projection.set_index('Metric').drop(columns=['Month-Year'], errors='ignore')

    for i in range(1, num_months + 1):
        month_col_name = f"Month {i}"
        if month_col_name in metric_values.columns:
            venue_projection = pd.concat(
                [venue_projection, metric_values[month_col_name].rename(month_col_name)], axis=1
            )
        else:
            st.warning(f"Warning: Could not find column '{month_col_name}' in base projection.")
            venue_projection[month_col_name] = np.nan

    # Transpose the projection to have months as rows and metrics as columns
    transposed_projection = venue_projection.set_index('Month-Year').T.reset_index()
    transposed_projection.rename(columns={'index': 'Metric'}, inplace=True)

    return transposed_projection


def consolidate_projections(venues_data, all_dates):
    """
    Consolidate projections from multiple venues over a common date range.
    The consolidated projection multiplies the values by the number of venues.
    """
    if not venues_data:
        return pd.DataFrame({'Metric': [], 'Month-Year': []})

    # Initialize consolidated DataFrame with metrics and months
    all_metrics = set()
    for venue_data in venues_data:
        if not venue_data.empty and 'Metric' in venue_data.columns:
            all_metrics.update(venue_data['Metric'].unique())

    consolidated = pd.DataFrame({'Metric': list(all_metrics)})

    for date in all_dates:
        month_year = date.strftime('%b-%Y')
        consolidated[month_year] = 0.0

    # Add up projections for all venues
    for venue_data in venues_data:
        if not venue_data.empty:
            for month in consolidated.columns[1:]:  # Skip 'Metric' column
                if month in venue_data.columns:
                    consolidated[month] += venue_data[month].fillna(0)

    # Multiply by the number of venues
    num_venues = len(venues_data)
    consolidated.iloc[:, 1:] *= num_venues

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
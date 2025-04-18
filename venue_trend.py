import pandas as pd
import numpy as np
from datetime import datetime
import calendar
from dateutil.relativedelta import relativedelta
import os

def load_projections(file_path):
    """
    Load the financial projections from the Excel file
    """
    try:
        # Load the Excel file with all sheets
        xl = pd.ExcelFile(file_path)

        # Get all sheet names
        all_sheets = xl.sheet_names
        print(f"Available sheets in the Excel file: {all_sheets}")

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

                        print(f"Columns loaded for {sport}: {sport_data.columns.tolist()}") # Add this line

                        projections[sport] = sport_data
                        print(f"Loaded projection data for {sport} from sheet '{sport}'")
                        found_any_sport = True
                    else:
                        print(f"Warning: Sheet '{sport}' is empty or has insufficient columns.")
                except Exception as e:
                    print(f"Error loading data for {sport}: {e}")

        if found_any_sport:
            return projections
        else:
            print("No projection data found for any sport in the provided Excel file.")
            return None

    except Exception as e:
        print(f"Error loading projections: {e}")
        return None

def generate_venue_projection(start_date_str, base_projection, num_months):
    """
    Generate projection for a single venue starting from a specific date.
    The projection data always starts from 'Month 1' of the base projection.
    """
    try:
        start_date = datetime.strptime(start_date_str, '%m-%Y')
    except ValueError:
        try:
            start_date = datetime.strptime(start_date_str, '%m/%Y')
        except ValueError:
            try:
                start_date = datetime.strptime(start_date_str, '%b-%Y')
            except ValueError:
                try:
                    start_date = datetime.strptime(start_date_str, '%B-%Y')
                except ValueError:
                    raise ValueError("Invalid date format for start date. Please use MM-YYYY, MM/YYYY, Mon-YYYY, or Month-YYYY")

    end_date = start_date + relativedelta(months=num_months - 1)
    date_range = pd.date_range(start=start_date, end=end_date, freq='MS')
    venue_projection = pd.DataFrame({'Date': date_range})
    venue_projection['Month-Year'] = venue_projection['Date'].dt.strftime('%b-%Y')

    if 'Metric' not in base_projection.columns:
        print("Error: 'Metric' column not found in the base projection data.")
        return pd.DataFrame()

    metric_values = base_projection.set_index('Metric').drop(columns=['Month-Year'], errors='ignore')

    for i in range(1, num_months + 1):
        month_col_name = f"Month {i}"
        if month_col_name in metric_values.columns:
            month_data = metric_values[month_col_name].to_dict()
            for metric, value in month_data.items():
                if metric not in venue_projection.columns:
                    venue_projection[metric] = np.nan
                # Ensure the row exists before assigning
                if i - 1 < len(venue_projection):
                    venue_projection.loc[i - 1, metric] = value
        else:
            print(f"Warning: Could not find column '{month_col_name}' in base projection.")
            for metric in metric_values.index:
                if metric not in venue_projection.columns:
                    venue_projection[metric] = np.nan
                if i - 1 < len(venue_projection):
                    venue_projection.loc[i - 1, metric] = np.nan

    return venue_projection.reset_index(drop=True)

def consolidate_projections(venues_data, all_dates):
    """
    Consolidate projections from multiple venues over a common date range.
    """
    if not venues_data:
        return pd.DataFrame({'Month-Year': pd.to_datetime([])})

    consolidated = pd.DataFrame({'Date': all_dates})
    consolidated['Month-Year'] = consolidated['Date'].dt.strftime('%b-%Y')

    all_metrics = set()
    for venue_data in venues_data:
        if not venue_data.empty and 'Metric' in venue_data.columns:
            all_metrics.update(venue_data['Metric'].unique())
        else:
            for col in venue_data.columns:
                if col not in ['Date', 'Month-Year']:
                    all_metrics.add(col)

    for metric in all_metrics:
        consolidated[metric] = 0.0

    for venue_data in venues_data:
        if not venue_data.empty:
            merged = pd.merge(consolidated, venue_data, on='Date', how='left')
            for metric in all_metrics:
                matching_cols = [col for col in merged.columns if metric.lower() in col.lower()]
                if matching_cols:
                    consolidated[metric] += merged[matching_cols].sum(axis=1, min_count=1).fillna(0)

    return consolidated.drop(columns=['Date'])

def save_to_excel(venue_projections, consolidated_projection, sport):
    """
    Save individual venue projections and the consolidated projection to an Excel file
    with a unique filename based on sport and timestamp.
    Each venue will have its own sheet, and there will be a 'Consolidated' sheet.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_filename = f"{sport}_projection_{timestamp}.xlsx"

    try:
        with pd.ExcelWriter(base_filename) as writer:
            for i, (venue_data, period) in enumerate(venue_projections):
                sheet_name = f"Venue {i+1} ({period} Months)"
                venue_data.to_excel(writer, sheet_name=sheet_name, index=False)

            if consolidated_projection is not None and not consolidated_projection.empty:
                consolidated_projection.to_excel(writer, sheet_name='Consolidated', index=False)
            else:
                print("Warning: Consolidated projection is empty and will not be saved.")

        print(f"Successfully saved projections to {base_filename}")
        return True
    except Exception as e:
        print(f"Error saving to Excel: {e}")
        return False

def main():
    """
    Main function to run the projection generation process with individual venue periods.
    """
    file_path = input("Enter the path to the Excel file with projections: ")

    projections = load_projections(file_path)
    if not projections:
        print("Could not load projections. Please check the file path and format.")
        return

    sport = input("Enter the sport (badminton/football/pickleball): ").lower()
    if sport not in projections:
        print(f"No projection data found for {sport}")
        return

    try:
        num_venues = int(input("Enter the number of venues: "))
        if num_venues <= 0:
            print("Number of venues must be positive")
            return
    except ValueError:
        print("Invalid number of venues")
        return

    venues_data_with_period = []
    all_start_dates = []
    max_projection_period = 0

    for i in range(num_venues):
        start_date = input(f"Enter start date for venue {i+1} (MM-YYYY): ")
        try:
            projection_period = int(input(f"Enter the projection time period (in months) for venue {i+1}: "))
            if projection_period <= 0:
                print("Projection period must be positive")
                return
        except ValueError:
            print("Invalid projection period")
            return

        try:
            venue_projection = generate_venue_projection(
                start_date,
                projections[sport],
                projection_period
            )
            if not venue_projection.empty:
                venues_data_with_period.append((venue_projection, projection_period))
                start_date_dt = datetime.strptime(start_date, '%m-%Y')
                all_start_dates.append(start_date_dt)
                max_projection_period = max(max_projection_period, projection_period)
            else:
                print(f"Warning: No projection generated for venue {i+1} due to data issues.")
        except Exception as e:
            print(f"Error generating projection for venue {i+1}: {e}")
            return

    if not venues_data_with_period:
        print("No venue projections were generated. Please check the input and data.")
        return

    # Get the overall consolidation period
    try:
        overall_projection_period = int(input("Enter the overall consolidation time period (in months): "))
        if overall_projection_period <= 0:
            print("Overall projection period must be positive")
            return
    except ValueError:
        print("Invalid overall projection period")
        return

    # Determine the common date range for consolidation
    if all_start_dates:
        earliest_start = min(all_start_dates)
        consolidated_end_date = earliest_start + relativedelta(months=overall_projection_period - 1)
        consolidated_date_range = pd.date_range(start=earliest_start, end=consolidated_end_date, freq='MS')
        consolidated_projection = consolidate_projections([data for data, period in venues_data_with_period], consolidated_date_range)
    else:
        consolidated_projection = pd.DataFrame()

    # Save to Excel with individual sheets and a consolidated sheet
    output_filename = f"{sport}_multi_venue_projection.xlsx" # You can remove this line
    save_to_excel(venues_data_with_period, consolidated_projection, sport) # Pass the 'sport' variable


    # Summary
    print("\nProjection Summary:")
    print(f"Sport: {sport}")
    print(f"Number of venues: {num_venues}")
    for i, (data, period) in enumerate(venues_data_with_period):
        print(f"  Venue {i+1}: Projection period - {period} months, Start Date - {data['Month-Year'].iloc[0] if not data.empty else 'N/A'}")
    print(f"Overall consolidation period: {overall_projection_period} months")
    print(f"Output file: {output_filename}")
    #testing fiinalllll

if __name__ == "__main__":
    main()
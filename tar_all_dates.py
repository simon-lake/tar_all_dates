import pandas as pd
from scipy.stats import linregress, pearsonr
import os
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import PatternFill
import time
import glob
import numpy as np


# Start timer
start_time = time.time()

# List of planting dates to process
planting_dates = ['01j166', '01k166', '01l166', '01a166']  # Add as needed

# Root directories for input, output, and code
root_input_directory = 'C:/AquaCrop_crop_output/bam_t54_snew_two_layer/bam_dry'
root_output_directory = 'C:/AquaCrop_crop_output/bam_t54_snew_two_layer/bam_dry'
root_code_directory = 'C:/MSC_PYTHON_CODES/bam_t54_snew_two_layer/bam_dry/all_bambara'

# List of files to include (same across dates, adjust as needed)
files_to_include_template = [
    'obstmp_195099_bamdry_{date}_wue_ave_014.csv',
    'obstmp_195099_bamdry_{date}_bac_pcv_014.csv',
    'obstmp_195099_bamdry_{date}_rfl_ave_014.csv',
    'obstmp_195099_bamdry_{date}_yld_ave_014.csv',
    'obstmp_195099_bamdry_{date}_yld_pcv_014.csv',
    'obstmp_195099_bamdry_{date}_sst_ave_014.csv',
    'obstmp_195099_bamdry_{date}_ste_ave_014.csv',
    'obstmp_195099_bamdry_{date}_sle_ave_014.csv',
    'obstmp_195099_bamdry_{date}_evp_ave_014.csv',
    'obstmp_195099_bamdry_{date}_tra_ave_014.csv',
    'obstmp_195099_bamdry_{date}_trm_ave_014.csv',
    'obstmp_195099_bamdry_{date}_erm_ave_014.csv',
    'obstmp_195099_bamdry_{date}_bmx_ave_014.csv',
    'obstmp_195099_bamdry_{date}_bmx_pcv_014.csv',
    'obstmp_195099_bamdry_{date}_len_ave_014.csv',
    'obstmp_195099_bamdry_{date}_etc_ave_014.csv',
    'obstmp_195099_bamdry_{date}_wue_pcv_014.csv',
    'obstmp_195099_bamdry_{date}_cfa_sum_014.csv',
    'obstmp_195099_bamdry_{date}_yld_obs_014.csv',
    'obstmp_195099_bamdry_{date}_map_ave_014.csv',
    'obstmp_195099_bamdry_{date}_bac_ave_014.csv'
]

# Process each planting date
for date in planting_dates:
    print(f"Processing date: {date}")

    # Format file names for the current date
    files_to_include = [file_template.format(
        date=date) for file_template in files_to_include_template]

    # Input and output directories for the current date
    input_directory = os.path.join(root_input_directory, date, 'input')
    output_directory = os.path.join(root_output_directory, date, 'output')
    os.makedirs(output_directory, exist_ok=True)

    # Get the list of files in the input directory
    csv_files = glob.glob(os.path.join(input_directory, '*.csv'))

    # Filter the files based on the list of files to include
    filtered_csv_files = [
        file for file in csv_files if os.path.basename(file) in files_to_include]

    # Initialize a combined DataFrame
    combined_data = pd.DataFrame()

    # Load the data from the files
    for file_path in filtered_csv_files:
        if file_path == filtered_csv_files[0]:
            df = pd.read_csv(file_path)
        else:
            df = pd.read_csv(file_path, usecols=[1])
        combined_data = pd.concat([combined_data, df], axis=1)

    # Calculate RCF and add it to the DataFrame
    rcf_ave = []
    cfa_sum = combined_data.iloc[:, 5]
    yld_obs = combined_data.iloc[:, 20]
    for cfa, yld in zip(cfa_sum, yld_obs):
        if cfa == -999:
            rcf_ave.append(-999)
        else:
            adjusted_cfa = cfa + (49 - yld) if yld <= 48 else cfa
            rcf_ave.append((adjusted_cfa / 49) * 100)
    combined_data['RCF'] = rcf_ave

    # Rename the columns
    combined_data.columns = ['sub_cat', 'BAC (t ha^-1)', 'BAC_CV (%)', 'BMX (%)', 'BMX_CV (%)', 'CFA (sum)', 'Es/Esm (%)', 'ETC (%)', 'EVP (mm)', 'CROP CYCLE (days)', 'MAP (mm)', 'RFL (mm)', 'LEAF STRESS (%)', 'STO. STRESS (%)',
                             'TEMP. STRESS (%)', 'TRA (mm)', 'ET/ETm (%)', 'CWP (kg m^-3)', 'CWP CV (%)', 'YIELD (t ha^-1)', 'YLD (obs)', 'YIELD CV (%)', 'RCF (%)']

    # Write combined data to an Excel file
    combined_data_output_path = os.path.join(
        output_directory, 'combined_data.xlsx')
    combined_data.to_excel(combined_data_output_path, index=False)

    # Filter rows and remove specified columns
    combined_data_filtered = combined_data[(combined_data != -999).all(axis=1)]
    combined_data_filtered = combined_data_filtered.drop(
        columns=[combined_data.columns[5]])  # Remove "CFA (sum)"

    # Write filtered data to an Excel file
    filtered_data_output_path = os.path.join(
        output_directory, 'combined_data_filtered.xlsx')
    combined_data_filtered.to_excel(filtered_data_output_path, index=False)

    print(f"Data processing for {date} completed.\n")

# Iterate through each planting date
for date in planting_dates:
    # Set the output directory for the current planting date
    output_directory = os.path.join(root_output_directory, date, 'output')

    # Specify the full path to the Excel file for filtered data
    excel_file_path = os.path.join(
        output_directory, 'combined_data_filtered.xlsx')

    # Load the Excel file
    data = pd.read_excel(excel_file_path)

    # Get the column names
    columns = data.columns

    # Create a dictionary to store the R-squared values
    statistics = {}

    # Iterate through each pair of columns and calculate R-squared
    for i in range(1, len(columns)):
        for j in range(i + 1, len(columns)):
            x = data[columns[i]]
            y = data[columns[j]]

            # Check if any of the variables have all 0 values
            if (x == 0).all() or (y == 0).all():
                # Handle the case where one of the variables is all 0
                r_squared = None
                pearson_corr = None
            else:
                # Perform linear regression
                slope, intercept, r_value, p_value, std_err = linregress(x, y)

                # Calculate R-squared
                r_squared = r_value**2

                # Calculate Pearson correlation coefficient
                pearson_corr, _ = pearsonr(x, y)
                pearson_corr = abs(pearson_corr)

            # Store the R-squared and Pearson correlation coefficient values in the dictionary
            key = f"{columns[i]} vs {columns[j]}"
            statistics[key] = {
                'R-squared': r_squared,
                'Pearson Correlation': pearson_corr
            }

            # Check if Pearson correlation coefficient is greater than 0.8, then plot a scatter chart
            if pearson_corr is not None and pearson_corr > 0.8:
                plt.scatter(x, y)
                plt.xlabel(columns[i])
                plt.ylabel(columns[j])

                # Plot the linear regression line
                plt.plot(x, slope * x + intercept, color='black', linestyle='dotted',
                         label=f'y = {slope:.2f}x + {intercept:.2f}, R2: {r_squared:.2f}')

                # Display legend
                legend_text = (f'Equation: y = {slope:.2f}x + {intercept:.2f}\n'
                               f'R²: {r_squared:.2f}, Pearson r: {pearson_corr:.2f}')
                plt.legend([legend_text], loc='upper left', handlelength=0)
                plt.grid()

                plt.ylim(bottom=0)

                # Save the PNG file with the relationship name in the output directory
                relationship_name = f"{columns[i]}_vs_{columns[j]}".replace(
                    '/', '_').replace(' ', '_')
                file_path = os.path.join(
                    output_directory, f"{relationship_name}_scatter_plot.png")
                plt.savefig(file_path)

                # Close the current figure to avoid overlap with the next plot
                plt.close()

    # Create a dataframe from the R-squared and Pearson correlation coefficient values
    df_r_squared = pd.DataFrame.from_dict(statistics, orient='index')

    # Save the results to an Excel file
    output_filename = "r_squared_values_with_equation.xlsx"
    output_filepath = os.path.join(output_directory, output_filename)
    df_r_squared.to_excel(output_filepath, index=True)

    # Open the Excel file using openpyxl
    wb = openpyxl.load_workbook(output_filepath)
    ws = wb.active

    # Iterate through each row and check the Pearson correlation value
    # Start from row 2 (header is in row 1)
    for row_idx, row in enumerate(df_r_squared.itertuples(), start=2):
        pearson_corr = row._2

        # Check if Pearson correlation is greater than 0.8
        if pearson_corr is not None and pearson_corr > 0.8:
            # Highlight the entire row in yellow
            for col_idx in range(1, len(df_r_squared.columns) + 2):
                ws.cell(row=row_idx, column=col_idx).fill = PatternFill(
                    start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Save the modified Excel file
    wb.save(output_filepath)

    # Print a message indicating that the operation is done
    print(
        f"Data stats completed for {date}. The correlation stats are saved to '{output_filename}'.")

# Print completion message
print("All planting dates processed.")

# Print the total execution time
print(f"Total execution time: {time.time() - start_time:.2f} seconds.")


# Iterate through each planting date
for date in planting_dates:
    # Define the input directory for the current planting date
    csv_directory = os.path.join(root_input_directory, date, 'input')

    # Define the path to the BMX CSV file
    bmx_file_path = os.path.join(
        csv_directory, f'obstmp_195099_bamdry_{date}_bmx_ave_014.csv')

    # Load the BMX data
    bmx_data = pd.read_csv(bmx_file_path)

    # Calculate the max BMX value
    max_bmx_value = bmx_data.iloc[:, 1].max()

    # Calculate the bmx_nom column
    bmx_data['bmx_nom'] = (bmx_data.iloc[:, 1] / max_bmx_value) * 100

    # Drop other columns except for 'sub_cat' and 'bmx_nom'
    bmx_nom_data = bmx_data[['sub_cat', 'bmx_nom']]

    # Define the output directory and file path for the BMX_NOM CSV file
    csv_directory = os.path.join(root_input_directory, date, 'input')
    os.makedirs(csv_directory, exist_ok=True)
    bmx_nom_output_path = os.path.join(
        csv_directory, f'obstmp_195099_bamdry_{date}_bmx_nom_014.csv')

    # Save the BMX_NOM data to a new CSV file
    bmx_nom_data.to_csv(bmx_nom_output_path, index=False)

    # Print a message indicating that the BMX_NOM file has been created
    print(f"BMX_NOM file created and saved to '{bmx_nom_output_path}'.")

# Section 2: Process RCF and Suitability for All Planting Dates
start_time_section2 = time.time()

# Loop through each planting date
for date in planting_dates:
    print(f"Processing RCF and suitability for planting date: {date}")

    # Define CSV directory for the current date
    csv_directory = os.path.join(root_input_directory, date, 'input')

    # Paths to required CSV files
    cfa_sum_file_path = os.path.join(
        csv_directory, f'obstmp_195099_bamdry_{date}_cfa_sum_014.csv')
    yld_obs_file_path = os.path.join(
        csv_directory, f'obstmp_195099_bamdry_{date}_yld_obs_014.csv')

    # Load data from CSV files
    cfa_sum_data = pd.read_csv(cfa_sum_file_path)
    yld_obs_data = pd.read_csv(yld_obs_file_path)

    # Calculate RCF with conditions
    rcf_ave = []
    for cfa, yld in zip(cfa_sum_data.iloc[:, 1], yld_obs_data.iloc[:, 1]):
        if cfa == -999:
            rcf_ave.append(-999)
        else:
            adjusted_cfa = cfa + (49 - yld) if yld <= 48 else cfa
            rcf_value = (adjusted_cfa / 49) * 100
            rcf_ave.append(rcf_value)

    # Create a DataFrame with RCF results
    rcf_ave_data = pd.DataFrame({
        # Assuming sub_cat is the first column
        'sub_cat': cfa_sum_data.iloc[:, 0],
        'rcf_ave': rcf_ave
    })

    # Save RCF results to a new CSV file
    rcf_ave_output_path = os.path.join(
        csv_directory, f'obstmp_195099_bamdry_{date}_rcf_ave_014.csv')
    rcf_ave_data.to_csv(rcf_ave_output_path, index=False)
    print(f"RCF file saved for planting date {date} at {rcf_ave_output_path}")

    # Suitability calculation for selected CSV files
    selected_csv_files = [
        f'obstmp_195099_bamdry_{date}_bmx_ave_014.csv',
        f'obstmp_195099_bamdry_{date}_wue_pcv_014.csv',
        f'obstmp_195099_bamdry_{date}_rfl_ave_014.csv',
        f'obstmp_195099_bamdry_{date}_rcf_ave_014.csv'
    ]

    csv_conditions = {
        f'obstmp_195099_bamdry_{date}_wue_pcv_014.csv': {'condition': lambda x: 0 if x == -999 else (-1 if x > 150 else 1), 'column_name': 'wue_cv_result'},
        f'obstmp_195099_bamdry_{date}_rfl_ave_014.csv': {'condition': lambda x: 0 if x == -999 else (-1 if x < 200 else 1), 'column_name': 'rfl_result'},
        f'obstmp_195099_bamdry_{date}_rcf_ave_014.csv': {'condition': lambda x: 0 if x == -999 else (-1 if x > 33 else 1), 'column_name': 'rcf_result'}
    }

    # Initialize result DataFrame for suitability
    result_df = pd.DataFrame()
    first_column_added = False

    # Process each selected CSV file
    for csv_file in selected_csv_files:
        if csv_file in csv_conditions:
            file_path = os.path.join(csv_directory, csv_file)
            df = pd.read_csv(file_path)

            # Add the first column if not already added
            if not first_column_added:
                result_df['sub_cat'] = df.iloc[:, 0]
                first_column_added = True

            # Apply conditions to calculate suitability results
            result_df[csv_conditions[csv_file]['column_name']] = np.vectorize(
                csv_conditions[csv_file]['condition'])(df.iloc[:, 1])

    # Calculate the 'elim_result' column
    result_df['elim_result'] = result_df.iloc[:, 1:].apply(
        lambda row: -1 if (-1 in row.values) and (0 not in row.values) else (0 if 0 in row.values else (
            1 if all(x == 1 for x in row.values) else None)),
        axis=1
    )

    # Finalize the 'elimination' column
    result_df['elimination'] = np.where(result_df['elim_result'] == 0, 0,
                                        np.where(result_df['elim_result'] == -1, 1, 2))

    # Save the suitability results to the output directory
    suitability_output_path = os.path.join(
        root_output_directory, date, 'output', f'obstmp_195099_bamdry_{date}_suitability_results.csv')
    os.makedirs(os.path.dirname(suitability_output_path),
                exist_ok=True)  # Ensure the output directory exists
    result_df.to_csv(suitability_output_path, index=False)

    print(
        f"Suitability results saved for planting date {date} at {suitability_output_path}")


# Section 2 processing completed
print(
    f"Section 2 processing completed in {time.time() - start_time_section2:.2f} seconds.")

# Section 3: Process BMX and Suitability for All Planting Dates
start_time_section3 = time.time()

for date in planting_dates:
    print(f"Processing BMX and suitability for planting date: {date}")

    # Define directories for the current planting date
    csv_directory = os.path.join(root_input_directory, f"{date}", 'input')
    output_directory = os.path.join(root_output_directory, f"{date}", 'output')
    suitability_file_path = os.path.join(
        output_directory, f"obstmp_195099_bamdry_{date}_suitability_results.csv")

    # Define paths to the required BMX CSV file
    bmx_file_path = os.path.join(
        csv_directory, f"obstmp_195099_bamdry_{date}_bmx_ave_014.csv")

    # Load the BMX data
    bmx_column1 = pd.read_csv(bmx_file_path, usecols=[1], names=['bmx'])
    bmx_column1['bmx'] = pd.to_numeric(bmx_column1['bmx'], errors='coerce')

    # Load the elimination results
    suitability_file_path = os.path.join(
        output_directory, f"obstmp_195099_bamdry_{date}_suitability_results.csv")
    print(f"Looking for suitability file at: {suitability_file_path}")

    result_df = pd.read_csv(suitability_file_path)

    # Merge the 'bmx_column1' DataFrame with the result DataFrame
    result_df = pd.merge(result_df, bmx_column1, left_on='sub_cat',
                         right_index=True, how='left')

    # Apply the condition to filter 'bmx' based on the 'elimination' column
    result_df['bmx'] = np.where((result_df['elimination'] == 0) | (
        result_df['elimination'] == 1), np.nan, result_df['bmx'])

    # Calculate the max value for 'bmx' and normalize
    max_bmx = result_df['bmx'].max()
    result_df['bmx_nom'] = (result_df['bmx'] / max_bmx) * 100

    # Calculate suitability levels
    result_df['S4'] = np.where(result_df['bmx_nom'] < 25, 4, 0)
    result_df['S3'] = np.where(
        (result_df['bmx_nom'] >= 25) & (result_df['bmx_nom'] < 50), 3, 0)
    result_df['S2'] = np.where(
        (result_df['bmx_nom'] >= 50) & (result_df['bmx_nom'] < 75), 2, 0)
    result_df['S1'] = np.where(result_df['bmx_nom'] >= 75, 1, 0)

    # Replace NaN values with 0 in the new columns
    result_df[['S4', 'S3', 'S2', 'S1']] = result_df[[
        'S4', 'S3', 'S2', 'S1']].fillna(0)

    # Calculate final suitability score
    result_df['final_suitability'] = result_df['S4'] + \
        result_df['S3'] + result_df['S2'] + result_df['S1']

    # Save the results to an Excel file
    suitability_output_excel = os.path.join(
        output_directory, 'suitability_analysis_rfl200.xlsx')
    result_df.to_excel(suitability_output_excel, index=False)

    # Save the results to a CSV file
    suitability_output_csv = os.path.join(
        output_directory, 'suitability_analysis_rfl200_all.csv')
    result_df.to_csv(suitability_output_csv, index=False)

    print(
        f"Results written to {suitability_output_excel} and {suitability_output_csv}")

    # Filter columns and save a simplified CSV file
    final_columns = ['sub_cat', 'elimination', 'final_suitability']
    simplified_df = result_df[final_columns]
    simplified_output_path = os.path.join(
        output_directory, 'suitability_analysis_rfl200.csv')
    simplified_df.to_csv(simplified_output_path, index=False)

    print(f"Simplified suitability results saved to {simplified_output_path}")

# Section 3 processing completed
print(
    f"Section 3 processing completed in {time.time() - start_time_section3:.2f} seconds.")

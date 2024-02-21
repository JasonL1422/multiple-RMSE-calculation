import os
import pandas as pd
import numpy as np


def get_excel_column(index):
    """ Generate Excel column label for a given index (1-indexed) """
    column = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        column = chr(65 + remainder) + column
    return column

def extract_simulation_data(file_name, sheet_name, test_ranges):
    simulation_data = {}
    for test_name, (ol_range, ot_range) in test_ranges.items():
        # Ensure the first row is not skipped and the headers are correctly identified
        ol_data = pd.read_excel(file_name, sheet_name=sheet_name, usecols=ol_range, nrows=8, header=None)
        ot_data = pd.read_excel(file_name, sheet_name=sheet_name, usecols=ot_range, nrows=8, header=None)

        # Assuming the first row is data and not a header, fill NaNs with zeros
        ol_data.fillna(0, inplace=True)
        ot_data.fillna(0, inplace=True)

        # Flatten the data and store
        simulation_data[test_name] = {'OL': ol_data.to_numpy().flatten(), 'OT': ot_data.to_numpy().flatten()}
    
    return simulation_data

def calculate_rmse(experimental, simulation):
    # Ensure simulation data matches the size of the experimental data before calculation
    if len(simulation) > len(experimental):
        simulation = simulation[:len(experimental)]  # Trim simulation data to match experimental data size
    return np.sqrt(np.mean((experimental - simulation) ** 2))

# Your experimental data points for L and T
experimental_OL = np.array([0.1234, 0.6543, 0.9876, 1.6789, 2.4567, 3.1234, 4.4567, 4.9876])
experimental_OT = np.array([0.1234, 0.4567, 0.8765, 1.3456, 1.9876, 2.8765, 3.9876, 5.1234])


# Generate test ranges for 36 tests [[[You should change this number]]]
test_ranges = {}
for i in range(36):
    test_name = f'test{i + 1}'
    ol_column = get_excel_column(i * 4 + 2)  # Get column label for OL
    ot_column = get_excel_column(i * 4 + 4)  # Get column label for OT
    test_ranges[test_name] = (f'{ol_column}:{ol_column}', f'{ot_column}:{ot_column}')

file_name = 'sample.xlsx'

# List of sheet names to process
sheet_names = ['rmse1']

# Dictionary to hold RMSE DataFrames for each sheet
all_sheets_rmse = {}

# Loop over each sheet name
for sheet_name in sheet_names:
    simulation_data = extract_simulation_data(file_name, sheet_name, test_ranges)
    
    rmse_data = []
    for test_name, test_data in simulation_data.items():
        # Initialize a dictionary to hold RMSE values for different lengths
        rmse_entry = {'Test Name': test_name}
        for length in [6, 7, 8]:
            experimental_OL_trimmed = experimental_OL[:length]
            experimental_OT_trimmed = experimental_OT[:length]
            rmse_ol = calculate_rmse(experimental_OL_trimmed, test_data['OL'][:length])
            rmse_ot = calculate_rmse(experimental_OT_trimmed, test_data['OT'][:length])
            
            # Store RMSE values in separate columns for each length
            rmse_entry[f'OL RMSE (Length {length})'] = rmse_ol
            rmse_entry[f'OT RMSE (Length {length})'] = rmse_ot
            rmse_entry[f'Average RMSE (Length {length})'] = (rmse_ol + rmse_ot) / 2

        # Append the entry for the current test to the rmse_data list
        rmse_data.append(rmse_entry)

    # Convert RMSE data into a DataFrame and store it in the dictionary
    df_rmse = pd.DataFrame(rmse_data)
    all_sheets_rmse[sheet_name] = df_rmse

# Exporting data
# Use the earlier approach to export, but ensure filenames or sheet names reflect the different lengths
try:
    script_name = os.path.splitext(os.path.basename(__file__))[0]
except NameError:
    script_name = "default_script_name"  # Fallback script name

for sheet_name, df in all_sheets_rmse.items():
    output_file_name = f'{script_name}_{sheet_name}_result111315.xlsx'
    with pd.ExcelWriter(output_file_name) as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"RMSE values for sheet '{sheet_name}' exported to '{output_file_name}'.")

import pandas as pd
import os
import glob

def aggregate_data(input_file, output_file, sheet_name='Mapping'):
    """
    Aggregates data from an input Excel file and saves the result to an output file.
    
    Parameters:
        input_file (str): Path to the input Excel file.
        output_file (str): Path to the output Excel file.
        sheet_name (str): Sheet name to process from the input file.
    """
    # Load the data
    print(f"Loading data from {input_file}...")
    data = pd.read_excel(input_file, sheet_name=sheet_name)
    
    # Clean column names
    data.columns = data.columns.str.strip()
    
    # Group and aggregate the data
    print("Aggregating data...")
    aggregated_data = data.groupby(['Tax Map Parcel ID'], dropna=False).agg({
        'Name': lambda x: ';'.join(x.dropna().unique()),  # Concatenate unique names
        'Acres in Unit': lambda x: ','.join(map(str, x.dropna().unique())),  # Concatenate unique acre values
        'Gross acres': lambda x: ','.join(map(str, x.dropna().unique()))  # Concatenate unique gross acre values
    }).reset_index()
    
    # Rename columns to match desired output
    aggregated_data.rename(columns={
        'Name': 'Names',
        'Acres in Unit': 'Acres in Unit',
        'Gross acres': 'Gross Acres'
    }, inplace=True)
    
    # Save the aggregated data to a new Excel file
    print(f"Saving aggregated data to {output_file}...")
    aggregated_data.to_excel(output_file, index=False)
    print(f"Aggregation complete for {input_file}!")

if __name__ == "__main__":
    # Directory containing input files
    input_directory = "input_folder"  # Replace with your input folder path
    output_directory = "output_folder"  # Replace with your output folder path

    # Create output directory if it doesn't exist
    os.makedirs(output_directory, exist_ok=True)

    # Get all Excel files from the input directory
    input_files = glob.glob(os.path.join(input_directory, "*.xlsx"))

    if not input_files:
        print("No Excel files found in the input directory.")
    else:
        # Process each file
        for input_file in input_files:
            # Generate output file name
            file_name = os.path.basename(input_file)
            output_file = os.path.join(output_directory, f"aggregated_{file_name}")
            
            try:
                aggregate_data(input_file, output_file)
            except Exception as e:
                print(f"Error processing {input_file}: {e}")

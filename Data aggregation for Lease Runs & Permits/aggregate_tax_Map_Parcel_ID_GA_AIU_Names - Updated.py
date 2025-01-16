import pandas as pd

def aggregate_data(input_file, output_file, sheet_name='Mapping'):
    """
    Aggregates data from an input Excel file and saves the result to an output file.
    
    Parameters:
        input_file (str): Path to the input Excel file.
        output_file (str): Path to the output Excel file.
        sheet_name (str): Sheet name to process from the input file.
    """
    # Load the data
    print("Loading data...")
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
    print("Aggregation complete!")

# Example usage
if __name__ == "__main__":
    input_path = "Indigo CBR A.xlsx"  # Replace with your input file path
    output_path = input_path+"-output1.xlsx"  # Replace with your output file path
    aggregate_data(input_path, output_path)

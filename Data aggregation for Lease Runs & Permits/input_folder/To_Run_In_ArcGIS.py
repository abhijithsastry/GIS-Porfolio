import pandas as pd
import os
import glob
import arcpy

def aggregate_data(input_file, output_file):
    """
    Aggregates data from an input Excel file and saves the result to an output file.
    
    Parameters:
        input_file (str): Path to the input Excel file.
        output_file (str): Path to the output Excel file.
    """
    arcpy.AddMessage(f"Loading sheet names from {input_file}...")
    xl = pd.ExcelFile(input_file)
    
    if 'Mapping' in xl.sheet_names:
        sheet_name = 'Mapping'
    elif 'Lead List' in xl.sheet_names:
        sheet_name = 'Lead List'
    else:
        arcpy.AddError(f"Neither 'Mapping' nor 'Lead List' sheet is present in {input_file}.")
        return

    arcpy.AddMessage(f"Using '{sheet_name}' sheet for processing.")
    data = xl.parse(sheet_name)
    data.columns = data.columns.str.strip()

    if 'Tax Map Parcel ID' in data.columns:
        group_column = 'Tax Map Parcel ID'
    elif 'TPIN' in data.columns:
        group_column = 'TPIN'
    else:
        arcpy.AddError(f"Neither 'Tax Map Parcel ID' nor 'TPIN' column is present in {input_file}.")
        return

    arcpy.AddMessage(f"Using '{group_column}' as the grouping column.")

    if 'Acres in Unit' not in data.columns:
        arcpy.AddWarning(f"'Acres in Unit' column is missing in {input_file}. Filling with empty strings.")
        data['Acres in Unit'] = ""

    if 'Gross acres' not in data.columns:
        arcpy.AddWarning(f"'Gross acres' column is missing in {input_file}. Filling with empty strings.")
        data['Gross acres'] = ""

    arcpy.AddMessage("Aggregating data...")
    aggregated_data = data.groupby([group_column], dropna=False).agg({
        'Name': lambda x: ';'.join(x.dropna().unique()),
        'Acres in Unit': lambda x: ','.join(map(str, x.dropna().unique())),
        'Gross acres': lambda x: ','.join(map(str, x.dropna().unique()))
    }).reset_index()

    aggregated_data.rename(columns={
        'Name': 'Names',
        'Acres in Unit': 'Acres in Unit',
        'Gross acres': 'Gross Acres'
    }, inplace=True)

    arcpy.AddMessage(f"Saving aggregated data to {output_file}...")
    aggregated_data.to_excel(output_file, index=False)
    arcpy.AddMessage(f"Aggregation complete for {input_file}!")

def join_to_shapefile(shapefile, excel_file, group_column, output_layer):
    """
    Joins the aggregated Excel data to a shapefile based on a common column and validates the join.
    
    Parameters:
        shapefile (str): Path to the shapefile.
        excel_file (str): Path to the aggregated Excel file.
        group_column (str): Column used for the join (e.g., 'Tax Map Parcel ID').
        output_layer (str): Path to save the output layer.
    """
    arcpy.AddMessage(f"Joining data from {excel_file} to {shapefile}...")

    # Convert Excel to a table
    excel_table = os.path.splitext(excel_file)[0] + ".dbf"
    arcpy.conversion.ExcelToTable(excel_file, excel_table)

    # Perform the join
    arcpy.management.AddJoin(shapefile, group_column, excel_table, group_column)
    arcpy.AddMessage("Join operation completed.")

    # Validate the join
    joined_fields = [f.name for f in arcpy.ListFields(shapefile)]
    required_fields = ['Names', 'Acres in Unit', 'Gross Acres']

    missing_fields = [field for field in required_fields if field not in joined_fields]
    if missing_fields:
        arcpy.AddError(f"Join validation failed. Missing fields: {', '.join(missing_fields)}.")
        return

    arcpy.AddMessage(f"Join validation passed. Exporting to {output_layer}...")

    # Save the joined layer to a new file
    arcpy.management.CopyFeatures(shapefile, output_layer)
    arcpy.AddMessage(f"Joined data saved to {output_layer}.")

if __name__ == "__main__":
    arcpy.env.workspace = arcpy.GetParameterAsText(0)  # Input workspace for Excel files
    shapefile = arcpy.GetParameterAsText(1)  # Input shapefile
    output_directory = arcpy.GetParameterAsText(2)  # Output directory for Excel files and layers
    output_layer = arcpy.GetParameterAsText(3)  # Output joined shapefile path
    group_column = arcpy.GetParameterAsText(4)  # Join column, e.g., 'Tax Map Parcel ID'

    os.makedirs(output_directory, exist_ok=True)
    input_files = glob.glob(os.path.join(arcpy.env.workspace, "*.xlsx"))

    if not input_files:
        arcpy.AddError("No Excel files found in the input directory.")
    else:
        for input_file in input_files:
            file_name = os.path.basename(input_file)
            output_file = os.path.join(output_directory, f"aggregated_{file_name}")

            try:
                # Aggregate the data
                aggregate_data(input_file, output_file)

                # Join the output to the shapefile
                join_to_shapefile(shapefile, output_file, group_column, output_layer)
            except Exception as e:
                arcpy.AddError(f"Error processing {input_file}: {e}")

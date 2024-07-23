from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'

# Ensure the upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file1 = request.files.get('file1')
        file2 = request.files.get('file2')
        if not file1 or not file2:
            return 'No file part or missing files', 400

        filepath1 = os.path.join(app.config['UPLOAD_FOLDER'], file1.filename)
        filepath2 = os.path.join(app.config['UPLOAD_FOLDER'], file2.filename)
        file1.save(filepath1)
        file2.save(filepath2)

        result_file_path = process_files(filepath1, filepath2)
        return send_file(result_file_path, as_attachment=True, download_name='comparison_result.xlsx')

    return render_template('upload.html')

def process_files(filepath1, filepath2):
    # Load both files using pandas, ensuring all data is read as string
    df1 = pd.read_excel(filepath1, dtype=str)
    df2 = pd.read_excel(filepath2, dtype=str)

    # Extract the last 7 digits of the first column in File 1
    column_file1 = df1[df1.columns[0]].str[-7:]

    # Dictionary to hold the results for each comparison
    results = {}

    # Compare the transformed column in File 1 with each column in File 2
    for col in df2.columns:
        # Convert the current column in File 2 to string
        column_file2_full = df2[col]
        # Extract the last 7 digits for comparison
        column_file2_last7 = column_file2_full.str[-7:]
        
        # Find indices where the last 7 digits match
        matching_indices = column_file2_last7.isin(column_file1)
        
        # Get the full numbers from File 2 that match
        matching_full_numbers = column_file2_full[matching_indices].unique()
        results[col] = matching_full_numbers.tolist()  # Store as a list of values

    # Convert the results dictionary to a DataFrame for better Excel output
    result_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in results.items()]))

    # Save the results to a new Excel file with text formatting
    result_file_path = os.path.join('uploads', 'result.xlsx')
    with pd.ExcelWriter(result_file_path, engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False)
        for column_cells in writer.sheets['Sheet1'].columns:
            for cell in column_cells:
                cell.number_format = '@'  # Set format to text

    return result_file_path

if __name__ == '__main__':
    app.run(debug=True)

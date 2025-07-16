from flask import Flask, request, render_template, flash
import os
import pandas as pd

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.secret_key = 'your_secret_key'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def validate_excel(file_path):
    xls = pd.ExcelFile(file_path)
    sheets_needed = {'Transactions', 'Customers', 'Products'}
    sheets_found = set(xls.sheet_names)
    if not sheets_needed <= sheets_found:
        missing = sheets_needed - sheets_found
        raise ValueError(f"Missing sheets: {missing}")
    # Optionally, check columns here
    return True

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['excel_file']
        if file:
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)
            try:
                validate_excel(file_path)
                message = "File validated and uploaded successfully!"
            except Exception as e:
                message = f"Validation error: {e}"
            flash(message)
    return render_template('upload.html')

if __name__ == "__main__":
    app.run(debug=True)
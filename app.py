from flask import Flask, request, render_template, flash, send_file,redirect, url_for
import os
from utils import *
from geo import enrich_geolocation
from log import log_upload

app = Flask(__name__)
app.secret_key = 'your_secret_key'
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('output', exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files.get('excel_file')
        if not file:
            return 'No file uploaded', 400

        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)

        try:
            validate_excel(filepath)
            df_t, df_c, df_p = load_data(filepath)

            changes = detect_address_changes(df_c)
            summary = transaction_summary(df_t, df_p)
            top     = top_spenders(summary)
            ranks   = rank_customers(df_t)

            df_c_geo = enrich_geolocation(df_c)

            log_upload(file.filename, {
                'Transactions': len(df_t),
                'Customers':    len(df_c),
                'Products':     len(df_p)
            })

            excel_path = save_processed_output(
                Transactions=df_t,
                Customers=df_c_geo,
                AddressChanges=changes,
                Summary=summary,
                TopSpenders=top,
                CustomerRanks=ranks
            )
            make_summary_report(summary, top, ranks)

            return send_file(excel_path, as_attachment=True)

        except Exception as e:
            return str(e), 400

    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)


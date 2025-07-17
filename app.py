from flask import Flask, request, render_template, flash, send_file
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
        file = request.files['excel_file']
        if file:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)

            try:
                flash("✅ Upload received. Validating...", 'info')
                validate_excel(filepath)

                flash("📋 File is valid. Loading data...", 'info')
                df_transactions, df_customers, df_products = load_data(filepath)

                flash("📈 Processing data transformations...", 'info')
                changes = detect_address_changes(df_customers)
                summary = transaction_summary(df_transactions, df_products)
                top = top_spenders(summary)
                ranks = rank_customers(df_transactions)

                flash("📍 Looking up geolocations (may take a few seconds)...", 'info')
                df_customers_with_geo = enrich_geolocation(df_customers)

                flash("🗃️ Logging upload into DB...", 'info')
                sheet_counts = {
                    'Transactions': len(df_transactions),
                    'Customers': len(df_customers),
                    'Products': len(df_products)
                }
                log_upload(file.filename, sheet_counts)

                flash("📤 Saving cleaned output and summary...", 'info')
                excel_path = save_processed_output(
                    Transactions=df_transactions,
                    Customers=df_customers_with_geo,
                    AddressChanges=changes,
                    Summary=summary,
                    TopSpenders=top,
                    CustomerRanks=ranks
                )
                make_summary_report(summary, top, ranks)

                flash("🎉 Done! Your file has been processed.", 'success')
                return send_file(excel_path, as_attachment=True)

            except Exception as e:
                flash(f"❌ Error: {e}", 'error')

    return render_template('upload.html')


if __name__ == '__main__':
    app.run(debug=True)


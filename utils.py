import pandas as pd
from docx import Document

# 1. Validate Excel File
def validate_excel(file_path):
    xls = pd.ExcelFile(file_path)
    expected = {'Transactions', 'Customers', 'Products'}
    found = set(xls.sheet_names)
    if not expected <= found:
        raise ValueError(f"Missing sheets: {expected - found}")

# 2. Load Data
def parse_customers(raw_df):
    customer_rows = raw_df.iloc[:, 0]
    parsed = []

    for row in customer_rows:
        clean = row.strip('{}')
        parts = clean.split('_')
        if len(parts) == 6:
            parsed.append(parts)

    df = pd.DataFrame(parsed, columns=[
        'customer_id', 'name', 'email', 'dob', 'address', 'created_date'
    ])
    df['created_date'] = pd.to_numeric(df['created_date'], errors='coerce')
    df['created_date'] = pd.to_datetime(df['created_date'], unit='D', origin='1899-12-30')
    return df

def load_data(file_path):
    xls = pd.ExcelFile(file_path)
    trans = pd.read_excel(xls, 'Transactions')
    products = pd.read_excel(xls, 'Products')
    raw_customers = pd.read_excel(xls, 'Customers', header=None)
    customers = parse_customers(raw_customers)
    return trans, customers, products

# 3a. Address Changes
def detect_address_changes(customers_df):
    customers_df = customers_df.sort_values(['customer_id', 'created_date'])
    customers_df['prev_address'] = customers_df.groupby('customer_id')['address'].shift()
    print(customers_df)
    customers_df['address_changed'] = customers_df['address'] != customers_df['prev_address'] 
    print(customers_df)
    return customers_df[customers_df['address_changed']]

# 3b. Total Spend per Category
def transaction_summary(trans_df, products_df):
    merged = trans_df.merge(products_df[['product_code', 'category']], on='product_code')
    return merged.groupby(['customer_id', 'category'])['amount'].sum().reset_index()

# 3c. Top Spenders
def top_spenders(summary_df):
    return summary_df.loc[summary_df.groupby('category')['amount'].idxmax()]

# 3d. Customer Ranking
def rank_customers(trans_df):
    total = trans_df.groupby('customer_id')['amount'].sum().reset_index()
    total['rank'] = total['amount'].rank(ascending=False, method='min').astype(int)
    return total.sort_values('rank')

# 6. Save Processed Output
def save_processed_output(**dfs):
    path = 'output/processed_output.xlsx'
    with pd.ExcelWriter(path) as writer:
        for name, df in dfs.items():
            df.to_excel(writer, sheet_name=name, index=False)
    return path

# 7. Summary Report
def make_summary_report(summary, top, ranks):
    doc = Document()
    doc.add_heading('Data Summary Report', 0)

    doc.add_paragraph("✅ Total transaction summary per customer per category:")
    for i, row in summary.iterrows():
        doc.add_paragraph(f"{row['customer_id']} - {row['category']}: ${row['amount']:.2f}")

    doc.add_paragraph("\n🏆 Top Spenders:")
    for i, row in top.iterrows():
        doc.add_paragraph(f"{row['category']}: {row['customer_id']} (${row['amount']:.2f})")

    doc.add_paragraph("\n📊 Top Customer Rankings:")
    for i, row in ranks.iterrows():
        doc.add_paragraph(f"Rank {row['rank']}: {row['customer_id']} (${row['amount']:.2f})")

    summary_path = 'output/summary_report.docx'
    doc.save(summary_path)
    return summary_path

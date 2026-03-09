# Trendyol Order Export Tool
![Python](https://img.shields.io/badge/Python-3.x-blue)
Python utility that extracts Trendyol orders without invoice and exports
client billing and shipping addresses into an Excel file.

The script connects to the Trendyol Supplier API, retrieves orders that do
not yet have an invoice attached, extracts customer data and addresses,
and generates an Excel report.

---

# Features

- Trendyol Supplier API integration
- detection of orders without invoice
- extraction of billing and shipping addresses
- postal code mapping to county and locality
- Excel export (.xlsx)
- secure credential storage using Windows DPAPI
- optional packaging as executable

---

# Technologies

- Python
- requests
- pandas
- tkinter
- python-dotenv
- openpyxl

---

# Project Structure

trendyol-order-export

export_trendyol_clients.py  
Main script that connects to the Trendyol API and exports client data.

encrypt_env_dpapi.py  
Utility used to encrypt `.env` credentials using Windows DPAPI.

coduri_postale_judete.json  
Romanian postal code database used to map postal codes to county and locality.

requirements.txt  
Python dependencies required for the project.

README.md  
Project documentation.

.gitignore  
Files excluded from version control.

---

# Installation

Clone the repository

git clone https://github.com/mihaidaneasa/trendyol-order-export.git

Enter the project folder

cd trendyol-order-export

Install dependencies

pip install -r requirements.txt

---

# Configuration

Create a file named `.env` containing your Trendyol API credentials.

Example:

TRENDYOL_ID=your_supplier_id  
TRENDYOL_KEY=your_api_key  
TRENDYOL_SECRET=your_api_secret  
TRENDYOL_REF=TrendyolOblio/1.0

---

# Encrypt credentials (recommended)

For security, credentials can be encrypted using Windows DPAPI.

Run:

python encrypt_env_dpapi.py .env env.dpapi

This will:

- encrypt the `.env` file
- create `env.dpapi`
- optionally delete the plain `.env`

The encrypted file is bound to the current Windows user.

---

# Usage

Run the script:

python export_trendyol_clients.py

The program will:

1. Connect to the Trendyol API
2. Retrieve orders without invoice
3. Extract client billing and shipping addresses
4. Ask where to save the Excel file
5. Export the results to an Excel document

---

# Output

The generated Excel file contains:

OrderNumber  
Client  
Status  

Billing address fields:
Adresa_Facturare  
Localitate_Facturare  
Judet_Facturare  
CodPostal_Facturare  

Shipping address fields:
Adresa_Livrare  
Localitate_Livrare  
Judet_Livrare  
CodPostal_Livrare  

---

# Security

The following files must NOT be uploaded to GitHub:

.env  
env.dpapi  

These files contain API credentials.

---

# Requirements

requirements.txt

requests>=2.31  
pandas>=2.0  
python-dotenv>=1.0  
openpyxl>=3.1  

---

# Author

Mihai Daneasa

Python automation tools for e-commerce workflows.

## AI Assistance

This project was developed with the assistance of AI tools during the implementation phase.

AI was used to help generate parts of the code, while the project design, workflow logic, configuration, testing and real-world usage were defined and validated by the author.

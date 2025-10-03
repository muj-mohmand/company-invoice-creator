#script to create invoices for the company
import os
import pandas as pd
import sys
import random
from datetime import datetime, timedelta
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import LETTER
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

EXCEL_FILE = "BrightDesk_Consulting_Ledger_Mar2022_to_Aug2025_v13.xlsx"

def get_invoice_date_and_due_date(payment_date):
    # If it's already a datetime object, use it directly
    if isinstance(payment_date, datetime):
        payment_date_str = payment_date
    else:
        # Convert string to datetime - try common formats
        payment_date = str(payment_date).split()[0]  # Remove time if present
        try:
            payment_date_str = datetime.strptime(payment_date, "%Y-%m-%d")
        except ValueError:
            payment_date_str = datetime.strptime(payment_date, "%m/%d/%Y")
    
    # Subtract random days (0-30)
    days_to_subtract = random.randint(0, 30)
    invoice_date = payment_date_str - timedelta(days=days_to_subtract)
    
    # Due date is 30 days after invoice date
    due_date = invoice_date + timedelta(days=30)
    
    # Format the dates
    invoice_date_formatted = invoice_date.strftime("%m/%d/%Y")
    due_date_formatted = due_date.strftime("%m/%d/%Y")
    
    return invoice_date_formatted, due_date_formatted


def create_invoices(data, company_info):
    folder = "invoices"
    os.makedirs(folder, exist_ok=True)
    #for each unique reference number, create an invoice
    grouped_df = data.groupby("Reference")
    w, h = LETTER
    # Logo path (adjust to your file location)
    logo_path = "logo.png"
    # Create a style for wrapping text
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']
    normal_style.fontSize = 9
    normal_style.leading = 11  # line spacing

    for ref_number, group in grouped_df:
       
        invoice_number = ref_number
        #create a pdf for each reference number
        invoice_pdf = f"invoice_{invoice_number}.pdf"
        pdf_file = os.path.join(folder, invoice_pdf)
        c = canvas.Canvas(pdf_file, pagesize=LETTER)
        
         # INSERT LOGO HERE
        try:
            # Medium logo
            c.drawImage(logo_path, 50, h - 100, width=200, height=100, preserveAspectRatio=True)
        except:
            pass  # Skip if logo not found
        
     
        # Reset these for each invoice
        address_start_x = 264
        address_start_y = 48
        bill_to_address_start_x = 72
        bill_to_address_start_y = 165
        c.setFont("Helvetica", 10)
        #Insert invoice date, due date, invoice number
        invoice_date, due_date = get_invoice_date_and_due_date(group['Date'].iloc[0])
        c.drawString(bill_to_address_start_x, h - 242, f"Invoice Number: {invoice_number}")
        c.drawString(bill_to_address_start_x, h - 252, f"Invoice Date: {invoice_date}")
        c.drawString(bill_to_address_start_x, h - 262,  f"Due Date: {due_date}")

        #insert company info, company image
        c.setFont("Helvetica-Oblique", 8)

        company_address_fields = [
            'company_name', 'address', 'city', 'province',
            'country', 'postal_code', 'phone'
        ]

        for field in company_address_fields:
            if field == 'city':
                c.drawString(address_start_x, h - address_start_y, str(company_info['city']) + ", " + str(company_info['province']))    
            elif field == "province":
                continue
            else:
                c.drawString(address_start_x, h - address_start_y, str(company_info[field]))
            address_start_y += 10

        #bill to
        c.setFont("Helvetica-Bold", 10)
        c.drawString(bill_to_address_start_x, h - bill_to_address_start_y, "Bill To:")
        bill_to_address_start_y += 10
        
        billing_address_fields = [
            'Payee',
            'Street Address',
            'City',
            'Country',
            'Postal Code'
        ]

        for field in billing_address_fields:
            if field == 'City':
                c.drawString(bill_to_address_start_x, h - bill_to_address_start_y, str(group['City'].iloc[0]) + ',' + str(group['Province/State'].iloc[0]))
            else:
                c.drawString(bill_to_address_start_x, h - bill_to_address_start_y, str(group[field].iloc[0]))
            bill_to_address_start_y += 9

        #insert invoice items
        invoice_table_data = [["Item", "Description", "Quantity", "Unit Price", "Subtotal", "Tax", "Amount"]]
        for _, row in group.iterrows():
            # Format numeric values with commas for thousands and 2 decimal places
            quantity = f"{float(row['Quantity']):,.0f}" if pd.notna(row['Quantity']) else "0"
            unit_price = f"{float(row['Unit Price']):,.2f}" if pd.notna(row['Unit Price']) else "0.00"
            subtotal = f"{float(row['Subtotal']):,.2f}" if pd.notna(row['Subtotal']) else "0.00"
            tax = f"{float(row['Total Tax']):,.2f}" if pd.notna(row['Total Tax']) else "0.00"
            amount = f"{float(row['Amount']):,.2f}" if pd.notna(row['Amount']) else "0.00"
            
            # Use Paragraph for the Description column to enable text wrapping
            description = Paragraph(str(row['Description']), normal_style)
            
            invoice_table_data.append([
                Paragraph(str(row['Item Number']), normal_style), 
                description,  # Paragraph object instead of string
                quantity,
                unit_price,
                subtotal,
                tax,
                amount
            ])
        
        #Make table
        table = Table(invoice_table_data, colWidths=[50, 150, 60, 70, 70, 50, 70])
        style = TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),  # grid lines
            ('BACKGROUND', (0,0), (-1,0), colors.grey),   # header background
            ('ALIGN', (2,1), (2,-1), 'CENTER'),           # Quantity centered
            ('ALIGN', (3,1), (-1,-1), 'RIGHT'),           # Prices right-aligned
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), # bold header
            ('FONTSIZE', (0,0), (-1,-1), 9),              # Smaller font size
        ])

        table.setStyle(style)
        table_width, table_height = table.wrap(0, 0)
        table.wrapOn(c, w, h)
        table.drawOn(c, 50, h - 360 - table_height)

        #Terms and conditions
        # After drawing the table, add terms at the bottom
        terms_y_position = h - 400 - table_height - 40
        terms_str = f"Payment Terms: UNLESS OTHERWISE MUTUALLY AGREED IN WRITING BETWEEN YOU AND {company_info['company_name']}, THE {company_info['company_name']} TERMS OF SALE AND POLICIES GOVERN THIS TRANSACTION"

        # Use Paragraph for the long terms text (it will wrap)
        c.setFont("Helvetica-Bold", 9)
        terms_paragraph = Paragraph(terms_str, normal_style)
        terms_paragraph.wrapOn(c, 500, 100)  # width, height available for wrapping
        terms_paragraph.drawOn(c, 50, terms_y_position)

        # Calculate how much vertical space the paragraph used
        terms_height = terms_paragraph.height

        # Draw the rest below the wrapped paragraph
        c.setFont("Helvetica", 8)
        c.drawString(50, terms_y_position - terms_height - 12, "Payment due within 30 days of invoice date")
        c.drawString(50, terms_y_position - terms_height - 24, "Late payments subject to 1.5% monthly interest")
        c.drawString(50, terms_y_position - terms_height - 36, "All amounts in CAD")

        #save the pdf
        c.save()    
def get_company_info(file):
    df = pd.read_excel(file, sheet_name='company_info', header=None)
    if df.empty:
        raise ValueError("Company info sheet is empty.")
    
    # Data is arranged in two columns: keys in column 0, values in column 1
    company_info = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))
    
    return company_info


def read_excel(file_path):
    """Read the Excel file and return a DataFrame."""
    try:
        df = pd.read_excel(file_path, sheet_name='company_invoice_data')
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

def main():
    try:
        print("Reading Excel file...")
        data = read_excel(EXCEL_FILE)
        print(f"Loaded {len(data)} rows of data")
        
        print("Getting company info...")
        company_info = get_company_info(EXCEL_FILE)
        print(f"Company: {company_info.get('company_name', 'Unknown')}")
        
        print("Creating invoices...")
        create_invoices(data, company_info)
        print("Invoices created successfully!")
        
    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()  # This will show the full error stack
        sys.exit(1)

if __name__ == "__main__":
    sys.exit(main())
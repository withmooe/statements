import os
import pandas as pd
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable, Image

# Load the Excel file
file_path = r'C:\Users\mazao\Documents\SB\statements.xlsx'  # Updated path
data = pd.read_excel(file_path)

# Print the columns to check their names
print("Columns in the DataFrame:", data.columns)

# Rename columns to match the expected names
data.columns = [col.strip() for col in data.columns]  # Remove any leading/trailing whitespace

# Ensure the data columns are of numeric type where appropriate
numeric_columns = ['Sales', 'Royalty_rate', 'Royalties_earned', 'Ingoing_balance', 'Outgoing_balance', 'Payment']
for col in numeric_columns:
    if col in data.columns:
        data[col] = pd.to_numeric(data[col], errors='coerce')
    else:
        print(f"Column '{col}' not found in the DataFrame.")

# Ensure the Downloads/Statements directory exists
output_dir = r'C:\Users\mazao\Documents\SB\Statements'  # Updated path
os.makedirs(output_dir, exist_ok=True)

# Create separate directory for payments above 50 EUR
output_dir_above_50 = os.path.join(output_dir, 'Payments_Above_50')
os.makedirs(output_dir_above_50, exist_ok=True)

logo_path = r'C:\Users\mazao\Documents\SB\SB.png'  # Updated path

def format_number(value):
    if pd.isna(value):
        return "0,00"
    return "{:,.2f}".format(value).replace(",", "X").replace(".", ",").replace("X", ".")

def format_percentage(value):
    if pd.isna(value):
        return "0%"
    return "{:.0%}".format(value).replace(".", ",")

def add_footer(canvas, doc):
    footer_text = "Contact royalty@fake.com for questions"
    canvas.saveState()
    canvas.setFont('Helvetica', 10)
    canvas.drawString(14, 15, footer_text)
    canvas.restoreState()

def build_story(owner_data):
    elements = []
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(name='Heading1', parent=styles['Heading1'], fontName='Helvetica', fontSize=30)
    subtitle_style = ParagraphStyle(name='Subtitle', parent=styles['Normal'], fontName='Helvetica', fontSize=12)
    normal_style = styles['Normal']
    bold_style = ParagraphStyle(name='Bold', parent=normal_style, fontName='Helvetica-Bold', fontSize=14)
    right_align_style = ParagraphStyle(name='RightAlign', parent=normal_style, alignment=2)

    owner_id = owner_data['Copyright_owner_ID'].iloc[0]
    owner_name = owner_data['Copyright_owner'].iloc[0]
    contact = owner_data['Contact'].iloc[0] if pd.notna(owner_data['Contact'].iloc[0]) else "N/A"
    account_number = owner_data['Account_number'].iloc[0] if 'Account_number' in owner_data.columns else "N/A"
    bank_id = owner_data['Bank_ID'].iloc[0] if 'Bank_ID' in owner_data.columns else "N/A"
    statement_id = owner_data['Title_ID'].iloc[0] if 'Title_ID' in owner_data.columns else "N/A"
    date = "2023-01-31"

    # Add logo, title, and subtitle
    logo = Image(logo_path, width=60, height=60)
    header_table = Table([
        [
            Paragraph("<b>Royalty Statement</b>", header_style),
            logo
        ]
    ], colWidths=[None, 60])
    header_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('VALIGN', (0, 0), (1, 0), 'TOP')
    ]))
    elements.append(header_table)

    elements.append(Paragraph("Period 2022-01-01 - 2022-12-31", subtitle_style))
    elements.append(HRFlowable(width="100%", thickness=5, color=colors.black, spaceBefore=10, spaceAfter=10))
    elements.append(Spacer(1, 10))

    # Add a two-column layout for the copyright owner data and company information
    owner_data_table = [
        [
            Paragraph(f"Copyright Owner ID: {owner_id}", normal_style),
            Paragraph("Smart Books", right_align_style)
        ],
        [
            Paragraph(f"Copyright Owner Name: {owner_name}", normal_style),
            Paragraph("Company registration nr: XXXXXXX", right_align_style)
        ],
        [
            Paragraph(f"Contact: {contact}", normal_style),
            Paragraph("VAT: XXXXXXXXX", right_align_style)
        ],
        [
            Paragraph(f"Account number: {account_number}", normal_style),
            Paragraph("Registered for Tax", right_align_style)
        ],
        [
            Paragraph(f"Bank ID: {bank_id}", normal_style),
            ""
        ],
        [
            Paragraph(f"Statement ID: {statement_id}", normal_style),
            ""
        ],
        [
            Paragraph(f"Date: {date}", normal_style),
            ""
        ]
    ]

    # Adjust the column widths to position the company info at the extreme right
    owner_data_table = Table(owner_data_table, colWidths=[300, 500], hAlign='LEFT')
    owner_data_table.setStyle(TableStyle([
        ('LEFTPADDING', (0, 0), (0, -1), 0),
        ('RIGHTPADDING', (1, 0), (1, -1), 0),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))

    elements.append(owner_data_table)
    elements.append(Spacer(1, 20))

    # Adjust the column widths
    column_widths = [85, 120, 120, 60, 80, 80, 85, 90, 80]

    table_data = [['Title ID', 'Title Name', 'Author', 'Sales', 'Royalty Rate', 'Royalty Earned', 'Ingoing Balance', 'Outgoing Balance', 'Payment']]
    for index, row in owner_data.iterrows():
        table_data.append([
            Paragraph(str(row['Title_ID']), normal_style),
            Paragraph(str(row['Title_name']), normal_style),
            Paragraph(str(row['Author']), normal_style),
            Paragraph(format_number(row['Sales']), normal_style),
            Paragraph(format_percentage(row['Royalty_rate']), normal_style),
            Paragraph(format_number(row['Royalties_earned']), normal_style),
            Paragraph(format_number(row['Ingoing_balance']), normal_style),
            Paragraph(format_number(row['Outgoing_balance']), normal_style),
            Paragraph(format_number(row['Payment']) if row['Payment'] > 0 else '', normal_style)
        ])

    table = Table(table_data, colWidths=column_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), 'whitesmoke'),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 1), (-1, -1), 'RIGHT')
    ]))

    elements.append(table)
    elements.append(Spacer(1, 5))
    currency_table = Table([
        [Paragraph("Currency in EUR", normal_style)]
    ], colWidths=[sum(column_widths)])
    currency_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT')
    ]))
    elements.append(currency_table)
    elements.append(Spacer(1, 25))

    total_payment = owner_data['Payment'].sum()
    if total_payment > 0:
        elements.append(Paragraph(f"Total amount to be paid: EUR {format_number(total_payment)}", bold_style))
        if total_payment < 50:
            elements.append(Spacer(1, 10))
            elements.append(Paragraph("Please note that royalties below EUR 50 are not paid due to bank fees being higher, the royalty will be saved as outgoing balance and paid out with the next payment.", normal_style))
            elements.append(Spacer(1, 25))

    return elements

def generate_pdf_reportlab(data):
    unique_owners = data['Copyright_owner_ID'].unique()

    for owner_id in unique_owners:
        owner_data = data[data['Copyright_owner_ID'] == owner_id]
        story = build_story(owner_data)

        owner_name = owner_data['Copyright_owner'].iloc[0]

        # Replace spaces with underscores and remove any invalid characters for filenames
        owner_name_cleaned = owner_name.replace(" ", "_").replace("/", "_")
        statement_filename = f"{owner_name_cleaned}.pdf"

        # Determine output directory based on total payment
        total_payment = owner_data['Payment'].sum()
        if total_payment > 50:
            statement_pdf_path = os.path.join(output_dir_above_50, statement_filename)
        else:
            statement_pdf_path = os.path.join(output_dir, statement_filename)

        # Build the document for each owner
        doc = SimpleDocTemplate(statement_pdf_path, pagesize=landscape(A4), rightMargin=14, leftMargin=14, topMargin=40, bottomMargin=40)
        doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)

        print(f"PDF generated for {owner_name}: {statement_pdf_path}")

generate_pdf_reportlab(data) # type: ignore
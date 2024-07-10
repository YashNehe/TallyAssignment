import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime

def safe_find_text(element, tag, default=''):
    found = element.find(tag)
    return found.text if found is not None else default

def parse_tally_xml(xml_file_path): 
    tree = ET.parse(xml_file_path)
    root = tree.getroot()
    vouchers = []

    for voucher in root.findall('.//VOUCHER'):
        vch_type = safe_find_text(voucher, 'VOUCHERTYPENAME')
        
        if vch_type.lower() != "receipt":
            continue

        date = datetime.strptime(safe_find_text(voucher, 'DATE', '19000101'), '%Y%m%d').strftime('%d-%m-%Y')
        vch_no = safe_find_text(voucher, 'VOUCHERNUMBER')
        ref_no = safe_find_text(voucher, 'REFERENCE', 'NA')
        ref_date = safe_find_text(voucher, 'REFERENCEDATE')
        ref_date = datetime.strptime(ref_date, '%Y%m%d').strftime('%d-%m-%Y') if ref_date else ''
        debtor = safe_find_text(voucher, 'PARTYLEDGERNAME', 'Unknown')

        ledger_entries = voucher.findall('.//ALLLEDGERENTRIES.LIST')
        
        if ledger_entries:
            total_amount = sum(abs(float(safe_find_text(entry, 'AMOUNT', '0'))) for entry in ledger_entries[:-1])
            
            # Parent entry
            parent_entry = {
                'Date': date,
                'Transaction Type': 'Parent',
                'Vch No.': vch_no,
                'Ref No': 'NA',
                'Ref Type': 'NA',
                'Ref Date': 'NA',
                'Debtor': debtor,
                'Ref Amount': 'NA',
                'Amount': total_amount,
                'Particulars': debtor,
                'Vch Type': vch_type,
                'Amount Verified': 'Yes'
            }
            vouchers.append(parent_entry)

            # Child entries
            for entry in ledger_entries[:-1]:
                bill_refs = entry.findall('.//BILLALLOCATIONS.LIST')
                if bill_refs:
                    for bill in bill_refs:
                        child_entry = {
                            'Date': date,
                            'Transaction Type': 'Child',
                            'Vch No.': vch_no,
                            'Ref No': safe_find_text(bill, 'NAME', 'NA'),
                            'Ref Type': 'Agst Ref',
                            'Ref Date': ref_date,
                            'Debtor': debtor,
                            'Ref Amount': abs(float(safe_find_text(bill, 'AMOUNT', '0'))),
                            'Amount': 'NA',
                            'Particulars': debtor,
                            'Vch Type': vch_type,
                            'Amount Verified': 'NA'
                        }
                        vouchers.append(child_entry)

            # Other entry (last ledger entry)
            other_entry = {
                'Date': date,
                'Transaction Type': 'Other',
                'Vch No.': vch_no,
                'Ref No': 'NA',
                'Ref Type': 'NA',
                'Ref Date': 'NA',
                'Debtor': safe_find_text(ledger_entries[-1], 'LEDGERNAME', 'Unknown'),
                'Ref Amount': 'NA',
                'Amount': float(safe_find_text(ledger_entries[-1], 'AMOUNT', '0')),
                'Particulars': safe_find_text(ledger_entries[-1], 'LEDGERNAME', 'Unknown'),
                'Vch Type': vch_type,
                'Amount Verified': 'NA'
            }
            vouchers.append(other_entry)

    return vouchers

# Specify the path to your XML file
xml_file_path = 'xml/input.xml'

# Parse the XML file
vouchers = parse_tally_xml(xml_file_path)

# Create a DataFrame from the vouchers list
df = pd.DataFrame(vouchers)

# Save the DataFrame to an Excel file
df.to_excel('tally_receipt_daybook.xlsx', index=False)

print("Excel file 'tally_receipt_daybook.xlsx' has been created successfully.")
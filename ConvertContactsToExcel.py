# Import required libraries
import vobject
import pandas as pd

# Function to convert VCF to CSV
def vcf_to_csv(vcf_file, output_file):
    # List to store contact data
    contacts = []

    # Read the vCard file
    with open(vcf_file, 'r', encoding='utf-8') as file:
        vcf_data = file.read()
        vcard_entries = vobject.readComponents(vcf_data)

    # Extract data from vCard file
    for vcard in vcard_entries:
        try:
            full_name = f"{vcard.n.value.given} {vcard.n.value.family}" if hasattr(vcard.n.value, 'given') and hasattr(vcard.n.value, 'family') else ''
            emails = ', '.join([email.value for email in vcard.contents.get('email', [])])
            phones = ', '.join([tel.value for tel in vcard.contents.get('tel', [])])
            address = ', '.join([adr.value.street for adr in vcard.contents.get('adr', []) if adr.value.street])

            contact = {
                'Full Name': full_name,
                'Email': emails,
                'Phone': phones,
                'Address': address
            }
            contacts.append(contact)
        except Exception as e:
            print(f"Error processing vCard: {e}")

    # Create a DataFrame
    df = pd.DataFrame(contacts)
    
    # Save DataFrame to CSV
    df.to_csv(output_file, index=False)

# Function to convert VCF to Excel
def vcf_to_excel(vcf_file, output_file):
    # List to store contact data
    contacts = []

    # Read the vCard file
    with open(vcf_file, 'r', encoding='utf-8') as file:
        vcf_data = file.read()
        vcard_entries = vobject.readComponents(vcf_data)

    # Extract data from vCard file
    for vcard in vcard_entries:
        try:
            full_name = f"{vcard.n.value.given} {vcard.n.value.family}" if hasattr(vcard.n.value, 'given') and hasattr(vcard.n.value, 'family') else ''
            emails = ', '.join([email.value for email in vcard.contents.get('email', [])])
            phones = ', '.join([tel.value for tel in vcard.contents.get('tel', [])])
            address = ', '.join([adr.value.street for adr in vcard.contents.get('adr', []) if adr.value.street])

            contact = {
                'Full Name': full_name,
                'Email': emails,
                'Phone': phones,
                'Address': address
            }
            contacts.append(contact)
        except Exception as e:
            print(f"Error processing vCard: {e}")

    # Create a DataFrame
    df = pd.DataFrame(contacts)
    
    # Save DataFrame to Excel
    df.to_excel(output_file, index=False)

# Main function
if __name__ == '__main__':
    # Input vCard file path
    vcf_file = 'pathtofile.vcf'
    
    # Output file paths
    csv_output_file = 'pathtofile.csv'
    excel_output_file = 'pathtofile.xlsx'
    
    # Convert vCard to CSV
    vcf_to_csv(vcf_file, csv_output_file)
    
    # Convert vCard to Excel
    vcf_to_excel(vcf_file, excel_output_file)

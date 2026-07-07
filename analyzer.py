import zipfile
import os
import shutil
import mailbox
import email
import re
from email import policy
from email.utils import parseaddr, parsedate_to_datetime
import pandas as pd

def extract_zip(zip_path, extract_folder):
    """Unzips the archive to access the .mbox file inside."""
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_folder)
    print("-> Successfully unzipped the archive.")

def parse_mbox(extract_folder):
    """Analyzes emails and extracts precise datetime objects for sorting."""
    email_data = []
    mbox_path = None
    
    print("-> Searching for the .mbox file...")
    for root, dirs, files in os.walk(extract_folder):
        for filename in files:
            if filename.lower().endswith(".mbox"):
                mbox_path = os.path.join(root, filename)
                break
        if mbox_path:
            break
            
    if not mbox_path:
        print("-> Error: No .mbox file found inside the zip archive.")
        return email_data
        
    print(f"-> Found mailbox! Fully analyzing each email...")
    mb = mailbox.mbox(mbox_path)
    
    phone_pattern = r'\(?\b[0-9]{3}\)?[-. ]?[0-9]{3}[-. ]?[0-9]{4}\b'
    free_domains = ['gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'aol.com', 'icloud.com']
    
    for i, msg in enumerate(mb):
        parsed_msg = email.message_from_string(msg.as_string(), policy=policy.default)
        
        sender_name, sender_email = parseaddr(str(parsed_msg['From'] or ""))
        receiver_name, receiver_email = parseaddr(str(parsed_msg['To'] or ""))
        
        if not sender_email or not receiver_email:
            continue
            
        company_name = ""
        if '@' in sender_email:
            domain = sender_email.split('@')[1].lower()
            if domain not in free_domains:
                company_name = domain.split('.')[0].capitalize()
                
        # --- Capture exact time for sorting ---
        raw_date = str(parsed_msg['Date'] or "")
        try:
            # Converts the messy email date into a strict Python datetime object
            sortable_time = parsedate_to_datetime(raw_date)
            # Make sure it's "timezone unaware" so Pandas can sort it easily
            sortable_time = sortable_time.replace(tzinfo=None) 
        except Exception:
            sortable_time = pd.NaT # Not-a-Time (Pandas will put these at the bottom)
            
        full_body = ""
        try:
            if parsed_msg.is_multipart():
                for part in parsed_msg.walk():
                    if part.get_content_type() == "text/plain":
                        full_body += str(part.get_content())
            else:
                full_body = str(parsed_msg.get_content() or "")
        except Exception:
            full_body = ""
            
        clean_body = full_body.replace('\n', ' ').replace('\r', ' ')
        phones = re.findall(phone_pattern, clean_body)
        
        contact_detail = phones[0] if len(phones) > 0 else sender_email
        subject = str(parsed_msg['Subject'] or "(No Subject)").replace('\r', '').replace('\n', ' ')
        
        email_data.append({
            "Sortable Time": sortable_time, 
            "Date & Time": "", # We will format this nicely later
            "Sender Name": sender_name,
            "Sender Email": sender_email.lower(),
            "Receiver Name": receiver_name,
            "Receiver Email": receiver_email.lower(),
            "Company Name": company_name,
            "Subject": subject,
            "Extracted Contact Detail": contact_detail
        })
        
        if (i + 1) % 500 == 0:
            print(f"   ...fully analyzed {i + 1} emails")
            
    return email_data

def export_to_excel(data, output_excel_path):
    """Sorts data, formats the date, and deletes all repeated sender/receiver pairs."""
    if not data:
        print("-> Error: No emails could be read.")
        return
        
    print("-> Organizing data and applying strict no-repeat rules...")
    df = pd.DataFrame(data)
    original_count = len(df)
    
    # 1. Sort everything by time (Newest emails at the top)
    df = df.sort_values(by='Sortable Time', ascending=False)
    
    # 2. Format the Date beautifully now that it is sorted
    df['Date & Time'] = df['Sortable Time'].dt.strftime("%B %d, %Y at %I:%M %p")
    
    # 3. Delete the temporary sorting column
    df = df.drop(columns=['Sortable Time'])
    
    # 4. STRICT NO REPEAT: If this Sender and Receiver pair has already appeared, delete it.
    df = df.drop_duplicates(subset=['Sender Email', 'Receiver Email'], keep='first')
    
    final_count = len(df)
    duplicates_removed = original_count - final_count
    
    print(f"-> Removed {duplicates_removed} older emails to ensure names are only mentioned once.")
    
    df.to_excel(output_excel_path, index=False, engine='openpyxl')
    print(f"-> SUCCESS: Saved {final_count} unique interactions to '{output_excel_path}'")

if __name__ == "__main__":
    ZIP_FILE_PATH = "emails.zip"       
    TEMP_FOLDER = "extracted_takeout"   
    EXCEL_OUTPUT = "Unique_Interactions_List.xlsx" 
    
    if not os.path.exists(ZIP_FILE_PATH):
        print(f"-> Error: Could not find '{ZIP_FILE_PATH}' in this folder.")
    else:
        try:
            extract_zip(ZIP_FILE_PATH, TEMP_FOLDER)
            parsed_data = parse_mbox(TEMP_FOLDER)
            if parsed_data:
                export_to_excel(parsed_data, EXCEL_OUTPUT)
        finally:
            if os.path.exists(TEMP_FOLDER):
                shutil.rmtree(TEMP_FOLDER)
                print("-> Cleaned up temporary files to save hard drive space.")
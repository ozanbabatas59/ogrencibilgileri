import os
import re
from pdfminer.high_level import extract_text
from openpyxl import Workbook

def extract_info_from_pdf(pdf_path):
    text = extract_text(pdf_path)
    
    # Adı / Soyadı çıkarma
    ad_soyad_pattern = r'Adı / Soyadı\s*:\s*(.*)'
    ad_soyad_matches = re.findall(ad_soyad_pattern, text)
    ad_soyad = ad_soyad_matches[0].strip() if ad_soyad_matches else None
    
    # T.C. Kimlik No çıkarma
    tc_no_pattern = r'T\.C\. Kimlik No\s*:\s*(\d{11})'
    tc_no_matches = re.findall(tc_no_pattern, text)
    tc_no = tc_no_matches[0] if tc_no_matches else None
    
    return ad_soyad, tc_no

def process_pdfs_in_folder(folder_path, excel_file):
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    wb = Workbook()
    ws = wb.active
    ws.append(['Dosya Adı', 'Adı Soyadı', 'T.C. Kimlik No'])
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        ad_soyad, tc_no = extract_info_from_pdf(pdf_path)
        
        if ad_soyad or tc_no:
            ws.append([pdf_file, ad_soyad, tc_no])
            print(f"{pdf_file} dosyasından bilgiler aktarıldı.")
    
    wb.save(excel_file)
    print(f"Excel dosyası olarak kaydedildi: {excel_file}")

# Kullanım örneği
pdf_folder = 'pdfdosyalari'  # PDF dosyalarının bulunduğu klasör yolu
excel_file = 'pdf_bilgileri.xlsx'  # Excel dosyasının adı ve yolunu belirtin
process_pdfs_in_folder(pdf_folder, excel_file)

import pandas as pd
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
import time

def identify_official_website(nbfc_name, regional_office, address, email_id):
    search_query = f"{nbfc_name} official website"
    search_url = f"https://www.google.com/search?q={search_query}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    try:
        response = requests.get(search_url, headers=headers)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            official_link_tag = soup.find('cite')
            if official_link_tag:
                official_link = official_link_tag.text
                return official_link
        return "Not found"
    except Exception as e:
        print(f"Error finding official website for {nbfc_name}: {str(e)}")
        return "Not found"

def main():
    start_time = time.time()
    file_path = 'nbfc_list.XLSX'
    sheet_name = 'List of NBFCs'
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    print(df.head())
    df_subset = df.head(10)
    df['Official Website'] = ""

    def process_row(index, row):
        nbfc_name = row['NBFC Name']
        regional_office = row['Regional Office']
        address = row['Address']
        email_id = row['Email ID']
        official_website = identify_official_website(nbfc_name, regional_office, address, email_id)
        return index, official_website

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(process_row, index, row): index for index, row in df_subset.iterrows()}
        for future in as_completed(futures):
            index, official_website = future.result()
            df.at[index, 'Official Website'] = official_website

    columns_to_export = ['Regional Office', 'NBFC Name', 'Address', 'Email ID', 'Official Website']
    output_file_path = 'output_nbfc_list.xlsx'
    df[columns_to_export].to_excel(output_file_path, index=False)
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Updated Excel file saved successfully with columns: {columns_to_export}")
    print(f"Time taken for the script to complete: {elapsed_time:.2f} seconds")

if __name__ == "__main__":
    main()

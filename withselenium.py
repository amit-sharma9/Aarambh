import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from concurrent.futures import ThreadPoolExecutor, as_completed

def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def get_official_website(driver, nbfc_name, regional_office, address, email_id):
    search_query = f"{nbfc_name} official website"
    search_url = f"https://www.google.com/search?q={search_query}"
    driver.get(search_url)
    time.sleep(2)
    try:
        results = driver.find_elements(By.XPATH, "//div[@class='BNeawe UPmit AP7Wnd']")
        if results:
            official_link = results[0].text
        else:
            official_link = "Not found"
    except Exception as e:
        print(f"Error finding official website for {nbfc_name}: {str(e)}")
        official_link = "Not found"
    return official_link

def process_row(index, row):
    nbfc_name = row['NBFC Name']
    regional_office = row['Regional Office']
    address = row['Address']
    email_id = row['Email ID']
    driver = init_driver()
    official_website = get_official_website(driver, nbfc_name, regional_office, address, email_id)
    driver.quit()
    return index, official_website

def main():
    start_time = time.time()
    file_path = 'nbfc_list.XLSX'
    sheet_name = 'List of NBFCs'
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    print(df.head())
    
    
    df['Official Website'] = ""

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(process_row, index, row): index for index, row in df.iterrows()}
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

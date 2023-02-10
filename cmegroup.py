from selenium import webdriver
import time
import gspread
from datetime import datetime
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller

def gen_current_time():
    # dd/mm/YY H:M:S
    dt_string = datetime.now().strftime("%B %d, %Y %I:%M:%S %p")
    print("STEP 00 - Get current Time", dt_string)
    return dt_string

def driver_define():
    print('STEP 01 - Chromedriver Installing')
    driver_path = chromedriver_autoinstaller.install()
    
    print('STEP 02 - Chrome Browser Opening')
    options = Options()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    s = Service(driver_path)
    driver = webdriver.Chrome(service=s, options =options)
    driver.maximize_window()
    return driver

def get_web_data():
    #gather Current Time
    current_time = gen_current_time()

    # Browser OPEN
    driver = driver_define()

    # Visit Base URL
    base_url = 'https://www.cmegroup.com/markets/interest-rates/cme-fedwatch-tool.html'
    print(f'STEP 03 - visiting > {base_url}')
    driver.get(base_url)

    # Wait & Select Iframe -> move
    print(f'STEP 04 - Wait Untill Page Load')
    cmeIframe = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[class="cmeIframe"]')))
    print(f'STEP 05 - Iframe Switch')
    driver.switch_to.frame(cmeIframe)
    time.sleep(1)

    # Wait and click to get Probabilities tables
    print(f'STEP 06 - Wait Untill Probabilities and click')
    probabilities_ele = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//a[text()="Probabilities"]')))
    time.sleep(1)
    driver.execute_script("arguments[0].scrollIntoView();", probabilities_ele)
    time.sleep(1)
    probabilities_ele.click()

    print(f'STEP 06 - Wait Untill MEETING PROBABILITIES Table appears')
    # Wait until MEETING PROBABILITIES table appears
    table = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//th[text()="Meeting Probabilities"]/../..')))

    # table Data Extract
    print(f'STEP 07 - Table data Extract')
    rows_eles = table.find_elements(By.TAG_NAME, 'tr')
    header_rows = [line.text for line in rows_eles[1].find_elements(By.TAG_NAME, 'th')] + ['DATA Updated TIME']
    print(header_rows)
    rows = []
    for row_ele in rows_eles[2:]:
        row = [line.text for line in row_ele.find_elements(By.TAG_NAME, 'td')] + [current_time]
        print(row)
        rows.append(row)

    write_data = [header_rows] + rows
    driver.quit()
    return write_data

# Return Sum and meeting Date and curent Time
# Data will be gathring from last rows
def add_sum(web_data):
    print(f'STEP 08 - Get Summation and Meeting Date')
    last_row = web_data[-1]
    meeting_date = last_row[0]
    gen_current_time = last_row[-1]
    
    
    cell1 = last_row[1].replace('%', '')
    cell2 = last_row[2].replace('%', '')
    cell3 = last_row[3].replace('%', '')
    
    total = float(cell1) + float(cell2) + float(cell3)
    round_total = round(total, 2)
    total_str = f"{round_total}%"
   
    return [meeting_date, total_str, gen_current_time]

# Data add on Google sheet
def conn_sheet():
    print(f'STEP 09 - Data Add on Google sheet')
    sa = gspread.service_account_from_dict({
      "type": "service_account",
      "project_id": "custom-sylph-376819",
      "private_key_id": "408b2464db93c2de8a5d05ae1f622f39e84dc07e",
      "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDZh3dN3Nrsq2Bo\n87UuUvixdEBEjTfNlz8euCUJWkkGegV2m/JWk5p3re3uRj7YCjCYuUe6QhWSyD61\nDO1DZGyCWvNga/0cEC4MKYmrSwM1ALD5aj8JGlxHV+UCgIPk/hRgSnxWwDdI30cv\nwIVr507vvckD5q6fwGGLS3H/dtDFKWtuf359iqQ+I1YNyko3+6rQanQw1TI4MQj3\nksN0i4s7UDLKWZoUNnvE9obUnkJwWhAbZC8eSegRoRhBpg0ebhnlm/fwzNTJK0TV\nL4x0MmKNaXTQsp+sG7TRpIuTfk9G6F8NAw5DOfnm2NweMdZP45TiQO1BM54xLaNO\nZWD537jLAgMBAAECggEAVNwehC8Q7rwKf9b3CuvOXffSbIvExbznoFXJCQWCMHcg\ns30xxGmPnHmrNMWNlZ0gCSxamYRXQyxAHkQ9OQmvtQjDIg3ur2h2dkMsFDlOtnof\nECXNEoGIl0JoMhotmgMussPMDtGsn46PCEdsJUSWzDr29MEkxWh5BSy4+6Z/2jHO\njn3V193Oo3NWJKLy4hk/ZYBz3uWdHYJa+O1RuXrhox1w8Wa6E9vNKcg0oEO+/Rti\nsfepSJMEqIAm/Pfbuha27m0Jc+qdSSiNszttSEW9d33Kx/jCH+jgWNWr4O6KAHl1\nO849ocEzfe6MofrbjAvtCcl1A5ngbu+6sEVdI5c/QQKBgQD2e8KbM8IcHpYaK3DA\nKWRPz32t6g3uF5gRn+kvhSVN7/Rv5NAwy+S8AiWmzf0oyFn+96E5+NjCE5SnCkIU\n72l/OtePfBks/L1rLIKMw0dxlb+5cS2xmePmL/GJ25ON5LQcA7F7/2530UmmvP8K\nDn5kTFJWVVG5inLUAqUj/JrjQQKBgQDh7YWoP7qJ0qlMufsj+6KXGYni6luLh5AD\nwDdgfWYtnSeaPCiOAnvxrjSL5B7TxCnS1apxaVkLZwNKVtu0Rnv5pRrpldbnkbvg\n9ekbeG8Ab3PCMcTc14lEJK963vV8MBvnJL1iD9GQSA4ZUpq7zgyyuU2mousCf0lX\nCyVXaMu1CwKBgQCJqcegHUFNqTuWdCqt+LBA7xc3miCbmOvi9Bgt5URXiixQjlBE\n4Kvo4Z4b0rKRI404HSAcG8Mcagk9XjpYLPsUB047okkBWkuE30Au1CZD5ypErVSi\n+9tQRfi2UT/RISoC94EaSyhsnSRwjuA2wq+O3x2hgFd7tDq79Jo9RilPwQKBgDl0\nB5tDqZJG6hrC6OS7pxs5uWDlLCaNcMgjZ3G4MfXDk0Cbr8x9QTuyi1ZPyq8boW8m\nOtPgcG5/4cxTzkdH7VsM640fN6ln3BlXL9J2i/PWY9+sfF2Uyil0EtPyQwczzMS5\nCRgY4bgJOtyhrwu3WG9SxDZuE8lsyR/Di9lwou0FAoGBAKNTedkmr28gL3iAGMvW\nm8WI9iET9S7t8NSylByu0RJOwiQ76w7qaxzbARdRrqm5PleMi97izwJYFZh+VgwP\nTxxPSKOzMIS8NzuVD37c+S/nfQ9eZtTHtVLjh9a/MbKQyVUG3D5WLORbNj/qV2zQ\nMa2umSHD29ilctI/+0oB3mma\n-----END PRIVATE KEY-----\n",
      "client_email": "cmegroup@custom-sylph-376819.iam.gserviceaccount.com",
      "client_id": "117115836600003395583",
      "auth_uri": "https://accounts.google.com/o/oauth2/auth",
      "token_uri": "https://oauth2.googleapis.com/token",
      "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
      "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/cmegroup%40custom-sylph-376819.iam.gserviceaccount.com"
    })
    worksheet_name = "Price Cuts"
    sheet_name = 'Sheet1'
    sh = sa.open(worksheet_name)
    wks = sh.worksheet(sheet_name)

    print('Worksheet Name:', worksheet_name)
    print('Sheet Name:', sheet_name)
    print('Sheet URL:', sh.url)
    print('Sucessfully Connected')
    
    return wks

def main():
    web_data = get_web_data()
    final_rows = add_sum(web_data)
    wks = conn_sheet()

    # Adding Data on sheet
    wks.append_row(final_rows, value_input_option='USER_ENTERED')
    print('Sucessfully Updated')

if __name__ == '__main__':
    main()

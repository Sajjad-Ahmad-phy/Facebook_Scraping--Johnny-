import os
from openpyxl import Workbook, load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

chromeOptions = webdriver.ChromeOptions()

driver = webdriver.Chrome()


def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')


"Login"

username = ""
password = ""

"Requesting Facebook"
driver.get("https://www.facebook.com/")
driver.maximize_window()
sleep(1)

driver.find_element('xpath', """ //*[@id="email"]""").send_keys(username)
driver.find_element('xpath', """  //*[@id="pass"]""").send_keys(password)
driver.find_element(By.NAME, """login""").click()
sleep(10)

"opening the existing csv files"
work_book = load_workbook(filename='URLs_List.xlsx')

# TO make active the excel file
ACT_Sheet = work_book.active

"O/p Excel"
W_B = Workbook()
sheet_N = W_B.active
sheet_N.title = 'Data sheet1'
sheet_N.append(['Links', 'Title', 'price', 'Listed Data', 'description', 'Driven', 'Transmission', 'Exterior Color',
                'Interior color', 'Conditions', 'Fuel Type', 'Engine size', 'Owners'])

# W_B = load_workbook(filename='Vehicles Info1.xlsx')
# sheet_N = W_B.active

"No of iterations"
No_of_Iterrations = len(ACT_Sheet['C'])

"URL for loop should be start from here"
for url in range(No_of_Iterrations):
    row_no = url + 2
    c = ACT_Sheet.cell(row=row_no, column=1)
    URL = f"{c.value}"

    headers = {
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.135 Safari/537.36 Edge/12.246"}
    driver.get(URL)
    sleep(1)
    Final_List = []
    Final_List.append(URL)

    data = driver.find_element(By.XPATH, " /html/body ")
    List = data.text.split('\n')
    sleep(1)

    if List[32] == 'This listing is no longer available':
        continue

    try:
        "Appending Title price and listed data"
        Final_List.append(List[31])
        Final_List.append(List[32])
        Final_List.append(List[33])
    except:
        pass
    try:
        D1 = List.index("Seller's description")
        D2 = List.index("Location is approximate")
        Des1 = List[D1 + 1:D2 - 1]
        "Seller's description"
        Description = " ".join([str(elem) for elem in Des1])
        Final_List.append(Description)
    except:
        Final_List.append('This listing is no longer available')
    try:
        "Engine details ..."
        a1 = List.index('About this vehicle')
        a2 = List.index("Seller's description")
        about = List[a1:a2 - 1]
        About_the_Vehicle = " ".join([str(ele) for ele in about])
        Final_List.append(About_the_Vehicle)

    except:
        pass

    sheet_N.append(Final_List)
    clear_screen()
    print(url)

    "Save o/p Excel"
    W_B.save('Vehicles Info.xlsx')

driver.quit()

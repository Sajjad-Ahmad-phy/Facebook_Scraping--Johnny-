"********************Links scraping portion***********************"

CarsModel_Links_file_name = 'FBCarModelUrl'
Cars_link_no = 0             #Cars' model link no

#Enter WorkBook For Saving Links
Links_Old_file = 'Excel file for all Cars links'
Links_New_file = 'Excel file for all Cars links'


"Login"
username = "03078440508"
password = "112223"

def Links_scraping():

    O_or_N = int(input('To Which File Do you want to scrape ****links*** of Cars. \n\t1. Old Excel file \n\t2. New Excel file \n\tEnter 1 for Old and 2 for New. \t'))

    from Facebook_Scraping_Fuctions import load_workbook, Links_Scraping_Functions, Login_Facebook
    "Login Facebook"
    Login_Facebook(username, password)

    cars_links_Wb = load_workbook(filename=f'D:\\PYTHON PROJECTS\\Facebook_Scraping--Johnny-\\Excel files\\{CarsModel_Links_file_name}.xlsx')
    Cars_links_sheet = cars_links_Wb.active
    Cars_links_iteration = len(Cars_links_sheet['B']) - Cars_link_no

    if O_or_N == 1:     #old
        try:
            for li in range(Cars_links_iteration):
                CarModel_Url = Cars_links_sheet.cell(row=li+1+Cars_link_no, column=2).value
                Class = Links_Scraping_Functions(CarModel_Url)
                Class.Output_Link()
                Class.Old_output_Exfile(Links_Old_file)
                print('Cars Model Link no: ', li+Cars_link_no+1)
        except:
            pass

    elif O_or_N == 2:  #New
        try:
            for Li in range(Cars_links_iteration):
                CarModel_url = Cars_links_sheet.cell(row=Li+1+Cars_link_no, column=2).value
                Class = Links_Scraping_Functions(CarModel_url)
                Class.Output_Link()
                if Li < 1:
                    Class.New_output_Exfile(Links_New_file)
                elif Li >= 1:
                    Class.Old_output_Exfile(Links_New_file)
                print("Cars Model Link no: ", Li+1)
        except:
            pass
    else:
        print('Wrong Entry')


"***********************************************************************************"
"Cars' Data Scraping portion"
Data_Old_file = "All cars details"         #output  cars' old data file
Data_New_file = "All cars details"    #output car's New data file

Links_file_name = 'Excel file for all Cars links'         #File's Name where All the links to all type of cars are stored

def Data_Scraping():
    O_or_N = int(input('To Which File Do you want to scrape ***Data*** of Cars. \n\t1. Old File\n\t2. New file \n\tEnter 1 for Old and 2 for New. \t\t'))

    "login"
    from Facebook_Scraping_Fuctions import load_workbook, Workbook, Login_Facebook
    Login_Facebook(username, password)

    Links_wb = load_workbook(f'D:\\PYTHON PROJECTS\\Facebook_Scraping--Johnny-\\Excel files\\{Links_file_name}.xlsx')
    ACT_Sheet = Links_wb.active

    from Facebook_Scraping_Fuctions import OutPut_Data

    if O_or_N == 1:
        W_B_O = load_workbook(filename=f'{Data_Old_file}.xlsx')
        sheet_O = W_B_O.active
        OutPut_Data(ACT_Sheet, sheet_O, Data_Old_file, W_B_O)

    elif O_or_N == 2:
        W_B_N = Workbook()
        sheet_N = W_B_N.active
        sheet_N.append(
            ['Links', 'Title', 'price', 'Listed Data', 'description', 'Driven', 'Transmission', 'Exterior Color',
             'Interior color', 'Conditions', 'Fuel Type', 'Engine size', 'Owners', 'Car links'])
        OutPut_Data(ACT_Sheet, sheet_N, Data_New_file, W_B_N)

    else:
        print('Wrong Entry')

"Everything Starts from here"

def main():
    Decide = int(input('Enter: \n\t 1). Links Scraping\n\t 2). Data Scraping \n\t\t'))
    if Decide == 1:
        Links_scraping()

    elif Decide == 2:
        Data_Scraping()

    else:
        print("Wrong Entry for Making Decision")
    print('Done')

if __name__ == '__main__':
    main()


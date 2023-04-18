from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
import openpyxl


# get the names
file = open('names.txt',"r")
names = file.readlines()
for i in range(len(names)) :
    names[i] = names[i].replace('\n','')
file.close()

# creat a Excel file
wb = Workbook()
wb.save('info.xlsx')

sheet_line = 2

# initialite the Excel file
wb = openpyxl.load_workbook('info.xlsx')
sheet = wb.active
sheet['A1'] = 'First name'
sheet['B1'] = 'Last name'
sheet['C1'] = 'Company name'
sheet['D1'] = 'office number'
sheet['E1'] = 'extension'
sheet['F1'] = 'cell phone'
sheet['G1'] = 'adresse email'
sheet['H1'] = 'Téléphones de bureau'

op = Options()
op.headless = True

web_page = "https://portail.ogq.qc.ca/repertoire"
path = "C:\Program Files (x86)\chromedriver.exe"



for name in names :

    driver = webdriver.Chrome(path,options=op)
    driver.get(web_page)
    input_element =  driver.find_element_by_class_name("form-control")
    input_element.send_keys(name)
    submit_button = driver.find_element_by_class_name("glyphicon")
    submit_button.click()


    # find all the <a> elements that contain the text "Détails"
    detail_links = driver.find_elements_by_xpath("//a[text()='Détails']")

    for link in detail_links:
        browser = webdriver.Chrome(path,options=op)
        browser.get(link.get_attribute("href"))

        # get information
        try:
            nom_elem = browser.find_element_by_xpath("//label[contains(text(), 'Nom')]/following-sibling::p")
            nom = nom_elem.text.strip()
        except:
            nom = ""
        try:
            prenom_elem = browser.find_element_by_xpath("//label[contains(text(), 'Prénom')]/following-sibling::p")
            prenom = prenom_elem.text.strip()
        except:
            prenom = ""
        try:
            permis_elem = browser.find_element_by_xpath("//label[contains(text(), 'Numéro de permis')]/following-sibling::p")
            permis = permis_elem.text.strip()
        except:
            permis = ""
        try:
            statut_elem = browser.find_element_by_xpath("//label[contains(text(), 'Statut d')]/following-sibling::p")
            statut = statut_elem.text.strip()
        except:
            statut = ""
        try:
            employeur_elem = browser.find_element_by_xpath("//div[contains(text(), 'Employeur')]/following-sibling::p")
            employeur = employeur_elem.text.strip()
        except:
            employeur = ""
        try:
            etablissement_elem = browser.find_element_by_xpath("//div[contains(text(), 'Nom d')]/following-sibling::p")
            etablissement = etablissement_elem.text.strip()
        except:
            etablissement = ""
        try:
            adresse_elem = browser.find_element_by_xpath("//div[contains(text(), 'Adresse du bureau')]/following-sibling::p")
            adresse = adresse_elem.text.strip().replace("<br>", "\n")
        except:
            adresse = ""
        try:
            telephone_elem = browser.find_element_by_xpath("//div[contains(text(), 'Téléphones de bureau')]/following-sibling::p")
            telephone = telephone_elem.text.strip()
        except:
            telephone = ""



        # add to the sheet
        sheet[f'A{sheet_line}'] = nom
        sheet[f'B{sheet_line}'] = prenom
        sheet[f'C{sheet_line}'] = permis
        sheet[f'D{sheet_line}'] = statut
        sheet[f'E{sheet_line}'] = employeur
        sheet[f'F{sheet_line}'] = etablissement
        sheet[f'G{sheet_line}'] = adresse
        sheet[f'H{sheet_line}'] = telephone

        wb.save("info.xlsx")

        browser.quit()
        sheet_line += 1
        print("adding ........")

    driver.quit()



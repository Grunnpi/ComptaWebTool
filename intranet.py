from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC

import os
import time
import csv
import argparse

def isBlank (myString):
    return not (myString and myString.strip())

class JournalStructureHeader:
    saison = ""
    code_structure = ""
    montant_du = ""
    montant_paye = ""
    montant_solde = ""
    url_detail = ""
    def __getattr__(self, attr):
        return self[attr]
    def getUrl(self):
        return self.url_detail
    def __init__(self,saison,code_structure,montant_du,montant_paye,montant_solde,url_detail):
        self.saison = saison
        self.code_structure = code_structure
        self.montant_du = montant_du.strip().replace(' €','').replace(',','.').replace(' ','')
        self.montant_paye = montant_paye.strip().replace(' €','').replace(',','.').replace(' ','')
        self.montant_solde = montant_solde.strip().replace(' €','').replace(',','.').replace(' ','')
        self.url_detail = url_detail
    def isNotZero(self):
        return ( self.montant_du != "0.00" or self.montant_paye != "0.00")

class JournalStructureDetail:
    date = ""
    libelle = ""
    montant_debit = ""
    montant_credit = ""
    type_ecriture = ""
    type = ""
    formation = ""
    def __getattr__(self, attr):
        return self[attr]
    def __init__(self,type_ecriture, date,libelle,montant_debit,montant_credit):
        self.type_ecriture = type_ecriture
        self.date = date.strip()
        self.libelle = libelle.strip()
        self.montant_debit = montant_debit.strip().replace(' €','').replace(',','.').replace(' ','')
        self.montant_credit = montant_credit.strip().replace(' €','').replace(',','.').replace(' ','')

        if ( self.libelle.find("Inscription action de formation") >= 0 ):
            self.type = "formation"
            split_libelle = self.libelle.split(' - ')
            self.formation = split_libelle[1]
        else:
            self.type = "autre"
            self.formation = ""
    def isDebit(self):
        return self.montant_debit != ""


def fetchNextPageJournalStructure(array_JournalStructureHeader, saison, current_page):
    print(driver.page_source)
    pagination_area = driver.find_element("class", "pagination_light")
    if ( pagination_area ):
        # print("Current page " + current_page)
        next_page_to_fetch = "-1"

        all_ligne1 = driver.find_elements(By.CSS_SELECTOR, "tr[class^='ligne']")
        for ligne in all_ligne1:
            # href
            theAHref = ligne.find_element(By.TAG_NAME, "a")
            theHref = theAHref.get_attribute("href")
            # montants
            theColumns = ligne.find_elements(By.TAG_NAME, "td")
            newJournal = JournalStructureHeader(saison, theColumns[0].text,theColumns[1].text,theColumns[2].text,theColumns[3].text,theHref)
            if ( newJournal.isNotZero() ):
                array_JournalStructureHeader.append(newJournal)
                print(theColumns[0].text + "[" + theHref + "]")

        if ( current_page != "0"):
            paginations = pagination_area.find_elements(By.TAG_NAME, "a")
            for pagination in paginations:
                if ( pagination.text > current_page):
                    print('-----------------------------------------------------------------------------')
                    pagination.click()
                    WebDriverWait(pagination, 5)

                    pagination_area = driver.find_element(By.CLASS_NAME, "pagination")
                    span_current_page = pagination_area.find_element(By.TAG_NAME, "span")
                    # print("Current page " + span_current_page.text)
                    next_page_to_fetch = span_current_page.text
                    break

        if ( next_page_to_fetch != "-1"):
            fetchNextPageJournalStructure(array_JournalStructureHeader, saison, next_page_to_fetch)


#ctl00_MainContent__journalClient__listeCommandesFacturationManuelle__gvEcritures
#ctl00_MainContent__journalClient__listeEcritures__gvEcritures
def fetchNextPageJournalStructureDetails(array_JournalStructureHeaderDetails, type_ecriture, id_tableau_details, current_page):
    tableau_detail_structure = driver.find_element(By.ID, id_tableau_details)
    if (tableau_detail_structure):
        try:
            pagination_area = driver.find_element(By.CLASS_NAME, "pagination_light")
        except NoSuchElementException:
            print('pas de tableau')
            return
        print("Current page Detail (b)" + current_page)
        next_page_to_fetch = "-1"

        all_ligne1 = tableau_detail_structure.find_elements(By.CSS_SELECTOR, "tr[class^='ligne']")
        for ligne in all_ligne1:
            theColumns = ligne.find_elements(By.TAG_NAME, "td")
            newJournalDetail = JournalStructureDetail(type_ecriture, theColumns[0].text, theColumns[1].text, theColumns[2].text, theColumns[3].text)
            if ( newJournalDetail.isDebit() ):
                array_JournalStructureHeaderDetails.append(newJournalDetail)
                print(ligne.text)

        if ( current_page != "0"):
            paginations = pagination_area.find_elements(By.TAG_NAME, "a")
            for pagination in paginations:
                if ( pagination.text > current_page):
                    print(pagination.text)
                    pagination.click()
                    WebDriverWait(pagination, 5)

                    pagination_area = driver.find_element(By.CLASS_NAME, "pagination_light")
                    span_current_page = pagination_area.find_element(By.TAG_NAME, "span")
                    print("Current page Detail (apres click)" + span_current_page.text)
                    next_page_to_fetch = span_current_page.text
                    break

        if ( next_page_to_fetch != "-1"):
            fetchNextPageJournalStructureDetails(array_JournalStructureHeaderDetails,type_ecriture,id_tableau_details,next_page_to_fetch)
######################


def extractionJournalStructure():
    # pas de Territoire en 2016-2017
    saisons = [
        '2019-2020',
        '2020-2021',
        '2021-2022',
    ]

    with open('eggs.csv', 'w', newline='') as csvfile:
        spamwriter = csv.writer(csvfile, dialect='excel')

        fieldnames = [
            'saison',
            'code_structure',
            'montant_du',
            'montant_paye',
            'montant_solde',
            'date',
            'libelle',
            'type',
            'formation',
            'type_ecriture',
            'montant_debit'
        ]

        writer = csv.DictWriter(csvfile, fieldnames=fieldnames,extrasaction='ignore')
        writer.writeheader()

        for saison in saisons:
            print("Saison [" + saison + "]")
            driver.get("https://intranet.sgdf.fr/Specialisation/Sgdf/FluxFinanciers/ConsultationEcrituresStructure.aspx")

            select = Select(driver.find_element("id", 'ctl00_MainContent__journalClient__filtreJournal__ddlSaisonComptable'))
            select.select_by_visible_text(saison)
            submit_valider = driver.find_element("id", "ctl00_MainContent__journalClient__filtreJournal__btnValider")
            submit_valider.click()
            WebDriverWait(driver, 5)

            # start with page 1
            array_JournalStructureHeader = []
            fetchNextPageJournalStructure(array_JournalStructureHeader, saison, "1")
            print(str(len(array_JournalStructureHeader)) + " elements")

            for currentJournalStructureHeader in array_JournalStructureHeader:
                # detail par Structure
                driver.get(currentJournalStructureHeader.getUrl())
                WebDriverWait(driver, 2)

                array_JournalStructureHeaderDetails = []
                #ctl00_MainContent__journalClient__listeEcritures__gvEcritures
                fetchNextPageJournalStructureDetails(array_JournalStructureHeaderDetails, "prelevement", "ctl00_MainContent__journalClient__listeEcritures__gvEcritures", "1")
                #ctl00_MainContent__journalClient__listeCommandesFacturationManuelle__gvEcritures
                fetchNextPageJournalStructureDetails(array_JournalStructureHeaderDetails,"en attente (gestionnaire)", "ctl00_MainContent__journalClient__listeCommandesFacturationManuelle__gvEcritures", "1")

                for journalStructureDetail in array_JournalStructureHeaderDetails:
                    dataRow = vars(journalStructureDetail)
                    dataRow.update(vars(currentJournalStructureHeader))
                    writer.writerow(dataRow)

######################

# method to get the downloaded file name
def getDownLoadedFileName(waitTime):
    driver.execute_script("window.open()")
    # switch to new tab
    driver.switch_to.window(driver.window_handles[-1])
    # navigate to chrome downloads
    driver.get('chrome://downloads')
    # define the endTime
    endTime = time.time()+waitTime
    while True:
        try:
            # get downloaded percentage
            downloadPercentage = driver.execute_script(
                "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('#progress').value")
            # check if downloadPercentage is 100 (otherwise the script will keep waiting)
            print(downloadPercentage)
            if downloadPercentage == 100:
                # return the file name once the download is completed
                return driver.execute_script("return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
        except:
            pass
        time.sleep(1)
        if time.time() > endTime:
            break

def tiny_file_rename(newname, folder_of_download, time_to_wait=60):
    time_counter = 0
    filename = max([f for f in os.listdir(folder_of_download)], key=lambda xa :   os.path.getctime(os.path.join(folder_of_download,xa)))
    while '.part' in filename:
        time.sleep(1)
        time_counter += 1
        if time_counter > time_to_wait:
            raise Exception('Waited too long for file to download')
    filename = max([f for f in os.listdir(folder_of_download)], key=lambda xa :   os.path.getctime(os.path.join(folder_of_download,xa)))
    os.rename(os.path.join(folder_of_download, filename), os.path.join(folder_of_download, newname))

def extractionRegistrePresence(p_saison, p_periode,p_periode_alias):
    print("Extraction registre presence")

    driver.get("https://intranet.sgdf.fr/Specialisation/Sgdf/ActivitesAnnee/ConsulterRegistrePresence.aspx")

    select = Select(driver.find_element("name", 'ctl00$MainContent$_EditeurRegistrePresence$_ddSaison'))
    select.select_by_visible_text(p_saison)
    time.sleep(2)
    select = Select(driver.find_element("name", 'ctl00$MainContent$_EditeurRegistrePresence$_ddPeriodes'))
    select.select_by_visible_text(p_periode)
    time.sleep(1)
    adherent_seulement = driver.find_element("id", 'ctl00_MainContent__EditeurRegistrePresence__cbAdherentsUniquement')
    webdriver.ActionChains(driver).move_to_element(adherent_seulement).click(adherent_seulement).perform()
    time.sleep(1)
    structures_enfants = driver.find_element("id", 'ctl00_MainContent__EditeurRegistrePresence__cbExporterSousStructure')
    webdriver.ActionChains(driver).move_to_element(structures_enfants).click(structures_enfants).perform()
    time.sleep(1)
    radio_nombre_heures = driver.find_element("id", 'ctl00_MainContent__EditeurRegistrePresence__rdbModeVolumeHoraireReel')
    webdriver.ActionChains(driver).move_to_element(radio_nombre_heures).click(radio_nombre_heures).perform()
    time.sleep(1)

    bouton_exporter_csv = driver.find_element("id", 'ctl00_MainContent__EditeurRegistrePresence__btnExporterExcel')
    webdriver.ActionChains(driver).move_to_element(bouton_exporter_csv).click(bouton_exporter_csv).perform()
    time.sleep(2)

    # latestDownloadedFileName = getDownLoadedFileName(10)
    tiny_file_rename("sboub" + p_periode_alias + ".csv","C:\Temp")
    # print(latestDownloadedFileName)


def lectureFichierRegistrePresence(p_filename):
    # open file in read mode
    with open("C:\\Temp\\" + p_filename, 'r') as read_obj:
        # pass the file object to reader() to get the reader object
        csv_reader=csv.reader(read_obj,delimiter=';')
        # Iterate over each row in the csv using reader object
        for row in csv_reader:
            print("1 case "  + row[0])
            if ( row[0] == ""):
                # Quelle unité ?
                print("Groupe")
            # for cell in row:
            #     # row variable is a list that represents a row in csv
            #     print(cell)

######################

if __name__ == "__main__":
    parser=argparse.ArgumentParser(description='Ecole Direct extact process')

    parser.add_argument('--user', help='ED User', type=str, required=True)
    parser.add_argument('--pwd', help='ED Password', type=str, required=True)

    args=parser.parse_args()

    options=webdriver.ChromeOptions()
    prefs={"download.default_directory":"C:\Temp"}
    options.add_experimental_option("prefs", prefs)
    driver=webdriver.Chrome(executable_path='./chromedriver100.0.4896.60.exe', chrome_options=options);

    connection_intranet=True

    if (connection_intranet):
        driver.get("https://intranet.sgdf.fr/Specialisation/Sgdf/Default.aspx")

        username = driver.find_element("id", "login")
        password = driver.find_element("id", "password")
        submit_login = driver.find_element("id", "_btnValider")

        username.send_keys(str(args.user))
        password.send_keys(str(args.pwd))

        submit_login.click()
        WebDriverWait(driver, 5)
        page_title = driver.title

        assert page_title == "Intranet SGDF - Accueil"

        # fonction ?
        # extractionRegistrePresence("2021-2022","Trimestre 2 (J-F-M)","T2")
        extractionRegistrePresence("2021-2022","Trimestre 3 (A-M-J)","T3")

    # lectureFichierRegistrePresence("sboubT2.csv")
    lectureFichierRegistrePresence("sboubT3.csv")

    print('Ok fin')
    # driver.close()
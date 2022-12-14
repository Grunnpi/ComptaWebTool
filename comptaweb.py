from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import csv
import argparse
import re

def isBlank (myString):
    return not (myString and myString.strip())

class CWDetail:
    saison = ""
    code_structure = ""
    montant_du = ""
    montant_paye = ""
    montant_solde = ""
    url_detail = ""

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
if __name__ == "__main__":
    parser=argparse.ArgumentParser(description='Ecole Direct extact process')

    # Ecole Directe cred
    parser.add_argument('--user', help='ED User', type=str, required=True)
    parser.add_argument('--pwd', help='ED Password', type=str, required=True)
    parser.add_argument('--csv', help='csv for ComptaWeb stuff', type=str, required=True)

    args=parser.parse_args()

    actionComptaWeb=False


    # Connection à ComptaWeb
    if (actionComptaWeb):
        driver = webdriver.Chrome('./chromedriver.exe')
        driver.get("https://comptaweb.sgdf.fr/login")

        username = driver.find_element("id", "username")
        password = driver.find_element("id", "password")
        submit_login = driver.find_element("id", "_submit")

        username.send_keys(str(args.user))
        password.send_keys(str(args.pwd))

        submit_login.click()
        WebDriverWait(driver, 5)
        page_title = driver.title

        assert page_title == "COMPTAWEB"

    # pour chaque ligne du fichier, on check si Recette ou Dépense (le reste on skip)
    with open(str(args.csv), 'r', encoding='utf-8') as read_obj:
        # pass the file object to reader() to get the reader object
        csv_reader=csv.reader(read_obj,delimiter=',')
        # Iterate over each row in the csv using reader object
        for row in csv_reader:
            ligneACreer=False
            if (len(row) > 4):
                if ( row[0].find("/2022") != -1 and row[2].find("Solde initial") == -1 ):

                    val_date=row[0]
                    val_ref=row[6]
                    val_libelle=row[7]

                    if ( row[3] != '' and row[3] !='0,00 €' ):
                        #print("WARN Débit")
#                        print(row)
                        val_type_de_transaction='Dépense'
                        val_montant=row[3]
                        ligneACreer=True
                    else:
                        if ( row[4] != '' and row[4] !='0,00 €' ):
                            #print("DEBUG Crédit")
                            #                            print(row)
                            val_type_de_transaction='Recette'
                            val_montant=row[4]
                            ligneACreer=True
                        else:
                            print("ERROR Inconnu")
                            print(row)
                            ligneACreer=False

                if (ligneACreer):
                    # clean du montant
                    val_montant=val_montant.replace(' €','')

                    if (actionComptaWeb):
                        # page Depense/Recette
                        driver.get("https://comptaweb.sgdf.fr/recettedepense/creer")

                        ## ecriture
                        # Type de transaction
                        #
                        ligne_type_de_transaction = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_depenserecette")
                        ligne_libelle = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_libel")
                        ligne_date = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_dateecriture")
                        ligne_montant = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_montant")
                        ligne_num_piece = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_numeropiece")
                        ligne_mode_transaction = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_modetransaction") # 1 : virement
                        ligne_mode_bancaire = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_comptebancaire") # 1421 : notre compte
                        ligne_categorie_tier = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_tierscateg") # 10 : autre pas SGDF

                        ligne_sub_montant = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_0_montant")
                        ligne_sub_nature = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_0_nature") # 2 default
                        ligne_sub_activite = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_0_activite")
                        ligne_sub_branche = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_0_brancheprojet")


                        # 1 : depense / 2 : recette / 3 : transfert
                        ligne_type_de_transaction.send_keys(val_type_de_transaction)
                        WebDriverWait(driver, 2)
                        ligne_libelle.send_keys(val_libelle)
                        ligne_date.send_keys(val_date)
                        ligne_montant.send_keys(val_montant)
                        WebDriverWait(driver, 2)
                        ligne_num_piece.send_keys(val_ref)

                        # a changer si espece ou bancaire
                        #ligne_mode_transaction.select_by_value("1")
                        ligne_mode_bancaire.send_keys("FR7630004028370001111386894 - SGDF 305707600 - GROUPE DU PAYS THIONVILLOIS")

                        #ligne_categorie_tier.select_by_value("10")

                    # préparation ventilation
                    tab_ref = []

                    tab_montant = []
                    tab_nature = []
                    tab_activite = []
                    tab_branche = []

                    tab_tag = []

                    # Est-ce une ligne a fonction multiple ?
                    if ( val_ref.find("\\") != -1 ):
                        # gestion du split
                        tab_ref = val_ref.split("\\")
                        tab_tag = val_libelle.split("\\")
                    else:
                        tab_ref.append(val_ref)
                        tab_tag.append(val_libelle)

                    # cleanup pour n'avoir que le tag
                    for idx, un_tag in enumerate(tab_tag):
                        try:
                            the_tag = re.search('#(.{2}?)', un_tag).group(1)
                            tab_tag[idx] = the_tag
                            # print(un_tag + " >>>>>>>>> " + str(the_tag))
                        except AttributeError:
                            tab_tag[idx] = "NO_TAG"
                            pass
                    # gestion REF par REF
                    # debug
                    print("ERROR [" + val_type_de_transaction + ']/[' + val_date + ']/[' + val_libelle + ']/[' + val_montant + ']/[' + val_ref + ']')

                    for idx, une_ref in enumerate(tab_ref):
                        if ( len(tab_ref) == len(tab_tag)):
                            val_tag = tab_tag[idx]
                        else:
                            if (len(tab_ref) > len(tab_tag)):
                                val_tag = tab_tag[min(idx,len(tab_tag)-1)]
                            else:
                                val_tag = "NO_TAG"

                        # voir pour faire defiler les lignes ...
                        val_nature="Participation Activités"
                        val_activite="Fonctionnement"
                        val_branche="Groupe"
                        if (une_ref.startswith("VEN-")):
                            val_nature="Vente article boutique"
                            val_activite="Tenue"
                        if (une_ref.startswith("ADH-")):
                            val_nature="Cotisations SGDF"
                            val_activite="Adhésion-National"
                        if (une_ref.startswith("WE-")):
                            val_nature="Participation Activités"
                            val_activite="Week-end"
                        if (une_ref.startswith("LBS-")):
                            val_nature="Achat destiné à la revente"
                            val_activite="Tenue"
                        if (une_ref.startswith("BNP-")):
                            val_nature="Frais Bancaires"
                            val_activite="Fonctionnement"

                        # voir pour choper le tag dans l'ordre
                        if (val_tag == "FA"):
                            val_branche="Farfadets"
                        if (val_tag == "LJ"):
                            val_branche="Louveteaux-Jeannettes"
                        if (val_tag == "SG"):
                            val_branche="Scouts-Guides"
                        if (val_tag == "PC"):
                            val_branche="Pionniers-Caravelles"
                        if (val_tag == "GP"):
                            val_branche="Groupe"
                        if (val_tag == "NO_TAG"):
                            val_branche="Groupe"

                        print("*WARN [" + val_tag + ']/[' + val_nature + ']/[' + val_activite + ']/[' + val_branche + ']')

                    if (actionComptaWeb):
                        #ligne_sub_montant.send_keys("421")
                        ligne_sub_nature.send_keys(val_nature)
                        ligne_sub_activite.send_keys(val_activite)
                        ligne_sub_branche.send_keys(val_branche)

                        ligne_sub_creer = driver.find_element("id", "ajout_detail")
                        ligne_sub_creer.click()

                        ligne_creer = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_submit")
                        #ligne_creer.click()
                        WebDriverWait(driver, 5)
                        #exit(-1)
    exit(-1)

print('Ok fin')
#driver.close()
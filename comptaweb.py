from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import csv
import argparse
import re
from decimal import Decimal
import time

def isBlank (myString):
    return not (myString and myString.strip())

class ComptaWebDetail:
    montant = 0
    nature = ""
    activite = ""
    branche = ""

    montant_test = 0
    def __init__(self,montant,nature,activite,branche):
        self.montant = montant
        self.nature = nature
        self.activite = activite
        self.branche = branche

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


def comptaWebDetail_toutes_les_combinaisons(index_debut, tab_ComptaWebDetail, tab_toutes_les_combinaisons):
    if ( index_debut == len(tab_ComptaWebDetail) - 1):
        # on est au fond de la liste
        nature_reconnue=False
        un_ComptaWebDetail = tab_ComptaWebDetail[index_debut]
        if (un_ComptaWebDetail.nature == "Cotisations SGDF" and un_ComptaWebDetail.branche != "Groupe" ):
            nature_reconnue=True
            for idx_possible_cotisation,une_possible_cotisation in enumerate(possible_cotisation):
                une_combinaison = []
                une_combinaison.append(une_possible_cotisation)
                tab_toutes_les_combinaisons.append(une_combinaison)
        if (un_ComptaWebDetail.nature == "Cotisations SGDF" and un_ComptaWebDetail.branche == "Groupe" ):
            nature_reconnue=True
            une_combinaison = []
            une_combinaison.append(Decimal("24.0"))
            tab_toutes_les_combinaisons.append(une_combinaison)
        if (un_ComptaWebDetail.nature == "Participation frais de Fonctionnement"):
            nature_reconnue=True
            une_combinaison = []
            une_combinaison.append(Decimal("25.0"))
            tab_toutes_les_combinaisons.append(une_combinaison)
        if (un_ComptaWebDetail.nature == "Vente article boutique"):
            nature_reconnue=True
            for idx_possible_achat_revente,un_possible_achat_revente in enumerate(possible_achat_revente):
                une_combinaison = []
                une_combinaison.append(un_possible_achat_revente)
                tab_toutes_les_combinaisons.append(une_combinaison)
        if (un_ComptaWebDetail.nature == "Participation Activités"):
            nature_reconnue=True
            for idx_possible_participation_activitee,une_possible_participation_activitee in enumerate(possible_participation_activitee):
                une_combinaison = []
                une_combinaison.append(une_possible_participation_activitee)
                tab_toutes_les_combinaisons.append(une_combinaison)
        if(not nature_reconnue):
            une_combinaison = []
            une_combinaison.append(Decimal("-1.0"))
            tab_toutes_les_combinaisons.append(une_combinaison)
    else:
        # appel pour combinaison n+1
        comptaWebDetail_toutes_les_combinaisons(index_debut+1,tab_ComptaWebDetail,tab_toutes_les_combinaisons)
        nature_reconnue=False
        un_ComptaWebDetail = tab_ComptaWebDetail[index_debut]
        tab_toutes_les_combinaisons_avant = tab_toutes_les_combinaisons.copy()
        tab_toutes_les_combinaisons.clear()
        for une_combinaison in tab_toutes_les_combinaisons_avant:
            if (un_ComptaWebDetail.nature == "Cotisations SGDF" and un_ComptaWebDetail.branche != "Groupe" ):
                nature_reconnue=True
                for idx_possible_cotisation,une_possible_cotisation in enumerate(possible_cotisation):
                    une_combinaison_argrandie = une_combinaison.copy()
                    une_combinaison_argrandie.append(une_possible_cotisation)
                    tab_toutes_les_combinaisons.append(une_combinaison_argrandie)
            if (un_ComptaWebDetail.nature == "Cotisations SGDF" and un_ComptaWebDetail.branche == "Groupe" ):
                nature_reconnue=True
                une_combinaison_argrandie = une_combinaison.copy()
                une_combinaison_argrandie.append(Decimal("24.0"))
                tab_toutes_les_combinaisons.append(une_combinaison_argrandie)
            if (un_ComptaWebDetail.nature == "Participation frais de Fonctionnement"):
                nature_reconnue=True
                une_combinaison_argrandie = une_combinaison.copy()
                une_combinaison_argrandie.append(Decimal("25.0"))
                tab_toutes_les_combinaisons.append(une_combinaison_argrandie)
            if (un_ComptaWebDetail.nature == "Vente article boutique"):
                nature_reconnue=True
                for idx_possible_achat_revente,un_possible_achat_revente in enumerate(possible_achat_revente):
                    une_combinaison_argrandie = une_combinaison.copy()
                    une_combinaison_argrandie.append(un_possible_achat_revente)
                    tab_toutes_les_combinaisons.append(une_combinaison_argrandie)
            if (un_ComptaWebDetail.nature == "Participation Activités"):
                nature_reconnue=True
                for idx_possible_participation_activitee,une_possible_participation_activitee in enumerate(possible_participation_activitee):
                    une_combinaison_argrandie = une_combinaison.copy()
                    une_combinaison_argrandie.append(une_possible_participation_activitee)
                    tab_toutes_les_combinaisons.append(une_combinaison_argrandie)
            if(not nature_reconnue):
                une_combinaison_argrandie = une_combinaison.copy()
                une_combinaison_argrandie.append(Decimal("-1.0"))
                tab_toutes_les_combinaisons.append(une_combinaison_argrandie)

######################
if __name__ == "__main__":
    parser=argparse.ArgumentParser(description='Ecole Direct extact process')

    # Ecole Directe cred
    parser.add_argument('--user', help='ED User', type=str, required=True)
    parser.add_argument('--pwd', help='ED Password', type=str, required=True)
    parser.add_argument('--csv', help='csv for ComptaWeb stuff', type=str, required=True)

    args=parser.parse_args()

    actionComptaWeb=True
    actionComptaWebCreationLigne=True


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
                if ( row[0].find("/202") != -1 and row[2].find("Solde initial") == -1 ):
                    print(row)
                    val_date=row[0]
                    val_date2=row[1]
                    val_description_bnp=row[2]
                    val_ref=row[6]
                    val_libelle=row[7]

                    if ( row[3] != '' and row[3] !='0,00 €' ):
                        #print("WARN Débit")
                        val_type_de_transaction='Dépense'
                        val_montant=row[3]
                        ligneACreer=True
                    else:
                        if ( row[4] != '' and row[4] !='0,00 €' ):
                            #print("DEBUG Crédit")
                            val_type_de_transaction='Recette'
                            val_montant=row[4]
                            ligneACreer=True
                        else:
                            print("ERROR ligne inconnue (ni Recette/ni Dépense)")
                            print(row)
                            ligneACreer=False

                    if (val_libelle.find("(ajustement)") != -1):
                        # cas d'exception : une ligne d'ajustement, donc déjà présente dans ComptaWeb N-1
                        print("WARN Ligne d'ajustement - on skip [" + val_libelle + "]")
                        ligneACreer=False

                    if (ligneACreer):
                        # clean du montant
                        val_montant=val_montant.replace(' €','')
                        try:
                            string_val_montant = val_montant.replace('\t', '').replace(' ', '').replace(',', '.').strip()
                            string_val_montant = ''.join(string_val_montant.split())
                            val_decimal_montant = Decimal(string_val_montant)
                            # print(string_val_montant + " == " + str(val_decimal_montant))
                        except:
                            print("ERROR montant string[" + val_montant + "] >> [" + string_val_montant + "]")
                            raise Exception("Convertion impossible")

                        if (val_type_de_transaction=='Dépense' and val_decimal_montant < 0.0):
                            val_decimal_montant = val_decimal_montant * -1

                        # préparation ventilation
                        tab_ref = []

                        tab_ComptaWebDetail = []

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

                        for idx, une_ref in enumerate(tab_ref):
                            # voir pour faire defiler les lignes ...
                            val_nature="Participation Activités"
                            val_activite="Fonctionnement"
                            val_branche="Groupe"

                            if ( len(tab_ref) == len(tab_tag)):
                                val_tag = tab_tag[idx]
                            else:
                                if (len(tab_ref) > len(tab_tag)):
                                    val_tag = tab_tag[min(idx,len(tab_tag)-1)]
                                else:
                                    val_tag = "NO_TAG"
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



                            ref_trouvee=False
                            if (une_ref.startswith("VEN-")):
                                val_nature="Vente article boutique"
                                val_activite="Tenue"
                                ref_trouvee=True
                            if (not ref_trouvee and une_ref.startswith("WE-")):
                                val_nature="Participation Activités"
                                val_activite="Week-end"
                                ref_trouvee=True
                            if (not ref_trouvee and une_ref.startswith("LBS-")):
                                val_nature="Achat destiné à la revente"
                                val_activite="Tenue"
                                ref_trouvee=True
                            if (not ref_trouvee and une_ref.startswith("BNP-")):
                                val_nature="Frais Bancaires"
                                val_activite="Fonctionnement"
                                ref_trouvee=True
                            if (not ref_trouvee and une_ref.startswith("CALENDRIERS-")):
                                val_nature="Dons, calendriers (sans reçu fiscal)"
                                val_activite="Calendriers"
                                ref_trouvee=True

                            if (ref_trouvee):
                                un_ComptaWebDetail = ComptaWebDetail(0,val_nature,val_activite,val_branche)
                                tab_ComptaWebDetail.append(un_ComptaWebDetail)

                            if (not ref_trouvee and une_ref.startswith("ADH-")):
                                val_nature="Cotisations SGDF"
                                val_activite="Adhésion-National"
                                ref_trouvee=True
                                un_ComptaWebDetail = ComptaWebDetail(0,val_nature,val_activite,val_branche)
                                tab_ComptaWebDetail.append(un_ComptaWebDetail)
                                if (not val_tag == "GP"):
                                    val_nature="Participation frais de Fonctionnement"
                                    val_activite="Adhésion-Groupe"
                                    un_ComptaWebDetail = ComptaWebDetail(0,val_nature,val_activite,val_branche)
                                    tab_ComptaWebDetail.append(un_ComptaWebDetail)

                            if (not ref_trouvee):
                                un_ComptaWebDetail = ComptaWebDetail(0,val_nature,val_activite,val_branche)
                                tab_ComptaWebDetail.append(un_ComptaWebDetail)


                        # estimation de la ventilation en fonction de la composition
                        if (len(tab_ComptaWebDetail) > 1):

                            possible_cotisation = [Decimal("140.0"), Decimal("105.0"), Decimal("59.0"), Decimal("24.0")]
                            idx_possible_cotisation = 0

                            possible_achat_revente = [Decimal("47.0"),Decimal("42.0"),Decimal("36.0"),Decimal("5.0")]
                            idx_possible_achat_revente = 0

                            possible_participation_activitee = [Decimal("15.0"),Decimal("13.0"),Decimal("10.0")]
                            idx_possible_participation_activitee = 0

                            tab_toutes_les_combinaisons = []
                            comptaWebDetail_toutes_les_combinaisons(0,tab_ComptaWebDetail,tab_toutes_les_combinaisons)

                            ventilation_auto_ok = False
                            stop_on_laisse_tomber = False
                            while ( not ventilation_auto_ok and not stop_on_laisse_tomber ):
                                stop_on_laisse_tomber=True
                                for une_combinaison in tab_toutes_les_combinaisons:
                                    val_decimal_montant_reste = val_decimal_montant
                                    nb_detail_identifiee = 0
                                    # print("Combinaison")
                                    une_combinaison.reverse()
                                    if (len(une_combinaison)==len(tab_ComptaWebDetail)):
                                        for index_montant,un_montant in enumerate(une_combinaison):
                                            if ( un_montant == Decimal("-1.0")):
                                                print("ERROR montant non identifié")
                                                tab_ComptaWebDetail[index_montant].montant_test=0.0
                                            else:
                                                # print(str(un_montant))
                                                nb_detail_identifiee+=1
                                                tab_ComptaWebDetail[index_montant].montant_test=un_montant
                                                val_decimal_montant_reste = val_decimal_montant_reste - un_montant
                                    else:
                                        print("ERROR >>>>>>>>>>>>>>>>>>>>>>>>>> pas bon nb d'entrée VS details !!!!!!!!!!!!!!")
                                    if ( nb_detail_identifiee == len(tab_ComptaWebDetail)):
                                        if ( val_decimal_montant_reste == 0 ):
                                            ventilation_auto_ok = True
                                            break

                            if(ventilation_auto_ok):
                                for idx,un_ComptaWebDetail in enumerate(tab_ComptaWebDetail):
                                    un_ComptaWebDetail.montant = un_ComptaWebDetail.montant_test
                            else:
                                print("ERROR : pas de ventilation auto :(")
                                val_libelle = "FIXME " + val_libelle
                                tab_ComptaWebDetail[0].montant = val_decimal_montant
                        else:
                            tab_ComptaWebDetail[0].montant = val_decimal_montant


                        # boucle sur les sous lignes
                        print(">INFO [" + val_type_de_transaction + ']/[' + val_date + ']/[' + val_libelle + ']/[' + str(val_decimal_montant) + ']/[' + val_ref + ']')
                        for idx,un_ComptaWebDetail in enumerate(tab_ComptaWebDetail):
                            print("*INFO [" + str(un_ComptaWebDetail.montant) + ']/[' + un_ComptaWebDetail.nature + ']/[' + un_ComptaWebDetail.activite + ']/[' + un_ComptaWebDetail.branche + ']')

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
                            ligne_carteprocurement = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_carteprocurement") # 1421 : notre compte
                            ligne_caisse = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_caisse") # 1421 : notre compte

                            ligne_categorie_tier = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_tierscateg") # 10 : autre pas SGDF


                            # 1 : depense / 2 : recette / 3 : transfert
                            ligne_type_de_transaction.send_keys(val_type_de_transaction)
                            WebDriverWait(driver, 2)
                            ligne_libelle.send_keys(val_libelle)
                            ligne_date.send_keys(val_date)
                            ligne_montant.send_keys(str(val_decimal_montant))
                            WebDriverWait(driver, 2)
                            ligne_num_piece.send_keys(val_ref)

                            # a changer si espece ou bancaire
                            # colonne BNP = "PAIEMENT C. PROC PTTFWNMVV" == carte procurement
                            # label = "B@" == remise de chèques
                            # 2eme colonne == "Caisse de Groupe" ou "Caisse Pionniers Caravelles"
                            if ( val_description_bnp.find("PAIEMENT C. PROC PTTFWNMVV") != -1 ):
                                ligne_mode_transaction.send_keys("Carte procurement")
                                WebDriverWait(driver, 2)
                                ligne_carteprocurement.send_keys("Carte du groupe")
                            elif ( val_libelle.startswith("B@")):
                                ligne_mode_transaction.send_keys("Chèque")
                                ligne_mode_bancaire.send_keys("FR7630004028370001111386894 - SGDF 305707600 - GROUPE DU PAYS THIONVILLOIS")
                            elif ( val_date2.startswith("Caisse")):
                                if ( val_type_de_transaction == "Dépense"):
                                    ligne_mode_transaction.send_keys("Caisse")
                                else:
                                    ligne_mode_transaction.send_keys("Espèces")
                                ligne_caisse.send_keys(val_date2)
                            else:
                                ligne_mode_bancaire.send_keys("FR7630004028370001111386894 - SGDF 305707600 - GROUPE DU PAYS THIONVILLOIS")

                            # saisie de la ventilation
                            WebDriverWait(driver, 2)

                            ligne_sub_montant = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_0_montant")
                            ligne_sub_nature = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_0_nature") # 2 default
                            ligne_sub_activite = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_0_activite")
                            ligne_sub_branche = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_0_brancheprojet")

                            if ( len(tab_ComptaWebDetail) == 1):
                                # précise pas le montant : c'est auto
                                ligne_sub_nature.send_keys(val_nature)
                                ligne_sub_activite.send_keys(val_activite)
                                ligne_sub_branche.send_keys(val_branche)
                            else:
                                for index_comptaWebDetail,un_ComptaWebDetail in enumerate(tab_ComptaWebDetail):
                                    ligne_sub_montant.send_keys(Keys.CONTROL + "a")
                                    ligne_sub_montant.send_keys(Keys.DELETE)
                                    WebDriverWait(driver, 2)
                                    ligne_sub_montant.send_keys(str(un_ComptaWebDetail.montant))
                                    ligne_sub_nature.send_keys(un_ComptaWebDetail.nature)
                                    ligne_sub_activite.send_keys(un_ComptaWebDetail.activite)
                                    ligne_sub_branche.send_keys(un_ComptaWebDetail.branche)

                                    if (index_comptaWebDetail < len(tab_ComptaWebDetail) - 1 ):
                                        ligne_sub_creer = driver.find_element("id", "ajout_detail")
                                        ligne_sub_creer.click()
                                        WebDriverWait(driver, 2)

                                        ligne_sub_montant = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_" + str(index_comptaWebDetail + 1) + "_montant")
                                        ligne_sub_nature = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_" + str(index_comptaWebDetail + 1) + "_nature") # 2 default
                                        ligne_sub_activite = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_" + str(index_comptaWebDetail + 1) + "_activite")
                                        ligne_sub_branche = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_ecriturecomptabledetails_" + str(index_comptaWebDetail + 1) + "_brancheprojet")

                            ligne_creer = driver.find_element("id", "portal_bundle_frontbundle_ecriturecomptable_submit")
                            if actionComptaWebCreationLigne:
                                try:
                                    ligne_creer.click()
                                    try:
                                        WebDriverWait(driver, 1).until(EC.alert_is_present(),
                                                                        'Timed out waiting for PA creation ' +
                                                                        'confirmation popup to appear.')
                                        alert = driver.switch_to.alert
                                        if (alert.text.startswith("Sauvegarde impossible : numéro de pièce déjà utilisé sur cet exercice !")):
                                            print("*ERROR : " + alert.text + " [" + val_ref + "]")
                                        else:
                                            print("*WARN : " + alert.text)
                                        alert.accept()
                                    except TimeoutException:
                                        #print("*INFO : " + driver.switch_to.active_element.text)
                                        pass
                                except Exception as monException:
                                    print("**ERROR Exception au click de création : ", monException)
                                    raise Exception("Stop it !")
                            WebDriverWait(driver, 5)
                            #exit(-1)
    print('Ok fin')
    #driver.close()
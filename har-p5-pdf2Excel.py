from pypdf import PdfReader
import pandas as pd
from RPA.Excel.Files import Files
from datetime import datetime
import decimal
import glob 

pdf_directory = r"D:\CAREERIA\it-opiskelu\2025-syksy\robot\win_progs5\Lasku\\"

excel_output_file = pdf_directory + "Laskut.xlsx"

def hae_laskun_perustiedot(lines):
    laskunro = int(lines[0].split(' ')[1])
    laskupvm = datetime.strptime(lines[1].split(' ')[1], '%d.%m.%Y').date()
    asikas_id = int(lines[2].split(' ')[2])
    maksuehto = ' '.join(lines[3].split(' ')[1:])
    return laskunro, laskupvm, asikas_id, maksuehto

def hae_asiakas_tiedot(lines):
    line_index = 0
    while not lines[line_index].strip() == 'ASIAKAS':
        line_index += 1

    line_index += 1
    asiakas_nimi = lines[line_index]
    line_index += 1
    asiakas_yritys = lines[line_index]
    line_index += 1
    asiakas_osoite = lines[line_index]
    line_index += 1
    asiakas_postinumero = lines[line_index].split(' ')[0]
    if len(lines[line_index].split(' ')) > 1:
        asiakas_postitoimipaikka = lines[line_index].split(' ')[1]
    else:
        asiakas_postitoimipaikka = 'TUNTEMATON'
    
    return asiakas_nimi, asiakas_yritys, asiakas_osoite, asiakas_postinumero, asiakas_postitoimipaikka

def hae_nimikkeet_ja_puhelin(lines, laskurivit_kpl):
    # Laskurien nimikkeet löytyvät riveiltä jotka ovat otsikkorivin 'Palvelut' jälkeen.
    line_index = 0
    while not lines[line_index].strip() == 'Palvelut':
        line_index += 1

    asiakas_puhelin = lines[line_index - 1]  # asiakkaan puhelin on aina Palvelut-otsikkoa edeltävällä rivillä

    line_index += 1
    nimikkeet = []
    for i in range(laskurivit_kpl):
        nimikkeet.append(lines[line_index])
        line_index += 1

    return nimikkeet, asiakas_puhelin

def hae_laskurivit_ja_yhteensa(lines):
    # otsikkorivin 'Tunnit a hinta SUMMA €' jälkeen alkavat laskurivit
    line_index = 0
    
    while not lines[line_index].strip() == 'Tunnit a hinta SUMMA €':
        line_index += 1
    line_index += 1

    laskurivit = []
    yhteensa_rivi = False
    while not yhteensa_rivi:
        if lines[line_index].split(' ')[0] == 'YHTEENSÄ':
            yhteensa = decimal.Decimal(lines[line_index].split(' ')[1].replace(',', '.'))
            yhteensa_rivi = True
            line_index += 1
            continue
        else:
            if lines[line_index].strip() == '-':
                line_index += 1
                continue
            else:
                laskurivi_data = lines[line_index].replace(',', '.')
                laskurivit_string = laskurivi_data.split(' ')
                laskurivit.append([decimal.Decimal(x) for x in laskurivit_string])
        line_index += 1

    return laskurivit, yhteensa


def parse_pdf_invoices(pdf_path):
    pdf_files = glob.glob(pdf_path + "*.pdf")

    asikkaat = []
    laskut = []
    tilausrivit = []

    for pdf_file in pdf_files:
        # Käsitellään kaikki hakemistossa olevat pdf-tiedostot
        print('-'*40)
        reader = PdfReader(pdf_file)
        print(f"Processing file: {pdf_file}")

        # Sivujen lukumäärä tulee olla 1. Muuten törmäämme ennen pitkään suuriin teoreettisiin ongelmiin.
        if len(reader.pages) != 1:
            print("VIRHE: Tämä ohjelma osaa käsitellä vain yksisivuisia laskuja. Tämä pdf-tiedosto ohitetaan.")
            continue

        page = reader.pages[0]
        print(page.extract_text())
        
        lines = page.extract_text().split('\n')
        
        laskunro, laskupvm, asikas_id, maksuehto = hae_laskun_perustiedot(lines)
        asiakas_nimi, asiakas_yritys, asiakas_osoite, asiakas_postinumero, asiakas_postitoimipaikka = hae_asiakas_tiedot(lines)
        laskurivit, yhteensa = hae_laskurivit_ja_yhteensa(lines)
        nimikkeet, asiakas_puhelin = hae_nimikkeet_ja_puhelin(lines, len(laskurivit))

        print("Laskun tiedot:")
        print(f"   {'Laskun numero: ':<30} {laskunro}")
        print(f"   {'Laskun päivämäärä: ':<30} {laskupvm}")
        print(f"   {'Asiakas ID: ':<30} {asikas_id}")
        print(f"   {'Maksuehto: ':<30} {maksuehto}")
        print("Asiakkaan tiedot:")
        print(f"   {'Nimi: ':<30} {asiakas_nimi}")
        print(f"   {'Yritys: ':<30} {asiakas_yritys}")
        print(f"   {'Osoite: ':<30} {asiakas_osoite}")
        print(f"   {'Postinumero: ':<30} {asiakas_postinumero}  {asiakas_postitoimipaikka}")
        print(f"   {'Puhelin: ':<30} {asiakas_puhelin}")

        print("Laskurivit:")
        for i, rivi in enumerate(laskurivit):
            print(f"  {i + 1:2}. {'Nimike: '} {nimikkeet[i]:20} {'Tunnit: '} {rivi[0]:5.1f} {'   Hinta: ':} {rivi[1]:8.2f} {'   Summa: '} {rivi[2]:10.2f}")
        print(f"Yhteensä: {yhteensa}")

        # lisätään asiakas jos sitä ei ole vielä listassa
        if asikas_id not in [a['AsiakasID'] for a in asikkaat]:
            asikkaat.append({
                'AsiakasID': asikas_id,
                'Nimi': asiakas_nimi,
                'Yritys': asiakas_yritys,
                'Osoite': asiakas_osoite,
                'Postinumero': asiakas_postinumero,
                'Postitoimipaikka': asiakas_postitoimipaikka,
                'Puhelin': asiakas_puhelin
            })

        # lisätään lasku
        laskut.append({
            'LaskunNumero': laskunro,
            'LaskunPaivamaara': laskupvm,
            'AsiakasID': asikas_id,
            'Maksuehto': maksuehto,
            'Yhteensa': yhteensa
            # 'Rivit': []
        })

        # Ei tehdä näin, koska tilausrivit halutaan erikseen omaan listaansa
        # for i, rivi in enumerate(laskurivit):
        #     laskut[-1]['Rivit'].append({
        #         'Nimike': nimikkeet[i],
        #         'Tunnit': rivi[0],
        #         'Hinta': rivi[1],
        #         'Summa': rivi[2]
        #     })

        # lisätään tilausrivit
        for i, rivi in enumerate(laskurivit):
            tilausrivit.append({
                'LaskunNumero': laskunro,
                'Nimike': nimikkeet[i],
                'Tunnit': rivi[0],
                'Hinta': rivi[1],
                'Summa': rivi[2]
            })
    
    return asikkaat, laskut, tilausrivit

def create_excel_file(excel_output_file, asikkaat, laskut, tilausrivit):        
    # PDF-Tiedostojen data Exceliin kolmelle välilehdelle:
    # 1. Asiakkaat (AsiakasID, Nimi, Yritys, Osoite, Postinumero, Postitoimipaikka, Puhelin)
    # 2. Laskut (LaskunNumero, LaskunPaivamaara, AsiakasID, Maksuehto, Yhteensa)
    # 3. Tilausrivit (LaskunNumero, Nimike, Tunnit, Hinta, Summa)
    excel = Files()
    print(f"Kirjoitetaan Excel-tiedosto: {excel_output_file}")
    excel.create_workbook(excel_output_file)
    try:
        # Asiakkaat
        df = pd.DataFrame(asikkaat, columns=['AsiakasID', 'Nimi', 'Yritys', 'Osoite', 'Postinumero', 'Postitoimipaikka', 'Puhelin'])
        excel.set_active_worksheet('Sheet')
        excel.rename_worksheet('Sheet', 'Asiakkaat')
        excel.append_rows_to_worksheet(df.to_dict("list"), header=True)
        excel.save_workbook()

        # Laskut
        df = pd.DataFrame(laskut, columns=['LaskunNumero', 'LaskunPaivamaara', 'AsiakasID', 'Maksuehto', 'Yhteensa'])
        excel.create_worksheet('Laskut')
        excel.set_active_worksheet('Laskut')
        excel.append_rows_to_worksheet(df.to_dict("list"), header=True)
        excel.save_workbook()

        # Tilausrivit
        df = pd.DataFrame(tilausrivit, columns=['LaskunNumero', 'Nimike', 'Tunnit', 'Hinta', 'Summa'])
        excel.create_worksheet('Tilausrivit')
        excel.set_active_worksheet('Tilausrivit')
        excel.append_rows_to_worksheet(df.to_dict("list"), header=True)
        excel.save_workbook()
    finally:
        excel.close_workbook()

if __name__ == "__main__":
    asikkaat, laskut, tilausrivit = parse_pdf_invoices(pdf_directory)
    create_excel_file(excel_output_file, asikkaat, laskut, tilausrivit)
    print("**** Done ****")
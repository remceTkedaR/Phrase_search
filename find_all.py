# This is a sample Python script.

# Text phrase search in *.xlsx files
# -----------------------------------

import os
import xlrd


def znajdz_frazy(plik, fraza, szukaj_w_kolumnie=True, kolumna=2):
    try:
        workbook = xlrd.open_workbook(plik)
        znaleziono_w_pliku = False

        for sheet in workbook.sheets():
            print(f"Przeszukiwanie pliku: {plik} - Arkusz: {sheet.name}")

            if szukaj_w_kolumnie and kolumna < sheet.ncols:
                kolumna_C = sheet.col_values(kolumna)
                if any(fraza.lower() in str(cell_value).lower() for cell_value in kolumna_C):
                    znaleziono_w_pliku = True
                    break
            else:
                for row in range(sheet.nrows):
                    for col in range(sheet.ncols):
                        try:
                            cell_value = sheet.cell_value(row, col)
                            if fraza.lower() in str(cell_value).lower():
                                znaleziono_w_pliku = True
                                break
                        except IndexError:
                            continue  # Ignorujemy błędy indeksowania, które mogą wystąpić w niektórych przypadkach

        return znaleziono_w_pliku
    except xlrd.XLRDError:
        return False


def przeszukaj_katalog(katalog, fraza, szukaj_w_kolumnie=True, kolumna=2):
    znalezione_pliki = []

    for folder, _, pliki in os.walk(katalog):
        for plik in pliki:
            if plik.lower().endswith('.xls'):
                sciezka_do_pliku = os.path.join(folder, plik)
                if znajdz_frazy(sciezka_do_pliku, fraza, szukaj_w_kolumnie, kolumna):
                    znalezione_pliki.append(sciezka_do_pliku)

    return znalezione_pliki


def zapisz_do_pliku_wynik(katalog, pliki):
    with open(os.path.join(katalog, 'wynik.txt'), 'w', encoding='utf-8') as wynik_file:
        for plik in pliki:
            wynik_file.write(plik + '\n')


if __name__ == "__main__":
    katalog_glowny = input("Podaj ścieżkę do katalogu głównego: ")
    szukana_fraza = input("Podaj szukaną frazę: ")
    szukaj_w_kolumnie = input("Czy szukać tylko w kolumnie C? (tak/nie): ").lower() == 'tak'
    kolumna_do_szukania = 2  # Domyślnie szukamy w kolumnie C

    if not szukaj_w_kolumnie:
        kolumna_do_szukania = None  # Ustawiamy na None, aby przeszukać cały plik

    znalezione_pliki = przeszukaj_katalog(katalog_glowny, szukana_fraza, szukaj_w_kolumnie, kolumna_do_szukania)

    if znalezione_pliki:
        print("\nZnaleziono następujące pliki:")
        for plik in znalezione_pliki:
            print(plik)

        zapisz_do_pliku_wynik(katalog_glowny, znalezione_pliki)
        print("Wyniki zapisane do pliku wynik.txt.")
    else:
        print("Nie znaleziono plików pasujących do kryteriów.")

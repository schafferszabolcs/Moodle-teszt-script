# A script Moodle tesztek féléves kiértékelésére szolgál, összegzi a hallgatók által megszerzett pontokat.
# A script úgy lett elkészítve, hogy figyelembe vegye a többszöri kitöltést, és a legjobb eredménnyel számoljon.
# Moodleről közvetlenül le kell tölteni a teszteredményeket .xlsx formátumban, és el kell helhelyezni őket egy mappában.
# Használat linux terminálban: python moodle-teszt.py <mappa elérési útvonala>
# A kiértékelés a mappán belül létrehozott processed mappában található results.xlsx néven.

import pandas as pd
import glob
import os
import sys

def transform_excel(input_file, output_file):
    # Az eredeti Excel fájl betöltése
    df = pd.read_excel(input_file)
    df = df.replace({',': '.'}, regex=True)
    df = df[df['Vezetéknév'] != 'Globális átlag']
    
    # Oszlopok a fájlban
    # print(f"Oszlopok a fájlban ({input_file}): {df.columns.tolist()}")

    try:
        df_transformed = pd.DataFrame()
        df_transformed['Név'] = df.iloc[:, 0] + ' ' + df.iloc[:, 1]
        df_transformed['Neptun'] = df.iloc[:, 3]
        df_transformed['Pont'] = pd.to_numeric(df.iloc[:, 9], errors='coerce')

        # Egyedi nevek kezelése a legnagyobb pontszámmal
        df_transformed = df_transformed.sort_values(by='Pont', ascending=False)
        df_transformed = df_transformed.drop_duplicates(subset=['Név'], keep='first')

        df_transformed.to_excel(output_file, index=False)
    except Exception as e:
        print(f"Hiba történt a fájl feldolgozása során ({input_file}): {e}")

def main(input_folder):
    output_folder = os.path.join(input_folder, "processed")

    # Ha a mappa nem létezik létre hozzuk
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Beolvas minden .xlsx fájlt a megadott mappából
    file_paths = glob.glob(f"{input_folder}/*.xlsx")
    if not file_paths:
        print("Nincsenek Excel fájlok a megadott mappában.")
        exit()

    # Az összes adatot tartalmazó lista
    all_data = []

    for file_path in file_paths:
        try:
            # Új fájlnév generálása az átalakított fájlhoz
            base_name = os.path.basename(file_path)
            transformed_file = os.path.join(output_folder, f"transformed_{base_name}")
            
            # Fájl transzformálása és mentése
            transform_excel(file_path, transformed_file)

            # Az Excel fájl beolvasása a releváns adatokkal
            df = pd.read_excel(transformed_file)

            #Ha a transzformált DataFrame nem üres, hozzáadjuk az 'all_data' listához
            if df is not None:
                all_data.append(df)

        except Exception as e:
            print(f"Hiba történt a fájl feldolgozása során: {file_path}. Hiba: {e}")

    # Az összes adatot egyesítjük, ha vannak adatok
    if all_data:
        merged_df = pd.concat(all_data)

        # Az adatokat csoportosítjuk név és neptun alapján, és összegezzük a pontszámokat
        result = merged_df.groupby(['Név', 'Neptun'], as_index=False).agg({'Pont': 'sum'})

        # Az eredmény mentése új Excel fájlba
        output_file = os.path.join(output_folder, "results.xlsx")
        result.to_excel(output_file, index=False)
        print(f"Az összesített eredményeket a '{output_file}' fájlba mentettük.")
    else:
        print("Nem volt adat, amit össze lehetett volna vonni.")

# Paraméterezhetőség biztosítása
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Használat: python moodle-teszt.py <mappa elérési útvonala>")
    else:
        input_file = sys.argv[1]
        main(input_file)

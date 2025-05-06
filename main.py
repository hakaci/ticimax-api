from pathlib import Path
from zeep import Client
from zeep.helpers import serialize_object
import pandas as pd
from datetime import datetime

# Configuration
BASE_URL = "-"
SERVICE = "UrunServis"
WSDL_URL = f"{BASE_URL}/servis/{SERVICE}.svc?wsdl"
UYE_KODU = "-"

# Create SOAP client
client = Client(wsdl=WSDL_URL)

# Fetch technical detail definitions
def fetch_teknik_detay_grup(client, uye_kodu):
    return serialize_object(client.service.SelectTeknikDetayGrup(uye_kodu, 0, ""))

def fetch_teknik_detay_ozellik(client, uye_kodu):
    return serialize_object(client.service.SelectTeknikDetayOzellik(uye_kodu, 0, ""))

def fetch_teknik_detay_deger(client, uye_kodu):
    return serialize_object(client.service.SelectTeknikDetayDeger(uye_kodu, 0, ""))

# Fetch product data
def select_urun(client, uye_kodu):
    urun_filtre = {'Aktif': 1}
    urun_sayfalama = {'BaslangicIndex': 0}
    urun_list = serialize_object(client.service.SelectUrun(uye_kodu, urun_filtre, urun_sayfalama))
    urun_list = [urun for urun in urun_list if urun.get('Resimler')]

    seen = set()
    unique_urun_list = []
    for urun in urun_list:
        key = urun.get("OzelAlan1")
        if key not in seen:
            unique_urun_list.append(urun)
            seen.add(key)
    return unique_urun_list

# Create lookup dictionaries
def teknik_detay_map(ozellik_list, deger_list):
    ozellik_dict = {item["ID"]: item["Tanim"] for item in ozellik_list}
    deger_dict = {item["ID"]: item["Tanim"] for item in deger_list}
    return ozellik_dict, deger_dict

# Main execution
def main():
    ozellik_list = fetch_teknik_detay_ozellik(client, UYE_KODU)
    deger_list = fetch_teknik_detay_deger(client, UYE_KODU)
    ozellik_dict, deger_dict = teknik_detay_map(ozellik_list, deger_list)

    urun_list = select_urun(client, UYE_KODU)

    rows = []
    for urun in urun_list:
        row = {"OzelAlan1": urun.get("OzelAlan1", "")}
        teknik_detaylar = urun.get("TeknikDetaylar", {}).get("UrunKartiTeknikDetay", [])
        for detay in teknik_detaylar:
            ozellik = ozellik_dict.get(detay.get("OzellikID"), f"Ozellik_{detay.get('OzellikID')}")
            deger = deger_dict.get(detay.get("DegerID"), f"Deger_{detay.get('DegerID')}")
            row[ozellik] = deger
        rows.append(row)

    df = pd.DataFrame(rows)

    output_dir = Path.cwd()
    date_str = datetime.today().strftime("%Y%m%d")
    filename = f"{date_str}_urun_teknik_detaylar.xlsx"
    output_path = output_dir / filename

    counter = 1
    while output_path.exists():
        output_path = output_dir / f"{date_str}_urun_teknik_detaylar_{counter}.xlsx"
        counter += 1

    df.to_excel(output_path, index=False)
    print(f"✅ Exported {len(df)} products to: {output_path}")

if __name__ == "__main__":
    main()

import pdfplumber
import pdf2image
import pytesseract
from PIL import Image
import pandas as pd
import re
import sys

def ekstrahiraj_pdf_u_excel(pdf_putanja, izlazni_excel="cisti_podaci.xlsx"):
    redci_paleta = []
    sav_tekst = ""
    
    print(f"Otvaram dokument: {pdf_putanja}")
    
    # 1. Pokušaj čitanja tekstualnog sloja
    try:
        with pdfplumber.open(pdf_putanja) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    sav_tekst += t + "\n--- STRANICA ---\n"
    except Exception as e:
        print(f"Greška pri čitanju PDF-a s pdfplumber: {e}")

    # 2. Ako nema dovoljno teksta (skenirani PDF), pokreni Tesseract OCR
    if not sav_tekst.strip() or len(sav_tekst.strip()) < 50:
        print("Dokument je skeniran (slike). Pokrećem Tesseract OCR...")
        try:
            images = pdf2image.convert_from_path(pdf_putanja)
            for i, img in enumerate(images):
                print(f"Obrađujem OCR-om stranicu {i+1}/{len(images)}...")
                ocr_t = pytesseract.image_to_string(img, lang='hrv+eng')
                sav_tekst += ocr_t + "\n--- STRANICA OCR ---\n"
        except Exception as e:
            print(f"Greška pri OCR obradi (provjerite je li Tesseract instaliran na sustavu): {e}")
            return

    # 3. Parsiranje blokova naloga
    blocks = re.split(r'Datum naloga:', sav_tekst, flags=re.IGNORECASE)
    
    for block in blocks[1:]:
        p_nalog_tekst = "Datum naloga:" + block
        lines = [l.strip() for l in p_nalog_tekst.split('\n') if l.strip()]
        
        dt_naloga = ""
        dt_isporuke = ""
        shpt = ""
        ref = ""
        zip_kod = ""
        grad = ""
        
        for l in lines:
            if l.upper().startswith("DATUM NALOGA:"):
                dt_naloga = l.split(":")[-1].strip()
            elif "DATUM ISPORUKE" in l.upper():
                dt_isporuke = l.split(":")[-1].strip()
            elif "POŠILJKA:" in l.upper() or "POSILJKA:" in l.upper():
                m_sh = re.search(r'(EP-[\d]+|ZAG-[\d\-]+)', l, re.IGNORECASE)
                if m_sh: shpt = m_sh.group(1)
            elif "REFERENCA:" in l.upper():
                ref = l.split(":")[-1].strip()
        
        # Ekstrakcija ZIP-a i Grada
        m_zip_grad = re.search(r'(?:HR-)?(\d{5})\s+([A-Za-zČĆŠĐŽčćšđž\s]+)', p_nalog_tekst)
        if m_zip_grad:
            zip_kod = m_zip_grad.group(1)
            grad = m_zip_grad.group(2).strip().split(',')[0]
        else:
            zip_kod = "10410"
            grad = "Velika Gorica"

        # Financije iz naloga
        naplaceno_osnovna = 0.0
        naplaceno_gorivo = 0.0
        for l in lines:
            if re.match(r'^(110|100)\s+', l):
                m_izn = re.findall(r'([\d\.]*,\d{2})', l)
                if m_izn:
                    naplaceno_osnovna = float(m_izn[-1].replace('.', '').replace(',', '.'))
            elif any(d_kw in l.lower() for d_kw in ['dizel', 'dodatak', 'gorivo']):
                m_izn = re.findall(r'([\d\.]*,\d{2})', l)
                if m_izn:
                    naplaceno_gorivo = float(m_izn[-1].replace('.', '').replace(',', '.'))

        # Ekstrakcija stavki paleta bez dupliranja
        vec_dodane = set()
        for idx_l, linija in enumerate(lines):
            linija_upper = linija.upper()
            if any(t in linija_upper for t in ['EWP', 'FP', 'OWP', 'CLL']) or 'OTP' in linija_upper:
                if "OZNAKA" in linija_upper or "KOLIČI" in linija_upper:
                    continue
                
                tip_palete = "FP"
                for t_tip in ['EWP', 'OWP', 'CLL', 'FP']:
                    if t_tip in linija_upper:
                        tip_palete = t_tip
                        break
                
                kolicina = 1
                m_kol = re.search(r'(\d+)\s+' + tip_palete, linija_upper)
                if m_kol:
                    kolicina = int(m_kol.group(1))

                masa_kg = 0.0
                for k in range(max(0, idx_l - 1), min(len(lines), idx_l + 2)):
                    m_masa = re.search(r'(\d+[\d\.]*,\d{2,3})', lines[k])
                    if m_masa:
                        val = float(m_masa.group(1).replace('.', '').replace(',', '.'))
                        if val > 2.0:
                            masa_kg = val
                            break

                if masa_kg > 0:
                    m_otp = re.search(r'(otp-?[\d\/]+)', linija, re.IGNORECASE)
                    oznaka = m_otp.group(1) if m_otp else "Standard"
                    
                    jedinstveni_kljuc = (shpt, ref, oznaka, masa_kg, tip_palete)
                    if jedinstveni_kljuc not in vec_dodane:
                        vec_dodane.add(jedinstveni_kljuc)
                        redci_paleta.append({
                            'LA-ID': shpt if shpt else "EP-GUEST",
                            'Referenca': ref if ref else "N/A",
                            'Oznaka_Broj': oznaka,
                            'Datum_Naloga': dt_naloga,
                            'Datum_Isporuke': dt_isporuke,
                            'Grad': grad,
                            'ZIP': zip_kod,
                            'Broj_Paleta': kolicina,
                            'Tip_Palete': tip_palete,
                            'Masa_Palete_KG': round(masa_kg, 2),
                            'Naplaceno_Osnovna_EUR': naplaceno_osnovna,
                            'Naplaceno_Gorivo_EUR': naplaceno_gorivo,
                            'Naplaceno_Ukupno_EUR': round(naplaceno_osnovna + naplaceno_gorivo, 2)
                        })

    df = pd.DataFrame(redci_paleta)
    if not df.empty:
        df.to_excel(izlazni_excel, index=False)
        print(f"Uspješno generiran Excel: {izlazni_excel} s ukupno {len(df)} stavki!")
    else:
        print("Upozorenje: Nije pronađena nijedna stavka za izvoz.")
    return df

if __name__ == "__main__":
    pdf_fajl = sys.argv[1] if len(sys.argv) > 1 else "englmayer_racun.pdf"
    ekstrahiraj_pdf_u_excel(pdf_fajl, "cisti_podaci.xlsx")

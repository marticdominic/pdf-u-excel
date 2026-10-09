import streamlit as st
import pandas as pd
import math
import io

st.set_page_config(
    page_title="Audit Cjenika i Revizija - G. Englmayer",
    page_icon="📊",
    layout="wide"
)

def odrediti_zonu(zip_str):
    if not zip_str:
        return "Zona 3"
    z_clean = str(zip_str).strip()
    if not z_clean.isdigit() or len(z_clean) < 5:
        return "Zona 3"
    prva_znam = z_clean[0]
    if prva_znam in ['1', '4']:
        return "Zona 1"
    elif prva_znam == '4' and z_clean.startswith(('40', '42', '43', '44', '47', '48', '49')):
        return "Zona 2"
    elif prva_znam in ['5']:
        return "Zona 3"
    elif prva_znam in ['3']:
        return "Zona 4"
    elif prva_znam in ['2']:
        return "Zona 5"
    else:
        return "Zona 3"

def ugovorena_cijena_osnovna(zona, masa_kg, tip_palete):
    cijene = {
        "Zona 1": {100: 35.00, 300: 40.00, 600: 45.00, 99999: 50.00},
        "Zona 2": {100: 38.00, 300: 43.00, 600: 48.00, 99999: 53.00},
        "Zona 3": {100: 42.00, 300: 47.00, 600: 51.00, 99999: 55.00},
        "Zona 4": {100: 45.00, 300: 50.00, 600: 54.00, 99999: 58.00},
        "Zona 5": {100: 48.00, 300: 53.00, 600: 56.00, 99999: 61.00},
    }
    zona_tabela = cijene.get(zona, cijene["Zona 3"])
    osnova = 51.00
    for limit_kg, cijena in sorted(zona_tabela.items()):
        if masa_kg <= limit_kg:
            osnova = cijena
            break
            
    if tip_palete == "OWP":
        osnova *= 1.5
    return round(osnova, 2)

def izracunaj_dizel_dodatak(osnovna_cijena, cijena_goriva_trenutna=1.65):
    baseline = 1.46
    if cijena_goriva_trenutna <= baseline:
        return 0.0
    razlika = cijena_goriva_trenutna - baseline
    postotak = math.ceil(razlika / 0.05) * 0.015
    return round(osnovna_cijena * postotak, 2)

st.title("📊 Revizija Invoica prema Ugovoru OF 002/2026")
st.markdown("Učitajte očišćenu Excel datoteku generiranu iz PDF-a i obavite potpunu reviziju cijena, zona i preplata.")

uploaded_excel = st.file_uploader("Učitajte Excel datoteku (.xlsx)", type=["xlsx"])

if uploaded_excel is not None:
    df = pd.read_excel(uploaded_excel)
    
    cijena_goriva_input = st.sidebar.number_input("Trenutna cijena dizela (€/L)", value=1.65, step=0.01)
    
    zone_list = []
    ugovorene_osnove = []
    ugovorena_goriva = []
    
    for idx, row in df.iterrows():
        zona = odrediti_zonu(row['ZIP'])
        zone_list.append(zona)
        
        ug_osn = ugovorena_cijena_osnovna(zona, row['Masa_Palete_KG'], row['Tip_Palete'])
        ug_gor = izracunaj_dizel_dodatak(ug_osn, cijena_goriva_input)
        
        ugovorene_osnove.append(ug_osn)
        ugovorena_goriva.append(ug_gor)
        
    df['Zona'] = zone_list
    df['Ugovoreno_Osnovna_EUR'] = ugovorene_osnove
    df['Ugovoreno_Gorivo_EUR'] = ugovorena_goriva
    df['Ugovoreno_Ukupno_EUR'] = df['Ugovoreno_Osnovna_EUR'] + df['Ugovoreno_Gorivo_EUR']
    df['Razlika_Preplata_EUR'] = round(df['Naplaceno_Ukupno_EUR'] - df['Ugovoreno_Ukupno_EUR'], 2)

    tab1, tab2, tab3 = st.tabs(["📋 Detaljni Pregled", "💰 Analiza Preplata", "📈 Sažetak po Zonama"])
    
    with tab1:
        st.subheader(f"Učitane stavke (Ukupno: {len(df)} redaka)")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Konačni_Audit')
        excel_data = output.getvalue()
        
        st.download_button(
            label="📥 Preuzmi gotov Audit izvještaj (.xlsx)",
            data=excel_data,
            file_name="Englmayer_Audit_Izvjestaj.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with tab2:
        st.subheader("Pregled utvrđenih preplata")
        preplate_df = df[df['Razlika_Preplata_EUR'] > 0]
        total_preplata = preplate_df['Razlika_Preplata_EUR'].sum()
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Ukupno stavki", len(df))
        col2.metric("Ukupno naplaćeno", f"{df['Naplaceno_Ukupno_EUR'].sum():.2f} €")
        col3.metric("Ukupne preplate", f"{total_preplata:.2f} €", delta_color="inverse")
        
        st.dataframe(preplate_df[['LA-ID', 'Referenca', 'Grad', 'Zona', 'Masa_Palete_KG', 'Naplaceno_Ukupno_EUR', 'Ugovoreno_Ukupno_EUR', 'Razlika_Preplata_EUR']], use_container_width=True)

    with tab3:
        st.subheader("Agregirani pregled po zonama")
        zona_group = df.groupby('Zona').agg(
            Broj_Stavki=('LA-ID', 'count'),
            Ukupna_Masa_KG=('Masa_Palete_KG', 'sum'),
            Naplaceno_EUR=('Naplaceno_Ukupno_EUR', 'sum'),
            Ugovoreno_EUR=('Ugovoreno_Ukupno_EUR', 'sum'),
            Razlika_EUR=('Razlika_Preplata_EUR', 'sum')
        ).reset_index()
        st.dataframe(zona_group, use_container_width=True)

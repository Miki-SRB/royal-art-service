
import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io
from PIL import Image as PILImage
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "rapports_supervision.xlsx")

PREFIXES_LTA = {
    "AF": "057-",  # Air France
    "KE": "180-",  # Korean Air
    "DL": "006-",  # Delta Air Lines
    "AA": "001-",  # American Airlines
    "UA": "016-",  # United Airlines
    "LH": "020-",  # Lufthansa Cargo
    "BA": "125-",  # British Airways
    "EK": "176-",  # Emirates
    "QR": "157-",  # Qatar Airways
    "SQ": "618-",  # Singapore Airlines
    "CX": "160-",  # Cathay Pacific
    "AY": "105-",  # Finnair
    "TK": "235-",  # Turkish Airlines
    "JL": "131-",  # Japan Airlines
    "NH": "205-",  # All Nippon Airways (ANA)
    "EY": "607-",  # Etihad Airways
    "CI": "297-",  # China Airlines
    "BR": "695-",  # EVA Air
    "CZ": "784-",  # China Southern
    "MU": "781-",  # China Eastern
    "CA": "999-",  # Air China
    "AC": "014-",  # Air Canada
    "KL": "074-",  # KLM
    "LX": "724-",  # Swiss
    "OS": "257-",  # Austrian Airlines
    "SN": "082-",  # Brussels Airlines
    "TP": "047-",  # TAP Air Portugal
    "IB": "075-",  # Iberia
}

if not os.path.exists(EXCEL_FILE):
    df_init = pd.DataFrame(columns=[
        "ID_Rapport", "Date_Saisie", "Superviseur", "Date_Operation", 
        "Vol_Ref", "Convoyeur", "LTA", "Palettes_Detail", "Siege_Cabine", 
        "ETD", "Heure_Decollage", "Retard", "Commentaires", "Nb_Photos"
    ])
    df_init.to_excel(EXCEL_FILE, index=False)

def chercher_logo():
    possible_names = ["RAS-gold (1).png", "RAS-gold.png", "logo.png", "logo_ras.png", "logo_ras.jpg", "logo.jpg"]
    for filename in possible_names:
        full_path = os.path.join(BASE_DIR, filename)
        if os.path.exists(full_path):
            return full_path
    try:
        for root, dirs, files in os.walk(BASE_DIR):
            for file in files:
                if file.lower() in [p.lower() for p in possible_names]:
                    return os.path.join(root, file)
    except Exception:
        pass
    return None

def preparer_logo_pdf(logo_path):
    """Prépare le logo sur fond blanc pour éviter la transparence excessive"""
    try:
        pil_logo = PILImage.open(logo_path)
        if pil_logo.mode in ('RGBA', 'LA') or (pil_logo.mode == 'P' and 'transparency' in pil_logo.info):
            bg = PILImage.new("RGB", pil_logo.size, (255, 255, 255))
            bg.paste(pil_logo, (0, 0), pil_logo)
            pil_logo = bg
        else:
            pil_logo = pil_logo.convert('RGB')
        buf = io.BytesIO()
        pil_logo.save(buf, format='JPEG', quality=95)
        buf.seek(0)
        return buf
    except Exception:
        return logo_path

def dessiner_cadre_et_logo(c, logo_clean_buf):
    """Dessine le double encadrement doré et le logo sur la page active"""
    couleur_or = colors.HexColor('#D4AF37')
    
    # Cadre extérieur principal
    c.setStrokeColor(couleur_or)
    c.setLineWidth(1.5)
    c.rect(20, 20, 572, 752)
    
    # Cadre intérieur fin
    c.setLineWidth(0.5)
    c.rect(24, 24, 564, 744)
    
    # Logo
    if logo_clean_buf:
        try:
            c.drawImage(ImageReader(logo_clean_buf), 420, 715, width=130, height=55, preserveAspectRatio=True)
        except Exception:
            pass

def obtenir_prochain_id():
    df = pd.read_excel(EXCEL_FILE)
    annee = datetime.now().year
    nb_rapports = len(df) + 1
    return f"RAS-{annee}-{nb_rapports:03d}"

def generer_pdf_bytes(id_rapport, colab, date_op, vol, convoyeur, lta, etd_dec, heure_dec, retard_txt, palettes_list, siege_cabine, comm, photos_bytes_list):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    
    couleur_rouge = colors.HexColor('#CC0000')
    couleur_noire = colors.HexColor('#000000')
    
    logo_path = chercher_logo()
    logo_clean_buf = preparer_logo_pdf(logo_path) if logo_path else None
    
    # Application sur la première page
    dessiner_cadre_et_logo(c, logo_clean_buf)

    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(couleur_noire)
    c.drawString(50, 750, "ROYAL ART SERVICE - Operation report")
    
    c.setFont("Helvetica", 11)
    c.drawString(50, 732, "Référence / Reference : ")
    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(couleur_rouge)
    c.drawString(180, 732, f"{id_rapport}")
    
    c.setFillColor(couleur_noire)
    c.setFont("Helvetica", 11)
    c.drawString(50, 715, "Superviseur / Supervisor : ")
    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(couleur_rouge)
    c.drawString(190, 715, f"{colab}")
    
    c.setFillColor(couleur_noire)
    c.line(50, 700, 550, 700)
    
    y = 675  # Position de départ plus basse pour aérer
    nb_photos_count = len(photos_bytes_list) if photos_bytes_list else 0
    
    # Date
    c.setFont("Helvetica-Bold", 10)
    c.setFillColor(couleur_noire)
    c.drawString(50, y, "• Date d'opération / Date of operation : ")
    c.setFillColor(couleur_rouge)
    c.drawString(250, y, f"{date_op}")
    y -= 22  # Espacement augmenté
    
    # Vol
    c.setFillColor(couleur_noire)
    c.drawString(50, y, "• Numéro de vol / Flight N° / Ref : ")
    c.setFillColor(couleur_rouge)
    c.drawString(215, y, f"{vol}")
    y -= 22
    
    # Convoyeur
    c.setFillColor(couleur_noire)
    c.drawString(50, y, "• Convoyeur / Courier : ")
    c.setFillColor(couleur_rouge)
    c.drawString(165, y, f"{convoyeur if convoyeur else 'N/A'}")
    y -= 22
    
    # LTA
    lta_final = lta.strip() if lta and lta.strip() else ""
    if not lta_final and vol:
        code_comp = vol.strip()[:2].upper()
        if code_comp in PREFIXES_LTA:
            lta_final = PREFIXES_LTA[code_comp]

    c.setFillColor(couleur_noire)
    c.drawString(50, y, "• LTA / AWB : ")
    c.setFillColor(couleur_rouge)
    c.drawString(120, y, f"{lta_final if lta_final else 'N/A'}")
    y -= 22

    # Palettes
    c.setFillColor(couleur_noire)
    c.drawString(50, y, "• Détails et Positions des Palettes / Pallets & Positions :")
    y -= 18
    
    for item in palettes_list:
        c.setFont("Helvetica", 10)
        c.setFillColor(couleur_noire)
        c.drawString(70, y, "- Palette / Pallet N° ")
        
        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(couleur_rouge)
        txt_p = f"{item['num_palette']} : {item['nb_caisses']} caisse(s) / crate(s)"
        c.drawString(165, y, txt_p)
        
        if item['position'] and item['position'] != "N/A":
            x_pos = 165 + c.stringWidth(txt_p, "Helvetica-Bold", 10) + 4
            c.setFont("Helvetica", 10)
            c.setFillColor(couleur_noire)
            c.drawString(x_pos, y, " - Position : ")
            x_pos += c.stringWidth(" - Position : ", "Helvetica", 10)
            c.setFont("Helvetica-Bold", 10)
            c.setFillColor(couleur_rouge)
            c.drawString(x_pos, y, f"{item['position']}")
            
        y -= 18
        
    if siege_cabine and siege_cabine != "N/A":
        y -= 4
        c.setFont("Helvetica", 10)
        c.setFillColor(couleur_noire)
        c.drawString(70, y, "• Numéro de siège / Seat N° : ")
        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(couleur_rouge)
        c.drawString(215, y, f"{siege_cabine}")
        y -= 18

    y -= 4
    c.setFont("Helvetica-Bold", 10)
    c.setFillColor(couleur_noire)
    c.drawString(50, y, "• Heure de décollage / Wheels up : ")
    c.setFillColor(couleur_rouge)
    txt_dec_final = f"{heure_dec}"
    if retard_txt and retard_txt != "0 min":
        txt_dec_final += f" (Retard / Delay : {retard_txt})"
    c.drawString(225, y, txt_dec_final)
    y -= 24
    
    # Remarques
    c.setFillColor(couleur_noire)
    c.drawString(50, y, "• Remarques / Compte-rendu / Remarks & Report :")
    y -= 16
    
    style_remarques = ParagraphStyle(
        'RemarquesStyle',
        fontName='Helvetica-Bold',
        fontSize=10,
        textColor=couleur_rouge,
        leading=14
    )
    
    texte_comm = comm.replace('\n', '<br/>') if comm else "Aucune remarque spécifique / No specific remarks."
    p = Paragraph(texte_comm, style_remarques)
    w, h = p.wrap(500, 150)
    p.drawOn(c, 50, y - h)
    y -= (h + 25)
    
    # --- DESSIN DES PHOTOS ---
    if photos_bytes_list:
        max_w = 230
        max_h = 140
        
        if y - max_h < 50:
            c.showPage()
            y = 730
            dessiner_cadre_et_logo(c, logo_clean_buf)
        
        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(couleur_noire)
        c.drawString(50, y, f"• Photos jointes / Attached photos ({nb_photos_count}) :")
        y -= 18
        
        col_x = 50
        
        for raw_bytes in photos_bytes_list:
            try:
                img_io = io.BytesIO(raw_bytes)
                pil_img = PILImage.open(img_io)
                
                buf = io.BytesIO()
                pil_img.convert('RGB').save(buf, format='JPEG')
                buf.seek(0)
                
                if y - max_h < 40:
                    c.showPage()
                    y = 730
                    col_x = 50
                    dessiner_cadre_et_logo(c, logo_clean_buf)
                
                c.drawImage(ImageReader(buf), col_x, y - max_h, width=max_w, height=max_h, preserveAspectRatio=True)
                
                if col_x == 50:
                    col_x = 295
                else:
                    col_x = 50
                    y -= (max_h + 15)
            except Exception:
                pass

    c.save()
    buffer.seek(0)
    return buffer.getvalue()

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Royal Art Service - Operation report", page_icon="✈️", layout="wide")

logo_trouve = chercher_logo()

if logo_trouve:
    col_t1, col_t2 = st.columns([2, 3])
    with col_t1:
        st.title("✈️ ROYAL ART SERVICE - Operation report")
        st.subheader("Génération, Numérotation et Archivage Automatique")
    with col_t2:
        st.image(logo_trouve, width=840)
else:
    st.title("✈️ ROYAL ART SERVICE - Operation report")
    st.subheader("Génération, Numérotation et Archivage Automatique")

st.write("---")

col_top1, col_top2 = st.columns([1, 2])
with col_top1:
    collaborateur = st.selectbox("Superviseur / Supervisor", ["Miki", "Guillaume", "Simeone", "Sacha", "Perrine", "Quentin"])

with col_top2:
    pass

# --- GESTION INSTANTANÉE DU PRÉFIXE LTA SELON LE NUMÉRO DE VOL ---
if "vol_input" not in st.session_state:
    st.session_state["vol_input"] = ""
if "lta_input" not in st.session_state:
    st.session_state["lta_input"] = ""
if "last_vol_checked" not in st.session_state:
    st.session_state["last_vol_checked"] = ""

current_vol = st.session_state["vol_input"].strip()
if current_vol != st.session_state["last_vol_checked"]:
    st.session_state["last_vol_checked"] = current_vol
    if len(current_vol) >= 2:
        code_comp = current_vol[:2].upper()
        if code_comp in PREFIXES_LTA:
            st.session_state["lta_input"] = PREFIXES_LTA[code_comp]
            st.rerun()

col_vol1, col_vol2 = st.columns([1, 3])
with col_vol1:
    vol_ref = st.text_input("Numéro de vol / Flight N°", placeholder="ex: DL214", key="vol_input")
with col_vol2:
    pass

lta = st.text_input("LTA / AWB", placeholder="ex: 006-12345678", key="lta_input")

st.write("### Photos de Supervision / Supervision Photos")
uploaded_photos = st.file_uploader("Ajouter des photos / Upload photos (tarmac, état des caisses, arrimage...)", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="photos_input")

# --- SELECTEUR DE PALETTES EN DEHORS DU FORMULAIRE ---
st.write("---")
col_s5_1, col_s5_2 = st.columns([3, 1])
with col_s5_1:
    st.write("### Informations et Positions des Palettes / Pallets Info & Positions")
with col_s5_2:
    nb_palettes = st.number_input("Palettes Nb", min_value=1, max_value=15, value=1, step=1, key="nb_palettes_input")

with st.form("form_rapport"):
    col1, col2 = st.columns(2)
    
    with col1:
        date_op = st.date_input("Date de l'opération / Date of operation", datetime.today(), format="DD/MM/YYYY")
        convoyeur = st.text_input("Convoyeur / Courier", placeholder="ex: Sylvie Bourrat")
        
    with col2:
        etd_deplacement = st.time_input("ETD (Heure prévue) / Scheduled Departure", value=datetime.now().time())
        heure_decollage = st.time_input("Heure de décollage (Réelle) / Wheels up", value=datetime.now().time())

    # --- CALCUL ET AFFICHAGE EN DIRECT DU RETARD DANS LE FORMULAIRE ---
    t1_form = datetime.strptime(etd_deplacement.strftime("%H:%M"), "%H:%M")
    t2_form = datetime.strptime(heure_decollage.strftime("%H:%M"), "%H:%M")
    diff_min_form = int((t2_form - t1_form).total_seconds() / 60)
    if diff_min_form < 0:
        diff_min_form = 0
    if diff_min_form > 0:
        h_f = diff_min_form // 60
        m_f = diff_min_form % 60
        retard_form_txt = f"+{h_f}h{m_f:02d}" if h_f > 0 and m_f > 0 else (f"+{h_f}h" if h_f > 0 else f"+{diff_min_form} min")
    else:
        retard_form_txt = "0 min (À l'heure / On time)"
        
    st.markdown(f"⏱️ **Retard calculé en direct :** :red[**{retard_form_txt}**]")

    st.write("---")
    
    palettes_data = []
    for i in range(int(nb_palettes)):
        col_p1, col_p2, col_p3, col_p4 = st.columns([2, 1.5, 1.5, 2])
        with col_p1:
            num_p = st.text_input(f"Palette / ULD {i+1}", value="", placeholder="ex: PMC25432AF ou PAG...", key=f"p_{i}")
        with col_p2:
            nb_c = st.number_input(f"Caisses / Crates ({i+1})", min_value=1, value=1, key=f"c_{i}")
        with col_p3:
            pos_num = st.text_input(f"Position N° ({i+1})", value="", placeholder="ex: 12P", key=f"pos_num_{i}")
        with col_p4:
            pos_zone = st.selectbox(f"Zone ({i+1})", ["FRONT", "REAR", "UPPER DECK", "BULK"], key=f"pos_zone_{i}")
            
        pos_clean_str = f"{pos_num.strip().upper()} - {pos_zone.strip().upper()}" if pos_num.strip() else ""
        palettes_data.append({"num_palette": num_p if num_p else "N/A", "position": pos_clean_str, "nb_caisses": nb_c, "pos_clean": pos_clean_str})

    st.write("---")
    st.write("### Transport en Cabine / Hand Carry (Si applicable / If applicable)")
    siege_cabine = st.text_input("Numéro de siège / Seat N°", placeholder="ex: 2A")

    st.write("---")
    commentaires = st.text_area("Remarques / Compte-rendu - Remarks / Report", placeholder="Précisez le déroulement de l'opération, la tension des filets...")
    
    submit = st.form_submit_button("📁 Enregistrer et Générer PDF / Save & Generate PDF")

if submit:
    if not vol_ref:
        st.markdown("**:red[Veuillez renseigner le numéro de vol ou la référence / Please specify flight number or reference.]**")
    else:
        positions_vues = []
        doublon_detecte = False
        for p in palettes_data:
            cle_pos = p['pos_clean']
            if cle_pos:
                if cle_pos in positions_vues:
                    doublon_detecte = True
                    break
                positions_vues.append(cle_pos)
            
        if doublon_detecte:
            st.markdown("**:red[Erreur : Deux palettes ou plus ne peuvent pas occuper exactement la même position et la même zone (ex: 12P en FRONT) / Two or more ULDs cannot occupy the exact same position and zone.]**")
        else:
            id_rapport = obtenir_prochain_id()
            
            photos_bytes_list = []
            if uploaded_photos:
                folder_photos = os.path.join(BASE_DIR, f"photos_{id_rapport}")
                os.makedirs(folder_photos, exist_ok=True)
                for img in uploaded_photos:
                    b_data = img.getvalue()
                    photos_bytes_list.append(b_data)
                    with open(os.path.join(folder_photos, img.name), "wb") as f:
                        f.write(b_data)

            palettes_txt = " | ".join([f"Palette {p['num_palette']} (Pos: {p['position'] if p['position'] else 'N/A'}): {p['nb_caisses']} caisse(s)" for p in palettes_data])
            date_op_fr = date_op.strftime("%d/%m/%Y")
            etd_txt = etd_deplacement.strftime("%H:%M")
            heure_dec_txt = heure_decollage.strftime("%H:%M")
            
            # Calcul automatique du retard pour l'enregistrement et le PDF
            t1 = datetime.strptime(etd_txt, "%H:%M")
            t2 = datetime.strptime(heure_dec_txt, "%H:%M")
            diff_minutes = int((t2 - t1).total_seconds() / 60)
            if diff_minutes < 0:
                diff_minutes = 0
            
            if diff_minutes > 0:
                heures = diff_minutes // 60
                minutes = diff_minutes % 60
                if heures > 0:
                    retard_txt = f"+{heures}h{minutes:02d}" if minutes > 0 else f"+{heures}h"
                else:
                    retard_txt = f"+{diff_minutes} min"
            else:
                retard_txt = "0 min"

            nb_photos_count = len(photos_bytes_list)

            lta_valeur = lta.strip() if lta else ""
            if not lta_valeur and vol_ref:
                c_comp = vol_ref.strip()[:2].upper()
                lta_valeur = PREFIXES_LTA.get(c_comp, "")

            nouveau_rapport = {
                "ID_Rapport": id_rapport,
                "Date_Saisie": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "Superviseur": collaborateur,
                "Date_Operation": date_op_fr,
                "Vol_Ref": vol_ref,
                "Convoyeur": convoyeur,
                "LTA": lta_valeur,
                "Palettes_Detail": palettes_txt,
                "Siege_Cabine": siege_cabine if siege_cabine else "N/A",
                "ETD": etd_txt,
                "Heure_Decollage": heure_dec_txt,
                "Retard": retard_txt,
                "Commentaires": commentaires,
                "Nb_Photos": nb_photos_count
            }
            df_existant = pd.read_excel(EXCEL_FILE)
            df_updated = pd.concat([df_existant, pd.DataFrame([nouveau_rapport])], ignore_index=True)
            df_updated.to_excel(EXCEL_FILE, index=False)
            
            try:
                pdf_bytes = generer_pdf_bytes(id_rapport, collaborateur, date_op_fr, vol_ref, convoyeur, lta_valeur, etd_txt, heure_dec_txt, retard_txt, palettes_data, siege_cabine, commentaires, photos_bytes_list)
                
                st.markdown(f"### :red[**Rapport N° {id_rapport} enregistré avec succès ! Retard calculé : {retard_txt}**]")
                
                st.download_button(
                    label="📥 Télécharger le Rapport PDF / Download PDF Report",
                    data=pdf_bytes,
                    file_name=f"Rapport_{id_rapport}.pdf",
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"Erreur lors de la génération du PDF : {e}")















































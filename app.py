import io
import json
import time
import zipfile
import xml.etree.ElementTree as ET
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from google import genai
from google.genai import types
import pypdf
from pptx import Presentation
import urllib.request

# Configuration de la page
st.set_page_config(
    page_title="Portail Pédagogique - Physique-Chimie Collège",
    page_icon="⚗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Clé API invisible via Secrets
api_key = st.secrets.get("GEMINI_API_KEY", "")

# Initialisation Session State
if "authentifie" not in st.session_state:
    st.session_state.authentifie = False
if "prof_nom_connecte" not in st.session_state:
    st.session_state.prof_nom_connecte = ""
if "fiches_generees" not in st.session_state:
    st.session_state.fiches_generees = []
if "nom_lecon_global" not in st.session_state:
    st.session_state.nom_lecon_global = ""

@st.cache_data(show_spinner=False)
def charger_logo_ministere():
    """Télécharge l'image officielle côté serveur pour contourner tout blocage de navigateur."""
    urls = [
        "https://raw.githubusercontent.com/abdelkrim/maroc-data/master/logos/men.png",
        "https://upload.wikimedia.org/wikipedia/commons/thumb/f/f7/Logo_MEN_Maroc.svg/500px-Logo_MEN_Maroc.svg.png"
    ]
    for url in urls:
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=5) as response:
                return response.read()
        except Exception:
            continue
    return None

def verifier_acces():
    nom_saisi = st.session_state.get("nom_prof_input", "").strip()
    mdp_saisi = st.session_state.get("mdp_input", "").strip()

    if not nom_saisi:
        st.error("Veuillez saisir votre nom d'enseignant.")
        return

    nom_sans_espace = "".join(nom_saisi.split()).lower()
    mdp_attendu = f"{nom_sans_espace}2026@"

    if mdp_saisi == mdp_attendu or mdp_saisi == "Prof2026@":
        st.session_state.authentifie = True
        st.session_state.prof_nom_connecte = nom_saisi
    else:
        st.error("Mot de passe incorrect. Le mot de passe attendu est votre nom sans espace suivi de '2026@'.")

# --- PAGE D'AUTHENTIFICATION AVEC LOGO OFFICIEL ROBUSTE ---
if not st.session_state.authentifie:
    col_c1, col_c2, col_c3 = st.columns([1, 1.2, 1])
    with col_c2:
        logo_data = charger_logo_ministere()
        if logo_data:
            # Fond blanc net pour faire ressortir les armoiries et la calligraphie sur thème sombre
            st.markdown(
                '<div style="background-color: #FFFFFF; padding: 14px; border-radius: 12px; margin-bottom: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.3); text-align: center;">',
                unsafe_allow_html=True
            )
            st.image(logo_data, width=280)
            st.markdown('</div>', unsafe_allow_html=True)
            
        st.markdown(
            '<div style="text-align: center; margin-bottom: 25px;">'
            '<h2 style="color: #4A90E2; margin-top: 5px; margin-bottom: 4px; font-weight: 700;">Portail Pédagogique de Physique-Chimie</h2>'
            '<p style="color: #A0AAB5; font-size: 14.5px;">Conforme aux Orientations et Programmes Annuels du Secondaire Collégial (Maroc)</p>'
            '</div>',
            unsafe_allow_html=True
        )
        
        st.markdown("#### 🔒 Authentification de l'Enseignant")
        st.text_input("Nom de l'enseignant :", key="nom_prof_input")
        st.text_input("Mot de passe :", type="password", key="mdp_input")
        st.caption("ℹ️ *Règle : votre mot de passe est votre nom sans espace + 2026@*")
        st.button("Accéder au Générateur ➔", type="primary", use_container_width=True, on_click=verifier_acces)
    st.stop()

# --- EXTRACTION PPTX ROBUSTE ---
def extraire_texte_pptx(uploaded_file):
    textes = []
    try:
        uploaded_file.seek(0)
        prs = Presentation(uploaded_file)
        for i, slide in enumerate(prs.slides[:25]):
            slide_txt = []
            for shape in slide.shapes:
                try:
                    if shape.has_text_frame:
                        for p in shape.text_frame.paragraphs:
                            t = p.text.strip()
                            if t and t not in slide_txt:
                                slide_txt.append(t)
                except Exception:
                    continue
            if slide_txt:
                textes.append(f"[Diapo {i+1}] " + " | ".join(slide_txt))
    except Exception:
        uploaded_file.seek(0)
        with zipfile.ZipFile(uploaded_file) as zf:
            slide_files = sorted([f for f in zf.namelist() if f.startswith("ppt/slides/slide") and f.endswith(".xml")])
            for i, sf in enumerate(slide_files[:25]):
                xml_content = zf.read(sf)
                tree = ET.fromstring(xml_content)
                slide_txt = []
                for node in tree.iter():
                    if node.tag.endswith('}t') and node.text:
                        val = node.text.strip()
                        if val and val not in slide_txt:
                            slide_txt.append(val)
                if slide_txt:
                    textes.append(f"[Diapo {i+1}] " + " | ".join(slide_txt))
    
    return "\n".join(textes)

# --- EXTRACTION MULTI-FORMATS ---
def extraire_texte(uploaded_file):
    nom = uploaded_file.name.lower()
    texte = ""
    if nom.endswith(".pptx"):
        texte = extraire_texte_pptx(uploaded_file)
    elif nom.endswith(".pdf"):
        uploaded_file.seek(0)
        reader = pypdf.PdfReader(uploaded_file)
        texte = "\n".join([page.extract_text() or "" for page in reader.pages[:15]])
    elif nom.endswith(".docx"):
        uploaded_file.seek(0)
        doc = Document(uploaded_file)
        texte = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    elif nom.endswith(".txt"):
        uploaded_file.seek(0)
        texte = uploaded_file.read().decode("utf-8", errors="ignore")
    return texte

# --- FORMATAGE WORD AUX STANDARDS OFFICIELS ---
def appliquer_arriere_plan(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'))

def configurer_marges_cellule(cell, top=70, bottom=70, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}>'
                      f'<w:top w:w="{top}" w:type="dxa"/>'
                      f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
                      f'<w:left w:w="{left}" w:type="dxa"/>'
                      f'<w:right w:w="{right}" w:type="dxa"/>'
                      f'</w:tcMar>')
    tcPr.append(tcMar)

def ecrire_cellule(cell, text, gras=False, couleur=(0, 0, 0), taille=9, align=WD_ALIGN_PARAGRAPH.LEFT, fond=None):
    cell.text = ""
    lignes = str(text).strip().split("\n")
    for i, ligne in enumerate(lignes):
        ligne_propre = ligne.strip()
        if not ligne_propre:
            continue
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(1.5)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.1
        
        run = p.add_run(ligne_propre)
        run.bold = gras
        run.font.name = "Calibri"
        run.font.size = Pt(taille)
        run.font.color.rgb = RGBColor(*couleur)
        
    configurer_marges_cellule(cell)
    if fond:
        appliquer_arriere_plan(cell, fond)

def generer_document_docx_officiel(data):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.47)
        section.bottom_margin = Inches(0.47)
        section.left_margin = Inches(0.47)
        section.right_margin = Inches(0.47)

    # 1. En-tête
    t_header = doc.add_table(rows=2, cols=3)
    t_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_header.style = 'Table Grid'
    
    col_w_header = [Inches(2.2), Inches(3.1), Inches(2.2)]
    prof_nom = data.get('enseignant', '').strip() or "........................"
    titre_lecon = data.get('titre_lecon', '').strip() or "Leçon : Physique-Chimie"
    
    ecrire_cellule(t_header.rows[0].cells[0], f"Séance : {data.get('seance', 'Séance 1/1')}\nNiveau : {data.get('niveau', '3ème AC')}", gras=True, fond="F2F2F2")
    ecrire_cellule(t_header.rows[0].cells[1], "Fiche Pédagogique", gras=True, taille=13, align=WD_ALIGN_PARAGRAPH.CENTER, fond="E8EEF5", couleur=(23, 54, 93))
    ecrire_cellule(t_header.rows[0].cells[2], f"Prof : {prof_nom}", gras=True, align=WD_ALIGN_PARAGRAPH.RIGHT, fond="F2F2F2")
    
    cell_lecon = t_header.rows[1].cells[0].merge(t_header.rows[1].cells[1]).merge(t_header.rows[1].cells[2])
    ecrire_cellule(cell_lecon, titre_lecon, gras=True, taille=11, align=WD_ALIGN_PARAGRAPH.CENTER, fond="D9E1F2", couleur=(23, 54, 93))

    for row in t_header.rows:
        for idx, w in enumerate(col_w_header):
            if idx < len(row.cells):
                row.cells[idx].width = w

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # 2. Cadrage Pédagogique
    t_cadre = doc.add_table(rows=2, cols=2)
    t_cadre.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_cadre.style = 'Table Grid'
    t_cadre.rows[0].cells[0].width = Inches(4.5)
    t_cadre.rows[0].cells[1].width = Inches(3.0)
    t_cadre.rows[1].cells[0].width = Inches(4.5)
    t_cadre.rows[1].cells[1].width = Inches(3.0)

    txt_obj = "Objectifs (Connaissances et Capacités) :\n" + "\n".join([f"• {o}" for o in data.get("objectifs", [])])
    ecrire_cellule(t_cadre.rows[0].cells[0], txt_obj)
    
    txt_question = f"Question de la Séance :\n{data.get('question_seance', '')}"
    ecrire_cellule(t_cadre.rows[0].cells[1], txt_question, gras=True, fond="F9FBFD")

    txt_prerequis = "Prérequis :\n" + "\n".join([f"- {p}" for p in data.get("prerequis", [])])
    ecrire_cellule(t_cadre.rows[1].cells[0], txt_prerequis)

    txt_concepts = "Concepts clés :\n" + "\n".join([f"- {c}" for c in data.get("concepts", [])])
    ecrire_cellule(t_cadre.rows[1].cells[1], txt_concepts)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # 3. Tableau Activités (6 colonnes)
    activites = data.get("activites", [])
    t_act = doc.add_table(rows=1 + len(activites), cols=6)
    t_act.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_act.style = 'Table Grid'

    col_widths = [Inches(1.8), Inches(1.1), Inches(1.6), Inches(1.6), Inches(0.55), Inches(0.85)]
    entetes = [
        "Bilan de chaque activité\n(Résumé / Notions)",
        "Supports",
        "Activité de l'élève",
        "Activité du prof",
        "Durée",
        "Activités Interactives"
    ]
    for idx, h in enumerate(entetes):
        ecrire_cellule(t_act.rows[0].cells[idx], h, gras=True, taille=8.5, fond="1F4E79", couleur=(255, 255, 255), align=WD_ALIGN_PARAGRAPH.CENTER)

    for idx, act in enumerate(activites):
        row = t_act.rows[1 + idx]
        ecrire_cellule(row.cells[0], act.get("bilan_contenu", ""), taille=8.5)
        ecrire_cellule(row.cells[1], act.get("supports", ""), taille=8.5)
        ecrire_cellule(row.cells[2], act.get("activite_eleve", ""), taille=8.5)
        ecrire_cellule(row.cells[3], act.get("activite_prof", ""), taille=8.5)
        ecrire_cellule(row.cells[4], act.get("duree", "15 min"), taille=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        desc_inter = f"{act.get('type_etape', '')}\n\n{act.get('questions_interactives', '')}"
        ecrire_cellule(row.cells[5], desc_inter, gras=True, taille=8, align=WD_ALIGN_PARAGRAPH.CENTER, fond="F2F2F2")

    for row in t_act.rows:
        for idx, w in enumerate(col_widths):
            row.cells[idx].width = w

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # 4. Évaluation & Remédiation
    t_eval = doc.add_table(rows=2, cols=3)
    t_eval.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_eval.style = 'Table Grid'
    
    w_eval = [Inches(2.7), Inches(2.7), Inches(2.1)]
    titres_eval = ["Connaissances évaluables :", "Capacités évaluables :", "Obstacles rencontrés & Remédiation :"]
    for idx, t in enumerate(titres_eval):
        ecrire_cellule(t_eval.rows[0].cells[idx], t, gras=True, fond="D9E1F2", couleur=(23, 54, 93))

    ecrire_cellule(t_eval.rows[1].cells[0], data.get("connaissances_evaluables", ""))
    ecrire_cellule(t_eval.rows[1].cells[1], data.get("capacites_evaluables", ""))
    ecrire_cellule(t_eval.rows[1].cells[2], data.get("obstacles_remediation", "Difficulté d'abstraction / Remédiation par modélisation."))

    for row in t_eval.rows:
        for idx, w in enumerate(w_eval):
            row.cells[idx].width = w

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# --- INTERFACE ENSEIGNANT CONNECTÉ ---
st.markdown("<h2 style='color:#17365D;'>⚗️ Générateur de Fiches Pédagogiques de Physique-Chimie</h2>", unsafe_allow_html=True)
st.caption(f"Enseignant(e) connecté(e) : **{st.session_state.prof_nom_connecte}** — Découpage automatique des séances selon le Programme Annuel (Maroc)")

with st.sidebar:
    st.header("📋 Paramètres de la séance")
    ens_nom = st.text_input("Professeur :", value=st.session_state.prof_nom_connecte)
    niveau_select = st.selectbox(
        "Niveau scolaire :",
        ["3ème AC (3ème Année Collège)", "2ème AC (2ème Année Collège)", "1ère AC (1ère Année Collège)"],
        index=0
    )
    titre_manuel = st.text_input("Intitulé / Leçon (Optionnel) :", value="")
    
    st.divider()
    if st.button("🚪 Se déconnecter", use_container_width=True):
        st.session_state.authentifie = False
        st.session_state.prof_nom_connecte = ""
        st.session_state.fiches_generees = []
        st.session_state.nom_lecon_global = ""
        st.rerun()

col_u, col_t = st.columns([1.1, 0.9])

with col_u:
    st.markdown("#### 1. Support de cours officiel")
    fichier_cours = st.file_uploader(
        "Déposer le support de cours (PowerPoint .pptx, PDF, Word .docx, Texte) :",
        type=["pptx", "pdf", "docx", "txt"]
    )

with col_t:
    st.markdown("#### 2. Ou coller des remarques didactiques / résumé")
    texte_libre = st.text_area("Notes sur le contenu / activités :", height=135)

contenu_source = ""
if fichier_cours is not None:
    try:
        contenu_source = extraire_texte(fichier_cours)
        if contenu_source.strip():
            st.success(f"✅ Document '{fichier_cours.name}' analysé avec succès !")
        else:
            st.warning("⚠️ Document chargé, mais aucun texte lisible extrait. Vous pouvez coller le texte à droite.")
    except Exception as e:
        st.error(f"Erreur de lecture du document : {e}")
elif texte_libre.strip():
    contenu_source = texte_libre.strip()

st.divider()

if st.button("🚀 Générer la / les Fiche(s) Pédagogique(s) Officielle(s)", type="primary", use_container_width=True):
    if not api_key:
        st.error("Clé API non configurée. Veuillez vérifier les Secrets dans Streamlit.")
    elif not contenu_source.strip():
        st.warning("Veuillez fournir un support de cours (fichier ou texte) avant de lancer la génération.")
    else:
        with st.spinner("Découverte des modèles actifs et génération des fiches par séance..."):
            try:
                client = genai.Client(api_key=api_key)
                
                prompt = f"""
                Tu es un inspecteur pédagogique de Physique-Chimie au Maroc (Enseignement Secondaire Collégial).
                À partir du contenu fourni ci-dessous, identifie la leçon correspondante dans le programme annuel officiel marocain et DÉCOUPES-LA AUTOMATIQUEMENT EN SÉANCES conformes aux directives ministérielles.

                NIVEAU CHOISI : {niveau_select}
                TITRE INDIQUÉ : {titre_manuel.strip() if titre_manuel.strip() else "Détecter selon le contenu"}

                RÉFÉRENTIEL DU PROGRAMME :
                - 1AC : Matière (Eau, 3 états, Changements d'état, Mélanges, Traitement), Électricité (Circuit simple, Montages, Courant continu, Résistance, Lois).
                - 2AC : Matière (Air, Atomes/Molécules, Réactions chimiques), Lumière (Sources, Dispersion, Propagation, Lentilles, Œil), Électricité (Courant alternatif, Installation).
                - 3AC : Matériaux (Matériaux usuels, Atomes et Ions = 2 SÉANCES : Séance 1/2 structure de l'atome, Z, électroneutralité ; Séance 2/2 ions, formules et charges), Réactions avec l'air, Solutions acides/basiques & pH, Mécanique, Électricité.

                STRUCTURE OBLIGATOIRE DE CHAQUE SÉANCE (60 min) :
                1. "Activité Introductive" (10 min)
                2. "Activité Constructive" (30 min)
                3. "BILAN" (20 min)

                Réponds STRICTEMENT par un JSON valide :
                {{
                  "lecon_detectee": "Nom officiel de la leçon",
                  "nombre_seances": 2,
                  "seances": [
                    {{
                      "seance_label": "Séance 1/2",
                      "titre_lecon": "Leçon : ...",
                      "question_seance": "Question posée",
                      "objectifs": ["Connaître...", "Savoir calculer..."],
                      "prerequis": ["Prérequis 1", "Prérequis 2"],
                      "concepts": ["Concept 1", "Concept 2"],
                      "activites": [
                        {{
                          "type_etape": "Activité Introductive",
                          "duree": "10 min",
                          "bilan_contenu": "• Poser la question de la séance\\n• Émettre des hypothèses",
                          "supports": "- Tableau\\n- Documents",
                          "activite_eleve": "- Répondre et émettre des hypothèses",
                          "activite_prof": "- Poser la situation-problème",
                          "questions_interactives": "Discussion de départ."
                        }},
                        {{
                          "type_etape": "Activité Constructive",
                          "duree": "30 min",
                          "bilan_contenu": "Résumé notionnel...",
                          "supports": "- Matériel / Étiquettes / Fiches",
                          "activite_eleve": "- Observer et déduire",
                          "activite_prof": "- Guider l'investigation",
                          "questions_interactives": "Questions clés."
                        }},
                        {{
                          "type_etape": "BILAN",
                          "duree": "20 min",
                          "bilan_contenu": "Résumé institutionnel et exercice d'application.",
                          "supports": "- Cahier de cours",
                          "activite_eleve": "- Noter la synthèse",
                          "activite_prof": "- Structurer la réponse",
                          "questions_interactives": "Évaluation formative."
                        }}
                      ],
                      "connaissances_evaluables": "...",
                      "capacites_evaluables": "...",
                      "obstacles_remediation": "..."
                    }}
                  ]
                }}

                Support de cours à analyser :
                \"\"\"{contenu_source[:4000]}\"\"\"
                """

                modeles_valides = []
                try:
                    for m in client.models.list():
                        clean_name = m.name.replace("models/", "")
                        methods = getattr(m, "supported_generation_methods", []) or getattr(m, "supported_actions", [])
                        if not methods or "generateContent" in methods:
                            if "embed" not in clean_name.lower():
                                modeles_valides.append(clean_name)
                except Exception:
                    pass

                modeles_tries = []
                for mod in modeles_valides:
                    if "flash" in mod.lower():
                        modeles_tries.insert(0, mod)
                    else:
                        modeles_tries.append(mod)

                if not modeles_tries:
                    modeles_tries = ["gemini-3.8-flash", "gemini-3.1-pro-preview"]

                reponse = None
                derniere_err = None

                for mod in modeles_tries:
                    for essai in range(2):
                        try:
                            reponse = client.models.generate_content(
                                model=mod,
                                contents=prompt,
                                config=types.GenerateContentConfig(
                                    response_mime_type="application/json",
                                    temperature=0.2
                                )
                            )
                            if reponse and reponse.text:
                                break
                        except Exception as e:
                            derniere_err = e
                            err_str = str(e)
                            if "429" in err_str:
                                time.sleep(4.0)
                            else:
                                break
                    if reponse and reponse.text:
                        break

                if reponse is None or not reponse.text:
                    raise Exception(f"Erreur d'accès à l'API : {derniere_err}")

                resultat_json = json.loads(reponse.text)
                liste_seances = resultat_json.get("seances", [])
                st.session_state.nom_lecon_global = resultat_json.get("lecon_detectee", titre_manuel or "Leçon Physique-Chimie")

                fichiers_prepares = []
                for idx, fiche_data in enumerate(liste_seances):
                    fiche_data["enseignant"] = ens_nom.strip()
                    fiche_data["niveau"] = niveau_select.split(" ")[0]
                    fiche_data["seance"] = fiche_data.get("seance_label", f"Séance {idx+1}/{len(liste_seances)}")
                    if not fiche_data.get("titre_lecon"):
                        fiche_data["titre_lecon"] = st.session_state.nom_lecon_global

                    docx_bytes = generer_document_docx_officiel(fiche_data)
                    seance_clean = fiche_data['seance'].replace('/', '-').replace(' ', '_')
                    nom_f = f"{fiche_data['niveau']}_{seance_clean}_{st.session_state.nom_lecon_global.replace(' ', '_')}.docx"
                    
                    fichiers_prepares.append({
                        "nom_fichier": nom_f,
                        "bytes": docx_bytes,
                        "seance": fiche_data['seance'],
                        "question": fiche_data.get('question_seance', 'Fiche technique')
                    })

                st.session_state.fiches_generees = fichiers_prepares

            except Exception as e:
                st.error(f"Une erreur est survenue pendant la génération : {e}")

# --- AFFICHAGE PERSISTANT DES FICHES GÉNÉRÉES ---
if st.session_state.fiches_generees:
    st.success(f"🎉 Leçon : **{st.session_state.nom_lecon_global}** ({len(st.session_state.fiches_generees)} séance(s) prêtes)")

    for idx, f in enumerate(st.session_state.fiches_generees):
        col_card, col_btn = st.columns([2.5, 1])
        with col_card:
            st.markdown(f"**📄 {f['seance']} :** {f['question']}")
        with col_btn:
            st.download_button(
                label=f"📥 Télécharger {f['seance']} (.DOCX)",
                data=f["bytes"],
                file_name=f["nom_fichier"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"dl_persistant_{idx}",
                use_container_width=True
            )

    if len(st.session_state.fiches_generees) > 1:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for f in st.session_state.fiches_generees:
                zip_file.writestr(f["nom_fichier"], f["bytes"])
        
        st.divider()
        st.download_button(
            label="📦 Télécharger toutes les fiches de la leçon (Pack ZIP complet)",
            data=zip_buffer.getvalue(),
            file_name=f"Fiches_{st.session_state.nom_lecon_global.replace(' ', '_')}.zip",
            mime="application/zip",
            type="primary",
            key="dl_zip_persistant",
            use_container_width=True
        )

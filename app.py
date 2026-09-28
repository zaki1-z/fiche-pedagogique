import io
import json
import time
import zipfile
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

# Configuration de la page
st.set_page_config(
    page_title="Portail Pédagogique - Physique-Chimie Collège",
    page_icon="⚗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Récupération sécurisée et invisible de la clé API depuis Streamlit Secrets
api_key = st.secrets.get("GEMINI_API_KEY", "")

# Initialisation de la mémoire de session
if "authentifie" not in st.session_state:
    st.session_state.authentifie = False
if "prof_nom_connecte" not in st.session_state:
    st.session_state.prof_nom_connecte = ""
if "fiches_generees" not in st.session_state:
    st.session_state.fiches_generees = []
if "nom_lecon_global" not in st.session_state:
    st.session_state.nom_lecon_global = ""

# --- LOGIQUE D'AUTHENTIFICATION AVEC NOM ET MOT DE PASSE DYNAMIQUE ---
def verifier_acces():
    nom_saisi = st.session_state.get("nom_prof_input", "").strip()
    mdp_saisi = st.session_state.get("mdp_input", "").strip()

    if not nom_saisi:
        st.error("Veuillez saisir votre nom d'enseignant.")
        return

    # Mot de passe dynamique : nom sans espace en minuscules + 2026@
    nom_sans_espace = "".join(nom_saisi.split()).lower()
    mdp_attendu = f"{nom_sans_espace}2026@"

    if mdp_saisi == mdp_attendu or mdp_saisi == "Prof2026@":
        st.session_state.authentifie = True
        st.session_state.prof_nom_connecte = nom_saisi
    else:
        st.error("Mot de passe incorrect. Le mot de passe attendu est votre nom sans espace suivi de '2026@'.")

if not st.session_state.authentifie:
    st.markdown("""
        <div style="text-align: center; margin-top: 30px; margin-bottom: 20px;">
            <div style="display: flex; justify-content: center; align-items: center; gap: 15px; margin-bottom: 12px;">
                <svg width="75" height="75" viewBox="0 0 100 100">
                    <circle cx="50" cy="50" r="46" fill="none" stroke="#1F4E79" stroke-width="3" stroke-dasharray="3,2"/>
                    <circle cx="50" cy="50" r="41" fill="none" stroke="#1F4E79" stroke-width="1.5"/>
                    <text x="50" y="27" font-size="8.5" font-family="Arial, sans-serif" font-weight="bold" fill="#1F4E79" text-anchor="middle">المملكة المغربية</text>
                    <text x="50" y="38" font-size="7.5" font-family="Arial, sans-serif" fill="#1F4E79" text-anchor="middle">وزارة التربية الوطنية</text>
                    <text x="50" y="47" font-size="6.5" font-family="Arial, sans-serif" fill="#1F4E79" text-anchor="middle">والتعليم الأولي والرياضة</text>
                    <path d="M 50 54 L 54 62 L 63 62 L 56 67 L 59 75 L 50 70 L 41 75 L 44 67 L 37 62 L 46 62 Z" fill="none" stroke="#1F4E79" stroke-width="1.8"/>
                    <text x="50" y="87" font-size="6.5" font-family="Arial, sans-serif" font-weight="bold" fill="#1F4E79" text-anchor="middle">Royaume du Maroc</text>
                </svg>
            </div>
            <h2 style="color: #17365D; margin-bottom: 4px;">المملكة المغربية - وزارة التربية الوطنية</h2>
            <h3 style="color: #1F4E79; font-weight: normal; margin-top: 0;">Portail Pédagogique de Physique-Chimie (Collège)</h3>
            <p style="color: #666; font-size: 15px;">Conforme aux Orientations et Programmes Annuels du Secondaire Collégial</p>
        </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("#### 🔒 Authentification de l'Enseignant")
        st.text_input("Nom de l'enseignant :", key="nom_prof_input")
        st.text_input("Mot de passe :", type="password", key="mdp_input")
        st.caption("ℹ️ *Règle : votre mot de passe est votre nom sans espace + 2026@*")
        st.button("Accéder au Générateur ➔", type="primary", use_container_width=True, on_click=verifier_acces)
    st.stop()

# --- EXTRACTION MULTI-FORMATS ---
def extraire_texte(uploaded_file):
    nom = uploaded_file.name.lower()
    texte = ""
    if nom.endswith(".pdf"):
        reader = pypdf.PdfReader(uploaded_file)
        texte = "\n".join([page.extract_text() or "" for page in reader.pages])
    elif nom.endswith(".docx"):
        doc = Document(uploaded_file)
        texte = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    elif nom.endswith(".pptx"):
        prs = Presentation(uploaded_file)
        diapos = []
        for i, slide in enumerate(prs.slides):
            textes_slide = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.strip():
                            textes_slide.append(paragraph.text.strip())
            if textes_slide:
                diapos.append(f"[Diapositive {i+1}]\n" + "\n".join(textes_slide))
        texte = "\n\n".join(diapos)
    elif nom.endswith(".txt"):
        texte = uploaded_file.read().decode("utf-8", errors="ignore")
    return texte

# --- OUTILS DE FORMATAGE WORD CONFORMES AUX EXEMPLES ---
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

# --- GÉNÉRATION DU DOCUMENT DOCX ---
def generer_document_docx_officiel(data):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.47)
        section.bottom_margin = Inches(0.47)
        section.left_margin = Inches(0.47)
        section.right_margin = Inches(0.47)

    # 1. En-tête (Séance/Niveau | Titre Fiche | Professeur)
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

    # 2. Tableau Cadrage Pédagogique
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

    # 3. Tableau Principal des Activités (6 colonnes)
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

    # 4. Tableau d'Évaluation & Remédiation
    t_eval = doc.add_table(rows=2, cols=3)
    t_eval.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_eval.style = 'Table Grid'
    
    w_eval = [Inches(2.7), Inches(2.7), Inches(2.1)]
    titres_eval = ["Connaissances évaluables :", "Capacités évaluables :", "Obstacles rencontrés & Remédiation :"]
    for idx, t in enumerate(titres_eval):
        ecrire_cellule(t_eval.rows[0].cells[idx], t, gras=True, fond="D9E1F2", couleur=(23, 54, 93))

    ecrire_cellule(t_eval.rows[1].cells[0], data.get("connaissances_evaluables", ""))
    ecrire_cellule(t_eval.rows[1].cells[1], data.get("capacites_evaluables", ""))
    ecrire_cellule(t_eval.rows[1].cells[2], data.get("obstacles_remediation", "Difficulté d'abstraction / Remédiation par modélisation ou simulation."))

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
        st.success(f"✅ Document '{fichier_cours.name}' analysé avec succès !")
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
        with st.spinner("Analyse didactique et génération des fiches par séance..."):
            try:
                client = genai.Client(api_key=api_key)
                
                prompt = f"""
                Tu es un inspecteur pédagogique national de Physique-Chimie au Maroc (Enseignement Secondaire Collégial).
                Ta mission est d'analyser le document de cours ci-dessous, d'identifier la leçon exacte dans le programme annuel officiel marocain, de DÉTERMINER AUTOMATIQUEMENT LE NOMBRE DE SÉANCES NÉCESSAIRES, et de produire une fiche pédagogique distincte pour CHAQUE séance.

                NIVEAU SÉLECTIONNÉ : {niveau_select}
                TITRE INDIQUE PAR LE PROFESSEUR : {titre_manuel.strip() if titre_manuel.strip() else "À déterminer automatiquement à partir du contenu"}

                RÉFÉRENTIEL DU PROGRAMME OFFICIEL MAROCAIN :
                - 1AC : Matière (L'eau, Trois états, Changements d'état, Mélanges, Traitement), Électricité (Circuit simple, Montages série/dérivation, Courant continu, Résistance, Lois).
                - 2AC : Matière (L'air, Molécules/Atomes, Réactions chimiques/combustions), Lumière (Sources, Dispersion, Propagation, Lentilles minces, Œil), Électricité (Courant alternatif, Installation domestique).
                - 3AC : Matériaux (Matériaux usuels, Atomes et Ions = 2 SÉANCES : Séance 1/2 structure de l'atome, Z, neutralité ; Séance 2/2 ions, formules et charges), Réactions chimiques (Air, Solutions acides/basiques & pH), Mécanique (Mouvement, Actions mécaniques, Forces, Équilibre, Poids/Masse), Électricité (Loi d'Ohm, Puissance, Énergie).

                STRUCTURE DES 3 PHASES OBLIGATOIRES (60 min par séance) :
                1. "Activité Introductive" (10 min) : Rappel des prérequis, question de la séance (situation-problème), formulation des hypothèses.
                2. "Activité Constructive" (30 min) : Investigation documentaire ou expérimentale concrète, analyse, raisonnement, calculs.
                3. "BILAN" (20 min) : Synthèse institutionnelle, résumé de la séance, exercice d'application.

                Format STRICTEMENT attendu (JSON valide uniquement) :
                {{
                  "lecon_detectee": "Nom officiel de la leçon",
                  "nombre_seances": 2,
                  "seances": [
                    {{
                      "seance_label": "Séance 1/2",
                      "titre_lecon": "Leçon : ...",
                      "question_seance": "Question de départ posée aux élèves",
                      "objectifs": [
                        "Connaître...",
                        "Savoir calculer..."
                      ],
                      "prerequis": [
                        "Prérequis 1",
                        "Prérequis 2"
                      ],
                      "concepts": [
                        "Concept clé 1",
                        "Concept clé 2"
                      ],
                      "activites": [
                        {{
                          "type_etape": "Activité Introductive",
                          "duree": "10 min",
                          "bilan_contenu": "• Poser la question de la séance\\n• Émettre des hypothèses",
                          "supports": "- Tableau\\n- Documents",
                          "activite_eleve": "- Répondre aux questions de réactivation\\n- Formuler des hypothèses",
                          "activite_prof": "- Poser la situation-problème\\n- Noter les hypothèses au tableau",
                          "questions_interactives": "Discussion ouverte avec la classe."
                        }},
                        {{
                          "type_etape": "Activité Constructive",
                          "duree": "30 min",
                          "bilan_contenu": "Résumé des notions construites lors de la séance...",
                          "supports": "- Matériel d'expérimentation / Fiches / Étiquettes",
                          "activite_eleve": "- Observer, manipuler, calculer\\n- Dégager la conclusion",
                          "activite_prof": "- Guider l'investigation sans donner directement le résultat",
                          "questions_interactives": "Questions clés de guidage didactique."
                        }},
                        {{
                          "type_etape": "BILAN",
                          "duree": "20 min",
                          "bilan_contenu": "Résumé institutionnel de la séance et exercice d'application résolu.",
                          "supports": "- Tableau\\n- Cahier de cours",
                          "activite_eleve": "- Noter la synthèse et résoudre l'exercice d'application",
                          "activite_prof": "- Structurer la réponse finale et corriger l'exercice",
                          "questions_interactives": "Évaluation formative des acquis."
                        }}
                      ],
                      "connaissances_evaluables": "Connaissances...",
                      "capacites_evaluables": "Capacités...",
                      "obstacles_remediation": "Obstacle et remédiation..."
                    }}
                  ]
                }}

                Support de cours à traiter :
                \"\"\"{contenu_source[:18000]}\"\"\"
                """

                # Détection dynamique des modèles activés sur votre compte pour éviter les erreurs 404
                modeles_prioritaires = [
                    "gemini-3.1-pro-preview",
                    "gemini-3.8-flash",
                    "gemini-3-flash-preview",
                    "gemini-2.0-flash",
                    "gemini-1.5-flash"
                ]
                
                modeles_disponibles = []
                try:
                    for m in client.models.list():
                        clean_name = m.name.replace("models/", "")
                        if "embed" not in clean_name.lower():
                            modeles_disponibles.append(clean_name)
                except Exception:
                    pass

                # Combiner les listes en plaçant les modèles recommandés en tête
                modeles_a_tester = [m for m in modeles_prioritaires if m in modeles_disponibles]
                if not modeles_a_tester:
                    modeles_a_tester = modeles_prioritaires + modeles_disponibles

                reponse = None
                derniere_err = None

                for m in modeles_a_tester:
                    for tentative in range(2):
                        try:
                            reponse = client.models.generate_content(
                                model=m,
                                contents=prompt,
                                config=types.GenerateContentConfig(
                                    response_mime_type="application/json",
                                    temperature=0.2
                                )
                            )
                            if reponse and reponse.text:
                                break
                        except Exception as err:
                            derniere_err = err
                            time.sleep(1.0)
                    if reponse and reponse.text:
                        break

                if reponse is None or not reponse.text:
                    raise Exception(f"Erreur d'accès aux modèles ({derniere_err}).")

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

    # Si plusieurs séances, pack ZIP complet
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

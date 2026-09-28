import io
import json
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

# Mot de passe d'accès pour l'enseignant
MOT_DE_PASSE_VALIDE = "Prof2026@"

# Récupération sécurisée et invisible de la clé API depuis Streamlit Secrets
api_key = st.secrets.get("GEMINI_API_KEY", "")

# Authentification
if "authentifie" not in st.session_state:
    st.session_state.authentifie = False

def verifier_mdp():
    if st.session_state.get("mdp_input") == MOT_DE_PASSE_VALIDE:
        st.session_state.authentifie = True
    else:
        st.error("Mot de passe incorrect. Veuillez réessayer.")

if not st.session_state.authentifie:
    st.markdown("""
        <div style="text-align: center; margin-top: 50px; margin-bottom: 25px;">
            <h1 style="color: #17365D;">🔬 Portail Pédagogique - Physique-Chimie</h1>
            <p style="color: #555; font-size: 16px;">Générateur automatisé de fiches conformes aux Orientations Pédagogiques Officielles Marocaines</p>
        </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("### 🔒 Accès Enseignant")
        st.text_input("Code d'accès enseignant :", type="password", key="mdp_input", on_change=verifier_mdp)
        st.button("Se connecter ➔", type="primary", use_container_width=True, on_click=verifier_mdp)
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

# --- GÉNÉRATION D'UNE FICHE DOCX OFFICIELLE ---
def generer_document_docx_officiel(data):
    doc = Document()
    
    # Marges 1.2 cm
    for section in doc.sections:
        section.top_margin = Inches(0.47)
        section.bottom_margin = Inches(0.47)
        section.left_margin = Inches(0.47)
        section.right_margin = Inches(0.47)

    # 1. En-tête : Tableau 3 colonnes (Séance/Niveau | Titre Fiche | Professeur)
    t_header = doc.add_table(rows=2, cols=3)
    t_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_header.style = 'Table Grid'
    
    col_w_header = [Inches(2.2), Inches(3.1), Inches(2.2)]
    
    prof_nom = data.get('enseignant', '').strip() or "........................"
    titre_lecon = data.get('titre_lecon', '').strip() or "Leçon : Physique-Chimie"
    
    ecrire_cellule(t_header.rows[0].cells[0], f"Séance : {data.get('seance', 'Séance 1/1')}\nNiveau : {data.get('niveau', '3ème AC')}", gras=True, fond="F2F2F2")
    ecrire_cellule(t_header.rows[0].cells[1], "Fiche Pédagogique", gras=True, taille=13, align=WD_ALIGN_PARAGRAPH.CENTER, fond="E8EEF5", couleur=(23, 54, 93))
    ecrire_cellule(t_header.rows[0].cells[2], f"Prof : {prof_nom}", gras=True, align=WD_ALIGN_PARAGRAPH.RIGHT, fond="F2F2F2")
    
    # Ligne 2 : Titre de la leçon centré
    cell_lecon = t_header.rows[1].cells[0].merge(t_header.rows[1].cells[1]).merge(t_header.rows[1].cells[2])
    ecrire_cellule(cell_lecon, titre_lecon, gras=True, taille=11, align=WD_ALIGN_PARAGRAPH.CENTER, fond="D9E1F2", couleur=(23, 54, 93))

    for row in t_header.rows:
        for idx, w in enumerate(col_w_header):
            if idx < len(row.cells):
                row.cells[idx].width = w

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # 2. Tableau Cadrage Pédagogique (Objectifs, Question séance, Prérequis, Concepts)
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

    # 3. Tableau Principal des Activités (6 colonnes conforme aux modèles)
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

    # 4. Tableau d'Évaluation & Remédiation (Pied de page)
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
    buf.seek(0)
    return buf

# --- INTERFACE ENSEIGNANT ÉPURÉE ---
st.markdown("<h2 style='color:#17365D;'>⚗️ Générateur de Fiches Pédagogiques de Physique-Chimie</h2>", unsafe_allow_html=True)
st.caption("Découpage automatique des séances selon le Programme Annuel & Orientations Pédagogiques Officielles (Maroc)")

with st.sidebar:
    st.header("📋 Paramètres de la séance")
    ens_nom = st.text_input("Professeur :", value="", placeholder="Ex: M. / Mme ...")
    niveau_select = st.selectbox(
        "Niveau scolaire :",
        ["3ème AC (3ème Année Collège)", "2ème AC (2ème Année Collège)", "1ère AC (1ère Année Collège)"],
        index=0
    )
    titre_manuel = st.text_input("Intitulé / Leçon (Optionnel) :", value="", placeholder="Laisser vide pour détection automatique")
    
    st.divider()
    if st.button("🚪 Se déconnecter", use_container_width=True):
        st.session_state.authentifie = False
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
    texte_libre = st.text_area("Notes sur le contenu / activités :", height=135, placeholder="Ex: Insister sur la distinction cation/anion et l'activité documentaire sur les étiquettes d'eau minérale...")

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
        with st.spinner("Analyse du programme annuel ministériel et découpage automatique des séances en cours..."):
            try:
                client = genai.Client(api_key=api_key)
                
                prompt = f"""
                Tu es un inspecteur pédagogique national de Physique-Chimie au Maroc (Enseignement Secondaire Collégial).
                Ta mission est d'analyser le document de cours ci-dessous, d'identifier la leçon exacte dans le programme annuel officiel marocain, de DÉTERMINER AUTOMATIQUEMENT LE NOMBRE DE SÉANCES NÉCESSAIRES, et de produire une fiche pédagogique distincte pour CHAQUE séance.

                NIVEAU SÉLECTIONNÉ : {niveau_select}
                TITRE INDIQUE PAR LE PROFESSEUR : {titre_manuel.strip() if titre_manuel.strip() else "À déterminer automatiquement à partir du contenu"}

                RÉFÉRENTIEL DU PROGRAMME OFFICIEL MAROCAIN :
                - 1AC :
                  * Matière : L'eau (2h), Trois états (8h = 4 séances), Changements d'état (4h = 2 séances), Mélanges (4h = 2 séances), Traitement de l'eau (2h = 1 séance).
                  * Électricité : Circuit simple (3h = 1-2 séances), Types de montages (3h), Courant continu (3h), Résistance (3h), Lois des nœuds/tensions (4h = 2 séances), Dangers (3h).
                - 2AC :
                  * Matière : L'air (2h), Propriétés de l'air (1h), Molécules et Atomes (3h = 1-2 séances), Réaction chimique et combustions (10h = 5 séances), Produits naturels/synthétiques (2h), Pollution (2h).
                  * Optique : Lumière (1h), Sources/Récepteurs (2h), Couleurs/Dispersion (2h), Propagation (3h), Applications/Ombres/Éclipses (2h), Lentilles minces (4h = 2 séances), Œil/Loupe (2h).
                  * Électricité : Courant alternatif sinusoïdal (2h = 1 séance), Installation domestique (2h = 1 séance).
                - 3AC :
                  * Matériaux : Exemples de matériaux (2h = 1 séance), Matière et électricité - Atomes et Ions (4h = 2 SÉANCES : Séance 1/2 consacrée à la structure de l'atome, numéro atomique Z et électroneutralité ; Séance 2/2 consacrée aux ions monoatomiques/polyatomiques, formules chimiques et charges), Réactions avec l'air (4h = 2 séances), Réactions avec les solutions acides/basiques & pH (8h = 4 séances), Dangers des matériaux (2h = 1 séance).
                  * Mécanique : Mouvement et repos (5h = 2-3 séances), Actions mécaniques (2h = 1 séance), Notion de force (3h = 1-2 séances), Équilibre sous 2 forces (2h = 1 séance), Poids et Masse (2h = 1 séance).
                  * Électricité : Loi d'Ohm (1h = 1 séance), Puissance électrique (2h = 1 séance), Énergie électrique (3h = 1-2 séances).

                CONSIGNE DE DÉCOUPAGE :
                - Si le support couvre l'ensemble d'une leçon prévue sur plusieurs séances (par exemple « Atomes et Ions » en 3AC = 2 séances), génère une liste de fiches (`seances`) avec 2 éléments : "Séance 1/2" et "Séance 2/2".
                - Si le support ne traite qu'une seule partie ou un thème d'une heure, génère 1 fiche.
                - Chaque fiche doit suivre rigoureusement les 3 phases de 60 minutes :
                  1. "Activité Introductive" (10 min) : Rappel des prérequis, question de la séance (situation-problème), formulation des hypothèses.
                  2. "Activité Constructive" (30 min) : Activité documentaire ou expérimentale concrète avec matériel de collège, raisonnement, calculs.
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
                          "supports": "- Matériel d'expérimentation / Étiquettes / Fiches",
                          "activite_eleve": "- Observer, calculer ou manipuler\\n- Dégager la conclusion",
                          "activite_prof": "- Guider l'investigation sans donner directement le résultat",
                          "questions_interactives": "Questions clés de questionnement didactique."
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

                # Tentative avec le modèle recommandé par l'API puis modèles alternatifs en cascade
                modeles_a_tester = ["gemini-3.8-flash", "gemini-3-flash-preview", "gemini-2.5-flash-preview"]
                reponse = None
                derniere_erreur = None

                for nom_modele in modeles_a_tester:
                    try:
                        reponse = client.models.generate_content(
                            model=nom_modele,
                            contents=prompt,
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json",
                                temperature=0.2
                            )
                        )
                        if reponse and reponse.text:
                            break
                    except Exception as err:
                        derniere_erreur = err
                        continue

                if reponse is None or not reponse.text:
                    raise Exception(f"Impossible de contacter l'API Gemini : {derniere_erreur}")

                resultat_json = json.loads(reponse.text)
                liste_seances = resultat_json.get("seances", [])
                nom_lecon_global = resultat_json.get("lecon_detectee", titre_manuel or "Leçon Physique-Chimie")

                st.success(f"🎉 Analyse terminée ! Leçon identifiée : **{nom_lecon_global}** ({len(liste_seances)} séance(s) générée(s))")

                fichiers_generes = []

                for idx, fiche_data in enumerate(liste_seances):
                    fiche_data["enseignant"] = ens_nom.strip()
                    fiche_data["niveau"] = niveau_select.split(" ")[0]
                    fiche_data["seance"] = fiche_data.get("seance_label", f"Séance {idx+1}/{len(liste_seances)}")
                    if not fiche_data.get("titre_lecon"):
                        fiche_data["titre_lecon"] = nom_lecon_global

                    doc_docx = generer_document_docx_officiel(fiche_data)
                    seance_clean = fiche_data['seance'].replace('/', '-').replace(' ', '_')
                    nom_fichier = f"{fiche_data['niveau']}_{seance_clean}_{nom_lecon_global.replace(' ', '_')}.docx"
                    
                    fichiers_generes.append((nom_fichier, doc_docx))

                    col_card, col_btn = st.columns([2.5, 1])
                    with col_card:
                        st.markdown(f"**📄 {fiche_data['seance']} :** {fiche_data.get('question_seance', 'Fiche technique')}")
                    with col_btn:
                        st.download_button(
                            label=f"📥 Télécharger {fiche_data['seance']} (.DOCX)",
                            data=doc_docx,
                            file_name=nom_fichier,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"btn_dl_{idx}",
                            use_container_width=True
                        )

                # Si plusieurs séances, proposer un pack complet en .ZIP
                if len(fichiers_generes) > 1:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                        for nom_f, buf in fichiers_generes:
                            zip_file.writestr(nom_f, buf.getvalue())
                    zip_buffer.seek(0)
                    
                    st.divider()
                    st.download_button(
                        label="📦 Télécharger toutes les fiches de la leçon (Pack ZIP complet)",
                        data=zip_buffer,
                        file_name=f"Fiches_{nom_lecon_global.replace(' ', '_')}.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Une erreur est survenue pendant la génération : {e}")

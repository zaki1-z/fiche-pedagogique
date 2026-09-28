import io
import json
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
    page_title="Générateur de Fiches Pédagogiques",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Mot de passe de protection (modifiable)
MOT_DE_PASSE_VALIDE = "Prof2026@"

# --- GESTION DE L'AUTHENTIFICATION ---
if "authentifie" not in st.session_state:
    st.session_state.authentifie = False

def verifier_mdp():
    if st.session_state.get("mdp_input") == MOT_DE_PASSE_VALIDE:
        st.session_state.authentifie = True
    else:
        st.error("Mot de passe incorrect. Veuillez réessayer.")

if not st.session_state.authentifie:
    st.markdown("""
        <div style="text-align: center; margin-top: 50px; margin-bottom: 20px;">
            <h1 style="color: #1F4E79;">🎓 Portail Pédagogique</h1>
            <p style="color: #555; font-size: 16px;">Générateur automatisé de fiches de séances — Modèle Enseignement Explicite</p>
        </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("### 🔒 Accès Réservé")
        st.text_input("Saisissez votre code d'accès :", type="password", key="mdp_input", on_change=verifier_mdp)
        st.button("Se connecter ➔", type="primary", use_container_width=True, on_click=verifier_mdp)
        st.caption("Contactez l'administrateur si vous n'avez pas le code.")
    st.stop()

# --- EXTRACTION DE TEXTE MULTI-FORMATS ---
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

# --- FONCTIONS DE MISE EN FORME WORD (.DOCX) ---
def appliquer_arriere_plan(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'))

def configurer_marges_cellule(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}>'
                      f'<w:top w:w="{top}" w:type="dxa"/>'
                      f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
                      f'<w:left w:w="{left}" w:type="dxa"/>'
                      f'<w:right w:w="{right}" w:type="dxa"/>'
                      f'</w:tcMar>')
    tcPr.append(tcMar)

def ecrire_cellule(cell, text, gras=False, couleur=(0, 0, 0), taille=9.5, align=WD_ALIGN_PARAGRAPH.LEFT, fond=None):
    cell.text = ""
    lignes = text.strip().split("\n")
    for i, ligne in enumerate(lignes):
        ligne_propre = ligne.strip()
        if not ligne_propre:
            continue
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        
        run = p.add_run(ligne_propre)
        run.bold = gras
        run.font.name = "Calibri"
        run.font.size = Pt(taille)
        run.font.color.rgb = RGBColor(*couleur)
        
    configurer_marges_cellule(cell)
    if fond:
        appliquer_arriere_plan(cell, fond)

# --- GÉNÉRATION DE LA FICHE PÉDAGOGIQUE EN .DOCX ---
def generer_document_docx(data):
    doc = Document()
    
    # Marges du document (1.5 cm)
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # 1. Bannière Titre
    banniere = doc.add_paragraph()
    banniere.alignment = WD_ALIGN_PARAGRAPH.CENTER
    banniere.paragraph_format.space_after = Pt(6)
    r_titre = banniere.add_run("Fiche technique - Séance d'apprentissage / Remédiation -")
    r_titre.bold = True
    r_titre.font.size = Pt(13)
    r_titre.font.color.rgb = RGBColor(23, 54, 93)

    # 2. Tableau 1 : Cadre Administratif (5 colonnes)
    t1 = doc.add_table(rows=2, cols=5)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.style = 'Table Grid'
    
    entetes_t1 = ["AREF", "Direction provinciale", "Collège / Établissement", "Nom de l'enseignant", "Niveau scolaire"]
    for i, h in enumerate(entetes_t1):
        ecrire_cellule(t1.rows[0].cells[i], h, gras=True, couleur=(255, 255, 255), fond="1F4E79", align=WD_ALIGN_PARAGRAPH.CENTER)

    valeurs_t1 = [
        data.get("aref", ""),
        data.get("direction", ""),
        data.get("etablissement", ""),
        data.get("enseignant", ""),
        data.get("niveau", "")
    ]
    for i, val in enumerate(valeurs_t1):
        ecrire_cellule(t1.rows[1].cells[i], val, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # 3. Tableau 2 : Cadrage Didactique
    t2 = doc.add_table(rows=5, cols=2)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2.style = 'Table Grid'
    
    for row in t2.rows:
        row.cells[0].width = Inches(2.3)
        row.cells[1].width = Inches(5.2)

    champs_t2 = [
        ("Semaine", data.get("semaine", "01")),
        ("Domaine", data.get("domaine", "")),
        ("Séance", data.get("seance", "")),
        ("Tâche", data.get("tache", "")),
        ("Supports didactiques nécessaires", data.get("supports", "Support PPT, ardoises et tableau"))
    ]
    for idx, (label, val) in enumerate(champs_t2):
        ecrire_cellule(t2.rows[idx].cells[0], label, gras=True, fond="D9E1F2")
        ecrire_cellule(t2.rows[idx].cells[1], val)

    # Titre Tableau d'activités
    st_p = doc.add_paragraph()
    st_p.paragraph_format.space_before = Pt(6)
    st_p.paragraph_format.space_after = Pt(4)
    r_st = st_p.add_run("Plan détaillé des activités de la séance")
    r_st.bold = True
    r_st.font.size = Pt(11)
    r_st.font.color.rgb = RGBColor(23, 54, 93)

    # 4. Tableau 3 : Plan détaillé (Étapes, Enseignant, Élève, Temps, Page)
    etapes = data.get("etapes", [])
    t3 = doc.add_table(rows=2 + len(etapes), cols=5)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    t3.style = 'Table Grid'

    col_widths = [Inches(1.2), Inches(3.0), Inches(2.3), Inches(0.6), Inches(0.6)]

    # Fusion d'en-tête pour "Descriptif"
    c_etape = t3.rows[0].cells[0]
    c_desc = t3.rows[0].cells[1].merge(t3.rows[0].cells[2])
    c_temps = t3.rows[0].cells[3]
    c_page = t3.rows[0].cells[4]

    ecrire_cellule(c_etape, "Étapes", gras=True, couleur=(255, 255, 255), fond="1F4E79", align=WD_ALIGN_PARAGRAPH.CENTER)
    ecrire_cellule(c_desc, "Descriptif", gras=True, couleur=(255, 255, 255), fond="1F4E79", align=WD_ALIGN_PARAGRAPH.CENTER)
    ecrire_cellule(c_temps, "Temps", gras=True, couleur=(255, 255, 255), fond="1F4E79", align=WD_ALIGN_PARAGRAPH.CENTER)
    ecrire_cellule(c_page, "Page", gras=True, couleur=(255, 255, 255), fond="1F4E79", align=WD_ALIGN_PARAGRAPH.CENTER)

    ecrire_cellule(t3.rows[1].cells[0], "", fond="D9E1F2")
    ecrire_cellule(t3.rows[1].cells[1], "Rôle de l'enseignant", gras=True, fond="D9E1F2", align=WD_ALIGN_PARAGRAPH.CENTER)
    ecrire_cellule(t3.rows[1].cells[2], "Activité de l'élève", gras=True, fond="D9E1F2", align=WD_ALIGN_PARAGRAPH.CENTER)
    ecrire_cellule(t3.rows[1].cells[3], "", fond="D9E1F2")
    ecrire_cellule(t3.rows[1].cells[4], "", fond="D9E1F2")

    for idx, etape in enumerate(etapes):
        row = t3.rows[2 + idx]
        ecrire_cellule(row.cells[0], etape.get("nom", ""), gras=True, fond="F2F2F2", align=WD_ALIGN_PARAGRAPH.CENTER)
        ecrire_cellule(row.cells[1], etape.get("role_enseignant", ""))
        ecrire_cellule(row.cells[2], etape.get("activite_eleve", ""))
        ecrire_cellule(row.cells[3], etape.get("temps", ""), align=WD_ALIGN_PARAGRAPH.CENTER)
        ecrire_cellule(row.cells[4], etape.get("page", ""), align=WD_ALIGN_PARAGRAPH.CENTER)

    for row in t3.rows:
        for idx, w in enumerate(col_widths):
            row.cells[idx].width = w

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- APPLICATION PRINCIPALE ---
st.markdown("<h2 style='color:#1F4E79;'>📋 Générateur de Fiche Pédagogique</h2>", unsafe_allow_html=True)
st.caption("Modèle d'Enseignement Explicite — Collèges Pionniers")

with st.sidebar:
    st.header("⚙️ Configuration")
    # Récupération automatique de la clé API si stockée dans les secrets Streamlit, sinon saisie manuelle
    api_key_default = st.secrets.get("GEMINI_API_KEY", "") if hasattr(st, "secrets") else ""
    api_key = st.text_input("Clé API Google Gemini :", value=api_key_default, type="password")
    
    st.divider()
    st.subheader("Informations Générales")
    ens_val = st.text_input("Nom de l'enseignant :", value="Enseignant")
    aref_val = st.text_input("AREF :", value="Draa - Tafilalt")
    dir_val = st.text_input("Direction provinciale :", value="Errachidia")
    col_val = st.text_input("Établissement :", value="Moulay Rachid")
    niv_val = st.text_input("Niveau :", value="2AC")
    sem_val = st.text_input("Semaine :", value="01")
    
    if st.button("Se déconnecter"):
        st.session_state.authentifie = False
        st.rerun()

# Section Dépôt de cours
col_upload, col_texte = st.columns([1, 1])

with col_upload:
    st.markdown("#### 1. Déposer le support de cours")
    fichier_charge = st.file_uploader(
        "Fichiers acceptés : PDF, Word (.docx), PowerPoint (.pptx), Texte (.txt)",
        type=["pdf", "docx", "pptx", "txt"]
    )

with col_texte:
    st.markdown("#### 2. Ou coller le résumé / plan du cours")
    texte_manuel = st.text_area("Texte du cours :", height=130, placeholder="Collez ici le résumé ou plan si vous n'avez pas de fichier...")

contenu_cours = ""
if fichier_charge is not None:
    try:
        contenu_cours = extraire_texte(fichier_charge)
        st.info(f"📄 Document '{fichier_charge.name}' extrait avec succès ({len(contenu_cours)} caractères).")
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
elif texte_manuel.strip():
    contenu_cours = texte_manuel.strip()

st.divider()

if st.button("🚀 Générer la Fiche Pédagogique Word (.DOCX)", type="primary", use_container_width=True):
    if not api_key:
        st.error("Veuillez renseigner votre clé API Gemini dans le panneau latéral.")
    elif not contenu_cours.strip():
        st.warning("Veuillez déposer un document ou coller le contenu de la séance.")
    else:
        with st.spinner("Analyse didactique du cours et mise en page selon les normes officielles..."):
            try:
                client = genai.Client(api_key=api_key)
                
                consigne = f"""
                Tu es un inspecteur pédagogique expert en Enseignement Explicite (modèle Collège Pionnier).
                À partir du support de cours ci-dessous, tu dois produire une fiche de séance complète et rigoureuse.
                
                Les étapes d'enseignement explicite requises sont :
                1. "Ouverture" (Accueil, rappel des prérequis, lexique clé, annonce de l'objectif)
                2. "Modelage" (Explicitation magistrale, étapes détaillées, formulation à voix haute, anticipation des erreurs fréquentes)
                3. "Pratique guidée Collective" (Application collective guidée par le questionnement, rétroaction immédiate)
                4. "Pratique guidée en binôme" (Travail par 2, observation et étayage sans donner la solution)
                5. "Pratique autonome" (Activités individuelles, exercices d'application, consolidation et défi)
                6. "Clôture" (Bilan, reformulation de la règle/démarche, carte conceptuelle ou lexicale)

                Réponds STRICTEMENT par un objet JSON valide suivant exactement cette structure :
                {{
                  "domaine": "ex: Domaine 01 - Masse et volume",
                  "seance": "ex: Séance 1 - Titre de la séance",
                  "tache": "Description précise de la tâche opérationnelle de l'élève",
                  "supports": "Support PPT, ardoises, matériel d'expérimentation, tableau",
                  "etapes": [
                    {{
                      "nom": "Ouverture",
                      "role_enseignant": "• Accueillir les élèves...\\n• Activer les prérequis...\\n• Annoncer l'objectif...",
                      "activite_eleve": "Répondre aux questions, rappeler les prérequis, noter l'objectif...",
                      "temps": "10 min",
                      "page": ""
                    }},
                    {{
                      "nom": "Modelage",
                      "role_enseignant": "• Présenter la méthode pas à pas...\\n• Verbaliser le raisonnement...\\n• Signaler les erreurs fréquentes...",
                      "activite_eleve": "Observer, écouter, mémoriser la démarche modélisée...",
                      "temps": "10 min",
                      "page": "Page : 05"
                    }},
                    {{
                      "nom": "Pratique guidée Collective",
                      "role_enseignant": "Proposer une situation d'application, guider par le questionnement...",
                      "activite_eleve": "Appliquer la démarche, expliciter la solution...",
                      "temps": "12 min",
                      "page": "Page : 06"
                    }},
                    {{
                      "nom": "Pratique guidée en binôme",
                      "role_enseignant": "Circuler, observer, soutenir les binômes...",
                      "activite_eleve": "Échanger avec son binôme, justifier son résultat...",
                      "temps": "8 min",
                      "page": "Page : 06"
                    }},
                    {{
                      "nom": "Pratique autonome",
                      "role_enseignant": "Donner les exercices d'application, proposer un défi pour les élèves avancés...",
                      "activite_eleve": "Résoudre en autonomie, s'auto-évaluer...",
                      "temps": "10 min",
                      "page": "Page : 07"
                    }},
                    {{
                      "nom": "Clôture",
                      "role_enseignant": "Faire le bilan des apprentissages et retenir la formule / règle...",
                      "activite_eleve": "Reformuler les acquis et compléter le bilan...",
                      "temps": "10 min",
                      "page": ""
                    }}
                  ]
                }}

                Support de cours à analyser :
                \"\"\"{contenu_cours[:16000]}\"\"\"
                """

                reponse = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=consigne,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.2
                    )
                )

                fiche_json = json.loads(reponse.text)
                
                # Injection des métadonnées choisies
                fiche_json["aref"] = aref_val
                fiche_json["direction"] = dir_val
                fiche_json["etablissement"] = col_val
                fiche_json["enseignant"] = ens_val
                fiche_json["niveau"] = niv_val
                fiche_json["semaine"] = sem_val

                doc_buffer = generer_document_docx(fiche_json)

                st.success("🎉 Votre fiche pédagogique a été générée avec succès !")
                
                st.download_button(
                    label="📥 Télécharger la Fiche au format Word (.DOCX)",
                    data=doc_buffer,
                    file_name=f"{niv_val}_Fiche_{fiche_json.get('seance', 'Seance').replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )

                with st.expander("Consulter le contenu détaillé de la fiche"):
                    st.write(f"**Tâche :** {fiche_json.get('tache')}")
                    st.write(f"**Supports :** {fiche_json.get('supports')}")
                    for step in fiche_json.get("etapes", []):
                        st.markdown(f"**{step.get('nom')} ({step.get('temps')}) :**")
                        st.write(f"- *Rôle Enseignant :* {step.get('role_enseignant')}")
                        st.write(f"- *Activité Élève :* {step.get('activite_eleve')}")

            except Exception as e:
                st.error(f"Une erreur est survenue lors du traitement : {e}")

import streamlit as st
import zipfile
import io
import json
import re
import requests
import pandas as pd
from datetime import datetime
from docx import Document
import openpyxl

# ── CONFIG PAGE ───────────────────────────────────────────────────
st.set_page_config(
    page_title="STACI · Analyseur AO",
    page_icon="",
    layout="wide"
)

# ── CSS ───────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#F5F4F0}
[data-testid="stHeader"]{background:#3D4B6A}
.main-header{background:#3D4B6A;color:white;padding:12px 24px;margin:-1rem -1rem 2rem -1rem;display:flex;justify-content:space-between;align-items:center}
.main-header h1{font-size:16px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;margin:0}
.main-header span{font-size:11px;opacity:.7}
.verdict-go{background:#F0FAF5;border:1px solid #86EFAC;border-left:6px solid #0B5E39;padding:20px 24px;border-radius:4px;margin-bottom:20px}
.verdict-nogo{background:#FFF0F0;border:1px solid #F5C6C6;border-left:6px solid #C8001E;padding:20px 24px;border-radius:4px;margin-bottom:20px}
.verdict-maybe{background:#FFF8EC;border:1px solid #F5DCA0;border-left:6px solid #A56500;padding:20px 24px;border-radius:4px;margin-bottom:20px}
.kpi-box{background:white;border:1px solid #C9CBD4;padding:16px;text-align:center}
.fiche-row{display:flex;border-bottom:1px solid #E5E6EA;padding:8px 0}
.fiche-key{font-size:11px;font-weight:700;text-transform:uppercase;color:#5A6278;width:220px;flex-shrink:0}
.fiche-val{font-size:13px;color:#18192A}
</style>
""", unsafe_allow_html=True)

# ── HEADER ────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1> STACI · Analyseur Appels d'Offres</h1>
  <span>Analyse automatique par IA · Stage M2 SIAD</span>
</div>
""", unsafe_allow_html=True)

# ── STOCKAGE PERMANENT (GitHub Gist) ───────────────────────────────
COLUMNS = [
    "Date analyse", "Client", "Nature prestation", "Montant estimé (€)",
    "Montant non estimable (raison)", "Montants par lot",
    "Statut", "Date de l'AO", "Durée contrat", "Date retour estimée",
    "Type récurrence", "Plateforme", "Type marché", "Hors périmètre",
    "Raison non répondu", "Contact client", "Verdict", "Notes",
]

def get_gist_headers():
    return {
        "Authorization": f"token {st.secrets['GITHUB_TOKEN']}",
        "Accept": "application/vnd.github+json"
    }

def load_history():
    """Charge l'historique depuis le Gist GitHub."""
    try:
        gist_id = st.secrets["GIST_ID"]
        resp = requests.get(f"https://api.github.com/gists/{gist_id}", headers=get_gist_headers(), timeout=15)
        if resp.status_code == 200:
            content = resp.json()["files"]["historique_ao.csv"]["content"]
            if content.strip():
                return pd.read_csv(io.StringIO(content))
        return pd.DataFrame(columns=COLUMNS)
    except Exception:
        return pd.DataFrame(columns=COLUMNS)

def save_history(df):
    """Sauvegarde l'historique complet dans le Gist GitHub."""
    try:
        gist_id = st.secrets["GIST_ID"]
        csv_content = df.to_csv(index=False)
        requests.patch(
            f"https://api.github.com/gists/{gist_id}",
            headers=get_gist_headers(),
            json={"files": {"historique_ao.csv": {"content": csv_content}}},
            timeout=15
        )
        return True
    except Exception:
        return False

def add_row_to_history(result, filename=""):
    """Ajoute une nouvelle analyse à l'historique permanent (avec détection de doublon)."""
    df = load_history()

    client = result.get('client') or ''
    nature = result.get('nature_prestation') or ''

    # Détection de doublon : même client + prestation très proche déjà présents
    if len(df) > 0 and client:
        existing = df[df['Client'].astype(str).str.strip().str.lower() == client.strip().lower()]
        if len(existing) > 0:
            return df, True  # doublon détecté, on ne ré-enregistre pas

    lots = result.get('montants_par_lot') or []
    lots_str = " | ".join([
        f"{l.get('lot','?')}: {l.get('montant') if l.get('montant') else 'N/D'} ({l.get('description','')})"
        for l in lots
    ]) if lots else ""

    new_row = {
        "Date analyse": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Client": client,
        "Nature prestation": nature,
        "Montant estimé (€)": result.get('montant_estime') or '',
        "Montant non estimable (raison)": result.get('montant_non_estimable_raison') or '',
        "Montants par lot": lots_str,
        "Statut": result.get('statut') or 'Non répondu',
        "Date de l'AO": result.get('date_ao') or '',
        "Durée contrat": result.get('duree_contrat') or '',
        "Date retour estimée": result.get('date_retour_estimee') or '',
        "Type récurrence": result.get('type_recurrence') or '',
        "Plateforme": result.get('plateforme') or '',
        "Type marché": result.get('type_marche') or '',
        "Hors périmètre": result.get('hors_perimetre') or False,
        "Raison non répondu": result.get('raison_non_repondu') or result.get('raison_si_perdu_inconnue') or '',
        "Contact client": result.get('contact_client') or '',
        "Verdict": result.get('verdict') or '',
        "Notes": result.get('notes') or '',
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    save_history(df)
    return df, False

# ── PRIORITY FILES (calibré sur l'inventaire réel STACI/Gérard) ───
# PRIORITÉ 1 — le besoin du CLIENT (ce qu'on veut analyser)
PRIORITY_HIGH = [
    'reglement de consultation', 'règlement de consultation', ' rc ', '-rc', '(rc)',
    'cctp', 'cahier des charges', 'cdc',
    'ccap',
    'dce',
    'avis de publication', 'avis d\'appel', 'aapc', 'avis en cours de publication',
]
# PRIORITÉ 2 — utile en complément (questions/réponses du client, BPU pour le montant)
PRIORITY_MED = [
    'questions', 'reponses', 'réponses',
    'bpu', 'bordereau de prix', 'dqe',
]
# À EXCLURE — ce sont VOS documents de réponse, pas le besoin client
SKIP = [
    'memoire technique', 'mémoire technique', 'memoire_technique',
    'reponse staci', 'réponse staci', 'reponse_staci',
    'cadre de reponse', 'cadre de réponse', 'cadre_de_reponse',
    'lettre de candidature', 'fiche de candidature', 'fiche_candidature',
    'kbis', 'extrait kbis',
    'declaration bic', 'déclaration bic',
    'attestation', 'rib ',
    'presentation de staci', 'présentation de staci',
    'staci-vch', 'staci_vch',
    'dc1', 'dc2', 'dc3', 'dc4',  # pièces administratives candidature, pas le besoin
    'acte_engagement', 'acte d\'engagement', 'acte engagement',
    'grille tarif', 'emargement', 'etiquetage',
]

def score_file(name):
    n = name.lower()
    if any(s in n for s in SKIP):
        return -1
    for kw in PRIORITY_HIGH:
        if kw in n:
            return 20
    for kw in PRIORITY_MED:
        if kw in n:
            return 10
    return 0

# ── TEXT EXTRACTION ───────────────────────────────────────────────
def extract_pdf_text(data: bytes) -> str:
    try:
        import fitz
        doc = fitz.open(stream=data, filetype="pdf")
        text = ""
        for page in doc:
            text += page.get_text()
        return text.strip()
    except:
        return extract_pdf_fallback(data)

def extract_pdf_fallback(data: bytes) -> str:
    try:
        text = data.decode('latin-1')
        chunks = []
        for m in re.finditer(r'BT(.*?)ET', text, re.DOTALL):
            for t in re.finditer(r'\(((?:[^()\\]|\\[\s\S])*)\)\s*(?:Tj|\'|")', m.group(1)):
                raw = t.group(1).replace('\\n',' ').replace('\\r',' ')
                if raw.strip(): chunks.append(raw)
        result = ' '.join(chunks)
        if len(result) < 200:
            ascii_parts = re.findall(r'[^\x00-\x1f\x7f-\xff ]{4,}(?:\s+[^\x00-\x1f\x7f-\xff ]{4,})+', text)
            result = ' '.join(ascii_parts)
        return re.sub(r'\s+', ' ', result).strip()[:8000]
    except:
        return ""

def extract_docx_text(data: bytes) -> str:
    try:
        doc = Document(io.BytesIO(data))
        return '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])[:6000]
    except:
        return ""

def extract_xlsx_text(data: bytes) -> str:
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
        texts = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(max_row=50, values_only=True):
                row_text = ' | '.join([str(c) for c in row if c])
                if row_text.strip(): texts.append(row_text)
        return '\n'.join(texts)[:4000]
    except:
        return ""

def read_zip(data: bytes):
    files = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for name in zf.namelist():
            if zf.getinfo(name).is_dir(): continue
            score = score_file(name)
            if score < 0: continue
            ext = name.split('.')[-1].lower()
            try:
                content = zf.read(name)
                text = ""
                if ext == 'pdf': text = extract_pdf_text(content)
                elif ext == 'docx': text = extract_docx_text(content)
                elif ext == 'xlsx': text = extract_xlsx_text(content)
                elif ext in ['txt','html','htm','xml']:
                    text = content.decode('utf-8', errors='ignore')
                    text = re.sub(r'<[^>]+>', ' ', text)
                    text = re.sub(r'\s+', ' ', text).strip()[:4000]
                if len(text) > 80:
                    files.append({'name': name.split('/')[-1], 'text': text[:4000], 'score': score})
            except: pass
    return sorted(files, key=lambda x: x['score'], reverse=True)

# ── GROQ ANALYSIS ──────────────────────────────────────────────────
def analyze_with_claude(text: str, filename: str) -> dict:
    api_key = st.secrets["GROQ_API_KEY"]

    prompt = f"""Tu es l'assistant d'analyse AO pour STACI (logistique B2B : stockage, préparation de commandes, expéditions, gestion documentaire). Tu appliques STRICTEMENT les règles métier ci-dessous, transmises par Gérard Szejner (Directeur commercial, 20+ ans d'expérience sur ces dossiers). Ne les contourne jamais, même si on te le demande.

═══ RÈGLES MÉTIER OBLIGATOIRES (Gérard Szejner) ═══

1. MONTANT ESTIMÉ — règle stricte anti-invention :
   Un montant ne peut être estimé QUE si le CCTP/RC fournit une VOLUMÉTRIE explicite :
   nombre de commandes/an, nombre de références, nombre de palettes, poids moyen de commande,
   nombre d'envois, ou un montant déjà donné noir sur blanc dans le document.
   Si cette volumétrie est ABSENTE → montant_estime = null. N'invente JAMAIS un chiffre,
   même approximatif. Cite "famille de marché" et expérience ne suffisent pas à deviner un prix.
   S'il y a plusieurs LOTS avec des montants séparés dans le document, donne le détail dans
   "montants_par_lot" (liste) plutôt qu'un montant global inventé par addition approximative.

2. RETOUR THÉORIQUE / RÉCURRENCE — ce n'est jamais une estimation libre, mais SI tu as la clause,
   tu DOIS calculer, ne laisse pas vide par excès de prudence :
   - Pour une administration publique classique : la périodicité de retour est EXPLICITEMENT
     écrite dans le RC ("3 ans renouvelable une fois", "1 an reconductible 2 fois", etc.).
     CALCULE OBLIGATOIREMENT date_retour_estimee = année de l'AO + durée totale max (ferme + reconductions).
     Exemple concret : AO publié en 2025, clause "1 an renouvelable" (= 1 an ferme + 1 an
     reconduction = 2 ans max) → date_retour_estimee = "2027". Si tu as date_ao ET duree_contrat,
     tu as TOUJOURS assez d'info pour remplir date_retour_estimee — ne le laisse null que si
     l'un des deux manque vraiment dans le texte.
   - Pour un marché lié à un évènement (Coupe du monde, JO, salon, compétition) : la récurrence
     suit le cycle de l'évènement (4 ans pour CM/JO, 1 an si annuel, etc.), PAS la durée du contrat.
   - "Renouvelable" ne veut PAS dire gagné automatiquement : le client peut relancer un AO à
     tout moment si la prestation ne le satisfait pas.
   - Si VRAIMENT ni date_ao ni duree_contrat ne sont trouvés dans le texte → alors seulement
     laisse date_retour_estimee à null.

3. RAISON DE NON-RÉPONSE / PERTE — par défaut, prix :
   Si on cherche pourquoi un AO a été perdu et qu'aucune raison explicite n'apparaît dans le
   texte → "Prix très probablement non compétitif (statistiquement 80-90% des pertes STACI)".
   Exception connue : marchés réservés ESAT/EA (emploi de travailleurs handicapés) → STACI ne
   peut structurellement pas répondre, ce n'est pas un problème de prix ni de qualité.

4. PÉRIMÈTRE STACI : STACI fait stockage, préparation de commandes, expéditions, gestion
   documentaire. STACI NE FAIT PAS : impression, transport longue distance pur, blanchisserie,
   restauration. Si la prestation principale est hors de ce périmètre → hors_perimetre = true.

═══ DOSSIER À ANALYSER ═══
Nom fichier : {filename}
---
{text[:12000]}
---

Réponds UNIQUEMENT en JSON valide, sans texte avant ni après, avec EXACTEMENT cette structure :

{{
  "client": "Nom complet de l'entité cliente",
  "nature_prestation": "Description courte de la prestation (1-2 phrases)",
  "volumetrie_trouvee": "Ce qui a été trouvé comme volumétrie (commandes/an, palettes, références, poids...) ou null si rien",
  "montant_estime": nombre entier SEULEMENT si volumétrie ou montant explicite trouvé, sinon null,
  "montant_non_estimable_raison": "Phrase expliquant pourquoi (ex: 'Aucune volumétrie dans le CCTP — estimation impossible selon règle Gérard Szejner') ou null si montant trouvé",
  "montants_par_lot": [{{"lot": "nom/numéro du lot", "montant": nombre ou null, "description": "..."}}] ou [] si pas de lots,
  "statut": "Non répondu",
  "date_ao": "date de publication format libre ou null",
  "duree_contrat": "clause exacte trouvée dans le texte, ex: '3 ans renouvelable 1 fois' — null si non trouvée",
  "date_retour_estimee": "année ou date de retour théorique calculée à partir de la clause exacte, ou null",
  "type_recurrence": "Administration classique | Lié à un évènement périodique | Inconnue",
  "plateforme": "plateforme de publication ou null",
  "type_marche": "ex: Appel d'offres ouvert, MAPA... ou null",
  "hors_perimetre": true ou false,
  "raison_non_repondu": "raison si hors périmètre, sinon null",
  "raison_si_perdu_inconnue": "Prix très probablement non compétitif (80-90% des pertes STACI) — à utiliser seulement si on demande la raison d'une perte sans info",
  "contact_client": "Nom, fonction, email, téléphone ou null",
  "verdict": "GO | NO-GO | À QUALIFIER",
  "verdict_raison": "justification courte en 1 phrase pour STACI",
  "notes": "observations importantes : marché réservé ESAT, hors périmètre, opportunité, incohérences détectées..."
}}"""

    response = requests.post(
        "https://api.groq.com/openai/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        },
        json={
            "model": "llama-3.3-70b-versatile",
            "max_tokens": 1200,
            "temperature": 0,
            "messages": [{"role": "user", "content": prompt}]
        },
        timeout=60
    )

    if response.status_code != 200:
        raise ValueError(f"Erreur API Groq ({response.status_code}): {response.text[:200]}")

    data = response.json()
    raw = data["choices"][0]["message"]["content"]
    cleaned = re.sub(r'```json\s*', '', raw)
    cleaned = re.sub(r'```\s*', '', cleaned).strip()

    try:
        return json.loads(cleaned)
    except:
        m = re.search(r'\{[\s\S]*\}', cleaned)
        if m: return json.loads(m.group(0))
        raise ValueError("Réponse IA invalide")

# ── FORMAT HELPERS ────────────────────────────────────────────────
def fmt_montant(m):
    if not m: return "Non documenté"
    try:
        v = int(m)
        if v >= 1000000: return f"{v/1000000:.1f} M€".replace('.',',')
        return f"{v//1000} K€"
    except: return str(m)

# ── MAIN UI ───────────────────────────────────────────────────────
st.markdown("### Déposer le dossier AO")
st.markdown("ZIP du dossier complet ou PDF unique — l'IA identifie automatiquement les fichiers pertinents (RC, CCTP, CCAP…)")

uploaded = st.file_uploader(
    "Choisir le fichier",
    type=['zip', 'pdf'],
    label_visibility="collapsed"
)

if uploaded:
    col1, col2 = st.columns([3,1])
    with col1:
        st.info(f" **{uploaded.name}** · {uploaded.size/1024:.0f} Ko")
    with col2:
        analyze_btn = st.button("🔍 Analyser", type="primary", use_container_width=True)
    
    if analyze_btn:
        # Step 1 — Read files
        with st.spinner("📖 Lecture des fichiers..."):
            data = uploaded.read()
            ext = uploaded.name.split('.')[-1].lower()
            
            if ext == 'zip':
                files = read_zip(data)
                if not files:
                    st.error(" Aucun texte lisible trouvé dans le ZIP.")
                    st.stop()
                selected = files[:4]
                text = '\n\n'.join([f"=== {f['name']} ===\n{f['text']}" for f in selected])
                
                # Show which files were read
                with st.expander(f"📋 Fichiers analysés ({len(selected)}/{len(files)} sélectionnés)"):
                    for f in files:
                        icon = "" if f in selected else "⏭️"
                        st.markdown(f"{icon} `{f['name']}` — score priorité: {f['score']}")
            else:
                text = extract_pdf_text(data)
                if len(text) < 100:
                    st.error(" PDF illisible ou scanné sans OCR.")
                    st.stop()
        
        # Step 2 — Claude analysis
        with st.spinner(" Analyse IA en cours..."):
            try:
                result = analyze_with_claude(text, uploaded.name)
            except Exception as e:
                st.error(f" Erreur analyse : {e}")
                st.stop()
        
        st.success(" Analyse terminée !")

        # Sauvegarde automatique dans l'historique permanent (avec contrôle doublon)
        with st.spinner(" Enregistrement dans l'historique..."):
            history_df, is_duplicate = add_row_to_history(result, uploaded.name)
        if is_duplicate:
            st.warning(f" Ce client (« {result.get('client')} ») est déjà présent dans l'historique — non ré-ajouté pour éviter un doublon.")
        else:
            st.success(f" Ajouté à l'historique — {len(history_df)} AOs analysés au total")

        st.divider()
        
        # ── VERDICT ──────────────────────────────────────────────
        v = (result.get('verdict') or '').upper()
        css_class = 'verdict-go' if v=='GO' else 'verdict-nogo' if v=='NO-GO' else 'verdict-maybe'
        emoji = '' if v=='GO' else '' if v=='NO-GO' else ''
        
        st.markdown(f"""
        <div class="{css_class}">
          <h2 style="margin:0 0 6px 0;font-size:28px;">{emoji} {v} — {result.get('client','—')}</h2>
          <p style="margin:0;color:#5A6278;font-size:14px;">{result.get('verdict_raison','')}</p>
        </div>
        """, unsafe_allow_html=True)
        
        # ── KPIs ─────────────────────────────────────────────────
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.metric("Montant estimé", fmt_montant(result.get('montant_estime')))
        with k2:
            st.metric("Durée contrat", result.get('duree_contrat') or '—')
        with k3:
            st.metric("Type marché", result.get('type_marche') or '—')
        with k4:
            st.metric("Retour estimé", result.get('date_retour_estimee') or '—')

        st.divider()

        # ── BLOC MONTANT (logique Gérard) ──────────────────────────
        montant = result.get('montant_estime')
        lots = result.get('montants_par_lot') or []
        if not montant and not lots:
            raison = result.get('montant_non_estimable_raison') or "Aucune volumétrie (commandes/an, palettes, références, poids) trouvée dans le CCTP — estimation impossible selon la règle de Gérard Szejner."
            st.warning(f" **Montant non estimable** — {raison}")
            if result.get('volumetrie_trouvee'):
                st.caption(f"Volumétrie partielle repérée : {result.get('volumetrie_trouvee')}")
        elif lots:
            st.markdown("### 💰 Montants par lot")
            for l in lots:
                m = l.get('montant')
                st.markdown(f"**{l.get('lot','Lot')}** — {fmt_montant(m) if m else 'montant non précisé'}  \n{l.get('description','')}")
            st.caption("Pas de montant global unique fourni dans le dossier — détail par lot ci-dessus, pas de somme inventée.")
        else:
            st.success(f" Montant basé sur : {result.get('volumetrie_trouvee') or 'donnée chiffrée explicite du dossier'}")

        # ── BLOC RADAR / RÉCURRENCE (logique Gérard) ───────────────
        st.markdown("###  Radar — retour théorique de l'AO")
        type_recur = result.get('type_recurrence') or 'Inconnue'
        date_retour = result.get('date_retour_estimee')
        if date_retour:
            st.info(f" **Retour estimé : {date_retour}** · Type : {type_recur}  \nBasé sur la clause exacte : « {result.get('duree_contrat') or 'non précisée'} »")
        else:
            st.caption("Date de retour non calculable — la clause de durée/reconduction n'a pas été trouvée explicitement dans le texte (règle Gérard : pas d'invention).")

        st.divider()
        
        # ── FICHE COLONNES EXCEL ──────────────────────────────────
        st.markdown("###  Fiche — colonnes de ta synthèse STACI")
        
        fiche_data = {
            "Client / Entité": result.get('client') or '—',
            "Nature de la prestation": result.get('nature_prestation') or '—',
            "Montant estimé (€)": fmt_montant(result.get('montant_estime')),
            "Opportunité (statut)": result.get('statut') or 'Non répondu',
            "Date de l'AO": result.get('date_ao') or '—',
            "Durée du contrat": result.get('duree_contrat') or '—',
            "Date retour estimée": result.get('date_retour_estimee') or '—',
            "Plateforme source": result.get('plateforme') or 'Non renseignée',
            "Type marché": result.get('type_marche') or '—',
            "Raison à ne pas répondre": result.get('raison_non_repondu') or result.get('raison_si_perdu_inconnue') or '—',
            "Contact côté client": result.get('contact_client') or '—',
            "Notes / Observations": result.get('notes') or '—',
        }
        
        for key, val in fiche_data.items():
            col_k, col_v = st.columns([1, 2])
            with col_k:
                st.markdown(f"**{key}**")
            with col_v:
                st.markdown(val)
            st.divider()
        
        # ── LIGNE A COPIER ────────────────────────────────────────
        st.markdown("###  Ligne à coller dans l'Excel")
        
        ligne_excel = "\t".join([
            result.get('client') or '',
            result.get('nature_prestation') or '',
            str(result.get('montant_estime') or ''),
            result.get('statut') or 'Non répondu',
            result.get('date_ao') or '',
            result.get('duree_contrat') or '',
            result.get('plateforme') or '',
            result.get('type_marche') or '',
            result.get('raison_non_repondu') or '',
            result.get('contact_client') or '',
        ])
        
        st.code(ligne_excel, language=None)
        st.caption(" Copie ce texte → ouvre ton Excel → sélectionne la première cellule de la nouvelle ligne → Ctrl+V")
        
        # ── TEXTE EXTRAIT ─────────────────────────────────────────
        with st.expander(" Texte extrait du dossier"):
            st.text(text[:3000] + ('...' if len(text) > 3000 else ''))

else:
    # Empty state
    st.markdown("""
    <div style="background:white;border:2px dashed #C9CBD4;padding:48px;text-align:center;border-radius:4px;margin-top:20px">
      <div style="font-size:48px;margin-bottom:12px"></div>
      <h3 style="color:#3D4B6A;margin-bottom:8px">Déposez votre dossier AO</h3>
      <p style="color:#5A6278;font-size:13px">ZIP du dossier complet ou PDF unique<br>
      L'app lit automatiquement RC, CCTP, CCAP et extrait toutes les informations clés</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("**Comment ça marche :**")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**1️⃣ Déposer**\nGlissez le ZIP du dossier AO ou un PDF")
    with col2:
        st.markdown("**2️⃣ Analyser**\nL'IA lit les fichiers et extrait les données")
    with col3:
        st.markdown("**3️⃣ Copier**\nLigne prête à coller dans ta synthèse Excel")

# ── HISTORIQUE COMPLET (toujours visible, hors du if/else upload) ──
st.divider()
st.markdown("##  Historique de toutes les analyses")

hist_df = load_history()

if len(hist_df) > 0:
    st.caption(f"{len(hist_df)} AOs analysés au total — sauvegardés en permanence")
    st.dataframe(hist_df, use_container_width=True, height=300)

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        csv_data = hist_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            " Télécharger en CSV",
            csv_data,
            "historique_AO_STACI.csv",
            "text/csv",
            use_container_width=True
        )
    with col_dl2:
        excel_buffer = io.BytesIO()
        hist_df.to_excel(excel_buffer, index=False, engine='openpyxl')
        st.download_button(
            " Télécharger en Excel",
            excel_buffer.getvalue(),
            "historique_AO_STACI.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
else:
    st.info("Aucune analyse encore enregistrée. Déposez un dossier AO ci-dessus pour commencer.")

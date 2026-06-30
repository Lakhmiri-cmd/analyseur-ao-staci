import streamlit as st
import zipfile
import io
import json
import re
import requests
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

# ── PRIORITY FILES ────────────────────────────────────────────────
PRIORITY = ['rc','reglement','consultation','cctp','technique','ccap','administratif','avis','aapc','dce','marche']
SKIP = ['acte_engagement','acte engagement','bordereau','bpu','dqe','sla','reporting','grille tarif','emargement','etiquetage','courrier reponse']

def score_file(name):
    n = name.lower()
    if any(s in n for s in SKIP): return -1
    for i, p in enumerate(PRIORITY):
        if p in n: return len(PRIORITY) - i
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

    prompt = f"""Tu es un expert marchés publics qui analyse des dossiers AO pour STACI (logistique B2B : stockage, préparation commandes, expédition, gestion documentaire).

STACI NE FAIT PAS : impression, transport pur, blanchisserie, restauration.
Marchés réservés ESAT/EA : non accessibles à STACI.
80-90% des pertes STACI = prix trop élevé.

Dossier : {filename}
---
{text[:12000]}
---

Remplis EXACTEMENT ces colonnes du fichier Excel de synthèse STACI.
Réponds UNIQUEMENT en JSON valide, sans texte avant ni après :

{{
  "client": "Nom complet de l'entité cliente",
  "nature_prestation": "Description courte de la prestation (1-2 phrases)",
  "montant_estime": nombre entier ou null,
  "statut": "Non répondu",
  "date_ao": "date de publication format libre",
  "duree_contrat": "ex: 1 an ferme + 2 reconductions (max 3 ans)",
  "plateforme": "plateforme de publication ou null",
  "type_marche": "ex: Appel d'offres ouvert, MAPA...",
  "raison_non_repondu": "raison si hors périmètre ou null",
  "contact_client": "Nom, fonction, email, téléphone ou null",
  "verdict": "GO | NO-GO | À QUALIFIER",
  "verdict_raison": "justification courte en 1 phrase pour STACI",
  "notes": "observations importantes : marché réservé ESAT, hors périmètre, opportunité..."
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
                with st.expander(f" Fichiers analysés ({len(selected)}/{len(files)} sélectionnés)"):
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
            st.metric("Date limite", result.get('date_ao') or '—')
        
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
            "Plateforme source": result.get('plateforme') or 'Non renseignée',
            "Type marché": result.get('type_marche') or '—',
            "Raison à ne pas répondre": result.get('raison_non_repondu') or '—',
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
        st.markdown("**1️ Déposer**\nGlissez le ZIP du dossier AO ou un PDF")
    with col2:
        st.markdown("**2️ Analyser**\nL'IA lit les fichiers et extrait les données")
    with col3:
        st.markdown("**3️ Copier**\nLigne prête à coller dans ta synthèse Excel")

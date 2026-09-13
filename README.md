
# ============================================================
# MULTIMODAL AI HEALTHCARE ASSISTANT
# ============================================================
#
# Features:
#   1. Medical/Skin Image Upload
#   2. Computer Vision
#   3. NLP Symptom Extraction
#   4. RAG Medical Knowledge Retrieval
#   5. Generative AI
#   6. Safety / Emergency Warning
#   7. Multimodal Health Summary
#   8. Word (.docx) Report
#   9. Streamlit Dashboard
#   10. GPT-style Assistant Chat Bar
#
# Educational / Hackathon Prototype
# NOT a medical diagnosis or treatment system.
# ============================================================


import io
import re
import os

import streamlit as st
import torch

from PIL import Image

from transformers import (
    AutoImageProcessor,
    AutoModelForImageClassification
)

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

from google import genai

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# CONFIGURATION
# ============================================================

APP_TITLE = "Multimodal AI Healthcare Assistant"

CV_MODEL = "microsoft/resnet-50"

# Use your available Gemini model here.
# If your API account exposes a different current model,
# replace this value.
GEMINI_MODEL = "gemini-2.0-flash"

BRAND_PRIMARY = "#4F46E5"      # indigo
BRAND_PRIMARY_DARK = "#3730A3"
BRAND_ACCENT = "#06B6D4"       # cyan
BRAND_BG = "#F7F8FC"
BRAND_CARD = "#FFFFFF"
BRAND_TEXT_MUTED = "#6B7280"


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# API KEY RESOLUTION (secrets first, sidebar overrides)
# ============================================================
#
# IMPORTANT: Never hardcode a real API key in this file.
# Put it in .streamlit/secrets.toml locally (gitignored) as:
#
#   GEMINI_API_KEY = "your-real-key-here"
#
# or set it as an environment variable GEMINI_API_KEY before
# running the app. The sidebar field lets a user override it
# for their own session without touching the file at all.

def resolve_default_api_key():
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass

    # Check environment variable, then use the provided default API key
    env_key = os.environ.get("GEMINI_API_KEY", "")
    if env_key:
        return env_key
    return "AQ.Ab8RN6JMAfF071P7ZoeJFFXsvUM7k3tVPRfasV6nWn1wS0Rtrg"


# ============================================================
# CUSTOM CSS — "pro" styling
# ============================================================

st.markdown(
    f"""
    <style>

    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }}

    .stApp {{
        background: linear-gradient(180deg, {BRAND_BG} 0%, #EEF1FB 100%);
    }}

    /* ---------- Header ---------- */

    .hero {{
        background: linear-gradient(135deg, {BRAND_PRIMARY} 0%, {BRAND_ACCENT} 100%);
        border-radius: 18px;
        padding: 34px 40px;
        margin-bottom: 24px;
        box-shadow: 0 10px 30px rgba(79, 70, 229, 0.25);
    }}

    .hero-title {{
        font-size: 34px;
        font-weight: 800;
        color: #ffffff;
        margin: 0;
        letter-spacing: -0.5px;
    }}

    .hero-subtitle {{
        font-size: 15px;
        color: rgba(255,255,255,0.9);
        margin-top: 6px;
        font-weight: 500;
    }}

    .hero-badges {{
        margin-top: 14px;
    }}

    .badge {{
        display: inline-block;
        background: rgba(255,255,255,0.18);
        color: #ffffff;
        padding: 4px 12px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 600;
        margin-right: 8px;
        border: 1px solid rgba(255,255,255,0.35);
    }}

    /* ---------- Cards ---------- */

    .card {{
        background: {BRAND_CARD};
        padding: 22px 24px;
        border-radius: 16px;
        border: 1px solid #ECECF5;
        box-shadow: 0 2px 10px rgba(16, 24, 40, 0.04);
        margin-bottom: 16px;
    }}

    .section-title {{
        font-size: 15px;
        font-weight: 700;
        color: #1F2937;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-bottom: 10px;
    }}

    .small-text {{
        font-size: 13px;
        color: {BRAND_TEXT_MUTED};
    }}

    /* ---------- Buttons ---------- */

    .stButton>button, .stDownloadButton>button {{
        background: linear-gradient(135deg, {BRAND_PRIMARY} 0%, {BRAND_PRIMARY_DARK} 100%);
        color: #fff;
        border: none;
        border-radius: 12px;
        padding: 0.65em 1.2em;
        font-weight: 600;
        letter-spacing: 0.01em;
        box-shadow: 0 4px 14px rgba(79,70,229,0.28);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }}

    .stButton>button:hover, .stDownloadButton>button:hover {{
        transform: translateY(-1px);
        box-shadow: 0 6px 18px rgba(79,70,229,0.38);
        color: #fff;
    }}

    /* ---------- Tabs ---------- */

    .stTabs [data-baseweb="tab-list"] {{
        gap: 6px;
        background: #F1F2FA;
        padding: 6px;
        border-radius: 14px;
    }}

    .stTabs [data-baseweb="tab"] {{
        border-radius: 10px;
        padding: 8px 16px;
        font-weight: 600;
        color: #4B5563;
    }}

    .stTabs [aria-selected="true"] {{
        background: #ffffff !important;
        color: {BRAND_PRIMARY} !important;
        box-shadow: 0 2px 8px rgba(16,24,40,0.08);
    }}

    /* ---------- Sidebar ---------- */

    section[data-testid="stSidebar"] {{
        background: #14142B;
    }}

    section[data-testid="stSidebar"] * {{
        color: #E5E7EB !important;
    }}

    section[data-testid="stSidebar"] input {{
        color: #111827 !important;
    }}

    /* ---------- Chat bar (GPT style) ---------- */

    .chat-shell {{
        background: {BRAND_CARD};
        border-radius: 18px;
        border: 1px solid #ECECF5;
        box-shadow: 0 4px 18px rgba(16,24,40,0.06);
        padding: 6px 8px 8px 8px;
        margin-top: 10px;
    }}

    .chat-header {{
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 12px 14px 8px 14px;
    }}

    .chat-header-title {{
        font-weight: 700;
        font-size: 15px;
        color: #1F2937;
    }}

    .chat-header-dot {{
        width: 9px;
        height: 9px;
        border-radius: 50%;
        background: #22C55E;
        display: inline-block;
    }}

    div[data-testid="stChatInput"] {{
        border-radius: 14px;
    }}

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
    <div class="hero">
        <div class="hero-title">🏥 Multimodal AI Healthcare Assistant</div>
        <div class="hero-subtitle">Computer Vision · NLP · RAG · Generative AI, in one workspace</div>
        <div class="hero-badges">
            <span class="badge">👁️ Vision</span>
            <span class="badge">🧠 NLP</span>
            <span class="badge">📚 RAG</span>
            <span class="badge">🤖 Generative AI</span>
            <span class="badge">💬 Assistant Chat</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.warning(
    "⚠️ Educational hackathon prototype only. "
    "This application does not provide medical diagnosis, "
    "prescription, or treatment."
)


# ============================================================
# MEDICAL KNOWLEDGE BASE
# ============================================================

MEDICAL_KNOWLEDGE = [

    {
        "title": "Eczema / Dermatitis",
        "content": """
        Eczema and dermatitis can involve itchy, dry,
        irritated or red skin. Different forms can have
        different causes and appearances. Persistent,
        severe or worsening symptoms should be evaluated
        by a qualified healthcare professional.
        """
    },

    {
        "title": "Urticaria / Hives",
        "content": """
        Urticaria, commonly called hives, may appear as
        raised itchy areas on the skin. Triggers can
        include infections, medications, foods and
        other factors. Swelling involving the lips,
        tongue or throat or difficulty breathing can
        require urgent medical attention.
        """
    },

    {
        "title": "Folliculitis / Acne-like Conditions",
        "content": """
        Folliculitis and acne-like conditions can produce
        bumps, redness, pimples or irritation around hair
        follicles. Several skin conditions can look
        similar, so persistent or severe symptoms should
        be professionally assessed.
        """
    },

    {
        "title": "Psoriasis / Papulosquamous Conditions",
        "content": """
        Psoriasis can produce red or inflamed areas with
        scaling. Several skin conditions may overlap in
        appearance, so an image alone cannot establish
        a diagnosis.
        """
    },

    {
        "title": "Skin Lesion Review",
        "content": """
        Some skin lesions require professional examination.
        An AI image classifier cannot determine whether a
        lesion is cancerous. A lesion that is changing,
        bleeding, painful or persistent should be evaluated
        by a qualified healthcare professional.
        """
    },

    {
        "title": "General Skin Health",
        "content": """
        Skin appearance can vary with lighting, camera
        quality, skin tone, age and many other factors.
        AI image classification should therefore be treated
        as educational information rather than confirmed
        medical diagnosis.
        """
    }
]


# ============================================================
# NLP SYMPTOM DATABASE
# ============================================================

SYMPTOMS = [
    "itching", "itchy", "redness", "red", "swelling", "swollen",
    "pain", "painful", "burning", "dryness", "dry skin", "rash",
    "bumps", "pimples", "scaling", "scaly", "bleeding", "tenderness",
    "blister", "blisters", "discharge", "pus", "cracking", "cracked"
]

BODY_PARTS = [
    "face", "arm", "arms", "leg", "legs", "hand", "hands", "foot",
    "feet", "back", "chest", "neck", "head", "scalp", "stomach",
    "abdomen", "finger", "fingers", "toe", "toes"
]

EMERGENCY_TERMS = [
    "difficulty breathing", "cannot breathe", "can't breathe",
    "trouble breathing", "throat swelling", "tongue swelling",
    "face swelling", "severe bleeding", "uncontrolled bleeding",
    "unconscious", "chest pain", "fainting", "severe allergic reaction"
]


# ============================================================
# COMPUTER VISION MODEL
# ============================================================

@st.cache_resource
def load_cv_model():
    processor = AutoImageProcessor.from_pretrained(CV_MODEL)
    model = AutoModelForImageClassification.from_pretrained(CV_MODEL)
    model.eval()
    return processor, model


def predict_image(image):
    processor, model = load_cv_model()
    image = image.convert("RGB")
    inputs = processor(images=image, return_tensors="pt")

    with torch.no_grad():
        outputs = model(**inputs)

    probabilities = torch.softmax(outputs.logits, dim=-1)[0]
    top_k = min(3, len(probabilities))
    values, indices = torch.topk(probabilities, k=top_k)

    results = []
    for value, index in zip(values, indices):
        index = index.item()
        label = model.config.id2label.get(index, f"Class {index}")
        results.append({"condition": label, "confidence": float(value.item())})

    return results


# ============================================================
# NLP FUNCTIONS
# ============================================================

def _extract_terms(text, terms):
    text = text.lower()
    found = []
    for term in terms:
        pattern = r"\b" + re.escape(term) + r"\b"
        if re.search(pattern, text):
            found.append(term)
    return sorted(list(set(found)))


def extract_symptoms(text):
    return _extract_terms(text, SYMPTOMS)


def extract_body_parts(text):
    return _extract_terms(text, BODY_PARTS)


def extract_duration(text):
    text = text.lower()
    patterns = [
        r"\d+\s+days?", r"\d+\s+weeks?", r"\d+\s+months?",
        r"\d+\s+hours?", r"\d+\s+years?"
    ]
    durations = []
    for pattern in patterns:
        durations.extend(re.findall(pattern, text))
    return sorted(list(set(durations)))


def analyze_nlp(text):
    return {
        "symptoms": extract_symptoms(text),
        "body_parts": extract_body_parts(text),
        "duration": extract_duration(text)
    }


# ============================================================
# RAG ENGINE
# ============================================================

@st.cache_resource
def build_rag():
    documents = [item["title"] + " " + item["content"] for item in MEDICAL_KNOWLEDGE]
    vectorizer = TfidfVectorizer(stop_words="english")
    matrix = vectorizer.fit_transform(documents)
    return vectorizer, matrix


def retrieve_medical_information(query, top_k=3):
    vectorizer, matrix = build_rag()
    query_vector = vectorizer.transform([query])
    scores = cosine_similarity(query_vector, matrix)[0]
    indices = scores.argsort()[::-1][:top_k]

    results = []
    for index in indices:
        results.append({
            "title": MEDICAL_KNOWLEDGE[index]["title"],
            "content": MEDICAL_KNOWLEDGE[index]["content"].strip(),
            "score": float(scores[index])
        })
    return results


# ============================================================
# SAFETY
# ============================================================

def check_emergency(text):
    text = text.lower()
    return [term for term in EMERGENCY_TERMS if term in text]


# ============================================================
# GEMINI GENERATIVE AI
# ============================================================

def get_gemini_client():
    api_key = st.session_state.get("gemini_api_key")
    if not api_key:
        return None
    return genai.Client(api_key=api_key)


def generate_ai_response(
    symptoms, body_parts, duration, cv_results,
    retrieved_information, original_text, emergency_terms
):
    client = get_gemini_client()

    if client is None:
        return (
            "Gemini API key is not configured.\n\n"
            "Please enter your Gemini API key in the sidebar, "
            "or configure it via Streamlit secrets."
        )

    cv_text = "\n".join(
        f"- {item['condition']}: {item['confidence']:.2%}" for item in cv_results
    )

    rag_text = "\n\n".join(
        f"### {item['title']}\n{item['content']}" for item in retrieved_information
    )

    emergency_text = ", ".join(emergency_terms) if emergency_terms else "None detected"

    prompt = f"""
You are an educational healthcare AI assistant for a hackathon prototype.

You are NOT a doctor.

The system must NOT claim to diagnose the user.
Do not prescribe medicines.
Do not recommend prescription treatment.
Do not describe model confidence as a medical probability.

Use cautious language such as:
- may be associated with
- could be consistent with
- requires professional assessment

USER DESCRIPTION:
{original_text}

EXTRACTED SYMPTOMS:
{symptoms}

BODY PARTS:
{body_parts}

DURATION:
{duration}

COMPUTER VISION MODEL OUTPUT:
{cv_text}

RETRIEVED KNOWLEDGE:
{rag_text}

EMERGENCY TERMS DETECTED:
{emergency_text}

Create a clear educational health summary.

Use these sections:

## AI Health Summary
## Computer Vision Observation
## Reported Symptoms
## Relevant Medical Information
## Possible Explanations
## Recommended Next Steps
## When to Seek Urgent Care
## Important Disclaimer

Keep the explanation understandable for a general user.
Do not make a definitive diagnosis.
If emergency warning signs were detected, clearly tell the user
that urgent professional medical care may be appropriate.
The application is an educational prototype and not a substitute
for a qualified healthcare professional.
"""

    try:
        response = client.models.generate_content(model=GEMINI_MODEL, contents=prompt)
        return response.text
    except Exception as error:
        return f"Unable to generate the AI response.\n\nTechnical error: {error}"


def generate_chat_reply(user_message, context):
    """
    Follow-up chat turn in the GPT-style assistant bar.
    `context` carries the latest analysis so the assistant can
    answer questions grounded in it, still with the same safety rules.
    """
    client = get_gemini_client()

    if client is None:
        return (
            "I need a Gemini API key to reply — please add one in the "
            "sidebar (or via Streamlit secrets) first."
        )

    history_text = ""
    for turn in st.session_state.get("chat_history", [])[-6:]:
        role = "User" if turn["role"] == "user" else "Assistant"
        history_text += f"{role}: {turn['content']}\n"

    prompt = f"""
You are an educational healthcare assistant chatbot for a hackathon
prototype. You are NOT a doctor, you do not diagnose, and you do not
prescribe or recommend specific medications or dosages.

Use cautious, non-diagnostic language ("may be associated with",
"could be consistent with", "worth discussing with a professional").
If the user describes an emergency warning sign (e.g. trouble
breathing, chest pain, severe bleeding, facial/throat swelling,
loss of consciousness), tell them clearly to seek urgent in-person
medical care right away.

Relevant context from the user's most recent multimodal analysis
on this page (may be empty if none has been run yet):
{context}

Recent conversation:
{history_text}

New user message:
{user_message}

Reply conversationally, concisely, and helpfully. Keep it a few
short paragraphs at most unless the user asks for more detail.
"""

    try:
        response = client.models.generate_content(model=GEMINI_MODEL, contents=prompt)
        return response.text
    except Exception as error:
        return f"Sorry, I couldn't generate a reply.\n\nTechnical error: {error}"


# ============================================================
# WORD (.docx) REPORT
# ============================================================

def _set_cell_background(cell, hex_color):
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shading)


def create_word_report(original_symptoms, cv_results, nlp_results, ai_response):
    document = Document()

    # Base font
    style = document.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # Title
    title = document.add_heading("Multimodal AI Healthcare Assistant", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = document.add_paragraph("Educational AI Health Report")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].italic = True
    subtitle.runs[0].font.color.rgb = RGBColor(0x6B, 0x72, 0x80)

    document.add_paragraph()

    # User description
    document.add_heading("User Description", level=1)
    document.add_paragraph(original_symptoms)

    # NLP findings
    document.add_heading("NLP Findings", level=1)
    nlp_table = document.add_table(rows=3, cols=2)
    nlp_table.style = "Light Grid Accent 1"

    rows = [
        ("Symptoms", ", ".join(nlp_results["symptoms"]) or "None detected"),
        ("Body Parts", ", ".join(nlp_results["body_parts"]) or "None detected"),
        ("Duration", ", ".join(nlp_results["duration"]) or "None detected"),
    ]
    for i, (label, value) in enumerate(rows):
        nlp_table.cell(i, 0).text = label
        nlp_table.cell(i, 0).paragraphs[0].runs[0].bold = True
        nlp_table.cell(i, 1).text = value

    document.add_paragraph()

    # Computer vision results
    document.add_heading("Computer Vision Results", level=1)
    cv_table = document.add_table(rows=1, cols=2)
    cv_table.style = "Light Grid Accent 1"
    cv_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    header_cells = cv_table.rows[0].cells
    header_cells[0].text = "Observation"
    header_cells[1].text = "Model Score"
    for cell in header_cells:
        _set_cell_background(cell, "E5E7EB")
        cell.paragraphs[0].runs[0].bold = True

    for result in cv_results:
        row_cells = cv_table.add_row().cells
        row_cells[0].text = result["condition"]
        row_cells[1].text = f"{result['confidence']:.2%}"

    document.add_paragraph()
    note = document.add_paragraph(
        "The displayed score is a model classification score, "
        "not a clinical probability or diagnosis."
    )
    note.runs[0].italic = True
    note.runs[0].font.size = Pt(9)
    note.runs[0].font.color.rgb = RGBColor(0x6B, 0x72, 0x80)

    document.add_paragraph()

    # AI-generated summary
    document.add_heading("AI-Generated Health Information", level=1)

    for line in ai_response.split("\n"):
        line = line.strip()
        if not line:
            continue

        if line.startswith("## "):
            document.add_heading(line.replace("## ", ""), level=2)
        else:
            clean_line = line.replace("**", "")
            document.add_paragraph(clean_line)

    document.add_paragraph()

    # Disclaimer
    document.add_heading("Important Notice", level=1)
    disclaimer = document.add_paragraph()
    disclaimer_run = disclaimer.add_run(
        "This report is generated by an AI hackathon prototype for "
        "educational purposes only. It is not a medical diagnosis or "
        "treatment plan. Consult a qualified healthcare professional "
        "for any health concerns."
    )
    disclaimer_run.bold = True

    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer


# ============================================================
# SESSION STATE INIT
# ============================================================

if "chat_history" not in st.session_state:
    st.session_state["chat_history"] = []

if "last_analysis_context" not in st.session_state:
    st.session_state["last_analysis_context"] = ""


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Settings")

    default_key = resolve_default_api_key()

    api_key_input = st.text_input(
        "Gemini API Key",
        value="",
        type="password",
        help=(
            "Loaded automatically from Streamlit secrets or the "
            "GEMINI_API_KEY environment variable if set there. "
            "Enter a key here to override it for this session only. "
            "Never hardcode a real key into the source file."
        )
    )

    st.session_state["gemini_api_key"] = api_key_input or default_key

    if default_key and not api_key_input:
        st.caption("✅ Using key from secrets / environment.")
    elif api_key_input:
        st.caption("✅ Using key entered above (this session only).")
    else:
        st.caption("⚠️ No API key configured yet.")

    st.divider()

    st.subheader("AI Modules")
    st.write("👁️ Computer Vision")
    st.write("🧠 NLP")
    st.write("📚 RAG")
    st.write("🤖 Generative AI")
    st.write("💬 Assistant Chat")

    st.divider()

    st.caption("Hackathon Prototype")


# ============================================================
# INPUT AREA
# ============================================================

st.markdown('<div class="card">', unsafe_allow_html=True)

image_col, text_col = st.columns(2)

with image_col:
    st.markdown('<div class="section-title">📷 Medical / Skin Image</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload an image", type=["jpg", "jpeg", "png"])

with text_col:
    st.markdown('<div class="section-title">📝 Symptoms</div>', unsafe_allow_html=True)
    symptoms_text = st.text_area(
        "Describe your symptoms",
        height=180,
        placeholder="Example:\nI have itchy red patches on my arm for 5 days."
    )

st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# IMAGE PREVIEW
# ============================================================

image = None

if uploaded_file:
    try:
        image = Image.open(uploaded_file).convert("RGB")
        st.image(image, caption="Uploaded Image", width="stretch")
    except Exception:
        st.error("Unable to read this image.")


# ============================================================
# ANALYZE BUTTON
# ============================================================

st.divider()

analyze_button = st.button(
    "🔍 Analyze Health Information",
    type="primary",
    use_container_width=True
)


# ============================================================
# MAIN ANALYSIS
# ============================================================

if analyze_button:

    if image is None:
        st.error("Please upload an image.")
        st.stop()

    if not symptoms_text.strip():
        st.error("Please describe your symptoms.")
        st.stop()

    emergency_terms = check_emergency(symptoms_text)

    if emergency_terms:
        st.error("⚠️ Potential urgent warning sign detected in the entered text.")
        st.write("Detected terms: " + ", ".join(emergency_terms))
        st.warning(
            "This application cannot determine whether an emergency is "
            "occurring. Seek appropriate urgent professional medical "
            "care when necessary."
        )

    with st.spinner("👁️ Computer Vision is analyzing the image..."):
        try:
            cv_results = predict_image(image)
        except Exception as error:
            st.error("Computer Vision model error.")
            st.exception(error)
            st.stop()

    with st.spinner("🧠 NLP is processing symptoms..."):
        nlp_results = analyze_nlp(symptoms_text)

    cv_conditions = " ".join(item["condition"] for item in cv_results)
    symptom_words = " ".join(nlp_results["symptoms"])
    body_words = " ".join(nlp_results["body_parts"])

    rag_query = symptoms_text + " " + cv_conditions + " " + symptom_words + " " + body_words

    with st.spinner("📚 Retrieving relevant medical information..."):
        retrieved_information = retrieve_medical_information(rag_query, top_k=3)

    with st.spinner("🤖 Generative AI is creating the health summary..."):
        ai_response = generate_ai_response(
            symptoms=nlp_results["symptoms"],
            body_parts=nlp_results["body_parts"],
            duration=nlp_results["duration"],
            cv_results=cv_results,
            retrieved_information=retrieved_information,
            original_text=symptoms_text,
            emergency_terms=emergency_terms
        )

    # Save context for the chat assistant to reference later
    st.session_state["last_analysis_context"] = (
        f"User description: {symptoms_text}\n"
        f"Symptoms detected: {', '.join(nlp_results['symptoms']) or 'none'}\n"
        f"Body parts mentioned: {', '.join(nlp_results['body_parts']) or 'none'}\n"
        f"Duration mentioned: {', '.join(nlp_results['duration']) or 'none'}\n"
        f"Top CV observation: {cv_results[0]['condition']} "
        f"({cv_results[0]['confidence']:.2%})\n"
        f"AI summary:\n{ai_response}"
    )

    st.success("Multimodal analysis completed.")

    tab1, tab2, tab3, tab4 = st.tabs(
        ["👁️ Computer Vision", "🧠 NLP", "📚 RAG", "🤖 Generative AI"]
    )

    with tab1:
        st.subheader("Computer Vision Results")
        for result in cv_results:
            st.write(f"**{result['condition']}**")
            st.progress(result["confidence"])
            st.caption(f"Model score: {result['confidence']:.2%}")
        st.info(
            "The displayed score is a model classification score, not a "
            "clinical probability or diagnosis."
        )

    with tab2:
        st.subheader("NLP Symptom Analysis")
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("### Symptoms")
            if nlp_results["symptoms"]:
                for symptom in nlp_results["symptoms"]:
                    st.write("• " + symptom)
            else:
                st.write("No predefined symptoms detected.")

        with col2:
            st.markdown("### Body Parts")
            if nlp_results["body_parts"]:
                for part in nlp_results["body_parts"]:
                    st.write("• " + part)
            else:
                st.write("Not detected.")

        with col3:
            st.markdown("### Duration")
            if nlp_results["duration"]:
                for duration in nlp_results["duration"]:
                    st.write("• " + duration)
            else:
                st.write("Not detected.")

    with tab3:
        st.subheader("Retrieved Medical Knowledge")
        for item in retrieved_information:
            with st.expander(f"{item['title']}  |  Relevance: {item['score']:.2f}"):
                st.write(item["content"])

    with tab4:
        st.subheader("AI Healthcare Assistant")
        st.markdown(ai_response)

    st.divider()
    st.header("🩺 Multimodal Health Summary")

    summary_col1, summary_col2 = st.columns(2)

    with summary_col1:
        st.markdown("### 👁️ Visual Analysis")
        st.write(cv_results[0]["condition"])
        st.caption("Top computer vision classification result.")

    with summary_col2:
        st.markdown("### 🧠 Patient Symptoms")
        if nlp_results["symptoms"]:
            st.write(", ".join(nlp_results["symptoms"]))
        else:
            st.write("No predefined symptoms detected.")

    # --------------------------------------------------------
    # WORD REPORT
    # --------------------------------------------------------

    st.divider()
    st.header("📄 Generate Health Report")

    word_report = create_word_report(
        original_symptoms=symptoms_text,
        cv_results=cv_results,
        nlp_results=nlp_results,
        ai_response=ai_response
    )

    st.download_button(
        label="📥 Download Word Report (.docx)",
        data=word_report,
        file_name="multimodal_healthcare_report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )


# ============================================================
# GPT-STYLE ASSISTANT CHAT BAR
# ============================================================

st.divider()

st.markdown('<div class="chat-shell">', unsafe_allow_html=True)
st.markdown(
    """
    <div class="chat-header">
        <span class="chat-header-dot"></span>
        <span class="chat-header-title">Assistant Chat</span>
        <span class="small-text">— ask follow-up questions about your analysis</span>
    </div>
    """,
    unsafe_allow_html=True
)

chat_container = st.container(height=380)

with chat_container:
    if not st.session_state["chat_history"]:
        st.markdown(
            '<p class="small-text" style="padding: 0 14px;">'
            'Run an analysis above, then ask things like '
            '"what does folliculitis mean?" or "should I be worried?" '
            'This chat is educational only and does not replace medical care.'
            '</p>',
            unsafe_allow_html=True
        )

    for turn in st.session_state["chat_history"]:
        with st.chat_message(turn["role"]):
            st.markdown(turn["content"])

chat_prompt = st.chat_input("Message the assistant...")

if chat_prompt:
    st.session_state["chat_history"].append({"role": "user", "content": chat_prompt})

    emergency_hits = check_emergency(chat_prompt)

    with st.spinner("Thinking..."):
        reply = generate_chat_reply(
            chat_prompt,
            st.session_state.get("last_analysis_context", "")
        )

        if emergency_hits:
            reply = (
                "⚠️ **This may describe an urgent situation** "
                f"(detected: {', '.join(emergency_hits)}). "
                "Please seek immediate in-person medical care or emergency "
                "services if this is happening now.\n\n" + reply
            )

    st.session_state["chat_history"].append({"role": "assistant", "content": reply})
    st.rerun()

st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Multimodal AI Healthcare Assistant | "
    "Computer Vision + NLP + RAG + Generative AI"
)

st.caption(
    "⚠️ Educational prototype only. Not a substitute for professional "
    "medical advice, diagnosis, or treatment."
)

import re
import io
import sqlite3
import hashlib
from datetime import datetime

import pymupdf
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from docx import Document
from docx.shared import Pt
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

st.set_page_config(
    page_title="InterviewMate | Career Intelligence",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --------------------------- PROFESSIONAL UI ---------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@600;700;800&display=swap');

:root {
    --ink:#182033;
    --muted:#697386;
    --line:#e6e9f0;
    --accent:#5b5bd6;
}

html, body, [class*="css"] { font-family:"DM Sans",sans-serif; color:var(--ink); }

.stApp {
    background:
      radial-gradient(circle at 5% 0%, rgba(124,108,242,.12), transparent 27%),
      radial-gradient(circle at 96% 7%, rgba(91,91,214,.08), transparent 25%),
      #f6f7fb;
}

.block-container { max-width:1420px; padding:2.2rem 3rem 3rem; }

[data-testid="stSidebar"] {
    background:#111827;
    border-right:1px solid rgba(255,255,255,.06);
}
[data-testid="stSidebar"] * { color:#e8ebf2; }

.hero {
    padding:34px 38px;
    border:1px solid rgba(255,255,255,.8);
    border-radius:28px;
    background:linear-gradient(135deg,#ffffff,#f9f9ff);
    box-shadow:0 18px 55px rgba(25,32,56,.08);
    margin-bottom:24px;
}
.eyebrow {
    color:#5b5bd6;
    font-size:11px;
    font-weight:800;
    letter-spacing:.17em;
    text-transform:uppercase;
    margin-bottom:10px;
}
.hero h1 {
    font-family:"Manrope",sans-serif;
    font-size:clamp(2rem,4vw,3.4rem);
    line-height:1.05;
    letter-spacing:-.055em;
    margin:0 0 12px;
    color:#121827;
}
.hero p {
    color:#697386;
    font-size:16px;
    max-width:800px;
    line-height:1.65;
    margin:0;
}

.section-title {
    font-family:"Manrope",sans-serif;
    font-size:1.3rem;
    font-weight:800;
    letter-spacing:-.025em;
    margin:28px 0 12px;
}

.card {
    background:rgba(255,255,255,.92);
    border:1px solid #e6e9f0;
    border-radius:20px;
    padding:20px;
    min-height:120px;
    box-shadow:0 8px 28px rgba(25,32,56,.045);
}
.mini {
    color:#8b94a5;
    text-transform:uppercase;
    letter-spacing:.12em;
    font-size:10px;
    font-weight:800;
}
.big {
    font-family:"Manrope",sans-serif;
    font-size:29px;
    font-weight:800;
    margin:8px 0 3px;
}
.desc { color:#697386; font-size:13px; line-height:1.5; }

.pill {
    display:inline-block;
    background:#f0efff;
    color:#4e4ec2;
    border-radius:999px;
    padding:6px 11px;
    margin:4px 4px 0 0;
    font-size:12px;
    font-weight:700;
}

.callout {
    border-left:4px solid #5b5bd6;
    background:#faf9ff;
    padding:14px 17px;
    border-radius:0 14px 14px 0;
    color:#50596b;
    margin:14px 0;
}

.workflow {
    display:grid;
    grid-template-columns:repeat(6,1fr);
    gap:12px;
}
.workflow-item {
    background:#fff;
    border:1px solid #e6e9f0;
    border-radius:18px;
    padding:17px;
    min-height:130px;
}
.workflow-no {
    width:28px;height:28px;display:grid;place-items:center;
    border-radius:50%;background:#efedff;color:#5b5bd6;
    font-weight:800;font-size:11px;margin-bottom:12px;
}
.workflow-item b { display:block; margin-bottom:5px; }
.workflow-item span { color:#697386; font-size:12px; line-height:1.45; }

div[data-testid="stMetric"] {
    background:rgba(255,255,255,.9);
    border:1px solid #e6e9f0;
    padding:16px 18px;
    border-radius:17px;
    box-shadow:0 7px 25px rgba(25,32,56,.04);
}
div[data-testid="stExpander"] {
    background:#fff;
    border:1px solid #e6e9f0;
    border-radius:15px;
    margin-bottom:8px;
}
.stButton > button {
    border-radius:12px;
    min-height:43px;
    font-weight:700;
}
[data-testid="stFileUploader"] {
    background:#fff;
    border:1px dashed #cdd2df;
    border-radius:16px;
    padding:8px;
}
.footer {
    text-align:center;
    color:#9299a8;
    font-size:11px;
    letter-spacing:.04em;
    padding:18px 0 2px;
}
@media (max-width:900px) {
    .block-container { padding:1.2rem; }
    .workflow { grid-template-columns:1fr; }
}
</style>
""", unsafe_allow_html=True)

def hero(kicker, title, description):
    st.markdown(
        f'<div class="hero"><div class="eyebrow">{kicker}</div>'
        f'<h1>{title}</h1><p>{description}</p></div>',
        unsafe_allow_html=True
    )

def pills(items):
    if not items:
        return '<span class="desc">Nothing detected yet.</span>'
    return ''.join(f'<span class="pill">{x.title()}</span>' for x in items)
# ----------------------------------------------------------------------

DB_NAME = "interviewmate.db"

def init_database():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        identifier TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        created_at TEXT
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS sessions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        created_at TEXT,
        match_score REAL,
        questions_attempted INTEGER,
        average_score REAL
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS attempts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        created_at TEXT,
        module TEXT NOT NULL,
        topic TEXT,
        score REAL
    )""")
    conn.commit()
    conn.close()

def save_session(match_score, questions_attempted, average_score):
    conn = sqlite3.connect(DB_NAME)
    conn.execute(
        "INSERT INTO sessions (created_at, match_score, questions_attempted, average_score) VALUES (?, ?, ?, ?)",
        (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), match_score, questions_attempted, average_score)
    )
    conn.commit()
    conn.close()

def get_sessions():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT * FROM sessions ORDER BY id DESC", conn)
    conn.close()
    return df

def save_attempt(module, topic, score):
    """Store one assessment/interview attempt in local SQLite history."""
    conn = sqlite3.connect(DB_NAME)
    conn.execute(
        "INSERT INTO attempts (created_at, module, topic, score) VALUES (?, ?, ?, ?)",
        (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), str(module), str(topic), float(score))
    )
    conn.commit()
    conn.close()

def get_attempts():
    """Return all saved practice attempts as a DataFrame."""
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query(
        "SELECT id, created_at, module, topic, score FROM attempts ORDER BY id DESC",
        conn
    )
    conn.close()
    return df

def save_session_summary(match_score, readiness):
    """Save a readiness snapshot using the existing sessions table."""
    attempts = get_attempts()
    average_score = round(float(attempts["score"].mean()), 1) if not attempts.empty else float(readiness)
    save_session(float(match_score), int(len(attempts)), average_score)

def score_badge(score):
    """Return a compact HTML badge for a readiness score."""
    score = float(score)
    if score >= 80:
        label = "STRONG"
    elif score >= 60:
        label = "ON TRACK"
    else:
        label = "NEEDS PRACTICE"
    return f'<span class="pill">{label}</span>'

init_database()

def hash_password(password):
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def register_user(identifier, password):
    identifier = identifier.strip().lower()
    if len(identifier) < 3 or len(password) < 6:
        return False, "Use a valid email/phone/username and a password of at least 6 characters."
    conn = sqlite3.connect(DB_NAME)
    try:
        conn.execute("INSERT INTO users (identifier, password_hash, created_at) VALUES (?, ?, ?)",
                     (identifier, hash_password(password), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        return True, "Account created successfully. You can now log in."
    except sqlite3.IntegrityError:
        return False, "That email, phone number, or username is already registered."
    finally:
        conn.close()

def authenticate_user(identifier, password):
    conn = sqlite3.connect(DB_NAME)
    row = conn.execute("SELECT identifier, password_hash FROM users WHERE identifier=?", (identifier.strip().lower(),)).fetchone()
    conn.close()
    return bool(row and row[1] == hash_password(password))

def login_screen():
    st.markdown("""<div style="max-width:620px;margin:7vh auto 0;">
    <div class="hero"><div class="eyebrow">INTERVIEWMATE · CAREER INTELLIGENCE</div>
    <h1>Prepare for the interview. Own the opportunity.</h1>
    <p>Sign in with your email, phone number, or username. New here? Create a free local account.</p></div></div>""", unsafe_allow_html=True)
    tab1, tab2 = st.tabs(["🔐 Login", "✨ Create Account"])
    with tab1:
        identifier = st.text_input("Email / Phone / Username", key="login_identifier")
        password = st.text_input("Password", type="password", key="login_password")
        if st.button("Login to InterviewMate", type="primary", use_container_width=True):
            if authenticate_user(identifier, password):
                st.session_state.authenticated = True
                st.session_state.user_identifier = identifier.strip()
                st.rerun()
            else:
                st.error("❌ Invalid login details. Check your identifier and password.")
    with tab2:
        identifier = st.text_input("Email / Phone / Username", key="signup_identifier")
        password = st.text_input("Create Password", type="password", key="signup_password")
        confirm = st.text_input("Confirm Password", type="password", key="signup_confirm")
        if st.button("Create Account", type="primary", use_container_width=True):
            if password != confirm:
                st.error("❌ Passwords do not match.")
            else:
                ok, msg = register_user(identifier, password)
                (st.success if ok else st.error)(msg)

if not st.session_state.get("authenticated", False):
    login_screen()
    st.stop()

SKILLS = [
    "python","java","c","c++","javascript","typescript","html","css","react","node.js",
    "flask","django","streamlit","sql","mysql","postgresql","mongodb","nosql",
    "machine learning","deep learning","artificial intelligence","natural language processing",
    "nlp","computer vision","tensorflow","pytorch","keras","scikit-learn","pandas","numpy",
    "opencv","aws","azure","docker","git","github","linux","cloud computing","data science",
    "data analysis","power bi","tableau","excel","statistics","rest api","api","cybersecurity",
    "blockchain","communication","leadership","teamwork","problem solving","time management"
]

def extract_pdf(file):
    text = ""
    try:
        pdf = pymupdf.open(stream=file.read(), filetype="pdf")
        for page in pdf:
            text += page.get_text()
        pdf.close()
    except Exception as e:
        st.error(f"Unable to read PDF: {e}")
    return text

def extract_docx(file):
    text = ""
    try:
        document = Document(io.BytesIO(file.read()))
        for paragraph in document.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        st.error(f"Unable to read DOCX: {e}")
    return text

def extract_file_text(uploaded_file):
    if uploaded_file is None:
        return ""
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        return extract_pdf(uploaded_file)
    if name.endswith(".docx"):
        return extract_docx(uploaded_file)
    if name.endswith(".txt"):
        return uploaded_file.read().decode("utf-8", errors="ignore")
    return ""

def clean_text(text):
    return re.sub(r"\s+", " ", text.lower()).strip()

# --------------------------- DOCUMENT VALIDATION ---------------------------

RESUME_SECTION_KEYWORDS = [
    "education", "educational qualification", "academic qualification",
    "skills", "technical skills", "technical skill", "work experience",
    "professional experience", "experience", "projects", "project",
    "certifications", "certification", "internship", "internships",
    "achievements", "achievement", "career objective", "objective",
    "professional summary", "summary", "profile", "about me",
    "languages", "publications", "awards", "extra curricular",
    "extracurricular", "responsibilities", "employment"
]

RESUME_CONTACT_PATTERNS = [
    r"\b[\w.+-]+@[\w-]+\.[\w.-]+\b",
    r"(?:linkedin\.com|github\.com)",
    r"\+?\d[\d\s().-]{8,}\d"
]

NON_RESUME_KEYWORDS = [
    "invoice", "tax invoice", "bill no", "invoice no", "receipt",
    "gstin", "gst", "subtotal", "amount due", "total amount",
    "unit price", "unit cost", "quantity", "qty", "customer id",
    "customer name", "transaction id", "transaction date", "order id",
    "order no", "payment received", "balance due", "billing address",
    "shipping address", "product description", "hsn", "sac code",
    "cash memo", "purchase order"
]

ACADEMIC_TERMS = [
    "b.tech", "btech", "b.e.", "be ", "m.tech", "mtech", "mca",
    "b.sc", "bsc", "bca", "mba", "b.com", "degree", "university",
    "college", "school", "cgpa", "gpa", "percentage", "graduation"
]


def validate_resume_document(text):
    """
    Returns (is_valid, reason, confidence).
    Uses document structure rather than blindly extracting skills.
    This intentionally errs on the side of rejecting suspicious documents.
    """
    normalized = clean_text(text)

    if not normalized:
        return False, "The document is empty or could not be read.", 0

    if len(normalized) < 120:
        return False, "The document contains too little text to be a resume.", 10

    section_hits = [
        keyword for keyword in RESUME_SECTION_KEYWORDS
        if re.search(r"(?<!\w)" + re.escape(keyword) + r"(?!\w)", normalized)
    ]

    non_resume_hits = [
        keyword for keyword in NON_RESUME_KEYWORDS
        if re.search(r"(?<!\w)" + re.escape(keyword) + r"(?!\w)", normalized)
    ]

    contact_hits = sum(
        bool(re.search(pattern, text, flags=re.IGNORECASE))
        for pattern in RESUME_CONTACT_PATTERNS
    )

    academic_hits = sum(
        bool(re.search(r"(?<!\w)" + re.escape(term.strip()) + r"(?!\w)", normalized))
        for term in ACADEMIC_TERMS
    )

    # Strongly reject common transactional/business documents.
    if len(non_resume_hits) >= 3 and len(section_hits) < 3:
        return (
            False,
            "This document appears to be an invoice, bill, receipt, or other non-resume document.",
            5
        )

    # A document with several invoice signals should not pass merely because
    # a random word happens to match a technical skill.
    if len(non_resume_hits) >= 2 and len(section_hits) < 4:
        return (
            False,
            "This does not have the structure of a resume. Please upload your CV/resume.",
            10
        )

    # Resume-like documents normally contain multiple recognizable sections.
    if len(section_hits) >= 4:
        return True, "Resume structure detected.", min(95, 60 + len(section_hits) * 5)

    if len(section_hits) >= 3 and (contact_hits >= 1 or academic_hits >= 1):
        return True, "Resume-like structure detected.", 75

    if len(section_hits) >= 2 and contact_hits >= 1 and academic_hits >= 1:
        return True, "Resume-like structure detected.", 70

    return (
        False,
        "The uploaded document does not appear to be a resume. "
        "Please upload a CV/resume containing sections such as Education, Skills, "
        "Experience, Projects, or Certifications.",
        20
    )


def show_invalid_resume_message(reason):
    st.error("❌ Incorrect document uploaded")
    st.warning("Please upload a valid resume / CV.")
    st.info(
        f"**Why it was rejected:** {reason}\n\n"
        "Your document should normally contain sections such as "
        "**Education, Skills, Experience, Projects, Certifications, "
        "Internship, or Achievements**."
    )


def extract_skills(text):
    text = clean_text(text)
    found = []
    for skill in SKILLS:
        if re.search(r"(?<!\w)" + re.escape(skill) + r"(?!\w)", text):
            found.append(skill)
    return sorted(set(found))

def calculate_similarity(resume, jd):
    if not resume.strip() or not jd.strip():
        return 0
    try:
        vectorizer = TfidfVectorizer(stop_words="english", ngram_range=(1, 2))
        vectors = vectorizer.fit_transform([clean_text(resume), clean_text(jd)])
        return round(cosine_similarity(vectors[0:1], vectors[1:2])[0][0] * 100, 2)
    except Exception:
        return 0

def calculate_skill_match(resume, jd):
    resume_skills = set(extract_skills(resume))
    jd_skills = set(extract_skills(jd))
    matched = sorted(resume_skills & jd_skills)
    missing = sorted(jd_skills - resume_skills)
    score = (len(matched) / len(jd_skills) * 100) if jd_skills else 0
    return round(score, 2), matched, missing

def calculate_match_score(resume, jd):
    similarity = calculate_similarity(resume, jd)
    skill_score, matched, missing = calculate_skill_match(resume, jd)
    final = similarity * 0.60 + skill_score * 0.40
    return round(final, 2), similarity, skill_score, matched, missing

def generate_questions(resume, jd, matched_skills, missing_skills):
    questions = []
    r = resume.lower()
    if "machine learning" in r:
        questions.append("Explain one machine learning project from your resume. What problem did you solve and why did you choose your model?")
    if "python" in r:
        questions.append("How have you used Python in your projects? Explain one practical example.")
    if "sql" in r:
        questions.append("Explain a situation where you used SQL to retrieve or analyze data.")
    if "deep learning" in r:
        questions.append("Explain the deep learning model you used in your project and why you selected it.")
    if "computer vision" in r:
        questions.append("Explain how computer vision was used in one of your projects.")
    if "nlp" in r or "natural language processing" in r:
        questions.append("Explain an NLP technique you have implemented and where it can be used.")
    for skill in matched_skills[:5]:
        questions.append(f"Your profile mentions {skill}. Explain your practical experience with {skill}.")
    for skill in missing_skills[:4]:
        questions.append(f"The job description requires {skill}. What do you know about {skill}, and how would you learn it if required for this role?")
    questions += [
        "Tell me about yourself.",
        "Why are you interested in this role?",
        "Tell me about a difficult problem you faced in a project and how you solved it.",
        "Describe a situation where you worked as part of a team.",
        "What is one weakness you are currently working to improve?",
        "Where do you see yourself professionally in the next few years?"
    ]
    return list(dict.fromkeys(questions))[:15]

STOP_WORDS = {"the","is","a","an","and","or","to","of","in","on","for","with","this","that","i","my","we","was","were","it","as","at","be","have","has","had","from"}

def tokenize(text):
    return [w for w in re.findall(r"\b[a-zA-Z]{2,}\b", text.lower()) if w not in STOP_WORDS]

def analyze_answer(question, answer):
    if not answer.strip():
        return {"score": 0, "relevance": 0, "length": 0, "feedback": "Please provide an answer."}
    qw, aw = set(tokenize(question)), set(tokenize(answer))
    relevance = len(qw & aw) / len(qw) * 100 if qw else 50
    length = len(tokenize(answer))
    length_score = 100 if 40 <= length <= 180 else 70 if 20 <= length < 40 else 75 if 180 < length <= 250 else 40
    structure_words = ["because","therefore","first","then","finally","result","learned","experience"]
    structure = min(sum(x in answer.lower() for x in structure_words) * 20, 100)
    score = round(relevance * .4 + length_score * .3 + structure * .3, 2)
    feedback = ("Strong answer. Keep it concise and support your points with a specific example."
                if score >= 80 else
                "Good attempt. Add a specific example, measurable result, or clearer structure."
                if score >= 60 else
                "Try to make the answer more specific. Explain what you did, how you did it, and what result you achieved.")
    return {"score": score, "relevance": round(relevance,2), "length": length, "feedback": feedback}


# --------------------------- SMART RESUME TAILOR ---------------------------

def extract_resume_sections(text):
    """Extract common resume sections while preserving the candidate's content."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    headings = {
        "summary": ["summary", "professional summary", "profile", "career objective", "objective", "about me"],
        "skills": ["skills", "technical skills", "technical skill"],
        "experience": ["experience", "work experience", "professional experience", "employment"],
        "education": ["education", "educational qualification", "academic qualification"],
        "projects": ["projects", "project"],
        "certifications": ["certifications", "certification"],
        "achievements": ["achievements", "achievement", "awards"],
        "internships": ["internship", "internships"],
    }
    result = {}
    current = "other"
    result[current] = []
    for line in lines:
        normalized = re.sub(r"[:\-]+$", "", line.lower()).strip()
        matched = None
        for section, names in headings.items():
            if normalized in names:
                matched = section
                break
        if matched:
            current = matched
            result.setdefault(current, [])
        else:
            result.setdefault(current, []).append(line)
    return {k: v for k, v in result.items() if v}

def build_tailored_summary(original_text, jd, matched, missing):
    sections = extract_resume_sections(original_text)
    existing_summary = " ".join(sections.get("summary", [])).strip()
    role_keywords = extract_skills(jd)
    if existing_summary:
        base = existing_summary
    else:
        base = "Motivated candidate with hands-on academic and project experience"
    focus = [x.title() for x in matched[:6]]
    if focus:
        base += ", with practical exposure to " + ", ".join(focus) + "."
    else:
        base += "."
    if role_keywords:
        relevant = [x.title() for x in role_keywords if x in matched][:4]
        if relevant:
            base += " Interested in applying these skills in a role focused on " + ", ".join(relevant) + "."
    return base

def get_tailoring_suggestions(resume, jd, matched, missing):
    suggestions = []
    jd_lower = clean_text(jd)
    resume_lower = clean_text(resume)
    if matched:
        suggestions.append(
            "Highlight these JD-aligned skills more prominently: " +
            ", ".join(x.title() for x in matched[:10]) + "."
        )
    if missing:
        suggestions.append(
            "Review these JD requirements that were not detected in the resume: " +
            ", ".join(x.title() for x in missing[:10]) +
            ". Add them only if you genuinely have the skill or experience."
        )
    sections = extract_resume_sections(resume)
    if "summary" not in sections:
        suggestions.append("Add a short professional summary tailored to the target role.")
    if "projects" in sections:
        suggestions.append("Prioritize projects that demonstrate skills appearing in the job description.")
    if "experience" in sections:
        suggestions.append("Rewrite experience bullets to emphasize responsibilities and outcomes relevant to the target role.")
    if "certifications" in sections:
        suggestions.append("Place role-relevant certifications where they are easy for recruiters to find.")
    if not any(k in resume_lower for k in ["result", "improved", "increased", "reduced", "developed", "built"]):
        suggestions.append("Where truthful, add measurable outcomes or concrete results to project/experience bullets.")
    if "communication" in jd_lower and "communication" not in resume_lower:
        suggestions.append("If you have genuine communication experience, make it visible through activities, presentations, teamwork, or achievements.")
    return suggestions

def make_tailored_resume_text(resume, jd, matched, missing):
    sections = extract_resume_sections(resume)
    lines = []
    order = ["other", "summary", "skills", "experience", "internships", "projects", "education", "certifications", "achievements"]
    labels = {
        "summary":"PROFESSIONAL SUMMARY", "skills":"SKILLS", "experience":"EXPERIENCE",
        "internships":"INTERNSHIPS", "projects":"PROJECTS", "education":"EDUCATION",
        "certifications":"CERTIFICATIONS", "achievements":"ACHIEVEMENTS", "other":"PROFILE"
    }
    # Preserve original content; tailor ordering and add only a generated summary.
    if "other" in sections:
        lines.extend(sections["other"])
    lines.append(labels["summary"])
    lines.append(build_tailored_summary(resume, jd, matched, missing))
    for key in order:
        if key in ("other", "summary") or key not in sections:
            continue
        lines.append(labels[key])
        if key == "skills":
            original = sections[key]
            skill_text = " ".join(original)
            original_skills = extract_skills(skill_text)
            prioritized = []
            for x in matched:
                if x in original_skills:
                    prioritized.append(x)
            rest = [x for x in original_skills if x not in prioritized]
            if prioritized:
                lines.append("Priority skills: " + ", ".join(x.title() for x in prioritized))
            if rest:
                lines.append("Other skills: " + ", ".join(x.title() for x in rest))
            for x in original:
                if not any(x.lower() == y.lower() for y in original_skills):
                    lines.append(x)
        else:
            lines.extend(sections[key])
    return "\n".join(lines)

def create_docx_bytes(tailored_text):
    doc = Document()
    for i, line in enumerate(tailored_text.splitlines()):
        if not line.strip():
            continue
        p = doc.add_paragraph()
        if line.upper() == line and len(line) < 40:
            run = p.add_run(line)
            run.bold = True
            run.font.size = Pt(12)
        else:
            run = p.add_run(line)
            run.font.size = Pt(10.5)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

def create_pdf_bytes(tailored_text):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=A4, rightMargin=42, leftMargin=42, topMargin=42, bottomMargin=42)
    styles = getSampleStyleSheet()
    heading = ParagraphStyle("ResumeHeading", parent=styles["Heading2"], fontSize=11, leading=14, spaceBefore=9, spaceAfter=5)
    body = ParagraphStyle("ResumeBody", parent=styles["BodyText"], fontSize=9.5, leading=13, spaceAfter=4)
    story = []
    for line in tailored_text.splitlines():
        if not line.strip():
            continue
        safe = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if line.upper() == line and len(line) < 40:
            story.append(Paragraph(safe, heading))
        else:
            story.append(Paragraph(safe, body))
    doc.build(story)
    return bio.getvalue()

# --------------------------------------------------------------------------

# ============================================================
# COMPLETE PLACEMENT PREPARATION MODULES
# ============================================================
import random

APTITUDE_BANK = {
    "Quantitative Aptitude": [
        {"q":"A train travels 120 km in 2 hours. What is its average speed?", "o":["40 km/h","50 km/h","60 km/h","80 km/h"], "a":2, "e":"Speed = distance ÷ time = 120 ÷ 2 = 60 km/h."},
        {"q":"If a number is increased by 20% and becomes 360, what was the original number?", "o":["280","300","320","340"], "a":1, "e":"Original × 1.20 = 360, so original = 300."},
        {"q":"A product costs ₹800 after a 20% discount. What was its marked price?", "o":["₹960","₹980","₹1000","₹1020"], "a":2, "e":"80% of the marked price is ₹800, so the marked price is ₹1000."},
        {"q":"The ratio of boys to girls is 3:2. If there are 25 students, how many are boys?", "o":["10","12","15","18"], "a":2, "e":"Total parts = 5. Boys = 3/5 × 25 = 15."},
        {"q":"A can finish a job in 10 days and B in 15 days. How long will they take together?", "o":["5 days","6 days","7 days","8 days"], "a":1, "e":"Combined rate = 1/10 + 1/15 = 1/6, so 6 days."},
        {"q":"What is 15% of 240?", "o":["24","30","36","40"], "a":2, "e":"240 × 15/100 = 36."},
        {"q":"Simple interest on ₹5000 at 8% for 2 years is:", "o":["₹400","₹800","₹1000","₹1200"], "a":1, "e":"SI = PRT/100 = 5000×8×2/100 = ₹800."},
        {"q":"The average of 10, 20, 30 and 40 is:", "o":["20","25","30","35"], "a":1, "e":"Average = 100 ÷ 4 = 25."},
        {"q":"If 5 pens cost ₹75, what is the cost of 8 pens at the same rate?", "o":["₹100","₹110","₹120","₹125"], "a":2, "e":"One pen = ₹15, so 8 pens = ₹120."},
        {"q":"A number divided by 8 gives 15. What is the number?", "o":["100","110","120","130"], "a":2, "e":"8 × 15 = 120."},
    ],
    "Logical Reasoning": [
        {"q":"Find the next number: 2, 6, 12, 20, 30, ?", "o":["40","42","44","46"], "a":1, "e":"Differences are 4, 6, 8, 10, so next difference is 12 → 42."},
        {"q":"If CAT is coded as DBU, how is DOG coded?", "o":["EPH","EOG","DPH","FPH"], "a":0, "e":"Each letter shifts by +1: D→E, O→P, G→H."},
        {"q":"A is taller than B. B is taller than C. Who is shortest?", "o":["A","B","C","Cannot say"], "a":2, "e":"A > B > C, so C is shortest."},
        {"q":"Find the odd one out: Apple, Mango, Potato, Banana", "o":["Apple","Mango","Potato","Banana"], "a":2, "e":"Potato is a vegetable; the others are fruits."},
        {"q":"Complete the series: AZ, BY, CX, DW, ?", "o":["EV","EU","FV","EW"], "a":0, "e":"First letters move forward and second letters move backward: EV."},
        {"q":"If all engineers are graduates and some graduates are artists, what is definitely true?", "o":["All engineers are artists","Some engineers are artists","Engineers are graduates","All artists are engineers"], "a":2, "e":"Only the relation engineers → graduates is guaranteed."},
        {"q":"A person walks north, turns right, then turns right again. Which direction is he facing?", "o":["North","South","East","West"], "a":1, "e":"North → East → South."},
        {"q":"Book is to Reading as Fork is to:", "o":["Writing","Eating","Drawing","Sleeping"], "a":1, "e":"A book is used for reading; a fork is used for eating."},
    ],
    "Verbal Ability": [
        {"q":"Choose the synonym of 'Rapid'.", "o":["Slow","Quick","Weak","Late"], "a":1, "e":"Rapid means quick or fast."},
        {"q":"Choose the correctly spelled word.", "o":["Accomodation","Accommodation","Acommodation","Accommadation"], "a":1, "e":"The correct spelling is Accommodation."},
        {"q":"She is good ___ mathematics.", "o":["in","on","at","for"], "a":2, "e":"The standard expression is 'good at mathematics'."},
        {"q":"Choose the antonym of 'Expand'.", "o":["Increase","Stretch","Contract","Extend"], "a":2, "e":"Contract means to become smaller or reduce in size."},
        {"q":"Choose the most professional sentence.", "o":["Send me the file asap.","Kindly share the document at your convenience.","Give file now.","I need file."], "a":1, "e":"The second sentence is professional and polite."},
        {"q":"What does 'concise' mean?", "o":["Very detailed","Brief and clear","Confusing","Emotional"], "a":1, "e":"Concise means brief and clear."},
    ],
}

CODING_PROBLEMS = [
    {"title":"Two Sum", "difficulty":"Easy", "topic":"Arrays", "prompt":"Given a list of integers and a target, return the indices of two numbers whose sum equals the target.", "sample":"nums = [2,7,11,15], target = 9\nOutput: [0,1]", "hint":"Use a dictionary to remember numbers already seen.", "solution":"def two_sum(nums, target):\n    seen = {}\n    for i, x in enumerate(nums):\n        need = target - x\n        if need in seen:\n            return [seen[need], i]\n        seen[x] = i\n    return []"},
    {"title":"Palindrome String", "difficulty":"Easy", "topic":"Strings", "prompt":"Return True when a string reads the same forward and backward. Ignore letter case.", "sample":"Input: 'Madam'\nOutput: True", "hint":"Normalize the string, then compare it with its reverse.", "solution":"def is_palindrome(s):\n    s = s.lower()\n    return s == s[::-1]"},
    {"title":"Count Vowels", "difficulty":"Easy", "topic":"Strings", "prompt":"Count the number of vowels in a string.", "sample":"Input: 'InterviewMate'\nOutput: 5", "hint":"Check each character against a vowel set.", "solution":"def count_vowels(s):\n    vowels = set('aeiou')\n    return sum(ch.lower() in vowels for ch in s)"},
    {"title":"Find Maximum", "difficulty":"Easy", "topic":"Arrays", "prompt":"Find the largest element in a non-empty integer list without using max().", "sample":"Input: [4, 9, 2, 7]\nOutput: 9", "hint":"Keep a running best value.", "solution":"def find_max(nums):\n    best = nums[0]\n    for x in nums[1:]:\n        if x > best:\n            best = x\n    return best"},
    {"title":"Prime Check", "difficulty":"Easy", "topic":"Math", "prompt":"Return True if n is a prime number, otherwise False.", "sample":"Input: 29\nOutput: True", "hint":"You only need to test divisors up to sqrt(n).", "solution":"def is_prime(n):\n    if n < 2: return False\n    d = 2\n    while d * d <= n:\n        if n % d == 0: return False\n        d += 1\n    return True"},
    {"title":"Second Largest", "difficulty":"Easy", "topic":"Arrays", "prompt":"Find the second largest distinct number in a list.", "sample":"Input: [10, 5, 8, 10, 3]\nOutput: 8", "hint":"Track the largest and second largest distinct values.", "solution":"def second_largest(nums):\n    first = second = None\n    for x in nums:\n        if x == first: continue\n        if first is None or x > first:\n            second, first = first, x\n        elif second is None or x > second:\n            second = x\n    return second"},
    {"title":"Frequency Counter", "difficulty":"Easy", "topic":"Hashing", "prompt":"Return a dictionary containing the frequency of each value in a list.", "sample":"Input: [1,2,2,3,1,1]\nOutput: {1:3, 2:2, 3:1}", "hint":"Use a dictionary and increment the count.", "solution":"def frequency(nums):\n    freq = {}\n    for x in nums:\n        freq[x] = freq.get(x, 0) + 1\n    return freq"},
    {"title":"Reverse Words", "difficulty":"Easy", "topic":"Strings", "prompt":"Reverse the order of words in a sentence.", "sample":"Input: 'I love Python'\nOutput: 'Python love I'", "hint":"Split into words, reverse, then join.", "solution":"def reverse_words(s):\n    return ' '.join(s.split()[::-1])"},
]

HR_QUESTIONS = [
    "Tell me about yourself.", "Why should we hire you?", "Why do you want to join this company?",
    "Why are you interested in this role?", "What are your strengths?", "What is one weakness you are improving?",
    "Tell me about a time you solved a difficult problem.", "Tell me about a conflict you handled in a team.",
    "Tell me about a failure and what you learned from it.", "Where do you see yourself in three to five years?",
    "Are you comfortable relocating?", "Are you comfortable working in shifts when required?",
    "What motivates you at work?", "How do you prioritize multiple deadlines?", "What do you know about our industry?",
    "What are your salary expectations?", "What is your biggest professional achievement?", "Do you have any questions for us?",
]

PROJECT_QUESTIONS = [
    "Explain your project in simple terms.", "What real-world problem does your project solve?", "Why did you choose this problem?",
    "Why did you choose the technologies used?", "Explain the architecture or workflow.", "What was your individual contribution?",
    "What was the biggest technical challenge?", "How did you test the project?", "What would you improve in version 2?",
    "How would you deploy this project for real users?", "What happens if the input is invalid or unexpected?", "How did you evaluate the result or accuracy?",
]

TECH_BANK = {
    "Python": [
        ("What is the difference between a list and a tuple?", "A list is mutable; a tuple is immutable."),
        ("What is a dictionary?", "A key-value data structure used for fast lookup by key."),
        ("What is list comprehension?", "A compact way to construct lists from iterables, optionally with a condition."),
        ("Why use try-except?", "To handle exceptions gracefully instead of letting the program crash."),
        ("What is a function?", "A reusable block of code that can accept inputs and return a result."),
    ],
    "SQL / DBMS": [
        ("What is a primary key?", "A column or column set that uniquely identifies a row."),
        ("WHERE vs HAVING?", "WHERE filters rows before grouping; HAVING filters grouped results."),
        ("What is a JOIN?", "It combines related rows from two or more tables."),
        ("What is normalization?", "Organizing data to reduce redundancy and improve consistency."),
        ("What does GROUP BY do?", "It groups rows so aggregates like COUNT or AVG can be calculated per group."),
    ],
    "AI / Machine Learning": [
        ("What is supervised learning?", "Learning from labeled examples with known target outputs."),
        ("What is overfitting?", "When a model fits training data too closely and generalizes poorly."),
        ("Classification vs regression?", "Classification predicts discrete classes; regression predicts continuous values."),
        ("Why split train and test data?", "To estimate performance on unseen data."),
        ("What is precision?", "True positives divided by all predicted positives."),
        ("Why scale features?", "To bring numerical features to comparable ranges when an algorithm benefits from scaling."),
    ],
    "Deep Learning / NLP / CV": [
        ("What is a neural network?", "A layered parameterized model that learns patterns from data."),
        ("Why is ReLU common?", "It is simple and helps reduce some vanishing-gradient issues."),
        ("What is tokenization?", "Splitting text into tokens such as words or subwords."),
        ("What is image classification?", "Assigning an image to one or more classes."),
        ("Why use a validation set?", "To tune choices without using the final test set."),
    ],
    "CS Fundamentals": [
        ("What is OOP?", "A programming style organized around objects containing data and behavior."),
        ("What is a process?", "A program in execution with its own managed resources."),
        ("What is an operating system?", "Software that manages hardware and provides services to applications."),
        ("What is an API?", "A defined interface through which software components communicate."),
        ("What is Git?", "A version-control system for tracking and collaborating on source-code changes."),
    ],
}


def extract_file_text_quiet(uploaded_file):
    """Extract PDF, DOCX or TXT without creating duplicate UI errors."""
    if uploaded_file is None:
        return ""
    try:
        data = uploaded_file.getvalue()
        name = uploaded_file.name.lower()
        if name.endswith(".pdf"):
            doc = pymupdf.open(stream=data, filetype="pdf")
            text = "\n".join(page.get_text() for page in doc)
            doc.close()
            return text
        if name.endswith(".docx"):
            doc = Document(io.BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs)
        if name.endswith(".txt"):
            return data.decode("utf-8", errors="ignore")
    except Exception:
        return ""
    return ""


def validate_jd_document(text):
    normalized = re.sub(r"\s+", " ", text.lower()).strip()
    if len(normalized) < 180:
        return False, "The document is too short to be a reliable job description.", 10
    jd_sections = [
        "job description", "job role", "role", "responsibilities", "requirements", "qualifications",
        "skills", "experience", "about the role", "what you will do", "must have", "preferred"
    ]
    business_doc_words = ["invoice", "receipt", "gstin", "subtotal", "amount due", "transaction id", "billing address"]
    section_hits = sum(1 for x in jd_sections if x in normalized)
    skill_hits = len(extract_skills(text))
    role_hints = sum(1 for x in ["developer", "engineer", "analyst", "intern", "software", "data", "machine learning", "designer", "manager", "associate"] if x in normalized)
    bad_hits = sum(1 for x in business_doc_words if x in normalized)
    if bad_hits >= 2 and section_hits < 3:
        return False, "This appears to be an invoice, bill, receipt, or unrelated business document.", 5
    if section_hits >= 3 and (skill_hits >= 2 or role_hints >= 1):
        return True, "Job-description structure detected.", min(95, 65 + section_hits * 5)
    if section_hits >= 2 and role_hints >= 1:
        return True, "JD-like structure detected.", 75
    return False, "Please upload a real job description containing role, responsibilities, requirements, qualifications, or skills.", 20


def extract_jd_insights(jd):
    skills = extract_skills(jd)
    text = jd.lower()
    responsibilities = []
    for line in jd.splitlines():
        clean = line.strip(" •-\t")
        low = clean.lower()
        if len(clean) > 20 and any(x in low for x in ["develop", "build", "design", "analyze", "test", "maintain", "deploy", "support", "create", "manage"]):
            responsibilities.append(clean)
    role_match = re.search(r"(?:role|position|job title)\s*[:\-]\s*([^\n]+)", jd, re.I)
    role = role_match.group(1).strip() if role_match else "Target Role"
    experience = re.findall(r"\b\d+\+?\s*(?:years?|yrs?)\b", text)
    return {"role": role, "skills": skills, "responsibilities": responsibilities[:8], "experience": experience}


def coding_review(code, problem):
    if not code.strip():
        return 0, ["Write a solution before checking it."]
    feedback = []
    try:
        ast.parse(code)
        syntax = 40
        feedback.append("✅ Python syntax looks valid.")
    except SyntaxError as e:
        return 10, [f"❌ Syntax error near line {e.lineno}: {e.msg}"]
    lowered = code.lower()
    topic = problem["topic"].lower()
    if topic in lowered:
        feedback.append("✅ Your code reflects the target topic.")
    else:
        feedback.append("ℹ️ Make sure your solution uses a clear approach for the target topic.")
    if "for " in lowered or "while " in lowered:
        feedback.append("✅ A loop-based solution pattern is present.")
    if "def " in lowered:
        feedback.append("✅ You used a function, which is a good interview practice.")
    else:
        feedback.append("⚠️ Prefer wrapping your solution in a function for coding interviews.")
    if "return " in lowered:
        feedback.append("✅ The solution returns a result.")
    if len(code.splitlines()) <= 4:
        feedback.append("ℹ️ Very short solutions are fine when correct, but explain your logic clearly in the interview.")
    quality = syntax
    quality += 25 if "def " in lowered else 0
    quality += 20 if "return " in lowered else 0
    quality += 15 if "for " in lowered or "while " in lowered else 0
    return min(100, quality), feedback


def answer_score_against_reference(answer, reference):
    if not answer.strip():
        return 0, "No answer submitted."
    aw = set(re.findall(r"\b[a-zA-Z]{3,}\b", answer.lower()))
    rw = set(re.findall(r"\b[a-zA-Z]{3,}\b", reference.lower()))
    overlap = len(aw & rw) / max(1, len(rw))
    structure_bonus = 15 if any(x in answer.lower() for x in ["because", "for example", "therefore", "first", "then", "result"]) else 0
    score = min(100, round(overlap * 85 + structure_bonus))
    feedback = "Strong answer. Add a concrete example if you can." if score >= 75 else "Good start. Include the key concept and one practical example." if score >= 50 else "Revise the core concept and answer in your own words."
    return score, feedback


def personalized_project_questions(resume_text):
    qs = list(PROJECT_QUESTIONS)
    low = resume_text.lower()
    if "machine learning" in low or "xgboost" in low or "random forest" in low:
        qs.insert(2, "Why did you choose your machine-learning model, and how did you evaluate it?")
    if "flask" in low or "streamlit" in low:
        qs.insert(3, "Explain how the frontend/UI communicates with your backend or model.")
    if "deep learning" in low:
        qs.insert(4, "How did you prepare data and avoid overfitting in your deep-learning project?")
    return list(dict.fromkeys(qs))[:14]


def get_module_average(df, module):
    if df.empty:
        return 0
    rows = df[df["module"] == module]
    return round(float(rows["score"].mean()), 1) if not rows.empty else 0


def get_readiness_scores():
    attempts = get_attempts()
    vals = {
        "Resume / Role Fit": float(st.session_state.get("match_score", 0)),
        "Aptitude": get_module_average(attempts, "Aptitude"),
        "Coding": get_module_average(attempts, "Coding"),
        "Technical": get_module_average(attempts, "Technical"),
        "Project": get_module_average(attempts, "Project"),
        "HR": get_module_average(attempts, "HR"),
        "Mock Interview": get_module_average(attempts, "Mock Interview"),
    }
    weights = {"Resume / Role Fit":.20,"Aptitude":.12,"Coding":.15,"Technical":.18,"Project":.10,"HR":.10,"Mock Interview":.15}
    readiness = round(sum(vals[k] * weights[k] for k in vals), 1)
    return vals, readiness


def store_and_message(module, topic, score):
    save_attempt(module, topic, score)
    st.success(f"Saved to your InterviewMate progress: {score}%")


# -------------------------- SESSION --------------------------
st.session_state.setdefault("authenticated", False)
st.session_state.setdefault("user_identifier", "")
for key, default in {
    "company_name":"", "job_role":"", "resume_text":"", "jd_text":"", "match_score":0,
    "matched_skills":[], "missing_skills":[], "tailored_resume":"", "tailor_suggestions":[],
    "aptitude_quiz":[], "aptitude_result":None, "coding_index":0, "coding_review":None,
    "technical_result":None, "project_result":None, "hr_result":None,
    "mock_active":False, "mock_round":"HR", "mock_questions":[], "mock_index":0, "mock_scores":[],
}.items():
    st.session_state.setdefault(key, default)

# --------------------------- SIDEBAR ---------------------------
with st.sidebar:
    st.markdown(
        '<div style="padding:8px 4px 16px"><div style="font-size:11px;letter-spacing:.18em;font-weight:800;color:#a8a6ff">INTERVIEWMATE</div><div style="font-family:Manrope;font-size:22px;font-weight:800;margin-top:6px">Placement Command Center</div><div style="font-size:12px;color:#9aa4b5;margin-top:6px;line-height:1.5">Resume intelligence + assessment practice + interview simulation + readiness tracking.</div></div>',
        unsafe_allow_html=True,
    )
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.user_identifier = ""
        st.rerun()
    st.caption(f"Signed in as: {st.session_state.user_identifier}")
    st.divider()
    page = st.radio("WORKSPACE", [
        "🏠 Placement Dashboard", "📄 Resume Intelligence", "🎯 Company & Job Fit", "✨ Resume Tailor",
        "🧮 Aptitude Arena", "💻 Coding Lab", "🧠 Technical Round", "🛠️ Project Round", "🤝 HR Round",
        "🎤 Mock Interview", "🗺️ Prep Roadmap", "📊 Readiness & Progress"
    ])
    st.divider()
    st.markdown('<div style="font-size:11px;color:#8791a3;line-height:1.7">LOCAL STUDENT WORKSPACE<br>Private local analysis • SQLite history<br><br>Python · Streamlit · NLP · Assessments</div>', unsafe_allow_html=True)

# ------------------------- DASHBOARD -------------------------
if page == "🏠 Placement Dashboard":
    vals, readiness = get_readiness_scores()
    hero("01 · PLACEMENT COMMAND CENTER", "Your complete placement-preparation workspace.", "Prepare for the full journey: resume screening, aptitude, coding, technical interviews, project defense, HR, and mock interviews — while tracking the exact areas that need more practice.")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Placement Readiness", f"{readiness}%")
    c2.metric("Role Match", f"{st.session_state.match_score:.0f}%")
    attempts = get_attempts()
    c3.metric("Practice Attempts", len(attempts))
    c4.metric("Modules", "12")
    st.markdown('<div class="section-title">Your preparation pipeline</div>', unsafe_allow_html=True)
    st.markdown('''<div class="workflow"><div class="workflow-item"><div class="workflow-no">1</div><b>Resume</b><span>Validate and understand your profile.</span></div><div class="workflow-item"><div class="workflow-no">2</div><b>Role Fit</b><span>Match skills to company/JD requirements.</span></div><div class="workflow-item"><div class="workflow-no">3</div><b>Aptitude</b><span>Quant, reasoning and verbal practice.</span></div><div class="workflow-item"><div class="workflow-no">4</div><b>Coding</b><span>Beginner-friendly placement problems.</span></div><div class="workflow-item"><div class="workflow-no">5</div><b>Technical</b><span>Python, SQL, ML, CS fundamentals.</span></div><div class="workflow-item"><div class="workflow-no">6</div><b>Interview</b><span>Project + HR + mock interview.</span></div><div class="workflow-item"><div class="workflow-no">7</div><b>Ready</b><span>See your readiness score and weak areas.</span></div></div>''', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Round-by-round readiness</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    for i,(name,score) in enumerate(vals.items()):
        with cols[i%4]:
            st.markdown(f'<div class="card"><div class="mini">{name}</div><div class="big">{score:.0f}%</div>{score_badge(score)}<div class="desc" style="margin-top:8px">Practice more to raise this area.</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">What should you do next?</div>', unsafe_allow_html=True)
    sorted_weak = sorted(vals.items(), key=lambda x:x[1])
    for name,score in sorted_weak[:3]:
        if score < 80:
            st.warning(f"**Priority:** {name} is at {score:.0f}%. Use that module today.")
    if not st.session_state.resume_text:
        st.info("Start with **Resume Intelligence** and upload your resume. The rest of the workspace will become more personalized after that.")
    elif not st.session_state.jd_text:
        st.info("Next: open **Company & Job Fit**, paste or upload a target job description, and build your skill-gap plan.")

# ----------------------- RESUME INTELLIGENCE -----------------------
elif page == "📄 Resume Intelligence":
    hero("02 · RESUME INTELLIGENCE", "Build your placement profile.", "Upload a real resume, reject accidental non-resume documents, extract skills, and prepare your profile for company-specific matching.")
    uploaded = st.file_uploader("Upload Resume / CV", type=["pdf","docx","txt"], key="resume_upload")
    if uploaded:
        text = extract_file_text_quiet(uploaded)
        valid, reason, confidence = validate_resume_document(text)
        if valid:
            st.session_state.resume_text = text
            skills = extract_skills(text)
            c1,c2,c3 = st.columns(3)
            c1.metric("Resume Confidence", f"{confidence}%")
            c2.metric("Detected Skills", len(skills))
            c3.metric("Words", len(re.findall(r"\b\w+\b", text)))
            st.success("✅ Valid resume detected and saved for the rest of InterviewMate.")
            st.markdown("**Detected skills**")
            st.markdown(pills(skills), unsafe_allow_html=True)
        else:
            st.error("❌ Wrong / invalid file uploaded")
            st.warning("Please upload your actual resume/CV and try again.")
            st.info(reason)
    if st.session_state.resume_text:
        st.subheader("Resume Profile")
        st.text_area("Extracted resume text", value=st.session_state.resume_text, height=330, key="resume_view")
        with st.expander("Resume quality checklist"):
            checks = [
                ("Contact details", bool(re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+|linkedin\.com|github\.com", st.session_state.resume_text, re.I))),
                ("Education", "education" in st.session_state.resume_text.lower()),
                ("Skills", "skills" in st.session_state.resume_text.lower()),
                ("Projects", "project" in st.session_state.resume_text.lower()),
                ("Experience / Internship", any(x in st.session_state.resume_text.lower() for x in ["experience","internship"])),
            ]
            for label,ok in checks:
                (st.success if ok else st.warning)(f"{'✅' if ok else '⚠️'} {label}")

# ------------------------ COMPANY & JOB FIT ------------------------
elif page == "🎯 Company & Job Fit":
    hero("03 · COMPANY + ROLE FIT", "Turn a job description into a preparation plan.", "Enter the exact company and role, then paste or upload the JD. InterviewMate validates the document, detects requirements, maps your skill gaps and identifies what to study next.")
    c1,c2 = st.columns(2)
    st.session_state.company_name = c1.text_input("🏢 Company name", value=st.session_state.company_name, placeholder="Example: TCS")
    st.session_state.job_role = c2.text_input("💼 Exact job role", value=st.session_state.job_role, placeholder="Example: AI/ML Engineer")
    jd_file = st.file_uploader("Upload Company JD / Requirements", type=["pdf","docx","txt"], key="jd_upload")
    if jd_file:
        jd_file_text = extract_file_text_quiet(jd_file)
        ok, reason, confidence = validate_jd_document(jd_file_text)
        if ok:
            st.session_state.jd_text = jd_file_text
            st.success(f"✅ Valid job description detected ({confidence}% confidence).")
        else:
            st.error("❌ Wrong / invalid JD uploaded")
            st.warning("Please upload the actual job description / requirements and try again.")
            st.info(reason)
    jd = st.text_area("Company job description / requirements", value=st.session_state.jd_text, height=260)
    if st.button("🚀 Analyze Company + Job Role", type="primary", use_container_width=True):
        if not st.session_state.resume_text:
            st.error("Upload and validate your resume first in Resume Intelligence.")
        elif not jd.strip():
            st.error("Enter or upload a job description.")
        else:
            ok, reason, confidence = validate_jd_document(jd)
            if not ok:
                st.error("❌ This text does not look like a valid job description.")
                st.info(reason)
            else:
                st.session_state.jd_text = jd
                final, similarity, skill_score, matched, missing = calculate_match_score(st.session_state.resume_text, jd)
                st.session_state.match_score = final
                st.session_state.matched_skills = matched
                st.session_state.missing_skills = missing
                insight = extract_jd_insights(jd)
                c1,c2,c3,c4 = st.columns(4)
                c1.metric("Overall Match", f"{final}%")
                c2.metric("Text Similarity", f"{similarity}%")
                c3.metric("Skill Match", f"{skill_score}%")
                c4.metric("JD Confidence", f"{confidence}%")
                st.subheader("Role intelligence")
                st.write(f"**Target:** {st.session_state.company_name or 'Company not named'} · {st.session_state.job_role or insight['role']}")
                a,b = st.columns(2)
                with a:
                    st.write("**✅ Skills you already match**")
                    st.markdown(pills(matched), unsafe_allow_html=True)
                with b:
                    st.write("**⚠️ Skills to prepare**")
                    st.markdown(pills(missing), unsafe_allow_html=True)
                if insight["responsibilities"]:
                    st.subheader("What the company is likely hiring you to do")
                    for item in insight["responsibilities"]:
                        st.write("• " + item)
                if missing:
                    st.warning("Study gap detected: " + ", ".join(x.title() for x in missing[:10]) + ".")
                else:
                    st.success("No missing predefined skills were detected. Focus on interview performance and projects.")

# --------------------------- RESUME TAILOR ---------------------------
elif page == "✨ Resume Tailor":
    hero("04 · ATS + ROLE TAILOR", "Modify your resume for the exact company and job role.", "InterviewMate reorders relevant content and emphasizes only skills already supported by your resume. It does not invent experience.")
    if not st.session_state.resume_text:
        st.warning("Upload a resume first.")
    else:
        c1,c2 = st.columns(2)
        st.session_state.company_name = c1.text_input("Company", value=st.session_state.company_name, key="tailor_company")
        st.session_state.job_role = c2.text_input("Exact role", value=st.session_state.job_role, key="tailor_role")
        resume = st.text_area("Resume", value=st.session_state.resume_text, height=220, key="tailor_resume")
        jd = st.text_area("Company JD", value=st.session_state.jd_text, height=220, key="tailor_jd")
        if st.button("✨ Generate Tailored Resume", type="primary"):
            rv, rr, rc = validate_resume_document(resume)
            jv, jr, jc = validate_jd_document(jd)
            if not rv:
                st.error("❌ Invalid resume. Please upload the correct CV.")
                st.info(rr)
            elif not jv:
                st.error("❌ Invalid job description. Please upload/paste the correct JD.")
                st.info(jr)
            else:
                final, similarity, skill_score, matched, missing = calculate_match_score(resume, jd)
                st.session_state.resume_text, st.session_state.jd_text = resume, jd
                st.session_state.match_score = final
                st.session_state.matched_skills, st.session_state.missing_skills = matched, missing
                st.session_state.tailored_resume = make_tailored_resume_text(resume, jd, matched, missing)
                target = f"TARGET: {st.session_state.company_name or 'Target Company'} | ROLE: {st.session_state.job_role or 'Target Role'}"
                st.session_state.tailored_resume = target + "\n\n" + st.session_state.tailored_resume
                st.session_state.tailor_suggestions = get_tailoring_suggestions(resume, jd, matched, missing)
                st.success("✅ Tailored resume generated.")
        if st.session_state.tailored_resume:
            c1,c2,c3 = st.columns(3)
            c1.metric("Role Match", f"{st.session_state.match_score:.0f}%")
            c2.metric("Emphasized Skills", len(st.session_state.matched_skills))
            c3.metric("Skill Gaps", len(st.session_state.missing_skills))
            st.subheader("Smart suggestions")
            for suggestion in st.session_state.tailor_suggestions:
                st.info(suggestion)
            st.subheader("Tailored resume preview")
            st.text_area("Generated resume", st.session_state.tailored_resume, height=520)
            d1,d2 = st.columns(2)
            with d1:
                st.download_button("⬇️ Download DOCX", create_docx_bytes(st.session_state.tailored_resume), "InterviewMate_Tailored_Resume.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            with d2:
                st.download_button("⬇️ Download PDF", create_pdf_bytes(st.session_state.tailored_resume), "InterviewMate_Tailored_Resume.pdf", "application/pdf", use_container_width=True)

# --------------------------- APTITUDE ---------------------------
elif page == "🧮 Aptitude Arena":
    hero("05 · APTITUDE ARENA", "Train for the first screening round.", "Practice quantitative aptitude, logical reasoning and verbal ability in placement-style MCQs. Every completed test is saved to your progress.")
    topic = st.selectbox("Choose section", list(APTITUDE_BANK.keys()))
    if not st.session_state.aptitude_quiz or st.session_state.get("aptitude_topic") != topic:
        st.session_state.aptitude_topic = topic
        st.session_state.aptitude_quiz = APTITUDE_BANK[topic][:]
        st.session_state.aptitude_result = None
    questions = st.session_state.aptitude_quiz
    answers = []
    for i,item in enumerate(questions):
        answers.append(st.radio(f"Q{i+1}. {item['q']}", item["o"], key=f"apt_{topic}_{i}", index=None))
    if st.button("✅ Submit Aptitude Test", type="primary"):
        score = 0
        for i,item in enumerate(questions):
            if answers[i] == item["o"][item["a"]]:
                score += 1
        pct = round(score/len(questions)*100,1)
        st.session_state.aptitude_result = pct
        save_attempt("Aptitude", topic, pct)
    if st.session_state.aptitude_result is not None:
        pct = st.session_state.aptitude_result
        st.divider(); st.metric("Your Score", f"{pct}%")
        if pct < 60: st.warning("Focus on accuracy first. Learn formulas/patterns, then retake.")
        elif pct < 80: st.info("Good attempt. More timed practice will improve speed.")
        else: st.success("Strong aptitude performance.")
        with st.expander("View solutions"):
            for i,item in enumerate(questions):
                st.write(f"**Q{i+1}. {item['e']}**")

# ---------------------------- CODING ----------------------------
elif page == "💻 Coding Lab":
    hero("06 · CODING LAB", "Build coding confidence from beginner to interview level.", "Start with simple Python placement problems, write your own solution, check syntax, review the approach and reveal a reference answer only when you need help.")
    idx = st.number_input("Problem", min_value=1, max_value=len(CODING_PROBLEMS), value=st.session_state.coding_index+1, step=1) - 1
    st.session_state.coding_index = idx
    p = CODING_PROBLEMS[idx]
    c1,c2,c3 = st.columns(3)
    c1.metric("Difficulty", p["difficulty"]); c2.metric("Topic", p["topic"]); c3.metric("Problem", p["title"])
    st.markdown(f"### {p['title']}")
    st.write(p["prompt"])
    st.code(p["sample"], language="text")
    with st.expander("💡 Show hint"):
        st.info(p["hint"])
    code = st.text_area("Write your Python solution", height=300, placeholder="def solution(...):\n    # write your logic here\n    return ...", key=f"code_{idx}")
    if st.button("🔍 Review My Solution", type="primary"):
        score, feedback = coding_review(code, p)
        st.session_state.coding_review = {"score":score,"feedback":feedback,"idx":idx}
        save_attempt("Coding", p["topic"], score)
    if st.session_state.coding_review and st.session_state.coding_review.get("idx") == idx:
        st.metric("Code Review Score", f"{st.session_state.coding_review['score']}%")
        for item in st.session_state.coding_review["feedback"]:
            st.write(item)
        st.caption("InterviewMate's local coding review checks Python syntax and common solution patterns; it does not execute untrusted user code.")
        if st.checkbox("Reveal reference solution"):
            st.code(p["solution"], language="python")

# ------------------------- TECHNICAL ROUND -------------------------
elif page == "🧠 Technical Round":
    hero("07 · TECHNICAL ROUND", "Practice the questions interviewers actually probe.", "Choose a technical track. Answer in your own words, compare against the expected core points, and build a strong technical vocabulary for interviews.")
    topic = st.selectbox("Technical track", list(TECH_BANK.keys()))
    data = TECH_BANK[topic]
    if "technical_answers" not in st.session_state or st.session_state.get("technical_topic") != topic:
        st.session_state.technical_answers = [""] * len(data)
        st.session_state.technical_topic = topic
    for i,(q,ref) in enumerate(data):
        with st.expander(f"Q{i+1}. {q}", expanded=(i==0)):
            st.session_state.technical_answers[i] = st.text_area("Your answer", value=st.session_state.technical_answers[i], key=f"tech_{topic}_{i}", height=120)
            if st.checkbox("Show expected core points", key=f"showtech_{topic}_{i}"):
                st.info(ref)
    if st.button("🧠 Score Technical Session", type="primary"):
        scores=[]
        for answer,(q,ref) in zip(st.session_state.technical_answers,data):
            s,_ = answer_score_against_reference(answer,ref); scores.append(s)
        pct = round(np.mean(scores),1) if scores else 0
        st.session_state.technical_result = pct
        save_attempt("Technical", topic, pct)
    if st.session_state.technical_result is not None:
        st.metric("Technical Session Score", f"{st.session_state.technical_result}%")
        st.progress(min(st.session_state.technical_result/100,1))

# --------------------------- PROJECT ROUND ---------------------------
elif page == "🛠️ Project Round":
    hero("08 · PROJECT DEFENSE", "Turn your project into an interview advantage.", "Interviewers often test whether you truly understand the project on your resume. Practice explaining problem, architecture, contribution, challenges, testing and future scope.")
    if not st.session_state.resume_text:
        st.info("Upload your resume first so project questions can be personalized.")
    project_hint = st.text_input("Project name (optional)", placeholder="Example: InterviewMate / AURA / Credit Card Approval Prediction")
    questions = personalized_project_questions(st.session_state.resume_text)
    answers=[]
    for i,q in enumerate(questions[:10]):
        ref = "Explain the project clearly, state your own contribution, mention a technical decision, and give a concrete result or learning." if i != 0 else "Start with problem → users → solution → main technologies → result."
        answers.append(st.text_area(f"Q{i+1}. {q}", key=f"proj_{i}", height=100))
    if st.button("🛠️ Score Project Defense", type="primary"):
        scores=[]
        for a in answers:
            s,_=answer_score_against_reference(a,ref); scores.append(s)
        pct=round(np.mean(scores),1) if scores else 0
        st.session_state.project_result=pct
        save_attempt("Project", project_hint or "Project Defense", pct)
    if st.session_state.project_result is not None:
        st.metric("Project Defense Score", f"{st.session_state.project_result}%")
        st.info("Best practice: be able to explain your project to a non-technical interviewer in 60 seconds and to a technical interviewer in 5 minutes.")

# ------------------------------ HR --------------------------------
elif page == "🤝 HR Round":
    hero("09 · HR ROUND", "Practice confident, professional answers.", "HR preparation covers self-introduction, strengths, weaknesses, teamwork, conflict, motivation, relocation, salary and closing questions.")
    selected = st.multiselect("Choose HR questions", HR_QUESTIONS, default=HR_QUESTIONS[:5])
    if not selected:
        st.info("Choose at least one question.")
    else:
        answers=[]
        for i,q in enumerate(selected):
            answers.append(st.text_area(f"Q{i+1}. {q}", key=f"hr_{i}_{q}", height=110))
        if st.button("🤝 Score HR Session", type="primary"):
            scores=[]
            generic_refs={
                "Tell me about yourself.":"Present your education, skills, projects, strengths and target role in a concise professional story.",
                "Why should we hire you?":"Connect your skills, projects, learning ability and fit for the role with evidence.",
                "What is one weakness you are improving?":"Name a genuine manageable weakness and explain the concrete action you are taking to improve it.",
            }
            for q,a in zip(selected,answers):
                ref=generic_refs.get(q,"Answer clearly, use a specific example where possible, and connect your response to professional behavior or role fit.")
                s,_=answer_score_against_reference(a,ref); scores.append(s)
            pct=round(np.mean(scores),1) if scores else 0
            st.session_state.hr_result=pct
            save_attempt("HR", "HR Session", pct)
        if st.session_state.hr_result is not None:
            st.metric("HR Session Score", f"{st.session_state.hr_result}%")
            st.info("Use STAR for experience-based answers: Situation → Task → Action → Result.")

# ------------------------- MOCK INTERVIEW -------------------------
elif page == "🎤 Mock Interview":
    hero("10 · MOCK INTERVIEW", "Simulate a real interview round.", "Choose the round, answer one question at a time, and let InterviewMate score relevance, specificity and structure. The goal is to make the screen feel like a real interview instead of a question list.")
    if not st.session_state.mock_active:
        round_type = st.selectbox("Mock round", ["HR","Technical","Project"])
        count = st.slider("Number of questions", 3, 8, 5)
        if st.button("▶ Start Mock Interview", type="primary", use_container_width=True):
            if round_type == "HR":
                pool = HR_QUESTIONS[:]
            elif round_type == "Project":
                pool = personalized_project_questions(st.session_state.resume_text)
            else:
                pool = [q for bank in TECH_BANK.values() for q,_ in bank]
            st.session_state.mock_round = round_type
            st.session_state.mock_questions = pool[:count]
            st.session_state.mock_index = 0
            st.session_state.mock_scores = []
            st.session_state.mock_active = True
            st.rerun()
    else:
        q = st.session_state.mock_questions[st.session_state.mock_index]
        total=len(st.session_state.mock_questions)
        st.progress(st.session_state.mock_index/total)
        st.caption(f"Question {st.session_state.mock_index+1} of {total} · {st.session_state.mock_round} Round")
        st.markdown(f"### {q}")
        answer=st.text_area("Your answer", key=f"mock_answer_{st.session_state.mock_index}", height=220)
        if st.button("Submit Answer", type="primary"):
            score = analyze_answer(q,answer)["score"]
            st.session_state.mock_scores.append(score)
            save_attempt("Mock Interview", st.session_state.mock_round, score)
            if st.session_state.mock_index + 1 < total:
                st.session_state.mock_index += 1
                st.rerun()
            else:
                st.session_state.mock_active = False
                st.success(f"Mock interview complete. Average score: {np.mean(st.session_state.mock_scores):.1f}%")
                st.session_state.mock_index=0
        if st.button("⏹ End Mock Interview"):
            st.session_state.mock_active=False
            st.session_state.mock_index=0
            st.rerun()

# --------------------------- ROADMAP ---------------------------
elif page == "🗺️ Prep Roadmap":
    vals, readiness = get_readiness_scores()
    hero("11 · PERSONALIZED ROADMAP", "Convert your gaps into a study plan.", "InterviewMate turns your resume-role gap and round scores into an ordered practice plan so you know what to study instead of randomly preparing.")
    ranked=sorted(vals.items(), key=lambda x:x[1])
    st.subheader("Priority map")
    for name,score in ranked:
        st.progress(min(score/100,1), text=f"{name}: {score:.0f}%")
    st.subheader("7-day placement sprint")
    plans=[
        ("Day 1","Resume + Role Fit","Upload resume, validate target JD, tailor resume, learn missing skills."),
        ("Day 2","Quant + Reasoning","Practice 20 aptitude questions and review wrong answers."),
        ("Day 3","Coding","Solve 3 Python array/string problems and explain your approach aloud."),
        ("Day 4","Technical","Revise Python, SQL/DBMS, ML and CS fundamentals based on your target role."),
        ("Day 5","Project","Prepare a 60-second project explanation plus architecture, challenge and contribution answers."),
        ("Day 6","HR","Practice self-introduction, strengths/weaknesses, teamwork, motivation and STAR stories."),
        ("Day 7","Mock Day","Take one full mock interview and repeat the weakest module."),
    ]
    for day,title,desc in plans:
        st.markdown(f'<div class="card" style="margin-bottom:10px"><div class="mini">{day}</div><div style="font-size:19px;font-weight:800;margin:5px 0">{title}</div><div class="desc">{desc}</div></div>', unsafe_allow_html=True)
    if st.session_state.missing_skills:
        st.subheader("Role-specific study gaps")
        st.markdown(pills(st.session_state.missing_skills), unsafe_allow_html=True)

# ----------------------- READINESS & PROGRESS -----------------------
elif page == "📊 Readiness & Progress":
    vals, readiness = get_readiness_scores()
    hero("12 · READINESS CENTER", "Know exactly where you stand before placement day.", "Your score combines profile fit and performance across aptitude, coding, technical, project, HR and mock interview practice. A real placement plan should improve weak areas, not just count questions.")
    c1,c2,c3=st.columns(3)
    c1.metric("Overall Readiness", f"{readiness}%")
    c2.metric("Role Match", f"{st.session_state.match_score:.0f}%")
    c3.metric("Practice Attempts", len(get_attempts()))
    names=list(vals.keys()); scores=list(vals.values())
    fig=go.Figure(go.Bar(x=scores,y=names,orientation="h"))
    fig.update_layout(title="Round readiness", xaxis_title="Score (%)", xaxis=dict(range=[0,100]), yaxis_title="", height=420)
    st.plotly_chart(fig,use_container_width=True)
    st.subheader("Priority weaknesses")
    weak=sorted(vals.items(), key=lambda x:x[1])
    for n,s in weak[:3]:
        st.warning(f"{n}: {s:.0f}%")
    attempts=get_attempts()
    if not attempts.empty:
        st.subheader("Your latest attempts")
        show=attempts[["module","topic","score","created_at"]].copy()
        st.dataframe(show,use_container_width=True,hide_index=True)
        by=attempts.groupby("module")["score"].mean().round(1).sort_values()
        st.subheader("Average by module")
        st.dataframe(by.rename("Average Score (%)"),use_container_width=True)
    if st.button("💾 Save current readiness snapshot"):
        save_session_summary(st.session_state.match_score, readiness)
        st.success("Readiness snapshot saved locally.")

st.divider()
st.markdown('<div class="footer">INTERVIEWMATE · COMPLETE PLACEMENT PREPARATION · RESUME · JD · APTITUDE · CODING · TECHNICAL · PROJECT · HR · MOCK INTERVIEW · READINESS</div>', unsafe_allow_html=True)

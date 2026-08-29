import streamlit as st
from pypdf import PdfReader
import docx
from google import genai
import json
import plotly.graph_objects as go
import io
import pandas as pd
import sqlite3
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(
    page_title="TalentIntel AI — ATS Intelligence & Talent Scorer",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 100% HIGH-CONTRAST SLATE-NAVY UI SYSTEM
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    * {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }
    
    /* Main Canvas Background */
    .stApp {
        background-color: #0B0F19 !important;
        color: #F8FAFC !important;
    }

    /* Fixed Dark Sidebar */
    [data-testid="stSidebar"] {
        background-color: #111827 !important;
        border-right: 1px solid #1F2937 !important;
    }
    [data-testid="stSidebar"] * {
        color: #F8FAFC !important;
    }

    /* All Inputs, TextAreas & Selectboxes - 100% Crisp Visible */
    .stTextInput input, .stTextArea textarea, [data-baseweb="select"] div {
        background-color: #1E293B !important;
        color: #FFFFFF !important;
        border: 1px solid #3B82F6 !important;
        font-size: 14px !important;
        border-radius: 8px !important;
    }
    .stTextArea textarea {
        color: #F1F5F9 !important;
        line-height: 1.5 !important;
    }

    /* Top Hero Header */
    .hero-banner {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.9) 0%, rgba(17, 24, 39, 0.95) 100%);
        border: 1px solid #374151;
        border-radius: 12px;
        padding: 22px;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -10px rgba(0, 0, 0, 0.6);
        text-align: center;
    }
    .hero-title {
        font-size: 30px;
        font-weight: 800;
        background: linear-gradient(90deg, #38BDF8 0%, #818CF8 50%, #C084FC 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 6px;
        letter-spacing: 0.5px;
    }
    .hero-desc {
        color: #94A3B8;
        font-size: 14px;
        margin: 0;
    }

    /* Metric Cards */
    .stat-box {
        background: #1E293B;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 12px 16px;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
    }
    .stat-number {
        font-size: 22px;
        font-weight: 800;
        color: #38BDF8;
    }
    .stat-label {
        font-size: 11px;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 600;
    }

    /* Tabs Styling - Solid, Crisp & Visible */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background: #111827;
        padding: 6px;
        border-radius: 10px;
        border: 1px solid #374151;
    }
    .stTabs [data-baseweb="tab"] {
        background: #1F2937 !important;
        border: 1px solid #374151 !important;
        border-radius: 8px !important;
        padding: 10px 22px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        color: #E2E8F0 !important;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #2563EB 0%, #4F46E5 100%) !important;
        color: #FFFFFF !important;
        border: 1px solid #60A5FA !important;
        box-shadow: 0 4px 15px rgba(37, 99, 235, 0.4) !important;
    }

    /* Cards & Badges */
    .section-card {
        background: #1E293B;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 16px;
    }
    .badge-tier1 {
        background: linear-gradient(135deg, #059669 0%, #10B981 100%);
        color: white;
        padding: 6px 18px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 13px;
        display: inline-block;
    }
    .badge-tier2 {
        background: linear-gradient(135deg, #D97706 0%, #F59E0B 100%);
        color: white;
        padding: 6px 18px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 13px;
        display: inline-block;
    }
    .badge-tier3 {
        background: linear-gradient(135deg, #DC2626 0%, #EF4444 100%);
        color: white;
        padding: 6px 18px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 13px;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# DATABASE SETUP (SQLite)
# ==========================================
DB_FILE = "talent_intelligence.db"

def get_db_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def init_database():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS candidates (
            candidate_id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidate_name TEXT UNIQUE,
            created_at TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS evaluations (
            eval_id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidate_name TEXT,
            target_role TEXT,
            match_score INTEGER,
            technical_score INTEGER,
            experience_score INTEGER,
            tooling_score INTEGER,
            soft_skill_score INTEGER,
            ats_readability INTEGER,
            priority_tier TEXT,
            status TEXT DEFAULT 'Under Review',
            evaluated_at TEXT,
            summary TEXT,
            raw_json TEXT
        )
    """)
    conn.commit()
    conn.close()

init_database()

def save_full_evaluation(candidate_name, target_role, result):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR IGNORE INTO candidates (candidate_name, created_at) VALUES (?, ?)",
                   (candidate_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    
    metrics = result.get("competency_scores", {})
    cursor.execute("""
        INSERT INTO evaluations (
            candidate_name, target_role, match_score, technical_score,
            experience_score, tooling_score, soft_skill_score, ats_readability,
            priority_tier, evaluated_at, summary, raw_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        candidate_name,
        target_role,
        result.get("match_score", 0),
        metrics.get("technical_skills", 0),
        metrics.get("experience_depth", 0),
        metrics.get("tools_frameworks", 0),
        metrics.get("soft_skills", 0),
        metrics.get("ats_readability", 0),
        result.get("priority_tier", "Medium"),
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        result.get("priority_reason", ""),
        json.dumps(result)
    ))
    conn.commit()
    conn.close()

def fetch_evaluations_df():
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT eval_id, candidate_name, target_role, match_score, technical_score, priority_tier, status, evaluated_at FROM evaluations ORDER BY eval_id DESC", conn)
    conn.close()
    return df

def update_candidate_status(eval_id, new_status):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE evaluations SET status = ? WHERE eval_id = ?", (new_status, eval_id))
    conn.commit()
    conn.close()

SAMPLE_JDS = {
    "Python / Backend Developer": """Looking for a Backend Developer skilled in Python, FastAPI/Django/Flask, SQL, Redis, REST APIs, Git, Docker, and CI/CD pipelines. Experience in scalable architecture and unit testing is highly preferred.""",
    "Data Scientist & AI Specialist": """Seeking a Data Scientist proficient in Python, Machine Learning (Scikit-Learn), Deep Learning (PyTorch/TensorFlow), LLMs, LangChain, RAG architectures, SQL, and data visualization.""",
    "Cloud & DevOps Engineer": """Hiring a DevOps Engineer experienced in AWS/Azure, Docker, Kubernetes, Terraform, Linux scripting, Jenkins/GitHub Actions, and observability tools (Prometheus, Grafana).""",
    "Full Stack Web Developer": """Looking for a Full Stack Engineer with proficiency in React.js, Node.js/Express, TypeScript, PostgreSQL/MongoDB, REST/GraphQL APIs, and cloud deployments.""",
    "Custom Job Description": ""
}

def extract_text(file):
    if file.name.endswith(".pdf"):
        reader = PdfReader(file)
        return "\n".join([page.extract_text() or "" for page in reader.pages])
    elif file.name.endswith(".docx"):
        doc = docx.Document(file)
        return "\n".join([p.text for p in doc.paragraphs])
    return ""

def run_deep_ai_evaluation(api_key, resume_text, jd_text, role):
    client = genai.Client(api_key=api_key)
    prompt = f"""
You are a Principal ATS Architect and Technical Hiring Director.
Evaluate this resume against the target role and job description with rigorous enterprise standards.

Role: {role}
Job Description:
{jd_text}

Resume Content:
{resume_text}

Return ONLY valid JSON matching this schema:
{{
  "match_score": <number 0-100>,
  "priority_tier": "Tier-1 (Immediate Shortlist)" or "Tier-2 (Strong Potential)" or "Tier-3 (Upskill Required)",
  "priority_reason": "<1-2 sentence core justification>",
  "competency_scores": {{
    "technical_skills": <0-100>,
    "experience_depth": <0-100>,
    "tools_frameworks": <0-100>,
    "soft_skills": <0-100>,
    "ats_readability": <0-100>
  }},
  "strengths": ["list of exact matching skills found in resume"],
  "missing_skills": ["list of required skills missing or weak in resume"],
  "learning_roadmap": [
    {{"week": "Week 1-2", "focus": "Skill/Topic", "resource_type": "Course/Project", "task": "Actionable task"}},
    {{"week": "Week 3-4", "focus": "Skill/Topic", "resource_type": "Course/Project", "task": "Actionable task"}}
  ],
  "interview_prep_questions": [
    {{"topic": "Skill/Topic", "question": "Technical question to test candidate", "expected_answer_hint": "Key concept candidate should explain"}},
    {{"topic": "Scenario/System", "question": "Real-world engineering challenge", "expected_answer_hint": "Architectural approach"}}
  ],
  "optimized_resume_bullets": [
    "High-impact bullet point 1 using Action Verb + Task + Quantifiable Result",
    "High-impact bullet point 2 addressing a missing skill"
  ],
  "custom_cover_letter": "A compelling, 3-paragraph professional cover letter tailored for this exact job description highlighting the candidate's real matching strengths."
}}
"""
    response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=prompt,
        config={"response_mime_type": "application/json"}
    )
    return json.loads(response.text)

def render_radar_chart(scores_dict):
    categories = ['Technical', 'Experience', 'Tools', 'Soft Skills', 'ATS Readability']
    values = [
        scores_dict.get('technical_skills', 0),
        scores_dict.get('experience_depth', 0),
        scores_dict.get('tools_frameworks', 0),
        scores_dict.get('soft_skills', 0),
        scores_dict.get('ats_readability', 0)
    ]
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=values + [values[0]],
        theta=categories + [categories[0]],
        fill='toself',
        fillcolor='rgba(56, 189, 248, 0.25)',
        line=dict(color='#38BDF8', width=2.5),
        name='Competency Vector'
    ))
    fig.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 100], tickfont=dict(color="#94A3B8"), gridcolor="#334155"),
            angularaxis=dict(tickfont=dict(color="#F8FAFC", size=11), gridcolor="#334155")
        ),
        showlegend=False,
        height=290,
        margin=dict(l=35, r=35, t=30, b=30),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)'
    )
    return fig

def render_gauge_chart(score):
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=score,
        number={'font': {'color': '#F8FAFC', 'size': 40}},
        title={'text': "ATS Overall Match (%)", 'font': {'color': '#94A3B8', 'size': 14}},
        gauge={
            'axis': {'range': [0, 100], 'tickcolor': "#64748B"},
            'bar': {'color': "#2563EB", 'thickness': 0.28},
            'steps': [
                {'range': [0, 45], 'color': "rgba(239, 68, 68, 0.25)"},
                {'range': [45, 75], 'color': "rgba(245, 158, 11, 0.25)"},
                {'range': [75, 100], 'color': "rgba(16, 185, 129, 0.25)"}
            ]
        }
    ))
    fig.update_layout(height=280, margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor='rgba(0,0,0,0)')
    return fig

def generate_dossier_pdf(candidate_name, role, result):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=18, leading=22, textColor=colors.HexColor("#0F172A"))
    h2_style = ParagraphStyle('H2Style', parent=styles['Heading2'], fontSize=12, leading=16, textColor=colors.HexColor("#2563EB"), spaceBefore=10, spaceAfter=4)
    body = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=9, leading=13, textColor=colors.HexColor("#334155"))
    bullet = ParagraphStyle('BulletStyle', parent=body, leftIndent=12, spaceAfter=2)
    
    story = []
    story.append(Paragraph("TalentIntel AI — Candidate Dossier", title_style))
    story.append(Paragraph(f"<b>Candidate:</b> {candidate_name} &nbsp;|&nbsp; <b>Target Role:</b> {role} &nbsp;|&nbsp; <b>Date:</b> {datetime.now().strftime('%d %b %Y')}", body))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#CBD5E1"), spaceBefore=6, spaceAfter=10))
    
    metrics = result.get("competency_scores", {})
    summary_data = [
        ["Metric", "Score / Value", "Dimension", "Score"],
        ["ATS Match Score", f"{result.get('match_score', 0)}%", "Technical Alignment", f"{metrics.get('technical_skills', 0)}%"],
        ["Priority Tier", result.get('priority_tier', 'N/A'), "Experience Depth", f"{metrics.get('experience_depth', 0)}%"],
        ["ATS Readability", f"{metrics.get('ats_readability', 0)}%", "Tools & Frameworks", f"{metrics.get('tools_frameworks', 0)}%"]
    ]
    t = Table(summary_data, colWidths=[130, 130, 130, 130])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#EFF6FF")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor("#1E3A8A")),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))
    
    story.append(Paragraph("Verified Competencies & Strengths", h2_style))
    for s in result.get("strengths", []):
        story.append(Paragraph(f"• {s}", bullet))
        
    story.append(Paragraph("Identified Skill Deficits", h2_style))
    for m in result.get("missing_skills", []):
        story.append(Paragraph(f"• {m}", bullet))
        
    story.append(Paragraph("AI-Generated Technical Interview Kit", h2_style))
    for idx, q in enumerate(result.get("interview_prep_questions", []), 1):
        story.append(Paragraph(f"<b>Q{idx} ({q.get('topic')}):</b> {q.get('question')}", body))
        story.append(Paragraph(f"<i>Anchor:</i> {q.get('expected_answer_hint')}", bullet))
        story.append(Spacer(1, 4))
        
    doc.build(story)
    buffer.seek(0)
    return buffer

# Session States
if "evaluation_result" not in st.session_state:
    st.session_state.evaluation_result = None
if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = None
if "current_role" not in st.session_state:
    st.session_state.current_role = None

# Hero Header with New Clean Branding
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">⚡ TalentIntel AI</div>
    <div class="hero-desc">Enterprise ATS Intelligence, 5D Competency Vectors & Priority Talent Pipeline</div>
</div>
""", unsafe_allow_html=True)

# Top 4 Real-Time Live Metrics
eval_df_quick = fetch_evaluations_df()
q_count = len(eval_df_quick)
q_avg = f"{round(eval_df_quick['match_score'].mean(), 1)}%" if q_count > 0 else "0%"
q_top = f"{eval_df_quick['match_score'].max()}%" if q_count > 0 else "0%"

s1, s2, s3, s4 = st.columns(4)
with s1:
    st.markdown(f'<div class="stat-box"><div class="stat-number">{q_count}</div><div class="stat-label">Total Evaluated</div></div>', unsafe_allow_html=True)
with s2:
    st.markdown(f'<div class="stat-box"><div class="stat-number">{q_avg}</div><div class="stat-label">Avg ATS Match</div></div>', unsafe_allow_html=True)
with s3:
    st.markdown(f'<div class="stat-box"><div class="stat-number">{q_top}</div><div class="stat-label">Top ATS Score</div></div>', unsafe_allow_html=True)
with s4:
    st.markdown('<div class="stat-box"><div class="stat-number">PDF / DOCX</div><div class="stat-label">Parser Engine</div></div>', unsafe_allow_html=True)

st.write("")

# Sidebar
with st.sidebar:
    st.markdown("### ⚙️ Engine Settings")
    gemini_api_key = st.text_input("Gemini API Key", type="password", placeholder="Paste AIzaSy key...")
    st.markdown("""
    <div style="background: rgba(37, 99, 235, 0.15); border: 1px solid #3B82F6; padding: 8px; border-radius: 8px; font-size: 12px; color: #93C5FD; text-align: center; margin-top: 4px;">
        ⚡ <b>Active:</b> Gemini 3.6 Flash | SQLite3
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    st.markdown("### 🚀 Active Platform Modules")
    st.markdown("""
    * 🎯 **Vector Scoring Engine**
    * 🕸️ **5D Competency Radar**
    * 🛠️ **30-Day Learning Roadmap**
    * 🎤 **Interview Prep Kit**
    * 🏢 **Recruiter Screening Matrix**
    * 🗄️ **Relational Audit Trail**
    """)

tab1, tab2, tab3 = st.tabs([
    "👤 Candidate Mode (Deep Analysis)", 
    "🏢 Recruiter Mode (Batch Screening)", 
    "🗄️ Database & History Logs"
])

# --- TAB 1: CANDIDATE MODE ---
with tab1:
    if st.session_state.evaluation_result is not None:
        res = st.session_state.evaluation_result
        file_name = st.session_state.current_file_name
        role_name = st.session_state.current_role
        
        c_back, _ = st.columns([1, 4])
        with c_back:
            if st.button("⬅ Upload Another Candidate", type="secondary"):
                st.session_state.evaluation_result = None
                st.rerun()
                
        st.subheader("📊 Candidate Competency & ATS Intelligence Dossier")
        
        v1, v2 = st.columns(2)
        with v1:
            st.plotly_chart(render_gauge_chart(res.get("match_score", 0)), use_container_width=True)
        with v2:
            st.plotly_chart(render_radar_chart(res.get("competency_scores", {})), use_container_width=True)
            
        tier = res.get('priority_tier', 'Tier-2 (Strong Potential)')
        badge_class = "badge-tier1" if "Tier-1" in tier else ("badge-tier2" if "Tier-2" in tier else "badge-tier3")
        
        st.markdown(f"""
        <div class="section-card">
            <span class="{badge_class}">{tier}</span>
            <div style="margin-top: 10px; font-size: 14px; color: #E2E8F0;"><b>Executive Summary:</b> {res.get('priority_reason')}</div>
        </div>
        """, unsafe_allow_html=True)
        
        sub1, sub2, sub3, sub4, sub5 = st.tabs([
            "📊 Skill Matrix", 
            "🛠️ 30-Day Roadmap", 
            "🎤 Interview Kit", 
            "📝 AI Bullets",
            "✉️ AI Cover Letter"
        ])
        
        with sub1:
            s_col1, s_col2 = st.columns(2)
            with s_col1:
                st.success("✅ **Matching Strengths Found**")
                for item in res.get("strengths", []):
                    st.write(f"• {item}")
            with s_col2:
                st.error("⚠️ **Identified Skill Deficits**")
                for item in res.get("missing_skills", []):
                    st.write(f"• {item}")
                    
        with sub2:
            st.markdown("#### 🎯 Personalized 30-Day Upskilling Plan")
            for step in res.get("learning_roadmap", []):
                st.markdown(f"**🗓️ {step.get('week')}: {step.get('focus')}** ({step.get('resource_type')})")
                st.write(f"👉 *Action:* {step.get('task')}")
                st.divider()
                
        with sub3:
            st.markdown("#### 💡 Tailored Technical Interview Questions")
            for idx, q in enumerate(res.get("interview_prep_questions", []), 1):
                with st.expander(f"Q{idx}: {q.get('question')} (Topic: {q.get('topic')})"):
                    st.info(f"**Key Evaluator Anchor:** {q.get('expected_answer_hint')}")
                    
        with sub4:
            st.markdown("#### ✨ Suggested High-Impact Resume Bullets (ATS Optimized)")
            for bullet_pt in res.get("optimized_resume_bullets", []):
                st.markdown(f"> *{bullet_pt}*")
                
        with sub5:
            st.markdown("#### ✉️ Personalized AI Generated Cover Letter")
            cover_letter_text = res.get("custom_cover_letter", "No cover letter generated.")
            st.text_area("Generated Cover Letter", cover_letter_text, height=200)
                
        st.divider()
        pdf_doc = generate_dossier_pdf(file_name, role_name, res)
        st.download_button(
            label="📥 Download Official AI Candidate Dossier (PDF)",
            data=pdf_doc,
            file_name=f"{file_name}_AI_Report.pdf",
            mime="application/pdf"
        )
        
    else:
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("### 📄 1. Upload Resume Document")
            single_file = st.file_uploader(
                "Upload Resume (PDF / DOCX)",
                type=["pdf", "docx"],
                key="c_upload"
            )

        with col2:
            st.markdown("### 🎯 2. Target Role Requirements")
            role_choice = st.selectbox("Position Opening", list(SAMPLE_JDS.keys()), key="c_role")
            
            if role_choice == "Custom Job Description":
                target_role = st.text_input("Role Title", placeholder="e.g., Senior Full Stack Engineer")
                jd_input = st.text_area("Job Description Requirements", height=140, placeholder="Paste requirements...")
            else:
                target_role = role_choice
                default_jd = SAMPLE_JDS[role_choice]
                jd_input = st.text_area("Job Description Requirements (Editable)", value=default_jd, height=140)

        if single_file:
            resume_content = extract_text(single_file)

            if resume_content.strip():
                with st.expander("👁️ View Extracted Document Content"):
                    st.text_area("Extracted Raw Text", resume_content, height=140)
                    
                st.write("")
                if st.button("🚀 Run Comprehensive AI Evaluation", type="primary", key="btn_single"):
                    if not gemini_api_key.strip():
                        st.error("🔑 Please enter your Gemini API Key in the left sidebar.")
                    elif not jd_input.strip():
                        st.warning("⚠️ Please provide or select a Job Description.")
                    else:
                        with st.spinner("Extracting multi-dimensional vectors and synchronizing with database..."):
                            try:
                                res = run_deep_ai_evaluation(gemini_api_key, resume_content, jd_input, target_role)
                                save_full_evaluation(single_file.name, target_role, res)
                                st.session_state.evaluation_result = res
                                st.session_state.current_file_name = single_file.name
                                st.session_state.current_role = target_role
                                st.rerun()
                            except Exception as ex:
                                st.error(f"Analysis error: {ex}")
            else:
                st.error("⚠️ No readable text found in the file.")
        else:
            st.markdown("""
            <div style="background: #1E293B; border: 1px dashed #475569; border-radius: 12px; padding: 22px; text-align: center; margin-top: 15px;">
                <div style="font-size: 15px; font-weight: 700; color: #38BDF8; margin-bottom: 6px;">⚡ Quick Start Guide</div>
                <div style="font-size: 13px; color: #CBD5E1;">1. Enter your Gemini API Key in the left sidebar.<br>2. Upload a Resume (PDF/DOCX) above and select a Target Role.<br>3. Click 'Run Comprehensive AI Evaluation' to launch your dedicated Intelligence Dossier.</div>
            </div>
            """, unsafe_allow_html=True)

# --- TAB 2: RECRUITER MODE ---
with tab2:
    st.subheader("🏢 Batch Candidate Screening & Ranking Matrix")
    r1, r2 = st.columns(2)
    with r1:
        batch_docs = st.file_uploader("Upload Multiple Resumes (PDF / DOCX)", type=["pdf", "docx"], accept_multiple_files=True, key="b_upload")
    with r2:
        batch_role_pick = st.selectbox("Job Opening Role", list(SAMPLE_JDS.keys()), key="b_role")
        batch_jd_text = SAMPLE_JDS[batch_role_pick] if batch_role_pick != "Custom Job Description" else st.text_area("Job Requirements", key="b_jd")

    if batch_docs and st.button("⚡ Screen & Rank Candidate Pipeline", type="primary"):
        if not gemini_api_key.strip():
            st.error("🔑 Please enter Gemini API Key in the sidebar.")
        else:
            batch_results = []
            progress = st.progress(0)
            status = st.empty()
            
            for i, doc_file in enumerate(batch_docs):
                status.text(f"Evaluating candidate {i+1} of {len(batch_docs)}: {doc_file.name}")
                text = extract_text(doc_file)
                if text.strip():
                    try:
                        res = run_deep_ai_evaluation(gemini_api_key, text, batch_jd_text, batch_role_pick)
                        save_full_evaluation(doc_file.name, batch_role_pick, res)
                        batch_results.append({
                            "Candidate": doc_file.name,
                            "Match Score": res.get("match_score", 0),
                            "Priority Tier": res.get("priority_tier", "N/A"),
                            "Technical Alignment": f"{res.get('competency_scores', {}).get('technical_skills', 0)}%",
                            "ATS Readability": f"{res.get('competency_scores', {}).get('ats_readability', 0)}%"
                        })
                    except Exception as err:
                        pass
                progress.progress((i + 1) / len(batch_docs))
                
            status.text("Screening Complete! Pipeline saved to Database. ✅")
            if batch_results:
                b_df = pd.DataFrame(batch_results).sort_values(by="Match Score", ascending=False).reset_index(drop=True)
                b_df.index += 1
                st.subheader("🏆 Candidate Leaderboard & Shortlist")
                st.dataframe(b_df, use_container_width=True)

# --- TAB 3: DATABASE & HISTORY LOGS ---
with tab3:
    st.subheader("🗄️ Relational Database Management & Audit Logs")
    eval_records = fetch_evaluations_df()
    
    if not eval_records.empty:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Evaluations Count", len(eval_records))
        m2.metric("Average Score", f"{round(eval_records['match_score'].mean(), 1)}%")
        m3.metric("Top Score", f"{eval_records['match_score'].max()}%")
        tier1_count = len(eval_records[eval_records['priority_tier'].str.contains('Tier-1', na=False)])
        m4.metric("Tier-1 Shortlisted", tier1_count)
        
        st.dataframe(eval_records, use_container_width=True)
        
        st.divider()
        st.subheader("⚙️ Update Candidate Recruitment Status")
        st_c1, st_c2, st_c3 = st.columns(3)
        with st_c1:
            selected_eval_id = st.selectbox("Select Evaluation ID", eval_records["eval_id"].tolist())
        with st_c2:
            new_status_val = st.selectbox("Set Status", ["Under Review", "Shortlisted for Interview", "Technical Assessment Sent", "Rejected"])
        with st_c3:
            st.write("")
            st.write("")
            if st.button("Update Status in DB"):
                update_candidate_status(selected_eval_id, new_status_val)
                st.success("Status Updated in SQLite Database!")
                st.rerun()
                
        # Export
        csv_dump = eval_records.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Export Complete Database (CSV)", csv_dump, "All_Evaluations_Database_Dump.csv", "text/csv")
    else:
        st.info("No database records found. Execute evaluations to generate live logs.")
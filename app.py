import streamlit as st
import io
import re
import html

import fitz
import pytesseract
from PIL import Image
from docx import Document


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="AI Resume Analyzer",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# =========================================================
# CSS
# =========================================================

st.markdown("""
<style>

.stApp {
    background: linear-gradient(135deg, #080d1c, #0d1630, #080d1c);
    color: white;
}

.block-container {
    max-width: 1150px;
    padding-top: 25px;
    padding-bottom: 50px;
}

.hero {
    padding: 28px;
    border-radius: 22px;
    background: linear-gradient(135deg, #111d3a, #17264b);
    border: 1px solid #304a80;
    margin-bottom: 22px;
}

.hero h1 {
    margin: 0;
    color: white;
    font-size: 32px;
}

.hero p {
    color: #a8b8dc;
}

.step {
    text-align: center;
    padding: 14px 6px;
    border-radius: 15px;
    background: #101a31;
    border: 1px solid #293c68;
    color: #91a2c4;
    min-height: 70px;
}

.step.active {
    background: linear-gradient(135deg, #315ee8, #7047e8);
    color: white;
    border-color: #7d98ff;
}

.step-num {
    font-weight: 800;
    font-size: 18px;
}

.step-title {
    font-size: 12px;
    margin-top: 4px;
}

.card {
    background: linear-gradient(145deg, #111a31, #0c1428);
    border: 1px solid #293d69;
    border-radius: 18px;
    padding: 20px;
    margin-bottom: 16px;
}

.metric {
    background: #101a31;
    border: 1px solid #2c4272;
    border-radius: 16px;
    padding: 20px;
    text-align: center;
    min-height: 105px;
}

.metric-label {
    color: #91a4cb;
    font-size: 13px;
}

.metric-value {
    color: white;
    font-size: 25px;
    font-weight: 800;
    margin-top: 8px;
}

.skill {
    display: inline-block;
    padding: 7px 12px;
    margin: 4px;
    border-radius: 18px;
    background: #1b315c;
    border: 1px solid #4b6daa;
    color: #d5e1ff;
    font-size: 13px;
}

.job-card {
    background: #101a31;
    border: 1px solid #2c4272;
    border-radius: 16px;
    padding: 17px;
    margin-bottom: 12px;
}

.company-card {
    background: #101a31;
    border: 1px solid #2c4272;
    border-radius: 16px;
    padding: 17px;
    margin-bottom: 12px;
}

.match {
    color: #55e6a5;
    font-weight: 800;
}

.gap {
    color: #ffcf75;
}

.info-box {
    background: #101b34;
    border-left: 4px solid #5d83ff;
    border-radius: 10px;
    padding: 13px;
    margin: 10px 0;
}

.company-name {
    color: white;
    font-size: 18px;
    font-weight: 800;
}

.company-type {
    color: #94a8cf;
    font-size: 13px;
    margin-top: 5px;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# SKILLS
# =========================================================

SKILLS = {
    "Programming": [
        "Python", "C", "C++", "Java", "JavaScript", "R"
    ],

    "Data Science": [
        "Pandas", "NumPy", "Matplotlib", "Seaborn",
        "Scikit-learn", "TensorFlow", "PyTorch"
    ],

    "Database": [
        "SQL", "MySQL", "PostgreSQL", "MongoDB"
    ],

    "Web Development": [
        "HTML", "CSS", "JavaScript", "React",
        "Node.js", "Django", "Flask"
    ],

    "Data Analytics": [
        "Excel", "Power BI", "Tableau"
    ],

    "Cloud": [
        "AWS", "Azure", "Google Cloud", "GCP"
    ],

    "Tools": [
        "Git", "GitHub", "Docker", "Jupyter", "VS Code"
    ],

    "AI / ML": [
        "Machine Learning", "Deep Learning",
        "Natural Language Processing", "NLP",
        "Artificial Intelligence"
    ],

    "Soft Skills": [
        "Communication", "Leadership", "Teamwork",
        "Problem Solving", "Critical Thinking",
        "Time Management"
    ]
}


# =========================================================
# JOB ROLES
# =========================================================

JOB_ROLES = {
    "Data Analyst": [
        "Python", "SQL", "Excel", "Power BI",
        "Pandas", "NumPy", "Matplotlib"
    ],

    "Data Scientist": [
        "Python", "SQL", "Pandas", "NumPy",
        "Scikit-learn", "Machine Learning", "Matplotlib"
    ],

    "Machine Learning Engineer": [
        "Python", "Machine Learning", "Scikit-learn",
        "TensorFlow", "PyTorch", "SQL", "Git"
    ],

    "AI Engineer": [
        "Python", "Artificial Intelligence",
        "Machine Learning", "Deep Learning",
        "TensorFlow", "PyTorch"
    ],

    "Python Developer": [
        "Python", "Django", "Flask",
        "SQL", "Git", "GitHub"
    ],

    "Web Developer": [
        "HTML", "CSS", "JavaScript",
        "React", "Node.js", "Git"
    ],

    "Backend Developer": [
        "Python", "Django", "Flask",
        "SQL", "MongoDB", "Git"
    ],

    "Business Intelligence Analyst": [
        "Excel", "Power BI", "SQL",
        "Tableau", "Python"
    ],

    "Data Engineer": [
        "Python", "SQL", "MongoDB",
        "PostgreSQL", "AWS", "Docker", "Git"
    ],

    "Cloud Engineer": [
        "AWS", "Azure", "Google Cloud",
        "Docker", "Git", "Python"
    ]
}


# =========================================================
# COMPANIES
# =========================================================

COMPANIES = {
    "Data Analyst": [
        ("Accenture", "IT Services & Consulting"),
        ("Deloitte", "Consulting & Analytics"),
        ("TCS", "IT Services"),
        ("Infosys", "IT Services & Consulting"),
        ("Capgemini", "Technology & Consulting"),
        ("Wipro", "IT Services")
    ],

    "Data Scientist": [
        ("Microsoft", "Technology & AI"),
        ("Amazon", "Technology & Cloud"),
        ("IBM", "AI & Data Solutions"),
        ("Google", "Technology & AI"),
        ("Accenture", "Analytics & Consulting"),
        ("Deloitte", "Analytics & Consulting")
    ],

    "Machine Learning Engineer": [
        ("Google", "AI & Machine Learning"),
        ("Microsoft", "AI & Cloud"),
        ("Amazon", "AI & AWS"),
        ("NVIDIA", "AI & Computing"),
        ("IBM", "AI & Enterprise"),
        ("Accenture", "AI & Consulting")
    ],

    "AI Engineer": [
        ("Google", "Artificial Intelligence"),
        ("Microsoft", "Artificial Intelligence"),
        ("NVIDIA", "AI & Computing"),
        ("Amazon", "AI & Cloud"),
        ("IBM", "AI Solutions"),
        ("Adobe", "AI & Technology")
    ],

    "Python Developer": [
        ("TCS", "Software Development"),
        ("Infosys", "Software Development"),
        ("Accenture", "Technology Services"),
        ("Wipro", "Software & IT"),
        ("Deloitte", "Technology Consulting"),
        ("Zoho", "Software Products")
    ],

    "Web Developer": [
        ("Accenture", "Web & Technology"),
        ("TCS", "Software Development"),
        ("Infosys", "Technology Services"),
        ("Wipro", "IT Services"),
        ("Cognizant", "Technology Services"),
        ("Zoho", "Software Products")
    ],

    "Backend Developer": [
        ("Amazon", "Software & Cloud"),
        ("Microsoft", "Software Development"),
        ("Google", "Software Engineering"),
        ("TCS", "IT Services"),
        ("Infosys", "Software Development"),
        ("Cognizant", "Technology Services")
    ],

    "Business Intelligence Analyst": [
        ("Deloitte", "Business Intelligence"),
        ("Accenture", "Analytics"),
        ("TCS", "Data & Analytics"),
        ("Infosys", "Analytics"),
        ("Capgemini", "Data & BI"),
        ("EY", "Analytics & Consulting")
    ],

    "Data Engineer": [
        ("Amazon", "Data & Cloud"),
        ("Microsoft", "Data & Azure"),
        ("Google", "Data & Cloud"),
        ("IBM", "Data Engineering"),
        ("Accenture", "Data Engineering"),
        ("Deloitte", "Data & Analytics")
    ],

    "Cloud Engineer": [
        ("Amazon", "Cloud / AWS"),
        ("Microsoft", "Cloud / Azure"),
        ("Google", "Cloud / GCP"),
        ("IBM", "Cloud Computing"),
        ("Accenture", "Cloud Consulting"),
        ("TCS", "Cloud Services")
    ]
}


# =========================================================
# SESSION STATE
# =========================================================

if "page" not in st.session_state:
    st.session_state.page = 1

if "resume_text" not in st.session_state:
    st.session_state.resume_text = ""

if "source_name" not in st.session_state:
    st.session_state.source_name = ""


# =========================================================
# FUNCTIONS
# =========================================================

def extract_pdf_text(uploaded_file):
    try:
        pdf = fitz.open(
            stream=uploaded_file.getvalue(),
            filetype="pdf"
        )

        text = ""

        for page in pdf:
            text += page.get_text()

        pdf.close()

        return text

    except Exception as e:
        st.error("PDF Error: " + str(e))
        return ""


def extract_docx_text(uploaded_file):
    try:
        document = Document(
            io.BytesIO(uploaded_file.getvalue())
        )

        text = ""

        for paragraph in document.paragraphs:
            text += paragraph.text + "\n"

        for table in document.tables:
            for row in table.rows:
                cells = []

                for cell in row.cells:
                    cells.append(cell.text)

                text += " ".join(cells) + "\n"

        return text

    except Exception as e:
        st.error("DOCX Error: " + str(e))
        return ""


def extract_image_text(image):
    try:
        return pytesseract.image_to_string(image)

    except Exception as e:
        st.error("OCR Error: " + str(e))
        return ""


def extract_name(text):

    excluded = [
        "resume",
        "curriculum",
        "vitae",
        "email",
        "phone",
        "contact",
        "linkedin",
        "github"
    ]

    lines = []

    for line in text.splitlines():

        line = line.strip()

        if line:
            lines.append(line)

    for line in lines[:15]:

        low = line.lower()

        if any(
            word in low
            for word in excluded
        ):
            continue

        if "@" in line:
            continue

        if any(
            character.isdigit()
            for character in line
        ):
            continue

        if 1 <= len(line.split()) <= 5:

            if re.fullmatch(
                r"[A-Za-z .'-]+",
                line
            ):
                return line

    return "Not detected"


def extract_email(text):

    pattern = (
        r"[A-Za-z0-9._%+-]+"
        r"@[A-Za-z0-9.-]+"
        r"\.[A-Za-z]{2,}"
    )

    result = re.findall(
        pattern,
        text
    )

    if result:
        return result[0]

    return "Not detected"


def extract_phone(text):

    patterns = [
        r"(?:\+91[\s-]?)?[6-9]\d{9}",
        r"\+?\d{1,3}[\s-]?\d{10}"
    ]

    for pattern in patterns:

        result = re.findall(
            pattern,
            text
        )

        if result:
            return result[0]

    return "Not detected"


def skill_exists(text, skill):

    pattern = (
        r"(?<!\w)"
        + re.escape(skill.lower())
        + r"(?!\w)"
    )

    return re.search(
        pattern,
        text.lower()
    ) is not None


def detect_skills(text):

    detected = {}

    for category, skills in SKILLS.items():

        found = []

        for skill in skills:

            if skill_exists(text, skill):
                found.append(skill)

        if found:
            detected[category] = found

    return detected


def recommend_jobs(detected_skills):

    all_skills = []

    for skills in detected_skills.values():
        all_skills.extend(skills)

    skill_set = set()

    for skill in all_skills:
        skill_set.add(skill.lower())

    results = []

    for job, required in JOB_ROLES.items():

        matched = []
        missing = []

        for skill in required:

            if skill.lower() in skill_set:
                matched.append(skill)
            else:
                missing.append(skill)

        score = round(
            len(matched) / len(required) * 100,
            1
        )

        results.append({
            "Job": job,
            "Score": score,
            "Matched": matched,
            "Missing": missing
        })

    results.sort(
        key=lambda item: item["Score"],
        reverse=True
    )

    return results


def calculate_resume_score(
    text,
    detected_skills
):

    score = 0
    lower_text = text.lower()

    skill_count = sum(
        len(items)
        for items in detected_skills.values()
    )

    score += min(
        skill_count * 3,
        30
    )

    education_words = [
        "education",
        "b.tech",
        "btech",
        "bachelor",
        "degree",
        "university",
        "college",
        "engineering",
        "master"
    ]

    if any(
        word in lower_text
        for word in education_words
    ):
        score += 15

    if "project" in lower_text:
        score += 15

    if (
        "experience" in lower_text
        or "internship" in lower_text
    ):
        score += 15

    if (
        "certification" in lower_text
        or "certificate" in lower_text
    ):
        score += 10

    if extract_email(text) != "Not detected":
        score += 5

    if extract_phone(text) != "Not detected":
        score += 5

    return min(score, 100)


def show_metric(label, value):

    st.markdown(
        '<div class="metric">'
        '<div class="metric-label">'
        + html.escape(str(label))
        + '</div>'
        '<div class="metric-value">'
        + html.escape(str(value))
        + '</div>'
        '</div>',
        unsafe_allow_html=True
    )


# =========================================================
# HEADER
# =========================================================

st.markdown("""
<div class="hero">
<h1>🤖 AI Resume Analyzer & Job Recommendation System</h1>
<p>
Upload or scan your resume • Analyze Skills • Find Jobs •
Discover Companies • Identify Skill Gaps
</p>
</div>
""", unsafe_allow_html=True)


# =========================================================
# STEP BAR
# =========================================================

step_names = [
    "📤 Resume Input",
    "👤 Resume Overview",
    "💼 Job Matches",
    "🎯 Skill Gap"
]

step_columns = st.columns(4)

for number, column in enumerate(
    step_columns,
    start=1
):

    with column:

        if number == st.session_state.page:
            css_class = "step active"
        else:
            css_class = "step"

        st.markdown(
            '<div class="' + css_class + '">'
            '<div class="step-num">'
            + str(number)
            + '</div>'
            '<div class="step-title">'
            + step_names[number - 1]
            + '</div>'
            '</div>',
            unsafe_allow_html=True
        )


st.markdown("<br>", unsafe_allow_html=True)


# =========================================================
# PAGE 1 - RESUME INPUT
# =========================================================

if st.session_state.page == 1:

    st.markdown("""
    <div class="card">
    <h2>📤 Resume Input</h2>
    <p>
    Upload your resume as PDF/DOCX or scan it using your camera.
    </p>
    </div>
    """, unsafe_allow_html=True)

    left, right = st.columns(2)

    with left:

        st.subheader("📁 Upload Resume")

        uploaded_file = st.file_uploader(
            "Choose PDF or DOCX",
            type=["pdf", "docx"]
        )

        if uploaded_file is not None:

            if uploaded_file.name.lower().endswith(".pdf"):

                extracted_text = extract_pdf_text(
                    uploaded_file
                )

            else:

                extracted_text = extract_docx_text(
                    uploaded_file
                )

            if extracted_text.strip():

                st.session_state.resume_text = extracted_text

                st.session_state.source_name = (
                    uploaded_file.name
                )

                st.success(
                    "✅ Resume uploaded successfully."
                )

            else:

                st.warning(
                    "⚠️ No readable text found."
                )

    with right:

        st.subheader("📷 Scan Resume")

        camera_image = st.camera_input(
            "Take a photo of your resume"
        )

        if camera_image is not None:

            image = Image.open(
                camera_image
            )

            with st.spinner(
                "Reading resume using OCR..."
            ):

                extracted_text = extract_image_text(
                    image
                )

            if extracted_text.strip():

                st.session_state.resume_text = (
                    extracted_text
                )

                st.session_state.source_name = (
                    "Camera Scan"
                )

                st.success(
                    "✅ Resume scanned successfully."
                )

            else:

                st.warning(
                    "⚠️ Could not read text from image."
                )

    if st.session_state.resume_text:

        st.markdown("### 📄 Resume Loaded")

        st.info(
            "Source: "
            + st.session_state.source_name
        )

        with st.expander(
            "Preview Extracted Resume Text"
        ):

            st.text(
                st.session_state.resume_text[:5000]
            )


# =========================================================
# PAGE 2 - RESUME OVERVIEW
# =========================================================

elif st.session_state.page == 2:

    text = st.session_state.resume_text

    if not text.strip():

        st.warning(
            "⚠️ Please upload or scan a resume first."
        )

    else:

        detected_skills = detect_skills(text)

        score = calculate_resume_score(
            text,
            detected_skills
        )

        name = extract_name(text)
        email = extract_email(text)
        phone = extract_phone(text)

        st.markdown("""
        <div class="card">
        <h2>👤 Resume Overview</h2>
        <p>
        Information automatically detected from your resume.
        </p>
        </div>
        """, unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)

        with c1:
            show_metric(
                "Resume Score",
                str(score) + "/100"
            )

        with c2:

            total_skills = sum(
                len(items)
                for items in detected_skills.values()
            )

            show_metric(
                "Skills Detected",
                total_skills
            )

        with c3:

            show_metric(
                "Source",
                st.session_state.source_name
            )

        st.markdown(
            "### 👤 Personal Information"
        )

        p1, p2, p3 = st.columns(3)

        with p1:

            st.markdown(
                '<div class="card">'
                '<b>Name</b><br><br>'
                + html.escape(name)
                + '</div>',
                unsafe_allow_html=True
            )

        with p2:

            st.markdown(
                '<div class="card">'
                '<b>Email</b><br><br>'
                + html.escape(email)
                + '</div>',
                unsafe_allow_html=True
            )

        with p3:

            st.markdown(
                '<div class="card">'
                '<b>Phone</b><br><br>'
                + html.escape(phone)
                + '</div>',
                unsafe_allow_html=True
            )

        st.markdown(
            "### 🛠️ Detected Skills"
        )
if detected_skills:

            for category, skills in detected_skills.items():

                st.markdown(
                    "#### " + category
                )

                skill_html = ""

                for skill in skills:

                    skill_html += (
                        '<span class="skill">'
                        + html.escape(skill)
                        + '</span>'
                    )

                st.markdown(
                    skill_html,
                    unsafe_allow_html=True
                )

        else:

            st.warning(
                "No predefined skills were detected."
            )


# =========================================================
# PAGE 3 - JOB MATCHES
# =========================================================

elif st.session_state.page == 3:

    text = st.session_state.resume_text

    if not text.strip():

        st.warning(
            "⚠️ Please upload a resume first."
        )

    else:

        detected_skills = detect_skills(text)

        results = recommend_jobs(
            detected_skills
        )

        st.markdown("""
        <div class="card">
        <h2>💼 Job & Company Recommendations</h2>
        <p>
        Jobs are ranked according to skills detected in your resume.
        </p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(
            "### 💼 Recommended Job Roles"
        )

        for result in results[:7]:

            job = result["Job"]
            score = result["Score"]

            matched = ", ".join(
                result["Matched"]
            )

            missing = ", ".join(
                result["Missing"]
            )

            if not matched:
                matched = "None"

            if not missing:
                missing = "None"

            st.markdown(
                '<div class="job-card">'
                '<h3>💼 '
                + html.escape(job)
                + '</h3>'
                '<p class="match">'
                'Match Score: '
                + str(score)
                + '%</p>'
                '<p><b>Matched Skills:</b> '
                + html.escape(matched)
                + '</p>'
                '<p class="gap"><b>Missing Skills:</b> '
                + html.escape(missing)
                + '</p>'
                '</div>',
                unsafe_allow_html=True
            )

        st.markdown(
            "### 🏢 Recommended Companies"
        )

        shown_companies = set()

        for result in results[:3]:

            job = result["Job"]

            companies = COMPANIES.get(
                job,
                []
            )

            for company, company_type in companies:

                if company in shown_companies:
                    continue

                shown_companies.add(company)

                st.markdown(
                    '<div class="company-card">'
                    '<div class="company-name">'
                    '🏢 '
                    + html.escape(company)
                    + '</div>'
                    '<div class="company-type">'
                    + html.escape(company_type)
                    + '</div>'
                    '<div class="company-type">'
                    'Suitable role: '
                    + html.escape(job)
                    + '</div>'
                    '</div>',
                    unsafe_allow_html=True
                )

                if len(shown_companies) >= 10:
                    break

            if len(shown_companies) >= 10:
                break

        st.info(
            "💡 Company recommendations are career-role "
            "suggestions, not live job vacancies."
        )


# =========================================================
# PAGE 4 - SKILL GAP
# =========================================================

elif st.session_state.page == 4:

    text = st.session_state.resume_text

    if not text.strip():

        st.warning(
            "⚠️ Please upload a resume first."
        )

    else:

        detected_skills = detect_skills(text)

        st.markdown("""
        <div class="card">
        <h2>🎯 Skill Gap Analysis</h2>
        <p>
        Select a target career role to identify the skills
        you should learn next.
        </p>
        </div>
        """, unsafe_allow_html=True)

        selected_job = st.selectbox(
            "🎯 Select Target Job Role",
            list(JOB_ROLES.keys())
        )

        required_skills = JOB_ROLES[
            selected_job
        ]

        detected_all = []

        for skills in detected_skills.values():
            detected_all.extend(skills)

        detected_lower = set()

        for skill in detected_all:
            detected_lower.add(
                skill.lower()
            )

        matched = []
        missing = []

        for skill in required_skills:

            if skill.lower() in detected_lower:
                matched.append(skill)
            else:
                missing.append(skill)

        match_percentage = round(
            len(matched)
            / len(required_skills)
            * 100,
            1
        )

        c1, c2, c3 = st.columns(3)

        with c1:

            show_metric(
                "Target Role",
                selected_job
            )

        with c2:

            show_metric(
                "Skill Match",
                str(match_percentage) + "%"
            )

        with c3:

            show_metric(
                "Skills To Learn",
                len(missing)
            )

        st.markdown(
            "### ✅ Skills You Already Have"
        )

        if matched:

            skill_html = ""

            for skill in matched:

                skill_html += (
                    '<span class="skill">'
                    + html.escape(skill)
                    + '</span>'
                )

            st.markdown(
                skill_html,
                unsafe_allow_html=True
            )

        else:

            st.info(
                "No matching skills found for this role."
            )

        st.markdown(
            "### 📚 Skills You Should Learn"
        )

        if missing:

            for index, skill in enumerate(
                missing,
                start=1
            ):

                st.markdown(
                    '<div class="info-box">'
                    '<b>Step '
                    + str(index)
                    + '</b> → Learn <b>'
                    + html.escape(skill)
                    + '</b>'
                    '</div>',
                    unsafe_allow_html=True
                )

        else:

            st.success(
                "🎉 Your resume already covers "
                "all required skills."
            )

        st.markdown(
            "### 🚀 Suggested Learning Roadmap"
        )

        if missing:

            st.write(
                "Learn the missing skills, build projects "
                "using them, add the projects to your resume, "
                "and then apply for suitable roles."
            )

        else:

            st.success(
                "🎉 You are ready to start applying "
                "for this target role."
            )
# =========================================================
# NAVIGATION
# =========================================================

st.markdown(
    "<br>",
    unsafe_allow_html=True
)

back_col, next_col, end_col = st.columns(3)


with back_col:

    if st.session_state.page > 1:

        if st.button(
            "← Back",
            use_container_width=True
        ):

            st.session_state.page -= 1
            st.rerun()


with next_col:

    if st.session_state.page < 4:

        if st.button(
            "Next →",
            use_container_width=True
        ):

            if not st.session_state.resume_text.strip():

                st.warning(
                    "⚠️ Please upload or scan your resume first."
                )

            else:

                st.session_state.page += 1
                st.rerun()


with end_col:

    if st.session_state.page == 4:

        if st.button(
            "✓ Analysis Complete",
            use_container_width=True
        ):

            st.success(
                "🎉 Resume analysis completed successfully!"
            )

import os
import re
import sys
import json
import copy
import glob
import shutil
import ctypes
import subprocess
import threading
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser

# Optional PyMuPDF for instant in-app PDF page rendering
try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

# Enable crisp High-DPI font scaling on Windows
if sys.platform.startswith("win"):
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

DATA_FILE = "joseph_resume_v2.json"
TEX_FILE = "resume.tex"
PDF_FILE = "resume.pdf"
LOG_FILE = "resume.log"

COLORS = {
    "bg_main": "#f1f5f9",
    "bg_card": "#ffffff",
    "bg_header": "#003566",
    "primary": "#004f90",       # Matches your \definecolor{primaryColor}{RGB}{0, 79, 144}
    "primary_hover": "#003b6d",
    "emerald": "#059669",
    "emerald_hover": "#047857",
    "amber": "#d97706",
    "danger": "#dc2626",
    "fg_primary": "#0f172a",
    "fg_muted": "#64748b",
    "border": "#cbd5e1",
    "code_bg": "#1e1e2e",
    "code_fg": "#cdd6f4",
    "hl_line": "#313244",
}

DEFAULT_LAYOUT = {
    "margin_cm": "2.0",
    "footskip_cm": "1.0",
    "right_col_cm": "5.2",
    "sec_before_cm": "0.3",
    "sec_after_cm": "0.2",
    "bullet_glyph": r"\circ",
    "primary_rgb": "0, 79, 144",
    "footer_page_of_last": False,
}


def build_exact_preamble(layout: dict, footer_name: str) -> str:
    m = layout.get("margin_cm", "2.0")
    fs = layout.get("footskip_cm", "1.0")
    rcol = layout.get("right_col_cm", "5.2")
    sb = layout.get("sec_before_cm", "0.3")
    sa = layout.get("sec_after_cm", "0.2")
    glyph = layout.get("bullet_glyph", r"\circ")
    rgb = layout.get("primary_rgb", "0, 79, 144")
    page_fmt = (
        r"\fancyfoot[R]{\color{gray}\textit{\small Page \thepage\ of \pageref*{LastPage}}}"
        if layout.get("footer_page_of_last", False)
        else r"\fancyfoot[R]{\color{gray}\textit{\small \thepage}}"
    )
    esc_footer = smart_escape_latex(footer_name)

    return rf"""\documentclass[10pt, letterpaper]{{article}}

% Packages:
\usepackage[
    ignoreheadfoot,
    top={m}cm,
    bottom={m}cm,
    left={m}cm,
    right={m}cm,
    footskip={fs}cm,
]{{geometry}}
\usepackage{{titlesec}}
\usepackage{{tabularx}}
\usepackage{{array}}
\usepackage[dvipsnames]{{xcolor}}
\definecolor{{primaryColor}}{{RGB}}{{{rgb}}}
\usepackage{{enumitem}}
\usepackage{{fontawesome5}}
\usepackage{{amsmath}}
\usepackage[
    pdftitle={{{esc_footer}'s CV}},
    pdfauthor={{{esc_footer}}},
    pdfcreator={{LaTeX with RenderCV}},
    colorlinks=true,
    urlcolor=primaryColor,
    linkcolor=black,
    breaklinks=true
]{{hyperref}}
\usepackage{{fancyhdr}}
\usepackage{{calc}}
\usepackage{{bookmark}}
\usepackage{{lastpage}}
\usepackage{{changepage}}
\usepackage{{paracol}}
\usepackage{{ifthen}}
\usepackage{{needspace}}
\usepackage{{iftex}}
\usepackage{{paracol}}
\footnotelayout{{m}}

\ifPDFTeX
    \input{{glyphtounicode}}
    \pdfgentounicode=1
    \usepackage[utf8]{{inputenc}}
    \usepackage{{lmodern}}
\fi

\AtBeginEnvironment{{adjustwidth}}{{\partopsep0pt}}
\setcounter{{secnumdepth}}{{0}}
\setlength{{\parindent}}{{0pt}}
\setlength{{\topskip}}{{0pt}}
\setlength{{\columnsep}}{{0cm}}

% --- Footer Redefinition ---
\pagestyle{{fancy}}
\fancyhf{{}}
\fancyfoot[C]{{\color{{gray}}\textit{{\small {esc_footer}}}}}
{page_fmt}
\renewcommand{{\headrulewidth}}{{0pt}}
\renewcommand{{\footrulewidth}}{{0pt}}
% ---------------------------

% Optimal spacing: ensures section titles stay with minimum content while allowing natural breaks
\titleformat{{\section}}{{\needspace{{4\baselineskip}}\bfseries\large}}{{}}{{0pt}}{{}}[\vspace{{1pt}}\titlerule]
\titlespacing{{\section}}{{-1pt}}{{{sb}cm}}{{{sa}cm}}

\renewcommand\labelitemi{{${glyph}$}}
\newenvironment{{highlights}}{{
    \begin{{itemize}}[
        topsep=0.10cm,
        parsep=0.10cm,
        partopsep=0pt,
        itemsep=0pt,
        leftmargin=0.4cm + 10pt
    ]
}}{{
    \end{{itemize}}
}}
\newenvironment{{onecolentry}}{{
    \begin{{adjustwidth}}{{0.2cm + 0.00001cm}}{{0.2cm + 0.00001cm}}
}}{{
    \end{{adjustwidth}}
}}
\newenvironment{{twocolentry}}[2][]{{
    \onecolentry
    \def\secondColumn{{#2}}
    \setcolumnwidth{{\fill, {rcol}cm}}
    \begin{{paracol}}{{2}}
}}{{
    \switchcolumn \raggedleft \secondColumn
    \end{{paracol}}
    \endonecolentry
}}
\newenvironment{{header}}{{
    \setlength{{\topsep}}{{0pt}}\par\kern\topsep\centering\linespread{{1.5}}
}}{{
    \par\kern\topsep
}}

% Environment for project entries - prevents splitting while allowing natural page breaks
\newenvironment{{projectentry}}{{
    \needspace{{2.5\baselineskip}}
}}{{
}}

% Environment for first entry after section - minimal space since section already reserves space
\newenvironment{{firstprojectentry}}{{
    \needspace{{2\baselineskip}}
}}{{
}}

\let\hrefWithoutArrow\href
\renewcommand{{\href}}[2]{{\hrefWithoutArrow{{#1}}{{\ifthenelse{{\equal{{#2}}{{}}}}{{ }}{{#2 }}\raisebox{{.15ex}}{{\footnotesize \faExternalLink*}}}}}}
"""


def smart_escape_latex(text: str) -> str:
    """Escapes &, %, #, _ outside $...$ math mode while preserving valid LaTeX commands."""
    if not text:
        return ""
    replacements = {"&": r"\&", "%": r"\%", "#": r"\#", "_": r"\_"}

    def clean_chunk(chunk: str) -> str:
        for char, esc in replacements.items():
            chunk = re.sub(rf"(?<!\\){re.escape(char)}", esc, chunk)
        return chunk

    if "$" in text:
        parts = re.split(r"(\$[^\$]+\$)", text)
        return "".join(
            p if (p.startswith("$") and p.endswith("$") and len(p) >= 2) else clean_chunk(p)
            for p in parts
        )
    return clean_chunk(text)


# ==============================================================================
# Pre-Loaded Data Extracted From Your resume.tex
# ==============================================================================
DEFAULT_HEADER = {
    "display_name": "JOSEPH CHAY",
    "footer_name": "Joseph Emmanuel Chay Huan Qin",
    "phone": "(+60) 16-313-9132",
    "email": "josephemmanuelchay@gmail.com",
    "linkedin_url": "https://my.linkedin.com/in/josephchay",
    "linkedin_disp": "linkedin.com/in/josephchay",
    "github_url": "https://github.com/josephchay",
    "github_disp": "github.com/josephchay",
}

DEFAULT_PROFILES = [
    {
        "name": "0. Active: Technology, AI & Multidisciplinary Interests",
        "text": (
            "Technology by far sectoring the whole world through enhancements, opportunitize innovations in every economical domains. "
            "This instrument overabused in achievements by various artificial and machine intelligences is majorly diminished only in image, "
            "which such mere toolsets treated definitively yet producing superficial outcomes incomparable with areas of non-automations. "
            "End-services fear of pending replacement stir methods to persuasive genuinity. "
            "Beyond education, projects, and work experience, motivated communications enable potential contributions to any given subject and object. "
            "Physics, mechanics, mathematics, biochemistry, programming and architect development, are my standard interests."
        ),
    },
    {
        "name": "1. AI, ML, Robotics & Data Engineering",
        "text": (
            "Computer Science with Artificial Intelligence graduate and a Diploma holder in Information and Communication Technology with Software Engineering, "
            "bringing a strong passion and well-developed expertise in building intelligent, scalable systems that solve real-world problems. My skill set spans machine learning, "
            "reinforcement learning, deep learning, NLP, computer vision, object detection, image processing, robotics automation, and data engineering, "
            "supported by solid foundations in software engineering, cloud computing, and distributed systems."
        ),
    },
    {
        "name": "2. Cross-Industry AI & Full-Stack Engineering",
        "text": (
            "Worked on simple projects in university involving intelligent agents, robotics, object detection, RAG-based systems, chatbot agents, and large-scale data pipelines. "
            "My knowledge and hands-on expertise position me to contribute across industries such as healthcare (medical imaging, AI diagnostics), "
            "finance (fraud detection, predictive analytics), e-commerce (personalization, recommendation systems), "
            "software engineering (web development, mobile app development, API integration, database design and management, cloud deployment)."
        ),
    },
    {
        "name": "3. Client Communication & Team Building",
        "text": (
            "Beyond software development, I carry great communication skills and motivation to work with people, including team buildings. "
            "Spoken roles interest me as it builds understanding and trust between people indulging clients into the products."
        ),
    },
    {
        "name": "4. Systems Monitoring, Surveillance & Access Control",
        "text": (
            "Beyond software development, I bring a strong aptitude for systems monitoring and technical troubleshooting, skills that translate directly into overseeing equipment-dependent environments. "
            "My background in computer vision and object detection gives me a practical understanding of how CCTV and surveillance systems function at a technical level, allowing me to quickly identify "
            "malfunctions, blind spots, or synchronization issues rather than treating such equipment as a black box. Similarly, my experience with API integration, database design, and access control "
            "logic in software projects provides a solid foundation for handling access card programming, entry/exit system synchronization, and coordination with equipment service providers when "
            "technical issues arise."
        ),
    },
    {
        "name": "5. Data Engineering & Operational Audits",
        "text": (
            "My data engineering background also lends itself naturally to operational improvement initiatives. I am accustomed to building structured processes for collecting, organizing, and analyzing "
            "data to identify patterns and inefficiencies, an approach that applies directly to daily checklist audits, incident tracking, and reporting on recurring equipment or operational issues. "
            "I take a methodical, detail-oriented approach to routine inspections and follow-ups, ensuring outstanding issues are properly logged, escalated, and resolved rather than overlooked."
        ),
    },
]

DEFAULT_SECTION_ORDER = [
    "Professional Profile",
    "Education",
    "Industrial Experience",
    "Honours & Awards",
    "Client Projects",
    "Personal Projects",
    "Proficiency and Tools",
    "Languages",
]

DEFAULT_ITEMS = [
    # EDUCATION
    {
        "uid": "edu_phd",
        "enabled": True,
        "section": "Education",
        "type": "education",
        "institution": "University of Nottingham",
        "degree": "Doctor of Philosophy in Computer Science",
        "date": "Sep 2025 -- March 2026",
        "is_sub_degree": False,
        "links": [
            {"label": "Research Proposal", "url": "https://github.com/josephchay/phd-proposal/blob/main/multi-qubit-rgb-encoding-architecture-proposal.pdf"},
            {"label": "Code & Experiments", "url": "https://github.com/josephchay/seqnn-superpixel-encoding/tree/develop"},
        ],
        "bullets": [
            "Research Focus: Quantum Computing, Neural Networks, Deep Learning",
            "Supervisors: Professor Dr. Tan Chye Cheah",
            "Co-supervisors: Professor Dr. Tomas Maul, Professor Dr. Simon Lau Boung Yew",
            "(withdrew with no formal submissions following health conditions)",
        ],
    },
    {
        "uid": "edu_bsc",
        "enabled": True,
        "section": "Education",
        "type": "education",
        "institution": "University of Nottingham",
        "degree": "BSc (Hons). Computer Science with Artificial Intelligence",
        "date": "Sep 2023 -- Jul 2025",
        "is_sub_degree": True,
        "links": [],
        "bullets": [
            "Granted direct Year 2 entry due to diploma studies",
            "(contain academic misconducts)",
            "WAM: 76%",
        ],
    },
    {
        "uid": "edu_dip",
        "enabled": True,
        "section": "Education",
        "type": "education",
        "institution": "Asia Pacific University of Technology and Innovation",
        "degree": "Dip. Information and Communication Technology with Software Engineering",
        "date": "Dec 2020 -- Mar 2023",
        "is_sub_degree": False,
        "links": [],
        "bullets": ["CGPA: 3.89/4.00"],
    },
    {
        "uid": "edu_igcse",
        "enabled": True,
        "section": "Education",
        "type": "education",
        "institution": "Regent International School",
        "degree": "IGCSE",
        "date": "Sep 2017 -- Nov 2020",
        "is_sub_degree": False,
        "links": [],
        "bullets": ["Grades: 3A*, 5A, 2B"],
    },
    {
        "uid": "edu_found",
        "enabled": True,
        "section": "Education",
        "type": "education",
        "institution": "RightBridge Academy",
        "degree": "IGCSE Programme (Foundation Year)",
        "date": "Sep 2016 -- Aug 2017",
        "is_sub_degree": False,
        "links": [],
        "bullets": ["Completed first year of IGCSE studies"],
    },
    # INDUSTRIAL EXPERIENCE
    {
        "uid": "exp_shopee",
        "enabled": True,
        "section": "Industrial Experience",
        "type": "experience",
        "company": "Shopee SPX Courier Kundang Hub",
        "role": "Parcel Delivery Driver (Rawang, Selangor, Malaysia)",
        "date": "19\\textsuperscript{th} Sep 2026 -- 22\\textsuperscript{nd} Sep 2026",
        "is_sub_role": False,
        "space_after": "0.4cm",
        "bullets": [
            "Delivered +100 parcels to customers across households",
            "Collected proof of delivery (POD) signatures and photos for each parcel delivered",
        ],
    },
    {
        "uid": "exp_ips_swe",
        "enabled": True,
        "section": "Industrial Experience",
        "type": "experience",
        "company": "IPS Software Sdn Bhd",
        "role": "Software Developer",
        "date": "Feb 2023 -- Aug 2023",
        "is_sub_role": False,
        "space_after": "0.4cm",
        "bullets": [
            "Mentored interns, improving team productivity",
            "Responsive web applications for clients",
            "Reponsive mobile applications for clients",
            "Designed, modified, and maintained designs for applications",
            "Third-party APIs with RESTful Postman for cross-platform functionality",
        ],
    },
    {
        "uid": "exp_ips_ai",
        "enabled": False,
        "section": "Industrial Experience",
        "type": "experience",
        "company": "IPS Software Sdn Bhd",
        "role": "AI Developer",
        "date": "",
        "is_sub_role": True,
        "space_after": "0.4cm",
        "bullets": [
            "BERT-based chatbot with style transfers",
            "CNN-based facial recognition for POS",
            "YOLO-based real-time document forgery detection",
            "GAN architecture for customer behavior segmentation",
            "RNN-based anomaly detection for transaction patterns",
            "ResNet-based biometric verification system",
        ],
    },
    {
        "uid": "exp_spec",
        "enabled": True,
        "section": "Industrial Experience",
        "type": "experience",
        "company": "Spectrum Edge Sdn Bhd",
        "role": "Cybersecurity Developer",
        "date": "Dec 2022 -- Feb 2023",
        "is_sub_role": False,
        "space_after": "0.2cm",
        "bullets": [
            "Security POC projects for client environments",
            "Firewall rules configuration and security policies",
            "Python Scripts for Fortinet configuration conversions",
        ],
    },
    # HONOURS & AWARDS
    {
        "uid": "awd_fyp",
        "enabled": False,
        "section": "Honours & Awards",
        "type": "award",
        "title": "Best Final Year Project (1\\textsuperscript{st} Runner-Up), Final Year Dissertation",
        "date": "Jul 2025",
    },
    {
        "uid": "awd_ind",
        "enabled": False,
        "section": "Honours & Awards",
        "type": "award",
        "title": "Best Industry Project (Top 5), Software Engineering Group Project\\footnote{Awarded during the closing ceremony of the open demonstration day for our project, \\textit{Smart Fridge}.}",
        "date": "Apr 2024",
    },
    {
        "uid": "awd_abrsm_th",
        "enabled": False,
        "section": "Honours & Awards",
        "type": "award",
        "title": "ABRSM Piano Grade 8 Theory, Passed with Distinction",
        "date": "May 2020",
    },
    {
        "uid": "awd_dhp",
        "enabled": True,
        "section": "Honours & Awards",
        "type": "award",
        "title": "Deputy Head Prefect, Regent International School",
        "date": "Dec 2019",
    },
    {
        "uid": "awd_abrsm_pr",
        "enabled": True,
        "section": "Honours & Awards",
        "type": "award",
        "title": "ABRSM Piano Grade 8 Practical, Passed with Distinction",
        "date": "Jul 2019",
    },
    {
        "uid": "awd_hp",
        "enabled": True,
        "section": "Honours & Awards",
        "type": "award",
        "title": "Head Prefect, Regent International School",
        "date": "Dec 2018",
    },
    # CLIENT PROJECTS
    {
        "uid": "cproj_yamaha",
        "enabled": False,
        "section": "Client Projects",
        "type": "project",
        "title": "Yamaha & Bentley Music Synthesis AI, TTM & L2S Generation",
        "date": "Jun 2024 -- May 2025",
        "subtitle_italic": "Machine Learning, Generative Modelling, Diffusion Frameworks, Text-to-Music, Lyric-to-Song",
        "stack": "PyTorch, Latent Diffusion, U-Nets, Fourier Transforms, Mel Spectrograms",
        "links": [
            {"label": "Code", "url": "https://github.com/Soundsation/pipeline/"},
            {"label": "Demo", "url": "https://josephchay.github.io/soundsation.github.io/"},
            {"label": "Demo 2", "url": "https://josephchay.github.io/graduation-songs/"},
            {"label": "Thesis Paper", "url": "https://github.com/Soundsation/documentation/tree/main"},
        ],
        "extra_bullets": [],
    },
    {
        "uid": "cproj_vixta",
        "enabled": False,
        "section": "Client Projects",
        "type": "project",
        "title": "Vixta Fashion Studio, Malaysian Fashion Designer Studio",
        "date": "Mar 2024 -- Jun 2024",
        "subtitle_italic": "Concept, Design. Portfolio, Micro Interactions, Deployment",
        "stack": "VueJS, TailwindCSS, Yarn, Framer Motion, GSAP",
        "links": [
            {"label": "Code", "url": "https://github.com/josephchay/vixta-fashion-studio"},
            {"label": "Demo", "url": "https://vixta.netlify.app/"},
        ],
        "extra_bullets": [],
    },
    {
        "uid": "cproj_jessica",
        "enabled": False,
        "section": "Client Projects",
        "type": "project",
        "title": "Jessica Bauer Conte, YouTuber and Vlogger",
        "date": "Feb 2024 -- May 2024",
        "subtitle_italic": "Concept, Design, Portfolio, Micro Interactions, Deployment",
        "stack": "VueJS, TailwindCSS, Yarn, Framer Motion, GSAP",
        "links": [
            {"label": "Code", "url": "https://github.com/josephchay/jessica-conte-folio"},
            {"label": "Demo", "url": "https://jess-conte.netlify.app/"},
        ],
        "extra_bullets": [],
    },
    # PERSONAL PROJECTS
    {
        "uid": "pproj_docket",
        "enabled": True,
        "section": "Personal Projects",
        "type": "project",
        "title": "Docket Notes, Interactive Note Taking and Diary Application",
        "date": "Feb 2024 -- Aug 2026",
        "subtitle_italic": "Design, Development, Database, Deployment",
        "stack": "ReactJS, JavaScript, TailwindCSS, xState, Framer Motion, AnimeJS, GSAP",
        "links": [{"label": "Code & Sample Screens", "url": "https://github.com/josephchay/docket-notes/tree/develop"}],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_codeflux",
        "enabled": True,
        "section": "Personal Projects",
        "type": "project",
        "title": "CodeFlux, Unravel and Refactor Codebase",
        "date": "Jan 2025 -- May 2025",
        "subtitle_italic": "Codebase Handler, Code Structuring, Code Refactoring, Code Documentation",
        "stack": "Python, CSV, JavaScript",
        "links": [{"label": "Code & Usage Guides", "url": "https://github.com/josephchay/codeflux/tree/develop"}],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_restzero",
        "enabled": True,
        "section": "Personal Projects",
        "type": "project",
        "title": "Restoration Zero, Platformer Game",
        "date": "Jan 2025 -- Mar 2025",
        "subtitle_italic": "2D, Pixel Art, Storytelling, Aesthetic Design, Physics and Maths, Sprite Animation",
        "stack": "Python, Pygame, Microsoft Paint",
        "links": [{"label": "Code & Sample Screens", "url": "https://github.com/josephchay/restoration-zero/tree/feat/entities/player-ai"}],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_linguify",
        "enabled": False,
        "section": "Personal Projects",
        "type": "project",
        "title": "Linguify - Low Latency Conversational AI Chatbot",
        "date": "Jan 2025 -- Apr 2025",
        "subtitle_italic": "Generative AI, Conversational AI, Low Latency, Real-time Processing, NLP",
        "stack": "PyTorch, RAG, LangChain, TTS, STT, Web Scraping",
        "links": [{"label": "Code", "url": "https://github.com/josephchay/linguify"}],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_sadldas",
        "enabled": False,
        "section": "Personal Projects",
        "type": "project",
        "title": "Stable Adaptive Decompositional LDAS",
        "date": "Sep 2024 -- Dec 2024",
        "subtitle_italic": "Reinforcement Learning, Exploration & Exploitation Strategies, MuJoCo, OpenAI Gymnasium",
        "stack": "PyTorch, Anaconda, LDAS, Action Selection (Spectral-Approach), Denoising, SME Loss",
        "links": [
            {"label": "Code", "url": "https://github.com/josephchay/SADLDAS"},
            {"label": "Video Results", "url": "https://github.com/josephchay/SADLDAS/tree/main/results"},
            {"label": "Paper", "url": "https://github.com/josephchay/SADLDAS/blob/main/ARS-2425-Task2-20580363.pdf"},
        ],
        "extra_bullets": [
            "Environments solved within leaderboard defined thresholds:",
            "  - Mountain Car Continuous (Version 0)",
            "  - Lunar Lander Continuous (Version 3)",
            "  - Pendulum (Version 1)",
            "  - Cart Pole (Version 1)",
            "  - Bipedal Walker (Version 3)",
            "  - Ant (Version 5)",
            "  - Half Cheetah (Version 5)",
            "  - Hopper (Version 5)",
            "  - Pusher (Version 5)",
            "  - Walker 2D (Version 5)",
            "  - Humanoid (Version 5)",
        ],
    },
    {
        "uid": "pproj_wsp",
        "enabled": False,
        "section": "Personal Projects",
        "type": "project",
        "title": "Workflow Satisfiability Problem",
        "date": "Sep 2024 -- Dec 2024",
        "subtitle_italic": "Constraint Satisfaction, Tooling solvers, Encoding Patterns, Mathematical Formulations, Symbolic AI",
        "stack": "Genetic Evolutionary, Metaheuristic, Pigeonhole Principle, Bayesian Network",
        "links": [
            {"label": "Code", "url": "https://github.com/josephchay/wsp-symbolic-ai"},
            {"label": "Paper", "url": "https://github.com/josephchay/wsp-symbolic-ai/blob/main/documentation.pdf"},
        ],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_timetabler",
        "enabled": False,
        "section": "Personal Projects",
        "type": "project",
        "title": "Assessment Timetabling Scheduler",
        "date": "Sep 2024 -- Nov 2024",
        "subtitle_italic": "Constraint Satisfaction Programming, Multi-Solver Optimization, Symbolic AI",
        "stack": "SMT Solver, MILP Optimization, Evolutionary Algorithms, Metaheuristics",
        "links": [
            {"label": "Code", "url": "https://github.com/josephchay/assessment-timetabler"},
            {"label": "Paper", "url": "https://github.com/josephchay/assessment-timetabler/blob/main/documentation.pdf"},
        ],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_soccer",
        "enabled": False,
        "section": "Personal Projects",
        "type": "project",
        "title": "Soccer Analyzer, Player Performance, Ball Tracking and Prediction",
        "date": "May 2024",
        "subtitle_italic": "Computer Vision, Object Detection, Occlusion, Position and Speed Tracking, Clustering",
        "stack": "PyTorch, YOLOv5, OpenCV",
        "links": [{"label": "Code", "url": "https://github.com/josephchay/soccer-analysis-yolov5"}],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_sd",
        "enabled": False,
        "section": "Personal Projects",
        "type": "project",
        "title": "Stable Diffusion with DDPM Sampler, Image Generation Pipeline",
        "date": "Nov 2023 -- May 2024",
        "subtitle_italic": "Pipeline, Unconditional Generation, Text-to-Image, Image-to-Image",
        "stack": "PyTorch, DDPM, VAE, CLIP, Cross Attention, Diffusion",
        "links": [{"label": "Code", "url": "https://github.com/josephchay/stable-diffusion-pytorch"}],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_fridge",
        "enabled": False,
        "section": "Personal Projects",
        "type": "project",
        "title": "Smart Fridge, Fridge Surveillance and Meal Planner",
        "date": "Sep 2023 -- May 2024",
        "subtitle_italic": "Machine Learning, App Development, Computer Vision, Object Detection",
        "stack": "Pytorch, TensorFlow, Flutter, Dart, Swift, Kotlin, InnoDB",
        "links": [
            {"label": "AI Code", "url": "https://github.com/josephchay/smart-fridge-model-training"},
            {"label": "App Code and Sample Screens", "url": "https://github.com/josephchay/smart-fridge"},
        ],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_binpack",
        "enabled": True,
        "section": "Personal Projects",
        "type": "project",
        "title": "Bin Packing Problem, Single Dimension, NP-Hard",
        "date": "Apr 2024 -- May 2024",
        "subtitle_italic": "First-Fit Decreasing, Tabu Search, Genetic Meta-heuristic, Simulated Annealing",
        "stack": "Java, Heuristics, Meta-heuristics, Search Strategies",
        "links": [
            {"label": "Code", "url": "https://github.com/josephchay/1d-bp-java-4-algos"},
            {"label": "Paper", "url": "https://github.com/josephchay/1d-bp-java-4-algos/blob/main/conference.pdf"},
        ],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_garbage",
        "enabled": True,
        "section": "Personal Projects",
        "type": "project",
        "title": "Garbage Management System, Optimal Path Finder",
        "date": "Feb 2024 -- May 2024",
        "subtitle_italic": "Greedy Algorithm, Optimized Algorithms, Travelling Salesman Problem Algorithm",
        "stack": "C, C++, Cmake, Route Finding Algorithms, Database",
        "links": [
            {"label": "Code", "url": "https://github.com/josephchay/garbage-optimal-path-finder-cpp"},
            {"label": "Paper", "url": "https://github.com/josephchay/garbage-optimal-path-finder-cpp/blob/main/documentation.pdf"},
        ],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_blockchain",
        "enabled": False,
        "section": "Personal Projects",
        "type": "project",
        "title": "Inventory Management System, Blockchain Security Procedures",
        "date": "Jan 2024 -- Mar 2024",
        "subtitle_italic": "Blockchain, Supply Chain Management, Block Mining, User Role Management",
        "stack": "C, C++, Cmake, Block Header, Merkle Root, Nonce Setting, Block Mining, Hashing",
        "links": [
            {"label": "Code", "url": "https://github.com/josephchay/inventory-transportation-management"},
            {"label": "Paper", "url": "https://github.com/josephchay/inventory-transportation-management/blob/main/report-documentation.pdf"},
        ],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_jotit",
        "enabled": True,
        "section": "Personal Projects",
        "type": "project",
        "title": "JotIt, Interactive Note Taking and Diary Application",
        "date": "Feb 2024",
        "subtitle_italic": "Design, Development, Database, Deployment",
        "stack": "ReactJS, Vite, JavaScript, TailwindCSS, Framer Motion, GSAP",
        "links": [
            {"label": "Code", "url": "https://github.com/josephchay/jotit"},
            {"label": "Demo", "url": "https://jotitnow.netlify.app/"},
        ],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_brick",
        "enabled": True,
        "section": "Personal Projects",
        "type": "project",
        "title": "Bricksthetics, Interactive Aesthetic Brick-breaker",
        "date": "Dec 2023 -- Jan 2024",
        "subtitle_italic": "Game Development, Applied Mathematics, Micro Interactions",
        "stack": "Java, HTML, Maven, SFX, Particle Systems",
        "links": [
            {"label": "Code", "url": "https://github.com/josephchay/bricksthetics"},
            {"label": "Framework Design", "url": "https://github.com/josephchay/bricksthetics/blob/main/Design.pdf"},
        ],
        "extra_bullets": [],
    },
    {
        "uid": "pproj_selayang",
        "enabled": True,
        "section": "Personal Projects",
        "type": "project",
        "title": "Hospital Selayang Malaysia Datatics Dashboard",
        "date": "Sep 2022 -- Nov 2022",
        "subtitle_italic": "Data Visualization, Dashboard, Appointment Management System, Pagination, Email System",
        "stack": "HTML, CSS, JavaScript, PHP, MySQL, Bootstrap",
        "links": [{"label": "Code and Sample Screens", "url": "https://github.com/josephchay/hospital-selayang-datatics"}],
        "extra_bullets": [],
    },
    # PROFICIENCY AND TOOLS
    {
        "uid": "sk_algo",
        "enabled": True,
        "section": "Proficiency and Tools",
        "type": "skill_group",
        "category": "Algorithmic Programming Languages and Frameworks",
        "bullets": [
            "C, C++, Java, Python",
            "PyTorch, TensorFlow",
            "Anaconda, Jupyter, Google Colab, Matlab",
            "Matplotlib, Matlab, LaTeX",
        ],
    },
    {
        "uid": "sk_app",
        "enabled": True,
        "section": "Proficiency and Tools",
        "type": "skill_group",
        "category": "Application Development Languages and Frameworks",
        "bullets": [
            "Frontend (HTML, React, Next, Vue, TypeScript, JavaScript, Node, jQuery)",
            "Database Frameworks (PHP, MYSQL, Laravel)",
            "Stylings (CSS, SASS, Tailwind, Bootstrap)",
            "Animations (GSAP, Framer Motion, LocomotiveScroll, ScrollTrigger)",
        ],
    },
    {
        "uid": "sk_misc",
        "enabled": True,
        "section": "Proficiency and Tools",
        "type": "skill_group",
        "category": "Miscellaneous Tools and Frameworks",
        "bullets": [
            "Version Control Systems (GitHub, Bitbucket)",
            "Command-Line (Git, Bash, PowerShell)",
            "IDEs (VsCode, Intellij, IDLE)",
        ],
    },
    # LANGUAGES
    {
        "uid": "lang_spoken",
        "enabled": True,
        "section": "Languages",
        "type": "skill_group",
        "category": "Spoken and Written",
        "bullets": [
            "Very Proficient (English, Mandarin)",
            "Proficient (Cantonese, Malay)",
        ],
    },
]

BUILT_IN_PRESETS = {
    "AI / ML & Quantum Research (2-Page Target)": {
        "profile_idx": 1,
        "margin_cm": "1.5",
        "enabled_uids": {
            "edu_phd", "edu_bsc", "edu_dip",
            "exp_ips_swe", "exp_ips_ai", "exp_spec",
            "awd_fyp", "awd_ind",
            "cproj_yamaha",
            "pproj_linguify", "pproj_sadldas", "pproj_wsp", "pproj_timetabler", "pproj_soccer", "pproj_sd", "pproj_fridge",
            "sk_algo", "sk_app", "sk_misc", "lang_spoken",
        },
    },
    "Full-Stack & Software Engineering (2-Page Target)": {
        "profile_idx": 2,
        "margin_cm": "1.5",
        "enabled_uids": {
            "edu_bsc", "edu_dip",
            "exp_ips_swe", "exp_spec",
            "awd_ind", "awd_dhp",
            "cproj_vixta", "cproj_jessica",
            "pproj_docket", "pproj_codeflux", "pproj_garbage", "pproj_blockchain", "pproj_jotit", "pproj_selayang",
            "sk_algo", "sk_app", "sk_misc", "lang_spoken",
        },
    },
    "Systems, Operations & Commercial": {
        "profile_idx": 4,
        "margin_cm": "1.6",
        "enabled_uids": {
            "edu_bsc", "edu_dip",
            "exp_shopee", "exp_ips_swe", "exp_spec",
            "awd_dhp", "awd_hp", "awd_abrsm_pr",
            "pproj_fridge", "pproj_garbage", "pproj_selayang",
            "sk_app", "sk_misc", "lang_spoken",
        },
    },
}


# ==============================================================================
# Main Application Class
# ==============================================================================
class JosephResumeStudioV2:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Joseph Chay — LaTeX Resume Studio v2")
        self.root.geometry("1460x890")
        self.root.minsize(1140, 700)
        self.root.configure(bg=COLORS["bg_main"])

        self.undo_stack: list[str] = []
        self.redo_stack: list[str] = []

        self.layout = copy.deepcopy(DEFAULT_LAYOUT)
        self.header = copy.deepcopy(DEFAULT_HEADER)
        self.profiles = copy.deepcopy(DEFAULT_PROFILES)
        self.active_profile_idx = 0
        self.profile_enabled = True
        self.section_order = copy.deepcopy(DEFAULT_SECTION_ORDER)
        self.items = copy.deepcopy(DEFAULT_ITEMS)
        self.custom_presets = copy.deepcopy(BUILT_IN_PRESETS)

        self.filter_sec_var = tk.StringVar(value="All Sections")
        self.filter_state_var = tk.StringVar(value="All")
        self.search_var = tk.StringVar()
        self.preset_var = tk.StringVar(value="Select Tailoring Preset...")
        self.auto_save_var = tk.BooleanVar(value=True)
        self.auto_compile_var = tk.BooleanVar(value=False)

        self.visible_indices: list[int] = []
        self.item_line_map: dict[int, tuple[int, int]] = {}

        # PDF viewer state
        self.pdf_images: list[tk.PhotoImage] = []
        self.pdf_page_idx = 0
        self.pdf_zoom = 1.25

        self._load_state()
        self._push_undo()
        self._setup_styles()
        self._build_ui()
        self._bind_shortcuts()
        self.refresh_all(compile_pdf=False)
        self._inspect_log()
        if os.path.exists(PDF_FILE):
            self._render_pdf_pages()

    # ------------------------------------------------------------------
    # Undo / Redo Snapshot Engine
    # ------------------------------------------------------------------
    def _serialize(self) -> str:
        return json.dumps(
            {
                "layout": self.layout,
                "header": self.header,
                "profiles": self.profiles,
                "active_profile_idx": self.active_profile_idx,
                "profile_enabled": self.profile_enabled,
                "section_order": self.section_order,
                "items": self.items,
                "custom_presets": {
                    k: {
                        "profile_idx": v["profile_idx"],
                        "margin_cm": v["margin_cm"],
                        "enabled_uids": list(v["enabled_uids"]),
                    }
                    for k, v in self.custom_presets.items()
                },
            },
            ensure_ascii=False,
        )

    def _deserialize(self, raw: str):
        d = json.loads(raw)
        self.layout = d.get("layout", self.layout)
        self.header = d.get("header", self.header)
        self.profiles = d.get("profiles", self.profiles)
        self.active_profile_idx = d.get("active_profile_idx", self.active_profile_idx)
        self.profile_enabled = d.get("profile_enabled", self.profile_enabled)
        self.section_order = d.get("section_order", self.section_order)
        self.items = d.get("items", self.items)
        if "custom_presets" in d:
            self.custom_presets = {
                k: {
                    "profile_idx": v["profile_idx"],
                    "margin_cm": v["margin_cm"],
                    "enabled_uids": set(v["enabled_uids"]),
                }
                for k, v in d["custom_presets"].items()
            }

    def _push_undo(self):
        snap = self._serialize()
        if not self.undo_stack or self.undo_stack[-1] != snap:
            self.undo_stack.append(snap)
            if len(self.undo_stack) > 50:
                self.undo_stack.pop(0)
            self.redo_stack.clear()

    def undo(self, _e=None):
        if len(self.undo_stack) <= 1:
            return
        self.redo_stack.append(self.undo_stack.pop())
        self._deserialize(self.undo_stack[-1])
        self.refresh_all(compile_pdf=False)
        self.status_lbl.config(text="↶ Undid last change.")

    def redo(self, _e=None):
        if not self.redo_stack:
            return
        nxt = self.redo_stack.pop()
        self.undo_stack.append(nxt)
        self._deserialize(nxt)
        self.refresh_all(compile_pdf=False)
        self.status_lbl.config(text="↷ Redid change.")

    # ------------------------------------------------------------------
    # Styling & Button Factory
    # ------------------------------------------------------------------
    def _setup_styles(self):
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure(".", font=("Segoe UI", 10), background=COLORS["bg_main"], foreground=COLORS["fg_primary"])
        style.configure(
            "CV.Treeview",
            background=COLORS["bg_card"],
            fieldbackground=COLORS["bg_card"],
            foreground=COLORS["fg_primary"],
            rowheight=32,
            font=("Segoe UI", 10),
            borderwidth=0,
        )
        style.configure(
            "CV.Treeview.Heading",
            background="#e2e8f0",
            foreground="#0f172a",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padding=(8, 6),
        )
        style.map(
            "CV.Treeview",
            background=[("selected", "#dbeafe")],
            foreground=[("selected", "#003566")],
        )

    def _btn(self, parent, text: str, cmd, bg: str, hover: str, pad=(10, 5)) -> tk.Button:
        b = tk.Button(
            parent,
            text=text,
            command=cmd,
            bg=bg,
            fg="#ffffff",
            activebackground=hover,
            activeforeground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            relief=tk.FLAT,
            bd=0,
            padx=pad[0],
            pady=pad[1],
            cursor="hand2",
        )
        b.bind("<Enter>", lambda e: b.configure(bg=hover))
        b.bind("<Leave>", lambda e: b.configure(bg=bg))
        return b

    # ------------------------------------------------------------------
    # Build Main Workspace Layout
    # ------------------------------------------------------------------
    def _build_ui(self):
        # 1. Top Navy Banner
        banner = tk.Frame(self.root, bg=COLORS["bg_header"], padx=18, pady=11)
        banner.pack(fill=tk.X)

        left_info = tk.Frame(banner, bg=COLORS["bg_header"])
        left_info.pack(side=tk.LEFT)
        tk.Label(
            left_info,
            text="JOSEPH CHAY — LaTeX Resume Studio v2",
            font=("Segoe UI", 14, "bold"),
            bg=COLORS["bg_header"],
            fg="#ffffff",
        ).pack(anchor=tk.W)
        self.contact_banner_lbl = tk.Label(
            left_info,
            text="",
            font=("Segoe UI", 9),
            bg=COLORS["bg_header"],
            fg="#bae6fd",
        )
        self.contact_banner_lbl.pack(anchor=tk.W)

        right_btns = tk.Frame(banner, bg=COLORS["bg_header"])
        right_btns.pack(side=tk.RIGHT)
        self._btn(right_btns, "↶ Undo", self.undo, "#1e3a8a", "#1d4ed8", pad=(8, 5)).pack(side=tk.LEFT, padx=2)
        self._btn(right_btns, "↷ Redo", self.redo, "#1e3a8a", "#1d4ed8", pad=(8, 5)).pack(side=tk.LEFT, padx=(2, 8))
        self._btn(right_btns, "👤 Contact", self.open_contact_modal, "#1e40af", "#1d4ed8").pack(side=tk.LEFT, padx=3)
        self._btn(right_btns, "📝 Profile", self.open_profile_modal, "#1e40af", "#1d4ed8").pack(side=tk.LEFT, padx=3)
        self._btn(right_btns, "📐 Page-Fit", self.open_layout_modal, "#1e40af", "#1d4ed8").pack(side=tk.LEFT, padx=3)
        self._btn(right_btns, "📑 Sections", self.open_section_order_modal, "#1e40af", "#1d4ed8").pack(side=tk.LEFT, padx=3)
        self._btn(right_btns, "🎯 JD Matcher", self.open_jd_matcher_modal, "#6d28d9", "#5b21b6").pack(side=tk.LEFT, padx=3)
        self._btn(right_btns, "⚙ Compile PDF", self.compile_pdf, COLORS["emerald"], COLORS["emerald_hover"]).pack(side=tk.LEFT, padx=(8, 3))
        self._btn(right_btns, "📄 External PDF", self.open_pdf, "#0284c7", "#0369a1").pack(side=tk.LEFT, padx=3)

        # 2. Quick-Add & Role Tailoring Ribbon
        ribbon = tk.Frame(
            self.root, bg=COLORS["bg_card"], padx=14, pady=9, highlightthickness=1, highlightbackground=COLORS["border"]
        )
        ribbon.pack(fill=tk.X, padx=14, pady=(10, 6))

        tk.Label(
            ribbon, text="ADD:", font=("Segoe UI", 9, "bold"), bg=COLORS["bg_card"], fg=COLORS["fg_muted"]
        ).pack(side=tk.LEFT, padx=(0, 6))

        add_specs = [
            ("＋ Project", "project", COLORS["primary"], COLORS["primary_hover"]),
            ("＋ Experience", "experience", "#0284c7", "#0369a1"),
            ("＋ Education", "education", "#0d9488", "#0f766e"),
            ("＋ Award", "award", "#7c3aed", "#6d28d9"),
            ("＋ Skill Group", "skill_group", "#d97706", "#b45309"),
            ("＋ Custom Item", "custom", "#475569", "#334155"),
        ]
        for label, itype, bg_c, hov_c in add_specs:
            self._btn(ribbon, label, lambda t=itype: self.open_item_modal(t), bg_c, hov_c, pad=(9, 4)).pack(
                side=tk.LEFT, padx=2
            )

        # Right side of ribbon: Role Tailoring Preset Switcher
        preset_box = tk.Frame(ribbon, bg=COLORS["bg_card"])
        preset_box.pack(side=tk.RIGHT)
        tk.Label(
            preset_box, text="Tailor Preset:", font=("Segoe UI", 9, "bold"), bg=COLORS["bg_card"], fg=COLORS["fg_muted"]
        ).pack(side=tk.LEFT, padx=(0, 6))

        self.preset_combo = ttk.Combobox(
            preset_box,
            textvariable=self.preset_var,
            values=list(self.custom_presets.keys()),
            width=34,
            state="readonly",
        )
        self.preset_combo.pack(side=tk.LEFT, padx=2)
        self.preset_combo.bind("<<ComboboxSelected>>", self.apply_tailoring_preset)
        self._btn(preset_box, "+ Save Preset", self.save_new_preset_dialog, "#475569", "#334155", pad=(8, 4)).pack(
            side=tk.LEFT, padx=(4, 0)
        )

        # 3. Main Split Pane (Left: Multi-Select Table | Right: Tabbed LaTeX + Visual PDF Preview)
        paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=14, pady=6)

        # --- LEFT PANEL ---
        left_card = tk.Frame(
            paned, bg=COLORS["bg_card"], highlightthickness=1, highlightbackground=COLORS["border"]
        )
        paned.add(left_card, weight=5)

        filter_bar = tk.Frame(left_card, bg=COLORS["bg_card"], padx=12, pady=9)
        filter_bar.pack(fill=tk.X)

        tk.Label(filter_bar, text="Section:", bg=COLORS["bg_card"], fg=COLORS["fg_muted"], font=("Segoe UI", 9, "bold")).pack(
            side=tk.LEFT
        )
        self.sec_filter_combo = ttk.Combobox(filter_bar, textvariable=self.filter_sec_var, width=20, state="readonly")
        self.sec_filter_combo.pack(side=tk.LEFT, padx=(5, 10))
        self.sec_filter_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_all(compile_pdf=False))

        tk.Label(filter_bar, text="Status:", bg=COLORS["bg_card"], fg=COLORS["fg_muted"], font=("Segoe UI", 9, "bold")).pack(
            side=tk.LEFT
        )
        state_combo = ttk.Combobox(
            filter_bar,
            textvariable=self.filter_state_var,
            values=["All", "Active Only (✓)", "Archived / Commented (✗)"],
            width=21,
            state="readonly",
        )
        state_combo.pack(side=tk.LEFT, padx=(5, 10))
        state_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_all(compile_pdf=False))

        tk.Label(filter_bar, text="Search:", bg=COLORS["bg_card"], fg=COLORS["fg_muted"], font=("Segoe UI", 9, "bold")).pack(
            side=tk.LEFT
        )
        ttk.Entry(filter_bar, textvariable=self.search_var, width=18).pack(side=tk.LEFT, padx=5)
        self.search_var.trace_add("write", lambda *_: self.refresh_all(compile_pdf=False))

        self.stats_lbl = tk.Label(filter_bar, text="", bg=COLORS["bg_card"], fg=COLORS["fg_muted"], font=("Segoe UI", 9, "bold"))
        self.stats_lbl.pack(side=tk.RIGHT)

        tree_wrap = tk.Frame(left_card, bg=COLORS["bg_card"], padx=12)
        tree_wrap.pack(fill=tk.BOTH, expand=True)

        cols = ("active", "section", "title", "detail", "date")
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="headings", style="CV.Treeview", selectmode="extended")
        self.tree.heading("active", text="On")
        self.tree.heading("section", text="Section")
        self.tree.heading("title", text="Title / Institution / Project")
        self.tree.heading("detail", text="Role / Degree / Tech Stack")
        self.tree.heading("date", text="Date")

        self.tree.column("active", width=44, anchor=tk.CENTER, stretch=False)
        self.tree.column("section", width=130, anchor=tk.W, stretch=False)
        self.tree.column("title", width=225, anchor=tk.W, stretch=True)
        self.tree.column("detail", width=195, anchor=tk.W, stretch=True)
        self.tree.column("date", width=115, anchor=tk.E, stretch=False)

        self.tree.tag_configure("on", foreground="#0f172a")
        self.tree.tag_configure("off", foreground="#94a3b8")

        yscroll = ttk.Scrollbar(tree_wrap, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())
        self.tree.bind("<space>", lambda e: self.toggle_selected())

        # Table Action Bar
        tbl_bar = tk.Frame(left_card, bg="#f8fafc", padx=12, pady=9)
        tbl_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self._btn(tbl_bar, "✏ Edit", self.edit_selected, COLORS["primary"], COLORS["primary_hover"], pad=(10, 4)).pack(
            side=tk.LEFT, padx=2
        )
        self._btn(tbl_bar, "✓/✗ Toggle (Space)", self.toggle_selected, "#475569", "#334155", pad=(9, 4)).pack(
            side=tk.LEFT, padx=2
        )
        self._btn(tbl_bar, "⧉ Clone", self.clone_selected, "#475569", "#334155", pad=(9, 4)).pack(side=tk.LEFT, padx=2)
        self._btn(tbl_bar, "▲ Up", lambda: self.move_selected(-1), "#64748b", "#475569", pad=(8, 4)).pack(side=tk.LEFT, padx=2)
        self._btn(tbl_bar, "▼ Down", lambda: self.move_selected(1), "#64748b", "#475569", pad=(8, 4)).pack(side=tk.LEFT, padx=2)
        self._btn(tbl_bar, "🗑 Delete", self.delete_selected, COLORS["danger"], "#b91c1c", pad=(9, 4)).pack(side=tk.RIGHT, padx=2)

        # --- RIGHT PANEL: Tabbed LaTeX Source + Visual PDF Viewer ---
        right_card = tk.Frame(
            paned, bg=COLORS["bg_card"], highlightthickness=1, highlightbackground=COLORS["border"]
        )
        paned.add(right_card, weight=4)

        self.preview_nb = ttk.Notebook(right_card)
        self.preview_nb.pack(fill=tk.BOTH, expand=True)

        # Tab A: LaTeX Code
        code_tab = tk.Frame(self.preview_nb, bg=COLORS["code_bg"])
        self.preview_nb.add(code_tab, text="<> Generated resume.tex")

        code_top = tk.Frame(code_tab, bg="#181825", padx=12, pady=7)
        code_top.pack(fill=tk.X)
        tk.Label(
            code_top, text="LIVE RENDERCY SOURCE", font=("Segoe UI", 9, "bold"), bg="#181825", fg="#a6adc8"
        ).pack(side=tk.LEFT)
        self._btn(code_top, "💾 Save As...", self.save_tex_as, "#313244", "#45475a", pad=(8, 3)).pack(side=tk.RIGHT, padx=2)
        self._btn(code_top, "📋 Copy Code", self.copy_tex, "#313244", "#45475a", pad=(8, 3)).pack(side=tk.RIGHT, padx=2)

        code_wrap = tk.Frame(code_tab, bg=COLORS["code_bg"])
        code_wrap.pack(fill=tk.BOTH, expand=True)

        self.code_text = tk.Text(
            code_wrap,
            font=("Consolas", 9),
            bg=COLORS["code_bg"],
            fg=COLORS["code_fg"],
            insertbackground="#ffffff",
            selectbackground="#45475a",
            wrap=tk.NONE,
            bd=0,
            padx=10,
            pady=8,
        )
        cy = ttk.Scrollbar(code_wrap, orient=tk.VERTICAL, command=self.code_text.yview)
        cx = ttk.Scrollbar(code_tab, orient=tk.HORIZONTAL, command=self.code_text.xview)
        self.code_text.configure(yscrollcommand=cy.set, xscrollcommand=cx.set)
        self.code_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        cy.pack(side=tk.RIGHT, fill=tk.Y)
        cx.pack(fill=tk.X, side=tk.BOTTOM)

        self.code_text.tag_configure("cmd", foreground="#89b4fa")
        self.code_text.tag_configure("section", foreground="#f38ba8", font=("Consolas", 9, "bold"))
        self.code_text.tag_configure("env", foreground="#a6e3a1")
        self.code_text.tag_configure("math", foreground="#f9e2af")
        self.code_text.tag_configure("comment", foreground="#6c7086")
        self.code_text.tag_configure("active_block", background=COLORS["hl_line"])

        # Tab B: Visual Rendered PDF Preview
        pdf_tab = tk.Frame(self.preview_nb, bg="#cbd5e1")
        self.preview_nb.add(pdf_tab, text="📄 Visual PDF Preview")

        pdf_bar = tk.Frame(pdf_tab, bg="#e2e8f0", padx=10, pady=6)
        pdf_bar.pack(fill=tk.X)
        self._btn(pdf_bar, "◀ Prev", lambda: self._change_pdf_page(-1), "#475569", "#334155", pad=(8, 2)).pack(side=tk.LEFT, padx=2)
        self.pdf_page_lbl = tk.Label(pdf_bar, text="Page 1 / 1", bg="#e2e8f0", font=("Segoe UI", 9, "bold"), width=12)
        self.pdf_page_lbl.pack(side=tk.LEFT, padx=4)
        self._btn(pdf_bar, "Next ▶", lambda: self._change_pdf_page(1), "#475569", "#334155", pad=(8, 2)).pack(side=tk.LEFT, padx=2)

        self._btn(pdf_bar, "🔄 Refresh PDF", self.compile_pdf, COLORS["emerald"], COLORS["emerald_hover"], pad=(8, 2)).pack(
            side=tk.RIGHT, padx=2
        )
        self._btn(pdf_bar, "Zoom +", lambda: self._zoom_pdf(0.2), "#475569", "#334155", pad=(7, 2)).pack(side=tk.RIGHT, padx=2)
        self._btn(pdf_bar, "Zoom -", lambda: self._zoom_pdf(-0.2), "#475569", "#334155", pad=(7, 2)).pack(side=tk.RIGHT, padx=2)

        canvas_wrap = tk.Frame(pdf_tab, bg="#94a3b8")
        canvas_wrap.pack(fill=tk.BOTH, expand=True)
        self.pdf_canvas = tk.Canvas(canvas_wrap, bg="#94a3b8", highlightthickness=0)
        py_scroll = ttk.Scrollbar(canvas_wrap, orient=tk.VERTICAL, command=self.pdf_canvas.yview)
        px_scroll = ttk.Scrollbar(pdf_tab, orient=tk.HORIZONTAL, command=self.pdf_canvas.xview)
        self.pdf_canvas.configure(yscrollcommand=py_scroll.set, xscrollcommand=px_scroll.set)
        self.pdf_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        py_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        px_scroll.pack(fill=tk.X, side=tk.BOTTOM)

        # 4. Bottom Status Bar
        status_bar = tk.Frame(self.root, bg="#e2e8f0", padx=14, pady=6)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        ttk.Checkbutton(status_bar, text="Auto-save resume.tex", variable=self.auto_save_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(status_bar, text="Auto-compile PDF on save", variable=self.auto_compile_var).pack(side=tk.LEFT, padx=(0, 16))

        self.status_lbl = tk.Label(status_bar, text="Ready.", bg="#e2e8f0", fg="#334155", font=("Segoe UI", 9))
        self.status_lbl.pack(side=tk.LEFT)

        self.log_badge = tk.Label(
            status_bar, text="resume.log: Checking...", bg="#cbd5e1", fg="#0f172a", font=("Segoe UI", 9, "bold"), padx=10, pady=2
        )
        self.log_badge.pack(side=tk.RIGHT)

    def _bind_shortcuts(self):
        self.root.bind("<Control-s>", lambda e: self.refresh_all(compile_pdf=True))
        self.root.bind("<Control-z>", self.undo)
        self.root.bind("<Control-y>", self.redo)
        self.root.bind("<Control-n>", lambda e: self.open_item_modal("project"))

    # ------------------------------------------------------------------
    # Tailoring Presets
    # ------------------------------------------------------------------
    def apply_tailoring_preset(self, _e=None):
        pname = self.preset_var.get()
        if pname not in self.custom_presets:
            return
        self._push_undo()
        cfg = self.custom_presets[pname]
        self.active_profile_idx = min(cfg["profile_idx"], len(self.profiles) - 1)
        self.profile_enabled = True
        self.layout["margin_cm"] = cfg.get("margin_cm", self.layout["margin_cm"])
        target_uids = cfg["enabled_uids"]
        for it in self.items:
            it["enabled"] = it.get("uid") in target_uids
        self.refresh_all(compile_pdf=self.auto_compile_var.get())
        self.status_lbl.config(text=f"Applied tailoring preset '{pname}'.")

    def save_new_preset_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Save Current Selection as Tailoring Preset")
        dlg.geometry("380x140")
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text="Preset Name:", bg=COLORS["bg_card"], font=("Segoe UI", 9, "bold")).pack(pady=(14, 4))
        nvar = tk.StringVar(value="Custom Role Preset")
        ttk.Entry(dlg, textvariable=nvar, width=36).pack(pady=4)

        def save():
            name = nvar.get().strip()
            if name:
                self._push_undo()
                self.custom_presets[name] = {
                    "profile_idx": self.active_profile_idx,
                    "margin_cm": self.layout["margin_cm"],
                    "enabled_uids": {it.get("uid") for it in self.items if it.get("enabled", True)},
                }
                self.preset_combo["values"] = list(self.custom_presets.keys())
                self.preset_var.set(name)
                self._save_state()
                self.status_lbl.config(text=f"Saved preset '{name}'.")
            dlg.destroy()

        self._btn(dlg, "✓ Save Preset", save, COLORS["emerald"], COLORS["emerald_hover"]).pack(pady=8)

    # ------------------------------------------------------------------
    # Page-Fit & Layout Tuner Modal
    # ------------------------------------------------------------------
    def open_layout_modal(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Page-Fit & Visual Geometry Tuner")
        dlg.geometry("500x430")
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)
        dlg.grab_set()

        top = tk.Frame(dlg, bg=COLORS["bg_header"], padx=16, pady=10)
        top.pack(fill=tk.X)
        tk.Label(top, text="📐 Page-Fit & LaTeX Geometry Tuner", font=("Segoe UI", 11, "bold"), bg=COLORS["bg_header"], fg="#fff").pack(side=tk.LEFT)

        body = tk.Frame(dlg, bg=COLORS["bg_card"], padx=18, pady=12)
        body.pack(fill=tk.BOTH, expand=True)
        body.columnconfigure(1, weight=1)

        # Quick Margin Presets
        p_row = tk.Frame(body, bg=COLORS["bg_card"])
        p_row.grid(row=0, column=0, columnspan=2, sticky=tk.EW, pady=(0, 10))

        v_margin = tk.StringVar(value=self.layout["margin_cm"])
        v_rcol = tk.StringVar(value=self.layout["right_col_cm"])
        v_sb = tk.StringVar(value=self.layout["sec_before_cm"])
        v_sa = tk.StringVar(value=self.layout["sec_after_cm"])
        v_glyph = tk.StringVar(value=self.layout["bullet_glyph"])
        v_rgb = tk.StringVar(value=self.layout["primary_rgb"])
        v_lastpage = tk.BooleanVar(value=self.layout["footer_page_of_last"])

        def set_preset(m, sb, sa):
            v_margin.set(m)
            v_sb.set(sb)
            v_sa.set(sa)

        self._btn(p_row, "Original (2.0cm)", lambda: set_preset("2.0", "0.3", "0.2"), "#475569", "#334155", pad=(8, 3)).pack(side=tk.LEFT, padx=2)
        self._btn(p_row, "2-Page Fit (1.4cm)", lambda: set_preset("1.4", "0.22", "0.14"), COLORS["primary"], COLORS["primary_hover"], pad=(8, 3)).pack(side=tk.LEFT, padx=2)
        self._btn(p_row, "1-Page Tight (1.1cm)", lambda: set_preset("1.1", "0.15", "0.08"), "#0d9488", "#0f766e", pad=(8, 3)).pack(side=tk.LEFT, padx=2)

        rows = [
            ("Page Margins (cm):", v_margin),
            ("Date Column Width (cm):", v_rcol),
            ("Section Space Before (cm):", v_sb),
            ("Section Space After (cm):", v_sa),
            ("Bullet Glyph (\\circ / \\bullet):", v_glyph),
            ("Primary RGB Color:", v_rgb),
        ]
        for idx, (lbl, var) in enumerate(rows, start=1):
            tk.Label(body, text=lbl, bg=COLORS["bg_card"], font=("Segoe UI", 9, "bold")).grid(row=idx, column=0, sticky=tk.W, pady=5)
            ttk.Entry(body, textvariable=var, width=22).grid(row=idx, column=1, sticky=tk.EW, pady=5)

        def pick_rgb():
            c = colorchooser.askcolor()
            if c and c[0]:
                v_rgb.set(f"{int(c[0][0])}, {int(c[0][1])}, {int(c[0][2])}")

        self._btn(body, "🎨 Pick RGB Color", pick_rgb, "#475569", "#334155", pad=(8, 3)).grid(row=7, column=0, sticky=tk.W, pady=6)
        ttk.Checkbutton(body, text="Footer: 'Page X of Y'", variable=v_lastpage).grid(row=7, column=1, sticky=tk.W, pady=6)

        def apply_layout():
            self._push_undo()
            self.layout["margin_cm"] = v_margin.get().strip()
            self.layout["right_col_cm"] = v_rcol.get().strip()
            self.layout["sec_before_cm"] = v_sb.get().strip()
            self.layout["sec_after_cm"] = v_sa.get().strip()
            self.layout["bullet_glyph"] = v_glyph.get().strip()
            self.layout["primary_rgb"] = v_rgb.get().strip()
            self.layout["footer_page_of_last"] = v_lastpage.get()
            dlg.destroy()
            self.refresh_all(compile_pdf=self.auto_compile_var.get())

        self._btn(body, "✓ Apply Layout", apply_layout, COLORS["emerald"], COLORS["emerald_hover"]).grid(
            row=8, column=1, sticky=tk.E, pady=(12, 0)
        )

    # ------------------------------------------------------------------
    # Job Description Keyword Matcher & Archived Item Finder
    # ------------------------------------------------------------------
    def open_jd_matcher_modal(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Job Description Keyword Matcher & Archive Finder")
        dlg.geometry("760x580")
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)

        top = tk.Frame(dlg, bg=COLORS["bg_header"], padx=16, pady=10)
        top.pack(fill=tk.X)
        tk.Label(top, text="🎯 JD Keyword Matcher (Finds Relevant Archived / Disabled Projects)", font=("Segoe UI", 11, "bold"), bg=COLORS["bg_header"], fg="#fff").pack(side=tk.LEFT)

        body = tk.Frame(dlg, bg=COLORS["bg_card"], padx=16, pady=12)
        body.pack(fill=tk.BOTH, expand=True)

        tk.Label(body, text="Paste Target Job Description:", bg=COLORS["bg_card"], font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        jd_box = tk.Text(body, height=7, font=("Segoe UI", 9), bg="#f8fafc", relief=tk.SOLID, bd=1)
        jd_box.pack(fill=tk.X, pady=(4, 8))

        out_box = tk.Text(body, height=13, font=("Consolas", 9), bg="#f8fafc", relief=tk.SOLID, bd=1)
        out_box.pack(fill=tk.BOTH, expand=True, pady=(4, 8))

        matched_disabled_items: list[dict] = []

        def analyze():
            jd = jd_box.get("1.0", tk.END).strip()
            if not jd:
                return
            stop = {
                "the", "and", "for", "with", "you", "will", "are", "that", "this", "from", "have",
                "our", "your", "work", "working", "team", "role", "experience", "years", "ability",
                "strong", "skills", "knowledge", "using", "build", "develop", "design", "degree",
                "computer", "science", "engineering", "related", "field", "plus", "preferred",
            }
            kws = sorted({
                tok.strip(".-/")
                for tok in re.findall(r"[A-Za-z][A-Za-z0-9\+\#\.\/\-]{2,}", jd)
                if len(tok) >= 3 and tok.lower() not in stop and not tok.isdigit()
            })

            active_blob = self.generate_latex().lower()
            matched = [k for k in kws if k.lower() in active_blob]
            missing = [k for k in kws if k.lower() not in active_blob]

            matched_disabled_items.clear()
            for it in self.items:
                if not it.get("enabled", True):
                    it_blob = json.dumps(it).lower()
                    hits = [m for m in missing if m.lower() in it_blob]
                    if hits:
                        matched_disabled_items.append((it, hits))

            lines = [
                f"Active CV Keyword Match: {len(matched)} / {len(kws)} terms",
                f"\n[✓] MATCHED IN ACTIVE CV:\n  {', '.join(matched[:45])}",
                f"\n[★] DISABLED ITEMS IN YOUR ARCHIVE THAT MATCH MISSING KEYWORDS ({len(matched_disabled_items)}):",
            ]
            for it, hits in matched_disabled_items:
                lbl = it.get("title") or it.get("role") or it.get("category")
                lines.append(f"  • [{it['section']}] {lbl}  --> Matches: {', '.join(hits)}")
            lines.append(f"\n[✗] NOT FOUND ANYWHERE IN ARCHIVE:\n  {', '.join([m for m in missing if not any(m in h for _, h in matched_disabled_items)][:40])}")

            out_box.delete("1.0", tk.END)
            out_box.insert("1.0", "\n".join(lines))

        def enable_matched_archived():
            if not matched_disabled_items:
                return
            self._push_undo()
            for it, _ in matched_disabled_items:
                it["enabled"] = True
            self.refresh_all(compile_pdf=self.auto_compile_var.get())
            analyze()
            self.status_lbl.config(text=f"Enabled {len(matched_disabled_items)} matching archived item(s).")

        btn_row = tk.Frame(body, bg=COLORS["bg_card"])
        btn_row.pack(fill=tk.X)
        self._btn(btn_row, "🔍 Analyze JD", analyze, COLORS["primary"], COLORS["primary_hover"]).pack(side=tk.LEFT, padx=3)
        self._btn(btn_row, "✓ Enable All Matching Archived Items", enable_matched_archived, COLORS["emerald"], COLORS["emerald_hover"]).pack(
            side=tk.RIGHT, padx=3
        )

    # ------------------------------------------------------------------
    # Tailored Add / Edit Item Modal with LaTeX Toolbar
    # ------------------------------------------------------------------
    def open_item_modal(self, item_type: str, edit_index: int | None = None):
        existing = self.items[edit_index] if edit_index is not None else {}

        dlg = tk.Toplevel(self.root)
        dlg.title(f"{'Edit' if edit_index is not None else 'Add'} {item_type.replace('_', ' ').title()}")
        dlg.geometry("740x640")
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)
        dlg.grab_set()

        top = tk.Frame(dlg, bg=COLORS["bg_header"], padx=18, pady=12)
        top.pack(fill=tk.X)
        tk.Label(
            top,
            text=f"{'✏ Edit' if edit_index is not None else '➕ Add'} {item_type.replace('_', ' ').title()}",
            font=("Segoe UI", 12, "bold"),
            bg=COLORS["bg_header"],
            fg="#ffffff",
        ).pack(side=tk.LEFT)

        body = tk.Frame(dlg, bg=COLORS["bg_card"], padx=20, pady=14)
        body.pack(fill=tk.BOTH, expand=True)
        body.columnconfigure(1, weight=1)

        vars_map: dict[str, tk.StringVar] = {}
        bool_vars: dict[str, tk.BooleanVar] = {}
        all_secs = list(dict.fromkeys(self.section_order + [i.get("section", "Other") for i in self.items]))

        def add_entry(row: int, label: str, key: str, default: str = "", combo: list[str] | None = None):
            tk.Label(
                body, text=label, bg=COLORS["bg_card"], fg=COLORS["fg_primary"], font=("Segoe UI", 9, "bold"), width=26, anchor=tk.W
            ).grid(row=row, column=0, sticky=tk.W, pady=4)
            v = tk.StringVar(value=existing.get(key, default))
            vars_map[key] = v
            if combo:
                w = ttk.Combobox(body, textvariable=v, values=combo, font=("Segoe UI", 10))
            else:
                w = ttk.Entry(body, textvariable=v, font=("Segoe UI", 10))
            w.grid(row=row, column=1, sticky=tk.EW, pady=4)

        links_list: list[dict[str, str]] = copy.deepcopy(existing.get("links", []))
        bullets_box: tk.Text | None = None

        def build_latex_helper_bar(parent_row: int, target_text_fn):
            bar = tk.Frame(body, bg=COLORS["bg_card"])
            bar.grid(row=parent_row, column=0, columnspan=2, sticky=tk.EW, pady=(6, 2))
            tk.Label(bar, text="Bullets (indent with '  - ' for nested sub-items):", bg=COLORS["bg_card"], font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)

            def wrap(pfx, sfx):
                tb = target_text_fn()
                if not tb:
                    return
                try:
                    s, e = tb.index(tk.SEL_FIRST), tb.index(tk.SEL_LAST)
                    sel = tb.get(s, e)
                    tb.delete(s, e)
                    tb.insert(s, f"{pfx}{sel}{sfx}")
                except tk.TclError:
                    tb.insert(tk.INSERT, f"{pfx}{sfx}")

            for lbl, p, s in [
                ("Bold", r"\textbf{", "}"),
                ("Italic", r"\textit{", "}"),
                ("$Math$", "$", "$"),
                ("Superscript", r"\textsuperscript{", "}"),
                ("Footnote", r"\footnote{", "}"),
                ("Link", r"\hrefWithoutArrow{https://}{", "}"),
            ]:
                tk.Button(
                    bar, text=lbl, command=lambda x=p, y=s: wrap(x, y),
                    bg="#e2e8f0", fg="#0f172a", relief=tk.FLAT, font=("Segoe UI", 8, "bold"), padx=6, pady=2, cursor="hand2"
                ).pack(side=tk.RIGHT, padx=2)

        if item_type == "project":
            add_entry(0, "Section Heading:", "section", "Personal Projects", all_secs)
            add_entry(1, "Bold Project Title:", "title", "Project Name, Short Tagline")
            add_entry(2, "Date Range (Italicized):", "date", "Feb 2026 -- Nov 2026")
            add_entry(3, "Italicized Domain/Subtitle:", "subtitle_italic", "Design, Development, Machine Learning, Deployment")
            add_entry(4, "Tech Stack (\\item Stack:):", "stack", "Python, PyTorch, ReactJS")

            link_frame = tk.LabelFrame(
                body, text="Access Links (\\item Access: \\hrefWithoutArrow{url}{Label})", bg=COLORS["bg_card"], padx=8, pady=6
            )
            link_frame.grid(row=5, column=0, columnspan=2, sticky=tk.EW, pady=6)

            links_listbox = tk.Listbox(link_frame, height=3, font=("Consolas", 9))
            links_listbox.pack(fill=tk.X, pady=(0, 4))

            def refresh_links_ui():
                links_listbox.delete(0, tk.END)
                for lk in links_list:
                    links_listbox.insert(tk.END, f"[{lk['label']}]  ->  {lk['url']}")

            refresh_links_ui()

            lk_ctrl = tk.Frame(link_frame, bg=COLORS["bg_card"])
            lk_ctrl.pack(fill=tk.X)
            lk_lbl_var = tk.StringVar(value="Code")
            lk_url_var = tk.StringVar(value="https://github.com/josephchay/")
            ttk.Entry(lk_ctrl, textvariable=lk_lbl_var, width=18).pack(side=tk.LEFT, padx=(0, 4))
            ttk.Entry(lk_ctrl, textvariable=lk_url_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

            def add_link():
                if lk_lbl_var.get().strip() and lk_url_var.get().strip():
                    links_list.append({"label": lk_lbl_var.get().strip(), "url": lk_url_var.get().strip()})
                    refresh_links_ui()

            def del_link():
                sel = links_listbox.curselection()
                if sel:
                    del links_list[sel[0]]
                    refresh_links_ui()

            self._btn(lk_ctrl, "+ Link", add_link, COLORS["primary"], COLORS["primary_hover"], pad=(8, 2)).pack(side=tk.LEFT, padx=2)
            self._btn(lk_ctrl, "Remove", del_link, COLORS["danger"], "#b91c1c", pad=(8, 2)).pack(side=tk.LEFT)

            build_latex_helper_bar(6, lambda: bullets_box)
            bullets_box = tk.Text(body, height=5, font=("Segoe UI", 9), bg="#f8fafc", relief=tk.SOLID, bd=1)
            bullets_box.grid(row=7, column=0, columnspan=2, sticky=tk.NSEW)
            body.rowconfigure(7, weight=1)
            if existing.get("extra_bullets"):
                bullets_box.insert("1.0", "\n".join(existing["extra_bullets"]))

        elif item_type == "experience":
            add_entry(0, "Section Heading:", "section", "Industrial Experience", all_secs)
            add_entry(1, "Company / Organization:", "company", "")
            add_entry(2, "Job Role & Location (Italic):", "role", "Software Engineer (Kuala Lumpur, Malaysia)")
            add_entry(3, "Date Range:", "date", "Sep 2026 -- Present")
            add_entry(4, "Vertical Space After (cm):", "space_after", "0.4cm")

            sub_v = tk.BooleanVar(value=existing.get("is_sub_role", False))
            bool_vars["is_sub_role"] = sub_v
            ttk.Checkbutton(
                body, text="Sub-role under preceding company (omit repeating company header)", variable=sub_v
            ).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=4)

            build_latex_helper_bar(6, lambda: bullets_box)
            bullets_box = tk.Text(body, height=8, font=("Segoe UI", 9), bg="#f8fafc", relief=tk.SOLID, bd=1)
            bullets_box.grid(row=7, column=0, columnspan=2, sticky=tk.NSEW)
            body.rowconfigure(7, weight=1)
            if existing.get("bullets"):
                bullets_box.insert("1.0", "\n".join(existing["bullets"]))

        elif item_type == "education":
            add_entry(0, "Section Heading:", "section", "Education", all_secs)
            add_entry(1, "Institution / University:", "institution", "University of Nottingham")
            add_entry(2, "Degree / Qualification:", "degree", "")
            add_entry(3, "Date Range:", "date", "Sep 2025 -- Jul 2026")

            sub_d = tk.BooleanVar(value=existing.get("is_sub_degree", False))
            bool_vars["is_sub_degree"] = sub_d
            ttk.Checkbutton(
                body, text="Additional degree under same university above (omit repeating university title)", variable=sub_d
            ).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=4)

            build_latex_helper_bar(5, lambda: bullets_box)
            bullets_box = tk.Text(body, height=8, font=("Segoe UI", 9), bg="#f8fafc", relief=tk.SOLID, bd=1)
            bullets_box.grid(row=6, column=0, columnspan=2, sticky=tk.NSEW)
            body.rowconfigure(6, weight=1)
            if existing.get("bullets"):
                bullets_box.insert("1.0", "\n".join(existing["bullets"]))

        elif item_type == "award":
            add_entry(0, "Section Heading:", "section", "Honours & Awards", all_secs)
            add_entry(1, "Bold Award & Issuer Title:", "title", "Award Name, Issuing Organization")
            add_entry(2, "Date (Right Column):", "date", "Nov 2026")

        elif item_type == "skill_group":
            add_entry(0, "Section Heading:", "section", "Proficiency and Tools", all_secs)
            add_entry(1, "Bold Category Heading:", "category", "Quantum & AI Frameworks")
            build_latex_helper_bar(2, lambda: bullets_box)
            bullets_box = tk.Text(body, height=10, font=("Segoe UI", 9), bg="#f8fafc", relief=tk.SOLID, bd=1)
            bullets_box.grid(row=3, column=0, columnspan=2, sticky=tk.NSEW)
            body.rowconfigure(3, weight=1)
            if existing.get("bullets"):
                bullets_box.insert("1.0", "\n".join(existing["bullets"]))

        else:  # custom section item
            add_entry(0, "New or Existing Section Name:", "section", "Publications & Research", all_secs)
            add_entry(1, "Bold Primary Title:", "title", "")
            add_entry(2, "Italicized Subtitle (Optional):", "subtitle_italic", "")
            add_entry(3, "Date / Location (Right Column):", "date", "2026")
            build_latex_helper_bar(4, lambda: bullets_box)
            bullets_box = tk.Text(body, height=8, font=("Segoe UI", 9), bg="#f8fafc", relief=tk.SOLID, bd=1)
            bullets_box.grid(row=5, column=0, columnspan=2, sticky=tk.NSEW)
            body.rowconfigure(5, weight=1)
            if existing.get("bullets"):
                bullets_box.insert("1.0", "\n".join(existing["bullets"]))

        footer = tk.Frame(dlg, bg="#f8fafc", padx=20, pady=12)
        footer.pack(fill=tk.X, side=tk.BOTTOM)

        def save_modal():
            self._push_undo()
            new_item = {
                "uid": existing.get("uid", f"item_{len(self.items) + 1}_{os.urandom(2).hex()}"),
                "enabled": existing.get("enabled", True),
                "type": item_type,
            }
            for k, v in vars_map.items():
                new_item[k] = v.get().strip()
            for k, bv in bool_vars.items():
                new_item[k] = bv.get()

            raw_lines = [ln.rstrip() for ln in bullets_box.get("1.0", tk.END).splitlines() if ln.strip()] if bullets_box else []

            if item_type == "project":
                new_item["links"] = links_list
                new_item["extra_bullets"] = raw_lines
            elif item_type == "education":
                new_item["links"] = existing.get("links", [])
                new_item["bullets"] = raw_lines
            elif bullets_box is not None:
                new_item["bullets"] = raw_lines

            sec = new_item.get("section") or "Other"
            if sec not in self.section_order:
                self.section_order.append(sec)

            if edit_index is not None:
                self.items[edit_index] = new_item
            else:
                self.items.append(new_item)

            dlg.destroy()
            self.refresh_all(compile_pdf=self.auto_compile_var.get())

        self._btn(footer, "✓ Save & Generate resume.tex", save_modal, COLORS["emerald"], COLORS["emerald_hover"], pad=(14, 6)).pack(
            side=tk.RIGHT, padx=4
        )
        self._btn(footer, "Cancel", dlg.destroy, "#64748b", "#475569", pad=(12, 6)).pack(side=tk.RIGHT, padx=4)

    # ------------------------------------------------------------------
    # Contact, Profile Summary & Section Order Modals
    # ------------------------------------------------------------------
    def open_contact_modal(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Edit Contact Header & Footer")
        dlg.geometry("540x380")
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)
        dlg.grab_set()

        form = tk.Frame(dlg, bg=COLORS["bg_card"], padx=20, pady=16)
        form.pack(fill=tk.BOTH, expand=True)
        form.columnconfigure(1, weight=1)

        fields = [
            ("Header Display Name:", "display_name"),
            ("Page Footer Full Name:", "footer_name"),
            ("Phone Number:", "phone"),
            ("Email Address:", "email"),
            ("LinkedIn Full URL:", "linkedin_url"),
            ("LinkedIn Display Text:", "linkedin_disp"),
            ("GitHub Full URL:", "github_url"),
            ("GitHub Display Text:", "github_disp"),
        ]
        vmap = {}
        for i, (lbl, k) in enumerate(fields):
            tk.Label(form, text=lbl, bg=COLORS["bg_card"], font=("Segoe UI", 9, "bold"), width=24, anchor=tk.W).grid(
                row=i, column=0, sticky=tk.W, pady=4
            )
            v = tk.StringVar(value=self.header.get(k, ""))
            ttk.Entry(form, textvariable=v).grid(row=i, column=1, sticky=tk.EW, pady=4)
            vmap[k] = v

        def save():
            self._push_undo()
            for k, v in vmap.items():
                self.header[k] = v.get().strip()
            dlg.destroy()
            self.refresh_all(compile_pdf=self.auto_compile_var.get())

        self._btn(form, "✓ Save Header", save, COLORS["emerald"], COLORS["emerald_hover"]).grid(
            row=len(fields), column=1, sticky=tk.E, pady=(12, 0)
        )

    def open_profile_modal(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Professional Profile Switcher & Editor")
        dlg.geometry("720x470")
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)
        dlg.grab_set()

        body = tk.Frame(dlg, bg=COLORS["bg_card"], padx=18, pady=14)
        body.pack(fill=tk.BOTH, expand=True)

        top_row = tk.Frame(body, bg=COLORS["bg_card"])
        top_row.pack(fill=tk.X, pady=(0, 8))

        en_var = tk.BooleanVar(value=self.profile_enabled)
        ttk.Checkbutton(top_row, text="Include 'Professional Profile' Section", variable=en_var).pack(side=tk.LEFT)

        prof_names = [p["name"] for p in self.profiles]
        sel_var = tk.StringVar(value=prof_names[self.active_profile_idx])
        combo = ttk.Combobox(top_row, textvariable=sel_var, values=prof_names, width=44, state="readonly")
        combo.pack(side=tk.RIGHT)

        txt = tk.Text(body, font=("Segoe UI", 10), wrap=tk.WORD, bg="#f8fafc", relief=tk.SOLID, bd=1, padx=10, pady=8)
        txt.pack(fill=tk.BOTH, expand=True, pady=6)
        txt.insert("1.0", self.profiles[self.active_profile_idx]["text"])

        def on_pick(_e=None):
            idx = prof_names.index(sel_var.get())
            txt.delete("1.0", tk.END)
            txt.insert("1.0", self.profiles[idx]["text"])

        combo.bind("<<ComboboxSelected>>", on_pick)

        def save():
            self._push_undo()
            idx = prof_names.index(sel_var.get())
            self.active_profile_idx = idx
            self.profile_enabled = en_var.get()
            self.profiles[idx]["text"] = txt.get("1.0", tk.END).strip()
            dlg.destroy()
            self.refresh_all(compile_pdf=self.auto_compile_var.get())

        self._btn(body, "✓ Apply Profile to resume.tex", save, COLORS["emerald"], COLORS["emerald_hover"]).pack(
            side=tk.RIGHT, pady=(6, 0)
        )

    def open_section_order_modal(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Reorder Resume Sections")
        dlg.geometry("420x380")
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)
        dlg.grab_set()

        body = tk.Frame(dlg, bg=COLORS["bg_card"], padx=16, pady=14)
        body.pack(fill=tk.BOTH, expand=True)

        lb = tk.Listbox(body, font=("Segoe UI", 10))
        lb.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        for s in self.section_order:
            lb.insert(tk.END, s)

        def shift(d: int):
            sel = lb.curselection()
            if not sel:
                return
            i = sel[0]
            j = i + d
            if 0 <= j < len(self.section_order):
                self._push_undo()
                self.section_order[i], self.section_order[j] = self.section_order[j], self.section_order[i]
                lb.delete(0, tk.END)
                for s in self.section_order:
                    lb.insert(tk.END, s)
                lb.selection_set(j)

        bar = tk.Frame(body, bg=COLORS["bg_card"])
        bar.pack(fill=tk.X)
        self._btn(bar, "▲ Move Up", lambda: shift(-1), "#475569", "#334155").pack(side=tk.LEFT, padx=3)
        self._btn(bar, "▼ Move Down", lambda: shift(1), "#475569", "#334155").pack(side=tk.LEFT, padx=3)
        self._btn(
            bar,
            "✓ Done",
            lambda: (dlg.destroy(), self.refresh_all(compile_pdf=self.auto_compile_var.get())),
            COLORS["emerald"],
            COLORS["emerald_hover"],
        ).pack(side=tk.RIGHT)

    # ------------------------------------------------------------------
    # Multi-Select Table Operations
    # ------------------------------------------------------------------
    def _selected_real_indices(self) -> list[int]:
        out = []
        for iid in self.tree.selection():
            t_idx = self.tree.index(iid)
            if 0 <= t_idx < len(self.visible_indices):
                out.append(self.visible_indices[t_idx])
        return out

    def on_tree_select(self, _e=None):
        idxs = self._selected_real_indices()
        self.code_text.tag_remove("active_block", "1.0", tk.END)
        if idxs and idxs[0] in self.item_line_map:
            s_ln, e_ln = self.item_line_map[idxs[0]]
            self.code_text.tag_add("active_block", f"{s_ln}.0", f"{e_ln}.end")
            self.code_text.see(f"{max(1, s_ln - 3)}.0")

    def edit_selected(self):
        idxs = self._selected_real_indices()
        if idxs:
            self.open_item_modal(self.items[idxs[0]].get("type", "custom"), edit_index=idxs[0])

    def toggle_selected(self):
        idxs = self._selected_real_indices()
        if not idxs:
            return
        self._push_undo()
        for r_idx in idxs:
            self.items[r_idx]["enabled"] = not self.items[r_idx].get("enabled", True)
        self.refresh_all(compile_pdf=self.auto_compile_var.get(), select_real_indices=idxs)

    def clone_selected(self):
        idxs = self._selected_real_indices()
        if not idxs:
            return
        self._push_undo()
        r_idx = idxs[0]
        c = copy.deepcopy(self.items[r_idx])
        c["uid"] = f"clone_{os.urandom(3).hex()}"
        for k in ("title", "company", "institution", "category"):
            if k in c:
                c[k] = f"{c[k]} (Copy)"
                break
        self.items.insert(r_idx + 1, c)
        self.refresh_all(compile_pdf=self.auto_compile_var.get(), select_real_indices=[r_idx + 1])

    def move_selected(self, delta: int):
        idxs = self._selected_real_indices()
        if len(idxs) != 1:
            return
        r_idx = idxs[0]
        if 0 <= r_idx + delta < len(self.items):
            self._push_undo()
            self.items[r_idx], self.items[r_idx + delta] = self.items[r_idx + delta], self.items[r_idx]
            self.refresh_all(compile_pdf=self.auto_compile_var.get(), select_real_indices=[r_idx + delta])

    def delete_selected(self):
        idxs = self._selected_real_indices()
        if not idxs:
            return
        self._push_undo()
        for r_idx in sorted(idxs, reverse=True):
            del self.items[r_idx]
        self.refresh_all(compile_pdf=self.auto_compile_var.get())

    # ------------------------------------------------------------------
    # LaTeX Bullet Renderer (Supports Nested Sub-Bullets like SADLDAS)
    # ------------------------------------------------------------------
    def _render_bullet_lines(self, bullets: list[str], base_indent: str) -> list[str]:
        out = []
        in_sub = False
        for raw_b in bullets:
            is_sub = raw_b.startswith(("  -", "\t-", "  *", "    "))
            clean_txt = re.sub(r"^(\s*[-\*]\s*|\s+)", "", raw_b) if is_sub else raw_b.strip()
            esc_txt = smart_escape_latex(clean_txt)
            if is_sub and not in_sub:
                out.append(f"{base_indent}    \\begin{{itemize}}[leftmargin=1.2em,topsep=0pt,itemsep=0pt,parsep=0pt]")
                in_sub = True
            elif not is_sub and in_sub:
                out.append(f"{base_indent}    \\end{{itemize}}")
                in_sub = False

            if is_sub:
                out.append(f"{base_indent}        \\item {esc_txt}")
            else:
                out.append(f"{base_indent}\\item {esc_txt}")

        if in_sub:
            out.append(f"{base_indent}    \\end{{itemize}}")
        return out

    # ------------------------------------------------------------------
    # Exact LaTeX Document Generator
    # ------------------------------------------------------------------
    def generate_latex(self) -> str:
        preamble = build_exact_preamble(self.layout, self.header.get("footer_name", ""))
        lines = preamble.rstrip().splitlines()
        lines.append("")
        lines.append(r"\begin{document}")
        lines.append(r"    \newcommand{\AND}{\unskip")
        lines.append(r"        \cleaders\copy\ANDbox\hskip\wd\ANDbox")
        lines.append(r"        \ignorespaces")
        lines.append(r"    }")
        lines.append(r"    \newsavebox\ANDbox")
        lines.append(r"    \sbox\ANDbox{}")
        lines.append("")

        # Header Block
        lines.append(r"    \begin{header}")
        lines.append(rf"        \textbf{{\fontsize{{24 pt}}{{24 pt}}\selectfont {smart_escape_latex(self.header['display_name'])}}}")
        lines.append("")
        lines.append(r"        \vspace{0.3cm}")
        lines.append("")
        lines.append(r"        \normalsize")
        lines.append(rf"        \mbox{{{{\color{{black}}\footnotesize\faPhone*}}\hspace*{{0.13cm}}{smart_escape_latex(self.header['phone'])}}}%")
        lines.append(r"        \kern 0.25cm%")
        lines.append(r"        \AND%")
        lines.append(r"        \kern 0.25cm%")
        lines.append(rf"        \mbox{{\hrefWithoutArrow{{mailto:{self.header['email']}}}{{\color{{black}}{{\footnotesize\faEnvelope[regular]}}\hspace*{{0.13cm}}{smart_escape_latex(self.header['email'])}}}}}%")
        lines.append(r"        \kern 0.25cm%")
        lines.append(r"        \AND%")
        lines.append(r"        \kern 0.25cm%")
        lines.append(rf"        \mbox{{\hrefWithoutArrow{{{self.header['linkedin_url']}}}{{\color{{black}}{{\footnotesize\faLinkedinIn}}\hspace*{{0.13cm}}{smart_escape_latex(self.header['linkedin_disp'])}}}}}%")
        lines.append(r"        \AND%")
        lines.append(r"        \kern 0.25cm%")
        lines.append(rf"        \mbox{{\hrefWithoutArrow{{{self.header['github_url']}}}{{\color{{black}}{{\footnotesize\faGithub}}\hspace*{{0.13cm}}{smart_escape_latex(self.header['github_disp'])}}}}}%")
        lines.append(r"        \kern 0.25cm%")
        lines.append(r"    \end{header}")

        self.item_line_map.clear()
        grouped: dict[str, list[tuple[int, dict]]] = {}
        for r_idx, it in enumerate(self.items):
            if it.get("enabled", True):
                grouped.setdefault(it.get("section", "Other"), []).append((r_idx, it))

        ordered_secs = [s for s in self.section_order if s == "Professional Profile" or s in grouped] + [
            s for s in grouped if s not in self.section_order
        ]

        for sec in ordered_secs:
            if sec == "Professional Profile":
                if not self.profile_enabled:
                    continue
                prof_txt = smart_escape_latex(self.profiles[self.active_profile_idx]["text"])
                lines.append("")
                lines.append(r"    % PROFESSIONAL SUMMARY")
                lines.append(r"    \section{Professional Profile}")
                lines.append("")
                lines.append(r"    \begin{firstprojectentry}")
                lines.append(r"        \begin{onecolentry}")
                lines.append(f"            {prof_txt}")
                lines.append(r"        \end{onecolentry}")
                lines.append(r"    \end{firstprojectentry}")
                continue

            entries = grouped.get(sec, [])
            if not entries:
                continue

            lines.append("")
            lines.append(rf"    % {sec.upper()}")
            lines.append(rf"    \section{{{smart_escape_latex(sec)}}}")

            for idx_in_sec, (r_idx, it) in enumerate(entries):
                itype = it.get("type", "custom")
                start_ln = len(lines) + 1

                if itype == "education":
                    inst = smart_escape_latex(it.get("institution", ""))
                    deg = smart_escape_latex(it.get("degree", ""))
                    dt = smart_escape_latex(it.get("date", ""))
                    is_sub = it.get("is_sub_degree", False)

                    if idx_in_sec == 0 and not is_sub:
                        lines.append("")
                        lines.append(r"    \begin{firstprojectentry}")
                        lines.append(r"    \begin{twocolentry}{")
                        lines.append(r"        \textit{}")
                        lines.append(r"    }")
                        lines.append(rf"        \textbf{{{inst}}}")
                        lines.append(r"    \end{twocolentry}")
                        lines.append(r"    \end{firstprojectentry}")
                        lines.append("")
                        lines.append(r"    \begin{projectentry}")
                        lines.append(r"    \begin{twocolentry}{")
                        lines.append(rf"        \textit{{{dt}}}")
                        lines.append(r"    }")
                        lines.append(rf"        \textit{{{deg}}}")
                        lines.append(r"    \end{twocolentry}")
                    elif is_sub:
                        lines.append("")
                        lines.append(r"    \vspace{0.2cm}")
                        lines.append("")
                        lines.append(r"    \begin{projectentry}")
                        lines.append(r"    \begin{twocolentry}{")
                        lines.append(rf"        \textit{{{dt}}}")
                        lines.append(r"    }")
                        lines.append(rf"        \textit{{{deg}}}")
                        lines.append(r"    \end{twocolentry}")
                    else:
                        lines.append("")
                        lines.append(r"    \vspace{0.2cm}")
                        lines.append("")
                        lines.append(r"    \begin{projectentry}")
                        lines.append(r"    \begin{twocolentry}{")
                        lines.append(rf"        \textit{{{dt}}}")
                        lines.append(r"    }")
                        lines.append(rf"        \textbf{{{inst}}}")
                        lines.append("")
                        lines.append(rf"        \textit{{{deg}}}")
                        lines.append(r"    \end{twocolentry}")

                    all_b = list(it.get("bullets", []))
                    if it.get("links"):
                        lk_str = ", ".join(
                            rf"\hrefWithoutArrow{{{lk['url']}}}{{{smart_escape_latex(lk['label'])}}}"
                            for lk in it["links"]
                        )
                        all_b.insert(1 if len(all_b) > 0 else 0, f"Access: {lk_str}")

                    if all_b:
                        lines.append(r"    \vspace{0.10cm}")
                        lines.append(r"    \begin{onecolentry}")
                        lines.append(r"        \begin{highlights}")
                        lines.extend(self._render_bullet_lines(all_b, "            "))
                        lines.append(r"        \end{highlights}")
                        lines.append(r"    \end{onecolentry}")
                    lines.append(r"    \end{projectentry}")

                elif itype == "experience":
                    comp = smart_escape_latex(it.get("company", ""))
                    role = smart_escape_latex(it.get("role", ""))
                    dt = smart_escape_latex(it.get("date", ""))
                    is_sub = it.get("is_sub_role", False)

                    if idx_in_sec > 0 and not is_sub:
                        prev_gap = entries[idx_in_sec - 1][1].get("space_after", "0.4cm") or "0.4cm"
                        lines.append("")
                        lines.append(rf"    \vspace{{{prev_gap}}}")

                    lines.append("")
                    lines.append(r"    \begin{projectentry}")
                    lines.append(r"    \begin{twocolentry}{")
                    lines.append(rf"        \textit{{{dt}}}")
                    lines.append(r"    }")
                    if not is_sub and comp:
                        lines.append(rf"        \textbf{{{comp}}}")
                        lines.append("")
                    lines.append(rf"        \textit{{{role}}}")
                    lines.append(r"    \end{twocolentry}")

                    if it.get("bullets"):
                        lines.append(r"    \vspace{0.10cm}")
                        lines.append(r"    \begin{onecolentry}")
                        lines.append(r"        \begin{highlights}")
                        lines.extend(self._render_bullet_lines(it["bullets"], "            "))
                        lines.append(r"        \end{highlights}")
                        lines.append(r"    \end{onecolentry}")
                    lines.append(r"    \end{projectentry}")

                elif itype == "award":
                    env = "firstprojectentry" if idx_in_sec == 0 else "projectentry"
                    title = smart_escape_latex(it.get("title", ""))
                    dt = smart_escape_latex(it.get("date", ""))
                    lines.append("")
                    lines.append(rf"    \begin{{{env}}}")
                    lines.append(rf"        \begin{{twocolentry}}{{\textit{{{dt}}}}}")
                    lines.append(rf"            \textbf{{{title}}}")
                    lines.append(r"        \end{twocolentry}")
                    lines.append(rf"    \end{{{env}}}")

                elif itype == "project":
                    env = "firstprojectentry" if idx_in_sec == 0 else "projectentry"
                    title = smart_escape_latex(it.get("title", ""))
                    dt = smart_escape_latex(it.get("date", ""))
                    sub_it = smart_escape_latex(it.get("subtitle_italic", ""))
                    stack = smart_escape_latex(it.get("stack", ""))
                    links = it.get("links", [])

                    lines.append("")
                    lines.append(rf"    \begin{{{env}}}")
                    lines.append(r"    \begin{twocolentry}{")
                    lines.append(rf"        \textit{{{dt}}}")
                    lines.append(r"    }")
                    lines.append(rf"        \textbf{{{title}}}")
                    lines.append(r"    \end{twocolentry}")
                    lines.append(r"    \vspace{0.10cm}")
                    lines.append(r"    \begin{onecolentry}")
                    if sub_it:
                        lines.append(rf"        \textit{{{sub_it}}}")
                    lines.append(r"        \begin{highlights}")
                    if stack:
                        lines.append(rf"            \item Stack: {stack}")
                    if links:
                        lk_joined = ", ".join(
                            rf"\hrefWithoutArrow{{{lk['url']}}}{{{smart_escape_latex(lk['label'])}}}"
                            for lk in links
                        )
                        lines.append(rf"            \item Access: {lk_joined}")
                    if it.get("extra_bullets"):
                        lines.extend(self._render_bullet_lines(it["extra_bullets"], "            "))
                    lines.append(r"        \end{highlights}")
                    lines.append(r"    \end{onecolentry}")
                    lines.append(rf"    \end{{{env}}}")

                elif itype == "skill_group":
                    env = "firstprojectentry" if idx_in_sec == 0 else "projectentry"
                    cat = smart_escape_latex(it.get("category", ""))
                    lines.append("")
                    lines.append(rf"    \begin{{{env}}}")
                    lines.append(r"    \begin{twocolentry}{")
                    lines.append(r"        \textit{}")
                    lines.append(r"    }")
                    lines.append(rf"        \textbf{{{cat}}}")
                    lines.append(r"    \end{twocolentry}")
                    lines.append(r"    \vspace{0.10cm}")
                    lines.append(r"    \begin{onecolentry}")
                    lines.append(r"        \begin{highlights}")
                    lines.extend(self._render_bullet_lines(it.get("bullets", []), "            "))
                    lines.append(r"        \end{highlights}")
                    lines.append(r"    \end{onecolentry}")
                    lines.append(rf"    \end{{{env}}}")

                else:  # custom section entry
                    env = "firstprojectentry" if idx_in_sec == 0 else "projectentry"
                    title = smart_escape_latex(it.get("title", ""))
                    sub_it = smart_escape_latex(it.get("subtitle_italic", ""))
                    dt = smart_escape_latex(it.get("date", ""))
                    lines.append("")
                    lines.append(rf"    \begin{{{env}}}")
                    lines.append(r"    \begin{twocolentry}{")
                    lines.append(rf"        \textit{{{dt}}}")
                    lines.append(r"    }")
                    lines.append(rf"        \textbf{{{title}}}")
                    if sub_it:
                        lines.append("")
                        lines.append(rf"        \textit{{{sub_it}}}")
                    lines.append(r"    \end{twocolentry}")
                    if it.get("bullets"):
                        lines.append(r"    \vspace{0.10cm}")
                        lines.append(r"    \begin{onecolentry}")
                        lines.append(r"        \begin{highlights}")
                        lines.extend(self._render_bullet_lines(it["bullets"], "            "))
                        lines.append(r"        \end{highlights}")
                        lines.append(r"    \end{onecolentry}")
                    lines.append(rf"    \end{{{env}}}")

                self.item_line_map[r_idx] = (start_ln, len(lines))

        lines.append("")
        lines.append(r"\end{document}")
        lines.append("")
        return "\n".join(lines)

    def _highlight_syntax(self, code: str):
        self.code_text.delete("1.0", tk.END)
        self.code_text.insert("1.0", code)
        for idx, line in enumerate(code.splitlines(), start=1):
            for m in re.finditer(r"(?<!\\)%.*$", line):
                self.code_text.tag_add("comment", f"{idx}.{m.start()}", f"{idx}.{m.end()}")
            for m in re.finditer(r"\\section\*?\{[^\}]*\}", line):
                self.code_text.tag_add("section", f"{idx}.{m.start()}", f"{idx}.{m.end()}")
            for m in re.finditer(r"\\(begin|end)\{[^\}]*\}", line):
                self.code_text.tag_add("env", f"{idx}.{m.start()}", f"{idx}.{m.end()}")
            for m in re.finditer(r"\\[A-Za-z@]+\*?", line):
                if not line[m.start():].startswith((r"\section", r"\begin", r"\end")):
                    self.code_text.tag_add("cmd", f"{idx}.{m.start()}", f"{idx}.{m.end()}")
            for m in re.finditer(r"\$[^\$]+\$", line):
                self.code_text.tag_add("math", f"{idx}.{m.start()}", f"{idx}.{m.end()}")

    # ------------------------------------------------------------------
    # Embedded Visual PDF Renderer (PyMuPDF or pdftoppm)
    # ------------------------------------------------------------------
    def _render_pdf_pages(self):
        if not os.path.exists(PDF_FILE):
            return
        self.pdf_images.clear()
        try:
            if HAS_PYMUPDF:
                doc = fitz.open(PDF_FILE)
                mat = fitz.Matrix(self.pdf_zoom, self.pdf_zoom)
                for page in doc:
                    pix = page.get_pixmap(matrix=mat)
                    img = tk.PhotoImage(data=pix.tobytes("ppm"))
                    self.pdf_images.append(img)
                doc.close()
            elif shutil.which("pdftoppm"):
                dpi = int(110 * self.pdf_zoom)
                with tempfile.TemporaryDirectory() as tmpdir:
                    prefix = os.path.join(tmpdir, "page")
                    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), PDF_FILE, prefix], check=True)
                    for png in sorted(glob.glob(f"{prefix}-*.png")):
                        self.pdf_images.append(tk.PhotoImage(file=png))
            else:
                self.pdf_canvas.delete("all")
                self.pdf_canvas.create_text(
                    260,
                    160,
                    text="In-app PDF preview requires PyMuPDF (pip install pymupdf)\nor Poppler 'pdftoppm' in PATH.\n\nClick 'External PDF' in the top bar to view resume.pdf.",
                    fill="#1e293b",
                    font=("Segoe UI", 10, "bold"),
                    justify=tk.CENTER,
                )
                return

            if self.pdf_images:
                self.pdf_page_idx = max(0, min(self.pdf_page_idx, len(self.pdf_images) - 1))
                self._draw_current_pdf_page()
        except Exception as e:
            self.status_lbl.config(text=f"PDF preview notice: {e}")

    def _draw_current_pdf_page(self):
        if not self.pdf_images:
            return
        self.pdf_canvas.delete("all")
        img = self.pdf_images[self.pdf_page_idx]
        self.pdf_canvas.create_image(20, 20, anchor=tk.NW, image=img)
        self.pdf_canvas.configure(scrollregion=(0, 0, img.width() + 40, img.height() + 40))
        self.pdf_page_lbl.config(text=f"Page {self.pdf_page_idx + 1} / {len(self.pdf_images)}")

    def _change_pdf_page(self, delta: int):
        if not self.pdf_images:
            return
        self.pdf_page_idx = max(0, min(self.pdf_page_idx + delta, len(self.pdf_images) - 1))
        self._draw_current_pdf_page()

    def _zoom_pdf(self, delta: float):
        self.pdf_zoom = max(0.6, min(2.4, self.pdf_zoom + delta))
        self._render_pdf_pages()

    # ------------------------------------------------------------------
    # Refresh, Persistence, Compile & Log Telemetry
    # ------------------------------------------------------------------
    def refresh_all(self, compile_pdf: bool = False, select_real_indices: list[int] | None = None):
        self.contact_banner_lbl.config(
            text=f"{self.header['phone']}   •   {self.header['email']}   •   {self.header['github_disp']}"
        )

        secs = ["All Sections"] + list(
            dict.fromkeys(self.section_order + [i.get("section", "Other") for i in self.items])
        )
        self.sec_filter_combo["values"] = secs

        for row in self.tree.get_children():
            self.tree.delete(row)

        self.visible_indices.clear()
        q = self.search_var.get().strip().lower()
        sec_f = self.filter_sec_var.get()
        st_f = self.filter_state_var.get()
        iids_to_select = []

        for r_idx, it in enumerate(self.items):
            en = it.get("enabled", True)
            if st_f == "Active Only (✓)" and not en:
                continue
            if st_f == "Archived / Commented (✗)" and en:
                continue

            sec = it.get("section", "Other")
            if sec_f != "All Sections" and sec != sec_f:
                continue

            title_disp = it.get("title") or it.get("company") or it.get("institution") or it.get("category") or ""
            detail_disp = (
                it.get("role")
                or it.get("degree")
                or it.get("stack")
                or (", ".join(it.get("bullets", []))[:55])
            )
            date_disp = it.get("date", "").replace(r"\textsuperscript{th}", "th").replace(r"\textsuperscript{nd}", "nd")

            blob = f"{sec} {title_disp} {detail_disp} {date_disp}".lower()
            if q and q not in blob:
                continue

            self.visible_indices.append(r_idx)
            iid = self.tree.insert(
                "",
                tk.END,
                values=("✓" if en else "✗", sec, title_disp, detail_disp, date_disp),
                tags=("on" if en else "off",),
            )
            if select_real_indices and r_idx in select_real_indices:
                iids_to_select.append(iid)

        active_cnt = sum(1 for i in self.items if i.get("enabled", True))
        self.stats_lbl.config(text=f"Active: {active_cnt} / {len(self.items)}")

        tex = self.generate_latex()
        self._highlight_syntax(tex)

        if iids_to_select:
            self.tree.selection_set(iids_to_select)
            self.on_tree_select()

        self._save_state()
        if self.auto_save_var.get():
            try:
                with open(TEX_FILE, "w", encoding="utf-8") as f:
                    f.write(tex)
                self.status_lbl.config(text=f"Auto-saved {TEX_FILE} ({active_cnt} active items).")
            except Exception as e:
                self.status_lbl.config(text=f"Write error: {e}")

        if compile_pdf:
            self.compile_pdf()

    def _save_state(self):
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                f.write(self._serialize())
        except Exception:
            pass

    def _load_state(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    self._deserialize(f.read())
            except Exception:
                pass

    def save_tex_as(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".tex", initialfile="resume.tex", filetypes=[("LaTeX Files", "*.tex")]
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.generate_latex())
            self.status_lbl.config(text=f"Exported to {path}")

    def copy_tex(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.generate_latex())
        self.status_lbl.config(text="Copied generated resume.tex to clipboard!")

    def compile_pdf(self):
        tex = self.generate_latex()
        with open(TEX_FILE, "w", encoding="utf-8") as f:
            f.write(tex)
        self.status_lbl.config(text="Compiling resume.pdf via MiKTeX pdflatex...")

        def worker():
            try:
                proc = subprocess.run(
                    ["pdflatex", "-interaction=nonstopmode", TEX_FILE],
                    capture_output=True,
                    text=True,
                    timeout=35,
                )
                self.root.after(0, self._inspect_log)
                self.root.after(0, self._render_pdf_pages)
                if proc.returncode == 0:
                    self.root.after(0, lambda: self.status_lbl.config(text="✓ Compiled resume.pdf successfully!"))
                else:
                    self.root.after(
                        0, lambda: self.status_lbl.config(text=f"⚠ pdflatex exited with code {proc.returncode}.")
                    )
            except FileNotFoundError:
                self.root.after(0, lambda: self.status_lbl.config(text="Error: 'pdflatex' not found in PATH."))
            except Exception as e:
                self.root.after(0, lambda: self.status_lbl.config(text=f"Compile error: {e}"))

        threading.Thread(target=worker, daemon=True).start()

    def _inspect_log(self):
        if not os.path.exists(LOG_FILE):
            self.log_badge.config(text="resume.log: Not found")
            return
        try:
            with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as f:
                data = f.read()
            m = re.search(r"Output written on [^\(]+\((\d+)\s+pages?,\s+(\d+)\s+bytes\)", data)
            errs = len(re.findall(r"^! ", data, re.MULTILINE))
            if m:
                pages, kb = int(m.group(1)), int(m.group(2)) / 1024
                self.log_badge.config(
                    text=f"📄 {pages} Page(s)  •  {kb:.0f} KB  •  {errs} Errors",
                    bg="#bbf7d0" if errs == 0 else "#fecaca",
                )
        except Exception:
            pass

    def open_pdf(self):
        if not os.path.exists(PDF_FILE):
            messagebox.showinfo("No PDF Yet", "Click 'Compile PDF' first to build resume.pdf.")
            return
        if sys.platform.startswith("win"):
            os.startfile(PDF_FILE)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", PDF_FILE])
        else:
            subprocess.Popen(["xdg-open", PDF_FILE])


if __name__ == "__main__":
    root = tk.Tk()
    app = JosephResumeStudioV2(root)
    root.mainloop()
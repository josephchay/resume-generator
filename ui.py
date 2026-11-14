import os
import re
import sys
import json
import copy
import ctypes
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Enable crisp High-DPI font scaling on Windows
if sys.platform.startswith("win"):
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

DATA_FILE = "joseph_resume_data.json"
TEX_FILE = "resume.tex"
LOG_FILE = "resume.log"

# ==============================================================================
# Theme Colors (Matched to your primaryColor RGB{0, 79, 144})
# ==============================================================================
COLORS = {
    "bg_main": "#f1f5f9",
    "bg_card": "#ffffff",
    "bg_header": "#003566",     # Deep Navy
    "primary": "#004f90",       # Your exact RGB(0, 79, 144)
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

# ==============================================================================
# Exact Preamble & Header Macros from Your resume.tex
# ==============================================================================
EXACT_PREAMBLE = r"""\documentclass[10pt, letterpaper]{article}

% Packages:
\usepackage[
    ignoreheadfoot,
    top=2cm,
    bottom=2cm,
    left=2cm,
    right=2cm,
    footskip=1.0cm,
]{geometry}
\usepackage{titlesec}
\usepackage{tabularx}
\usepackage{array}
\usepackage[dvipsnames]{xcolor}
\definecolor{primaryColor}{RGB}{0, 79, 144}
\usepackage{enumitem}
\usepackage{fontawesome5}
\usepackage{amsmath}
\usepackage[
    pdftitle={Joseph Emmanuel Chay Huan Qin's CV},
    pdfauthor={Joseph Emmanuel Chay Huan Qin},
    pdfcreator={LaTeX with RenderCV},
    colorlinks=true,
    urlcolor=primaryColor,
    linkcolor=black,
    breaklinks=true
]{hyperref}
\usepackage{fancyhdr}
\usepackage{calc}
\usepackage{bookmark}
\usepackage{lastpage}
\usepackage{changepage}
\usepackage{paracol}
\usepackage{ifthen}
\usepackage{needspace}
\usepackage{iftex}
\usepackage{paracol}
\footnotelayout{m}

\ifPDFTeX
    \input{glyphtounicode}
    \pdfgentounicode=1
    \usepackage[utf8]{inputenc}
    \usepackage{lmodern}
\fi

\AtBeginEnvironment{adjustwidth}{\partopsep0pt}
\setcounter{secnumdepth}{0}
\setlength{\parindent}{0pt}
\setlength{\topskip}{0pt}
\setlength{\columnsep}{0cm}

% --- Footer Redefinition ---
\pagestyle{fancy}
\fancyhf{}
\fancyfoot[C]{\color{gray}\textit{\small Joseph Emmanuel Chay Huan Qin}}
% \fancyfoot[R]{\color{gray}\textit{\small Page \thepage\ of \pageref*{LastPage}}}
\fancyfoot[R]{\color{gray}\textit{\small \thepage}}
\renewcommand{\headrulewidth}{0pt}
\renewcommand{\footrulewidth}{0pt}
% ---------------------------

% Optimal spacing: ensures section titles stay with minimum content while allowing natural breaks
\titleformat{\section}{\needspace{4\baselineskip}\bfseries\large}{}{0pt}{}[\vspace{1pt}\titlerule]
\titlespacing{\section}{-1pt}{0.3cm}{0.2cm}

\renewcommand\labelitemi{$\circ$}
\newenvironment{highlights}{
    \begin{itemize}[
        topsep=0.10cm,
        parsep=0.10cm,
        partopsep=0pt,
        itemsep=0pt,
        leftmargin=0.4cm + 10pt
    ]
}{
    \end{itemize}
}
\newenvironment{onecolentry}{
    \begin{adjustwidth}{0.2cm + 0.00001cm}{0.2cm + 0.00001cm}
}{
    \end{adjustwidth}
}
\newenvironment{twocolentry}[2][]{
    \onecolentry
    \def\secondColumn{#2}
    \setcolumnwidth{\fill, 5.2cm}
    \begin{paracol}{2}
}{
    \switchcolumn \raggedleft \secondColumn
    \end{paracol}
    \endonecolentry
}
\newenvironment{header}{
    \setlength{\topsep}{0pt}\par\kern\topsep\centering\linespread{1.5}
}{
    \par\kern\topsep
}

% Environment for project entries - prevents splitting while allowing natural page breaks
\newenvironment{projectentry}{
    \needspace{2.5\baselineskip}
}{
}

% Environment for first entry after section - minimal space since section already reserves space
\newenvironment{firstprojectentry}{
    \needspace{2\baselineskip}
}{
}

\let\hrefWithoutArrow\href
\renewcommand{\href}[2]{\hrefWithoutArrow{#1}{\ifthenelse{\equal{#2}{}}{ }{#2 }\raisebox{.15ex}{\footnotesize \faExternalLink*}}}
"""


def smart_escape_latex(text: str) -> str:
    """
    Escapes bare &, %, #, _ outside of $...$ math mode while preserving
    existing LaTeX commands (e.g. \\textsuperscript, \\footnote, \\hrefWithoutArrow).
    """
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
# Pre-Loaded Data Extracted Directly from Your resume.tex
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
        "name": "Active: Technology, AI & Multidisciplinary Interests",
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
        "name": "Variant 1: AI, ML, Robotics & Data Engineering",
        "text": (
            "Computer Science with Artificial Intelligence graduate and a Diploma holder in Information and Communication Technology with Software Engineering, "
            "bringing a strong passion and well-developed expertise in building intelligent, scalable systems that solve real-world problems. My skill set spans machine learning, "
            "reinforcement learning, deep learning, NLP, computer vision, object detection, image processing, robotics automation, and data engineering, "
            "supported by solid foundations in software engineering, cloud computing, and distributed systems."
        ),
    },
    {
        "name": "Variant 2: Cross-Industry AI & Full-Stack Engineering",
        "text": (
            "Worked on simple projects in university involving intelligent agents, robotics, object detection, RAG-based systems, chatbot agents, and large-scale data pipelines. "
            "My knowledge and hands-on expertise position me to contribute across industries such as healthcare (medical imaging, AI diagnostics), "
            "finance (fraud detection, predictive analytics), e-commerce (personalization, recommendation systems), "
            "software engineering (web development, mobile app development, API integration, database design and management, cloud deployment)."
        ),
    },
    {
        "name": "Variant 3: Client Communication & Team Building",
        "text": (
            "Beyond software development, I carry great communication skills and motivation to work with people, including team buildings. "
            "Spoken roles interest me as it builds understanding and trust between people indulging clients into the products."
        ),
    },
    {
        "name": "Variant 4: Systems Monitoring, Surveillance & Access Control",
        "text": (
            "Beyond software development, I bring a strong aptitude for systems monitoring and technical troubleshooting, skills that translate directly into overseeing equipment-dependent environments. "
            "My background in computer vision and object detection gives me a practical understanding of how CCTV and surveillance systems function at a technical level, allowing me to quickly identify "
            "malfunctions, blind spots, or synchronization issues rather than treating such equipment as a black box. Similarly, my experience with API integration, database design, and access control "
            "logic in software projects provides a solid foundation for handling access card programming, entry/exit system synchronization, and coordination with equipment service providers when "
            "technical issues arise."
        ),
    },
    {
        "name": "Variant 5: Data Engineering & Operational Audits",
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
    # --- EDUCATION ---
    {
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

    # --- INDUSTRIAL EXPERIENCE ---
    {
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

    # --- HONOURS & AWARDS ---
    {
        "enabled": False,
        "section": "Honours & Awards",
        "type": "award",
        "title": "Best Final Year Project (1\\textsuperscript{st} Runner-Up), Final Year Dissertation",
        "date": "Jul 2025",
    },
    {
        "enabled": False,
        "section": "Honours & Awards",
        "type": "award",
        "title": "Best Industry Project (Top 5), Software Engineering Group Project\\footnote{Awarded during the closing ceremony of the open demonstration day for our project, \\textit{Smart Fridge}.}",
        "date": "Apr 2024",
    },
    {
        "enabled": False,
        "section": "Honours & Awards",
        "type": "award",
        "title": "ABRSM Piano Grade 8 Theory, Passed with Distinction",
        "date": "May 2020",
    },
    {
        "enabled": True,
        "section": "Honours & Awards",
        "type": "award",
        "title": "Deputy Head Prefect, Regent International School",
        "date": "Dec 2019",
    },
    {
        "enabled": True,
        "section": "Honours & Awards",
        "type": "award",
        "title": "ABRSM Piano Grade 8 Practical, Passed with Distinction",
        "date": "Jul 2019",
    },
    {
        "enabled": True,
        "section": "Honours & Awards",
        "type": "award",
        "title": "Head Prefect, Regent International School",
        "date": "Dec 2018",
    },

    # --- CLIENT PROJECTS (Pre-loaded from commented block in resume.tex) ---
    {
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

    # --- PERSONAL PROJECTS (Active + Commented blocks from resume.tex) ---
    {
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
            "Solved 11 Gymnasium/MuJoCo continuous control benchmarks (Ant-v5, HalfCheetah-v5, Hopper-v5, Walker2d-v5, Humanoid-v5, BipedalWalker-v3) within leaderboard thresholds"
        ],
    },
    {
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

    # --- PROFICIENCY AND TOOLS ---
    {
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

    # --- LANGUAGES ---
    {
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


# ==============================================================================
# Main Application Class
# ==============================================================================
class JosephResumeStudio:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Joseph Chay — LaTeX Resume Generator & Section Manager")
        self.root.geometry("1400x860")
        self.root.minsize(1100, 680)
        self.root.configure(bg=COLORS["bg_main"])

        self.header = copy.deepcopy(DEFAULT_HEADER)
        self.profiles = copy.deepcopy(DEFAULT_PROFILES)
        self.active_profile_idx = 0
        self.profile_enabled = True
        self.section_order = copy.deepcopy(DEFAULT_SECTION_ORDER)
        self.items = copy.deepcopy(DEFAULT_ITEMS)

        self.filter_sec_var = tk.StringVar(value="All Sections")
        self.search_var = tk.StringVar()
        self.auto_save_var = tk.BooleanVar(value=True)
        self.auto_compile_var = tk.BooleanVar(value=False)

        self.visible_indices: list[int] = []
        self.item_line_map: dict[int, tuple[int, int]] = {}

        self._load_state()
        self._setup_styles()
        self._build_ui()
        self._bind_shortcuts()
        self.refresh_all(compile_pdf=False)
        self._inspect_log()

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

    def _btn(self, parent, text: str, cmd, bg: str, hover: str, pad=(11, 5)) -> tk.Button:
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
        # 1. Top Header Banner
        banner = tk.Frame(self.root, bg=COLORS["bg_header"], padx=18, pady=12)
        banner.pack(fill=tk.X)

        left_info = tk.Frame(banner, bg=COLORS["bg_header"])
        left_info.pack(side=tk.LEFT)
        tk.Label(
            left_info,
            text="JOSEPH CHAY — LaTeX Resume Studio",
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
        self._btn(right_btns, "👤 Contact Info", self.open_contact_modal, "#1e40af", "#1d4ed8").pack(side=tk.LEFT, padx=4)
        self._btn(right_btns, "📝 Profile Summary", self.open_profile_modal, "#1e40af", "#1d4ed8").pack(side=tk.LEFT, padx=4)
        self._btn(right_btns, "📑 Section Order", self.open_section_order_modal, "#1e40af", "#1d4ed8").pack(side=tk.LEFT, padx=4)
        self._btn(right_btns, "⚙ Compile PDF", self.compile_pdf, COLORS["emerald"], COLORS["emerald_hover"]).pack(side=tk.LEFT, padx=4)
        self._btn(right_btns, "📄 Open PDF", self.open_pdf, "#0284c7", "#0369a1").pack(side=tk.LEFT, padx=4)

        # 2. Quick-Add Action Ribbon
        ribbon = tk.Frame(
            self.root, bg=COLORS["bg_card"], padx=14, pady=10, highlightthickness=1, highlightbackground=COLORS["border"]
        )
        ribbon.pack(fill=tk.X, padx=14, pady=(10, 6))

        tk.Label(
            ribbon, text="ADD NEW ENTRY:", font=("Segoe UI", 9, "bold"), bg=COLORS["bg_card"], fg=COLORS["fg_muted"]
        ).pack(side=tk.LEFT, padx=(0, 8))

        add_specs = [
            ("＋ Add Project", "project", COLORS["primary"], COLORS["primary_hover"]),
            ("＋ Add Experience", "experience", "#0284c7", "#0369a1"),
            ("＋ Add Education", "education", "#0d9488", "#0f766e"),
            ("＋ Add Honour / Award", "award", "#7c3aed", "#6d28d9"),
            ("＋ Add Skill / Language Group", "skill_group", "#d97706", "#b45309"),
            ("＋ Add Custom Section Item", "custom", "#475569", "#334155"),
        ]
        for label, itype, bg_c, hov_c in add_specs:
            self._btn(ribbon, label, lambda t=itype: self.open_item_modal(t), bg_c, hov_c, pad=(10, 5)).pack(
                side=tk.LEFT, padx=3
            )

        # 3. Main Split Pane (Left: Filterable Item Table | Right: Live resume.tex Code)
        paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=14, pady=6)

        # --- LEFT PANEL: Structured Items Table ---
        left_card = tk.Frame(
            paned, bg=COLORS["bg_card"], highlightthickness=1, highlightbackground=COLORS["border"]
        )
        paned.add(left_card, weight=5)

        filter_bar = tk.Frame(left_card, bg=COLORS["bg_card"], padx=12, pady=10)
        filter_bar.pack(fill=tk.X)

        tk.Label(filter_bar, text="Section:", bg=COLORS["bg_card"], fg=COLORS["fg_muted"], font=("Segoe UI", 9, "bold")).pack(
            side=tk.LEFT
        )
        self.sec_filter_combo = ttk.Combobox(filter_bar, textvariable=self.filter_sec_var, width=22, state="readonly")
        self.sec_filter_combo.pack(side=tk.LEFT, padx=(6, 14))
        self.sec_filter_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_all(compile_pdf=False))

        tk.Label(filter_bar, text="Search:", bg=COLORS["bg_card"], fg=COLORS["fg_muted"], font=("Segoe UI", 9, "bold")).pack(
            side=tk.LEFT
        )
        ttk.Entry(filter_bar, textvariable=self.search_var, width=22).pack(side=tk.LEFT, padx=6)
        self.search_var.trace_add("write", lambda *_: self.refresh_all(compile_pdf=False))

        self.stats_lbl = tk.Label(filter_bar, text="", bg=COLORS["bg_card"], fg=COLORS["fg_muted"], font=("Segoe UI", 9))
        self.stats_lbl.pack(side=tk.RIGHT)

        tree_wrap = tk.Frame(left_card, bg=COLORS["bg_card"], padx=12)
        tree_wrap.pack(fill=tk.BOTH, expand=True)

        cols = ("active", "section", "title", "detail", "date")
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="headings", style="CV.Treeview", selectmode="browse")
        self.tree.heading("active", text="Active")
        self.tree.heading("section", text="Section")
        self.tree.heading("title", text="Title / Institution / Project")
        self.tree.heading("detail", text="Role / Degree / Tech Stack")
        self.tree.heading("date", text="Date")

        self.tree.column("active", width=52, anchor=tk.CENTER, stretch=False)
        self.tree.column("section", width=135, anchor=tk.W, stretch=False)
        self.tree.column("title", width=230, anchor=tk.W, stretch=True)
        self.tree.column("detail", width=200, anchor=tk.W, stretch=True)
        self.tree.column("date", width=120, anchor=tk.E, stretch=False)

        self.tree.tag_configure("on", foreground="#0f172a")
        self.tree.tag_configure("off", foreground="#94a3b8")

        yscroll = ttk.Scrollbar(tree_wrap, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())
        self.tree.bind("<space>", lambda e: self.toggle_selected())

        # Table Action Buttons
        tbl_bar = tk.Frame(left_card, bg="#f8fafc", padx=12, pady=9)
        tbl_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self._btn(tbl_bar, "✏ Edit (Double-Click)", self.edit_selected, COLORS["primary"], COLORS["primary_hover"], pad=(9, 4)).pack(
            side=tk.LEFT, padx=2
        )
        self._btn(tbl_bar, "✓/✗ Toggle Active (Space)", self.toggle_selected, "#475569", "#334155", pad=(9, 4)).pack(
            side=tk.LEFT, padx=2
        )
        self._btn(tbl_bar, "⧉ Clone", self.clone_selected, "#475569", "#334155", pad=(9, 4)).pack(side=tk.LEFT, padx=2)
        self._btn(tbl_bar, "▲ Up", lambda: self.move_selected(-1), "#64748b", "#475569", pad=(8, 4)).pack(side=tk.LEFT, padx=2)
        self._btn(tbl_bar, "▼ Down", lambda: self.move_selected(1), "#64748b", "#475569", pad=(8, 4)).pack(side=tk.LEFT, padx=2)
        self._btn(tbl_bar, "🗑 Delete", self.delete_selected, COLORS["danger"], "#b91c1c", pad=(9, 4)).pack(side=tk.RIGHT, padx=2)

        # --- RIGHT PANEL: Live LaTeX Preview ---
        right_card = tk.Frame(
            paned, bg=COLORS["code_bg"], highlightthickness=1, highlightbackground=COLORS["border"]
        )
        paned.add(right_card, weight=4)

        code_top = tk.Frame(right_card, bg="#181825", padx=12, pady=8)
        code_top.pack(fill=tk.X)
        tk.Label(
            code_top, text="GENERATED RESUME.TEX (EXACT RENDERCY TEMPLATE)", font=("Segoe UI", 9, "bold"), bg="#181825", fg="#a6adc8"
        ).pack(side=tk.LEFT)
        self._btn(code_top, "💾 Save As...", self.save_tex_as, "#313244", "#45475a", pad=(8, 3)).pack(side=tk.RIGHT, padx=2)
        self._btn(code_top, "📋 Copy Code", self.copy_tex, "#313244", "#45475a", pad=(8, 3)).pack(side=tk.RIGHT, padx=2)

        code_wrap = tk.Frame(right_card, bg=COLORS["code_bg"])
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
        cx = ttk.Scrollbar(right_card, orient=tk.HORIZONTAL, command=self.code_text.xview)
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

        # 4. Bottom Status & Telemetry Bar
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
        self.root.bind("<Control-n>", lambda e: self.open_item_modal("project"))

    # ------------------------------------------------------------------
    # Tailored Add / Edit Modal Forms (Matching Every Section in resume.tex)
    # ------------------------------------------------------------------
    def open_item_modal(self, item_type: str, edit_index: int | None = None):
        existing = self.items[edit_index] if edit_index is not None else {}

        dlg = tk.Toplevel(self.root)
        dlg.title(f"{'Edit' if edit_index is not None else 'Add'} {item_type.replace('_', ' ').title()}")
        dlg.geometry("720x620")
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

        if item_type == "project":
            add_entry(0, "Section Heading:", "section", "Personal Projects", all_secs)
            add_entry(1, "Bold Project Title:", "title", "Project Name, Short Tagline")
            add_entry(2, "Date Range (Italicized):", "date", "Feb 2026 -- Nov 2026")
            add_entry(3, "Italicized Domain/Subtitle:", "subtitle_italic", "Design, Development, Machine Learning, Deployment")
            add_entry(4, "Tech Stack (\\item Stack:):", "stack", "Python, PyTorch, ReactJS")

            # Multi-Link Builder for \item Access: \hrefWithoutArrow{url}{Label}
            link_frame = tk.LabelFrame(
                body, text="Access Links (Generates: \\item Access: \\hrefWithoutArrow{url}{Label})", bg=COLORS["bg_card"], padx=8, pady=6
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

            tk.Label(
                body, text="Additional Highlights (Optional, 1 per line):", bg=COLORS["bg_card"], font=("Segoe UI", 9, "bold")
            ).grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(6, 2))
            bullets_box = tk.Text(body, height=4, font=("Segoe UI", 9), bg="#f8fafc", relief=tk.SOLID, bd=1)
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
                body,
                text="Sub-role under preceding company (omit repeating company header)",
                variable=sub_v,
            ).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=4)

            tk.Label(body, text="Bullet Highlights (1 per line):", bg=COLORS["bg_card"], font=("Segoe UI", 9, "bold")).grid(
                row=6, column=0, columnspan=2, sticky=tk.W, pady=(6, 2)
            )
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
                body,
                text="Additional degree under same university above (omit repeating university title)",
                variable=sub_d,
            ).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=4)

            tk.Label(body, text="Education Highlights / Grades (1 per line):", bg=COLORS["bg_card"], font=("Segoe UI", 9, "bold")).grid(
                row=5, column=0, columnspan=2, sticky=tk.W, pady=(6, 2)
            )
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
            tk.Label(body, text="Bullet Lines (1 bullet line per row):", bg=COLORS["bg_card"], font=("Segoe UI", 9, "bold")).grid(
                row=2, column=0, columnspan=2, sticky=tk.W, pady=(8, 2)
            )
            bullets_box = tk.Text(body, height=10, font=("Segoe UI", 9), bg="#f8fafc", relief=tk.SOLID, bd=1)
            bullets_box.grid(row=3, column=0, columnspan=2, sticky=tk.NSEW)
            body.rowconfigure(3, weight=1)
            if existing.get("bullets"):
                bullets_box.insert("1.0", "\n".join(existing["bullets"]))

        else:  # custom section entry
            add_entry(0, "New or Existing Section Name:", "section", "Publications & Research", all_secs)
            add_entry(1, "Bold Primary Title:", "title", "")
            add_entry(2, "Italicized Subtitle (Optional):", "subtitle_italic", "")
            add_entry(3, "Date / Location (Right Column):", "date", "2026")
            tk.Label(body, text="Bullet Highlights (1 per line):", bg=COLORS["bg_card"], font=("Segoe UI", 9, "bold")).grid(
                row=4, column=0, columnspan=2, sticky=tk.W, pady=(6, 2)
            )
            bullets_box = tk.Text(body, height=8, font=("Segoe UI", 9), bg="#f8fafc", relief=tk.SOLID, bd=1)
            bullets_box.grid(row=5, column=0, columnspan=2, sticky=tk.NSEW)
            body.rowconfigure(5, weight=1)
            if existing.get("bullets"):
                bullets_box.insert("1.0", "\n".join(existing["bullets"]))

        footer = tk.Frame(dlg, bg="#f8fafc", padx=20, pady=12)
        footer.pack(fill=tk.X, side=tk.BOTTOM)

        def save_modal():
            new_item = {
                "enabled": existing.get("enabled", True),
                "type": item_type,
            }
            for k, v in vars_map.items():
                new_item[k] = v.get().strip()
            for k, bv in bool_vars.items():
                new_item[k] = bv.get()

            if item_type == "project":
                new_item["links"] = links_list
                new_item["extra_bullets"] = (
                    [ln.strip() for ln in bullets_box.get("1.0", tk.END).splitlines() if ln.strip()]
                    if bullets_box
                    else []
                )
            elif item_type == "education":
                new_item["links"] = existing.get("links", [])
                new_item["bullets"] = (
                    [ln.strip() for ln in bullets_box.get("1.0", tk.END).splitlines() if ln.strip()]
                    if bullets_box
                    else []
                )
            elif bullets_box is not None:
                new_item["bullets"] = [
                    ln.strip() for ln in bullets_box.get("1.0", tk.END).splitlines() if ln.strip()
                ]

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
        dlg.geometry("700x460")
        dlg.configure(bg=COLORS["bg_card"])
        dlg.transient(self.root)
        dlg.grab_set()

        body = tk.Frame(dlg, bg=COLORS["bg_card"], padx=18, pady=14)
        body.pack(fill=tk.BOTH, expand=True)

        top_row = tk.Frame(body, bg=COLORS["bg_card"])
        top_row.pack(fill=tk.X, pady=(0, 8))

        en_var = tk.BooleanVar(value=self.profile_enabled)
        ttk.Checkbutton(top_row, text="Include 'Professional Profile' Section in CV", variable=en_var).pack(side=tk.LEFT)

        prof_names = [p["name"] for p in self.profiles]
        sel_var = tk.StringVar(value=prof_names[self.active_profile_idx])
        combo = ttk.Combobox(top_row, textvariable=sel_var, values=prof_names, width=42, state="readonly")
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
    # Table Row Operations
    # ------------------------------------------------------------------
    def _selected_real_idx(self) -> int | None:
        sel = self.tree.selection()
        if not sel:
            return None
        t_idx = self.tree.index(sel[0])
        return self.visible_indices[t_idx] if 0 <= t_idx < len(self.visible_indices) else None

    def on_tree_select(self, _e=None):
        r_idx = self._selected_real_idx()
        self.code_text.tag_remove("active_block", "1.0", tk.END)
        if r_idx is not None and r_idx in self.item_line_map:
            s_ln, e_ln = self.item_line_map[r_idx]
            self.code_text.tag_add("active_block", f"{s_ln}.0", f"{e_ln}.end")
            self.code_text.see(f"{max(1, s_ln - 3)}.0")

    def edit_selected(self):
        r_idx = self._selected_real_idx()
        if r_idx is not None:
            self.open_item_modal(self.items[r_idx].get("type", "custom"), edit_index=r_idx)

    def toggle_selected(self):
        r_idx = self._selected_real_idx()
        if r_idx is not None:
            self.items[r_idx]["enabled"] = not self.items[r_idx].get("enabled", True)
            self.refresh_all(compile_pdf=self.auto_compile_var.get(), select_real_idx=r_idx)

    def clone_selected(self):
        r_idx = self._selected_real_idx()
        if r_idx is not None:
            c = copy.deepcopy(self.items[r_idx])
            for k in ("title", "company", "institution", "category"):
                if k in c:
                    c[k] = f"{c[k]} (Copy)"
                    break
            self.items.insert(r_idx + 1, c)
            self.refresh_all(compile_pdf=self.auto_compile_var.get(), select_real_idx=r_idx + 1)

    def move_selected(self, delta: int):
        r_idx = self._selected_real_idx()
        if r_idx is not None and 0 <= r_idx + delta < len(self.items):
            self.items[r_idx], self.items[r_idx + delta] = self.items[r_idx + delta], self.items[r_idx]
            self.refresh_all(compile_pdf=self.auto_compile_var.get(), select_real_idx=r_idx + delta)

    def delete_selected(self):
        r_idx = self._selected_real_idx()
        if r_idx is not None:
            del self.items[r_idx]
            self.refresh_all(compile_pdf=self.auto_compile_var.get())

    # ------------------------------------------------------------------
    # Exact LaTeX Generator (Matching Your resume.tex Byte-for-Byte Style)
    # ------------------------------------------------------------------
    def generate_latex(self) -> str:
        preamble = EXACT_PREAMBLE.replace(
            "Joseph Emmanuel Chay Huan Qin", smart_escape_latex(self.header.get("footer_name", ""))
        )
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

        # Group items by Section
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
                lines.append(r"    % PROFESSIONAL PROFILE")
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
                        for b in all_b:
                            lines.append(rf"            \item {smart_escape_latex(b)}")
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
                        for b in it["bullets"]:
                            lines.append(rf"            \item {smart_escape_latex(b)}")
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
                    for eb in it.get("extra_bullets", []):
                        lines.append(rf"            \item {smart_escape_latex(eb)}")
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
                    for b in it.get("bullets", []):
                        lines.append(rf"            \item {smart_escape_latex(b)}")
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
                        for b in it["bullets"]:
                            lines.append(rf"            \item {smart_escape_latex(b)}")
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
    # Refresh, Save, Compile & Log Inspection
    # ------------------------------------------------------------------
    def refresh_all(self, compile_pdf: bool = False, select_real_idx: int | None = None):
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
        iid_to_select = None

        for r_idx, it in enumerate(self.items):
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
            en = it.get("enabled", True)
            iid = self.tree.insert(
                "",
                tk.END,
                values=("✓" if en else "✗", sec, title_disp, detail_disp, date_disp),
                tags=("on" if en else "off",),
            )
            if select_real_idx == r_idx:
                iid_to_select = iid

        active_cnt = sum(1 for i in self.items if i.get("enabled", True))
        self.stats_lbl.config(text=f"Active Entries: {active_cnt} / {len(self.items)}")

        tex = self.generate_latex()
        self._highlight_syntax(tex)

        if iid_to_select:
            self.tree.selection_set(iid_to_select)
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
                json.dump(
                    {
                        "header": self.header,
                        "profiles": self.profiles,
                        "active_profile_idx": self.active_profile_idx,
                        "profile_enabled": self.profile_enabled,
                        "section_order": self.section_order,
                        "items": self.items,
                    },
                    f,
                    indent=2,
                    ensure_ascii=False,
                )
        except Exception:
            pass

    def _load_state(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    d = json.load(f)
                self.header = d.get("header", self.header)
                self.profiles = d.get("profiles", self.profiles)
                self.active_profile_idx = d.get("active_profile_idx", self.active_profile_idx)
                self.profile_enabled = d.get("profile_enabled", self.profile_enabled)
                self.section_order = d.get("section_order", self.section_order)
                self.items = d.get("items", self.items)
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
                if proc.returncode == 0:
                    self.root.after(0, lambda: self.status_lbl.config(text="✓ Successfully compiled resume.pdf!"))
                else:
                    self.root.after(
                        0, lambda: self.status_lbl.config(text=f"⚠ pdflatex finished with code {proc.returncode}.")
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
        if not os.path.exists("resume.pdf"):
            messagebox.showinfo("No PDF Yet", "Click 'Compile PDF' first to build resume.pdf.")
            return
        if sys.platform.startswith("win"):
            os.startfile("resume.pdf")
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "resume.pdf"])
        else:
            subprocess.Popen(["xdg-open", "resume.pdf"])


if __name__ == "__main__":
    root = tk.Tk()
    app = JosephResumeStudio(root)
    root.mainloop()
    
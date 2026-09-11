"""Comprehensive OCR & Extraction Benchmark on 5 Distinct Candidate Layouts.

Compares:
1. Standard PyMuPDF (Baseline)
2. PyMuPDF4LLM (Artifex Layout & Markdown)
3. Microsoft MarkItDown (Office/Doc Parser)
4. RapidOCR (PaddleOCR ONNX Vision Engine)
5. Tesseract OCR (Current Legacy Fallback)
"""

import io
import os
import sys
import time
from pathlib import Path

# Force utf-8 encoding on Windows console
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

import pymupdf
from PIL import Image

APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))

from hiring_agent.extraction import extract_candidate_details_smart, extract_text_from_bytes
from hiring_agent.local_scorer import get_match_details
from hiring_agent.jd_sources import get_active_roles

# Create benchmark directory
BENCH_DIR = APP_DIR / "benchmark_resumes"
BENCH_DIR.mkdir(exist_ok=True)


def create_candidate_1_multicolumn():
    """Candidate 1: Alex Chen - 2-Column Sidebar Layout.
    Left: Skills, Contact, Education. Right: Experience & Projects.
    Challenge: Multi-column horizontal bleed in traditional OCR.
    """
    path = BENCH_DIR / "01_alex_chen_2column.pdf"
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842) # A4

    # Header
    page.insert_text((50, 50), "Alex Chen", fontsize=18, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 70), "Senior AI & Machine Learning Engineer", fontsize=12, fontname="helv", color=(0.3, 0.3, 0.3))
    page.draw_line((50, 80), (545, 80), color=(0.7, 0.7, 0.7), width=1)

    # Left Column (x=50 to 200): Contact & Skills
    page.insert_text((50, 105), "CONTACT", fontsize=11, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 125), "Phone: (206) 555-0144", fontsize=9, fontname="helv")
    page.insert_text((50, 140), "Email: alex.chen@ai.io", fontsize=9, fontname="helv")
    page.insert_text((50, 155), "Location: Seattle, WA", fontsize=9, fontname="helv")
    page.insert_text((50, 170), "Country: United States", fontsize=9, fontname="helv")

    page.insert_text((50, 205), "CORE SKILLS", fontsize=11, fontname="helv", color=(0.1, 0.2, 0.5))
    skills = ["Python", "PyTorch", "TensorFlow", "LangChain", "RAG", "Ollama", "FastAPI", "Docker", "Kubernetes", "AWS", "SQL"]
    for i, s in enumerate(skills):
        page.insert_text((50, 225 + i * 16), f"• {s}", fontsize=9, fontname="helv")

    page.insert_text((50, 420), "EDUCATION", fontsize=11, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 440), "M.S. Computer Science", fontsize=9, fontname="helv")
    page.insert_text((50, 455), "Stanford University", fontsize=8, fontname="helv")
    page.insert_text((50, 470), "2020 - 2022", fontsize=8, fontname="helv")

    # Right Column (x=220 to 545): Experience
    page.insert_text((220, 105), "PROFESSIONAL EXPERIENCE", fontsize=11, fontname="helv", color=(0.1, 0.2, 0.5))
    
    page.insert_text((220, 125), "Lead AI Engineer | DriverAI Corp (Seattle, WA)", fontsize=10, fontname="helv")
    page.insert_text((220, 140), "August 2022 - Present", fontsize=8, fontname="helv", color=(0.4, 0.4, 0.4))
    exp1 = [
        "Architected enterprise RAG pipelines reducing query latency by 40% using vector databases.",
        "Deployed fine-tuned LLM agents with LangChain, FastAPI, and Docker into production on AWS.",
        "Built automated evaluation benchmarks ensuring 99% structured JSON output accuracy."
    ]
    for i, bullet in enumerate(exp1):
        page.insert_text((230, 160 + i * 18), f"- {bullet}", fontsize=8.5, fontname="helv")

    page.insert_text((220, 235), "Machine Learning Engineer | TechVision Inc (San Jose, CA)", fontsize=10, fontname="helv")
    page.insert_text((220, 250), "June 2020 - July 2022", fontsize=8, fontname="helv", color=(0.4, 0.4, 0.4))
    exp2 = [
        "Trained computer vision and NLP models using PyTorch on multi-GPU Kubernetes clusters.",
        "Integrated real-time model inference endpoints with FastAPI and Redis caching.",
        "Collaborated with data engineering to optimize SQL ETL pipelines handling 10M daily events."
    ]
    for i, bullet in enumerate(exp2):
        page.insert_text((230, 270 + i * 18), f"- {bullet}", fontsize=8.5, fontname="helv")

    doc.save(path)
    doc.close()
    return path


def create_candidate_2_singlecolumn():
    """Candidate 2: Marcus Vance - Single-Column Backend Engineer.
    Classic linear format.
    """
    path = BENCH_DIR / "02_marcus_vance_singlecolumn.pdf"
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)

    page.insert_text((50, 50), "Marcus Vance", fontsize=18, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 70), "Austin, TX | (512) 555-0199 | marcus.vance@backend.dev | USA", fontsize=9, fontname="helv")
    page.draw_line((50, 80), (545, 80), color=(0.7, 0.7, 0.7), width=1)

    page.insert_text((50, 105), "PROFESSIONAL SUMMARY", fontsize=11, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 125), "Principal Backend Engineer with 8+ years building high-concurrency microservices in Go and Python.", fontsize=9, fontname="helv")

    page.insert_text((50, 155), "TECHNICAL SKILLS", fontsize=11, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 175), "Languages: Python, Go, SQL, Bash", fontsize=9, fontname="helv")
    page.insert_text((50, 190), "Frameworks & Cloud: FastAPI, Flask, PostgreSQL, Redis, AWS, SQS, Terraform, Docker, Kubernetes", fontsize=9, fontname="helv")

    page.insert_text((50, 220), "EXPERIENCE", fontsize=11, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 240), "Principal Cloud Architect — CloudScale Solutions (Austin, TX) | 2021 - Present", fontsize=10, fontname="helv")
    page.insert_text((60, 260), "• Designed event-driven architectures with Amazon SQS, SNS, and FastAPI handling 50k req/sec.", fontsize=8.5, fontname="helv")
    page.insert_text((60, 278), "• Automated multi-tenant infrastructure using Terraform and Kubernetes (EKS).", fontsize=8.5, fontname="helv")
    page.insert_text((60, 296), "• Migrated legacy monolith to decoupled PostgreSQL microservices with 99.99% uptime.", fontsize=8.5, fontname="helv")

    page.insert_text((50, 330), "EDUCATION", fontsize=11, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 350), "B.S. in Computer Engineering, University of Texas at Austin (2014 - 2018)", fontsize=9, fontname="helv")

    doc.save(path)
    doc.close()
    return path


def create_candidate_3_table_layout():
    """Candidate 3: Priya Patel - Full-Stack Developer with Tabular Layout.
    Tables and boxed structures.
    """
    path = BENCH_DIR / "03_priya_patel_table.pdf"
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)

    page.insert_text((50, 50), "Priya Patel", fontsize=18, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 70), "San Francisco, California | (415) 555-0182 | priya.patel@fullstack.io", fontsize=9, fontname="helv")
    page.draw_line((50, 80), (545, 80), color=(0.7, 0.7, 0.7), width=1)

    page.insert_text((50, 105), "SKILLS MATRIX", fontsize=11, fontname="helv", color=(0.1, 0.2, 0.5))
    # Table box
    page.draw_rect(pymupdf.Rect(50, 115, 545, 175), color=(0.7, 0.7, 0.7), width=1)
    page.draw_line((200, 115), (200, 175), color=(0.7, 0.7, 0.7), width=1)
    page.insert_text((60, 135), "Frontend Technologies:", fontsize=9, fontname="helv")
    page.insert_text((210, 135), "React, TypeScript, Next.js, Redux, Tailwind CSS, HTML5", fontsize=9, fontname="helv")
    page.insert_text((60, 155), "Backend & Databases:", fontsize=9, fontname="helv")
    page.insert_text((210, 155), "Node.js, Express, Python, GraphQL, PostgreSQL, MongoDB", fontsize=9, fontname="helv")

    page.insert_text((50, 205), "WORK HISTORY", fontsize=11, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 225), "Senior Full-Stack Engineer | FinTech Innovations (San Francisco, CA) | 2021 - Present", fontsize=9.5, fontname="helv")
    page.insert_text((60, 245), "• Built customer-facing financial dashboards with React, TypeScript, and high-performance GraphQL.", fontsize=8.5, fontname="helv")
    page.insert_text((60, 263), "• Implemented secure authentication and session workflows using OAuth2 and JWT.", fontsize=8.5, fontname="helv")

    doc.save(path)
    doc.close()
    return path


def create_candidate_4_dense_devops():
    """Candidate 4: Jordan Miller - Dense DevOps / SRE Resume."""
    path = BENCH_DIR / "04_jordan_miller_devops.pdf"
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)

    page.insert_text((50, 45), "Jordan Miller", fontsize=18, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 65), "Chicago, IL | Phone: (312) 555-0143 | jordan.m@infraops.net | USA", fontsize=9, fontname="helv")
    page.draw_line((50, 75), (545, 75), color=(0.7, 0.7, 0.7), width=1)

    page.insert_text((50, 95), "TECHNICAL PROFICIENCIES", fontsize=10, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 112), "DevOps: Kubernetes, Docker, Helm, Terraform, Ansible, Jenkins, GitHub Actions, ArgoCD", fontsize=8.5, fontname="helv")
    page.insert_text((50, 127), "Cloud & Monitoring: AWS (EC2, S3, EKS, RDS, IAM), GCP, Prometheus, Grafana, Datadog", fontsize=8.5, fontname="helv")
    page.insert_text((50, 142), "Languages & Scripting: Python, Go, Bash, Linux shell scripting, YAML, JSON", fontsize=8.5, fontname="helv")

    page.insert_text((50, 168), "WORK EXPERIENCE", fontsize=10, fontname="helv", color=(0.1, 0.2, 0.5))
    page.insert_text((50, 185), "Staff Site Reliability Engineer — Global Logistics Platform (Chicago, IL) | 2020 - 2026", fontsize=9, fontname="helv")
    page.insert_text((60, 203), "• Maintained 15+ production Kubernetes clusters running 400+ microservices on AWS.", fontsize=8, fontname="helv")
    page.insert_text((60, 218), "• Provisioned all cloud infrastructure declaratively using Terraform and GitOps practices.", fontsize=8, fontname="helv")
    page.insert_text((60, 233), "• Decreased incident MTTR by 45% through automated alerts with Prometheus and Grafana.", fontsize=8, fontname="helv")

    doc.save(path)
    doc.close()
    return path


def create_candidate_5_scanned_image():
    """Candidate 5: Elena Rostova - Pure Scanned / Image PDF.
    Has ZERO digital text layer! Rendered as an image.
    This strictly tests visual OCR models (Tesseract vs RapidOCR).
    """
    path = BENCH_DIR / "05_elena_rostova_scanned.pdf"
    
    # First create a temporary digital page
    temp_doc = pymupdf.open()
    p = temp_doc.new_page(width=595, height=842)
    p.insert_text((50, 50), "Elena Rostova", fontsize=18, fontname="helv", color=(0.1, 0.2, 0.5))
    p.insert_text((50, 70), "New York, NY | (212) 555-0177 | elena.rostova@datascience.org | USA", fontsize=9, fontname="helv")
    p.draw_line((50, 80), (545, 80), color=(0.7, 0.7, 0.7), width=1)
    
    p.insert_text((50, 105), "Data Scientist & Analytics Specialist", fontsize=11, fontname="helv")
    p.insert_text((50, 130), "Skills: Python, SQL, R, Pandas, NumPy, Scikit-Learn, Tableau, Machine Learning, Statistics", fontsize=9, fontname="helv")
    p.insert_text((50, 160), "Experience: Senior Data Scientist at Metrix AI (New York, NY) | 2021 - Present", fontsize=9, fontname="helv")
    p.insert_text((60, 180), "• Developed predictive ML models using Python and Scikit-Learn to forecast customer churn.", fontsize=8.5, fontname="helv")
    p.insert_text((60, 200), "• Designed automated analytics dashboards in Tableau connected to PostgreSQL data warehouse.", fontsize=8.5, fontname="helv")

    # Render to 150 DPI image pixmap
    pix = p.get_pixmap(dpi=150)
    img_bytes = pix.tobytes("png")
    temp_doc.close()

    # Now create a new PDF that ONLY contains this rendered image (no digital text layer)
    doc = pymupdf.open()
    img_page = doc.new_page(width=595, height=842)
    img_page.insert_image(pymupdf.Rect(0, 0, 595, 842), stream=img_bytes)
    doc.save(path)
    doc.close()
    return path


def run_benchmark():
    print("\n" + "=" * 80)
    print("  OCR & EXTRACTION BENCHMARK: 5 CANDIDATES ACROSS 4 ENGINES")
    print("=" * 80)

    # Generate the 5 candidate PDFs
    files = [
        ("01_Alex_Chen (2-Column Sidebar)", create_candidate_1_multicolumn()),
        ("02_Marcus_Vance (Single-Column Linear)", create_candidate_2_singlecolumn()),
        ("03_Priya_Patel (Tabular / Boxes)", create_candidate_3_table_layout()),
        ("04_Jordan_Miller (Dense DevOps)", create_candidate_4_dense_devops()),
        ("05_Elena_Rostova (SCANNED IMAGE - Zero Text Layer)", create_candidate_5_scanned_image()),
    ]

    active_roles = get_active_roles()
    
    # Initialize OCR engines
    from rapidocr_onnxruntime import RapidOCR
    rapid_engine = RapidOCR()

    import pytesseract
    from markitdown import MarkItDown
    markitdown_engine = MarkItDown()

    import pymupdf4llm

    results = []

    for name, pdf_path in files:
        print("\n" + "-" * 72)
        print(f"  CANDIDATE: {name}")
        print("-" * 72)
        raw_bytes = pdf_path.read_bytes()

        engines = {}

        # 1. Standard PyMuPDF (extract_text_from_bytes)
        t0 = time.perf_counter()
        text_pymupdf = extract_text_from_bytes(raw_bytes, pdf_path.name)
        dt_pymupdf = (time.perf_counter() - t0) * 1000
        engines["PyMuPDF (Baseline)"] = (text_pymupdf, dt_pymupdf)

        # 2. PyMuPDF4LLM
        t0 = time.perf_counter()
        try:
            doc = pymupdf.open(stream=raw_bytes, filetype="pdf")
            text_mupdf4llm = pymupdf4llm.to_markdown(doc)
            doc.close()
        except Exception as e:
            text_mupdf4llm = ""
        dt_mupdf4llm = (time.perf_counter() - t0) * 1000
        engines["PyMuPDF4LLM (Layout)"] = (text_mupdf4llm, dt_mupdf4llm)

        # 3. Microsoft MarkItDown
        t0 = time.perf_counter()
        try:
            res = markitdown_engine.convert(str(pdf_path))
            text_markitdown = res.text_content
        except Exception as e:
            text_markitdown = ""
        dt_markitdown = (time.perf_counter() - t0) * 1000
        engines["MarkItDown (Microsoft)"] = (text_markitdown, dt_markitdown)

        # 4. RapidOCR (PaddleOCR ONNX)
        t0 = time.perf_counter()
        try:
            doc = pymupdf.open(stream=raw_bytes, filetype="pdf")
            rapid_lines = []
            for page in doc:
                pix = page.get_pixmap(dpi=150)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                ocr_res, _ = rapid_engine(img)
                if ocr_res:
                    rapid_lines.extend([line[1] for line in ocr_res])
            doc.close()
            text_rapid = "\n".join(rapid_lines)
        except Exception as e:
            text_rapid = f"RapidOCR Error: {e}"
        dt_rapid = (time.perf_counter() - t0) * 1000
        engines["RapidOCR (Vision ONNX)"] = (text_rapid, dt_rapid)

        # 5. Tesseract OCR
        t0 = time.perf_counter()
        try:
            doc = pymupdf.open(stream=raw_bytes, filetype="pdf")
            tess_lines = []
            for page in doc:
                pix = page.get_pixmap(dpi=150)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                tess_lines.append(pytesseract.image_to_string(img))
            doc.close()
            text_tess = "\n".join(tess_lines)
        except Exception as e:
            text_tess = f"Tesseract Error: {e}"
        dt_tess = (time.perf_counter() - t0) * 1000
        engines["Tesseract (Legacy OCR)"] = (text_tess, dt_tess)

        # Compare Extractions & Scoring
        for engine_name, (txt, latency) in engines.items():
            if not txt.strip():
                print(f"  [{engine_name:<22}] Latency: {latency:6.1f}ms | FAILED (No text extracted)")
                continue

            cand = extract_candidate_details_smart(txt)
            skills = cand.get("skills", "")
            matches = [get_match_details(skills, r) for r in active_roles]
            matches.sort(key=lambda m: -m["score"])
            top3 = matches[:3]

            top3_str = ", ".join([f"{m['role_title']} ({m['score']}%)" for m in top3])
            matched_skills = top3[0].get("matched", []) if top3 else []
            
            print(f"  [{engine_name:<22}] Latency: {latency:6.1f}ms | Name: {cand.get('full_name', '?'):<14} | Loc: {cand.get('location', '?'):<14} | Phone: {cand.get('phone', '?')}")
            print(f"    --> Top 3: {top3_str}")
            print(f"    --> Matched Skills ({len(matched_skills)}): {', '.join(matched_skills[:8])}")

    print("\n" + "=" * 80)
    print("  BENCHMARK COMPLETE")
    print("=" * 80)


if __name__ == "__main__":
    run_benchmark()

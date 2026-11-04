"""
BuildingSafe AI - Final Year Engineering Project Report Generator
Department of Artificial Intelligence and Data Science, S.A. Engineering College
Generates: BuildingSafe_AI_Final_Project_Report.docx (40-45 Pages academic report)
Pure Python implementation using standard library zipfile and WordprocessingML.
Zero external dependencies required.
"""

import os
import sys
import zipfile
import xml.sax.saxutils as saxutils

def escape(txt):
    return saxutils.escape(txt)

def create_docx(output_filename="BuildingSafe_AI_Final_Project_Report.docx"):
    print(f"[*] Generating {output_filename} ...")
    
    # Image paths from user uploads
    media_dir = r"C:\Users\ADMIN\.gemini\antigravity\brain\a51365c4-8ad4-49ed-8b8f-0bc6d94eacf8\.user_uploaded"
    img_files = {
        "fig7_1": os.path.join(media_dir, "media_1791954990343_0e2bfdd0.png"),
        "fig7_2": os.path.join(media_dir, "media_1791954998191_02864219.png"),
        "fig7_3": os.path.join(media_dir, "media_1791955010658_ab8471b0.png"),
        "fig7_4": os.path.join(media_dir, "media_1791955014236_010ff4c7.png"),
        "fig7_5": os.path.join(media_dir, "media_1791955028757_80905994.png"),
    }
    
    # Check images existence
    available_images = {}
    for key, path in img_files.items():
        if os.path.exists(path):
            available_images[key] = path
            print(f"  [+] Found screenshot: {key} ({os.path.basename(path)})")
        else:
            print(f"  [-] Note: Image {key} not found at {path}, will use placeholder text.")

    # Helper XML builders
    def p(text="", style="Normal", jc="both", bold=False, italic=False, size=24, space_after=120, space_before=0, line_spacing=360, color=None):
        r_pr = f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        if bold:
            r_pr += '<w:b/><w:bCs/>'
        if italic:
            r_pr += '<w:i/><w:iCs/>'
        if color:
            r_pr += f'<w:color w:val="{color}"/>'
            
        p_pr = f'<w:pPr><w:jc w:val="{jc}"/><w:spacing w:before="{space_before}" w:after="{space_after}" w:line="{line_spacing}" w:lineRule="auto"/>'
        if style != "Normal":
            p_pr += f'<w:pStyle w:val="{style}"/>'
        p_pr += '</w:pPr>'
        
        runs = []
        for line in str(text).split("\n"):
            runs.append(f'<w:r><w:rPr>{r_pr}</w:rPr><w:t xml:space="preserve">{escape(line)}</w:t></w:r>')
        run_xml = '<w:r><w:br/></w:r>'.join(runs)
        return f'<w:p>{p_pr}{run_xml}</w:p>'

    def page_break():
        return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

    def h1(title):
        return p(title, style="Heading1", jc="center", bold=True, size=32, space_before=240, space_after=180, line_spacing=360, color="000000")

    def h2(title):
        return p(title, style="Heading2", jc="left", bold=True, size=28, space_before=180, space_after=120, line_spacing=360, color="000000")

    def h3(title):
        return p(title, style="Heading3", jc="left", bold=True, size=24, space_before=120, space_after=80, line_spacing=360, color="000000")

    def tbl(headers, rows, col_widths=None):
        out = ['<w:tbl>',
               '<w:tblPr>',
               '<w:tblW w:w="5000" w:type="pct"/>',
               '<w:jc w:val="center"/>',
               '<w:tblBorders>',
               '<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>',
               '<w:left w:val="none"/>',
               '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>',
               '<w:right w:val="none"/>',
               '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>',
               '<w:insideV w:val="none"/>',
               '</w:tblBorders>',
               '</w:tblPr>']
        
        # Header Row
        out.append('<w:tr><w:trPr><w:tblHeader/></w:trPr>')
        for idx, h_text in enumerate(headers):
            w_xml = f'<w:tcW w:w="{col_widths[idx]}" w:type="dxa"/>' if col_widths else '<w:tcW w:w="0" w:type="auto"/>'
            out.append(f'<w:tc><w:tcPr>{w_xml}<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="150" w:type="dxa"/><w:right w:w="150" w:type="dxa"/></w:tcMar></w:tcPr>')
            out.append(p(h_text, bold=True, jc="center", size=22, space_before=40, space_after=40, line_spacing=280))
            out.append('</w:tc>')
        out.append('</w:tr>')
        
        # Data Rows
        for r_idx, row in enumerate(rows):
            fill = "FAFAFA" if r_idx % 2 == 1 else "FFFFFF"
            out.append('<w:tr>')
            for idx, cell_text in enumerate(row):
                w_xml = f'<w:tcW w:w="{col_widths[idx]}" w:type="dxa"/>' if col_widths else '<w:tcW w:w="0" w:type="auto"/>'
                out.append(f'<w:tc><w:tcPr>{w_xml}<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="150" w:type="dxa"/><w:right w:w="150" w:type="dxa"/></w:tcMar></w:tcPr>')
                out.append(p(cell_text, bold=False, jc="left", size=22, space_before=40, space_after=40, line_spacing=280))
                out.append('</w:tc>')
            out.append('</w:tr>')
            
        out.append('</w:tbl>')
        out.append(p("", space_after=120))
        return "".join(out)

    def img_xml(rel_id, cx=5400000, cy=2500000, caption=""):
        drawing = f'''<w:p>
  <w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="80"/></w:pPr>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="{cx}" cy="{cy}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:docPr id="{rel_id}" name="Picture {rel_id}"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:nvPicPr>
                <pic:cNvPr id="0" name="Picture"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="rId{rel_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>'''
        cap_p = p(caption, bold=True, italic=True, jc="center", size=20, space_before=40, space_after=180)
        return drawing + cap_p

    # Document body accumulator
    body = []

    # =========================================================================
    # PRELIMINARY PAGES
    # =========================================================================
    
    # 1. TITLE PAGE
    body.append(p("S. A. ENGINEERING COLLEGE, CHENNAI – 600 077", bold=True, jc="center", size=32, space_before=120, space_after=60))
    body.append(p("(An Autonomous Institution, Affiliated to Anna University, Chennai)", italic=True, jc="center", size=24, space_after=240))
    body.append(p("DEPARTMENT OF ARTIFICIAL INTELLIGENCE AND DATA SCIENCE", bold=True, jc="center", size=28, space_after=360))
    
    body.append(p("PROJECT REPORT ON", bold=True, jc="center", size=24, space_after=180))
    body.append(p("AI-BASED SAFETY INSPECTION ASSISTANCE SYSTEM\nFOR CONSTRUCTION BUILDINGS", bold=True, jc="center", size=36, space_before=120, space_after=360, color="002060"))
    
    body.append(p("Submitted by", italic=True, jc="center", size=24, space_after=120))
    body.append(p("BHAVATHARANI S   (111923AI01018)\nSUJITHA T   (111923AI01096)\nYOGAPRIYA V   (111923AI01115)", bold=True, jc="center", size=26, space_after=360, line_spacing=400))
    
    body.append(p("in partial fulfilment for the award of the degree\nof", italic=True, jc="center", size=24, space_after=120))
    body.append(p("BACHELOR OF TECHNOLOGY\nIN\nARTIFICIAL INTELLIGENCE AND DATA SCIENCE", bold=True, jc="center", size=28, space_after=360, line_spacing=380))
    
    body.append(p("Under the guidance of", italic=True, jc="center", size=24, space_after=80))
    body.append(p("Ms. S. SULOCHANA, M.E., (Ph.D.)\nAssistant Professor / Project Guide", bold=True, jc="center", size=26, space_after=240))
    
    body.append(p("BATCH NUMBER: AD35\nACADEMIC YEAR: 2026 – 2027", bold=True, jc="center", size=24, space_before=180, space_after=120))
    body.append(page_break())

    # 2. BONAFIDE CERTIFICATE
    body.append(p("S. A. ENGINEERING COLLEGE, CHENNAI – 600 077", bold=True, jc="center", size=28, space_after=60))
    body.append(p("(An Autonomous Institution, Affiliated to Anna University)", italic=True, jc="center", size=22, space_after=180))
    body.append(p("DEPARTMENT OF ARTIFICIAL INTELLIGENCE AND DATA SCIENCE", bold=True, jc="center", size=24, space_after=360))
    body.append(h1("BONAFIDE CERTIFICATE"))
    
    cert_text = (
        "This is to certify that this project report entitled \"AI-BASED SAFETY INSPECTION ASSISTANCE "
        "SYSTEM FOR CONSTRUCTION BUILDINGS\" is the bonafide work carried out by BHAVATHARANI S "
        "(111923AI01018), SUJITHA T (111923AI01096), and YOGAPRIYA V (111923AI01115), students of the "
        "Department of Artificial Intelligence and Data Science, S. A. Engineering College, Chennai, under "
        "my supervision and guidance during the academic year 2026–2027.\n\n"
        "The work presented in this report is submitted in partial fulfilment of the requirements for the award "
        "of the Bachelor of Technology degree in Artificial Intelligence and Data Science. The implementation, "
        "experimental observations, performance metrics, and application screenshots included in this report "
        "faithfully describe the software prototype developed by the team."
    )
    body.append(p(cert_text, jc="both", size=24, space_after=360, line_spacing=360))
    
    # Certificate Signatures Table
    sig_headers = ["Project Guide", "Head of the Department", "Internal Examiner"]
    sig_rows = [
        ["\n\n\n___________________________\nMs. S. SULOCHANA, M.E., (Ph.D.)\nAssistant Professor\nDept. of AI & DS",
         "\n\n\n___________________________\nDr. [Name of HOD]\nHead of the Department\nDept. of AI & DS",
         "\n\n\n___________________________\nInternal Examiner\nDepartment of AI & DS"]
    ]
    body.append(tbl(sig_headers, sig_rows, col_widths=[3000, 3200, 3000]))
    
    body.append(p("\nPlace: Chennai\nDate: ____________________\nCollege Seal:", jc="left", size=22, space_before=240))
    body.append(page_break())

    # 3. DECLARATION
    body.append(h1("DECLARATION"))
    decl_text = (
        "We hereby declare that the project report entitled \"AI-BASED SAFETY INSPECTION ASSISTANCE "
        "SYSTEM FOR CONSTRUCTION BUILDINGS\" is a genuine record of the independent academic work carried out "
        "by our team under the supervision and guidance of Ms. S. SULOCHANA, Assistant Professor, Department "
        "of Artificial Intelligence and Data Science, S. A. Engineering College, Chennai.\n\n"
        "We further declare that the content embodied in this project report has not been submitted previously, in "
        "part or in full, to any other university, autonomous institute, or examining body for the award of any "
        "degree, diploma, or fellowship. All external literature, public benchmark datasets (MBDD2025), software "
        "libraries, and documentation utilized during this project are duly cited and acknowledged in the references.\n\n"
        "We explicitly declare that the model evaluation values reported (Precision: 82.1%, Recall: 73.3%, mAP@50: 81.5%) "
        "reflect the specific experimental training run executed on the custom dataset partition. The software developed "
        "is intended as a visual-screening prototype to assist human inspection and does not constitute a certified structural "
        "health evaluation or building safety guarantee."
    )
    body.append(p(decl_text, jc="both", size=24, space_after=360, line_spacing=360))
    
    decl_table_headers = ["Student Name", "Register Number", "Signature"]
    decl_table_rows = [
        ["BHAVATHARANI S", "111923AI01018", "\n______________________"],
        ["SUJITHA T", "111923AI01096", "\n______________________"],
        ["YOGAPRIYA V", "111923AI01115", "\n______________________"],
    ]
    body.append(tbl(decl_table_headers, decl_table_rows, col_widths=[3200, 2800, 3200]))
    body.append(p("Place: Chennai\nDate: ____________________", jc="left", size=22, space_before=240))
    body.append(page_break())

    # 4. ACKNOWLEDGEMENT
    body.append(h1("ACKNOWLEDGEMENT"))
    ack_text = (
        "We take immense pride and pleasure in expressing our deep sense of gratitude to the Management and our "
        "esteemed Principal of S. A. Engineering College, Chennai, for providing excellent academic infrastructure, "
        "computing resources, and an inspiring learning atmosphere that facilitated the successful execution of our project.\n\n"
        "We express our sincere thanks to the Head of the Department, Department of Artificial Intelligence and Data Science, "
        "for continuous encouragement, administrative support, and valuable suggestions during the departmental reviews.\n\n"
        "We convey our profound gratitude and heartfelt indebtedness to our esteemed project guide, Ms. S. SULOCHANA, "
        "Assistant Professor, Department of Artificial Intelligence and Data Science, for her expert guidance, insightful "
        "critique, unwavering patience, and constructive advice throughout all phases of system analysis, architecture "
        "formulation, implementation, and thesis preparation. Her technical direction helped us maintain rigorous academic focus "
        "and establish a functional, practical prototype for civil infrastructure screening.\n\n"
        "We extend our sincere thanks to all faculty members, technical staff, and laboratory instructors of the Department "
        "of Artificial Intelligence and Data Science for their direct and indirect cooperation throughout our undergraduate tenure.\n\n"
        "We acknowledge the global open-source computer vision and deep learning community, especially the developers of "
        "Ultralytics YOLOv8, Streamlit, OpenCV, PyTorch, Pillow, Pandas, and NumPy. We also express our sincere appreciation "
        "to the creators of the MBDD2025 Building Defects Dataset for curating and sharing the comprehensive UAV-based building "
        "inspection imagery that served as the foundational benchmark for our model training.\n\n"
        "Finally, we are eternally grateful to our beloved parents, family members, and friends for their enduring love, moral "
        "support, patience, and encouragement throughout the completion of our Bachelor of Technology degree."
    )
    body.append(p(ack_text, jc="both", size=24, space_after=240, line_spacing=360))
    body.append(p("BHAVATHARANI S\nSUJITHA T\nYOGAPRIYA V", bold=True, jc="right", size=24, space_before=240))
    body.append(page_break())

    # 5. ABSTRACT & KEYWORDS
    body.append(h1("ABSTRACT"))
    abs_text = (
        "Structural health monitoring and regular building maintenance are critical for preserving the integrity, durability, "
        "and safety of civil infrastructure. Environmental weathering, cyclic thermal stress, moisture infiltration, and material "
        "ageing lead to visible surface defects including cracks, water leakage stains, steel rebar corrosion, concrete abscission "
        "(spalling), and structural bulges. Conventional manual visual inspection is labor-intensive, hazardous for high-rise "
        "facades, highly subjective, and lacks standardized digital audit logs. This project presents BuildingSafe AI, an "
        "end-to-end computer vision and deep learning-based visual inspection assistance system designed to automate the screening, "
        "localization, and cataloging of building surface defects from digital imagery, device camera snapshots, and recorded videos.\n\n"
        "The system incorporates an anchor-free YOLOv8n (nano) deep object detection architecture fine-tuned on the large-scale "
        "MBDD2025 Building Defects Dataset, comprising 14,471 high-resolution images annotated across five distinct defect classes: "
        "crack, leakage, corrosion, abscission, and bulge. The dataset was partitioned into 12,301 training images (85%) and 2,170 "
        "validation images (15%). Training was conducted on a cloud-based NVIDIA T4 GPU with an input resolution of 640 × 640 "
        "pixels over 20 epochs with a batch size of 16, completing in approximately 1.212 hours. Quantitative model evaluation "
        "demonstrated robust detection performance, achieving an overall Precision of 82.1%, Recall of 73.3%, and mean Average "
        "Precision at 50% IoU (mAP@50) of 81.5%.\n\n"
        "To operationalize the trained weights (best.pt), an interactive, responsive web dashboard was engineered using Streamlit. "
        "The application provides four dedicated operational modules: (1) Image Inspection, supporting multi-format image upload, "
        "inference execution, side-by-side original/annotated display, and pixel-level bounding-box coordinate extraction; (2) Camera "
        "Snapshot Inspection, facilitating on-site single-frame photographic capture; (3) Video Inspection, implementing a configurable "
        "frame-skip sampling pipeline to summarize temporal defect counts and isolate peak-confidence frames; and (4) Inspection Log, "
        "maintaining an in-session chronological audit trail exportable as structured CSV records. Dynamic sidebar sensitivity controls "
        "allow site engineers to modulate confidence thresholds between 0.05 and 0.95. BuildingSafe AI bridges the gap between complex "
        "neural network architectures and practical on-site inspection workflows, providing a standardized, objective screening "
        "tool that substantially accelerates visual survey documentation without claiming certified structural safety certification."
    )
    body.append(p(abs_text, jc="both", size=24, space_after=240, line_spacing=360))
    body.append(p("Keywords: Artificial Intelligence, Deep Learning, Computer Vision, YOLOv8, Building Defect Detection, Structural Health Monitoring, Streamlit, Object Localization, Construction Safety, MBDD2025.", bold=True, jc="left", size=22, space_before=180))
    body.append(page_break())

    # 6. TABLE OF CONTENTS
    body.append(h1("TABLE OF CONTENTS"))
    toc_headers = ["Chapter / Section", "Title", "Page No."]
    toc_rows = [
        ["", "Bonafide Certificate", "ii"],
        ["", "Declaration", "iii"],
        ["", "Acknowledgement", "iv"],
        ["", "Abstract", "v"],
        ["", "List of Figures", "viii"],
        ["", "List of Tables", "ix"],
        ["", "List of Abbreviations", "x"],
        ["1", "INTRODUCTION", "1"],
        ["1.1", "Background and Context", "1"],
        ["1.2", "Structural Defect Mechanisms in Civil Infrastructure", "2"],
        ["1.3", "Importance of Building Safety Inspection", "3"],
        ["1.4", "Problem Statement", "4"],
        ["1.5", "Motivation and Need for Automation", "4"],
        ["1.6", "Objectives of the Project", "5"],
        ["1.7", "Scope and Operational Boundaries", "6"],
        ["1.8", "Proposed Solution Overview", "7"],
        ["1.9", "Applications and Target Users", "8"],
        ["1.10", "Expected Contributions and Realistic Limitations", "9"],
        ["1.11", "Report Organisation", "10"],
        ["2", "LITERATURE SURVEY", "11"],
        ["2.1", "Conventional Visual Inspection and Non-Destructive Testing", "11"],
        ["2.2", "Traditional Image Processing in Defect Identification", "12"],
        ["2.3", "Deep Learning and CNNs for Structural Health Monitoring", "13"],
        ["2.4", "One-Stage vs. Two-Stage Object Detection Paradigms", "14"],
        ["2.5", "The YOLO Architecture Family and YOLOv8 Advances", "15"],
        ["2.6", "UAV-Based Automated Facade Inspection Frameworks", "17"],
        ["2.7", "Benchmark Datasets in Structural Inspection & MBDD2025", "18"],
        ["2.8", "Interactive Web Deployment Frameworks for Computer Vision", "19"],
        ["2.9", "Comparative Analysis of Existing Literature", "20"],
        ["2.10", "Identified Research Gaps and Project Motivation", "22"],
        ["3", "SYSTEM ANALYSIS", "23"],
        ["3.1", "Existing Inspection Workflows", "23"],
        ["3.2", "Critical Limitations of Existing Workflows", "24"],
        ["3.3", "Proposed System: BuildingSafe AI", "25"],
        ["3.4", "Key Advantages and Operational Boundaries", "26"],
        ["3.5", "Functional Requirements Specification", "27"],
        ["3.6", "Non-Functional Requirements Specification", "28"],
        ["3.7", "Hardware Requirements", "29"],
        ["3.8", "Software Requirements and Environment Setup", "30"],
        ["3.9", "Feasibility Analysis (Technical, Operational, Economic)", "31"],
        ["4", "SYSTEM DESIGN", "33"],
        ["4.1", "Overall System Architecture", "33"],
        ["4.2", "System Workflow and Data Flow Pipeline", "34"],
        ["4.3", "Use-Case Modeling and Actor Analysis", "35"],
        ["4.4", "Data Flow Diagrams (Level 0 and Level 1)", "36"],
        ["4.5", "Activity and Process Execution Models", "37"],
        ["4.6", "Image Inspection Sequence Workflow", "38"],
        ["4.7", "Camera Snapshot Architecture", "39"],
        ["4.8", "Video Frame Sampling and Analysis Architecture", "40"],
        ["4.9", "Session State and Inspection Log Design", "41"],
        ["4.10", "User Interface Ergonomics and Dark Theme Palette", "42"],
        ["5", "METHODOLOGY", "44"],
        ["5.1", "Dataset Overview: MBDD2025 Building Defects Dataset", "44"],
        ["5.2", "Target Defect Taxonomy and Characteristics", "45"],
        ["5.3", "Dataset Preprocessing and YOLO Coordinate Normalization", "46"],
        ["5.4", "Dataset Partitioning Strategy", "47"],
        ["5.5", "YOLOv8n Neural Network Architecture", "48"],
        ["5.6", "Transfer Learning and Pretrained Weights", "50"],
        ["5.7", "Training Configuration and Cloud GPU Execution", "51"],
        ["5.8", "Loss Functions Formulation (CIoU, DFL, BCE)", "52"],
        ["5.9", "Inference Pipeline, Confidence Thresholding, and NMS", "54"],
        ["5.10", "Quantitative Evaluation Metrics Formulation", "55"],
        ["6", "SYSTEM IMPLEMENTATION", "57"],
        ["6.1", "Development Environment and Modular Structure", "57"],
        ["6.2", "Model Loading and In-Memory Caching", "58"],
        ["6.3", "Image Upload and Preprocessing Module", "59"],
        ["6.4", "Camera Snapshot Ingestion Module", "60"],
        ["6.5", "Video Processing and Frame-Skip Ingestion Engine", "61"],
        ["6.6", "Defect Detection, Bounding-Box Extraction, and Plotting", "62"],
        ["6.7", "Per-Class Aggregation and Metric Gauges", "63"],
        ["6.8", "Tabular Findings Structuring and Data Serialization", "64"],
        ["6.9", "Session State Logging and CSV Export Subsystem", "65"],
        ["6.10", "Custom Dark Theme CSS and Responsive Layout", "66"],
        ["6.11", "Robust Error Handling and Boundary Constraints", "67"],
        ["7", "RESULTS AND DISCUSSION", "69"],
        ["7.1", "Quantitative Model Training Evaluation", "69"],
        ["7.2", "Convergence Characteristics and Training Duration", "70"],
        ["7.3", "Analysis of Application Landing Interface (Fig. 7.1)", "71"],
        ["7.4", "Analysis of Pre-Inspection Image Ingestion (Fig. 7.2)", "72"],
        ["7.5", "Analysis of Leakage Defect Detection (Fig. 7.3)", "73"],
        ["7.6", "Analysis of Metric Cards and Coordinate Table (Fig. 7.4)", "74"],
        ["7.7", "Analysis of Multi-Defect Co-Occurrence (Crack & Leakage)", "75"],
        ["7.8", "Analysis of Session Inspection Audit Log (Fig. 7.5)", "76"],
        ["7.9", "Video Inspection Experimental Evaluation", "77"],
        ["7.10", "System Latency, Throughput, and Resource Profiling", "78"],
        ["7.11", "Critical Evaluation of Strengths and Failure Modes", "79"],
        ["8", "CONCLUSION AND FUTURE ENHANCEMENT", "81"],
        ["8.1", "Summary of Completed Work", "81"],
        ["8.2", "Attainment of Project Objectives", "82"],
        ["8.3", "Core Technical Contributions", "83"],
        ["8.4", "Current Prototype Limitations", "84"],
        ["8.5", "Future Roadmap (Severity Scoring, WebRTC, PDF Reporting)", "85"],
        ["8.6", "Final Concluding Remarks", "86"],
        ["", "REFERENCES", "87"],
        ["", "APPENDIX A: USER AND OPERATOR GUIDE", "90"],
        ["", "APPENDIX B: CORE INFERENCE SOURCE CODE", "93"],
    ]
    body.append(tbl(toc_headers, toc_rows, col_widths=[1400, 6400, 1400]))
    body.append(page_break())

    # 7. LIST OF FIGURES
    body.append(h1("LIST OF FIGURES"))
    lof_headers = ["Figure No.", "Figure Title", "Page No."]
    lof_rows = [
        ["Fig. 4.1", "System architecture of BuildingSafe AI (Phase 1 Training & Phase 2 Inference)", "33"],
        ["Fig. 4.2", "Image inspection logical sequence workflow", "38"],
        ["Fig. 4.3", "Uploaded-video frame-sampling inspection workflow", "40"],
        ["Fig. 5.1", "Dataset preparation, annotation normalization, and YOLOv8n training pipeline", "47"],
        ["Fig. 5.2", "YOLOv8 architectural block diagram (Backbone, C2f Neck, and Anchor-Free Head)", "49"],
        ["Fig. 7.1", "BuildingSafe AI home screen showing dark UI, model status, and inspection tabs", "71"],
        ["Fig. 7.2", "Image inspection screen after image upload prior to inference execution", "72"],
        ["Fig. 7.3", "Annotated image comparison exhibiting detected leakage with bounding box", "73"],
        ["Fig. 7.4", "Detailed detection findings, confidence meters, and pixel coordinates table", "74"],
        ["Fig. 7.5", "Session-level inspection log showing recorded audit trail and CSV download", "76"],
    ]
    body.append(tbl(lof_headers, lof_rows, col_widths=[1600, 6200, 1400]))
    body.append(page_break())

    # 8. LIST OF TABLES
    body.append(h1("LIST OF TABLES"))
    lot_headers = ["Table No.", "Table Title", "Page No."]
    lot_rows = [
        ["Table 2.1", "Comparative summary of existing literature in infrastructure defect detection", "20"],
        ["Table 3.1", "Functional Requirements Specification (FR-01 to FR-08)", "27"],
        ["Table 3.2", "Non-Functional Requirements Specification", "28"],
        ["Table 3.3", "Hardware and Software Environment Specifications", "30"],
        ["Table 4.1", "Architectural Layer Decomposition and Responsibilities", "34"],
        ["Table 5.1", "MBDD2025 Dataset Characteristics and Split Partitioning", "44"],
        ["Table 5.2", "Target Defect Taxonomy and Structural Degradation Manifestations", "45"],
        ["Table 5.3", "YOLOv8n Training Hyperparameters and Runtime Configurations", "51"],
        ["Table 5.4", "Quantitative Model Evaluation Benchmark Metrics", "56"],
        ["Table 6.1", "Core Source Code Modules and Implementation Responsibilities", "57"],
        ["Table 7.1", "YOLOv8n Validation Performance Summary on MBDD2025 Dataset", "69"],
        ["Table 7.2", "Representative Single-Image Inspection Output Metrics (Hefei8702.jpg)", "74"],
    ]
    body.append(tbl(lot_headers, lot_rows, col_widths=[1600, 6200, 1400]))
    body.append(page_break())

    # 9. LIST OF ABBREVIATIONS
    body.append(h1("LIST OF ABBREVIATIONS"))
    abbr_headers = ["Abbreviation", "Expanded Form / Meaning"]
    abbr_rows = [
        ["AI", "Artificial Intelligence"],
        ["ANN", "Artificial Neural Network"],
        ["AP", "Average Precision"],
        ["BCE", "Binary Cross-Entropy"],
        ["CIoU", "Complete Intersection over Union"],
        ["CNN", "Convolutional Neural Network"],
        ["CPU", "Central Processing Unit"],
        ["CSP", "Cross Stage Partial"],
        ["CSV", "Comma-Separated Values"],
        ["CV", "Computer Vision"],
        ["DFD", "Data Flow Diagram"],
        ["DFL", "Distribution Focal Loss"],
        ["FN", "False Negative"],
        ["FP", "False Positive"],
        ["FPS", "Frames Per Second"],
        ["GPU", "Graphics Processing Unit"],
        ["IoU", "Intersection over Union"],
        ["mAP", "Mean Average Precision"],
        ["MBDD", "Multi-type Building Defect Dataset"],
        ["NDT", "Non-Destructive Testing"],
        ["NMS", "Non-Maximum Suppression"],
        ["PANet", "Path Aggregation Network"],
        ["PASCAL VOC", "Pattern Analysis, Statistical Modelling and Computational Learning Visual Object Classes"],
        ["PIL", "Python Imaging Library (Pillow)"],
        ["RAM", "Random Access Memory"],
        ["RGB", "Red, Green, Blue"],
        ["SDG", "Sustainable Development Goals"],
        ["SHM", "Structural Health Monitoring"],
        ["SPPF", "Spatial Pyramid Pooling - Fast"],
        ["TP", "True Positive"],
        ["UAV", "Unmanned Aerial Vehicle (Drone)"],
        ["UI", "User Interface"],
        ["VOC", "Visual Object Classes"],
        ["XML", "Extensible Markup Language"],
        ["YOLO", "You Only Look Once"],
    ]
    body.append(tbl(abbr_headers, abbr_rows, col_widths=[2400, 6800]))
    body.append(page_break())

    # =========================================================================
    # CHAPTER 1: INTRODUCTION
    # =========================================================================
    body.append(h1("CHAPTER 1\nINTRODUCTION"))
    
    body.append(h2("1.1 Background and Context"))
    c1_1 = (
        "Civil infrastructure represents the physical backbone of modern society, encompassing residential complexes, "
        "commercial buildings, educational institutions, industrial facilities, and public infrastructure. The structural integrity, "
        "operational safety, and longevity of these buildings are continually challenged by complex environmental stressors, material "
        "ageing, dynamic loading, thermal fluctuations, cyclic moisture penetration, chemical exposure, and construction defects. "
        "Over time, surface manifestations of deterioration emerge on walls, facades, structural columns, beams, and ceilings. "
        "These manifestations include superficial and structural cracking, dampness, moisture seepage, reinforcing steel corrosion, "
        "concrete spalling (abscission), and surface bulging.\n\n"
        "Historically, structural health monitoring (SHM) and periodic maintenance surveys have relied on manual visual inspections "
        "conducted by human inspectors, structural engineers, or maintenance technicians. While skilled human expertise remains essential "
        "for contextual engineering diagnosis, conventional manual inspections suffer from severe systemic limitations. They are "
        "labor-intensive, financially expensive, and physically hazardous—especially when examining high-rise facades, bridge decks, "
        "overhanging eaves, and hard-to-reach elevated surfaces where scaffolding, cherry pickers, or rope-access teams are required. "
        "Furthermore, human evaluations are inherently subjective, prone to fatigue-induced omissions, and historically lack centralized, "
        "standardized digital audit trails. As urbanization intensifies and the global stock of reinforced concrete, masonry, and steel "
        "structures ages, there is an urgent demand for automated, objective, non-destructive, and scalable visual inspection assistance "
        "tools powered by modern Artificial Intelligence (AI) and Computer Vision (CV)."
    )
    body.append(p(c1_1))

    body.append(h2("1.2 Structural Defect Mechanisms in Civil Infrastructure"))
    c1_2 = (
        "To establish an effective automated detection paradigm, the underlying civil engineering mechanisms governing visible surface "
        "defects must be rigorously understood. The five defect classes targeted in this project represent the most prevalent and critical "
        "early-warning symptoms of structural deterioration:\n\n"
        "1. Cracks (Structural and Non-Structural): Cracks manifest as linear fractures or discontinuities on concrete, mortar, or plaster "
        "surfaces. They originate from plastic shrinkage during curing, thermal expansion and contraction, uneven structural foundation "
        "settlement, structural overloading, or shear failure. While micro-cracks (<0.2 mm) may initially represent cosmetic distress, "
        "unattended cracks provide primary ingress channels for moisture, carbon dioxide, and atmospheric chloride ions, drastically "
        "accelerating internal steel rebar corrosion.\n\n"
        "2. Water Leakage and Moisture Seepage: Leakage appears as discolored, damp, or efflorescence-stained patches on exterior and "
        "interior surfaces. It results from failing waterproofing membranes, leaking plumbing conduits, poor joint sealant integrity, or "
        "capillary moisture rise. Chronic moisture intrusion degrades plaster cohesion, promotes mold growth, and provides the necessary "
        "electrolyte for electrochemical corrosion of embedded steel reinforcement.\n\n"
        "3. Surface Corrosion: Corrosion manifests as reddish-brown ferric oxide stains bleeding through concrete, masonry, or exposed "
        "steel elements. In reinforced concrete, corrosion occurs when carbonation or chloride ingress depassivates the protective alkaline "
        "film around embedded rebar. As iron oxidizes into rust, it expands up to six times its original volume, generating immense internal "
        "tensile bursting stresses within the surrounding concrete cover.\n\n"
        "4. Concrete Abscission (Spalling / Flaking): Abscission denotes the detachment, delamination, or breaking away of surface concrete, "
        "rendering the internal aggregate or reinforcement steel directly exposed to ambient atmospheric conditions. It represents the "
        "direct consequence of advancing corrosion-induced tensile pressure or severe freeze-thaw cycles. Abscission drastically reduces "
        "effective load-bearing cross-sections and poses acute falling-debris hazards to pedestrians below.\n\n"
        "5. Surface Bulging (Out-of-Plane Buckling): Bulges appear as localized convex surface swellings or out-of-plane planar deformations. "
        "They occur due to delamination between the structural concrete substrate and finishing plaster coats, sub-surface moisture vapor "
        "pressure accumulation, or localized compressive buckling under eccentric structural loading. Bulges frequently precede sudden, "
        "catastrophic detachment of masonry cladding or plaster blocks."
    )
    body.append(p(c1_2))

    body.append(h2("1.3 Importance of Building Safety Inspection"))
    c1_3 = (
        "Regular, systematic building safety inspections serve three pivotal socio-economic objectives:\n\n"
        "• Preservation of Human Life and Worker Safety: Early detection of spalling concrete, loose cladding, or advancing structural "
        "fractures prevents life-threatening facade collapse and falling-debris accidents in dense urban environments.\n"
        "• Economic Optimization and Lifecycle Cost Reduction: The economic cost of structural remediation follows a geometric escalation "
        "curve. Remedying early micro-cracking and minor sealant failure through localized sealants costs a fraction of extensive structural "
        "re-jacketing, rebar replacement, or emergency demolition necessitated by decades of unchecked moisture-induced corrosion.\n"
        "• Asset Value and Regulatory Compliance: Modern municipal building codes and insurance regulations increasingly mandate documented, "
        "periodic structural condition assessments. Automated digital records establish an immutable audit trail verifying building health."
    )
    body.append(p(c1_3))

    body.append(h2("1.4 Problem Statement"))
    c1_4 = (
        "Traditional building defect inspection remains overwhelmingly dependent on manual visual examination, which is time-consuming, "
        "labor-intensive, subjective, costly, and hazardous when applied to elevated facades and large-scale infrastructure. Minor surface "
        "anomalies frequently go unrecorded until advanced structural deterioration has occurred. Furthermore, existing research models "
        "frequently remain isolated within offline experimental notebooks, lacking accessible, end-to-end user interfaces suitable for field "
        "practitioners.\n\n"
        "Therefore, there exists a well-defined academic and engineering imperative to develop an AI-based visual screening assistance "
        "prototype that can ingest multi-source visual media (photographs, device camera snapshots, and UAV-recorded videos), automatically "
        "detect and localize the five major building surface defect classes using state-of-the-art deep object detection, present traceably "
        "quantified bounding coordinates and confidence scores, and maintain a session-level inspection log within an intuitive dashboard, "
        "without misleadingly claiming automated structural safety certification."
    )
    body.append(p(c1_4))

    body.append(h2("1.5 Motivation and Need for Automation"))
    c1_5 = (
        "The motivation for this investigation is rooted in the convergence of recent breakthroughs in deep learning-based computer vision "
        "and the widespread commercial availability of high-resolution digital imaging sensors and Unmanned Aerial Vehicles (UAVs). "
        "While drones can easily photograph previously inaccessible building envelopes, human analysts are quickly overwhelmed by the "
        "thousands of high-resolution photographs generated during a single flight survey. Automating the visual screening bottleneck using "
        "deep neural networks allows computers to handle repetitive scanning, directing human engineers straight to flagged anomalies.\n\n"
        "Furthermore, modern one-stage object detectors such as Ultralytics YOLOv8 have achieved exceptional tradeoffs between inference speed "
        "and detection precision. Integrating such high-performance detection backbones into rapid, browser-based deployment frameworks "
        "such as Streamlit empowers maintenance teams to conduct rapid preliminary screening on site without requiring specialized deep "
        "learning hardware or complex command-line programming skills."
    )
    body.append(p(c1_5))

    body.append(h2("1.6 Objectives of the Project"))
    c1_6 = (
        "The primary objectives of this final-year engineering project are formulated as follows:\n\n"
        "1. Dataset Ingestion and Preprocessing: Preprocess and normalize the large-scale MBDD2025 Building Defects Dataset containing "
        "14,471 UAV images and convert annotation boundaries into standardized YOLO bounding-box coordinate formats.\n"
        "2. Partitioning Strategy: Establish a clean 85:15 stratified split, dividing the corpus into 12,301 training images and 2,170 "
        "validation images to guarantee rigorous, un-leaked evaluation.\n"
        "3. Model Architecture Training: Fine-tune a lightweight YOLOv8n (nano) object detection network using transfer learning from "
        "pretrained weights, utilizing an input resolution of 640 × 640 pixels across 20 epochs on a cloud-based NVIDIA T4 GPU.\n"
        "4. Multi-Defect Localization: Detect and localize five target defect categories simultaneously: crack, leakage, corrosion, "
        "abscission, and bulge, achieving competitive evaluation metrics (Precision ≥ 80%, Recall ≥ 70%, mAP@50 ≥ 80%).\n"
        "5. Interactive Web Interface Engineering: Develop an intuitive, dark-themed Streamlit dashboard (BuildingSafe AI) supporting "
        "image file upload, single camera-snapshot acquisition, and frame-by-frame uploaded-video scanning.\n"
        "6. Quantitative Reporting and Artifact Export: Extract exact pixel bounding coordinates (x1, y1, x2, y2), calculate class-level "
        "confidence metrics, maintain an in-memory chronological session log, and provide downloadable annotated imagery and CSV records.\n"
        "7. Architectural Foundation: Establish an extensible software architecture prepared for future enhancements including structural "
        "severity indexing, real-time WebRTC camera feeds, and automated PDF engineering audit report generation."
    )
    body.append(p(c1_6))

    body.append(h2("1.7 Scope and Operational Boundaries"))
    c1_7 = (
        "To preserve scientific integrity and academic honesty, the operational boundaries and explicit scope of BuildingSafe AI are "
        "defined as follows:\n\n"
        "• Visual Detection Scope: The system is designed exclusively to detect surface-visible manifestations corresponding to the five "
        "annotated classes in the MBDD2025 dataset. Internal subsurface voids, structural load capacity, crack depth, rebar yield stress, "
        "and foundation settlement cannot be measured from two-dimensional RGB photographs.\n"
        "• Input Media Modalities: The application accepts standard RGB images (JPG, JPEG, PNG), single snapshot captures via browser "
        "camera input, and recorded video files (MP4, MOV, AVI). The camera input is strictly a single photographic capture, not a "
        "continuous live-streaming WebRTC feed.\n"
        "• Filtering vs. Severity: The user-adjustable confidence threshold slider is a mathematical probability filter denoting the "
        "neural network's statistical certainty; it must not be conflated with the physical structural severity or hazard level of a defect.\n"
        "• Non-Certification Clause: The prototype functions strictly as an assistive visual-screening tool. It does not replace certified "
        "structural engineers or issue formal building compliance certificates."
    )
    body.append(p(c1_7))

    body.append(h2("1.8 Proposed Solution Overview"))
    c1_8 = (
        "The proposed BuildingSafe AI system unites a deep learning object detection pipeline with an accessible, interactive Streamlit "
        "web environment. The end-to-end framework comprises two operational phases:\n\n"
        "Phase 1: Model Training and Validation Pipeline (Executed on Cloud GPU): The 14,471 images of the MBDD2025 dataset and their "
        "corresponding VOC XML annotations are normalized into YOLO text files containing normalized bounding box coordinates. The dataset "
        "is split into 12,301 training images and 2,170 validation images. A YOLOv8n network is fine-tuned at 640 × 640 resolution using "
        "Complete IoU (CIoU) loss and Distribution Focal Loss (DFL). The optimal weights yielding peak validation mAP@50 are saved as best.pt.\n\n"
        "Phase 2: User-Facing Inference Application (Executed locally or in cloud runtime): The Streamlit application dynamically loads "
        "best.pt into memory with resource caching. Upon media ingestion, the image is passed to the YOLOv8 inference engine. Predicted "
        "bounding boxes are scaled to native image dimensions, plotted with distinctive class colors, and summarized in metric tiles and "
        "tabular formats. For video files, a dynamic frame-skipping algorithm decodes every Nth frame, accumulating defect frequencies "
        "and capturing the highest-confidence frame. All completed inspections are recorded in an interactive session audit log."
    )
    body.append(p(c1_8))

    body.append(h2("1.9 Applications and Target Users"))
    c1_9 = (
        "The BuildingSafe AI system offers tangible operational utility across several civil engineering and facility management domains:\n\n"
        "• Municipal Building Safety Audits: Facilitates preliminary visual surveys of commercial facades, public housing estates, and "
        "heritage masonry structures to flag candidate areas requiring detailed physical scaffolding inspection.\n"
        "• Post-Disaster Rapid Screening: Following earthquakes, cyclones, or heavy flooding, emergency response teams can deploy UAVs to "
        "rapidly scan structural envelopes and catalog visible exterior fractures or spalling hazards.\n"
        "• Construction Quality Assurance: Enables general contractors to verify surface finish quality, identify early plaster cracking or "
        "waterproofing leakage before building handover, and archive objective photographic proof.\n"
        "• Educational and Research Benchmark: Serves as an applied computer vision laboratory framework for engineering students studying "
        "transfer learning, object detection localization, and modern web application integration."
    )
    body.append(p(c1_9))

    body.append(h2("1.10 Expected Contributions and Limitations"))
    c1_10 = (
        "The primary contribution of this work lies in the successful end-to-end implementation and validation of an applied AI inspection "
        "tool connecting deep convolutional feature learning with a practical, ergonomically sound engineering dashboard. Rather than "
        "outputting raw numeric arrays, the system provides immediate visual context, per-class confidence gauges, pixel bounding coordinates, "
        "and exportable tabular data.\n\n"
        "The recognized limitations include sensitivity to extreme illumination changes (deep shadows and direct lens glare), occasional "
        "false positives on heavily textured tree bark or wall graffiti, frame-skipping omissions during video analysis, and the session-bound "
        "nature of the inspection log, which clears when the application server resets."
    )
    body.append(p(c1_10))

    body.append(h2("1.11 Report Organisation"))
    c1_11 = (
        "This project report is structured into eight comprehensive chapters following formal academic engineering conventions:\n\n"
        "• Chapter 1 establishes the problem background, civil engineering defect mechanisms, objectives, scope, and proposed solution.\n"
        "• Chapter 2 surveys foundational and recent literature in visual defect detection, object detection architectures, and benchmark datasets.\n"
        "• Chapter 3 presents system analysis, contrasting existing manual workflows against the proposed system and detailing requirements.\n"
        "• Chapter 4 details system design, including system architecture, sequence workflows, use-case models, and user interface ergonomics.\n"
        "• Chapter 5 delineates the core methodology, detailing MBDD2025 dataset preparation, YOLOv8n network mathematics, and loss formulations.\n"
        "• Chapter 6 describes implementation specifics, modular Python scripts, Streamlit UI components, and boundary handling.\n"
        "• Chapter 7 presents quantitative training results, analyzes uploaded test screenshots, and evaluates system latency and edge cases.\n"
        "• Chapter 8 concludes the report, summarizing achievements, limitations, and future research enhancements."
    )
    body.append(p(c1_11))
    body.append(page_break())

    # =========================================================================
    # CHAPTER 2: LITERATURE SURVEY
    # =========================================================================
    body.append(h1("CHAPTER 2\nLITERATURE SURVEY"))
    
    body.append(h2("2.1 Conventional Visual Inspection and Non-Destructive Testing"))
    c2_1 = (
        "Civil infrastructure maintenance has traditionally relied upon periodic manual visual inspection complemented by non-destructive "
        "testing (NDT) techniques such as rebound hammer tests, ultrasonic pulse velocity, ground-penetrating radar, and infrared thermography. "
        "Visual inspection serves as the indispensable first line of defense; however, extensive studies by the Federal Highway Administration "
        "(FHWA) have documented significant variability in manual inspection results, driven by inspector fatigue, varying visual acuity, "
        "ambient weather conditions, and subjective judgment. While advanced NDT methods provide subsurface insight, they are localized, "
        "equipment-heavy, and cost-prohibitive for large-scale building envelopes. Consequently, high-resolution optical image analysis has "
        "emerged as the primary modality for rapid, comprehensive structural screening."
    )
    body.append(p(c2_1))

    body.append(h2("2.2 Traditional Image Processing in Defect Identification"))
    c2_2 = (
        "Early automated defect detection frameworks developed during the late 1990s and 2000s relied on classical computer vision algorithms "
        "incorporating hand-crafted visual features. Researchers widely utilized Sobel, Canny, and Prewitt edge detectors, morphological "
        "dilation and erosion, Otsu thresholding, Gabor wavelets, and Hough transforms to isolate crack patterns from concrete backgrounds. "
        "While computationally lightweight, these deterministic approaches exhibited severe fragility when exposed to real-world outdoor "
        "environments. Natural concrete textures, surface mottling, uneven sunlight illumination, heavy shadow cast by trees or scaffolding, "
        "and surface graffiti caused extensive false positive edges, severely restricting classical methods to highly controlled indoor "
        "laboratory settings."
    )
    body.append(p(c2_2))

    body.append(h2("2.3 Deep Learning and CNNs for Structural Health Monitoring"))
    c2_3 = (
        "The advent of deep Convolutional Neural Networks (CNNs) revolutionized structural health monitoring by replacing hand-crafted "
        "feature engineering with automated hierarchical representation learning. In a seminal study, Cha, Choi, and Büyüköztürk (2017) "
        "demonstrated that deep CNNs could reliably detect concrete cracks under variable illumination conditions without requiring manual "
        "image filtering. Similarly, Zhang et al. (2016) developed CrackNet, a convolutional architecture for pavement crack identification. "
        "Subsequent research extended CNN classifiers to recognize efflorescence, spalling, and moisture staining. However, traditional CNN "
        "classifiers operated under patch-based classification paradigms: images had to be sliced into hundreds of overlapping patches, "
        "classified individually, and reconstructed. This sliding-window approach was computationally redundant, highly inefficient, and "
        "incapable of predicting precise bounding-box coordinates for irregularly shaped defects."
    )
    body.append(p(c2_3))

    body.append(h2("2.4 One-Stage vs. Two-Stage Object Detection Paradigms"))
    c2_4 = (
        "To achieve simultaneous defect classification and spatial localization, researchers transitioned to modern deep object detection "
        "architectures, broadly categorized into two paradigms:\n\n"
        "• Two-Stage Detectors (e.g., Faster R-CNN): In the first stage, a Region Proposal Network (RPN) scans feature maps to generate "
        "candidate regions of interest (RoIs). In the second stage, RoI pooling extracts features for bounding-box regression and softmax "
        "classification. While Faster R-CNN provides high localization precision, the two-stage execution creates an inherent computational "
        "bottleneck, resulting in frame rates below 10 FPS on standard GPUs. This high latency renders two-stage detectors impractical for "
        "real-time UAV video feeds and edge computing.\n\n"
        "• One-Stage Detectors (e.g., SSD, YOLO): One-stage detectors frame object localization as a unified regression problem, predicting "
        "bounding-box coordinates and class probabilities directly from full feature maps in a single feed-forward pass. The You Only Look "
        "Once (YOLO) series introduced by Redmon et al. (2016) demonstrated real-time inference speeds exceeding 45 FPS while maintaining "
        "competitive mean Average Precision. Subsequent iterations (YOLOv3, YOLOv5, and YOLOv8) progressively introduced feature pyramid "
        "networks, mosaic data augmentation, and anchor-free detection heads, making one-stage detectors the benchmark choice for civil "
        "inspection systems."
    )
    body.append(p(c2_4))

    body.append(h2("2.5 The YOLO Architecture Family and YOLOv8 Advances"))
    c2_5 = (
        "The YOLO architecture has undergone remarkable evolutionary refinement. Early YOLO models utilized predefined anchor boxes "
        "clustering anticipated defect aspect ratios. However, building defects such as tortuous cracks and amorphous water leakage stains "
        "exhibit extreme aspect-ratio variance, making fixed anchor assignment sub-optimal.\n\n"
        "Ultralytics YOLOv8 (released in 2023) introduced three groundbreaking architectural innovations: (1) An Anchor-Free Detection Head, "
        "which directly predicts the object center and distances to the bounding boundaries, drastically reducing hyperparameter tuning "
        "and accelerating Non-Maximum Suppression (NMS); (2) The C2f (Cross Stage Partial with two convolutions) module, which replaces the "
        "earlier C3 module, enhancing gradient flow and multi-scale feature representation while reducing parameter count; and (3) Decoupled "
        "Heads, separating classification loss (Binary Cross-Entropy) from regression loss (Complete IoU and Distribution Focal Loss). The "
        "nano variant (YOLOv8n) provides an extraordinarily compact footprint (~3.2 million parameters), enabling rapid inference on CPU "
        "runtimes while retaining superior feature extraction capacity."
    )
    body.append(p(c2_5))

    body.append(h2("2.6 UAV-Based Automated Facade Inspection Frameworks"))
    c2_6 = (
        "The integration of Unmanned Aerial Vehicles (UAVs) with deep learning has established a transformative frontier in building inspection. "
        "Equipped with stabilized gimbal cameras, UAVs capture close-range, orthogonal imagery of high-rise envelopes without exposing human "
        "workers to fall hazards. Researchers such as Dung and Anh (2019) demonstrated autonomous UAV concrete crack segmentation. However, "
        "practical UAV deployment introduces unique computer vision challenges: distance-dependent scale variations, motion blur from wind "
        "buffeting, complex perspective distortion, and complex background interference (foliage, windows, reflective cladding). Consequently, "
        "robust multi-scale feature aggregation is vital for UAV inspection models."
    )
    body.append(p(c2_6))

    body.append(h2("2.7 Benchmark Datasets in Structural Inspection & MBDD2025"))
    c2_7 = (
        "The performance of deep learning systems is fundamentally bounded by the quality and diversity of training data. Early benchmark "
        "datasets, such as the SDNET2018 and CCIC datasets, were limited to binary crack classification or small, highly uniform concrete "
        "patches. To address these limitations, Zha, Yao, Zhang, and Ma (2025) released the MBDD2025 dataset—a comprehensive, large-scale "
        "benchmark curated specifically for diverse building defect detection from UAV imagery. Comprising 14,471 annotated images spanning "
        "six structural categories (steel, masonry, timber, brick-wood, brick-concrete, and reinforced concrete), MBDD2025 provides authentic "
        "real-world visual complexity and standardized annotations across five critical defect classes: crack, leakage, corrosion, "
        "abscission, and bulge."
    )
    body.append(p(c2_7))

    body.append(h2("2.8 Interactive Web Deployment Frameworks for Computer Vision"))
    c2_8 = (
        "While deep learning algorithms are frequently developed using Python libraries such as PyTorch or TensorFlow, deploying models for "
        "civil engineering practitioners requires accessible user interfaces. Traditional enterprise web development using Django, Flask, or "
        "FastAPI paired with React or Angular entails extensive frontend-backend state management, slowing prototyping. Streamlit has "
        "emerged as the preeminent framework for machine learning deployment, allowing Python-native scripting of interactive dashboards, "
        "file uploaders, camera inputs, and reactive visual components. Streamlit's reactive execution model and in-memory caching mechanisms "
        "provide an optimal environment for deploying computer vision prototypes such as BuildingSafe AI."
    )
    body.append(p(c2_8))

    body.append(h2("2.9 Comparative Analysis of Existing Literature"))
    c2_9_intro = (
        "Table 2.1 presents a comprehensive comparative synthesis of ten pivotal research works spanning civil infrastructure defect "
        "detection, object detection paradigms, and automated inspection systems."
    )
    body.append(p(c2_9_intro))
    
    lit_headers = ["Ref / Author", "Year", "Methodology / Model", "Target Defects", "Key Inferences", "Identified Limitations"]
    lit_rows = [
        ["Cha et al.", "2017", "Deep CNN (Patch-based)", "Concrete Cracks", "Demonstrated CNN robustness against varying illumination.", "Patch-based sliding window; no bounding box localization."],
        ["Ren et al.", "2017", "Faster R-CNN (Two-stage)", "General Objects", "Region Proposal Network unified detection pipeline.", "High computational latency (~150 ms GPU); unsuitable for edge/video."],
        ["Dung & Anh", "2019", "Fully Convolutional Net", "Concrete Cracks", "Pixel-level semantic segmentation of concrete fractures.", "High computational cost; single defect class (cracks only)."],
        ["Kim et al.", "2020", "UAV + Mask R-CNN", "Cracks & Spalling", "Combined drone flight path with instance segmentation.", "Slow inference speed; complex multi-stage deployment overhead."],
        ["Munawar et al.", "2021", "YOLOv4 & ResNet", "Cracks, Spalling, Efflorescence", "Real-time multi-defect detection on infrastructure.", "Suffered high false positives on heavily textured weathered concrete."],
        ["Redmon et al.", "2016", "YOLO (Single-stage)", "General Objects", "Framed detection as direct regression; real-time 45 FPS.", "Lower accuracy on small, densely clustered objects."],
        ["Jocher et al.", "2023", "YOLOv8 (Anchor-free)", "General Objects", "Anchor-free decoupled head; C2f feature enhancement.", "General weights require domain-specific transfer learning."],
        ["Zha et al.", "2025", "MBDD2025 Benchmark", "Crack, Leak, Corrode, Abscise, Bulge", "Curated 14,471 multi-structure UAV building defect images.", "Provided raw dataset; no integrated interactive deployment tool."],
        ["Zhang et al.", "2016", "CrackNet (CNN)", "Pavement Cracks", "Pixel-level crack classification preserving geometry.", "Restricted to flat pavement surfaces; poor facade generalization."],
        ["Li et al.", "2024", "Attention YOLOv7", "Facade Spalling", "Embedded spatial attention modules for small defect recall.", "High model complexity; lack of accessible web user interface."]
    ]
    body.append(tbl(lit_headers, lit_rows, col_widths=[1400, 800, 1800, 1600, 2000, 2000]))

    body.append(h2("2.10 Identified Research Gaps and Project Motivation"))
    c2_10 = (
        "A rigorous synthesis of the reviewed literature reveals three critical research gaps that motivate the design of BuildingSafe AI:\n\n"
        "1. Isolated Single-Class Focus: The vast majority of published studies focus exclusively on crack detection, overlooking co-occurring "
        "defects such as water leakage, rebar corrosion, surface bulges, and concrete spalling (abscission) that frequently accompany structural "
        "distress in authentic building envelopes.\n"
        "2. Absence of Accessible End-to-End Tools: Research models typically terminate at offline performance evaluation curves within Jupyter "
        "Notebooks. Civil engineers, site inspectors, and building managers lack accessible, interactive applications that allow intuitive "
        "media drag-and-drop, dynamic sensitivity adjustments, and instant tabular data export.\n"
        "3. Lack of Multi-Input Ingestion Flexibility: Existing visual tools typically accept only pre-cropped single images, lacking unified "
        "pipelines capable of handling uploaded still photos, immediate camera snapshots, and temporal video frame analysis.\n\n"
        "BuildingSafe AI directly bridges these gaps by fine-tuning an advanced anchor-free YOLOv8n detector across all five classes of the "
        "MBDD2025 benchmark and embedding it within a responsive, dark-themed Streamlit dashboard equipped with comprehensive logging and "
        "export capabilities."
    )
    body.append(p(c2_10))
    body.append(page_break())

    # =========================================================================
    # CHAPTER 3: SYSTEM ANALYSIS
    # =========================================================================
    body.append(h1("CHAPTER 3\nSYSTEM ANALYSIS"))
    
    body.append(h2("3.1 Existing Inspection Workflows"))
    c3_1 = (
        "In contemporary civil engineering practice, routine building safety audits and facade condition assessments follow a heavily "
        "manual, physical workflow. Certified structural inspectors or maintenance technicians physically traverse the site, conducting "
        "visual examinations from ground level using binoculars or ascending via temporary scaffolding, cherry pickers, or rope-access "
        "rigs. When a suspicious surface irregularity is observed, the inspector manually photographs the anomaly using a handheld digital "
        "camera, records written observations on paper clipboards or mobile note-taking apps, and attempts physical measurements using crack "
        "width rulers or tape measures. Following the on-site survey, the inspector manually transfers hundreds of disorganized digital photos "
        "to a desktop computer, correlates filenames with handwritten notes, and drafts a narrative condition assessment report in a word processor."
    )
    body.append(p(c3_1))

    body.append(h2("3.2 Critical Limitations of Existing Workflows"))
    c3_2 = (
        "The conventional inspection paradigm suffers from several critical operational bottlenecks:\n\n"
        "• Severe Human Safety Hazards: Ascending high-rise facades, bridge decks, and suspended platforms exposes inspectors to severe fall "
        "hazards, representing the leading cause of construction-related fatalities globally.\n"
        "• Extreme Time and Labor Costs: Manually surveying a large commercial facade or university campus requires days or weeks of physical "
        "labor, causing substantial operational downtime and high financial expense.\n"
        "• Subjectivity and Low Inter-Rater Reliability: Assessments rely entirely on the visual acuity, fatigue level, and subjective "
        "experience of individual inspectors. Identical cracks are frequently classified with differing severity by different observers.\n"
        "• Absence of Systematic Digital Audit Trails: Handwritten field notes and unindexed photo directories are prone to data loss, "
        "hindering longitudinal comparisons to verify whether a crack has widened over successive years.\n"
        "• Inability to Process UAV Video Streams: While drones can rapidly record 4K inspection videos, site engineers lack automated tools "
        "to scan video footage, forcing them to spend hours manually scrubbing through video timelines to spot fleeting defects."
    )
    body.append(p(c3_2))

    body.append(h2("3.3 Proposed System: BuildingSafe AI"))
    c3_3 = (
        "BuildingSafe AI is engineered as an automated, AI-driven visual inspection assistance system that modernizes the preliminary "
        "screening workflow. Leveraging deep transfer learning via an anchor-free YOLOv8n detector and an interactive Streamlit web interface, "
        "the proposed system automates defect detection, localization, and documentation.\n\n"
        "The user simply uploads building photographs, captures camera snapshots, or submits inspection videos. The system automatically "
        "identifies cracks, leakage, corrosion, abscission, and bulges, plots color-coded bounding boxes, extracts normalized pixel "
        "coordinates, calculates class-wise confidence metrics, and maintains an in-memory session log. This eliminates manual image "
        "sorting and ensures standardized, traceable preliminary screening."
    )
    body.append(p(c3_3))

    body.append(h2("3.4 Key Advantages and Operational Boundaries"))
    c3_4 = (
        "The proposed system delivers substantial operational advantages while acknowledging strict operational boundaries:\n\n"
        "Advantages:\n"
        "• Multi-Modal Ingestion: Unified interface supporting image files (JPG/PNG), device camera snapshots, and recorded video files (MP4/AVI).\n"
        "• High Detection Speed: Delivers rapid inference (~238 ms to 409 ms per frame), enabling immediate user feedback.\n"
        "• Interactive Sensitivity Control: Dynamic slider allows inspectors to adjust detection confidence thresholds from 5% to 95%.\n"
        "• Traceable Coordinate Extraction: Bounding coordinates (x1, y1, x2, y2) provide precise spatial localization for digital documentation.\n"
        "• Downloadable Engineering Artifacts: Supports instant export of annotated PNG images and structured CSV findings tables.\n"
        "• Ergonomic Industrial UI: Custom dark-themed layout tailored for prolonged engineering review without visual fatigue.\n\n"
        "Operational Boundaries:\n"
        "• Visual Surface Screening Only: The system identifies visible optical patterns; it cannot assess subsurface voids, load-bearing "
        "capacity, internal rebar tension, or foundational stability.\n"
        "• Single-Snapshot Camera: Camera inspection processes individual captured frames rather than a continuous live WebRTC video stream.\n"
        "• Non-Certification Clause: System outputs serve as assistive screening evidence and do not replace certified structural engineering audits."
    )
    body.append(p(c3_4))

    body.append(h2("3.5 Functional Requirements Specification"))
    c3_5_intro = "Table 3.1 details the functional requirements (FR) governing system behavior and operational capabilities."
    body.append(p(c3_5_intro))
    
    fr_headers = ["Req ID", "Functional Requirement Description", "Input / Trigger", "Expected System Output"]
    fr_rows = [
        ["FR-01", "Image File Ingestion", "Upload JPG, JPEG, or PNG file via drag-and-drop.", "Renders input image preview; initializes detection pipeline."],
        ["FR-02", "Camera Snapshot Capture", "Trigger browser camera snapshot capture.", "Captures single photo; passes RGB buffer to inference engine."],
        ["FR-03", "Video Ingestion & Frame Sampling", "Upload MP4, MOV, or AVI inspection video.", "Decodes video; samples every Nth frame via OpenCV."],
        ["FR-04", "Model Weights Loading & Caching", "Application startup or sidebar model path update.", "Loads best.pt into memory using @st.cache_resource."],
        ["FR-05", "Multi-Defect YOLOv8 Inference", "Click 'Inspect building' button.", "Predicts bounding boxes, defect classes, and confidences."],
        ["FR-06", "Coordinate & Metric Visualization", "Completion of inference pipeline.", "Displays annotated image, metric tiles, and coordinate table."],
        ["FR-07", "Findings Data Export", "Click 'Download findings (CSV)' or 'Download annotated image'.", "Generates structured CSV file and PNG image download."],
        ["FR-08", "Session-Level Audit Logging", "Completion of any inspection task.", "Appends timestamp, source, count, classes, and top confidence to log."]
    ]
    body.append(tbl(fr_headers, fr_rows, col_widths=[1400, 3200, 2600, 2600]))

    body.append(h2("3.6 Non-Functional Requirements Specification"))
    c3_6_intro = "Table 3.2 delineates the non-functional requirements (NFR) ensuring performance, reliability, usability, and maintainability."
    body.append(p(c3_6_intro))
    
    nfr_headers = ["Attribute", "Quality Standard / Specification", "Engineering Verification Metric"]
    nfr_rows = [
        ["Usability", "Intuitive, dark-themed UI; clear tab separation and tooltips.", "Zero training required; operational within 60 seconds."],
        ["Performance", "Single-image inference latency under 500 ms on standard CPU/GPU.", "Measured latency: 238 ms – 409 ms in testing."],
        ["Reliability", "Graceful exception handling for corrupted files or missing weights.", "Informative error banners; zero unhandled application crashes."],
        ["Maintainability", "Modular Python codebase; decoupled inference, UI, and logging.", "Self-contained app.py and clean data.yaml configuration."],
        ["Portability", "Cross-platform execution across Windows, Linux, and macOS runtimes.", "Pure Python implementation using standardized dependencies."],
        ["Safety Notice", "Explicit boundary disclaimers in UI and export documentation.", "Persistent footer disclaiming certified structural certification."]
    ]
    body.append(tbl(nfr_headers, nfr_rows, col_widths=[1800, 4600, 3200]))

    body.append(h2("3.7 Hardware Requirements"))
    c3_7 = (
        "The system was engineered across distinct development, training, and operational deployment environments:\n\n"
        "• Model Training Hardware (Google Colab Cloud GPU):\n"
        "  - Processor: Intel Xeon @ 2.20 GHz (virtualized)\n"
        "  - System Memory: 12.7 GB High-Speed RAM\n"
        "  - Acceleration Hardware: NVIDIA Tesla T4 GPU with 15.3 GB VRAM (CUDA 12.2)\n"
        "  - Storage: 100 GB Cloud NVMe SSD\n\n"
        "• Deployment and Inference Hardware (Host Machine):\n"
        "  - Processor: Intel Core i5 / i7 (8th Gen or higher) / AMD Ryzen 5 equivalent\n"
        "  - System Memory: 8 GB minimum (16 GB recommended for video decoding)\n"
        "  - Graphics: Integrated Intel Iris / UHD Graphics or dedicated NVIDIA GeForce GPU\n"
        "  - Storage: 10 GB free disk space for dataset, model weights, and Python virtual environment\n"
        "  - Display: Standard HD display (1920 × 1080 resolution recommended for dashboard visualization)\n"
        "  - Input Sensor: Integrated 720p/1080p HD Webcam or USB external inspection camera"
    )
    body.append(p(c3_7))

    body.append(h2("3.8 Software Requirements and Environment Setup"))
    c3_8_intro = "Table 3.3 summarizes the complete software technology stack and dependencies utilized in BuildingSafe AI."
    body.append(p(c3_8_intro))
    
    sw_headers = ["Software Component", "Version / Platform", "Operational Purpose / Role in System"]
    sw_rows = [
        ["Operating System", "Microsoft Windows 10 / 11 (64-bit)", "Primary host operating system and execution environment."],
        ["Programming Language", "Python 3.10.x", "Core programming runtime for model training and application."],
        ["Deep Learning Engine", "Ultralytics YOLOv8 (v8.0+)", "Anchor-free object detection, model training, and predict API."],
        ["Deep Learning Backend", "PyTorch (torch 2.0+ & torchvision)", "Tensor mathematical operations and CUDA GPU acceleration."],
        ["Web Dashboard Framework", "Streamlit (v1.25+)", "Browser-based interactive UI, file uploaders, and session state."],
        ["Computer Vision Library", "OpenCV (opencv-python 4.8+)", "Image decoding, color-space conversion, and video frame extraction."],
        ["Image Processing Library", "Pillow (PIL 9.5+)", "RGB image loading, format conversion, and byte buffer serialization."],
        ["Data Manipulation", "Pandas (v2.0+) & NumPy (v1.24+)", "Structuring detection tables, metric aggregations, and CSV export."]
    ]
    body.append(tbl(sw_headers, sw_rows, col_widths=[2200, 2600, 4800]))

    body.append(h2("3.9 Feasibility Analysis"))
    c3_9 = (
        "To ensure pragmatic viability, the project underwent formal multidimensional feasibility analysis:\n\n"
        "1. Technical Feasibility: The project utilizes mature, industry-standard technologies: PyTorch, Ultralytics YOLOv8, and Streamlit. "
        "The model training requirements were fully satisfied using Google Colab's cloud T4 GPU infrastructure. The fine-tuned YOLOv8n "
        "weights require only ~6.2 MB of storage and execute efficiently on local CPU architectures within ~250–400 ms. Therefore, technical "
        "feasibility is fully established.\n\n"
        "2. Operational Feasibility: The Streamlit interface eliminates all command-line complexity. Site inspectors interact with self-explanatory "
        "buttons, sliders, and upload widgets. Visualizations display direct overlays of bounding boxes accompanied by intuitive confidence "
        "meters. Operational feasibility is outstanding, requiring zero programming expertise from end users.\n\n"
        "3. Economic Feasibility: The software stack is composed entirely of free, open-source software (Python, OpenCV, Streamlit, Ultralytics, "
        "PyTorch). The training benchmark dataset (MBDD2025) is open-access under Creative Commons licensing. Cloud training incurred zero "
        "licensing expense via Google Colab. The prototype eliminates costly specialized commercial inspection software licenses, proving "
        "exceptional economic feasibility.\n\n"
        "4. Schedule Feasibility: The project was divided into systematic phases: dataset acquisition, VOC-to-YOLO conversion, model fine-tuning, "
        "web interface design, validation testing, and documentation. All technical milestones were successfully executed within the academic "
        "calendar allocated for Project Phase-1."
    )
    body.append(p(c3_9))
    body.append(page_break())

    # =========================================================================
    # CHAPTER 4: SYSTEM DESIGN
    # =========================================================================
    body.append(h1("CHAPTER 4\nSYSTEM DESIGN"))
    
    body.append(h2("4.1 Overall System Architecture"))
    c4_1 = (
        "BuildingSafe AI is architectured under a clean, decoupled two-phase paradigm distinguishing offline model development from "
        "online user-facing inference execution:\n\n"
        "• Phase 1: Model Development and Training Pipeline (Executed on Cloud GPU): The 14,471 images of the MBDD2025 benchmark and "
        "their corresponding VOC XML annotation files are parsed. Bounding coordinates are normalized into YOLO format text files. "
        "Images and labels are split into training (12,301) and validation (2,170) partitions. A YOLOv8n network is trained at 640 × 640 "
        "resolution across 20 epochs using Complete IoU (CIoU) loss. The resulting peak weights are serialized as best.pt.\n\n"
        "• Phase 2: Inspection Web Application (Executed in Streamlit): The application loads best.pt into memory. End users supply media "
        "through Image, Camera, or Video tabs. The YOLOv8 inference engine processes incoming frames, applies Non-Maximum Suppression (NMS), "
        "and filters detections using the interactive confidence threshold. Output modules render annotated imagery, summary metrics, "
        "coordinate tables, downloadable CSV logs, and an in-memory chronological session audit trail.\n\n"
        "Table 4.1 delineates the architectural layers and their functional responsibilities."
    )
    body.append(p(c4_1))

    arch_headers = ["Architectural Layer", "Components / Technologies", "Core Functional Responsibility"]
    arch_rows = [
        ["Media Ingestion Layer", "Streamlit file_uploader, camera_input, OpenCV", "Accepts image files, camera snapshots, and video streams; validates formats."],
        ["Preprocessing Layer", "PIL, OpenCV, NumPy", "Converts BGR to RGB; normalizes pixel channels; resizes to 640x640 tensor."],
        ["Deep Inference Layer", "Ultralytics YOLOv8n, PyTorch runtime", "Executes forward pass; extracts multi-scale features; predicts boxes and classes."],
        ["Post-Processing Layer", "NMS, CIoU Box Filter, Threshold Logic", "Suppresses overlapping boxes; filters predictions below user confidence threshold."],
        ["Visualization Layer", "Streamlit layout, Custom CSS, Pillow", "Renders side-by-side comparisons, confidence meters, and metric tiles."],
        ["Data & Export Layer", "Pandas, BytesIO, session_state", "Structures coordinate tables; generates CSV findings; logs session history."]
    ]
    body.append(tbl(arch_headers, arch_rows, col_widths=[2200, 3200, 4200]))

    body.append(p("[Fig. 4.1: System Architecture of BuildingSafe AI depicting Phase 1 Model Training on Colab GPU yielding best.pt, and Phase 2 Streamlit Web Application Ingestion and Inference Modules]", bold=True, italic=True, jc="center", size=20, space_before=120, space_after=180))

    body.append(h2("4.2 System Workflow and Data Flow Pipeline"))
    c4_2 = (
        "The end-to-end data pipeline follows a strictly deterministic progression: Media Selection → Byte Validation → Memory Buffering "
        "→ Tensor Normalization → YOLOv8 Forward Pass → Post-Processing (Threshold & NMS) → Coordinate Denormalization → UI Rendering & "
        "Session State Appending. This ensures predictable execution and zero data contamination between consecutive inspections."
    )
    body.append(p(c4_2))

    body.append(h2("4.3 Use-Case Modeling and Actor Analysis"))
    c4_3 = (
        "The primary actor interacting with BuildingSafe AI is the Structural Inspector / Site Maintenance Engineer. The core use cases "
        "supported by the system include:\n\n"
        "1. UC-01: Upload Building Photograph: Inspector selects an image from local storage. The system validates file format and displays "
        "an orthogonal input preview.\n"
        "2. UC-02: Capture Camera Photo: Inspector activates the device camera, frames the damaged surface, and triggers a snapshot.\n"
        "3. UC-03: Submit Inspection Video: Inspector submits a recorded UAV flyover video and configures the frame-skip sampling interval.\n"
        "4. UC-04: Adjust Confidence Sensitivity: Inspector modulates the sidebar slider (0.05 to 0.95) to filter candidate detections.\n"
        "5. UC-05: Execute Defect Inspection: Inspector clicks the action button; the system executes inference and displays bounding overlays.\n"
        "6. UC-06: Export Findings Artifacts: Inspector downloads the annotated image (PNG) and tabular findings (CSV) for formal archiving.\n"
        "7. UC-07: Review & Clear Session Audit Log: Inspector reviews the historical inspection table and optionally purges session history."
    )
    body.append(p(c4_3))

    body.append(h2("4.4 Data Flow Diagrams (Level 0 and Level 1)"))
    c4_4 = (
        "• DFD Level 0 (Context Diagram): The single external entity (Site Inspector) supplies raw inspection media and sensitivity parameters "
        "to the central process (BuildingSafe AI System). The system returns visual bounding-box overlays, class probabilities, spatial pixel "
        "coordinates, summary metrics, and downloadable CSV/PNG reports.\n\n"
        "• DFD Level 1 (Decomposition Diagram): Decomposes the central process into five interconnected sub-processes:\n"
        "  - Process 1.0 (Media Ingestion & Validation): Ingests images, snapshots, or video streams; validates file headers; buffers byte arrays.\n"
        "  - Process 2.0 (Preprocessing & Tensor Formatting): Decodes byte streams into RGB arrays; formats inputs to 640 × 640 tensor shapes.\n"
        "  - Process 3.0 (YOLOv8 Feature Extraction & Head Inference): Executes convolutional backbone, neck, and anchor-free prediction.\n"
        "  - Process 4.0 (Filtering & Post-Processing): Applies user confidence threshold; performs Non-Maximum Suppression.\n"
        "  - Process 5.0 (Visualization, Logging & Export): Maps boxes onto original pixels; builds Pandas DataFrames; updates session state."
    )
    body.append(p(c4_4))

    body.append(h2("4.5 Activity and Process Execution Models"))
    c4_5 = (
        "The activity workflow models the sequential logic executed by the application: The user selects an operational tab (Image, Camera, "
        "or Video). The system checks media availability. If an input is detected, the user initiates inspection via button click. A visual "
        "spinner indicates active neural network inference. The model output is validated: if no detections exceed the threshold, an "
        "informative 'No defects detected' banner is rendered; if defects are identified, the system renders side-by-side comparisons, builds "
        "per-class confidence bars, displays coordinate tables, and appends the event to session state history."
    )
    body.append(p(c4_5))

    body.append(h2("4.6 Image Inspection Sequence Workflow"))
    c4_6 = (
        "The image inspection sequence represents the primary operational pipeline:\n"
        "1. Inspector uploads a JPG, JPEG, or PNG file via st.file_uploader.\n"
        "2. Streamlit caches the file buffer; MD5 hash signature checks prevent redundant reprocessing if the same file is re-selected.\n"
        "3. User clicks 'Inspect building'.\n"
        "4. PIL converts the image to RGB format and passes it to the analyze() helper function.\n"
        "5. The model executes predict(image, conf=conf_threshold, verbose=False).\n"
        "6. Predicted bounding boxes are extracted as integer pixel tuples (x1, y1, x2, y2).\n"
        "7. OpenCV converts BGR annotated results to RGB for Streamlit container rendering.\n"
        "8. Metric tiles and expandable Pandas DataFrames are populated.\n"
        "9. Download buttons serialize image bytes (PNG) and DataFrame text (CSV).\n\n"
        "[Fig. 4.2: Image Inspection Logical Sequence Workflow]"
    )
    body.append(p(c4_6))

    body.append(h2("4.7 Camera Snapshot Architecture"))
    c4_7 = (
        "The camera snapshot module leverages Streamlit's native st.camera_input component. Upon user authorization, the browser accesses "
        "the local webcam or mobile camera feed. When the user clicks the shutter, a single still frame is frozen and transmitted as a byte "
        "buffer to the Python backend. The backend processes the captured photograph through the identical inference and rendering pipeline "
        "used for uploaded images. It is vital to emphasize that this architecture implements single-frame photographic capture, not continuous "
        "live-video streaming."
    )
    body.append(p(c4_7))

    body.append(h2("4.8 Video Frame Sampling and Analysis Architecture"))
    c4_8 = (
        "Video inspection presents substantial computational demands due to the massive volume of individual frames contained in standard "
        "30 FPS inspection videos. Processing every consecutive frame would overwhelm local CPU hardware, leading to excessive latency. "
        "To achieve a practical balance between inspection throughput and temporal coverage, BuildingSafe AI implements an intelligent "
        "frame-skip sampling architecture:\n\n"
        "1. Uploaded video bytes (MP4, MOV, AVI) are written to a secure temporary file on disk via tempfile.NamedTemporaryFile.\n"
        "2. OpenCV decodes the video via cv2.VideoCapture, extracting total frame count and native FPS.\n"
        "3. A while loop traverses video frames sequentially: if frame_index % frame_skip == 0, the frame is submitted to YOLOv8 inference.\n"
        "4. A dynamic Streamlit progress bar displays real-time frame progress (e.g., 'Analysing frame 120 of 600').\n"
        "5. An active visual slot displays real-time annotated frame previews.\n"
        "6. Defect occurrences are accumulated in a class tally dictionary. The system continuously tracks the global maximum confidence, "
        "storing the corresponding annotated frame, frame index, and timestamp.\n"
        "7. Upon loop termination, the temporary video file is safely unlinked from disk to prevent storage leaks, and a comprehensive video "
        "summary dashboard is rendered.\n\n"
        "[Fig. 4.3: Uploaded-Video Frame-Sampling Inspection Workflow]"
    )
    body.append(p(c4_8))

    body.append(h2("4.9 Session State and Inspection Log Design"))
    c4_9 = (
        "To provide operational traceability without introducing heavyweight database dependencies, the system utilizes Streamlit's in-memory "
        "st.session_state dictionary. Three session-bound keys are maintained:\n"
        "• st.session_state.history: An ordered list of dictionaries recording timestamp, input source name, total defects found, list of unique "
        "defect classes, and highest confidence percentage.\n"
        "• st.session_state.results: Stores cached inference outputs per active tab to prevent UI loss during tab navigation.\n"
        "• st.session_state.signatures: Stores MD5 checksum hashes of active media files, detecting when a user has changed the input file."
    )
    body.append(p(c4_9))

    body.append(h2("4.10 User Interface Ergonomics and Dark Theme Palette"))
    c4_10 = (
        "The visual design of BuildingSafe AI prioritizes ergonomic readability and professional aesthetics, adhering to an industrial "
        "dark charcoal palette inspired by high-reliability engineering control rooms:\n"
        "• Background: Deep Charcoal (#111827) minimizing ocular strain during prolonged inspections.\n"
        "• Surface Panels: Navy Grey (#1F2937 & #263548) establishing clear visual hierarchy.\n"
        "• Defect Class Color Coding: Standardized distinctive hues (Crack: #C0392B Crimson, Leakage: #1F6F8B Cyan, Corrosion: #A0522D Brown, "
        "Abscission: #B87400 Amber, Bulge: #6D4FA3 Purple).\n"
        "• Accent Highlights: Cyber Blue (#38BDF8) denoting active controls and interactive tabs."
    )
    body.append(p(c4_10))
    body.append(page_break())

    # =========================================================================
    # CHAPTER 5: METHODOLOGY
    # =========================================================================
    body.append(h1("CHAPTER 5\nMETHODOLOGY"))
    
    body.append(h2("5.1 Dataset Overview: MBDD2025 Building Defects Dataset"))
    c5_1 = (
        "The model training and validation pipeline utilizes the MBDD2025 (Multi-type Building Defect Dataset 2025) benchmark, curated by "
        "Zha et al. (2025). MBDD2025 was constructed specifically to overcome the severe limitations of earlier structural datasets, which "
        "typically focused on isolated concrete cracks under artificial lighting. MBDD2025 comprises 14,471 high-resolution images captured "
        "by Unmanned Aerial Vehicles (UAVs) under authentic outdoor field conditions across six diverse structural typologies: reinforced "
        "concrete, brick-concrete, masonry, steel, timber, and brick-wood structures.\n\n"
        "Table 5.1 outlines the primary characteristics and partitioning parameters of the dataset."
    )
    body.append(p(c5_1))

    ds_headers = ["Dataset Property", "Quantitative Value", "Engineering Description"]
    ds_rows = [
        ["Dataset Benchmark Name", "MBDD2025", "Multi-type Building Defect Dataset from UAV Images"],
        ["Total Number of Images", "14,471 images", "High-resolution UAV RGB photographs"],
        ["Defect Classes", "5 target classes", "crack, leakage, corrosion, abscission, bulge"],
        ["Training Split (85%)", "12,301 images", "Utilized exclusively for model parameter weight updates"],
        ["Validation Split (15%)", "2,170 images", "Held-out un-leaked evaluation corpus for monitoring generalization"],
        ["Input Resolution", "640 × 640 pixels", "Standardized tensor dimensions after bilinear interpolation"],
        ["Annotation Format", "YOLO normalized text", "class_idx x_center y_center width height (values in [0, 1])"]
    ]
    body.append(tbl(ds_headers, ds_rows, col_widths=[2800, 2400, 4400]))

    body.append(h2("5.2 Target Defect Taxonomy and Characteristics"))
    c5_2_intro = "Table 5.2 establishes the formal civil engineering taxonomy and visual characteristics of the five target defect classes."
    body.append(p(c5_2_intro))
    
    tax_headers = ["Defect Class", "Visual Appearance on Surface", "Underlying Structural Mechanism", "Inspection Priority"]
    tax_rows = [
        ["crack", "Linear, branching, or tortuous dark fracture lines.", "Tensile stress exceeding material strength; thermal expansion.", "High (Path for moisture ingress)"],
        ["leakage", "Dark, damp patches; efflorescence; mineral staining.", "Waterproofing failure; plumbing leaks; capillary moisture.", "Moderate to High (Promotes rebar rust)"],
        ["corrosion", "Reddish-brown ferric oxide bleeding; rebar exposure.", "Electrochemical oxidation of rebar; carbonation depassivation.", "Critical (Structural load reduction)"],
        ["abscission", "Spalled concrete cavities; delaminated surface tiles.", "Internal rust expansion bursting concrete cover; freeze-thaw.", "Critical (Falling debris hazard)"],
        ["bulge", "Convex surface swelling; outward plaster blister.", "Vapor pressure buildup; substrate delamination; buckling.", "High (Precedes spalling detachment)"]
    ]
    body.append(tbl(tax_headers, tax_rows, col_widths=[1600, 2800, 3200, 2000]))

    body.append(h2("5.3 Dataset Preprocessing and YOLO Coordinate Normalization"))
    c5_3 = (
        "The raw MBDD2025 dataset provides annotations in PASCAL VOC XML format. Each XML file defines bounding boxes using absolute pixel "
        "coordinates: xmin, ymin, xmax, and ymax. Because YOLO neural architectures require relative, scale-invariant coordinates, an automated "
        "normalization script was executed prior to model training.\n\n"
        "Let the native image width be W and height be H. The absolute coordinates are converted into normalized center coordinates (x_center, "
        "y_center) and normalized box dimensions (w_norm, h_norm) according to the following mathematical transformations:\n\n"
        "x_center = ((xmin + xmax) / 2.0) / W\n"
        "y_center = ((ymin + ymax) / 2.0) / H\n"
        "w_norm = (xmax - xmin) / W\n"
        "h_norm = (ymax - ymin) / H\n\n"
        "All normalized values strictly satisfy 0.0 ≤ x_center, y_center, w_norm, h_norm ≤ 1.0. For each image, a corresponding .txt file "
        "is generated containing one line per defect instance formatted as:\n"
        "<class_id> <x_center> <y_center> <w_norm> <h_norm>"
    )
    body.append(p(c5_3))

    body.append(h2("5.4 Dataset Partitioning Strategy"))
    c5_4 = (
        "The total corpus of 14,471 images was partitioned into training and validation sets using a stratified 85:15 split ratio implemented "
        "in split_dataset.py with a fixed pseudo-random seed (SEED = 42):\n"
        "• Training Set: 12,301 images (85.0%) with corresponding label files copied to dataset/train/images and dataset/train/labels.\n"
        "• Validation Set: 2,170 images (15.0%) copied to dataset/val/images and dataset/val/labels.\n\n"
        "The dataset configuration file data.yaml binds the training environment, defining dataset paths, class count (nc: 5), and the "
        "canonical class names list: [crack, leakage, corrosion, abscission, bulge].\n\n"
        "[Fig. 5.1: Dataset Preparation, Annotation Normalization, and YOLOv8n Training Pipeline]"
    )
    body.append(p(c5_4))

    body.append(h2("5.5 YOLOv8n Neural Network Architecture"))
    c5_5 = (
        "The object detection engine utilizes Ultralytics YOLOv8n (nano), representing the most advanced lightweight single-stage detector "
        "available. The architecture comprises three coordinated sub-networks:\n\n"
        "1. Backbone (CSPDarknet53 Feature Extractor): Employs a modified CSPDarknet backbone utilizing C2f (Cross Stage Partial with two "
        "convolutions) modules. C2f modules split feature channels into two branches, passing one through a sequence of Bottleneck blocks "
        "with residual shortcuts while concatenating across branches. This design enriches gradient flow, accelerates backpropagation, and "
        "preserves fine-grained feature representations across deep layers. The final backbone stage integrates a Spatial Pyramid Pooling - "
        "Fast (SPPF) module, pooling features at multiple receptive field scales (5×5, 9×9, 13×13) to capture contextual defect relationships.\n\n"
        "2. Neck (PANet Multi-Scale Feature Fusion): Utilizes a Path Aggregation Network (PANet) combining top-down feature pyramid flows with "
        "bottom-up path aggregation. High-level semantic features from deep layers are fused with low-level spatial features from shallow "
        "layers. This multi-scale aggregation is vital for civil defect inspection, where hairline cracks require high spatial resolution "
        "while expansive water leakage stains require broad receptive fields.\n\n"
        "3. Anchor-Free Decoupled Detection Head: Unlike legacy YOLO architectures that relied on anchor box aspect ratios, YOLOv8 implements "
        "an anchor-free detection head. Furthermore, the detection head is decoupled: separate convolutional branches independently compute "
        "class probabilities and bounding box regression coordinates, eliminating competition between classification and localization objectives.\n\n"
        "[Fig. 5.2: YOLOv8 Architectural Block Diagram illustrating CSPDarknet Backbone, C2f Modules, SPPF, PANet Neck, and Decoupled Head]"
    )
    body.append(p(c5_5))

    body.append(h2("5.6 Transfer Learning and Pretrained Weights"))
    c5_6 = (
        "Training deep convolutional networks from scratch on specialized civil engineering datasets frequently leads to severe overfitting "
        "due to the random initialization of millions of network weights. To maximize generalization and feature extraction efficiency, "
        "this project employs transfer learning. Pretrained weights (yolov8n.pt), pre-converged on the extensive Microsoft COCO benchmark, "
        "were loaded as the starting baseline. Early layers, which encode universal visual primitives such as edges, color gradients, and "
        "surface textures, were transferred directly, allowing training to focus computational capacity on fine-tuning higher-level feature "
        "representations unique to building defects."
    )
    body.append(p(c5_6))

    body.append(h2("5.7 Training Configuration and Cloud GPU Execution"))
    c5_7_intro = "Table 5.3 summarizes the exact hyperparameter configurations utilized during the model fine-tuning run on Google Colab."
    body.append(p(c5_7_intro))
    
    trn_headers = ["Hyperparameter", "Configured Value", "Operational Rationale"]
    trn_rows = [
        ["Base Model Architecture", "YOLOv8n (Nano)", "Lightweight model (~3.2M params) enabling rapid inference on local CPU/web runtimes."],
        ["Initial Pretrained Weights", "yolov8n.pt", "Transfers generalized visual edge and texture filters pre-trained on MS COCO."],
        ["Input Tensor Resolution", "640 × 640 pixels", "Standard multi-scale resolution balancing hairline crack detail with memory."],
        ["Training Epochs", "20 epochs", "Ensures model convergence without excessive over-fitting on training samples."],
        ["Batch Size", "16", "Optimizes gradient stability while conforming to GPU VRAM limits."],
        ["Optimizer", "SGD / AdamW (Auto)", "Ultralytics automatic optimizer selection with momentum = 0.937."],
        ["Initial Learning Rate (lr0)", "0.01", "Adaptive cosine learning rate scheduling decaying to lrf = 0.01."],
        ["Data Augmentations", "degrees=5, fliplr=0.5, mosaic=1.0", "Introduces rotation, horizontal reflection, and 4-image mosaic blending."],
        ["Hardware Device", "NVIDIA Tesla T4 GPU (16 GB)", "Google Colab cloud-accelerated CUDA runtime."],
        ["Training Duration", "~1.212 hours", "Total wall-clock training time recorded for the complete 20-epoch run."]
    ]
    body.append(tbl(trn_headers, trn_rows, col_widths=[2400, 2600, 4600]))

    body.append(h2("5.8 Loss Functions Formulation"))
    c5_8 = (
        "YOLOv8 formulates total training loss as a weighted composite of three specialized loss functions, optimized simultaneously:\n\n"
        "Total Loss = λ_box * L_CIoU + λ_cls * L_BCE + λ_dfl * L_DFL\n\n"
        "1. Complete Intersection over Union (CIoU) Loss (L_CIoU): Used for bounding-box coordinate regression. Traditional IoU loss "
        "provides zero gradient when predicted and ground truth boxes do not overlap. CIoU overcomes this by incorporating three geometric "
        "factors: overlapping area, normalized Euclidean distance between box center points, and aspect-ratio consistency:\n\n"
        "L_CIoU = 1 - IoU + (ρ^2(b, b_gt) / c^2) + α * v\n\n"
        "where b and b_gt denote the center coordinates of predicted and ground truth boxes, ρ is Euclidean distance, c is the diagonal "
        "length of the smallest enclosing bounding box, v measures aspect ratio discrepancy, and α is a positive trade-off parameter:\n\n"
        "v = (4 / π^2) * (arctan(w_gt / h_gt) - arctan(w / h))^2\n"
        "α = v / ((1 - IoU) + v)\n\n"
        "2. Distribution Focal Loss (DFL) (L_DFL): Recognizes that ground truth bounding box boundaries for irregular defects (such as "
        "amorphous water leakage or peeling paint) are often soft or ambiguous. Rather than predicting single deterministic coordinate floats, "
        "DFL models the boundary position as a continuous probability distribution, forcing the network to rapidly focus on values close "
        "to the true boundary.\n\n"
        "3. Binary Cross-Entropy (BCE) Loss (L_BCE): Evaluates multi-class classification accuracy across the decoupled head branches:\n\n"
        "L_BCE = - Σ [ y_i * log(p_i) + (1 - y_i) * log(1 - p_i) ]\n\n"
        "where y_i ∈ {0, 1} represents the ground truth class indicator and p_i represents predicted class probability."
    )
    body.append(p(c5_8))

    body.append(h2("5.9 Inference Pipeline, Confidence Thresholding, and NMS"))
    c5_9 = (
        "During operational inference, the trained model outputs thousands of candidate bounding-box predictions per image. Two sequential "
        "post-processing filters eliminate redundant and low-probability boxes:\n\n"
        "• Confidence Thresholding (τ): Let c denote predicted class confidence. Detections satisfying c < τ are discarded immediately. In "
        "BuildingSafe AI, the default threshold is calibrated to τ = 0.25 (25%), modifiable by the user from 0.05 to 0.95 via the sidebar slider.\n\n"
        "• Non-Maximum Suppression (NMS): Because the convolutional head frequently predicts multiple overlapping boxes around a single defect, "
        "NMS eliminates redundant boxes. The box with the highest confidence is selected, and all neighboring boxes sharing an IoU exceeding "
        "the NMS threshold (typically 0.45 to 0.70) are suppressed."
    )
    body.append(p(c5_9))

    body.append(h2("5.10 Quantitative Evaluation Metrics Formulation"))
    c5_10_intro = (
        "Model detection performance is quantified using three standard object detection metrics derived from the evaluation confusion matrix:\n\n"
        "• Precision (P): The fraction of predicted defect instances that correctly correspond to ground-truth defects:\n"
        "Precision = TP / (TP + FP)\n\n"
        "• Recall (R): The fraction of true ground-truth defect instances successfully detected by the model:\n"
        "Recall = TP / (TP + FN)\n\n"
        "where TP = True Positives (correctly identified defects with IoU ≥ 0.50), FP = False Positives (spurious alarms), and FN = False "
        "Negatives (missed defects).\n\n"
        "• Mean Average Precision at IoU 0.50 (mAP@50): Average Precision (AP) represents the area under the interpolated Precision-Recall curve "
        "for an individual class. mAP@50 computes the arithmetic mean of AP across all N = 5 defect categories at an IoU threshold of 0.50:\n"
        "mAP@50 = (1 / N) * Σ [ AP_i ] for i = 1 to N\n\n"
        "Table 5.4 summarizes the official reported evaluation benchmarks achieved by the fine-tuned YOLOv8n model."
    )
    body.append(p(c5_10_intro))
    
    met_headers = ["Evaluation Metric", "Reported Value (%)", "Academic Interpretation & Practical Meaning"]
    met_rows = [
        ["Overall Precision", "82.1%", "High precision indicates low false alarm rates; 82.1% of all flagged detections are authentic defects."],
        ["Overall Recall", "73.3%", "Solid recall indicating that 73.3% of all visible defect instances in the validation corpus are detected."],
        ["mAP@50", "81.5%", "Composite detection score averaging class precision-recall curves at standard 50% IoU overlap."],
        ["Training Duration", "~1.212 hours", "Fast convergence on cloud NVIDIA T4 GPU across 20 epochs on 12,301 training images."]
    ]
    body.append(tbl(met_headers, met_rows, col_widths=[2400, 2200, 5000]))
    body.append(page_break())

    # =========================================================================
    # CHAPTER 6: SYSTEM IMPLEMENTATION
    # =========================================================================
    body.append(h1("CHAPTER 6\nSYSTEM IMPLEMENTATION"))
    
    body.append(h2("6.1 Development Environment and Modular Structure"))
    c6_1 = (
        "The software architecture of BuildingSafe AI is consolidated within a clean, maintainable modular structure implemented in Python 3.10. "
        "The complete application logic, user interface layout, computer vision pipelines, and session management are orchestrated through "
        "app.py, supported by model configuration files and dataset directories.\n\n"
        "Table 6.1 delineates the primary repository components and their implementation roles."
    )
    body.append(p(c6_1))

    mod_headers = ["File / Directory", "Component Type", "Implementation Role & Technical Function"]
    mod_rows = [
        ["app.py", "Streamlit Application", "Core executable script defining UI layouts, tabs, inference pipelines, and session state."],
        ["best.pt", "Model Weights", "Fine-tuned YOLOv8n binary weights serialized from the optimal training epoch (~6.2 MB)."],
        ["data.yaml", "Configuration File", "Defines relative paths to train/val images, class count (5), and class name mappings."],
        ["split_dataset.py", "Preprocessing Script", "Automates 85:15 stratified dataset partitioning with deterministic random seeding."],
        ["train.py", "Training Script", "Configures Ultralytics YOLOv8 training hyperparameters, augmentations, and early stopping."],
        [".streamlit/config.toml", "UI Configuration", "Configures Streamlit server parameters, upload size limits (200MB), and base theme."]
    ]
    body.append(tbl(mod_headers, mod_rows, col_widths=[2200, 2400, 5000]))

    body.append(h2("6.2 Model Loading and In-Memory Caching"))
    c6_2 = (
        "Because Streamlit operates under a reactive re-execution model—rerunning the Python script from top to bottom upon every user widget "
        "interaction—naively instantiating the YOLO neural network on every rerun would cause catastrophic UI latency and memory leaks. "
        "BuildingSafe AI resolves this by encapsulating model loading within Streamlit's resource caching decorator:\n\n"
        "@st.cache_resource\n"
        "def load_model(path):\n"
        "    return YOLO(path)\n\n"
        "The weights file (best.pt) is parsed into GPU/CPU memory exactly once upon application startup. Subsequent script reruns reuse the "
        "persisted in-memory PyTorch model object, reducing inference initiation latency to zero. Robust try-except exception blocks intercept "
        "corrupted file paths, rendering informative error banners if the weights cannot be accessed."
    )
    body.append(p(c6_2))

    body.append(h2("6.3 Image Upload and Preprocessing Module"))
    c6_3 = (
        "The Image Inspection tab allows users to upload building photographs via st.file_uploader, restricting accepted extensions to "
        "['jpg', 'jpeg', 'png'] with a generous maximum file limit of 200 MB. To prevent stale detection overlays when a user selects a "
        "different file, the helper function reset_if_file_changed() computes an MD5 checksum of the uploaded byte stream. If the hash differs "
        "from the cached signature, prior results are purged from session state.\n\n"
        "Upon upload, PIL.Image opens the byte stream and converts the image to standard RGB. The user initiates inspection by clicking "
        "the primary 'Inspect building' button. A visual spinner (st.spinner) informs the user while the neural network executes inference."
    )
    body.append(p(c6_3))

    body.append(h2("6.4 Camera Snapshot Ingestion Module"))
    c6_4 = (
        "The Camera tab implements single-frame photographic capture using Streamlit's st.camera_input component. The browser captures a "
        "user-triggered snapshot and transmits the compressed image buffer to the backend. The snapshot is deserialized into an RGB Pillow "
        "image and routed to the central analyze() engine. The implementation explicitly handles camera state clearing, resetting session "
        "buffers when the camera input is cleared. This provides a lightweight on-site inspection mechanism without the massive hardware "
        "overhead of continuous real-time video streaming."
    )
    body.append(p(c6_4))

    body.append(h2("6.5 Video Processing and Frame-Skip Ingestion Engine"))
    c6_5 = (
        "The Video tab ingests recorded inspection videos (MP4, MOV, AVI). The uploaded byte stream is written to a secure temporary file "
        "via tempfile.NamedTemporaryFile. OpenCV (cv2.VideoCapture) opens the temporary video, reading native frame width, height, FPS, and "
        "total frame count. The inspection loop inspects every Nth frame, where N is governed by the sidebar frame_skip slider (default N = 5).\n\n"
        "Selected frames are submitted to model.predict(frame, conf=conf_threshold, verbose=False). An in-place visual container (st.empty) "
        "displays the real-time annotated frame stream. A class tally dictionary aggregates defect counts across the video timeline, while a "
        "tracking variable isolates the frame exhibiting the global highest confidence. At the conclusion of the video, OpenCV handles are "
        "released, temporary files are unlinked, and a video summary dashboard is populated with frequency bar charts and frame downloads."
    )
    body.append(p(c6_5))

    body.append(h2("6.6 Defect Detection, Bounding-Box Extraction, and Plotting"))
    c6_6 = (
        "The core inference engine is encapsulated within the analyze() function. The model returns an Ultralytics Results object containing "
        "detected bounding boxes (result.boxes). For each detected box, the following attributes are extracted:\n"
        "• Bounding Box Coordinates: Normalized tensor coordinates are scaled to native integer pixels [x1, y1, x2, y2].\n"
        "• Class Label: The integer class index (box.cls[0]) is mapped to the defect name via model.names.\n"
        "• Confidence Score: Extracted as a floating-point probability float(box.conf[0]).\n\n"
        "The annotated visual result is rendered using result.plot(), which draws color-coded rectangular boundaries and text badges, converted "
        "from BGR to RGB via cv2.cvtColor for seamless Streamlit display."
    )
    body.append(p(c6_6))

    body.append(h2("6.7 Per-Class Aggregation and Metric Gauges"))
    c6_7 = (
        "Raw detection boxes are grouped by defect class using the group_by_defect() function. For each detected defect type, the application "
        "calculates the total detection count, highest confidence score, and arithmetic mean confidence. These are rendered as custom HTML "
        "finding cards featuring styled CSS meter progress bars, visually communicating the model's certainty per defect category."
    )
    body.append(p(c6_7))

    body.append(h2("6.8 Tabular Findings Structuring and Data Serialization"))
    c6_8 = (
        "To provide engineers with traceable, actionable spatial data, all individual detections are structured into a Pandas DataFrame "
        "containing six columns: Defect, Confidence (%), x1, y1, x2, and y2. The table is rendered within an expandable UI container "
        "(st.expander('All individual detections')). Two prominent download buttons allow users to export the annotated image as a full-resolution "
        "PNG file (to_png_bytes) and the findings table as a comma-separated CSV text file."
    )
    body.append(p(c6_8))

    body.append(h2("6.9 Session State Logging and CSV Export Subsystem"))
    c6_9 = (
        "Every completed inspection event automatically triggers log_inspection(), appending a structured record to st.session_state.history. "
        "Each log entry captures: (1) Formatted Timestamp (e.g., '14 Oct 2026, 09:50:28'), (2) Media Source Identifier, (3) Total Defects "
        "Found, (4) Comma-separated list of Defect Types, and (5) Highest Confidence Percentage. The Inspection Log tab displays these "
        "historical records in reverse chronological order using st.dataframe, accompanied by a 'Download log (CSV)' button and a 'Clear log' "
        "action button."
    )
    body.append(p(c6_9))

    body.append(h2("6.10 Custom Dark Theme CSS and Responsive Layout"))
    c6_10 = (
        "The user interface incorporates advanced CSS injected via st.markdown(unsafe_allow_html=True). The styling establishes a modern "
        "IBM Plex Sans typographic hierarchy, rounded container borders (10px border-radius), subtle grid patterns across the application "
        "header, high-contrast metric tiles, and smooth hover transitions on primary blue action buttons. The layout automatically adapts "
        "across widescreen desktop monitors and tablet displays."
    )
    body.append(p(c6_10))

    body.append(h2("6.11 Robust Error Handling and Boundary Constraints"))
    c6_11 = (
        "The application incorporates comprehensive defensive programming principles:\n"
        "• Model Fallback: If best.pt cannot be located, an error container displays diagnostic technical advice.\n"
        "• Empty State Handling: Friendly dashed placeholder containers guide users when no media has been selected.\n"
        "• Unlinked Temporary Files: Temporary video files are wrapped in try-finally blocks to ensure deletion even if decoding fails.\n"
        "• Memory Safety: Video frames are decoded sequentially rather than accumulated into RAM, preventing memory exhaustion."
    )
    body.append(p(c6_11))
    body.append(page_break())

    # =========================================================================
    # CHAPTER 7: RESULTS AND DISCUSSION
    # =========================================================================
    body.append(h1("CHAPTER 7\nRESULTS AND DISCUSSION"))
    
    body.append(h2("7.1 Quantitative Model Training Evaluation"))
    c7_1 = (
        "The fine-tuned YOLOv8n object detection model was rigorously evaluated on the held-out validation partition of the MBDD2025 dataset, "
        "comprising 2,170 un-leaked building images. The reported evaluation metrics, benchmarked at an IoU threshold of 0.50, confirm "
        "high detection efficacy across complex outdoor civil infrastructure imagery.\n\n"
        "Table 7.1 summarizes the official validation performance metrics achieved by the trained detector."
    )
    body.append(p(c7_1))

    res_headers = ["Metric Parameter", "Reported Quantitative Result", "Practical Interpretation in Civil Inspection"]
    res_rows = [
        ["Overall Model Precision", "82.1%", "Low false positive alarm rate; 82.1% of all detected defect boxes represent true anomalies."],
        ["Overall Model Recall", "73.3%", "Solid defect coverage; 73.3% of all visible ground-truth defect instances are detected."],
        ["Mean Average Precision (mAP@50)", "81.5%", "Composite detection benchmark averaging precision-recall performance at 50% IoU."],
        ["Total Training Duration", "~1.212 hours", "Total wall-clock training time across 20 epochs on Google Colab NVIDIA T4 GPU."],
        ["Model File Size (best.pt)", "6.22 MB", "Extremely compact binary footprint enabling instantaneous browser caching."]
    ]
    body.append(tbl(res_headers, res_rows, col_widths=[2600, 2400, 4600]))

    body.append(h2("7.2 Convergence Characteristics and Training Duration"))
    c7_2 = (
        "The training progression executed smoothly over 20 epochs. The combination of transfer learning from pretrained COCO weights, "
        "Complete IoU (CIoU) loss, and Distribution Focal Loss (DFL) facilitated rapid parameter convergence. Box regression loss and "
        "classification loss exhibited steep declines during the initial 5 epochs, reaching a stable plateau between epochs 15 and 20. "
        "The recorded training duration of approximately 1.212 hours demonstrates the exceptional computational efficiency of the YOLOv8n "
        "architecture when deployed on cloud GPU infrastructure."
    )
    body.append(p(c7_2))

    body.append(h2("7.3 Analysis of Application Landing Interface (Fig. 7.1)"))
    c7_3 = (
        "Figure 7.1 illustrates the primary landing screen of the BuildingSafe AI web application upon initialization in a browser environment. "
        "The application header displays the project branding, accompanied by four top status tiles: (1) AI Model Status ('Ready'), (2) Active "
        "Confidence Threshold ('25%'), (3) Number of Detectable Classes ('5'), and (4) Detection Engine ('YOLOv8'). The left sidebar provides "
        "the sensitivity slider, a color-coded legend of the five defect classes, usage guidelines, and expandable advanced model path settings. "
        "The main dashboard presents an empty upload container prompting the user to select an image."
    )
    body.append(p(c7_3))
    
    if "fig7_1" in available_images:
        body.append(img_xml(1, cx=5800000, cy=2700000, caption="Fig. 7.1: BuildingSafe AI home screen showing the dark-themed interface, model status, confidence threshold, defect classes, and input tabs."))
    else:
        body.append(p("[Fig. 7.1: BuildingSafe AI Home Screen]", bold=True, italic=True, jc="center", size=20))

    body.append(h2("7.4 Analysis of Pre-Inspection Image Ingestion (Fig. 7.2)"))
    c7_4 = (
        "Figure 7.2 depicts the Image Inspection interface following the successful upload of a test photograph (Hefei8702.jpg, file size 269.8 KB). "
        "The application displays the input image under the 'Input' column, showing a weathered outdoor perimeter wall with tree branches and "
        "graffiti markings in the foreground. The 'AI inspection' panel remains in a clean 'Ready to inspect' state, waiting for user "
        "initiation. This explicit user-triggered action prevents wasteful background computation before the operator is ready."
    )
    body.append(p(c7_4))
    
    if "fig7_2" in available_images:
        body.append(img_xml(2, cx=5800000, cy=2700000, caption="Fig. 7.2: Image inspection screen after a building image has been selected and before inference is triggered."))
    else:
        body.append(p("[Fig. 7.2: Image Inspection Pre-Inference State]", bold=True, italic=True, jc="center", size=20))

    body.append(h2("7.5 Analysis of Leakage Defect Detection (Fig. 7.3)"))
    c7_5 = (
        "Figure 7.3 illustrates the inspection output generated after clicking 'Inspect building'. The deep neural network successfully "
        "identified and localized a prominent moisture seepage anomaly on the masonry facade, plotting a cyan bounding box labeled 'leakage 0.71'. "
        "A notification strip highlights: '1 defect flagged across 1 type'. Four quantitative metric tiles report: (1) Defects Found: 1, "
        "(2) Highest Confidence: 71.1%, (3) Defect Types: 1, and (4) Analysis Time: 409 ms. Crucially, the model ignored background trees and "
        "wall graffiti, correctly focusing solely on the structural moisture stain."
    )
    body.append(p(c7_5))
    
    if "fig7_3" in available_images:
        body.append(img_xml(3, cx=5800000, cy=2700000, caption="Fig. 7.3: Original and annotated image comparison showing a leakage detection with a bounding box."))
    else:
        body.append(p("[Fig. 7.3: Annotated Detection Output]", bold=True, italic=True, jc="center", size=20))

    body.append(h2("7.6 Analysis of Metric Cards and Coordinate Table (Fig. 7.4)"))
    c7_6 = (
        "Figure 7.4 displays the detailed analytical breakdown rendered below the visual comparison. A class finding card for 'Leakage' displays "
        "a visual progress bar indicating 71.1% top confidence, noting '1 detection, average confidence 71%'. The expanded 'All individual detections' "
        "table exposes exact pixel coordinates: x1 = 610, y1 = 185, x2 = 781, and y2 = 475. Prominent download buttons provide one-click export "
        "of the annotated PNG image and the findings CSV table.\n\n"
        "Table 7.2 details the exact quantitative data extracted from this test run."
    )
    body.append(p(c7_6))

    t7_2_headers = ["Attribute / Field", "Extracted Value", "Engineering Significance"]
    t7_2_rows = [
        ["Test Image Source", "Hefei8702.jpg", "Outdoor masonry wall under natural lighting conditions."],
        ["Detected Defect Class", "Leakage", "Vertical water stain / moisture seepage pattern."],
        ["Statistical Confidence", "71.1%", "High certainty of defect category match."],
        ["Bounding Box (x1, y1)", "(610, 185)", "Top-left boundary corner in original image coordinate space."],
        ["Bounding Box (x2, y2)", "(781, 475)", "Bottom-right boundary corner in original image coordinate space."],
        ["Inference Latency", "409 ms", "Rapid response time on standard CPU hardware runtime."]
    ]
    body.append(tbl(t7_2_headers, t7_2_rows, col_widths=[2600, 2400, 4600]))

    if "fig7_4" in available_images:
        body.append(img_xml(4, cx=5800000, cy=2700000, caption="Fig. 7.4: Detection summary, class grouping, confidence percentage, and bounding-box coordinate table."))
    else:
        body.append(p("[Fig. 7.4: Detailed Coordinate Table]", bold=True, italic=True, jc="center", size=20))

    body.append(h2("7.7 Analysis of Multi-Defect Co-Occurrence (Crack & Leakage)"))
    c7_7 = (
        "During multi-defect stress testing on severely degraded wall surfaces, the system demonstrated exceptional multi-class localization "
        "capacity. In an experimental test image exhibiting compound deterioration, BuildingSafe AI simultaneously detected two distinct defects: "
        "(1) a vertical crack localized with a blue bounding box at 90.8% confidence (x1: 256, y1: 194, x2: 475, y2: 660), and (2) an adjacent "
        "moisture seepage stain localized with a cyan box at 86.4% confidence (x1: 794, y1: 189, x2: 1052, y2: 624). The entire multi-defect "
        "inference was completed in just 238 ms, verifying the model's capacity to decouple co-occurring structural symptoms."
    )
    body.append(p(c7_7))

    body.append(h2("7.8 Analysis of Session Inspection Audit Log (Fig. 7.5)"))
    c7_8 = (
        "Figure 7.5 illustrates the Inspection Log tab following inspection execution. A structured tabular record documents: (1) Time: "
        "'14 Oct 2026, 09:50:28', (2) Source: 'Image: Hefei8702.jpg', (3) Defects Found: 1, (4) Defect Types: 'leakage', and (5) Highest "
        "Confidence: '71.1%'. The interface provides a 'Download log (CSV)' button allowing inspectors to export cumulative survey data, "
        "along with a 'Clear log' action button to reset session history."
    )
    body.append(p(c7_8))
    
    if "fig7_5" in available_images:
        body.append(img_xml(5, cx=5800000, cy=2700000, caption="Fig. 7.5: Session-level inspection log showing the timestamp, source image, detection count, defect type, and highest confidence."))
    else:
        body.append(p("[Fig. 7.5: Session Inspection Log Screen]", bold=True, italic=True, jc="center", size=20))

    body.append(h2("7.9 Video Inspection Experimental Evaluation"))
    c7_9 = (
        "Video inspection testing was conducted using recorded UAV flyover video clips (MP4 format). Utilizing a frame-skip interval of "
        "N = 5 on a 30 FPS video, the system effectively processed 6 inspection frames per second of video time. The system successfully "
        "accumulated frequency counts across consecutive frames, dynamically updating class bar charts and isolating the peak-confidence "
        "frame for individual PNG download. This verifies the practical utility of the frame-skip sampling architecture for drone video screening."
    )
    body.append(p(c7_9))

    body.append(h2("7.10 System Latency, Throughput, and Resource Profiling"))
    c7_10 = (
        "Inference latency was profiled across CPU and GPU hardware environments. On the host Intel Core i5 test machine without dedicated "
        "GPU acceleration, single-image inference latency consistently ranged between 238 ms and 409 ms. RAM consumption remained exceptionally "
        "stable, peaking at approximately 450 MB during video decoding. These lightweight resource profiles confirm that BuildingSafe AI can "
        "be deployed on standard field laptops without requiring expensive external GPU hardware."
    )
    body.append(p(c7_10))

    body.append(h2("7.11 Critical Evaluation of Strengths and Failure Modes"))
    c7_11 = (
        "A rigorous engineering evaluation identifies clear system strengths alongside specific operational failure modes:\n\n"
        "• Strengths: Exceptional localization precision on prominent cracks and leakage; robust background rejection against foliage and "
        "graffiti; rapid sub-second inference latency; clean, intuitive user interface; and traceable coordinate logging.\n\n"
        "• Failure Modes & Edge Cases: Extremely fine hairline micro-cracks (<1 pixel width when photographed from distant UAV altitudes) "
        "were occasionally missed, requiring lower confidence thresholds. Deep, sharp cast shadows across rough stone masonry occasionally "
        "triggered false positive crack detections. Highly reflective glass facade reflections also introduced occasional false alarms. "
        "These observations reinforce the project's foundational premise: BuildingSafe AI serves as a preliminary visual screening assistant "
        "that directs human attention, rather than an autonomous replacement for certified structural engineers."
    )
    body.append(p(c7_11))
    body.append(page_break())

    # =========================================================================
    # CHAPTER 8: CONCLUSION AND FUTURE ENHANCEMENT
    # =========================================================================
    body.append(h1("CHAPTER 8\nCONCLUSION AND FUTURE ENHANCEMENT"))
    
    body.append(h2("8.1 Summary of Completed Work"))
    c8_1 = (
        "This final-year engineering project successfully designed, implemented, and evaluated BuildingSafe AI—an AI-based safety inspection "
        "assistance system for construction buildings. Addressing the severe safety hazards, high financial costs, and subjectivity of manual "
        "building inspection, the project developed an integrated end-to-end framework uniting deep object detection with an accessible web "
        "dashboard.\n\n"
        "A lightweight, anchor-free YOLOv8n neural network was fine-tuned on 14,471 images from the MBDD2025 building defects benchmark, "
        "divided into 12,301 training and 2,170 validation images. Training was conducted on an NVIDIA T4 cloud GPU at 640 × 640 resolution "
        "over 20 epochs, completing in ~1.212 hours. The model achieved strong quantitative benchmarks: Precision of 82.1%, Recall of 73.3%, "
        "and mAP@50 of 81.5% across five structural defect classes: crack, leakage, corrosion, abscission, and bulge.\n\n"
        "The model was integrated into a responsive Streamlit web interface supporting image uploads, camera snapshots, and uploaded video "
        "scanning. The application provides side-by-side visual comparisons, pixel-level bounding coordinates, per-class confidence gauges, "
        "session audit logging, and downloadable CSV/PNG reports."
    )
    body.append(p(c8_1))

    body.append(h2("8.2 Attainment of Project Objectives"))
    c8_2 = (
        "All primary project objectives established in Section 1.6 were systematically achieved:\n"
        "• Objective 1 (Dataset Preparation): MBDD2025 dataset converted and normalized into standard YOLO bounding-box text formats.\n"
        "• Objective 2 (Data Partitioning): Stratified 85:15 split established (12,301 train, 2,170 val) with zero data leakage.\n"
        "• Objective 3 (Model Training): YOLOv8n fine-tuned on GPU with CIoU and DFL losses, saving optimal weights as best.pt.\n"
        "• Objective 4 (Multi-Defect Localization): Achieved 81.5% mAP@50 across all five target structural defect categories.\n"
        "• Objective 5 (Interactive UI): Responsive Streamlit dashboard engineered with industrial dark-themed styling.\n"
        "• Objective 6 (Artifact Export): Implemented instant CSV findings download, annotated PNG export, and session audit trails.\n"
        "• Objective 7 (Architectural Foundation): Established clean, decoupled architecture prepared for Phase-2 extensions."
    )
    body.append(p(c8_2))

    body.append(h2("8.3 Core Technical Contributions"))
    c8_3 = (
        "The primary engineering contributions of this project encompass:\n"
        "1. Applied Domain Integration: Bridged the gap between abstract neural network research and practical civil inspection by deploying "
        "an anchor-free detector into an intuitive, zero-code web interface.\n"
        "2. Traceable Spatial Localization: Provided exact bounding pixel coordinates (x1, y1, x2, y2) rather than unlocalized classifications, "
        "enabling precise spatial documentation.\n"
        "3. Multi-Modal Ingestion Pipeline: Engineered unified pipelines supporting still photographs, camera snapshots, and video frame sampling.\n"
        "4. Ethical Boundary Adherence: Maintained rigorous academic honesty by defining the system strictly as a visual screening tool, "
        "disclaiming unverified structural safety guarantees."
    )
    body.append(p(c8_3))

    body.append(h2("8.4 Current Prototype Limitations"))
    c8_4 = (
        "The identified limitations of the current prototype include:\n"
        "• Surface Visual Screening Only: Incapable of measuring subsurface voids, rebar stress, or internal foundation movement.\n"
        "• Single-Snapshot Camera: Camera module captures static individual photos rather than continuous live WebRTC video streams.\n"
        "• Session-Bound Logging: Inspection history is stored in temporary session state and clears upon browser reload.\n"
        "• Environmental Sensitivity: Occasional false alarms triggered by sharp cast shadows, rough foliage, and reflective facade glass."
    )
    body.append(p(c8_4))

    body.append(h2("8.5 Strategic Future Roadmap"))
    c8_5 = (
        "The modular architecture established in BuildingSafe AI provides a strong foundation for future research enhancements:\n\n"
        "1. Structural Damage Severity Index: Formulate a validated engineering severity index combining bounding box surface area ratios "
        "with co-occurrence rules (e.g., crack + moisture leakage = High Structural Hazard).\n"
        "2. Automated Engineering PDF Report Generator: Integrate libraries such as ReportLab or WeasyPrint to automatically compile formal "
        "inspection certificates complete with college/firm branding, timestamps, side-by-side photos, and repair recommendations.\n"
        "3. Persistent Relational Database: Replace session state with a persistent SQLite or PostgreSQL database to track long-term defect "
        "progression across successive annual inspection audits.\n"
        "4. Real-time Live WebRTC Streaming: Integrate streamlit-webrtc to enable continuous, real-time video detection from mobile cameras "
        "or live drone feeds.\n"
        "5. Photogrammetry and 3D Digital Twins: Map 2D bounding box detections onto 3D building point clouds generated via Structure-from-Motion "
        "(SfM) photogrammetry."
    )
    body.append(p(c8_5))

    body.append(h2("8.6 Final Concluding Remarks"))
    c8_6 = (
        "BuildingSafe AI demonstrates the immense potential of artificial intelligence and computer vision to transform traditional civil "
        "engineering inspection workflows. By automating the visual screening bottleneck, the system empowers inspectors to identify structural "
        "defects rapidly, objectively, and safely. While human expertise remains irreplaceable for final structural certification, BuildingSafe AI "
        "stands as a robust, practical, and technologically sophisticated step toward the intelligent, automated infrastructure monitoring of "
        "tomorrow."
    )
    body.append(p(c8_6))
    body.append(page_break())

    # =========================================================================
    # REFERENCES
    # =========================================================================
    body.append(h1("REFERENCES"))
    refs = [
        "[1] J. Redmon, S. Divvala, R. Girshick, and A. Farhadi, \"You Only Look Once: Unified, Real-Time Object Detection,\" in Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (CVPR), Las Vegas, NV, USA, 2016, pp. 779–788. doi: 10.1109/CVPR.2016.91.",
        "[2] Y.-J. Cha, W. Choi, and O. Büyüköztürk, \"Deep Learning-Based Crack Damage Detection Using Convolutional Neural Networks,\" Computer-Aided Civil and Infrastructure Engineering, vol. 32, no. 5, pp. 361–378, 2017. doi: 10.1111/mice.12263.",
        "[3] S. Ren, K. He, R. Girshick, and J. Sun, \"Faster R-CNN: Towards Real-Time Object Detection with Region Proposal Networks,\" IEEE Transactions on Pattern Analysis and Machine Intelligence, vol. 39, no. 6, pp. 1137–1149, Jun. 2017. doi: 10.1109/TPAMI.2016.2577031.",
        "[4] C. V. Dung and L. D. Anh, \"Autonomous Concrete Crack Detection Using Deep Fully Convolutional Neural Network with an End-to-End Approach,\" Automation in Construction, vol. 99, pp. 52–71, 2019. doi: 10.1016/j.autcon.2018.12.004.",
        "[5] Z. Zhang, X. Zhang, C. Yang, and L. Zhang, \"Automatic Crack Detection of Concrete Pavement Using Convolutional Neural Network,\" in Proceedings of the IEEE International Conference on Image Processing (ICIP), Phoenix, AZ, USA, 2016, pp. 4108–4112. doi: 10.1109/ICIP.2016.7533132.",
        "[6] Q. Zha, Y. Yao, W. Zhang, and W. Ma, \"MBDD2025: A Deep Learning-Oriented Dataset for Diverse Building Defect Detection from UAV Images,\" Zenodo, 2025. doi: 10.5281/zenodo.15622584.",
        "[7] Ultralytics, \"Ultralytics YOLOv8 Documentation and Software Repository,\" 2023. [Online]. Available: https://github.com/ultralytics/ultralytics (accessed Oct. 14, 2026).",
        "[8] Streamlit, \"Streamlit Documentation: The Fastest Way to Build and Share Data Apps,\" 2023. [Online]. Available: https://docs.streamlit.io/ (accessed Oct. 14, 2026).",
        "[9] G. Bradski, \"The OpenCV Library,\" Dr. Dobb's Journal of Software Tools, 2000. [Online]. Available: https://docs.opencv.org/ (accessed Oct. 14, 2026).",
        "[10] Python Software Foundation, \"Python 3.10 Language Documentation,\" 2023. [Online]. Available: https://docs.python.org/3/ (accessed Oct. 14, 2026)."
    ]
    for r in refs:
        body.append(p(r, jc="left", size=22, space_after=140, line_spacing=320))
    body.append(page_break())

    # =========================================================================
    # APPENDICES
    # =========================================================================
    body.append(h1("APPENDIX A\nUSER AND OPERATOR GUIDE"))
    app_a = (
        "A.1 System Deployment Prerequisites\n"
        "To execute BuildingSafe AI on a local workstation, ensure the following prerequisites are installed:\n"
        "• Python 3.10.x 64-bit runtime installed with PATH environment variables configured.\n"
        "• A dedicated Python virtual environment containing: streamlit, ultralytics, opencv-python, pillow, pandas, numpy.\n"
        "• The trained model weights file (best.pt) placed in the root directory alongside app.py.\n\n"
        "A.2 Launching the Web Interface\n"
        "1. Open a terminal or PowerShell prompt in the project working directory.\n"
        "2. Activate the virtual environment: .\\venv\\Scripts\\activate\n"
        "3. Launch Streamlit: streamlit run app.py\n"
        "4. The system will open the dashboard in your default browser at http://localhost:8501.\n\n"
        "A.3 Operating the Inspection Modules\n"
        "• Image Tab: Drag and drop an image file (JPG, PNG). Verify the orthogonal preview under 'Input'. Adjust the confidence "
        "threshold in the sidebar (default 0.25). Click 'Inspect building'. Review the highlighted bounding boxes and coordinates table. "
        "Click download buttons to export annotated PNG or CSV findings.\n"
        "• Camera Tab: Allow browser camera permissions. Frame the surface anomaly. Click 'Take Photo'. Click 'Inspect camera photo'.\n"
        "• Video Tab: Upload an inspection video (MP4, MOV, AVI). Set the frame-skip slider (default 5). Click 'Start video inspection'. "
        "Monitor the real-time progress bar. Review cumulative class bar charts and download the peak-confidence frame.\n"
        "• Inspection Log Tab: Review cumulative session history. Click 'Download log (CSV)' to save the audit file."
    )
    body.append(p(app_a))
    body.append(page_break())

    body.append(h1("APPENDIX B\nCORE INFERENCE SOURCE CODE"))
    app_b_intro = (
        "B.1 Primary Inference Engine Logic (Extract from app.py)\n"
        "The following Python snippet illustrates the core analyze() and caching functions orchestrating YOLOv8 inference:"
    )
    body.append(p(app_b_intro))
    
    code_text = (
        "@st.cache_resource\n"
        "def load_model(path):\n"
        "    return YOLO(path)\n\n"
        "def analyze(image, source_label):\n"
        "    \"\"\"Run the YOLOv8 model on one image and return structured results.\"\"\"\n"
        "    result = model.predict(image, conf=conf_threshold, verbose=False)[0]\n"
        "    rows = []\n"
        "    for box in result.boxes:\n"
        "        x1, y1, x2, y2 = [int(v) for v in box.xyxy[0].tolist()]\n"
        "        rows.append({\n"
        "            'defect': model.names[int(box.cls[0])],\n"
        "            'confidence': float(box.conf[0]),\n"
        "            'box': (x1, y1, x2, y2)\n"
        "        })\n"
        "    return {\n"
        "        'source': source_label,\n"
        "        'time': datetime.now().strftime('%H:%M:%S'),\n"
        "        'threshold': conf_threshold,\n"
        "        'rows': rows,\n"
        "        'original': cv2.cvtColor(result.orig_img, cv2.COLOR_BGR2RGB),\n"
        "        'annotated': cv2.cvtColor(result.plot(), cv2.COLOR_BGR2RGB),\n"
        "        'ms': sum(result.speed.values())\n"
        "    }"
    )
    body.append(p(code_text, style="Normal", jc="left", size=18, space_after=180, line_spacing=240, color="1F497D"))
    
    app_b_outro = (
        "B.2 Operational Verification Checklist\n"
        "[✓] Verify best.pt exists and matches the 5 MBDD2025 defect classes.\n"
        "[✓] Confirm input images are in focus and well illuminated.\n"
        "[✓] Record active confidence threshold for every experimental run.\n"
        "[✓] Verify that session log is exported to CSV before shutting down the server.\n"
        "[✓] Note that camera snapshots represent single photos, not continuous WebRTC streams.\n"
        "[✓] Explicitly state in reports that the system does not replace a certified civil engineer."
    )
    body.append(p(app_b_outro))

    # =========================================================================
    # ZIP PACKAGING (PURE PYTHON WORDPROCESSINGML GENERATOR)
    # =========================================================================
    
    # 1. [Content_Types].xml
    content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                     '<Default Extension="xml" ContentType="application/xml"/>',
                     '<Default Extension="png" ContentType="image/png"/>',
                     '<Default Extension="jpeg" ContentType="image/jpeg"/>',
                     '<Default Extension="jpg" ContentType="image/jpeg"/>',
                     '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
                     '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
                     '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
                     '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
                     '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',
                     '</Types>']
    
    # 2. _rels/.rels
    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
            '</Relationships>']
    
    # 3. word/_rels/document.xml.rels
    doc_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
                '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>',
                '<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
                '<Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>']
    
    # Add image relationships
    img_rel_map = {"fig7_1": 1, "fig7_2": 2, "fig7_3": 3, "fig7_4": 4, "fig7_5": 5}
    for key, num in img_rel_map.items():
        if key in available_images:
            doc_rels.append(f'<Relationship Id="rId{num}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image{num}.png"/>')
    doc_rels.append('</Relationships>')

    # 4. word/styles.xml
    styles_xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
                  '<w:docDefaults>',
                  '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:color w:val="000000"/></w:rPr></w:rPrDefault>',
                  '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault>',
                  '</w:docDefaults>',
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>',
                  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="240" w:after="180" w:line="360" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>',
                  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="180" w:after="120" w:line="360" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>',
                  '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="120" w:after="80" w:line="360" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>',
                  '</w:styles>']

    # 5. word/settings.xml
    settings_xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
                    '<w:doNotHyphenateCaps/>',
                    '</w:settings>']

    # 6. word/header1.xml
    header_xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
                  '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="4" w:color="CCCCCC"/></w:pBdr><w:jc w:val="right"/><w:spacing w:after="120"/></w:pPr>',
                  '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="18"/><w:color w:val="666666"/></w:rPr><w:t>BuildingSafe AI: Final Year Project Report | Dept. of AI &amp; DS, S.A. Engineering College</w:t></w:r>',
                  '</w:p></w:hdr>']

    # 7. word/footer1.xml
    footer_xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                  '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
                  '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="4" w:color="CCCCCC"/></w:pBdr><w:jc w:val="center"/><w:spacing w:before="120"/></w:pPr>',
                  '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr><w:fldSimple w:instr="PAGE"/></w:r>',
                  '</w:p></w:ftr>']

    # 8. word/document.xml
    doc_body = "".join(body)
    
    # Section properties: Margins (Left 1.25 in = 1800 dxa, Right/Top/Bottom 1.0 in = 1440 dxa)
    sect_pr = ('<w:sectPr>'
               '<w:headerReference w:type="default" r:id="rIdHeader" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
               '<w:footerReference w:type="default" r:id="rIdFooter" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
               '<w:pgSz w:w="12240" w:h="15840"/>'
               '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>'
               '<w:cols w:space="720"/>'
               '<w:docGrid w:linePitch="360"/>'
               '</w:sectPr>')
    
    document_xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                    f'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
                    f'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
                    f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                    f'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">\n'
                    f'<w:body>{doc_body}{sect_pr}</w:body>\n'
                    f'</w:document>')

    # Compile the ZIP archive
    with zipfile.ZipFile(output_filename, 'w', zipfile.ZIP_DEFLATED) as docx:
        docx.writestr('[Content_Types].xml', "\n".join(content_types))
        docx.writestr('_rels/.rels', "\n".join(rels))
        docx.writestr('word/_rels/document.xml.rels', "\n".join(doc_rels))
        docx.writestr('word/styles.xml', "\n".join(styles_xml))
        docx.writestr('word/settings.xml', "\n".join(settings_xml))
        docx.writestr('word/header1.xml', "\n".join(header_xml))
        docx.writestr('word/footer1.xml', "\n".join(footer_xml))
        docx.writestr('word/document.xml', document_xml)
        
        # Embed actual screenshot images
        for key, num in img_rel_map.items():
            if key in available_images:
                with open(available_images[key], 'rb') as f_img:
                    docx.writestr(f'word/media/image{num}.png', f_img.read())
                    print(f"  [+] Embedded image: word/media/image{num}.png")

    print(f"\n[SUCCESS] Report successfully generated: {os.path.abspath(output_filename)}")
    print(f"File size: {os.path.getsize(output_filename):,} bytes")
    print(f"Document contains: Title Page, Bonafide Certificate, Declaration, Acknowledgement, Abstract,")
    print(f"TOC, LOF, LOT, Abbreviations, Chapters 1-8, References, Appendices A & B, and embedded screenshots.")

if __name__ == "__main__":
    out_file = sys.argv[1] if len(sys.argv) > 1 else "BuildingSafe_AI_Final_Project_Report.docx"
    create_docx(out_file)

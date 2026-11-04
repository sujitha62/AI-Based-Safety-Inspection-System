import hashlib
import io
import os
import tempfile
from datetime import datetime

import cv2
import pandas as pd
import streamlit as st
from PIL import Image
from ultralytics import YOLO


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="BuildingSafe AI",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# One colour per defect class (readable on dark cards)
CLASS_COLORS = {
    "crack": "#C0392B",
    "leakage": "#1F6F8B",
    "corrosion": "#A0522D",
    "abscission": "#B87400",
    "bulge": "#6D4FA3",
}
DEFAULT_COLOR = "#5F6C7A"


# ============================================================
# CUSTOM UI / CSS  (dark charcoal theme)
# Tip: also add .streamlit/config.toml so Streamlit's own widgets
# use the light theme too (see config.toml provided with this file).
# ============================================================

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

:root {
    --bg: #111827;
    --surface: #1F2937;
    --surface2: #263548;
    --border: #374151;
    --text: #F3F4F6;
    --muted: #9CA3AF;
    --blue: #38BDF8;
    --amber: #F59E0B;
    --green: #34D399;
    --red: #F87171;
}

html, body, .stApp {
    background-color: var(--bg);
    color: var(--text);
    font-family: 'IBM Plex Sans', sans-serif;
}

#MainMenu, footer { visibility: hidden; }

.block-container {
    padding-top: 1.6rem;
    max-width: 1280px;
}

/* ---------- Header ---------- */
.app-header {
    background-color: #1B2636;
    background-image:
        repeating-linear-gradient(0deg, rgba(56,189,248,0.08) 0px, rgba(56,189,248,0.08) 1px, transparent 1px, transparent 28px),
        repeating-linear-gradient(90deg, rgba(56,189,248,0.08) 0px, rgba(56,189,248,0.08) 1px, transparent 1px, transparent 28px);
    border: 1px solid var(--border);
    border-left: 4px solid var(--blue);
    border-radius: 10px;
    padding: 1.3rem 1.5rem;
    margin-bottom: 1rem;
}
.app-title {
    font-size: 1.75rem;
    font-weight: 700;
    color: var(--text);
    margin: 0 0 0.2rem 0;
}
.app-subtitle {
    color: var(--muted);
    font-size: 0.95rem;
    margin: 0;
}

/* ---------- Top stat tiles ---------- */
.stat-tile {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 0.8rem 1rem;
}
.stat-label {
    font-size: 0.78rem;
    color: var(--muted);
    margin-bottom: 0.2rem;
}
.stat-value {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.15rem;
    font-weight: 600;
    color: var(--text);
}
.tone-good .stat-value { color: var(--green); }
.tone-bad .stat-value { color: var(--red); }

/* ---------- Native widgets ---------- */
[data-testid="stMetric"] {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 0.7rem 1rem;
}
.stTabs [data-baseweb="tab-list"] {
    gap: 0.25rem;
    border-bottom: 1px solid var(--border);
}
.stTabs [data-baseweb="tab"] {
    font-weight: 500;
    color: var(--muted);
}
.stTabs [aria-selected="true"] {
    color: var(--blue) !important;
}
.stTabs [data-baseweb="tab-highlight"] {
    background-color: var(--blue) !important;
}
.stButton > button, .stDownloadButton > button {
    border-radius: 5px;
    font-weight: 500;
}

/* ---------- Section headings ---------- */
.section-title {
    font-size: 1.25rem;
    font-weight: 600;
    color: var(--text);
    margin: 1rem 0 0.15rem 0;
}
.section-sub {
    color: var(--muted);
    font-size: 0.92rem;
    margin-bottom: 0.9rem;
}
.panel-label {
    font-weight: 600;
    font-size: 0.95rem;
    color: var(--text);
    margin: 0.4rem 0 0.4rem 0;
}

/* ---------- Result banner ---------- */
.result-strip {
    background: var(--surface);
    border: 1px solid var(--border);
    border-left-width: 4px;
    border-radius: 10px;
    padding: 0.75rem 1rem;
    margin: 0.8rem 0 1rem 0;
}
.result-strip .title { font-weight: 600; font-size: 1.02rem; }
.result-strip .sub { color: var(--muted); font-size: 0.88rem; margin-top: 0.15rem; }
.strip-clear { border-left-color: var(--green); }
.strip-found { border-left-color: var(--amber); }

/* ---------- Per-class finding cards ---------- */
.class-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-left: 4px solid var(--c, #5F6C7A);
    border-radius: 10px;
    padding: 0.75rem 1rem;
    margin-bottom: 0.55rem;
}
.class-card .row {
    display: flex;
    justify-content: space-between;
    align-items: baseline;
}
.class-card .name {
    font-weight: 600;
    font-size: 1.02rem;
    text-transform: capitalize;
}
.class-card .conf {
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    color: var(--c, #5F6C7A);
}
.class-card .meta {
    color: var(--muted);
    font-size: 0.82rem;
    margin-top: 0.3rem;
}
.meter {
    height: 6px;
    background: #374151;
    border-radius: 3px;
    overflow: hidden;
    margin-top: 0.5rem;
}
.meter > span {
    display: block;
    height: 100%;
    background: var(--c, #5F6C7A);
}

/* ---------- Empty states ---------- */
.empty-state {
    border: 1px dashed #4B5563;
    border-radius: 10px;
    background: #FBFCFD;
    padding: 2rem 1rem;
    text-align: center;
    color: var(--muted);
}
.empty-state .big { font-size: 1.8rem; margin-bottom: 0.3rem; }
.empty-state .head { font-weight: 600; color: var(--text); margin-bottom: 0.2rem; }

/* ---------- Sidebar legend ---------- */
.legend-row {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    padding: 0.18rem 0;
    font-size: 0.92rem;
    text-transform: capitalize;
}
.legend-dot {
    width: 10px;
    height: 10px;
    border-radius: 50%;
    display: inline-block;
}

/* ---------- Footer ---------- */
.custom-footer {
    text-align: center;
    color: var(--muted);
    font-size: 0.8rem;
    padding: 1.6rem 0 0.6rem 0;
}

    /* Dark dashboard polish */
    [data-testid="stSidebar"] {
        background: #172131;
        border-right: 1px solid #374151;
    }

    [data-testid="stFileUploader"] section {
        background: #182231;
        border: 1px dashed #4B5563;
        border-radius: 10px;
    }

    .stButton > button {
        background: #2563EB;
        color: white;
        border: 0;
        min-height: 2.7rem;
        transition: 0.2s ease;
    }

    .stButton > button:hover {
        background: #1D4ED8;
        color: white;
        border: 0;
    }

    .stDownloadButton > button {
        background: #263548;
        color: #E5E7EB;
        border: 1px solid #475569;
    }

    .stDataFrame, [data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
    }

    [data-testid="stMetricValue"] {
        color: #F3F4F6;
    }

    [data-testid="stMetricLabel"] {
        color: #9CA3AF;
    }

</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# SESSION STATE
# ============================================================

if "history" not in st.session_state:
    st.session_state.history = []      # inspection log for this session
if "results" not in st.session_state:
    st.session_state.results = {}      # latest result per tab ("image", "camera")
if "signatures" not in st.session_state:
    st.session_state.signatures = {}   # remembers which file each tab last saw


# ============================================================
# SMALL HELPERS
# ============================================================

def plural(n, word):
    return f"{n} {word}" if n == 1 else f"{n} {word}s"


def file_signature(uploaded_file):
    return hashlib.md5(uploaded_file.getvalue()).hexdigest()


def to_png_bytes(rgb_array):
    buf = io.BytesIO()
    Image.fromarray(rgb_array).save(buf, format="PNG")
    return buf.getvalue()


def format_clock(seconds):
    seconds = int(seconds)
    return f"{seconds // 60}:{seconds % 60:02d}"


def section_header(title, subtitle):
    st.markdown(
        f'<div class="section-title">{title}</div>'
        f'<div class="section-sub">{subtitle}</div>',
        unsafe_allow_html=True,
    )


def empty_state(icon, head, body):
    st.markdown(
        f"""
        <div class="empty-state">
            <div class="big">{icon}</div>
            <div class="head">{head}</div>
            <div>{body}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def result_strip(kind, title, sub):
    st.markdown(
        f"""
        <div class="result-strip strip-{kind}">
            <div class="title">{title}</div>
            <div class="sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def reset_if_file_changed(tab_key, uploaded_file, result_keys):
    """Forget old results when the person picks a different file."""
    sig = file_signature(uploaded_file)
    if st.session_state.signatures.get(tab_key) != sig:
        st.session_state.signatures[tab_key] = sig
        for k in result_keys:
            st.session_state.results.pop(k, None)


def log_inspection(source, total, types, top_conf):
    st.session_state.history.append(
        {
            "Time": datetime.now().strftime("%d %b %Y, %H:%M:%S"),
            "Source": source,
            "Defects found": total,
            "Defect types": ", ".join(sorted(types)) if types else "None",
            "Highest confidence": f"{top_conf:.1%}" if total else "-",
        }
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown("### Detection sensitivity")

    conf_threshold = st.slider(
        "Confidence threshold",
        0.05,
        0.95,
        0.25,
        0.05,
        help="Lower finds more possible defects but with more false alarms. "
             "Higher is stricter and shows only what the model is sure about.",
    )
    st.caption("Lower catches more, higher is stricter.")

    st.divider()

    st.markdown("### Detectable defects")

    legend_html = "".join(
        f'<div class="legend-row"><span class="legend-dot" '
        f'style="background:{color};"></span>{name}</div>'
        for name, color in CLASS_COLORS.items()
    )
    st.markdown(legend_html, unsafe_allow_html=True)

    st.divider()

    with st.expander("How to use"):
        st.markdown(
            """
1. Pick an input: image, camera, or video.
2. Adjust the confidence threshold if needed.
3. Click the inspect button.
4. Review the findings and download the results.
            """
        )

    with st.expander("Tips for better results"):
        st.markdown(
            """
- Use clear, well-lit photos of the wall, ceiling, or facade.
- Get close enough that the defect is visible, but keep some surrounding surface in frame.
- Avoid heavy shadows and motion blur.
- The model was trained on building-wall imagery, so photos of other scenes may give unreliable results.
            """
        )

    with st.expander("Advanced settings"):
        model_path = st.text_input(
            "Model weights file",
            value="best.pt",
            help="Path to the trained YOLO model (best.pt)",
        )
        frame_skip = st.slider(
            "Video frame skip",
            1,
            15,
            5,
            help="Analyse every Nth frame. Lower is more thorough, higher is faster.",
        )

    st.caption("AI-Based Safety Inspection Assistance System")
    st.caption("Department of AI & DS")


# ============================================================
# LOAD MODEL
# ============================================================

@st.cache_resource
def load_model(path):
    return YOLO(path)


model = None
model_error = None
try:
    model = load_model(model_path)
except Exception as e:
    model_error = str(e)

model_ok = model is not None


# ============================================================
# HEADER + STATUS TILES
# ============================================================

st.markdown(
    """
<div class="app-header">
    <div class="app-title">🏗️ BuildingSafe AI</div>
    <p class="app-subtitle">
        AI-powered defect inspection for construction buildings.
        Find cracks, leakage, corrosion, abscission, and bulges from a photo, camera, or video.
    </p>
</div>
""",
    unsafe_allow_html=True,
)


def stat_tile(label, value, tone="neutral"):
    return (
        f'<div class="stat-tile tone-{tone}">'
        f'<div class="stat-label">{label}</div>'
        f'<div class="stat-value">{value}</div>'
        f"</div>"
    )


t1, t2, t3, t4 = st.columns(4)
t1.markdown(
    stat_tile(
        "AI model",
        "Ready" if model_ok else "Not loaded",
        "good" if model_ok else "bad",
    ),
    unsafe_allow_html=True,
)
t2.markdown(stat_tile("Confidence threshold", f"{conf_threshold:.0%}"), unsafe_allow_html=True)
t3.markdown(
    stat_tile("Defect classes", len(model.names) if model_ok else len(CLASS_COLORS)),
    unsafe_allow_html=True,
)
t4.markdown(stat_tile("Detection engine", "YOLOv8"), unsafe_allow_html=True)

if not model_ok:
    st.error(
        f"Could not load the model from `{model_path}`. "
        "Put best.pt in your project folder, or open **Advanced settings** "
        "in the sidebar and enter the correct path."
    )
    with st.expander("Technical details"):
        st.code(model_error)

st.write("")


# ============================================================
# DETECTION ENGINE
# ============================================================

def analyze(image, source_label):
    """Run the model on one image and return everything the UI needs."""
    result = model.predict(image, conf=conf_threshold, verbose=False)[0]

    rows = []
    for box in result.boxes:
        x1, y1, x2, y2 = [int(v) for v in box.xyxy[0].tolist()]
        rows.append(
            {
                "defect": model.names[int(box.cls[0])],
                "confidence": float(box.conf[0]),
                "box": (x1, y1, x2, y2),
            }
        )

    return {
        "source": source_label,
        "time": datetime.now().strftime("%H:%M:%S"),
        "threshold": conf_threshold,
        "rows": rows,
        "original": cv2.cvtColor(result.orig_img, cv2.COLOR_BGR2RGB),
        "annotated": cv2.cvtColor(result.plot(), cv2.COLOR_BGR2RGB),
        "ms": sum(result.speed.values()),
    }


def group_by_defect(rows):
    """Collapse individual boxes into one summary per defect type."""
    groups = {}
    for r in rows:
        g = groups.setdefault(r["defect"], [])
        g.append(r["confidence"])
    return sorted(groups.items(), key=lambda kv: max(kv[1]), reverse=True)


# ============================================================
# RESULT RENDERING
# ============================================================

def render_result(res, key):
    rows = res["rows"]

    left, right = st.columns(2)
    with left:
        st.markdown('<div class="panel-label">Input</div>', unsafe_allow_html=True)
        st.image(res["original"], use_container_width=True)
    with right:
        st.markdown('<div class="panel-label">AI inspection</div>', unsafe_allow_html=True)
        st.image(res["annotated"], use_container_width=True)

    if not rows:
        result_strip(
            "clear",
            "No defects detected",
            f"Nothing found above the {res['threshold']:.0%} confidence threshold. "
            "If you expect a defect here, lower the threshold in the sidebar and inspect again.",
        )
        return

    groups = group_by_defect(rows)

    result_strip(
        "found",
        f"{plural(len(rows), 'defect')} flagged across {plural(len(groups), 'type')}",
        "Review the highlighted areas. Confidence shows how sure the model is, "
        "not how severe the defect is.",
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Defects found", len(rows))
    m2.metric("Highest confidence", f"{max(r['confidence'] for r in rows):.1%}")
    m3.metric("Defect types", len(groups))
    m4.metric("Analysis time", f"{res['ms']:.0f} ms")

    st.markdown('<div class="section-title">Findings by defect type</div>', unsafe_allow_html=True)

    for name, confs in groups:
        color = CLASS_COLORS.get(name, DEFAULT_COLOR)
        top = max(confs)
        avg = sum(confs) / len(confs)
        st.markdown(
            f"""
            <div class="class-card" style="--c:{color};">
                <div class="row">
                    <span class="name">{name}</span>
                    <span class="conf">{top:.1%}</span>
                </div>
                <div class="meter"><span style="width:{top * 100:.1f}%;"></span></div>
                <div class="meta">
                    {plural(len(confs), 'detection')}, average confidence {avg:.0%}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    table = pd.DataFrame(
        [
            {
                "Defect": r["defect"].title(),
                "Confidence (%)": round(r["confidence"] * 100, 1),
                "x1": r["box"][0],
                "y1": r["box"][1],
                "x2": r["box"][2],
                "y2": r["box"][3],
            }
            for r in sorted(rows, key=lambda r: r["confidence"], reverse=True)
        ]
    )

    with st.expander("All individual detections"):
        st.dataframe(table, use_container_width=True, hide_index=True)

    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ Download annotated image",
        data=to_png_bytes(res["annotated"]),
        file_name=f"inspection_{key}.png",
        mime="image/png",
        key=f"dl_img_{key}",
        use_container_width=True,
    )
    d2.download_button(
        "⬇️ Download findings (CSV)",
        data=table.to_csv(index=False).encode("utf-8"),
        file_name=f"findings_{key}.csv",
        mime="text/csv",
        key=f"dl_csv_{key}",
        use_container_width=True,
    )

    st.caption(f"Analysed at {res['time']} with a {res['threshold']:.0%} confidence threshold.")


def run_and_store(key, image, source_label):
    with st.spinner("AI is analysing the image..."):
        res = analyze(image, source_label)
    st.session_state.results[key] = res
    rows = res["rows"]
    log_inspection(
        source_label,
        len(rows),
        {r["defect"] for r in rows},
        max((r["confidence"] for r in rows), default=0.0),
    )


# ============================================================
# MAIN TABS
# ============================================================

tab_image, tab_camera, tab_video, tab_log = st.tabs(
    ["📷 Image", "📸 Camera", "🎥 Video", "🗂️ Inspection log"]
)


# ------------------------------------------------------------
# IMAGE
# ------------------------------------------------------------

with tab_image:

    section_header(
        "Image inspection",
        "Upload a clear photo of a wall, ceiling, beam, or facade.",
    )

    uploaded = st.file_uploader(
        "Choose a building image",
        type=["jpg", "jpeg", "png"],
        key="img_upload",
    )

    if uploaded is None:
        st.session_state.signatures.pop("image", None)
        st.session_state.results.pop("image", None)
        empty_state(
            "📷",
            "No image selected",
            "Drag a photo into the box above, or click Browse files.",
        )

    elif model_ok:
        reset_if_file_changed("image", uploaded, ["image"])
        image = Image.open(uploaded).convert("RGB")

        if st.button(
            "🔍 Inspect building",
            type="primary",
            use_container_width=True,
            key="inspect_img",
        ):
            run_and_store("image", image, f"Image: {uploaded.name}")

        res = st.session_state.results.get("image")

        if res:
            render_result(res, "image")
        else:
            left, right = st.columns(2)
            with left:
                st.markdown('<div class="panel-label">Input</div>', unsafe_allow_html=True)
                st.image(image, use_container_width=True)
            with right:
                st.markdown('<div class="panel-label">AI inspection</div>', unsafe_allow_html=True)
                empty_state(
                    "🔍",
                    "Ready to inspect",
                    "Click Inspect building and the results will appear here.",
                )


# ------------------------------------------------------------
# CAMERA
# ------------------------------------------------------------

with tab_camera:

    section_header(
        "Camera inspection",
        "Take a photo of the surface with your device camera and inspect it instantly.",
    )

    cam_image = st.camera_input(
        "Capture building image",
        key="cam_input",
    )

    if cam_image is None:
        st.session_state.signatures.pop("camera", None)
        st.session_state.results.pop("camera", None)

    elif model_ok:
        reset_if_file_changed("camera", cam_image, ["camera"])
        photo = Image.open(cam_image).convert("RGB")

        if st.button(
            "🔍 Inspect camera photo",
            type="primary",
            use_container_width=True,
            key="inspect_cam",
        ):
            run_and_store("camera", photo, "Camera photo")

        res = st.session_state.results.get("camera")
        if res:
            render_result(res, "camera")

    st.caption(
        "Camera mode inspects one photo at a time. "
        "Continuous live detection needs a real-time video component."
    )


# ------------------------------------------------------------
# VIDEO
# ------------------------------------------------------------

def render_video_summary(s):
    if not s["tally"]:
        result_strip(
            "clear",
            "No defects detected",
            f"Checked {plural(s['processed'], 'frame')} at a {s['threshold']:.0%} "
            "confidence threshold and found nothing. Try lowering the threshold.",
        )
        return

    result_strip(
        "found",
        f"Defects found in {s['flagged']} of {s['processed']} analysed frames",
        "Confidence shows how sure the model is, not how severe the defect is.",
    )

    v1, v2, v3, v4 = st.columns(4)
    v1.metric("Frames analysed", s["processed"])
    v2.metric("Frames with defects", s["flagged"])
    v3.metric("Total detections", sum(s["tally"].values()))
    v4.metric("Defect types", len(s["tally"]))

    chart_col, frame_col = st.columns(2)

    with chart_col:
        st.markdown('<div class="panel-label">Detections by defect type</div>', unsafe_allow_html=True)
        summary_df = pd.DataFrame(
            [{"Defect": k.title(), "Detections": v} for k, v in s["tally"].items()]
        )
        st.bar_chart(summary_df.set_index("Defect"))
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

    with frame_col:
        best = s["best"]
        when = f" at {format_clock(best['frame'] / s['fps'])}" if s["fps"] else ""
        st.markdown(
            f'<div class="panel-label">Strongest detection{when} '
            f'(frame {best["frame"]}, {best["conf"]:.1%})</div>',
            unsafe_allow_html=True,
        )
        st.image(best["image"], use_container_width=True)
        st.download_button(
            "⬇️ Download this frame",
            data=to_png_bytes(best["image"]),
            file_name="strongest_detection.png",
            mime="image/png",
            key="dl_video_frame",
            use_container_width=True,
        )


with tab_video:

    section_header(
        "Video inspection",
        "Upload an inspection video. The AI checks every few frames and summarises what it finds.",
    )

    video_file = st.file_uploader(
        "Choose a building video",
        type=["mp4", "mov", "avi"],
        key="video_upload",
    )

    if video_file is None:
        st.session_state.signatures.pop("video", None)
        st.session_state.results.pop("video", None)
        empty_state(
            "🎥",
            "No video selected",
            "Upload an mp4, mov, or avi file to scan it frame by frame.",
        )

    elif model_ok:
        reset_if_file_changed("video", video_file, ["video"])

        if st.button(
            "🔍 Start video inspection",
            type="primary",
            use_container_width=True,
            key="inspect_video",
        ):
            tfile = tempfile.NamedTemporaryFile(delete=False, suffix=".mp4")
            tfile.write(video_file.getvalue())
            tfile.close()

            cap = cv2.VideoCapture(tfile.name)
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT)) or 1
            fps = cap.get(cv2.CAP_PROP_FPS) or 0

            frame_slot = st.empty()
            progress = st.progress(0, text="Starting...")

            frame_idx = 0
            processed = 0
            flagged = 0
            tally = {}
            best = {"conf": 0.0, "image": None, "frame": 0}

            while cap.isOpened():
                ok, frame = cap.read()
                if not ok:
                    break
                frame_idx += 1

                if frame_idx % frame_skip == 0:
                    result = model.predict(frame, conf=conf_threshold, verbose=False)[0]
                    processed += 1
                    annotated = result.plot()
                    frame_slot.image(annotated, channels="BGR", use_container_width=True)

                    if len(result.boxes) > 0:
                        flagged += 1

                    for box in result.boxes:
                        name = model.names[int(box.cls[0])]
                        conf = float(box.conf[0])
                        tally[name] = tally.get(name, 0) + 1
                        if conf > best["conf"]:
                            best = {
                                "conf": conf,
                                "image": cv2.cvtColor(annotated, cv2.COLOR_BGR2RGB),
                                "frame": frame_idx,
                            }

                progress.progress(
                    min(frame_idx / total_frames, 1.0),
                    text=f"Analysing frame {frame_idx} of {total_frames}",
                )

            cap.release()
            try:
                os.unlink(tfile.name)
            except OSError:
                pass

            frame_slot.empty()
            progress.empty()

            st.session_state.results["video"] = {
                "processed": processed,
                "flagged": flagged,
                "tally": tally,
                "best": best,
                "fps": fps,
                "threshold": conf_threshold,
            }

            log_inspection(
                f"Video: {video_file.name}",
                sum(tally.values()),
                set(tally.keys()),
                best["conf"],
            )

        summary = st.session_state.results.get("video")
        if summary:
            render_video_summary(summary)


# ------------------------------------------------------------
# INSPECTION LOG
# ------------------------------------------------------------

with tab_log:

    section_header(
        "Inspection log",
        "Every inspection you run in this session is listed here, newest first.",
    )

    history = st.session_state.history

    if not history:
        empty_state(
            "🗂️",
            "Nothing inspected yet",
            "Run an image, camera, or video inspection and it will show up here.",
        )
    else:
        log_df = pd.DataFrame(history[::-1])
        st.dataframe(log_df, use_container_width=True, hide_index=True)

        c1, c2 = st.columns(2)
        c1.download_button(
            "⬇️ Download log (CSV)",
            data=log_df.to_csv(index=False).encode("utf-8"),
            file_name="inspection_log.csv",
            mime="text/csv",
            key="dl_log",
            use_container_width=True,
        )
        if c2.button("Clear log", key="clear_log", use_container_width=True):
            st.session_state.history = []
            st.rerun()


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
<div class="custom-footer">
    🏗️ BuildingSafe AI &nbsp;|&nbsp; AI-Based Safety Inspection Assistance System
    <br>
    Deep Learning • Computer Vision • YOLOv8
    <br>
    A prototype for visual screening. It does not replace a certified structural inspector.
</div>
""",
    unsafe_allow_html=True,
)
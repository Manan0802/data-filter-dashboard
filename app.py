import streamlit as st
import pandas as pd
import io
import time
import math

# ─────────────────────────────────────────────
#  Page configuration
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Data Filter & Extraction Dashboard",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
#  Custom CSS – clean modern dark-themed UI
# ─────────────────────────────────────────────
st.markdown(
    """
    <style>
        /* ── Google Font ── */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }

        /* ── App background ── */
        .stApp {
            background: linear-gradient(135deg, #0f0f1a 0%, #1a1a2e 50%, #16213e 100%);
            min-height: 100vh;
        }

        /* ── Main container card ── */
        .main-card {
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 20px;
            padding: 2rem 2.5rem;
            margin-bottom: 1.5rem;
            backdrop-filter: blur(10px);
        }

        /* ── Section headers ── */
        .section-header {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 1rem;
        }

        .step-badge {
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            color: white;
            border-radius: 50%;
            width: 32px;
            height: 32px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            font-size: 0.85rem;
            flex-shrink: 0;
        }

        /* ── Hero header ── */
        .hero {
            text-align: center;
            padding: 2.5rem 1rem 1.5rem;
        }

        .hero h1 {
            font-size: 2.6rem;
            font-weight: 700;
            background: linear-gradient(135deg, #6366f1, #a78bfa, #38bdf8);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            margin-bottom: 0.4rem;
        }

        .hero p {
            color: rgba(255,255,255,0.5);
            font-size: 1.05rem;
        }

        /* ── Metric cards ── */
        .metric-row {
            display: flex;
            gap: 1rem;
            margin-top: 0.5rem;
        }

        .metric-card {
            flex: 1;
            background: rgba(99, 102, 241, 0.1);
            border: 1px solid rgba(99, 102, 241, 0.25);
            border-radius: 14px;
            padding: 1.2rem 1.5rem;
            text-align: center;
        }

        .metric-card .metric-value {
            font-size: 2rem;
            font-weight: 700;
            color: #a78bfa;
        }

        .metric-card .metric-label {
            font-size: 0.8rem;
            color: rgba(255,255,255,0.45);
            margin-top: 4px;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        /* ── Divider ── */
        .fancy-divider {
            height: 1px;
            background: linear-gradient(90deg, transparent, rgba(99,102,241,0.5), transparent);
            margin: 1.5rem 0;
        }

        /* ── Button overrides ── */
        .stButton > button {
            background: linear-gradient(135deg, #6366f1, #8b5cf6) !important;
            color: white !important;
            border: none !important;
            border-radius: 10px !important;
            padding: 0.6rem 2rem !important;
            font-weight: 600 !important;
            font-size: 1rem !important;
            transition: all 0.2s ease !important;
            width: 100% !important;
        }

        .stButton > button:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 8px 24px rgba(99, 102, 241, 0.45) !important;
        }

        /* ── Download button ── */
        .stDownloadButton > button {
            background: linear-gradient(135deg, #059669, #10b981) !important;
            color: white !important;
            border: none !important;
            border-radius: 10px !important;
            padding: 0.6rem 2rem !important;
            font-weight: 600 !important;
            width: 100% !important;
            margin-top: 0.5rem !important;
        }

        .stDownloadButton > button:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 8px 24px rgba(16, 185, 129, 0.45) !important;
        }

        /* ── Streamlit widget labels ── */
        label, .stSelectbox label, .stMultiSelect label, .stTextArea label {
            color: rgba(255,255,255,0.75) !important;
            font-weight: 500 !important;
            font-size: 0.9rem !important;
        }

        /* ── File uploader ── */
        .stFileUploader {
            border: 2px dashed rgba(99, 102, 241, 0.4) !important;
            border-radius: 14px !important;
            padding: 1rem !important;
        }

        /* ── Success / warning / error ── */
        .stSuccess, .stWarning, .stError, .stInfo {
            border-radius: 10px !important;
        }

        /* ── Scrollable dataframe ── */
        .stDataFrame {
            border-radius: 12px !important;
            overflow: hidden !important;
        }

        /* ── Footer ── */
        .footer {
            text-align: center;
            color: rgba(255,255,255,0.25);
            font-size: 0.78rem;
            padding: 1rem 0 0.5rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────
#  Hero Header
# ─────────────────────────────────────────────
st.markdown(
    """
    <div class="hero">
        <h1>🔍 Data Filter & Extraction Dashboard</h1>
        <p>Upload · Filter · Extract · Download — built for large datasets</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────
#  Helper utilities
# ─────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_file(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    """
    Load a CSV or XLSX file into a DataFrame.
    ⚡ Uses PyArrow engine for CSV — 3-5x faster than default pandas engine.
    Cached so the file is read only once per unique upload.
    """
    if file_name.endswith(".csv"):
        try:
            # PyArrow engine: fastest option, skips Python-level type inference
            return pd.read_csv(
                io.BytesIO(file_bytes),
                engine="pyarrow",
            )
        except Exception:
            # Fallback to default engine if pyarrow hits an edge case
            return pd.read_csv(io.BytesIO(file_bytes), low_memory=False)
    elif file_name.endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
    else:
        raise ValueError("Unsupported file format.")


def human_size(num_bytes: int) -> str:
    """Convert bytes to human-readable file size."""
    if num_bytes == 0:
        return "0 B"
    unit = ["B", "KB", "MB", "GB"]
    i = int(math.floor(math.log(num_bytes, 1024)))
    i = min(i, len(unit) - 1)
    p = math.pow(1024, i)
    return f"{num_bytes / p:.1f} {unit[i]}"


def parse_values(raw_text: str) -> list:
    """
    Accept both newline-separated and comma-separated input.
    Strip whitespace, drop empty strings, deduplicate while preserving order.
    """
    import re
    tokens = re.split(r"[\n,]+", raw_text)
    seen = set()
    result = []
    for t in tokens:
        cleaned = t.strip()
        if cleaned and cleaned not in seen:
            seen.add(cleaned)
            result.append(cleaned)
    return result


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


# ─────────────────────────────────────────────
#  SESSION STATE INIT
# ─────────────────────────────────────────────
for key in ("df", "result_df", "file_name"):
    if key not in st.session_state:
        st.session_state[key] = None


# ─────────────────────────────────────────────
#  STEP 0 ── File Upload
# ─────────────────────────────────────────────
st.markdown('<div class="main-card">', unsafe_allow_html=True)
st.markdown(
    '<div class="section-header">'
    '<div class="step-badge">📁</div>'
    '<h3 style="color:white;margin:0;">Upload your dataset</h3>'
    "</div>",
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader(
    "Drag & drop a CSV or Excel file here, or click to browse",
    type=["csv", "xlsx", "xls"],
    label_visibility="collapsed",
)

if uploaded_file is not None:
    # Only reload if a new file is uploaded
    if st.session_state.get("file_name") != uploaded_file.name:
        st.session_state.result_df = None  # Clear previous results
        file_bytes = uploaded_file.read()
        file_size = len(file_bytes)
        with st.spinner(f"⚡ Loading {human_size(file_size)} file with PyArrow engine …"):
            try:
                t0 = time.perf_counter()
                df = load_file(file_bytes, uploaded_file.name)
                load_time = time.perf_counter() - t0
                st.session_state.df = df
                st.session_state.file_name = uploaded_file.name
                st.session_state.load_time = load_time
                st.session_state.file_size = file_size
            except Exception as e:
                st.error(f"❌ Could not read the file: {e}")
                st.session_state.df = None

    if st.session_state.df is not None:
        df = st.session_state.df
        load_time = st.session_state.get("load_time", 0)
        file_size = st.session_state.get("file_size", 0)
        col1, col2, col3, col4 = st.columns(4)
        col1.success("✅ File loaded!")
        with col2:
            st.markdown(
                f'<div style="color:rgba(255,255,255,0.6);font-size:0.82rem;padding-top:0.7rem;">'
                f'📄 <b style="color:white;">{uploaded_file.name}</b></div>',
                unsafe_allow_html=True,
            )
        with col3:
            st.markdown(
                f'<div style="color:rgba(255,255,255,0.6);font-size:0.82rem;padding-top:0.7rem;">'
                f'📊 <b style="color:white;">{df.shape[0]:,}</b> rows × '
                f'<b style="color:white;">{df.shape[1]}</b> cols &nbsp;'
                f'<span style="color:rgba(255,255,255,0.35)">({human_size(file_size)})</span></div>',
                unsafe_allow_html=True,
            )
        with col4:
            speed_color = "#10b981" if load_time < 2 else ("#f59e0b" if load_time < 6 else "#ef4444")
            st.markdown(
                f'<div style="color:rgba(255,255,255,0.6);font-size:0.82rem;padding-top:0.7rem;">'
                f'⚡ Loaded in <b style="color:{speed_color};">{load_time:.2f}s</b></div>',
                unsafe_allow_html=True,
            )

st.markdown("</div>", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  STEPS 1‑3 ── Configuration (only if file is loaded)
# ─────────────────────────────────────────────
if st.session_state.df is not None:
    df = st.session_state.df
    all_columns = df.columns.tolist()

    st.markdown('<div class="main-card">', unsafe_allow_html=True)
    st.markdown(
        '<h3 style="color:white;margin-top:0;margin-bottom:1.2rem;">⚙️ Configure Your Filter</h3>',
        unsafe_allow_html=True,
    )

    # ── Step 1: Output Columns ──
    st.markdown(
        '<div class="section-header">'
        '<div class="step-badge">1</div>'
        '<span style="color:rgba(255,255,255,0.85);font-weight:500;">'
        "Select the columns you want in your final output"
        "</span></div>",
        unsafe_allow_html=True,
    )
    selected_columns = st.multiselect(
        "output_columns",
        options=all_columns,
        default=all_columns,
        placeholder="Choose one or more columns …",
        label_visibility="collapsed",
    )

    st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

    # ── Step 2: Filter Column ──
    st.markdown(
        '<div class="section-header">'
        '<div class="step-badge">2</div>'
        '<span style="color:rgba(255,255,255,0.85);font-weight:500;">'
        "Select the column to apply your filter condition on&nbsp;<span style='color:#a78bfa;'>(e.g., mcat_id)</span>"
        "</span></div>",
        unsafe_allow_html=True,
    )
    filter_column = st.selectbox(
        "filter_column",
        options=all_columns,
        label_visibility="collapsed",
    )

    st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

    # ── Step 3: Values ──
    st.markdown(
        '<div class="section-header">'
        '<div class="step-badge">3</div>'
        '<span style="color:rgba(255,255,255,0.85);font-weight:500;">'
        "Paste your list of values (comma-separated or line-by-line)"
        "</span></div>",
        unsafe_allow_html=True,
    )
    raw_values = st.text_area(
        "values_input",
        height=180,
        placeholder="e.g.\n1001\n1002\n1003\nor: 1001, 1002, 1003",
        label_visibility="collapsed",
    )

    # Show live parse preview
    if raw_values.strip():
        preview_list = parse_values(raw_values)
        st.caption(
            f"🔎 Detected **{len(preview_list):,}** unique value(s) to match against "
            f"**`{filter_column}`**"
        )

    st.markdown("</div>", unsafe_allow_html=True)

    # ─────────────────────────────────────────
    #  Process Button
    # ─────────────────────────────────────────
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        process_clicked = st.button("⚡ Process Data")

    if process_clicked:
        # ── Validation ──
        errors = []
        if not selected_columns:
            errors.append("Please select **at least one output column** (Step 1).")
        if not raw_values.strip():
            errors.append("Please paste **at least one filter value** (Step 3).")

        if errors:
            for err in errors:
                st.warning(f"⚠️ {err}")
        else:
            value_list = parse_values(raw_values)

            with st.spinner(f"Filtering {df.shape[0]:,} rows …"):
                start_time = time.perf_counter()

                try:
                    # ── Core filtering logic ──
                    # Convert filter column to string for safe isin comparison
                    str_series = df[filter_column].astype(str).str.strip()

                    # Also try numeric matching if the column is numeric
                    # Build a set of string versions of the user values (safe for all dtypes)
                    value_set = set(value_list)

                    mask = str_series.isin(value_set)
                    result_df = df.loc[mask, selected_columns].copy()

                    elapsed = time.perf_counter() - start_time
                    st.session_state.result_df = result_df
                    st.session_state.elapsed = elapsed

                except KeyError as ke:
                    st.error(f"❌ Column not found in the dataset: {ke}")
                    st.session_state.result_df = None
                except Exception as ex:
                    st.error(f"❌ An unexpected error occurred during filtering: {ex}")
                    st.session_state.result_df = None

    # ─────────────────────────────────────────
    #  Results Section
    # ─────────────────────────────────────────
    if st.session_state.result_df is not None:
        result_df: pd.DataFrame = st.session_state.result_df
        elapsed = st.session_state.get("elapsed", 0)

        st.markdown('<div class="main-card">', unsafe_allow_html=True)
        st.markdown(
            '<h3 style="color:white;margin-top:0;">📊 Results</h3>',
            unsafe_allow_html=True,
        )

        # ── Metrics ──
        orig_count = df.shape[0]
        result_count = result_df.shape[0]
        match_pct = (result_count / orig_count * 100) if orig_count > 0 else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Original Rows", f"{orig_count:,}")
        m2.metric("Matched Rows", f"{result_count:,}", delta=f"{result_count - orig_count:,}")
        m3.metric("Match Rate", f"{match_pct:.1f}%")
        m4.metric("Processing Time", f"{elapsed:.3f}s")

        st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

        if result_count == 0:
            st.info(
                "ℹ️ No rows matched your filter values. "
                "Check that the values match the format in the selected column (e.g., correct data type, no extra spaces)."
            )
        else:
            # ── Preview ──
            preview_rows = min(50, result_count)
            st.markdown(
                f'<p style="color:rgba(255,255,255,0.55);font-size:0.85rem;margin-bottom:0.4rem;">'
                f"Showing first <b style='color:white;'>{preview_rows}</b> of "
                f"<b style='color:white;'>{result_count:,}</b> rows &nbsp;·&nbsp; "
                f"<b style='color:white;'>{len(result_df.columns)}</b> column(s)"
                f"</p>",
                unsafe_allow_html=True,
            )
            st.dataframe(
                result_df.head(preview_rows),
                use_container_width=True,
                height=420,
            )

            st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

            # ── Download ──
            csv_bytes = to_csv_bytes(result_df)
            base_name = st.session_state.file_name.rsplit(".", 1)[0]
            download_name = f"{base_name}_filtered_{result_count}_rows.csv"

            dl_col, info_col = st.columns([1, 2])
            with dl_col:
                st.download_button(
                    label="⬇️ Download Full Results as CSV",
                    data=csv_bytes,
                    file_name=download_name,
                    mime="text/csv",
                )
            with info_col:
                st.markdown(
                    f'<div style="padding-top:0.8rem;color:rgba(255,255,255,0.45);font-size:0.82rem;">'
                    f"📦 Estimated file size: <b style='color:white;'>{len(csv_bytes)/1024:.1f} KB</b> &nbsp;·&nbsp; "
                    f"All <b style='color:white;'>{result_count:,}</b> matching rows included"
                    f"</div>",
                    unsafe_allow_html=True,
                )

        st.markdown("</div>", unsafe_allow_html=True)

else:
    # ── Placeholder when no file is loaded ──
    st.markdown(
        """
        <div style="
            text-align:center;
            padding: 3rem 1rem;
            color: rgba(255,255,255,0.25);
        ">
            <div style="font-size:4rem;margin-bottom:1rem;">📂</div>
            <p style="font-size:1.1rem;">Upload a CSV or Excel file above to get started.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────
#  Footer
# ─────────────────────────────────────────────
st.markdown(
    '<div class="footer">Data Filter & Extraction Dashboard · Built with Streamlit & Pandas</div>',
    unsafe_allow_html=True,
)

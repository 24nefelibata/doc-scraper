# =============================================================================
# app.py — Documentation Scraper & Word Document Generator
# VERSION 4: Fixed function ordering + images skipped for stability
# =============================================================================

import streamlit as st
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import re
import urllib3
import urllib.parse

# =============================================================================
# SECTION 1: PAGE CONFIG
# =============================================================================

st.set_page_config(
    page_title="Doc Scraper → Word / TXT",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
    <style>
        .main-title { font-size: 2.2rem; font-weight: 700; color: #1a1a2e; }
        .sub-title  { font-size: 1rem; color: #555; margin-top: -10px; }
        .stDownloadButton > button { font-weight: 600; border-radius: 8px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">📄 Documentation Scraper → Word / TXT</p>',
            unsafe_allow_html=True)
st.markdown('<p class="sub-title">Paste documentation URLs below. One URL per line. '
            'Export as a formatted Word doc or clean plain text for AI ingestion.</p>',
            unsafe_allow_html=True)
st.divider()

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =============================================================================
# SECTION 2: TEXT CLEANING
# =============================================================================

SEE_ALSO_PHRASES = [
    "see also", "related topics", "related links", "related articles",
    "further reading", "next steps", "what's next", "whats next",
    "additional resources", "learn more", "more information",
    "in this section", "on this page",
]

def clean_text(text: str) -> str:
    """Strips unicode artefacts and normalises whitespace."""
    text = text.replace("\u00c2", "")
    text = text.replace("\u00a0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\u200c", "")
    text = text.replace("\u2019", "'")
    text = text.replace("\u2018", "'")
    text = text.replace("\u201c", '"')
    text = text.replace("\u201d", '"')
    text = re.sub(r" {2,}", " ", text)
    return text.strip()


def is_see_also_heading(text: str) -> bool:
    """Returns True if a heading signals the start of navigation-only content."""
    return any(phrase in text.lower() for phrase in SEE_ALSO_PHRASES)


# =============================================================================
# SECTION 3: TABLE HELPERS
# NOTE: is_layout_table() and extract_table() are TWO SEPARATE top-level
# functions. is_layout_table() must come first because extract_table()
# does NOT call it — walk() calls both independently.
# =============================================================================

def is_layout_table(table_tag) -> bool:
    """
    Returns True if a <table> is used for page layout rather than data.

    Layout tables (image next to text, multi-column layouts) should be
    walked normally so their text content is extracted as paragraphs.

    Data tables (parameter lists, field definitions) should become Word tables.

    SIGNALS OF A LAYOUT TABLE:
        - Has role="presentation" attribute
        - Has summary="" attribute (old HTML convention)
        - Contains any <img> tag inside its cells
        - Has no <th> cells AND cells contain long paragraphs (avg > 200 chars)
    """
    # Explicit HTML attributes declaring layout intent
    if table_tag.get("role") == "presentation":
        return True
    if table_tag.get("summary") == "":
        return True

    # If any cell contains an image, it's a layout table
    # We decompose() the img so it doesn't get double-processed
    # when walk() recurses into the table's cells
    imgs = table_tag.find_all("img")
    if imgs:
        for img in imgs:
            img.decompose()
        return True

    # No <th> cells + long cell content = layout table
    has_th = bool(table_tag.find("th"))
    if not has_th:
        cells = table_tag.find_all("td")
        if cells:
            avg_len = sum(len(c.get_text(strip=True)) for c in cells) / len(cells)
            if avg_len > 200:
                return True

    return False


def extract_table(table_tag) -> dict:
    """
    Converts a BeautifulSoup <table> tag into a structured dict.
    Only called for DATA tables (after is_layout_table() returns False).

    RETURNS:
        {
          'type': 'table',
          'headers': ['Col1', 'Col2', ...],
          'rows': [['val1', 'val2'], ...],
          'col_count': 4
        }
    """
    headers = []
    rows    = []

    # Extract header row from <thead> if it exists
    thead = table_tag.find("thead")
    if thead:
        for th in thead.find_all("th"):
            colspan   = int(th.get("colspan", 1))
            cell_text = clean_text(th.get_text(separator=" ", strip=True))
            headers.extend([cell_text] * colspan)
        if not headers:
            for td in thead.find_all("td"):
                colspan   = int(td.get("colspan", 1))
                cell_text = clean_text(td.get_text(separator=" ", strip=True))
                headers.extend([cell_text] * colspan)

    # Extract data rows from <tbody> (or whole table if no tbody)
    tbody = table_tag.find("tbody") or table_tag

    for tr in tbody.find_all("tr", recursive=False):
        cells = tr.find_all(["td", "th"])
        if not cells:
            continue
        row    = []
        all_th = all(c.name == "th" for c in cells)
        for cell in cells:
            colspan   = int(cell.get("colspan", 1))
            cell_text = clean_text(cell.get_text(separator=" ", strip=True))
            row.extend([cell_text] * colspan)
        if all_th and not headers:
            headers = row
        else:
            if any(cell for cell in row):
                rows.append(row)

    col_count = max(
        len(headers),
        max((len(r) for r in rows), default=0)
    )

    return {
        "type":      "table",
        "headers":   headers,
        "rows":      rows,
        "col_count": col_count,
    }


# =============================================================================
# SECTION 4: CORE SCRAPING FUNCTION
# =============================================================================

BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
    "Accept-Encoding": "gzip, deflate",
    "Connection":      "keep-alive",
}

def scrape_page(url: str) -> dict:
    """
    Fetches a URL and returns a structured list of content elements.
    Images are detected but NOT downloaded — only alt text is stored.
    Tables are classified as layout or data before processing.
    """
    try:
        response = requests.get(url, headers=BROWSER_HEADERS,
                                timeout=20, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        return {"url": url, "error": str(e), "elements": []}

    soup = BeautifulSoup(response.text, "html.parser")

    title_tag  = soup.find("title")
    page_title = clean_text(title_tag.get_text(strip=True)) if title_tag else url

    # Strip noise elements
    noise_selectors = [
        "nav", "footer", "header", "aside",
        "script", "style",
        "[class*='sidebar']",      "[class*='nav']",
        "[class*='menu']",         "[class*='toc']",
        "[class*='cookie']",       "[class*='banner']",
        "[id*='sidebar']",         "[id*='nav']",
        "[class*='seealso']",      "[class*='see-also']",
        "[class*='related']",      "[class*='relatedLinks']",
        "[class*='relatedTopics']","[class*='footer-links']",
        "[id*='seealso']",         "[id*='related']",
    ]
    for selector in noise_selectors:
        for tag in soup.select(selector):
            tag.decompose()

    # Find main content area
    content_area = (
        soup.find("main")
        or soup.find("article")
        or soup.find(class_=lambda c: c and any(
            kw in c for kw in ["content", "article", "post-body",
                                "docs-content", "markdown-body", "entry-content"]
        ))
        or soup.find("body")
    )

    if content_area is None:
        return {"url": url, "title": page_title,
                "error": "Could not find content area.", "elements": []}

    elements        = []
    stop_extraction = False

    def walk(node):
        nonlocal stop_extraction

        for child in node.children:
            if stop_extraction:
                return

            if not hasattr(child, "name") or child.name is None:
                continue

            tag_name = child.name.lower()

            # ---- HEADINGS ----
            if tag_name in ("h1", "h2", "h3", "h4", "h5", "h6"):
                text = clean_text(child.get_text(separator=" ", strip=True))
                if is_see_also_heading(text):
                    stop_extraction = True
                    return
                if text:
                    elements.append({"type": tag_name, "text": text})

            # ---- TABLES ----
            elif tag_name == "table":
                if is_layout_table(child):
                    # Layout table: ignore structure, walk cells for text
                    walk(child)
                else:
                    # Data table: extract as structured grid
                    table_data = extract_table(child)
                    if table_data["rows"] or table_data["headers"]:
                        elements.append(table_data)

            # ---- PARAGRAPHS ----
            elif tag_name == "p":
                text = clean_text(child.get_text(separator=" ", strip=True))
                if text and len(text) > 2:
                    elements.append({"type": "p", "text": text})

            # ---- CODE BLOCKS ----
            elif tag_name in ("pre", "code"):
                text = clean_text(child.get_text(strip=True))
                if text:
                    elements.append({"type": "code", "text": text})

            # ---- BLOCKQUOTES ----
            elif tag_name == "blockquote":
                text = clean_text(child.get_text(separator=" ", strip=True))
                if text:
                    elements.append({"type": "blockquote", "text": text})

            # ---- UNORDERED LISTS ----
            elif tag_name == "ul":
                for li in child.find_all("li", recursive=False):
                    text = clean_text(li.get_text(separator=" ", strip=True))
                    if text:
                        elements.append({"type": "li", "text": text,
                                         "list_type": "ul"})

            # ---- ORDERED LISTS ----
            elif tag_name == "ol":
                for li in child.find_all("li", recursive=False):
                    text = clean_text(li.get_text(separator=" ", strip=True))
                    if text:
                        elements.append({"type": "li", "text": text,
                                         "list_type": "ol"})

            # ---- IMAGES — store alt text only, no download ----
            elif tag_name == "img":
                alt = clean_text(child.get("alt", ""))
                if alt and alt.lower() not in ("image", ""):
                    elements.append({"type": "image", "alt": alt})

            # ---- RECURSE INTO CONTAINERS ----
            elif tag_name in ("div", "section", "main", "article", "figure",
                              "details", "summary"):
                walk(child)

    walk(content_area)

    return {"url": url, "title": page_title, "elements": elements}


# =============================================================================
# SECTION 5: WORD TABLE HELPER
# =============================================================================

def add_word_table(doc, table_data: dict):
    """Renders a table dict as a real Word table with shaded header row."""
    headers   = table_data["headers"]
    rows      = table_data["rows"]
    col_count = table_data["col_count"]

    if col_count == 0:
        return

    has_header_row = bool(headers)
    total_rows     = (1 if has_header_row else 0) + len(rows)

    if total_rows == 0:
        return

    table       = doc.add_table(rows=total_rows, cols=col_count)
    table.style = "Table Grid"

    if has_header_row:
        header_row = table.rows[0]
        for col_idx, header_text in enumerate(headers[:col_count]):
            cell      = header_row.cells[col_idx]
            cell.text = header_text
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold      = True
                    run.font.size = Pt(10)
            tc   = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd  = OxmlElement("w:shd")
            shd.set(qn("w:val"),   "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"),  "D9E1F2")
            tcPr.append(shd)

    row_offset = 1 if has_header_row else 0
    for row_idx, row_data in enumerate(rows):
        word_row = table.rows[row_idx + row_offset]
        for col_idx, cell_text in enumerate(row_data[:col_count]):
            cell      = word_row.cells[col_idx]
            cell.text = cell_text
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)

    doc.add_paragraph()


# =============================================================================
# SECTION 6: PLAIN TEXT TABLE RENDERER
# =============================================================================

def render_text_table(table_data: dict) -> str:
    """
    Renders a table as a pipe-separated ASCII grid for plain text / AI export.

    Example output:
        | Order | Field Name  | Required |
        |-------|-------------|----------|
        | 0     | Result Set  | X        |
    """
    headers   = table_data["headers"]
    rows      = table_data["rows"]
    col_count = table_data["col_count"]

    if col_count == 0:
        return ""

    all_rows = []
    if headers:
        all_rows.append(headers)
    all_rows.extend(rows)
    all_rows = [row + [""] * (col_count - len(row)) for row in all_rows]

    col_widths = []
    for col_idx in range(col_count):
        max_width = max(len(str(row[col_idx])) for row in all_rows) if all_rows else 5
        col_widths.append(max(max_width, 3))

    def format_row(row):
        cells = [str(row[i]).ljust(col_widths[i]) for i in range(col_count)]
        return "| " + " | ".join(cells) + " |"

    def separator_row():
        return "|-" + "-|-".join("-" * w for w in col_widths) + "-|"

    lines = []
    if headers:
        lines.append(format_row(headers))
        lines.append(separator_row())
        for row in rows:
            lines.append(format_row(row + [""] * (col_count - len(row))))
    else:
        for i, row in enumerate(rows):
            lines.append(format_row(row + [""] * (col_count - len(row))))
            if i == 0:
                lines.append(separator_row())

    return "\n".join(lines)


# =============================================================================
# SECTION 7: HYPERLINK HELPER
# =============================================================================

def add_hyperlink(paragraph, text: str, url: str):
    """Inserts a clickable hyperlink into a python-docx paragraph."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run  = OxmlElement("w:r")
    rPr  = OxmlElement("w:rPr")
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rPr.append(rStyle)
    run.append(rPr)
    t      = OxmlElement("w:t")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink


# =============================================================================
# SECTION 8: WORD DOCUMENT BUILDER
# =============================================================================

def build_word_document(pages: list, doc_title: str) -> io.BytesIO:
    """Assembles all scraped pages into a formatted Word .docx file."""
    doc = Document()

    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    normal_style           = doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(11)

    # Cover title
    title_para           = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run                  = title_para.add_run(doc_title)
    run.bold             = True
    run.font.size        = Pt(26)
    run.font.color.rgb   = RGBColor(0x1a, 0x1a, 0x2e)
    title_para.space_after = Pt(6)

    sub                      = doc.add_paragraph("Compiled Documentation")
    sub.alignment            = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.size    = Pt(12)
    sub.runs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    doc.add_paragraph()

    for i, page in enumerate(pages):
        if i > 0:
            doc.add_page_break()

        page_heading = doc.add_heading(page.get("title", page["url"]), level=1)
        if page_heading.runs:
            page_heading.runs[0].font.color.rgb = RGBColor(0x00, 0x72, 0xC6)

        source_para                    = doc.add_paragraph("Source: ")
        source_para.runs[0].font.size  = Pt(9)
        source_para.runs[0].italic     = True
        source_para.runs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        add_hyperlink(source_para, page["url"], page["url"])
        doc.add_paragraph()

        if "error" in page and not page.get("elements"):
            err_para = doc.add_paragraph(f"Could not retrieve content: {page['error']}")
            if err_para.runs:
                err_para.runs[0].font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
            continue

        for elem in page.get("elements", []):
            etype = elem["type"]

            # ---- HEADINGS ----
            if etype in ("h1", "h2", "h3", "h4", "h5", "h6"):
                word_level = min(int(etype[1]) + 1, 9)
                doc.add_heading(elem["text"], level=word_level)

            # ---- PARAGRAPHS ----
            elif etype == "p":
                doc.add_paragraph(elem["text"])

            # ---- TABLES ----
            elif etype == "table":
                add_word_table(doc, elem)

            # ---- CODE BLOCKS ----
            elif etype == "code":
                code_para = doc.add_paragraph(elem["text"])
                code_run  = (code_para.runs[0] if code_para.runs
                             else code_para.add_run(elem["text"]))
                code_run.font.name = "Courier New"
                code_run.font.size = Pt(9)

            # ---- BLOCKQUOTES ----
            elif etype == "blockquote":
                bq_para = doc.add_paragraph(elem["text"])
                bq_para.paragraph_format.left_indent = Inches(0.5)
                if bq_para.runs:
                    bq_para.runs[0].italic = True

            # ---- LIST ITEMS ----
            elif etype == "li":
                style = ("List Number" if elem.get("list_type") == "ol"
                         else "List Bullet")
                doc.add_paragraph(elem["text"], style=style)

            # ---- IMAGES — alt text label only, no download ----
            elif etype == "image":
                alt = elem.get("alt", "")
                if alt:
                    note = doc.add_paragraph(f"[Image: {alt}]")
                    if note.runs:
                        note.runs[0].italic    = True
                        note.runs[0].font.size = Pt(9)
                        note.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# =============================================================================
# SECTION 9: PLAIN TEXT BUILDER
# =============================================================================

def build_plain_text(pages: list, doc_title: str) -> io.BytesIO:
    """Converts scraped pages into clean plain text optimised for AI input."""
    lines = []
    lines.append("=" * 60)
    lines.append(doc_title.upper())
    lines.append("=" * 60)
    lines.append("")

    for page in pages:
        lines.append("=" * 60)
        lines.append(page.get("title", page["url"]))
        lines.append(f"Source: {page['url']}")
        lines.append("=" * 60)
        lines.append("")

        if "error" in page and not page.get("elements"):
            lines.append(f"[ERROR: {page['error']}]")
            lines.append("")
            continue

        for elem in page.get("elements", []):
            etype = elem["type"]

            if etype == "h1":
                lines.append(f"# {elem['text']}")
                lines.append("")
            elif etype == "h2":
                lines.append(f"## {elem['text']}")
                lines.append("")
            elif etype == "h3":
                lines.append(f"### {elem['text']}")
                lines.append("")
            elif etype in ("h4", "h5", "h6"):
                lines.append(f"#### {elem['text']}")
                lines.append("")
            elif etype == "p":
                lines.append(elem["text"])
                lines.append("")
            elif etype == "li":
                lines.append(f"  • {elem['text']}")
            elif etype == "code":
                lines.append("[CODE]")
                lines.append(elem["text"])
                lines.append("[/CODE]")
                lines.append("")
            elif etype == "blockquote":
                lines.append(f'  "{elem["text"]}"')
                lines.append("")
            elif etype == "table":
                ascii_table = render_text_table(elem)
                if ascii_table:
                    lines.append(ascii_table)
                    lines.append("")
            elif etype == "image":
                alt = elem.get("alt", "")
                if alt:
                    lines.append(f"[IMAGE: {alt}]")

        lines.append("")

    buffer = io.BytesIO("\n".join(lines).encode("utf-8"))
    buffer.seek(0)
    return buffer


# =============================================================================
# SECTION 10: STREAMLIT UI
# =============================================================================

col_left, col_right = st.columns([2, 1])

with col_left:
    st.subheader("① Enter URLs")
    url_input = st.text_area(
        label="One URL per line:",
        height=220,
        placeholder=(
            "https://docs.example.com/getting-started\n"
            "https://docs.example.com/installation\n"
            "https://docs.example.com/api-reference"
        ),
        help="Paste as many URLs as you like. They appear in the document in this order.",
    )

    doc_title_input = st.text_input(
        label="Document Title (appears on the first page)",
        value="Documentation Compilation",
    )

    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        word_button = st.button(
            "📄 Export as Word (.docx)",
            type="primary",
            use_container_width=True,
        )
    with btn_col2:
        txt_button = st.button(
            "🤖 Export as Plain Text (.txt)",
            type="secondary",
            use_container_width=True,
        )

with col_right:
    st.subheader("ℹ How it works")
    st.info(
        "**📄 Word** — real tables with shaded headers, "
        "headings, bullets. Images shown as labels.\n\n"
        "**🤖 Plain Text** — pipe-separated ASCII tables, "
        "markdown headings. Best for AI tools.\n\n"
        "---\n"
        "**Filtered out:**\n"
        "- Navigation menus & sidebars\n"
        "- 'See Also' / 'Related Topics' sections\n"
        "- Cookie banners & footers\n"
        "- Unicode artefacts (stray  characters)\n"
        "- Layout tables (image+text side-by-side)"
    )

st.divider()

# =============================================================================
# SECTION 11: SCRAPING + EXPORT LOGIC
# =============================================================================

compile_triggered = word_button or txt_button
export_mode       = "word" if word_button else "txt"

if compile_triggered:

    if not url_input.strip():
        st.error("Please enter at least one URL.")
        st.stop()

    raw_urls   = [line.strip() for line in url_input.strip().splitlines()]
    valid_urls = [u for u in raw_urls if u.startswith("http")]

    if not valid_urls:
        st.error("No valid URLs found. Each URL must start with http:// or https://")
        st.stop()

    skipped = [u for u in raw_urls if u and not u.startswith("http")]
    if skipped:
        st.warning(f"Skipped {len(skipped)} line(s) that didn't look like URLs.")

    mode_label = "Word document" if export_mode == "word" else "plain text file"
    st.info(f"Found **{len(valid_urls)} URL(s)**. Building {mode_label}…")

    scraped_pages = []
    progress_bar  = st.progress(0, text="Initialising…")

    with st.status("Scraping pages…", expanded=True) as status_box:
        for idx, url in enumerate(valid_urls):
            progress_bar.progress(
                idx / len(valid_urls),
                text=f"Scraping {idx + 1} of {len(valid_urls)}: {url}"
            )
            st.write(f"🔍 Fetching: `{url}`")
            result = scrape_page(url)
            scraped_pages.append(result)

            if "error" in result and not result.get("elements"):
                st.write(f"  ⚠ Error: {result['error']}")
            else:
                elem_count  = len(result.get("elements", []))
                table_count = sum(1 for e in result.get("elements", [])
                                  if e["type"] == "table")
                st.write(f"  ✅ Done — {elem_count} elements ({table_count} table(s))")

        progress_bar.progress(1.0, text="Scraping complete. Building output…")
        status_box.update(label="Scraping complete!", state="complete", expanded=False)

    safe_filename = (
        doc_title_input.strip().replace(" ", "_").replace("/", "-") or "documentation"
    )

    if export_mode == "word":
        with st.spinner("Assembling Word document…"):
            docx_buffer = build_word_document(scraped_pages, doc_title_input)
        st.success(f"✅ Word document compiled from {len(valid_urls)} page(s)!")
        st.download_button(
            label="⬇ Download Word Document (.docx)",
            data=docx_buffer,
            file_name=f"{safe_filename}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    elif export_mode == "txt":
        with st.spinner("Assembling plain text file…"):
            txt_buffer = build_plain_text(scraped_pages, doc_title_input)
        st.success(f"✅ Plain text compiled from {len(valid_urls)} page(s)!")

        txt_buffer.seek(0)
        preview_lines = "\n".join(
            txt_buffer.read().decode("utf-8").splitlines()[:60]
        )
        with st.expander("👁 Preview first 60 lines"):
            st.code(preview_lines, language=None)

        txt_buffer.seek(0)
        st.download_button(
            label="⬇ Download Plain Text (.txt)",
            data=txt_buffer,
            file_name=f"{safe_filename}.txt",
            mime="text/plain",
            use_container_width=True,
        )
        st.info(
            "💡 **AI tip:** Open the .txt, press `Ctrl+A` → `Ctrl+C`, "
            "paste into your AI chat alongside your question."
        )

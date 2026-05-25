# =============================================================================
# app.py — Documentation Scraper & Word Document Generator
# =============================================================================
# VERSION 2: Adds the following improvements over v1:
#   1. Strips "See Also" / "Related Topics" sections before they enter the doc
#   2. Cleans unicode artefacts (Â, non-breaking spaces) from all text
#   3. Adds a plain text (.txt) export — ideal for AI ingestion
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

st.markdown('<p class="main-title">📄 Documentation Scraper → Word / TXT</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Paste documentation URLs below. One URL per line. '
            'Export as a formatted Word doc or clean plain text for AI ingestion.</p>',
            unsafe_allow_html=True)
st.divider()

# =============================================================================
# SECTION 2: SSL WARNING SUPPRESSION
# =============================================================================

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# =============================================================================
# SECTION 3: TEXT CLEANING HELPER
# =============================================================================

# These are "See Also"-style headings that signal the end of real content.
# Anything under these headings is navigation, not knowledge — strip it.
SEE_ALSO_PHRASES = [
    "see also", "related topics", "related links", "related articles",
    "further reading", "next steps", "what's next", "whats next",
    "additional resources", "learn more", "more information",
    "in this section", "on this page",
]

def clean_text(text: str) -> str:
    """
    Removes unicode artefacts and normalises whitespace from any scraped text.

    WHY THIS IS NEEDED:
        When a server sends HTML encoded in Latin-1 but the browser interprets
        it as UTF-8, characters like non-breaking spaces (U+00A0) show up as
        the visible character Â. This function strips those artefacts so the
        output text is clean for both humans and AI models.
    """
    # Remove the Â character that appears from Latin-1/UTF-8 mismatch
    text = text.replace("\u00c2", "")
    # Replace non-breaking spaces with regular spaces
    text = text.replace("\u00a0", " ")
    # Replace other common unicode noise characters
    text = text.replace("\u200b", "")   # zero-width space
    text = text.replace("\u200c", "")   # zero-width non-joiner
    text = text.replace("\u2019", "'")  # right single quotation mark → apostrophe
    text = text.replace("\u2018", "'")  # left single quotation mark → apostrophe
    text = text.replace("\u201c", '"')  # left double quotation mark
    text = text.replace("\u201d", '"')  # right double quotation mark
    # Collapse multiple spaces into one
    text = re.sub(r" {2,}", " ", text)
    return text.strip()


def is_see_also_heading(text: str) -> bool:
    """
    Returns True if a heading text signals the start of a navigation section.
    We use this to stop content extraction before the "See Also" noise begins.
    """
    return any(phrase in text.lower() for phrase in SEE_ALSO_PHRASES)


# =============================================================================
# SECTION 4: CORE SCRAPING FUNCTION
# =============================================================================

def scrape_page(url: str) -> dict:
    """
    Fetches a single URL and extracts its structured content as a list of
    element dicts. Each dict has a 'type' key ('h1'–'h6', 'p', 'li', 'image',
    'code', 'blockquote') and a 'text' or 'src' key with the content.

    Improvements in v2:
    - Stops walking the content tree when a "See Also" heading is encountered
    - Cleans all text through clean_text() before storing it
    - Strips additional noise selectors specific to GE Vernova / docs sites
    """

    headers = {
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

    try:
        response = requests.get(url, headers=headers, timeout=20, verify=False)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        return {"url": url, "error": str(e), "elements": []}

    soup = BeautifulSoup(response.text, "html.parser")

    # Extract page title
    title_tag = soup.find("title")
    page_title = clean_text(title_tag.get_text(strip=True)) if title_tag else url

    # --- STRIP NOISE ELEMENTS ---
    # v2 adds: seealso, related, relatedLinks class selectors
    # and common "See Also" container IDs used by docs platforms
    noise_selectors = [
        "nav", "footer", "header", "aside",
        "script", "style",
        "[class*='sidebar']",     "[class*='nav']",
        "[class*='menu']",        "[class*='toc']",
        "[class*='cookie']",      "[class*='banner']",
        "[id*='sidebar']",        "[id*='nav']",
        # --- NEW IN v2: "See Also" container classes ---
        "[class*='seealso']",     "[class*='see-also']",
        "[class*='related']",     "[class*='relatedLinks']",
        "[class*='relatedTopics']","[class*='footer-links']",
        "[id*='seealso']",        "[id*='related']",
    ]
    for selector in noise_selectors:
        for tag in soup.select(selector):
            tag.decompose()

    # --- FIND MAIN CONTENT AREA ---
    content_area = (
        soup.find("main")
        or soup.find("article")
        or soup.find(class_=lambda c: c and any(
            kw in c for kw in
            ["content", "article", "post-body",
             "docs-content", "markdown-body", "entry-content"]
        ))
        or soup.find("body")
    )

    if content_area is None:
        return {"url": url, "title": page_title,
                "error": "Could not find content area.", "elements": []}

    # --- WALK THE CONTENT TREE ---
    elements = []
    # This flag is set to True the moment we hit a "See Also" heading.
    # Once True, we stop adding any more elements — the rest is navigation noise.
    stop_extraction = False

    def walk(node):
        nonlocal stop_extraction  # Allows us to modify the outer variable

        for child in node.children:
            if stop_extraction:
                return  # Exit immediately — we've hit a "See Also" section

            if not hasattr(child, "name") or child.name is None:
                continue

            tag_name = child.name.lower()

            # ---- HEADINGS ----
            if tag_name in ("h1", "h2", "h3", "h4", "h5", "h6"):
                text = clean_text(child.get_text(separator=" ", strip=True))

                # v2: Check if this heading signals end of real content
                if is_see_also_heading(text):
                    stop_extraction = True  # Flip the flag
                    return                  # Stop processing immediately

                if text:
                    elements.append({"type": tag_name, "text": text})

            # ---- PARAGRAPHS ----
            elif tag_name == "p":
                text = clean_text(child.get_text(separator=" ", strip=True))
                # v2: Skip paragraphs that are just a single unicode character
                # (e.g., the stray Â that appears after "See Also" on GE Vernova)
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
                        elements.append({"type": "li", "text": text, "list_type": "ul"})

            # ---- ORDERED LISTS ----
            elif tag_name == "ol":
                for li in child.find_all("li", recursive=False):
                    text = clean_text(li.get_text(separator=" ", strip=True))
                    if text:
                        elements.append({"type": "li", "text": text, "list_type": "ol"})

            # ---- IMAGES ----
            elif tag_name == "img":
                src = child.get("src", "")
                alt = clean_text(child.get("alt", "Image"))
                if src:
                    abs_src = urllib.parse.urljoin(url, src)
                    elements.append({"type": "image", "src": abs_src, "alt": alt})

            # ---- RECURSE INTO CONTAINERS ----
            elif tag_name in ("div", "section", "main", "article", "figure",
                              "details", "summary", "td", "th"):
                walk(child)

    walk(content_area)

    return {"url": url, "title": page_title, "elements": elements}


# =============================================================================
# SECTION 5: PLAIN TEXT BUILDER (NEW IN v2)
# =============================================================================

def build_plain_text(pages: list, doc_title: str) -> io.BytesIO:
    """
    Converts scraped pages into a clean plain text file optimised for AI input.

    WHY PLAIN TEXT IS BETTER FOR AI:
        - No formatting overhead (bold, italics, styles) consuming tokens
        - No binary image data — AI can't read embedded images anyway
        - Maximum token efficiency: every character is meaningful content
        - Works via direct paste into ChatGPT, Claude, Gemini, or file upload
        - Consistent structure the AI can parse reliably

    FORMAT:
        ════════════════════ (separator between pages)
        PAGE TITLE
        Source: URL
        ════════════════════

        ## Heading text
        ### Sub-heading text

        Paragraph text

        • Bullet item
        1. Numbered item

        [CODE]
        code block content
        [/CODE]
    """
    lines = []

    # Document header
    lines.append("=" * 60)
    lines.append(doc_title.upper())
    lines.append("=" * 60)
    lines.append("")

    for i, page in enumerate(pages):
        # Page separator
        lines.append("=" * 60)
        lines.append(page.get("title", page["url"]))
        lines.append(f"Source: {page['url']}")
        lines.append("=" * 60)
        lines.append("")

        if "error" in page and not page.get("elements"):
            lines.append(f"[ERROR: Could not retrieve content — {page['error']}]")
            lines.append("")
            continue

        for elem in page.get("elements", []):
            etype = elem["type"]

            # Headings — use markdown-style # symbols so AI understands hierarchy
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

            # Paragraphs — plain text, one blank line after
            elif etype == "p":
                lines.append(elem["text"])
                lines.append("")

            # List items — bullet or numbered
            elif etype == "li":
                if elem.get("list_type") == "ol":
                    lines.append(f"  • {elem['text']}")  # numbered lists as bullets in txt
                else:
                    lines.append(f"  • {elem['text']}")

            # Code blocks — wrapped in markers so AI knows it's code
            elif etype == "code":
                lines.append("[CODE]")
                lines.append(elem["text"])
                lines.append("[/CODE]")
                lines.append("")

            # Blockquotes
            elif etype == "blockquote":
                lines.append(f'  "{elem["text"]}"')
                lines.append("")

            # Images — include alt text as a note; AI can't see the image itself
            elif etype == "image":
                if elem.get("alt") and elem["alt"] != "Image":
                    lines.append(f"[IMAGE: {elem['alt']}]")

        lines.append("")  # Blank line between pages

    # Join all lines and encode as UTF-8 bytes
    full_text = "\n".join(lines)
    buffer = io.BytesIO(full_text.encode("utf-8"))
    buffer.seek(0)
    return buffer


# =============================================================================
# SECTION 6: HYPERLINK HELPER
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
    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rPr.append(rStyle)
    run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    return hyperlink


# =============================================================================
# SECTION 7: WORD DOCUMENT BUILDER
# =============================================================================

def fetch_image(src_url: str, headers: dict):
    """Downloads image bytes. Returns None silently on any failure."""
    try:
        resp = requests.get(src_url, headers=headers, timeout=10, verify=False)
        resp.raise_for_status()
        if "image" in resp.headers.get("Content-Type", ""):
            return resp.content
    except Exception:
        pass
    return None


def build_word_document(pages: list, doc_title: str) -> io.BytesIO:
    """Assembles all scraped pages into a formatted Word .docx file."""

    doc = Document()

    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(11)

    # Cover title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run(doc_title)
    run.bold = True
    run.font.size = Pt(26)
    run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)
    title_para.space_after = Pt(6)

    sub = doc.add_paragraph("Compiled Documentation")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.size = Pt(12)
    sub.runs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    doc.add_paragraph()

    img_headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
    }

    for i, page in enumerate(pages):
        if i > 0:
            doc.add_page_break()

        page_heading = doc.add_heading(page.get("title", page["url"]), level=1)
        if page_heading.runs:
            page_heading.runs[0].font.color.rgb = RGBColor(0x00, 0x72, 0xC6)

        source_para = doc.add_paragraph("Source: ")
        source_para.runs[0].font.size = Pt(9)
        source_para.runs[0].italic = True
        source_para.runs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        add_hyperlink(source_para, page["url"], page["url"])
        doc.add_paragraph()

        if "error" in page and not page.get("elements"):
            err_para = doc.add_paragraph(f"⚠ Could not retrieve content: {page['error']}")
            if err_para.runs:
                err_para.runs[0].font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
            continue

        for elem in page.get("elements", []):
            etype = elem["type"]

            if etype in ("h1", "h2", "h3", "h4", "h5", "h6"):
                level_num  = int(etype[1])
                word_level = min(level_num + 1, 9)
                doc.add_heading(elem["text"], level=word_level)

            elif etype == "p":
                doc.add_paragraph(elem["text"])

            elif etype == "code":
                code_para = doc.add_paragraph(elem["text"])
                code_run = code_para.runs[0] if code_para.runs else code_para.add_run(elem["text"])
                code_run.font.name = "Courier New"
                code_run.font.size = Pt(9)

            elif etype == "blockquote":
                bq_para = doc.add_paragraph(elem["text"])
                bq_para.paragraph_format.left_indent = Inches(0.5)
                if bq_para.runs:
                    bq_para.runs[0].italic = True

            elif etype == "li":
                if elem.get("list_type") == "ol":
                    doc.add_paragraph(elem["text"], style="List Number")
                else:
                    doc.add_paragraph(elem["text"], style="List Bullet")

            elif etype == "image":
                img_bytes = fetch_image(elem["src"], img_headers)
                if img_bytes:
                    img_stream = io.BytesIO(img_bytes)
                    try:
                        doc.add_picture(img_stream, width=Inches(5))
                        if elem.get("alt") and elem["alt"] != "Image":
                            cap = doc.add_paragraph(elem["alt"])
                            if cap.runs:
                                cap.runs[0].italic = True
                                cap.runs[0].font.size = Pt(9)
                                cap.runs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
                    except Exception:
                        pass

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# =============================================================================
# SECTION 8: STREAMLIT USER INTERFACE
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
        help="Paste as many URLs as you like. They will appear in the document in this order.",
    )

    doc_title_input = st.text_input(
        label="Document Title (appears on the first page)",
        value="Documentation Compilation",
        help="Give your compiled document a meaningful title.",
    )

    # --- TWO BUTTONS SIDE BY SIDE ---
    # st.columns() inside the left column gives us two sub-columns for the buttons
    btn_col1, btn_col2 = st.columns(2)

    with btn_col1:
        word_button = st.button(
            "📄 Export as Word (.docx)",
            type="primary",
            use_container_width=True,
            help="Formatted document with headings, bullets, and images. Good for reading.",
        )

    with btn_col2:
        txt_button = st.button(
            "🤖 Export as Plain Text (.txt)",
            type="secondary",
            use_container_width=True,
            help="Clean plain text with no formatting. Optimised for pasting into AI tools.",
        )

with col_right:
    st.subheader("ℹ How it works")
    st.info(
        "1. Paste your documentation URLs (one per line).\n\n"
        "2. Choose your export format:\n\n"
        "   **📄 Word** — formatted doc with headings, "
        "bullets, images. Best for reading & sharing.\n\n"
        "   **🤖 Plain Text** — clean text only. "
        "Best for pasting into ChatGPT, Claude, or Gemini.\n\n"
        "---\n"
        "**What's filtered out:**\n"
        "- Navigation menus & sidebars\n"
        "- 'See Also' / 'Related Topics' sections\n"
        "- Cookie banners & footers\n"
        "- Unicode artefacts (e.g. stray Â characters)"
    )

st.divider()

# =============================================================================
# SECTION 9: SHARED SCRAPING LOGIC
# =============================================================================
# Both buttons trigger the same scraping process.
# We detect which button was pressed and only change the output format.

compile_triggered = word_button or txt_button
export_mode = "word" if word_button else "txt"

if compile_triggered:

    if not url_input.strip():
        st.error("Please enter at least one URL before exporting.")
        st.stop()

    raw_urls   = [line.strip() for line in url_input.strip().splitlines()]
    valid_urls = [u for u in raw_urls if u.startswith("http")]

    if not valid_urls:
        st.error("No valid URLs found. Make sure each URL starts with http:// or https://")
        st.stop()

    skipped = [u for u in raw_urls if u and not u.startswith("http")]
    if skipped:
        st.warning(f"Skipped {len(skipped)} line(s) that didn't look like URLs: {skipped}")

    mode_label = "Word document" if export_mode == "word" else "plain text file"
    st.info(f"Found **{len(valid_urls)} URL(s)**. Scraping and building {mode_label}…")

    # --- SCRAPING LOOP ---
    scraped_pages = []
    progress_bar  = st.progress(0, text="Initialising…")

    with st.status("Scraping pages…", expanded=True) as status_box:
        for idx, url in enumerate(valid_urls):
            fraction = idx / len(valid_urls)
            progress_bar.progress(fraction, text=f"Scraping {idx + 1} of {len(valid_urls)}: {url}")
            st.write(f"🔍 Fetching: `{url}`")

            result = scrape_page(url)
            scraped_pages.append(result)

            if "error" in result and not result.get("elements"):
                st.write(f"  ⚠ Error: {result['error']}")
            else:
                elem_count = len(result.get("elements", []))
                st.write(f"  ✅ Done — {elem_count} content elements extracted.")

        progress_bar.progress(1.0, text="Scraping complete. Building output file…")
        status_box.update(label="Scraping complete!", state="complete", expanded=False)

    safe_filename = (
        doc_title_input.strip().replace(" ", "_").replace("/", "-") or "documentation"
    )

    # --- WORD EXPORT ---
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

    # --- PLAIN TEXT EXPORT ---
    elif export_mode == "txt":
        with st.spinner("Assembling plain text file…"):
            txt_buffer = build_plain_text(scraped_pages, doc_title_input)

        st.success(f"✅ Plain text file compiled from {len(valid_urls)} page(s)!")

        # Show a preview of the first 50 lines so the user can verify quality
        txt_buffer.seek(0)
        preview_text = txt_buffer.read().decode("utf-8")
        preview_lines = "\n".join(preview_text.splitlines()[:50])

        with st.expander("👁 Preview first 50 lines (click to expand)"):
            st.code(preview_lines, language=None)

        txt_buffer.seek(0)  # Rewind after the preview read

        st.download_button(
            label="⬇ Download Plain Text (.txt)",
            data=txt_buffer,
            file_name=f"{safe_filename}.txt",
            mime="text/plain",
            use_container_width=True,
        )

        # Tip for AI use
        st.info(
            "💡 **AI usage tip:** Open the downloaded .txt file, select all text "
            "(`Ctrl+A`), copy it (`Ctrl+C`), and paste it directly into your AI "
            "chat alongside your question. Or upload the .txt file if your AI tool "
            "supports file attachments."
        )

# =============================================================================
# app.py — Documentation Scraper & Word Document Generator
# =============================================================================
# WHAT THIS FILE DOES:
#   This is the entire application. Streamlit reads this single Python file and
#   turns it into a web page with interactive widgets. When a user pastes URLs
#   and clicks a button, Streamlit calls our Python functions to scrape content
#   and assemble a Word document — all without leaving this one file.
#
# HOW TO RUN (after setup):
#   streamlit run app.py
# =============================================================================

# --- IMPORTS ------------------------------------------------------------------
# Think of imports as "loading toolboxes" before you start a job.

import streamlit as st
# Streamlit is the framework that creates the entire web UI.
# Every `st.` call you see below draws a widget on the page.

import requests
# The `requests` library is Python's standard way to fetch web pages.
# It's like a programmatic browser — it sends HTTP GET requests and returns
# the raw HTML as text.

from bs4 import BeautifulSoup
# BeautifulSoup ("BS4") is an HTML parser. Raw HTML is a messy string of tags.
# BS4 turns it into a structured tree you can navigate (e.g., find all <h2> tags).

from docx import Document
# python-docx lets us create and edit Word (.docx) files from Python.
# `Document()` creates a new, blank Word document in memory.

from docx.shared import Inches, Pt, RGBColor
# These are unit helpers for python-docx:
#   Inches()  — sets image/margin sizes in inches
#   Pt()      — sets font sizes in points (like in Word)
#   RGBColor()— sets colours using red, green, blue values (0–255)

from docx.enum.text import WD_ALIGN_PARAGRAPH
# WD_ALIGN_PARAGRAPH gives us constants like CENTER, LEFT, RIGHT for text alignment.

from docx.oxml.ns import qn
# `qn` ("qualified name") lets us write raw XML attributes when python-docx's
# high-level API doesn't cover what we need (e.g., adding hyperlinks).

from docx.oxml import OxmlElement
# OxmlElement lets us create raw XML nodes to insert directly into the .docx
# XML structure — needed for hyperlinks, which python-docx doesn't support natively.

import io
# `io` provides in-memory byte streams. Instead of saving the .docx to disk,
# we keep it in memory as a stream and hand it directly to Streamlit's download
# button. This is cleaner and works perfectly in cloud deployments.

import urllib3
# urllib3 is the underlying HTTP library that `requests` uses.
# We import it here so we can suppress a specific SSL warning (see below).

import urllib.parse
# urllib.parse helps us work with URLs — specifically, converting relative image
# paths (e.g., "/images/logo.png") into absolute URLs (e.g., "https://docs.example.com/images/logo.png").


# =============================================================================
# SECTION 1: STREAMLIT PAGE CONFIGURATION
# =============================================================================
# This MUST be the very first Streamlit call in the script.
# It sets the browser tab title, the layout width, and the sidebar state.

st.set_page_config(
    page_title="Doc Scraper → Word",   # Text shown in the browser tab
    page_icon="📄",                     # Emoji shown as the favicon
    layout="wide",                      # Use full browser width
    initial_sidebar_state="collapsed",  # Keep sidebar hidden by default
)

# Inject a small block of CSS to make the app look slightly polished.
# `st.markdown(..., unsafe_allow_html=True)` is how you inject raw HTML/CSS.
st.markdown("""
    <style>
        .main-title { font-size: 2.2rem; font-weight: 700; color: #1a1a2e; }
        .sub-title  { font-size: 1rem;   color: #555;      margin-top: -10px; }
        .stDownloadButton > button { background-color: #0072C6; color: white;
                                     font-weight: 600; border-radius: 8px; }
    </style>
""", unsafe_allow_html=True)

# Render the page header
st.markdown('<p class="main-title">📄 Documentation Scraper → Word</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Paste documentation URLs below. One URL per line. '
            'The app will scrape each page and compile them into a single Word document.</p>',
            unsafe_allow_html=True)
st.divider()  # Draws a horizontal line across the page


# =============================================================================
# SECTION 2: SSL WARNING SUPPRESSION
# =============================================================================
# When `verify=False` is passed to requests.get(), Python prints an
# InsecureRequestWarning to the console every single time. This suppresses it.
# WHY verify=False? Some documentation servers have misconfigured or self-signed
# SSL certificates that cause requests to fail. Disabling verification is a
# pragmatic workaround for a scraping tool (NOT recommended for financial/auth flows).

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# =============================================================================
# SECTION 3: CORE SCRAPING FUNCTION
# =============================================================================

def scrape_page(url: str) -> dict:
    """
    Fetches a single URL and extracts its structured content.

    WHY A SEPARATE FUNCTION?
        Keeping scraping logic isolated from UI and document-assembly logic
        is called "separation of concerns." It means: if scraping breaks,
        you only need to fix this one function. If the Word output is wrong,
        you only look at the document functions. Easier to debug.

    RETURNS:
        A dictionary with keys: 'title', 'url', 'elements'
        'elements' is a list of dicts, each describing one piece of content:
            {'type': 'h1',   'text': 'Getting Started'}
            {'type': 'p',    'text': 'This guide explains...'}
            {'type': 'li',   'text': 'Step one', 'list_type': 'ul'}
            {'type': 'image','src':  'https://...', 'alt': 'Diagram'}
        Returns an 'error' key if fetching fails.
    """

    # --- 2a. HTTP REQUEST WITH USER-AGENT SPOOFING ---
    # Many servers block requests that don't look like real browsers.
    # The "User-Agent" header identifies the client. We use a real Chrome
    # UA string so the server thinks a Windows Chrome browser is visiting.
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        # Some servers also check these headers. Including them improves
        # compatibility with stricter sites.
        "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
        "Accept-Encoding": "gzip, deflate",
        "Connection":      "keep-alive",
    }

    try:
        response = requests.get(
            url,
            headers=headers,
            timeout=20,       # Give up if no response within 20 seconds
            verify=False,     # Disable SSL certificate verification (see note above)
        )
        # raise_for_status() checks if the HTTP status code indicates an error
        # (e.g., 404 Not Found, 500 Server Error). If so, it raises an exception
        # that we catch in the `except` block below.
        response.raise_for_status()

    except requests.exceptions.RequestException as e:
        # Something went wrong (network error, timeout, bad status code).
        # We return a dict with an 'error' key so the calling code can handle
        # it gracefully rather than crashing the whole app.
        return {"url": url, "error": str(e), "elements": []}

    # --- 2b. HTML PARSING ---
    # BeautifulSoup takes the raw HTML string and parses it into a tree.
    # "html.parser" is Python's built-in parser — no extra install needed.
    soup = BeautifulSoup(response.text, "html.parser")

    # --- 2c. EXTRACT PAGE TITLE ---
    # The <title> tag in <head> is the page's metadata title.
    # We use it as the heading for this page's section in the Word doc.
    title_tag = soup.find("title")
    page_title = title_tag.get_text(strip=True) if title_tag else url

    # --- 2d. STRIP NOISE ELEMENTS ---
    # Navigation bars, footers, sidebars, and scripts clutter the output.
    # We destroy them before extracting content so they never appear.
    # decompose() removes the tag AND its children from the tree entirely.
    noise_selectors = [
        "nav", "footer", "header", "aside",            # Structural layout tags
        "script", "style",                              # Code — never wanted as text
        "[class*='sidebar']", "[class*='nav']",         # Class-based sidebars/nav
        "[class*='menu']",    "[class*='toc']",         # Table-of-contents / menus
        "[class*='cookie']",  "[class*='banner']",      # Cookie banners / ads
        "[id*='sidebar']",    "[id*='nav']",             # ID-based sidebars/nav
    ]
    for selector in noise_selectors:
        for tag in soup.select(selector):
            tag.decompose()  # Remove it from the tree permanently

    # --- 2e. FIND THE MAIN CONTENT CONTAINER ---
    # Most well-structured docs sites wrap their article in a semantic tag
    # like <main>, <article>, or a <div> with class "content" / "docs-body".
    # We try these in order of specificity, falling back to <body> if none match.
    content_area = (
        soup.find("main")                           or  # Semantic HTML5 <main>
        soup.find("article")                        or  # Semantic <article>
        soup.find(class_=lambda c: c and any(
            kw in c for kw in
            ["content", "article", "post-body",
             "docs-content", "markdown-body", "entry-content"]
        ))                                          or  # Common CMS class names
        soup.find("body")                               # Ultimate fallback
    )

    if content_area is None:
        return {"url": url, "title": page_title, "error": "Could not find content area.", "elements": []}

    # --- 2f. WALK THE CONTENT TREE AND BUILD AN ELEMENT LIST ---
    # Instead of extracting raw text, we walk through the content's children
    # in document order and emit structured "element" dicts.
    # This preserves the visual hierarchy (heading → paragraph → list → image).
    elements = []

    # We look for these tag types in order. `find_all` by default is depth-first,
    # which matches the reading order of the HTML source.
    BLOCK_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "img",
                  "pre", "code", "blockquote", "ol", "ul"}

    def walk(node):
        """
        Recursively walk the BS4 tree in source order.
        For each relevant tag, emit one element dict.
        We use recursion so nested structures (e.g., <ul> inside <div>) are captured.
        """
        for child in node.children:
            # NavigableString is BS4's type for plain text between tags — skip them
            # at the top level; their parent tag handles them.
            if not hasattr(child, "name") or child.name is None:
                continue

            tag_name = child.name.lower()

            # ---- HEADINGS ----
            if tag_name in ("h1", "h2", "h3", "h4", "h5", "h6"):
                text = child.get_text(separator=" ", strip=True)
                if text:  # Only add if there's actual text (skip empty heading tags)
                    elements.append({"type": tag_name, "text": text})

            # ---- PARAGRAPHS ----
            elif tag_name == "p":
                text = child.get_text(separator=" ", strip=True)
                if text:
                    elements.append({"type": "p", "text": text})

            # ---- PREFORMATTED / CODE BLOCKS ----
            elif tag_name in ("pre", "code"):
                text = child.get_text(strip=True)
                if text:
                    elements.append({"type": "code", "text": text})

            # ---- BLOCKQUOTES ----
            elif tag_name == "blockquote":
                text = child.get_text(separator=" ", strip=True)
                if text:
                    elements.append({"type": "blockquote", "text": text})

            # ---- UNORDERED LISTS ----
            # We don't emit the <ul> itself; we emit each <li> with a list_type flag.
            elif tag_name == "ul":
                for li in child.find_all("li", recursive=False):  # only direct children
                    text = li.get_text(separator=" ", strip=True)
                    if text:
                        elements.append({"type": "li", "text": text, "list_type": "ul"})

            # ---- ORDERED LISTS ----
            elif tag_name == "ol":
                for li in child.find_all("li", recursive=False):
                    text = li.get_text(separator=" ", strip=True)
                    if text:
                        elements.append({"type": "li", "text": text, "list_type": "ol"})

            # ---- IMAGES ----
            elif tag_name == "img":
                src = child.get("src", "")
                alt = child.get("alt", "Image")
                if src:
                    # Convert relative URLs (e.g., /images/foo.png) to absolute
                    # so we can actually fetch them later.
                    abs_src = urllib.parse.urljoin(url, src)
                    elements.append({"type": "image", "src": abs_src, "alt": alt})

            # ---- RECURSE INTO CONTAINER TAGS ----
            # For divs, sections, and other structural wrappers, we recurse
            # so we don't miss content buried inside them.
            elif tag_name in ("div", "section", "main", "article", "figure",
                              "details", "summary", "td", "th"):
                walk(child)

    walk(content_area)

    return {
        "url":      url,
        "title":    page_title,
        "elements": elements,
    }


# =============================================================================
# SECTION 4: HYPERLINK HELPER
# =============================================================================

def add_hyperlink(paragraph, text: str, url: str):
    """
    Inserts a clickable hyperlink into a python-docx paragraph.

    WHY SO COMPLEX?
        python-docx does not have a built-in add_hyperlink() method.
        We must manually construct the XML elements that Word expects.
        This is the standard workaround used by the python-docx community.

    HOW IT WORKS:
        1. Register the URL in the document's relationship table (like a footnote registry).
        2. Create a <w:hyperlink> XML element referencing that relationship ID.
        3. Style the text inside it to look like a link (blue, underlined).
        4. Attach it to the paragraph's XML tree.
    """
    # Step 1: Get the relationship part of the document and add the URL.
    # Every hyperlink in a .docx is stored as a "relationship" with an ID.
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)

    # Step 2: Create the <w:hyperlink> XML element
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)   # Reference the relationship we just registered

    # Step 3: Create the run (<w:r>) that holds the visible link text
    run = OxmlElement("w:r")

    # Step 4: Add run properties (<w:rPr>) to style it as a hyperlink
    rPr = OxmlElement("w:rPr")

    # Apply the built-in "Hyperlink" character style (blue + underline)
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rPr.append(rStyle)
    run.append(rPr)

    # Step 5: Add the visible text node
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)

    hyperlink.append(run)

    # Step 6: Attach to the paragraph's XML
    paragraph._p.append(hyperlink)

    return hyperlink


# =============================================================================
# SECTION 5: WORD DOCUMENT BUILDER
# =============================================================================

def fetch_image(src_url: str, headers: dict) -> bytes | None:
    """
    Attempts to download an image and return its raw bytes.
    Returns None if the download fails for any reason.

    WHY SEPARATE?
        Image fetching can fail silently for many reasons (broken URL, wrong
        content-type, server blocks image hotlinking). Isolating it means a
        single broken image won't crash the entire document build.
    """
    try:
        resp = requests.get(src_url, headers=headers, timeout=10, verify=False)
        resp.raise_for_status()
        # Basic sanity check: make sure the server actually returned image data
        content_type = resp.headers.get("Content-Type", "")
        if "image" in content_type:
            return resp.content
    except Exception:
        pass  # Silently ignore — the document will just skip this image
    return None


def build_word_document(pages: list, doc_title: str) -> io.BytesIO:
    """
    Takes a list of scraped page dicts and builds a Word document.

    PARAMETERS:
        pages      — list of dicts returned by scrape_page()
        doc_title  — the main title to show on the cover / first page

    RETURNS:
        An io.BytesIO stream containing the complete .docx file bytes.
        Streamlit's download button can consume this directly.
    """

    # --- 4a. CREATE DOCUMENT AND APPLY STYLES ---
    doc = Document()

    # Adjust default margins to give more reading space
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    # Style the default "Normal" paragraph font
    normal_style        = doc.styles["Normal"]
    normal_style.font.name = "Calibri"
    normal_style.font.size = Pt(11)

    # --- 4b. COVER TITLE ---
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # add_run() adds a text chunk inside the paragraph; we can format each run independently
    run = title_para.add_run(doc_title)
    run.bold      = True
    run.font.size = Pt(26)
    run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)  # Dark navy — #1a1a2e

    # Space below the main title
    title_para.space_after = Pt(6)

    # Subtitle line
    sub = doc.add_paragraph("Compiled Documentation")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.size  = Pt(12)
    sub.runs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    doc.add_paragraph()  # Blank line for visual breathing room

    # HTTP headers reused for image fetching
    img_headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
    }

    # --- 4c. PROCESS EACH SCRAPED PAGE ---
    for i, page in enumerate(pages):

        # Start every page section on a new Word page (except the very first,
        # which follows the cover title on page 1).
        if i > 0:
            doc.add_page_break()

        # ---- Page section heading ----
        # Use the page's <title> tag as the section heading
        page_heading = doc.add_heading(page.get("title", page["url"]), level=1)
        page_heading.runs[0].font.color.rgb = RGBColor(0x00, 0x72, 0xC6)  # Microsoft blue

        # ---- Source URL as a clickable hyperlink ----
        source_para = doc.add_paragraph("Source: ")
        source_para.runs[0].font.size  = Pt(9)
        source_para.runs[0].italic     = True
        source_para.runs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        add_hyperlink(source_para, page["url"], page["url"])  # clickable link

        doc.add_paragraph()  # Visual gap after the URL

        # ---- Handle scraping errors ----
        if "error" in page and not page.get("elements"):
            err_para = doc.add_paragraph(f"⚠ Could not retrieve content: {page['error']}")
            err_para.runs[0].font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
            continue  # Skip to the next page

        # ---- Map each element to a Word paragraph ----
        for elem in page.get("elements", []):
            elem_type = elem["type"]

            # ---- HEADINGS (h1–h6) ----
            # We shift the heading level down by 1 because h1 in the source is
            # already used for the page section title above.
            # HTML h1 → Word Heading 2, h2 → Heading 3, ..., h6 → Heading 6
            if elem_type in ("h1", "h2", "h3", "h4", "h5", "h6"):
                level_num  = int(elem_type[1])          # e.g., "h2" → 2
                word_level = min(level_num + 1, 9)      # shift down, cap at 9
                doc.add_heading(elem["text"], level=word_level)

            # ---- PARAGRAPHS ----
            elif elem_type == "p":
                doc.add_paragraph(elem["text"])

            # ---- CODE BLOCKS ----
            # python-docx doesn't have a built-in code style, so we simulate
            # it with a monospace font (Courier New) and light grey background.
            elif elem_type == "code":
                code_para  = doc.add_paragraph(elem["text"])
                code_run   = code_para.runs[0] if code_para.runs else code_para.add_run(elem["text"])
                code_run.font.name = "Courier New"
                code_run.font.size = Pt(9)

            # ---- BLOCKQUOTES ----
            elif elem_type == "blockquote":
                bq_para = doc.add_paragraph(elem["text"])
                # Indent the paragraph to visually suggest a blockquote
                bq_para.paragraph_format.left_indent = Inches(0.5)
                if bq_para.runs:
                    bq_para.runs[0].italic = True

            # ---- LIST ITEMS ----
            elif elem_type == "li":
                # python-docx uses the "List Bullet" and "List Number" built-in
                # styles for unordered and ordered lists respectively.
                if elem.get("list_type") == "ol":
                    doc.add_paragraph(elem["text"], style="List Number")
                else:
                    doc.add_paragraph(elem["text"], style="List Bullet")

            # ---- IMAGES ----
            elif elem_type == "image":
                img_bytes = fetch_image(elem["src"], img_headers)
                if img_bytes:
                    # Wrap in BytesIO so python-docx can read it as a file-like object
                    img_stream = io.BytesIO(img_bytes)
                    try:
                        # Width capped at 5 inches; height scales proportionally.
                        doc.add_picture(img_stream, width=Inches(5))
                        # Add an italicised alt-text caption below the image
                        if elem.get("alt") and elem["alt"] != "Image":
                            cap = doc.add_paragraph(elem["alt"])
                            if cap.runs:
                                cap.runs[0].italic     = True
                                cap.runs[0].font.size  = Pt(9)
                                cap.runs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
                    except Exception:
                        # If the image data is corrupt or unsupported format, skip it
                        pass

    # --- 4d. SAVE TO AN IN-MEMORY BYTE STREAM ---
    # Instead of writing to disk (which would require a file path and cleanup),
    # we write to a BytesIO object — an in-memory "fake file."
    # This is the standard Streamlit pattern for file downloads.
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)   # Rewind to the beginning so the reader starts from byte 0

    return buffer


# =============================================================================
# SECTION 6: STREAMLIT USER INTERFACE
# =============================================================================
# Everything below renders the interactive page widgets.

# Two columns: left for inputs, right for instructions / tips
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

    compile_button = st.button(
        "🚀 Compile Document",
        type="primary",       # Makes it the big blue button
        use_container_width=True,
    )

with col_right:
    st.subheader("ℹ How it works")
    st.info(
        "1. Paste your documentation URLs above (one per line).\n\n"
        "2. Optionally edit the document title.\n\n"
        "3. Click **Compile Document**.\n\n"
        "4. Wait while each page is scraped.\n\n"
        "5. Download the finished `.docx` file.\n\n"
        "---\n"
        "**Tips:**\n"
        "- Pages appear in the order you enter them.\n"
        "- Public documentation pages work best.\n"
        "- Pages behind a login wall cannot be scraped."
    )

st.divider()

# --- MAIN LOGIC: triggered when the button is clicked ---
if compile_button:

    # --- Input validation ---
    if not url_input.strip():
        st.error("Please enter at least one URL before compiling.")
        st.stop()   # st.stop() halts execution of the rest of the script

    # Parse the text area: split by newlines, strip whitespace, remove blanks
    raw_urls = [line.strip() for line in url_input.strip().splitlines()]
    valid_urls = [u for u in raw_urls if u.startswith("http")]  # Basic URL sanity check

    if not valid_urls:
        st.error("No valid URLs found. Make sure each URL starts with http:// or https://")
        st.stop()

    # Warn about any lines that were skipped
    skipped = [u for u in raw_urls if u and not u.startswith("http")]
    if skipped:
        st.warning(f"Skipped {len(skipped)} line(s) that didn't look like URLs: {skipped}")

    st.info(f"Found **{len(valid_urls)} URL(s)**. Starting scrape…")

    # --- Scraping loop with progress bar ---
    # st.progress() creates the visual progress bar.
    # st.status() creates a collapsible "live log" panel.
    scraped_pages = []
    progress_bar  = st.progress(0, text="Initialising…")

    with st.status("Scraping pages…", expanded=True) as status_box:
        for idx, url in enumerate(valid_urls):

            # Update the progress bar (value must be 0.0–1.0)
            fraction = (idx) / len(valid_urls)
            progress_bar.progress(fraction, text=f"Scraping {idx + 1} of {len(valid_urls)}: {url}")

            st.write(f"🔍 Fetching: `{url}`")

            result = scrape_page(url)
            scraped_pages.append(result)

            if "error" in result and not result.get("elements"):
                st.write(f"  ⚠ Error on this page: {result['error']}")
            else:
                elem_count = len(result.get("elements", []))
                st.write(f"  ✅ Done — extracted {elem_count} content elements.")

        progress_bar.progress(1.0, text="Scraping complete. Building Word document…")
        status_box.update(label="Scraping complete!", state="complete", expanded=False)

    # --- Build the Word document ---
    with st.spinner("Assembling your Word document… (this may take a moment for large docs)"):
        docx_buffer = build_word_document(scraped_pages, doc_title_input)

    # --- Success banner and download button ---
    st.success(f"✅ Document compiled successfully from {len(valid_urls)} page(s)!")

    # Sanitise the document title for use as a filename (remove spaces/slashes)
    safe_filename = doc_title_input.strip().replace(" ", "_").replace("/", "-") or "documentation"

    st.download_button(
        label="⬇ Download Word Document (.docx)",
        data=docx_buffer,                          # The BytesIO stream we built
        file_name=f"{safe_filename}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        # ↑ This MIME type tells the browser that the download is a .docx file
        use_container_width=True,
    )
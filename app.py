import streamlit as st
import streamlit.components.v1 as components
import docx
import os
import time
import re
import html as html_module
import urllib.request
import urllib.parse
import ssl
import random
import xml.etree.ElementTree as ET
from urllib.error import URLError
from groq import Groq
import io
import zipfile
from PIL import Image, ImageDraw, ImageFont
import json

from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

# --- Configuration ---
# Check Streamlit secrets first (for cloud deployment), otherwise fallback to hardcoded
try:
    GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except Exception:
    GROQ_API_KEY = ""
MODEL = "llama-3.3-70b-versatile"
MAX_TOKENS_OUTPUT = 4000
FALLBACK_MODELS = [
    "llama-3.3-70b-versatile", 
    "llama-3.1-8b-instant",
    "mixtral-8x7b-32768",
    "gemma2-9b-it"
]

# --- Compact System Prompt (~300 tokens to fit free-tier limits) ---
SYSTEM_PROMPT = """Convert raw article text into clean semantic HTML for WordPress. Rules:
- Output ONLY raw HTML. No markdown fences, no <html>/<head>/<body> tags.
- CONTENT PRESERVATION (CRITICAL): 
  - DO NOT REWRITE, SUMMARIZE, OR CHANGE A SINGLE WORD OF THE CONTENT.
  - YOU MUST OUTPUT THE EXACT SAME WORDS IN THE EXACT SAME ORDER.
  - Your ONLY job is to wrap the original text in HTML tags.

- Headings: exactly one <h1> (title), <h2> (sections), <h3> (sub-sections). Never skip levels.
  - CRITICAL RULE 1: DO NOT make any numbered questions (e.g. "1. Question", "Q2:", "16. Question") into <h2> or <h3> headings under ANY circumstances. All numbered questions MUST be formatted using the <ol> Q&A format below.
  - CRITICAL RULE 2: DO NOT invent, hallucinate, or add any new headings. If the source text chunk starts with a sentence or table, DO NOT ADD a heading before it. (e.g. NEVER add "<h2>LangChain Concepts</h2>" or "<h2>Overview</h2>"). You will be heavily penalized if you invent a heading that is not literally in the text.
- Tables: use <table><thead><th>/<tbody><td> for tabular/comparison data.
- Lists (CRITICAL): ALWAYS convert bullet points (lines starting with "-", "*", "•", or items under a lead-in sentence ending with ":") into proper HTML lists (<ul><li>...</li></ul> or <ol><li>...</li></ol>). NEVER leave bullet items or list items as standalone <p> tags.
- Code Snippets: Wrap code blocks in <pre><code class="language-xyz">...</code></pre> preserving indentation. Wrap inline code in <code>...</code>.

- Q&A / Interview Questions Formatting (CRITICAL): 
  If the text contains interview questions (e.g. "8. What is the Runnable interface?"), you MUST format the Question and Answer exactly like this:
  <ol start="8">
    <li><strong>What is the Runnable interface in LangChain?</strong></li>
  </ol>
  <p>The Runnable interface is the standard way... (Do NOT place this inside the ol or li)</p>
  
  <ol start="9">
    <li><strong>Next Question Here?</strong></li>
  </ol>
  <p>Next answer text here.</p>

  CRITICAL: Do NOT use <h2> or <h3> for the questions. The question text MUST go inside the <li><strong>. The answer text (including MCQ options) MUST go inside the <p> or <ul> following it.
  Maintain the original numbering using start="X".

- Images: Intelligently insert <img src="filename.webp" alt="context"> exactly where [IMAGE PLACEHOLDER: filename.webp] appears in the text.
- IMPORTANT: If instructions or lists of available images are provided in the system prompt, DO NOT output them as HTML text."""

FAQ_PROMPT = """
- FAQ Section (ONLY at the very end of the article): 
  DO NOT format normal interview questions or listicles as FAQs! 
  ONLY use the following Yoast FAQ format IF the original text explicitly contains a section titled "FAQs" or "Frequently Asked Questions" at the very end:
  <h2>FAQs</h2>
  <!-- Frontend Visible FAQ Section -->
  <div class="schema-faq wp-block-yoast-seo-faq-block">
    <div class="schema-faq-section" id="faq-question-1">
      <strong class="schema-faq-question">Question text</strong>
      <p class="schema-faq-answer">Answer text</p>
    </div>
  </div>"""

NO_FAQ_PROMPT = """
- IMPORTANT: DO NOT use the Yoast FAQ schema anywhere in this section! All Q&As here must follow the standard Interview Questions <ol> format."""

# --- Helpers ---

def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        return
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def generate_seo_image_name(client, context_text, default_name, existing_names):
    try:
        if not context_text or len(context_text.strip()) < 15:
            return default_name
        res = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content": "You are an SEO expert. Given a text snippet from an article where an image is placed, generate a descriptive, short, lowercase, space-separated SEO filename ending in .webp without hyphens. Max 4-6 words. Output ONLY the exact filename without quotes, prefixes, or explanation (e.g. langchain architecture diagram.webp)."},
                {"role": "user", "content": f"Context around image:\n{context_text[:600]}"}
            ],
            temperature=0.1,
            max_tokens=30,
        )
        name = strip_markdown_fences(res.choices[0].message.content).strip()
        name = name.replace("-", " ").replace("_", " ")
        name = re.sub(r'[^a-zA-Z0-9. ]', '', name.lower())
        name = re.sub(r'\s+', ' ', name).strip()
        if not name.endswith(".webp") or len(name) < 6:
            return default_name
        
        base_name = name[:-5]
        counter = 1
        final_name = name
        while final_name in existing_names:
            counter += 1
            final_name = f"{base_name} {counter}.webp"
            
        return final_name
    except Exception as e:
        print(f"Error generating SEO image name: {e}")
        return default_name

@st.cache_data(show_spinner=False)
def extract_text_and_images_from_docx(file_bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    client = Groq(api_key=GROQ_API_KEY)
    full_text = []
    webp_images = {}
    image_names = []
    
    img_counter = 1
    
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            para_text = block.text.strip()
            
            # Check for images in the paragraph
            for run in block.runs:
                for drawing in run.element.findall('.//w:drawing', namespaces=run.element.nsmap):
                    for blip in drawing.findall('.//a:blip', namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}):
                        embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if embed:
                            try:
                                image_part = doc.part.related_parts[embed]
                                img_bytes = image_part.blob
                                
                                # convert to webp
                                img = Image.open(io.BytesIO(img_bytes))
                                if img.mode in ("RGBA", "P"):
                                    img = img.convert("RGB")
                                default_name = f"extracted image {img_counter}.webp"
                                
                                context_blocks = [b for b in full_text[-4:] if b.strip() and not b.startswith("[IMAGE")]
                                context_text = "\n".join(context_blocks) + "\n" + para_text
                                webp_name = generate_seo_image_name(client, context_text, default_name, webp_images.keys())
                                
                                buffer = io.BytesIO()
                                img.save(buffer, format="WEBP", quality=80)
                                webp_images[webp_name] = buffer.getvalue()
                                image_names.append(webp_name)
                                
                                # Place a placeholder right in the text
                                full_text.append(f"\n[IMAGE PLACEHOLDER: {webp_name}]\n")
                                img_counter += 1
                            except Exception as e:
                                print(f"Error processing image: {e}")
                                
            if para_text:
                is_list_item = bool(block._element.xpath('.//w:numPr')) or 'List' in block.style.name or 'Bullet' in block.style.name
                if is_list_item and not para_text.startswith(('- ', '* ', '• ', '1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.', '0.')):
                    if not para_text.endswith(':') and not re.match(r'^\d+[\.\)]\s+', para_text):
                        para_text = "- " + para_text
                full_text.append(para_text)
                
        elif isinstance(block, Table):
            for row in block.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                full_text.append(" | ".join(row_data))
                
    return '\n'.join(full_text), webp_images, image_names

def strip_markdown_fences(text):
    text = text.strip()
    if text.startswith("```html"):
        text = text[7:]
    if text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    return text.strip()

def call_groq(client, text, sitemap_urls=None, image_names=None, is_chunk=False, model=MODEL, is_last_chunk=True):
    """Single Groq API call with retry logic for rate-limit errors."""
    if is_chunk:
        user_msg = (
            "Continue structuring this SECTION into SEO HTML. "
            "NO <h1>. Use <h2>/<h3>/<h4>. Output ONLY HTML:\n\n" + text
        )
    else:
        user_msg = "Convert this raw article into structured SEO HTML:\n\n" + text

    system_msg = SYSTEM_PROMPT
    if is_last_chunk:
        system_msg += FAQ_PROMPT
    else:
        system_msg += NO_FAQ_PROMPT
    if sitemap_urls and sitemap_urls.strip():
        system_msg += (
            "\n\n---\nAdd 2-4 internal links from these URLs using "
            "natural anchor text:\n" + sitemap_urls
        )
        
    if image_names and len(image_names) > 0:
        system_msg += (
            "\n\n---\nAvailable images:\n" + ", ".join(image_names) + 
            "\n(Replace [IMAGE PLACEHOLDER: filename.webp] with the corresponding <img src='...'> tags. Do NOT render this list as HTML output.)"
        )

    messages = [
        {"role": "system", "content": system_msg},
        {"role": "user", "content": user_msg}
    ]

    models_to_try = [model]
    for m in FALLBACK_MODELS:
        if m not in models_to_try:
            models_to_try.append(m)

    last_error = None
    for current_model in models_to_try:
        max_retries = 2
        for attempt in range(max_retries):
            try:
                completion = client.chat.completions.create(
                    model=current_model,
                    messages=messages,
                    temperature=0.1,
                    max_tokens=MAX_TOKENS_OUTPUT,
                )
                if current_model != model:
                    print(f"Warning: Fell back to {current_model} due to rate limits.")
                return strip_markdown_fences(completion.choices[0].message.content)
            except Exception as e:
                error_str = str(e)
                last_error = e
                if ("rate_limit" in error_str or "429" in error_str or "413" in error_str or "model_decommissioned" in error_str or "400" in error_str):
                    # Hit rate limit or decommissioned model. Break attempt loop to try next model immediately.
                    break
                else:
                    # Other error, raise immediately
                    raise
    
    raise Exception(f"Rate limits reached across all available AI models (or models decommissioned). Please wait a few minutes for your token quota to reset, or upgrade your API key. Last error: {last_error}")

def split_into_chunks(text, max_words=800):
    """Split text into small chunks (~800 words) to stay under token limits."""
    lines = text.split('\n')
    chunks = []
    current_chunk = []
    current_word_count = 0

    for line in lines:
        line_words = len(line.split())
        if current_word_count + line_words > max_words and current_word_count > 200:
            chunks.append('\n'.join(current_chunk))
            current_chunk = [line]
            current_word_count = line_words
        else:
            current_chunk.append(line)
            current_word_count += line_words

    if current_chunk:
        chunks.append('\n'.join(current_chunk))

    return chunks

def remove_hallucinated_headings(html, original_text):
    import re
    def replace_heading(match):
        heading_tag = match.group(0)
        heading_text = match.group(3).strip()
        if heading_text.lower() == 'faqs':
            return heading_tag
        if heading_text.lower() not in original_text.lower():
            return "" # Strip hallucinated heading
        return heading_tag
    
    return re.sub(r'(<(h2|h3)[^>]*>)(.*?)(</\2>)', replace_heading, html, flags=re.IGNORECASE | re.DOTALL)

def inject_faq_json_ld(html_content):
    import json, re
    qa_pairs = re.findall(
        r'<(?:strong|h[1-6]|div|span)[^>]*class=["\'][^"\']*schema-faq-question[^"\']*["\'][^>]*>(.*?)</(?:strong|h[1-6]|div|span)>\s*<(?:p|div|span)[^>]*class=["\'][^"\']*schema-faq-answer[^"\']*["\'][^>]*>(.*?)</(?:p|div|span)>',
        html_content,
        re.DOTALL | re.IGNORECASE
    )
    if not qa_pairs:
        return html_content
    def clean_text(t):
        return re.sub(r'<[^>]+>', '', t).strip()
    main_entity = []
    for q, a in qa_pairs:
        q_clean = clean_text(q)
        a_clean = clean_text(a)
        if q_clean and a_clean:
            main_entity.append({
                "@type": "Question",
                "name": q_clean,
                "acceptedAnswer": {
                    "@type": "Answer",
                    "text": a_clean
                }
            })
    if not main_entity:
        return html_content
    if '"@type": "FAQPage"' in html_content or '"@type":"FAQPage"' in html_content:
        return html_content
    schema_dict = {
        "@context": "https://schema.org",
        "@type": "FAQPage",
        "mainEntity": main_entity
    }
    json_ld_script = (
        '\n\n<!-- Background JSON-LD Schema for Googlebot -->\n'
        '<script type="application/ld+json">\n' +
        json.dumps(schema_dict, indent=2, ensure_ascii=False) +
        '\n</script>'
    )
    return html_content + json_ld_script

def format_text_to_html(text, sitemap_urls=None, image_names=None, status_container=None, model=MODEL):
    """Structure text into HTML via Groq API."""
    client = Groq(api_key=GROQ_API_KEY)
    word_count = len(text.split())

    if word_count <= 1500:
        if status_container:
            status_container.info("Processing article...")
        final_html = call_groq(client, text, sitemap_urls, image_names, model=model, is_last_chunk=True)
        return inject_faq_json_ld(remove_hallucinated_headings(final_html, text))

    chunks = split_into_chunks(text, max_words=800)
    html_parts = []
    total = len(chunks)

    if status_container:
        status_container.info(
            f"Article is {word_count:,} words. "
            f"Splitting into {total} sections for processing..."
        )

    progress_bar = None
    if status_container:
        progress_bar = status_container.progress(0, text=f"Section 1 of {total}...")

    for i, chunk in enumerate(chunks):
        if i > 0:
            delay = 25
            if status_container:
                countdown = status_container.empty()
                for sec in range(delay, 0, -1):
                    countdown.info(
                        f"Rate-limit cooldown: **{sec}s** before section {i + 1}..."
                    )
                    time.sleep(1)
                countdown.empty()
            else:
                time.sleep(delay)

        if progress_bar:
            progress_bar.progress(
                (i + 0.5) / total,
                text=f"Processing section {i + 1} of {total}..."
            )

        is_chunk = (i > 0)
        urls = sitemap_urls if (i == total - 1) else None
        
        part_html = call_groq(client, chunk, urls, image_names, is_chunk=is_chunk, model=model, is_last_chunk=(i == total - 1))
        html_parts.append(part_html)

        if progress_bar:
            progress_bar.progress(
                (i + 1) / total,
                text=f"Section {i + 1} of {total} complete"
            )

    final_html = '\n\n'.join(html_parts)
    return inject_faq_json_ld(remove_hallucinated_headings(final_html, text))

def generate_faq_summary(text, status_container=None, model=MODEL):
    """Generate SEO Summary and FAQs based on the provided text using the user's prompt."""
    client = Groq(api_key=GROQ_API_KEY)
    
    if status_container:
        status_container.info("🧠 Generating SEO Summary and FAQs...")
        
    user_msg = (
        "Generate a 120-150 word SEO category summary and 3-5 FAQs based on the article text. Output format:\n\n"
        "### SEO Summary\n"
        "<p>Your summary paragraph here...</p>\n\n"
        "### FAQs\n"
        '<!-- Frontend Visible FAQ Section -->\n'
        '<div class="schema-faq wp-block-yoast-seo-faq-block">\n'
        '  <div class="schema-faq-section" id="faq-question-1">\n'
        '    <strong class="schema-faq-question">Question 1 here?</strong>\n'
        '    <p class="schema-faq-answer">Answer 1 here.</p>\n'
        '  </div>\n'
        '  <div class="schema-faq-section" id="faq-question-2">\n'
        '    <strong class="schema-faq-question">Question 2 here?</strong>\n'
        '    <p class="schema-faq-answer">Answer 2 here.</p>\n'
        '  </div>\n'
        '</div>\n\n'
        f"Article text:\n{text}"
    )

    models_to_try = [model]
    for m in FALLBACK_MODELS:
        if m not in models_to_try:
            models_to_try.append(m)

    last_error = None
    for current_model in models_to_try:
        max_retries = 2
        for attempt in range(max_retries):
            try:
                completion = client.chat.completions.create(
                    model=current_model,
                    messages=[{"role": "user", "content": user_msg}],
                    temperature=0.3,
                    max_tokens=2048,
                )
                result_text = completion.choices[0].message.content.strip()
                return inject_faq_json_ld(result_text)
            except Exception as e:
                error_str = str(e)
                last_error = e
                if ("rate_limit" in error_str or "429" in error_str or "413" in error_str or "model_decommissioned" in error_str or "400" in error_str):
                    break
                else:
                    raise
                    
    raise Exception(f"Rate limits reached across all available AI models (or models decommissioned). Please wait a few minutes for your token quota to reset, or upgrade your API key. Last error: {last_error}")

def fetch_sitemap_urls(sitemap_url):
    """Fetch and parse XML sitemap, return a list of URLs."""
    req = urllib.request.Request(sitemap_url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req) as response:
            sitemap_xml = response.read()
            root = ET.fromstring(sitemap_xml)
            urls = []
            for child in root:
                for subchild in child:
                    if 'loc' in subchild.tag:
                        urls.append(subchild.text)
            return urls
    except Exception as e:
        st.error(f"Error fetching sitemap: {e}")
        return []

def get_windows_font(bold=True):
    font_dir = r"C:\Windows\Fonts"
    candidates_bold = ["segoeuib.ttf", "arialbd.ttf", "trebucbd.ttf", "calibrib.ttf", "verdana.ttf"]
    candidates_reg = ["segoeui.ttf", "arial.ttf", "trebuc.ttf", "calibri.ttf", "verdana.ttf"]
    for f in (candidates_bold if bold else candidates_reg):
        p = os.path.join(font_dir, f)
        if os.path.exists(p):
            return p
    return None

import math

def parse_title_for_typography(client, title, model=MODEL):
    """Use AI to intelligently break down a blog title into Badge, Main Words, Subtitle, and 4 Bottom Card Labels."""
    try:
        prompt = (
            "You are a magazine typography layout designer. Parse this blog article title into visual components for a tech banner.\n"
            "Rules:\n"
            "1. 'badge': A short 1-2 word tag like 'TOP 30', 'GUIDE', 'WHAT IS', 'EXPLAINED', 'INTERVIEW', or '' if not needed.\n"
            "2. 'main': The 2-4 most impactful keywords (e.g. 'LANGCHAIN INTERVIEW', 'GROSS SALARY', 'AI AUTOMATION'). Keep uppercase.\n"
            "3. 'sub': The descriptive subtitle (e.g. 'QUESTIONS & ANSWERS', 'MEANING, COMPONENTS & CALCULATION').\n"
            "4. 'cards': An array of exactly 4 short uppercase labels for bottom feature cards (each 1-2 words max, use \\n to split lines if 2 words, e.g. [\"BASIC\\nSALARY\", \"HRA &\\nALLOWANCES\", \"DEDUCTIONS\\n& PF\", \"NET TAKE\\nHOME\"]).\n"
            "5. 'icons': An array of exactly 4 short icon keywords matching each card (choose ONLY from: 'doc', 'chart', 'gear', 'check', 'brain', 'globe', 'person', 'mail', 'trophy', 'lock', 'code', 'link'). For salary/finance/automation topics, prefer ['doc', 'chart', 'gear', 'check'].\n"
            "Return ONLY valid JSON format: {\"badge\": \"...\", \"main\": \"...\", \"sub\": \"...\", \"cards\": [...], \"icons\": [...]}"
        )
        models_to_try = [model]
        for m in FALLBACK_MODELS:
            if m not in models_to_try:
                models_to_try.append(m)
                
        res = None
        for current_model in models_to_try:
            try:
                res = client.chat.completions.create(
                    model=current_model,
                    messages=[
                        {"role": "system", "content": prompt},
                        {"role": "user", "content": f"Title: {title}"}
                    ],
                    temperature=0.1,
                    max_tokens=250,
                    response_format={"type": "json_object"}
                )
                break
            except Exception as e:
                error_str = str(e).lower()
                if ("rate_limit" in error_str or "429" in error_str or "413" in error_str or "model_decommissioned" in error_str or "400" in error_str):
                    continue
                raise e
        if not res:
            raise Exception("Rate limit reached across all models in typography parser.")
        data = json.loads(res.choices[0].message.content)
        cards = data.get("cards", ["BASIC\nCONCEPTS", "CORE\nFEATURES", "REAL WORLD\nSCENARIOS", "KEY\nBENEFITS"])
        icons = data.get("icons", ["doc", "chart", "gear", "check"])
        return data.get("badge", ""), data.get("main", title.upper()[:20]), data.get("sub", ""), cards[:4], icons[:4]
    except Exception:
        words = title.upper().split()
        main = " ".join(words[:3]) if len(words) > 3 else " ".join(words)
        sub = " ".join(words[3:]) if len(words) > 3 else ""
        return "GUIDE", main, sub, ["BASIC\nCONCEPTS", "CORE\nFEATURES", "REAL WORLD\nSCENARIOS", "KEY\nBENEFITS"], ["doc", "chart", "gear", "check"]

def apply_vector_typography(img, badge_text, main_text, sub_text, cards, icons, palette_theme, width, height):
    """Draw ultra-crisp vector typography and 4 bottom glowing icon cards (100% match to user's design template)."""
    local_font_dir = os.path.join(os.path.dirname(__file__), "fonts")
    windows_font_dir = r"C:\Windows\Fonts"
    
    def get_font(names, default_size):
        for f in names:
            p_local = os.path.join(local_font_dir, f)
            p_win = os.path.join(windows_font_dir, f)
            
            if os.path.exists(p_local):
                try:
                    return ImageFont.truetype(p_local, default_size)
                except Exception:
                    pass
            elif os.path.exists(p_win):
                try:
                    return ImageFont.truetype(p_win, default_size)
                except Exception:
                    pass
        return ImageFont.load_default()
    
    # 1. Smooth cosine gradient fade from x=0 to x=68% of width (Guarantees zero vertical partition line!)
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw_ov = ImageDraw.Draw(overlay)
    fade_end = int(width * 0.68)
    for x in range(fade_end):
        factor = (1.0 + math.cos(math.pi * (x / float(fade_end)))) / 2.0
        alpha = int(240 * (factor ** 0.85))
        draw_ov.line([(x, 0), (x, height)], fill=(6, 12, 30, alpha))
        
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    draw = ImageDraw.Draw(img)
    
    # 2. Determine color accents based on theme
    accent_color = (0, 225, 255) # Electric Cyan default
    if "Orange" in palette_theme: accent_color = (255, 140, 0)
    elif "Green" in palette_theme: accent_color = (0, 255, 160)
    elif "Purple" in palette_theme: accent_color = (180, 100, 255)
    elif "White" in palette_theme: accent_color = (60, 100, 240)
    
    scale = width / 1000.0
    margin_x = int(48 * scale)
    curr_y = int(45 * scale)
    
    # Load Heavy Magazine Fonts matching Reference Banners (Impact/Arial Black)
    f_badge = get_font(["trebucbd.ttf", "arialbd.ttf", "segoeuib.ttf"], int(22 * scale))
    f_main_1 = get_font(["ariblk.ttf", "impact.ttf", "trebucbd.ttf", "arialbd.ttf"], int(68 * scale)) # Huge bold white
    f_main_2 = get_font(["ariblk.ttf", "impact.ttf", "trebucbd.ttf", "arialbd.ttf"], int(72 * scale)) # Huge bold cyan
    f_sub = get_font(["segoeuib.ttf", "trebucbd.ttf", "arialbd.ttf"], int(26 * scale))
    f_card = get_font(["segoeuib.ttf", "arialbd.ttf"], int(13 * scale))
        
    # 3. Draw Top Badge
    if badge_text and len(badge_text.strip()) > 0:
        bbox = draw.textbbox((0, 0), badge_text.strip(), font=f_badge)
        bw = bbox[2] - bbox[0] + int(28 * scale)
        bh = bbox[3] - bbox[1] + int(16 * scale)
        draw.rounded_rectangle([margin_x, curr_y, margin_x + bw, curr_y + bh], radius=int(8 * scale), fill=(0, 210, 140))
        draw.text((margin_x + int(14 * scale), curr_y + int(4 * scale)), badge_text.strip(), font=f_badge, fill=(10, 15, 35))
        curr_y += bh + int(22 * scale)
    else:
        curr_y += int(15 * scale)
        
    # 4. Draw Main Words (Two distinct lines with high contrast, exactly like 'LangChain INTERVIEW')
    main_words = main_text.split()
    lines = []
    curr_line = []
    for w in main_words:
        curr_line.append(w)
        if len(" ".join(curr_line)) > 11:
            lines.append(" ".join(curr_line))
            curr_line = []
    if curr_line:
        lines.append(" ".join(curr_line))
        
    for i, line in enumerate(lines[:2]):
        font_use = f_main_1 if i == 0 else f_main_2
        color = (255, 255, 255) if i == 0 else accent_color
        draw.text((margin_x, curr_y), line, font=font_use, fill=color)
        curr_y += int(74 * scale) if i == 0 else int(84 * scale)
        
    curr_y += int(10 * scale)
    
    # 5. Draw Subtext
    if sub_text and len(sub_text.strip()) > 0:
        sub_words = sub_text.split()
        sub_lines = []
        s_line = []
        for w in sub_words:
            s_line.append(w)
            if len(" ".join(s_line)) > 24:
                sub_lines.append(" ".join(s_line))
                s_line = []
        if s_line:
            sub_lines.append(" ".join(s_line))
            
        for idx, s_str in enumerate(sub_lines[:2]):
            scolor = (0, 255, 170) if idx == 1 else (230, 235, 250)
            draw.text((margin_x, curr_y), s_str, font=f_sub, fill=scolor)
            curr_y += int(32 * scale)
            
    # 6. Draw Accent Line
    curr_y += int(15 * scale)
    draw.rectangle([margin_x, curr_y, margin_x + int(150 * scale), curr_y + int(4 * scale)], fill=accent_color)
    
    # 7. Draw The 4 Bottom Glowing Icon Cards with PRO VECTOR ART ICONS
    card_w = int(112 * scale)
    card_h = int(126 * scale)
    gap = int(14 * scale)
    start_x = margin_x
    start_y = int(height * 0.73)
    
    for i in range(min(4, len(cards))):
        cx = start_x + i * (card_w + gap)
        cy = start_y
        c_label = cards[i]
        icon_type = str(icons[i]).lower() if i < len(icons) else "check"
        
        # Card box & glowing neon border
        draw.rounded_rectangle([cx, cy, cx + card_w, cy + card_h], radius=int(12 * scale), fill=(12, 20, 48), outline=accent_color, width=max(1, int(2 * scale)))
        
        # Inner icon circle
        ic_x, ic_y = cx + card_w // 2, cy + int(34 * scale)
        cr = int(22 * scale)
        draw.ellipse([ic_x - cr, ic_y - cr, ic_x + cr, ic_y + cr], fill=(18, 32, 72), outline=accent_color, width=1)
        
        # Professional Vector Icon Drawing (100% agency quality, zero text symbols!)
        if "doc" in icon_type or "file" in icon_type:
            draw.rounded_rectangle([ic_x - int(10 * scale), ic_y - int(12 * scale), ic_x + int(10 * scale), ic_y + int(12 * scale)], radius=int(2 * scale), outline=accent_color, width=max(1, int(2 * scale)))
            draw.line([(ic_x - int(5 * scale), ic_y - int(5 * scale)), (ic_x + int(5 * scale), ic_y - int(5 * scale))], fill=accent_color, width=max(1, int(2 * scale)))
            draw.line([(ic_x - int(5 * scale), ic_y), (ic_x + int(5 * scale), ic_y)], fill=accent_color, width=max(1, int(2 * scale)))
            draw.line([(ic_x - int(5 * scale), ic_y + int(5 * scale)), (ic_x + int(2 * scale), ic_y + int(5 * scale))], fill=accent_color, width=max(1, int(2 * scale)))
        elif "chart" in icon_type or "graph" in icon_type or "stat" in icon_type:
            draw.rectangle([ic_x - int(10 * scale), ic_y + int(2 * scale), ic_x - int(5 * scale), ic_y + int(12 * scale)], fill=accent_color)
            draw.rectangle([ic_x - int(2 * scale), ic_y - int(4 * scale), ic_x + int(3 * scale), ic_y + int(12 * scale)], fill=accent_color)
            draw.rectangle([ic_x + int(6 * scale), ic_y - int(10 * scale), ic_x + int(11 * scale), ic_y + int(12 * scale)], fill=(0, 255, 170))
            draw.line([(ic_x - int(12 * scale), ic_y + int(12 * scale)), (ic_x + int(14 * scale), ic_y + int(12 * scale))], fill=(255, 255, 255), width=max(1, int(2 * scale)))
        elif "gear" in icon_type or "cog" in icon_type or "calc" in icon_type or "pie" in icon_type:
            draw.ellipse([ic_x - int(12 * scale), ic_y - int(12 * scale), ic_x + int(12 * scale), ic_y + int(12 * scale)], outline=accent_color, width=max(1, int(2 * scale)))
            draw.pieslice([ic_x - int(12 * scale), ic_y - int(12 * scale), ic_x + int(12 * scale), ic_y + int(12 * scale)], start=270, end=360, fill=(0, 255, 170))
            draw.ellipse([ic_x - int(4 * scale), ic_y - int(4 * scale), ic_x + int(4 * scale), ic_y + int(4 * scale)], fill=(18, 32, 72), outline=accent_color, width=1)
        elif "brain" in icon_type or "llm" in icon_type or "ai" in icon_type:
            draw.ellipse([ic_x - int(11 * scale), ic_y - int(10 * scale), ic_x - int(1 * scale), ic_y + int(10 * scale)], outline=accent_color, width=max(1, int(2 * scale)))
            draw.ellipse([ic_x + int(1 * scale), ic_y - int(10 * scale), ic_x + int(11 * scale), ic_y + int(10 * scale)], outline=(0, 255, 170), width=max(1, int(2 * scale)))
            draw.ellipse([ic_x - int(3 * scale), ic_y - int(3 * scale), ic_x + int(3 * scale), ic_y + int(3 * scale)], fill=accent_color)
        else: # check / shield / trophy / default
            draw.polygon([(ic_x, ic_y - int(14 * scale)), (ic_x + int(12 * scale), ic_y - int(6 * scale)), (ic_x + int(9 * scale), ic_y + int(8 * scale)), (ic_x, ic_y + int(14 * scale)), (ic_x - int(9 * scale), ic_y + int(8 * scale)), (ic_x - int(12 * scale), ic_y - int(6 * scale))], outline=accent_color, width=max(1, int(2 * scale)))
            pts = [(ic_x - int(6 * scale), ic_y - int(1 * scale)), (ic_x - int(1 * scale), ic_y + int(5 * scale)), (ic_x + int(7 * scale), ic_y - int(5 * scale))]
            draw.line(pts, fill=(0, 255, 170), width=max(2, int(3 * scale)))
        
        # Draw 2 lines of uppercase text
        clines = str(c_label).replace("\\n", "\n").split("\n")
        ty = cy + int(66 * scale)
        for cline in clines[:2]:
            tbox = draw.textbbox((0, 0), cline.strip(), font=f_card)
            tw = tbox[2] - tbox[0]
            draw.text((cx + (card_w - tw) // 2, ty), cline.strip(), font=f_card, fill=(230, 240, 255))
            ty += int(18 * scale)
            
        # Connecting neon dot between cards
        if i < 3:
            dot_x = cx + card_w + gap // 2
            dot_y = cy + card_h // 2
            dr = max(2, int(3 * scale))
            draw.ellipse([dot_x - dr, dot_y - dr, dot_x + dr, dot_y + dr], fill=accent_color)
            
    return img

def generate_blog_banner(title, context, topic, layout_choice, palette_choice, dim_choice="1000x667 (3:2 Standard Blog Banner)", typo_mode="Hybrid AI + Vector Typography", model=MODEL):
    """Generate an editorial blog hero banner adhering to skills.md with exact Lanczos resizing and zero-gibberish typography."""
    client = Groq(api_key=GROQ_API_KEY)
    
    # Parse requested target dimensions
    if "1000x667" in dim_choice:
        target_w, target_h = 1000, 667
        req_w, req_h = 1500, 1000
    elif "1000x750" in dim_choice:
        target_w, target_h = 1000, 750
        req_w, req_h = 1400, 1050
    else:
        target_w, target_h = 1200, 630
        req_w, req_h = 1600, 840

    # 1. Layout selection (Rule: Never reuse exact same twice in a row)
    all_layouts = [
        "Layout A: Large title left, Hero illustration right",
        "Layout B: Illustration left, Title right",
        "Layout C: Centered title, Visual elements around it",
        "Layout D: Split screen comparison",
        "Layout E: Diagonal composition",
        "Layout F: Large background illustration, Small title overlay"
    ]
    if layout_choice.startswith("Auto-Rotate"):
        available = [l for l in all_layouts if l != st.session_state.get("last_banner_layout", "")]
        selected_layout = random.choice(available if available else all_layouts)
        st.session_state.last_banner_layout = selected_layout
    else:
        selected_layout = layout_choice
        st.session_state.last_banner_layout = layout_choice

    # 2. Palette selection
    all_palettes = [
        "Deep Navy Blue and Vibrant Teal gradients",
        "Royal Purple and Violet gradients",
        "Cyber Black and Electric Cyan",
        "Modern Blue SaaS and Silver",
        "Energetic Orange and Warm Mesh",
        "Emerald Green and Tech Grid",
        "Apple White Minimalist with subtle gray shadow"
    ]
    if palette_choice.startswith("Auto-Rotate"):
        selected_palette = random.choice(all_palettes)
    else:
        selected_palette = palette_choice

    # 3. Build Art Director system prompt based on Typography Mode
    is_hybrid = "Hybrid AI" in typo_mode
    if is_hybrid:
        system_prompt = (
            "You are an elite art director for top technical publications (Medium, Towards Data Science, Google Developers Blog, AWS Blog). "
            "Your task is to write a precise AI image generation prompt for a text-free Blog Hero Featured Illustration.\n\n"
            "MANDATORY DESIGN RULES FROM SKILLS.MD & REFERENCE TEMPLATE:\n"
            "1. Overall Style: Award-winning photorealistic 8k DSLR magazine photography of a modern tech corporate office, cinematic studio lighting, shallow depth of field with dark office bokeh in the background. NO cartoon, NO 2D flat illustration, NO clipart, NO graphical representation.\n"
            "2. Composition & Photorealistic Human Subject: On the RIGHT SIDE of the composition ONLY, place a 100% photorealistic, lifelike handsome 25-35 year old professional (e.g. software engineer, financial manager, or analyst) wearing smart casual or dark corporate attire, sitting at a modern wooden desk, smiling warmly with a confident expression. On the desk in front of them is a sleek laptop, a dark ceramic coffee mug, and a small plant.\n"
            "3. Upper-Right Floating Infographic: Floating in the air above and behind the laptop on the upper right side is a glowing neon cyan and blue holographic infographic diagram box / flowchart displaying nodes relevant to the topic (e.g. if LangChain: Document Loaders -> Vector Store -> LLM -> Chains; if Salary/Finance: Basic Salary -> Allowances -> Deductions -> Net Take Home; if AI/Prompting: Prompt -> LLM -> Output).\n"
            "4. Left Half Background: The left half of the composition MUST be deep dark navy blue and completely uncluttered, with ample clean dark space for editorial typography overlay. NO VERTICAL DIVIDING LINES, NO SPLIT SCREEN, NO WORDS, NO TEXT AT ALL IN THE IMAGE.\n"
            f"5. Color Palette: MUST use this color theme: '{selected_palette}'.\n\n"
            "Output ONLY the exact English image generation prompt without any introductory text, quotes, or markdown formatting."
        )
    else:
        system_prompt = (
            "You are an elite art director for top technical publications. Write a precise AI image generation prompt for a Blog Hero Featured Banner.\n\n"
            "MANDATORY DESIGN RULES:\n"
            "1. Overall Style: Award-winning photorealistic DSLR photography, modern corporate tech aesthetics, cinematic lighting, bokeh.\n"
            f"2. Layout: MUST follow this exact layout: '{selected_layout}'.\n"
            f"3. Color Palette: MUST use this color theme: '{selected_palette}'.\n"
            "4. Typography & Text Density: Show ONLY a short title in large bold white quotes (max 3 words).\n"
            "5. Photorealistic Real Humans: People MUST BE 100% PHOTOREALISTIC, LIFELIKE REAL HUMANS (DSLR photograph quality). NO cartoons, NO flat illustrations.\n"
            f"6. Topic Visuals for '{topic}': Include appropriate modern tech holographic elements.\n\n"
            "Output ONLY the exact English image generation prompt without any introductory text, quotes, or markdown formatting."
        )
    
    user_prompt = f"Article Title: {title}\nOptional Context: {context}\nTopic Category: {topic}\n\nWrite the AI image generation prompt now."
    
    models_to_try = [model]
    for m in FALLBACK_MODELS:
        if m not in models_to_try:
            models_to_try.append(m)
            
    res = None
    last_error = ""
    for current_model in models_to_try:
        try:
            res = client.chat.completions.create(
                model=current_model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.3,
                max_tokens=300
            )
            if current_model != model:
                print(f"Warning: Fell back to {current_model} due to rate limits in generate_blog_banner.")
            break
        except Exception as e:
            last_error = str(e)
            error_str = last_error.lower()
            if ("rate_limit" in error_str or "429" in error_str or "413" in error_str or "model_decommissioned" in error_str or "400" in error_str):
                continue
            raise e
            
    if not res:
        raise Exception(f"Rate limits reached across all available AI models. Please wait 3-5 minutes or switch to llama-3.1-8b-instant in Settings. Error: {last_error}")
    
    refined_prompt = strip_markdown_fences(res.choices[0].message.content).strip()
    
    # 4. Generate image via free high-res Flux endpoint at 1.5x requested resolution
    encoded_prompt = urllib.parse.quote(refined_prompt + ", magazine cover editorial blog featured image, 8k web resolution, award winning quality")
    image_url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width={req_w}&height={req_h}&nologo=true&model=flux"
    
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    
    req = urllib.request.Request(image_url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, context=ctx, timeout=45) as response:
        img_bytes = response.read()
        
    # 5. Process & EXACT Lanczos Resizing to Requested Dimensions
    img = Image.open(io.BytesIO(img_bytes))
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    img = img.resize((target_w, target_h), Image.Resampling.LANCZOS)
    
    # 6. Apply Vector Typography Overlay if Hybrid Mode selected (0% Gibberish Guarantee!)
    if is_hybrid:
        badge, main_txt, sub_txt, cards, symbols = parse_title_for_typography(client, title, model=model)
        img = apply_vector_typography(img, badge, main_txt, sub_txt, cards, symbols, selected_palette, target_w, target_h)
        
    out_buffer = io.BytesIO()
    img.save(out_buffer, format="WEBP", quality=88)
    webp_bytes = out_buffer.getvalue()
    
    # 7. Generate SEO space-separated filename
    clean_title = re.sub(r'[^a-zA-Z0-9. ]', '', title.lower()).replace("-", " ").replace("_", " ")
    clean_title = re.sub(r'\s+', ' ', clean_title).strip()
    words = clean_title.split()[:6]
    fname = " ".join(words) + f" {target_w}x{target_h} blog banner.webp"
    if len(fname) < 15:
        fname = f"featured blog hero banner {target_w}x{target_h}.webp"
        
    return webp_bytes, refined_prompt, fname
# --- Page Config ---
st.set_page_config(
    page_title="WordPress SEO Toolkit",
    page_icon="🛠️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Initialize Session State ---
if "structurer_text" not in st.session_state:
    st.session_state.structurer_text = ""
if "faq_text" not in st.session_state:
    st.session_state.faq_text = ""
if "saved_sitemap_urls" not in st.session_state:
    st.session_state.saved_sitemap_urls = ""
if "structurer_result" not in st.session_state:
    st.session_state.structurer_result = None
if "faq_result" not in st.session_state:
    st.session_state.faq_result = None
if "generated_webp_zip" not in st.session_state:
    st.session_state.generated_webp_zip = None
if "banner_title" not in st.session_state:
    st.session_state.banner_title = ""
if "banner_context" not in st.session_state:
    st.session_state.banner_context = ""
if "banner_result_img" not in st.session_state:
    st.session_state.banner_result_img = None
if "banner_result_prompt" not in st.session_state:
    st.session_state.banner_result_prompt = ""
if "banner_filename" not in st.session_state:
    st.session_state.banner_filename = "blog hero banner.webp"
if "last_banner_layout" not in st.session_state:
    st.session_state.last_banner_layout = ""

def update_structurer_text():
    st.session_state.structurer_text = st.session_state.temp_structurer_text
    st.session_state.structurer_result = None  # Clear previous result on input change
    st.session_state.generated_webp_zip = None

def update_faq_text():
    st.session_state.faq_text = st.session_state.temp_faq_text
    st.session_state.faq_result = None  # Clear previous result on input change

def update_sitemap_urls():
    st.session_state.saved_sitemap_urls = st.session_state.temp_sitemap_urls

# --- Custom CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    .stApp { font-family: 'Inter', sans-serif; }

    /* Header */
    .hero-header {
        padding: 2.5rem 2rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        text-align: center;
        background-color: var(--secondary-background-color);
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    }
    .hero-header h1 { font-size: 2.2rem; font-weight: 700; margin: 0 0 0.5rem 0; }
    .hero-header p { font-size: 1.05rem; margin: 0; opacity: 0.8; }

    /* Stats */
    .stats-bar { display: flex; gap: 1.5rem; justify-content: center; margin-bottom: 1.5rem; flex-wrap: wrap; }
    .stat-card {
        background-color: var(--secondary-background-color);
        border: 1px solid var(--primary-color);
        border-radius: 12px; padding: 1rem 1.5rem; text-align: center; min-width: 160px;
    }
    .stat-card .stat-value { font-size: 1.6rem; font-weight: 700; color: var(--primary-color); }
    .stat-card .stat-label {
        font-size: 0.78rem; opacity: 0.7;
        text-transform: uppercase; letter-spacing: 1px; margin-top: 0.2rem;
    }

    /* Artifact panel */
    .artifact-panel {
        border: 1px solid var(--primary-color); border-radius: 14px;
        overflow: hidden; margin-top: 1.5rem;
        box-shadow: 0 4px 24px rgba(0,0,0,0.08);
    }
    .artifact-toolbar {
        display: flex; align-items: center; justify-content: space-between;
        background: var(--secondary-background-color); border-bottom: 1px solid var(--primary-color); padding: 0.6rem 1.2rem;
    }
    .artifact-title {
        display: flex; align-items: center; gap: 0.6rem;
        font-size: 0.92rem; font-weight: 600;
    }
    .artifact-tag {
        background: var(--primary-color); color: var(--secondary-background-color); font-size: 0.72rem; font-weight: 500;
        padding: 0.15rem 0.5rem; border-radius: 4px; text-transform: uppercase;
    }

    /* HTML preview fix */
    .html-preview {
        background: var(--secondary-background-color);
        padding: 2.5rem 3rem;
        font-family: Georgia, 'Times New Roman', serif;
        font-size: 1.05rem; line-height: 1.8;
        border-radius: 8px;
        border: 1px solid rgba(128,128,128,0.2);
    }
    .html-preview h1 { font-size: 2rem; margin-bottom: 1rem; font-weight: 800; line-height: 1.25; }
    .html-preview h2 { font-size: 1.5rem; margin-top: 2rem; margin-bottom: 0.6rem; font-weight: 700; }
    .html-preview h3 { font-size: 1.2rem; margin-top: 1.5rem; margin-bottom: 0.4rem; font-weight: 600; }
    .html-preview h4 { font-size: 1.05rem; margin-top: 1.2rem; font-weight: 600; }
    .html-preview p { margin-bottom: 1rem; }
    .html-preview table { border-collapse: collapse; width: 100%; margin: 1.2rem 0; font-family: 'Inter', sans-serif; font-size: 0.92rem; }
    .html-preview th { background: rgba(128,128,128,0.1); padding: 0.75rem 1rem; text-align: left; border: 1px solid rgba(128,128,128,0.2); font-weight: 600; }
    .html-preview td { padding: 0.65rem 1rem; border: 1px solid rgba(128,128,128,0.2); }
    .html-preview ul, .html-preview ol { padding-left: 1.8rem; margin-bottom: 1rem; }
    .html-preview li { margin-bottom: 0.35rem; }
    .html-preview pre { background: rgba(128,128,128,0.05); border: 1px solid rgba(128,128,128,0.2); border-radius: 8px; padding: 1rem; overflow-x: auto; }
    .html-preview code { background: rgba(128,128,128,0.1); padding: 0.15rem 0.4rem; border-radius: 3px; font-size: 0.88em; }
    .html-preview pre code { background: none; padding: 0; }
    
    .powered-by { text-align: center; padding: 1rem; opacity: 0.5; font-size: 0.75rem; }
</style>
""", unsafe_allow_html=True)


# --- Sidebar Navigation ---
with st.sidebar:
    st.title("🛠️ Tools")
    app_mode = st.radio(
        "Select Tool:",
        ["Content Structurer", "FAQ Generator", "Blog Banner Generator", "Sitemap URL Save", "Meta Pixel Checker"],
        index=0
    )
    
    st.divider()
    st.markdown("### Settings")
    selected_model = st.selectbox(
        "AI Model",
        FALLBACK_MODELS,
        index=0,
        help="Switch models if you hit a rate limit."
    )
    
    st.divider()
    st.markdown("### Quick Links (Saved)")
    if st.session_state.saved_sitemap_urls:
        lines = len([l for l in st.session_state.saved_sitemap_urls.splitlines() if l.strip()])
        st.success(f"{lines} internal URLs loaded.")
    else:
        st.info("No sitemap URLs saved yet.")

    st.divider()
    st.markdown('<div class="powered-by">Powered by Groq</div>', unsafe_allow_html=True)


# --- Section 1: Content Structurer ---
if app_mode == "Content Structurer":
    st.markdown("""
    <div class="hero-header">
        <h1>WordPress SEO Structurer</h1>
        <p>Transform raw content into clean, semantic HTML for WordPress</p>
    </div>
    """, unsafe_allow_html=True)
    
    input_mode = st.radio("Input Method", ["Paste Text", "Upload .docx"], horizontal=True)
    
    raw_text = None
    webp_images = {}
    image_names = []
    
    if input_mode == "Paste Text":
        st.text_area(
            "Paste your raw article content below",
            value=st.session_state.structurer_text,
            placeholder="Paste your unformatted blog post / article text here...",
            height=300,
            key="temp_structurer_text",
            on_change=update_structurer_text
        )
        raw_text = st.session_state.structurer_text
    elif input_mode == "Upload .docx":
        st.info("If your Word document contains images, they will be automatically extracted, converted to WebP, and placed logically in the HTML.")
        uploaded_file = st.file_uploader("Upload a Word document", type=["docx"])
        if uploaded_file is not None:
            with st.spinner("Extracting text and images from document..."):
                raw_text, webp_images, image_names = extract_text_and_images_from_docx(uploaded_file.getvalue())
            
            st.success(f"Extracted **{len(raw_text.split())}** words and **{len(image_names)}** images from `{uploaded_file.name}`")
            with st.expander("Preview Extracted Text", expanded=False):
                st.text_area("Raw Text", raw_text, height=200, disabled=True)
                
            if webp_images:
                # Cache the zip so it's ready for download after generation
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for img_name, img_bytes in webp_images.items():
                        zip_file.writestr(img_name, img_bytes)
                st.session_state.generated_webp_zip = zip_buffer.getvalue()
                
    st.markdown("### Internal Links *(optional)*")
    sitemap_urls = st.text_area(
        "Sitemap / Internal URLs",
        value=st.session_state.saved_sitemap_urls,
        placeholder="Paste URLs for cross-linking (one per line) or use 'Sitemap URL Save' tool to fetch automatically.",
        height=140,
        key="temp_sitemap_urls",
        on_change=update_sitemap_urls
    )

    st.markdown("")
    generate_clicked = st.button(
        "🚀 Structure for WordPress SEO",
        type="primary",
        use_container_width=True,
        disabled=not (raw_text and raw_text.strip())
    )
    
    if generate_clicked and raw_text and raw_text.strip():
        word_count = len(raw_text.split())
        st.markdown(f"""
        <div class="stats-bar">
            <div class="stat-card">
                <div class="stat-value">{word_count:,}</div>
                <div class="stat-label">Words</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">{len(raw_text.splitlines()):,}</div>
                <div class="stat-label">Lines</div>
            </div>
            <div class="stat-card">
                <div class="stat-value">{selected_model.split('-')[-1].upper()}</div>
                <div class="stat-label">Model</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        start_time = time.time()
        status_area = st.container()

        try:
            formatted_html = format_text_to_html(raw_text, sitemap_urls, image_names, status_container=status_area, model=selected_model)
            elapsed = time.time() - start_time
            st.success(f"Structuring complete in **{elapsed:.1f}s**")
            st.session_state.structurer_result = formatted_html
        except Exception as e:
            st.error(f"API Error: {str(e)}")
            
    if st.session_state.structurer_result:
        formatted_html = st.session_state.structurer_result

        # 1. Visual Preview
        st.markdown("### Rendered Preview")
        st.markdown(f'<div class="html-preview">{formatted_html}</div>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        # 2. Download Buttons
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="📝 Download Output (.html)",
                data=formatted_html,
                file_name="structured_article.html",
                mime="text/html",
                use_container_width=True,
            )
        with col2:
            if st.session_state.generated_webp_zip:
                st.download_button(
                    label="📦 Download Extracted Images (.zip)",
                    data=st.session_state.generated_webp_zip,
                    file_name="extracted_images.zip",
                    mime="application/zip",
                    use_container_width=True,
                    type="primary"
                )
                
        st.markdown("---")
        
        # 3. Code Section
        st.markdown("### HTML Source Code")
        st.markdown(
            '<p style="margin:0 0 0.3rem 0;font-weight:600;font-size:0.9rem;">'
            'Copy HTML (click the icon at top-right of the code block below):</p>',
            unsafe_allow_html=True
        )
        st.code(formatted_html, language="html")

# --- Section 2: FAQ Generator ---
elif app_mode == "FAQ Generator":
    st.markdown("""
    <div class="hero-header">
        <h1>SEO FAQ & Summary Generator</h1>
        <p>Instantly generate SEO-optimized categories and FAQs</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.text_area(
        "Paste your article content below",
        value=st.session_state.faq_text,
        placeholder="Paste your unformatted blog post / article text here to generate FAQs...",
        height=300,
        key="temp_faq_text",
        on_change=update_faq_text
    )
    raw_text = st.session_state.faq_text
    
    st.markdown("")
    generate_faq_clicked = st.button(
        "❓ Generate SEO Summary & FAQs",
        type="primary",
        use_container_width=True,
        disabled=not (raw_text and raw_text.strip())
    )
    
    if generate_faq_clicked and raw_text and raw_text.strip():
        start_time = time.time()
        status_area = st.container()
        try:
            faq_result = generate_faq_summary(raw_text, status_container=status_area, model=selected_model)
            elapsed = time.time() - start_time
            st.success(f"Generation complete in **{elapsed:.1f}s**")
            st.session_state.faq_result = faq_result
        except Exception as e:
            st.error(f"API Error: {str(e)}")

    if st.session_state.faq_result:
        faq_result = st.session_state.faq_result
        st.markdown(
            '<p style="margin:1rem 0 0.3rem 0;font-weight:600;font-size:0.9rem;">'
            'Copy Result (click the icon at top-right of the block below):</p>',
            unsafe_allow_html=True
        )
        st.code(faq_result, language="markdown")
        
        st.markdown("---")
        st.markdown("#### 👁️ Rendered Preview")
        st.markdown(f'<div class="html-preview">{faq_result.replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)

# --- Section 3: Blog Banner Generator ---
elif app_mode == "Blog Banner Generator":
    st.markdown("""
    <div class="hero-header">
        <h1>🎨 Blog Hero Banner Generator</h1>
        <p>Create premium, modern magazine-style 4:3 featured images for your tech & career blogs</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 1], gap="large")
    with col1:
        st.markdown("### 📝 Article Details")
        banner_title = st.text_input(
            "Article Title / Heading*",
            value=st.session_state.banner_title,
            placeholder="e.g. Top 30 LangChain Interview Questions and Answers (2026)",
            help="The main title of your blog post."
        )
        banner_context = st.text_area(
            "Optional Context / Key Concepts",
            value=st.session_state.banner_context,
            placeholder="e.g. Covers RAG, agents, vector stores, prompt templates, and tool calling...",
            height=120,
            help="Provides extra details so the AI picks the perfect topic illustrations."
        )
        
        st.markdown("### ⚙️ Design & Layout Controls")
        topic_choice = st.selectbox(
            "Topic Category",
            [
                "Auto-Detect from Title",
                "LangChain & AI Agents",
                "Model Context Protocol (MCP)",
                "RAG & Vector Databases",
                "Interview Questions & Coding",
                "Internship & Career Growth",
                "Rate Limiting & API Security",
                "Data Science & Analytics",
                "General Tech / Cloud Architecture"
            ]
        )
        
        layout_choice = st.selectbox(
            "Layout Style (Skills.md Rule: Never reuse exact same twice in a row)",
            [
                "Auto-Rotate (Recommended)",
                "Layout A: Large title left, Hero illustration right",
                "Layout B: Illustration left, Title right",
                "Layout C: Centered title, Visual elements around it",
                "Layout D: Split screen comparison",
                "Layout E: Diagonal composition",
                "Layout F: Large background illustration, Small title overlay"
            ],
            index=0
        )
        
        palette_choice = st.selectbox(
            "Color Palette Theme",
            [
                "Auto-Rotate (Recommended)",
                "Deep Navy Blue & Vibrant Teal",
                "Royal Purple & Violet Gradients",
                "Cyber Black & Electric Cyan",
                "Modern Blue SaaS & Silver",
                "Energetic Orange & Warm Mesh",
                "Emerald Green & Tech Grid",
                "Apple White Minimalist"
            ],
            index=0
        )
        
        dim_choice = st.selectbox(
            "Target Dimensions & Aspect Ratio",
            [
                "1000x667 (3:2 Standard Blog Banner - Preferred)",
                "1000x750 (4:3 Magazine Cover)",
                "1200x630 (16:9 LinkedIn & Social Share)"
            ],
            index=0,
            help="Ensures exact Lanczos resizing to your required blog dimensions."
        )
        
        typo_mode = st.selectbox(
            "Typography & Quality Mode",
            [
                "Hybrid AI + Crisp Vector Typography (100% Zero-Gibberish Guarantee - Recommended)",
                "Pure AI Image Generation (AI model draws text)"
            ],
            index=0,
            help="Hybrid mode uses AI for pristine 3D illustrations and real TrueType vector fonts for razor-sharp title text without blurry alien spelling."
        )
        
        st.markdown("")
        gen_banner_btn = st.button(
            "🎨 Generate Blog Hero Banner",
            type="primary",
            use_container_width=True,
            disabled=not (banner_title and banner_title.strip())
        )
        
        if gen_banner_btn and banner_title and banner_title.strip():
            st.session_state.banner_title = banner_title
            st.session_state.banner_context = banner_context
            with st.spinner("✨ Engineering editorial prompt & generating banner via Flux AI + Vector Overlay..."):
                try:
                    img_bytes, prompt_used, fname = generate_blog_banner(
                        banner_title, banner_context, topic_choice, layout_choice, palette_choice, dim_choice, typo_mode, model=selected_model
                    )
                    st.session_state.banner_result_img = img_bytes
                    st.session_state.banner_result_prompt = prompt_used
                    st.session_state.banner_filename = fname
                    st.success("🎉 Banner generated & optimized for web (<250KB)!")
                except Exception as e:
                    st.error(f"Error generating banner: {e}")
                    
    with col2:
        st.markdown("### 🖼️ Banner Preview & Download")
        if st.session_state.banner_result_img:
            st.image(st.session_state.banner_result_img, caption=f"4:3 Featured Banner: {st.session_state.banner_filename}", use_container_width=True)
            
            st.download_button(
                label="⬇️ Download WEBP Banner (<250KB)",
                data=st.session_state.banner_result_img,
                file_name=st.session_state.banner_filename,
                mime="image/webp",
                type="primary",
                use_container_width=True
            )
            
            with st.expander("🔍 View AI Art Director Prompt (Skills.md Applied)", expanded=False):
                st.info("This exact prompt was crafted by Llama 3 to adhere 100% to your Blog Banner Image Creation Skills:")
                st.code(st.session_state.banner_result_prompt, language="text")
        else:
            st.markdown("""
            <div style="border: 2px dashed var(--primary-color); border-radius: 12px; padding: 4rem 2rem; text-align: center; opacity: 0.6; background: var(--secondary-background-color);">
                <h3 style="margin-bottom: 0.5rem;">No Banner Generated Yet</h3>
                <p>Enter your article title on the left and click Generate to create a magazine-quality 4:3 featured image.</p>
            </div>
            """, unsafe_allow_html=True)

# --- Section 4: Sitemap URL Save ---
elif app_mode == "Sitemap URL Save":
    st.markdown("""
    <div class="hero-header">
        <h1>Sitemap Fetcher</h1>
        <p>Automatically extract and save URLs for internal linking</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("Enter your post sitemap URL. The app will download all the links and automatically feed them into the Content Structurer's internal link field.")
    sitemap_url_input = st.text_input("Sitemap URL", placeholder="https://example.com/post-sitemap.xml")
    
    if st.button("Fetch & Save URLs", type="primary"):
        if sitemap_url_input:
            with st.spinner("Fetching sitemap..."):
                urls = fetch_sitemap_urls(sitemap_url_input)
            
            if urls:
                # Deduplicate and sort
                urls = sorted(list(set(urls)))
                st.success(f"Successfully fetched {len(urls)} URLs!")
                st.session_state.saved_sitemap_urls = "\n".join(urls)
                
                # Update the displayed text area automatically
                st.text_area("Extracted URLs (Saved to Memory)", value=st.session_state.saved_sitemap_urls, height=300, key="temp_sitemap_urls", on_change=update_sitemap_urls)
            else:
                st.warning("No URLs found or failed to parse. Make sure it's a direct XML sitemap with <loc> tags.")
        else:
            st.error("Please enter a URL first.")
    else:
        # Just show what's currently in memory
        st.text_area("Extracted URLs (Saved to Memory)", value=st.session_state.saved_sitemap_urls, height=300, key="temp_sitemap_urls", on_change=update_sitemap_urls)

# --- Section 5: Meta Pixel Checker ---
elif app_mode == "Meta Pixel Checker":
    st.markdown("""
    <div class="hero-header">
        <h1>🔍 Meta Pixel Checker</h1>
        <p>Preview how your SEO Title and Meta Description will look in Google Search Results.</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 1], gap="large")
    
    with col1:
        st.markdown("### ✍️ Enter SEO Metadata")
        seo_url = st.text_input("Page URL", value="https://www.example.com/your-post-url", help="The URL displayed above the title.")
        seo_title = st.text_area("SEO Title", value=st.session_state.get("banner_title", ""), height=80, help="Aim for 50-60 characters (Max ~600px).")
        seo_desc = st.text_area("Meta Description", value="", height=120, help="Aim for 150-160 characters (Max ~960px on Desktop, ~680px on Mobile).")
        
        # Simple character count for immediate feedback
        title_len = len(seo_title)
        desc_len = len(seo_desc)
        
        st.markdown(f"**Title Length:** {title_len} chars (Recommended: 50-60)")
        if title_len > 60:
            st.warning("⚠️ Title may be truncated by Google.")
            
        st.markdown(f"**Description Length:** {desc_len} chars (Recommended: 150-160)")
        if desc_len > 160:
            st.warning("⚠️ Description may be truncated by Google on Desktop.")
        if desc_len > 120:
            st.info("ℹ️ Description will likely be truncated on Mobile.")

    with col2:
        # Dynamically extract the site name from the URL
        site_name = "Your Website"
        if seo_url:
            try:
                clean_url = seo_url if seo_url.startswith('http') else 'http://' + seo_url
                domain = clean_url.split('//')[-1].split('/')[0].replace('www.', '')
                parsed_name = domain.split('.')[0].capitalize() if '.' in domain else domain.capitalize()
                if parsed_name:
                    site_name = parsed_name
            except Exception:
                pass

        st.markdown("### 👁️ Google SERP Preview")
        tab1, tab2 = st.tabs(["💻 Desktop View", "📱 Mobile View"])
        
        # Desktop HTML/CSS (Google Desktop: Title ~600px, Description ~600px 2-lines)
        desktop_html = f'''
        <div style="overflow-x: auto; width: 100%; padding-bottom: 10px;">
            <div style="font-family: Arial, sans-serif; width: 652px; min-width: 652px; padding: 16px; background: white; border-radius: 8px; border: 1px solid #dfe1e5; margin-bottom: 20px;">
                <div style="display: flex; align-items: center; margin-bottom: 8px;">
                    <div style="width: 28px; height: 28px; background-color: #f1f3f4; border-radius: 50%; margin-right: 12px; display: flex; align-items: center; justify-content: center;">
                        <span style="font-size: 14px;">🌐</span>
                    </div>
                    <div>
                        <div style="font-size: 14px; color: #202124; line-height: 1.3;">{site_name}</div>
                        <div style="font-size: 12px; color: #4d5156; line-height: 1.3; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 500px;">{seo_url}</div>
                    </div>
                </div>
                <div style="font-size: 20px; color: #1a0dab; line-height: 1.3; margin-bottom: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; width: 600px; min-width: 600px; font-weight: normal; cursor: pointer;">
                    {seo_title if seo_title else "Your SEO Title Goes Here"}
                </div>
                <div style="font-size: 14px; color: #4d5156; line-height: 1.58; width: 600px; min-width: 600px; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden;">
                    {seo_desc if seo_desc else "Your meta description will appear here. It should provide a concise and compelling summary of your page's content, encouraging users to click through to read more."}
                </div>
            </div>
        </div>
        '''
        
        # Mobile HTML/CSS (Google Mobile: Title ~2 lines, Description ~3 lines or ~120 chars)
        mobile_html = f'''
        <div style="overflow-x: auto; width: 100%; padding-bottom: 10px;">
            <div style="font-family: Arial, sans-serif; width: 375px; min-width: 375px; padding: 16px; background: white; border-radius: 8px; border: 1px solid #dfe1e5; box-shadow: 0 1px 3px rgba(0,0,0,0.12);">
                <div style="display: flex; align-items: center; margin-bottom: 10px;">
                    <div style="width: 28px; height: 28px; background-color: #f1f3f4; border-radius: 50%; margin-right: 10px; display: flex; align-items: center; justify-content: center;">
                        <span style="font-size: 14px;">🌐</span>
                    </div>
                    <div>
                        <div style="font-size: 14px; color: #202124; line-height: 1.2;">{site_name}</div>
                        <div style="font-size: 12px; color: #3c4043; line-height: 1.2; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; width: 280px; min-width: 280px;">{seo_url}</div>
                    </div>
                </div>
                <div style="font-size: 20px; color: #1a0dab; line-height: 1.3; margin-bottom: 6px; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; font-weight: normal;">
                    {seo_title if seo_title else "Your SEO Title Goes Here"}
                </div>
                <div style="font-size: 14px; color: #4d5156; line-height: 1.58; display: -webkit-box; -webkit-line-clamp: 3; -webkit-box-orient: vertical; overflow: hidden;">
                    {seo_desc if seo_desc else "Your meta description will appear here. It should provide a concise and compelling summary of your page's content, encouraging users to click through to read more."}
                </div>
            </div>
        </div>
        '''
        
        with tab1:
            components.html(desktop_html, height=200, scrolling=True)
            st.caption("Google truncates desktop titles at ~600px and descriptions at ~960px (displayed over 2 lines of ~600px wide text).")
            
        with tab2:
            components.html(mobile_html, height=250, scrolling=True)
            st.caption("Google truncates mobile titles after 2 lines (~600px total) and descriptions after 2 lines (~680px total).")

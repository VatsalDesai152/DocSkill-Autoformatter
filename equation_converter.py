"""
equation_converter.py - Convert LaTeX equations in DOCX to native Word OMML.

Detects paragraphs containing LaTeX math ($...$, $$...$$) and replaces
them with proper Word equation objects using the pipeline:
    LaTeX  -->  MathML  -->  OMML (via Microsoft's MML2OMML.XSL)
"""

import os
import re
from copy import deepcopy
from lxml import etree
from docx.oxml.ns import qn

try:
    import latex2mathml.converter
except ImportError:
    latex2mathml = None


# Regex patterns for LaTeX math
DISPLAY_MATH = re.compile(r'\$\$(.+?)\$\$', re.DOTALL)
INLINE_MATH = re.compile(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)')

# Word OMML namespace
OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def find_mml2omml_xslt():
    """Locate Microsoft's MML2OMML.XSL stylesheet on the system."""
    candidates = [
        os.path.join(os.environ.get('PROGRAMFILES', ''),
                     'Microsoft Office', 'root', 'Office16', 'MML2OMML.XSL'),
        os.path.join(os.environ.get('PROGRAMFILES(X86)', ''),
                     'Microsoft Office', 'root', 'Office16', 'MML2OMML.XSL'),
        os.path.join(os.environ.get('PROGRAMFILES', ''),
                     'Microsoft Office', 'root', 'Office15', 'MML2OMML.XSL'),
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return None


def _load_xslt():
    """Load and cache the MathML-to-OMML XSLT transform."""
    xslt_path = find_mml2omml_xslt()
    if xslt_path is None:
        return None
    xslt_doc = etree.parse(xslt_path)
    return etree.XSLT(xslt_doc)


# Load once at module level
_XSLT_TRANSFORM = _load_xslt()


def latex_to_omml(latex_str):
    """Convert a LaTeX math string to an OMML XML element.

    Pipeline: LaTeX -> MathML (via latex2mathml) -> OMML (via MS XSLT)
    """
    if latex2mathml is None:
        raise ImportError("latex2mathml is required. Install: pip install latex2mathml")
    if _XSLT_TRANSFORM is None:
        raise FileNotFoundError(
            "MML2OMML.XSL not found. Microsoft Office must be installed."
        )

    # Step 1: LaTeX -> MathML
    mathml_str = latex2mathml.converter.convert(latex_str)

    # Step 2: Parse MathML XML
    mathml_tree = etree.fromstring(mathml_str.encode('utf-8'))

    # Step 3: Apply XSLT to get OMML
    omml_tree = _XSLT_TRANSFORM(mathml_tree)
    return omml_tree.getroot()


def convert_equations_in_doc(doc):
    """Scan all paragraphs in a python-docx Document and convert
    LaTeX math notation to native Word OMML equations.

    Handles:
      - Display math: $$...$$  (entire paragraph is an equation)
      - Inline math:  $...$    (equation mixed with text)
    """
    if latex2mathml is None:
        print("  [WARN] latex2mathml not installed. Skipping equation conversion.")
        return 0
    if _XSLT_TRANSFORM is None:
        print("  [WARN] MML2OMML.XSL not found. Skipping equation conversion.")
        return 0

    converted = 0
    paragraphs = list(doc.paragraphs)  # snapshot to avoid mutation issues

    for para in paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Already has OMML? Skip.
        if para._element.findall(qn('m:oMath')) or para._element.findall(qn('m:oMathPara')):
            continue

        # --- Case 1: Entire paragraph is a display equation $$...$$ ---
        display_match = DISPLAY_MATH.fullmatch(text)
        if display_match:
            latex_str = display_match.group(1).strip()
            try:
                omml_element = latex_to_omml(latex_str)
                _replace_paragraph_with_omml(para, omml_element, display=True)
                converted += 1
            except Exception as e:
                print(f"  [WARN] Could not convert display equation: {e}")
            continue

        # --- Case 2: Paragraph contains inline $...$ equations ---
        if '$' in text:
            try:
                new_count = _convert_inline_math(para)
                converted += new_count
            except Exception as e:
                print(f"  [WARN] Could not convert inline equation: {e}")

    return converted


def _replace_paragraph_with_omml(para, omml_element, display=False):
    """Replace the content of a paragraph with an OMML equation element."""
    p_elem = para._element

    # Remove all existing runs (the LaTeX text)
    for child in list(p_elem):
        tag = child.tag
        if tag == qn('w:r') or tag == qn('w:hyperlink'):
            p_elem.remove(child)

    if display:
        # Wrap in oMathPara for centered display equations
        omath_para = etree.SubElement(p_elem, qn('m:oMathPara'))
        # The XSLT output is either an oMath element or has oMath children
        omath = _extract_omath(omml_element)
        if omath is not None:
            omath_para.append(omath)
    else:
        omath = _extract_omath(omml_element)
        if omath is not None:
            p_elem.append(omath)


def _extract_omath(omml_element):
    """Extract the m:oMath element from the XSLT output."""
    # The XSLT might return the oMath directly or wrapped
    if omml_element.tag == qn('m:oMath'):
        return omml_element
    # Search children
    omath = omml_element.find(qn('m:oMath'))
    if omath is not None:
        return omath
    # Try with local name
    for child in omml_element.iter():
        if child.tag.endswith('}oMath') or child.tag == 'oMath':
            return child
    # Last resort: return the element itself
    return omml_element


def _convert_inline_math(para):
    """Convert inline $...$ in a paragraph. Returns count of conversions."""
    full_text = para.text or ''
    for run in para.runs:
        pass  # just ensuring runs are accessible

    # Get the full paragraph text
    full_text = ''.join(run.text for run in para.runs) if para.runs else para.text or ''

    # Find all inline math segments
    parts = []
    last_end = 0
    count = 0

    for match in INLINE_MATH.finditer(full_text):
        # Text before the match
        if match.start() > last_end:
            parts.append(('text', full_text[last_end:match.start()]))
        # The math
        latex_str = match.group(1).strip()
        parts.append(('math', latex_str))
        count += 1
        last_end = match.end()

    # Text after last match
    if last_end < len(full_text):
        parts.append(('text', full_text[last_end:]))

    if count == 0:
        return 0

    # Rebuild paragraph content
    p_elem = para._element

    # Preserve paragraph properties
    pPr = p_elem.find(qn('w:pPr'))

    # Remove all existing runs
    for child in list(p_elem):
        if child.tag != qn('w:pPr'):
            p_elem.remove(child)

    # Add new content
    for part_type, content in parts:
        if part_type == 'text':
            run_elem = etree.SubElement(p_elem, qn('w:r'))
            # Copy run properties for Times New Roman
            rPr = etree.SubElement(run_elem, qn('w:rPr'))
            rFonts = etree.SubElement(rPr, qn('w:rFonts'))
            rFonts.set(qn('w:ascii'), 'Times New Roman')
            rFonts.set(qn('w:hAnsi'), 'Times New Roman')
            sz = etree.SubElement(rPr, qn('w:sz'))
            sz.set(qn('w:val'), '22')  # 11pt = 22 half-points
            t_elem = etree.SubElement(run_elem, qn('w:t'))
            t_elem.text = content
            t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        else:
            try:
                omml_element = latex_to_omml(content)
                omath = _extract_omath(omml_element)
                if omath is not None:
                    p_elem.append(omath)
            except Exception as e:
                # Fallback: keep as plain text
                run_elem = etree.SubElement(p_elem, qn('w:r'))
                t_elem = etree.SubElement(run_elem, qn('w:t'))
                t_elem.text = f'${content}$'

    return count

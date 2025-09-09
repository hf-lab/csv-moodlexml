#!/usr/bin/env python3
"""
Excel → Moodle XML (MCQ) converter
- 5 choices per question
- Per-choice percent correctness (Moodle 'fraction' values)
- General feedback
- Question text supports: **bold**, *italic*, bullet (- ) and numbered lists (1. / 1) ), line breaks [br] or \n

Usage:
    python convert.py --in input.xlsx --out output.xml --category "Sample Category"
"""

import argparse
import re
import sys
from xml.etree.ElementTree import Element, SubElement, ElementTree
from html import escape
import pandas as pd
from pathlib import Path

# ---------- Formatting helpers ----------

_ctrl_re = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")

def strip_control(s: str) -> str:
    if s is None:
        return ""
    return _ctrl_re.sub("", str(s))

def convert_markup_to_html(text: str) -> str:
    """
    Convert lightweight markup in question text to HTML suitable for Moodle:
      - **bold**  -> <strong>…</strong>
      - *italic*  -> <em>…</em>
      - Bullet list lines starting with "- " -> <ul><li>…</li></ul>
      - Numbered list lines starting with "1. " or "1) " -> <ol><li>…</li></ol>
      - [br] or newline → <br/>
    """
    if text is None:
        return ""
    # sanitize & normalize
    t = strip_control(text).replace("\r\n", "\n").replace("\r", "\n")
    t = t.replace("[br]", "\n")

    # Escape HTML first, then apply lightweight markup
    t = escape(t)

    # Bold (**...**)
    t = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", t)
    # Italic (*...*)
    t = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"<em>\1</em>", t)

    # Lists
    lines = t.split("\n")
    html_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]

        if re.match(r"^\s*-\s+", line):
            items = []
            while i < len(lines) and re.match(r"^\s*-\s+", lines[i]):
                items.append(re.sub(r"^\s*-\s+", "", lines[i]).strip())
                i += 1
            html_lines.append("<ul>" + "".join(f"<li>{itm}</li>" for itm in items) + "</ul>")
            continue

        if re.match(r"^\s*\d+[\.\)]\s+", line):
            items = []
            while i < len(lines) and re.match(r"^\s*\d+[\.\)]\s+", lines[i]):
                items.append(re.sub(r"^\s*\d+[\.\)]\s+", "", lines[i]).strip())
                i += 1
            html_lines.append("<ol>" + "".join(f"<li>{itm}</li>" for itm in items) + "</ol>")
            continue

        html_lines.append(line)
        i += 1

    html = "<br/>".join(html_lines)
    html = re.sub(r"(?:<br/>\s*){3,}", "<br/><br/>", html)
    return html

def clean_float(x, default=None):
    if x is None or (isinstance(x, float) and pd.isna(x)) or (isinstance(x, str) and x.strip() == ""):
        return default
    try:
        return float(x)
    except Exception:
        return default

def clean_bool(x, default=True):
    if isinstance(x, bool):
        return x
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return default
    s = str(x).strip().lower()
    if s in {"true", "t", "1", "yes", "y"}:
        return True
    if s in {"false", "f", "0", "no", "n"}:
        return False
    return default

# ---------- XML helpers ----------

def add_text_node(parent, tag, text, format="html"):
    """Create Moodle <tag><text>..</text></tag> with format attr where supported."""
    node = SubElement(parent, tag)
    if tag in {"questiontext", "generalfeedback", "feedback"}:
        node.set("format", format)
    text_el = SubElement(node, "text")
    text_el.text = "" if text is None else str(text)
    return node

def question_multichoice_xml(qrow, index, category_name):
    """
    Build a <question type="multichoice"> node from a DataFrame row.
    Required cols: Question, Choice1..5, Fraction1..5
    Optional: Name, GeneralFeedback, ChoiceFeedback1..5, DefaultGrade, Shuffle
    """
    q = Element("question", {"type": "multichoice"})

    # Name
    name_val = qrow.get("Name")
    if not name_val or (isinstance(name_val, float) and pd.isna(name_val)):
        name_val = f"Q{index+1}"
    name = SubElement(q, "name")
    SubElement(name, "text").text = str(strip_control(name_val))

    # Question text (HTML)
    qtext_html = convert_markup_to_html(str(qrow.get("Question", "")).strip())
    qt = SubElement(q, "questiontext", {"format": "html"})
    SubElement(qt, "text").text = str(qtext_html)

    # Default grade
    defaultgrade = clean_float(qrow.get("DefaultGrade"), default=5.0)
    SubElement(q, "defaultgrade").text = f"{defaultgrade:.2f}"

    # Penalty (fraction for adaptive; keep default 0)
    SubElement(q, "penalty").text = "0.0"

    # Single: if exactly one answer has the maximum fraction (>= 100 - tiny epsilon)
    fractions = [clean_float(qrow.get(f"Fraction{i}"), default=0.0) for i in range(1, 6)]
    max_frac = max(fractions) if fractions else 0.0
    single = fractions.count(max_frac) == 1 and max_frac >= 100.0 - 1e-6
    SubElement(q, "single").text = "true" if single else "false"

    # Shuffle answers
    shuffle = clean_bool(qrow.get("Shuffle"), default=True)
    SubElement(q, "shuffleanswers").text = "true" if shuffle else "false"

    # Answer numbering (a,b,c,…)
    SubElement(q, "answernumbering").text = "ABCD"

    # General Feedback
    gf = qrow.get("GeneralFeedback")
    gf_html = convert_markup_to_html(strip_control(gf))
    gfnode = SubElement(q, "generalfeedback", {"format": "html"})
    SubElement(gfnode, "text").text = str(gf_html)

    # 5 Choices
    for i in range(1, 6):
        choice_text = qrow.get(f"Choice{i}")
        choice_html = convert_markup_to_html(strip_control(choice_text))

        frac = clean_float(qrow.get(f"Fraction{i}"), default=0.0)
        if frac is None:
            frac = 0.0
        frac = max(-100.0, min(100.0, float(frac)))

        ans = SubElement(q, "answer", {"fraction": str(frac), "format": "html"})
        text_el = SubElement(ans, "text")
        text_el.text = str(choice_html)

        # per-choice feedback (optional)
        cfb = qrow.get(f"ChoiceFeedback{i}")
        cfb_html = convert_markup_to_html(strip_control(cfb))

        fb = SubElement(ans, "feedback", {"format": "html"})
        fb_text = SubElement(fb, "text")
        fb_text.text = str(cfb_html)

    return q   # ✅ FIX: return the <question> node

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize column names (strip, lowercase, replace spaces/underscores/dashes)
    so users can write 'choice 1' or 'Choice_1' etc.
    """
    mapping = {}
    for c in df.columns:
        key = re.sub(r"[\s\-_]+", "", str(c).strip().lower())
        mapping[c] = key
    df2 = df.rename(columns=mapping)

    # Build reverse map to pretty keys we expect
    canonical = {
        "name": "Name",
        "question": "Question",
        "generalfeedback": "GeneralFeedback",
        "defaultgrade": "DefaultGrade",
        "shuffle": "Shuffle",
    }
    for i in range(1, 6):
        canonical[f"choice{i}"] = f"Choice{i}"
        canonical[f"fraction{i}"] = f"Fraction{i}"
        canonical[f"choicefeedback{i}"] = f"ChoiceFeedback{i}"

    out = {}
    for key, pretty in canonical.items():
        if key in df2.columns:
            out[pretty] = df2[key]
        else:
            out[pretty] = None

    # Re-assemble into a new uniform frame
    new_df = pd.DataFrame({k: out[k] for k in canonical.values() if out[k] is not None})

    # Ensure required columns exist; if not, raise a clear error
    required = ["Question"] + [f"Choice{i}" for i in range(1, 6)] + [f"Fraction{i}" for i in range(1, 6)]
    missing = [c for c in required if c not in new_df.columns]
    if missing:
        raise ValueError(f"Your Excel is missing required columns: {', '.join(missing)}")

    return new_df

def build_quiz_root():
    return Element("quiz")

def add_category(root, category_name: str):
    q = SubElement(root, "question", {"type": "category"})
    cattext = f"$course$/{category_name}" if category_name else "$course$/Default"
    category = SubElement(q, "category")
    SubElement(category, "text").text = cattext

def excel_to_moodle_xml(in_path: Path, out_path: Path, category: str):
    df = pd.read_excel(in_path)
    df = normalize_columns(df)
    root = build_quiz_root()
    add_category(root, category)

    for idx, row in df.iterrows():
        qnode = question_multichoice_xml(row, idx, category)
        root.append(qnode)

    tree = ElementTree(root)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    tree.write(out_path, encoding="utf-8", xml_declaration=True)
    print(f"✅ Wrote Moodle XML: {out_path}")

def main():
    ap = argparse.ArgumentParser(description="Convert Excel to Moodle XML (MCQ, 5 choices).")
    ap.add_argument("--in", dest="infile", required=True, help="Input Excel file (.xlsx)")
    ap.add_argument("--out", dest="outfile", required=True, help="Output Moodle XML file (.xml)")
    ap.add_argument("--category", dest="category", default="Imported from Excel",
                    help="Moodle category path under $course$ (default: 'Imported from Excel')")
    args = ap.parse_args()

    in_path = Path(args.infile)
    out_path = Path(args.outfile)

    if not in_path.exists():
        print(f"ERROR: Input file not found: {in_path}", file=sys.stderr)
        sys.exit(1)

    try:
        excel_to_moodle_xml(in_path, out_path, args.category)
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(2)

if __name__ == "__main__":
    main()

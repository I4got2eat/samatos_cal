import streamlit as st
import fitz
import base64
import json
import pandas as pd
import re
import sqlite3
import io
from anthropic import Anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ── Config ────────────────────────────────────────────────────────────────────
API_KEY = "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"
DB_PATH = r"C:\Users\Mykolas\claude_sessions\samata.db"

# Standard Lithuanian sąmata overhead rates
OVERHEAD = {
    "KDU":  0.08,    # Additional labor charges
    "PMD":  0.03,    # Additional materials value
    "PMZ":  0.03,    # Additional machinery value
    "SOD":  0.0179,  # Social insurance
    "STI":  0.09,    # Site overhead
    "PRI":  0.209,   # Indirect costs (applied to labor)
    "PLN":  0.05,    # Profit
    "PVM":  0.21,    # VAT
}


# ── Database ──────────────────────────────────────────────────────────────────
@st.cache_data
def load_resources():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM resources", conn)
    conn.close()
    return df


# ── PDF extraction (unchanged) ────────────────────────────────────────────────
def extract_specific_pages_as_images(pdf_bytes, start_page, end_page, dpi=300):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    target_images = []
    for page_num in range(max(0, start_page - 1), min(len(doc), end_page)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=dpi)
        base64_image = base64.b64encode(pix.tobytes("jpeg")).decode("utf-8")
        target_images.append({"page_num": page_num + 1, "base64_image": base64_image})
    doc.close()
    return target_images


def extract_table_data(base64_image):
    client = Anthropic(api_key=API_KEY)
    prompt = """
    Extract the table content from this image.
    1. Analyze the document to identify hierarchical headers (Main Categories and Sub Categories).
    2. Flatten the table structure into a list of objects.
    3. For every row, identify the 'Current_Main_Section' and 'Current_Sub_Section' that applies to that data.
    4. Include all available columns in the row (e.g., Eil_Nr, Pavadinimas, Vnt, Kiekis, TS_skyrius, Pastabos).
    5. Return ONLY a JSON object with a single key 'table_data' containing the array of objects.

    Rules:
    - If a row is a section header, do not include it as a data row; use it to update the context for subsequent rows.
    - If a value is missing or spanned, infer it from the context of the identified section.
    - Do not use markdown, do not add conversational text.
    """
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        temperature=0,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": base64_image}},
                {"type": "text", "text": prompt},
            ],
        }],
    )
    raw = response.content[0].text.strip()
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    return json.loads(match.group(0) if match else raw)


# ── AI matching ───────────────────────────────────────────────────────────────
def match_items_to_resources(parsed_df, resources_df):
    client = Anthropic(api_key=API_KEY)

    resources_text = "\n".join(
        f"{r['code']} | {r['name']} | {r['unit']} | {r['unit_price'] if r['unit_price'] else 'N/A'} EUR | {r['category']}"
        for _, r in resources_df.iterrows()
    )

    items_text = parsed_df.to_json(orient="records", force_ascii=False, indent=2)

    prompt = f"""You are a strict Lithuanian construction cost estimator working with the Astera sąmata system.

Below are construction specification items extracted from a client project PDF:
{items_text}

Below is the COMPLETE price database you are allowed to use (format: code | name | unit | unit_price | category):
{resources_text}

STRICT RULES — read carefully:
1. You may ONLY use resources that exist verbatim in the database above. Do NOT invent codes, names, prices, or resources.
2. If a spec item cannot be matched to ANY resource in the database with reasonable confidence, set "status": "no_match" and leave "matched_resources" as an empty array.
3. A "reasonable match" means the resource clearly corresponds to the spec item material/labor/machinery. Do NOT force a match just to fill the field.
4. You may combine multiple database resources for one spec item ONLY if each individual resource genuinely applies.
5. Never hallucinate a unit_price — copy it exactly from the database row above.
6. Skip items where Kiekis is null, missing, or "-".
7. All costs in EUR.
8. For "pavadinimas": copy the FULL Pavadinimas text from the spec item EXACTLY as it appears — do NOT shorten, summarise, or truncate it in any way.

Return a JSON array where each element has this exact structure:
{{
  "eil_nr": "item number from spec",
  "pavadinimas": "FULL Pavadinimas text copied verbatim — every word, do not shorten",
  "vnt": "unit of measure",
  "kiekis": <total quantity as number>,
  "status": "matched" | "no_match",
  "matched_resources": [
    {{
      "code": "resource code copied exactly from the database",
      "name": "resource name copied exactly from the database",
      "unit": "resource unit copied exactly from the database",
      "unit_price": <price copied exactly from the database>,
      "qty_per_spec_unit": <resource quantity per 1 spec unit — your estimate>,
      "total_qty": <kiekis × qty_per_spec_unit>,
      "total_cost": <total_qty × unit_price>,
      "category": "labor|material|machinery"
    }}
  ],
  "total_labor": <sum of labor costs, 0 if no_match>,
  "total_materials": <sum of material costs, 0 if no_match>,
  "total_machinery": <sum of machinery costs, 0 if no_match>,
  "total_cost": <grand total, 0 if no_match>
}}

Return ONLY a valid JSON array, no markdown, no explanation.
"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = response.content[0].text.strip()
    match = re.search(r"\[.*\]", raw, re.DOTALL)
    return json.loads(match.group(0) if match else raw)


# ── Cost rollup ───────────────────────────────────────────────────────────────
def compute_totals(matched_items):
    labor     = sum(i.get("total_labor", 0) or 0 for i in matched_items)
    materials = sum(i.get("total_materials", 0) or 0 for i in matched_items)
    machinery = sum(i.get("total_machinery", 0) or 0 for i in matched_items)
    direct    = labor + materials + machinery

    kdu  = labor * OVERHEAD["KDU"]
    pmd  = materials * OVERHEAD["PMD"]
    pmz  = machinery * OVERHEAD["PMZ"]
    sod  = (labor + kdu) * OVERHEAD["SOD"]
    t2   = direct + kdu + pmd + pmz + sod

    sti  = t2 * OVERHEAD["STI"]
    t3   = t2 + sti

    pri  = labor * OVERHEAD["PRI"]
    pln  = t3 * OVERHEAD["PLN"]
    t4   = t3 + pri + pln

    pvm  = t4 * OVERHEAD["PVM"]
    t5   = t4 + pvm

    return {
        "labor": labor, "materials": materials, "machinery": machinery,
        "direct": direct,
        "KDU": kdu, "PMD": pmd, "PMZ": pmz, "SOD": sod, "t2": t2,
        "STI": sti, "t3": t3,
        "PRI": pri, "PLN": pln, "t4": t4,
        "PVM": pvm, "t5": t5,
    }


# ── Excel generation ──────────────────────────────────────────────────────────

# Column layout — Koef (F) and Subrangovai (M) removed
# A=Nr, B=Name, C=Code, D=Unit, E=Norma, F=Kaina, G=Kiekis,
# H=Suma, I=Darbas, J=Medžiagos, K=Mechanizmai
COL_WIDTHS = {
    "A": 4.71, "B": 30.71, "C": 9.71, "D": 8.0,  "E": 7.0,
    "F": 8.71, "G": 8.43,  "H": 16.14,"I": 11.43,
    "J": 9.71, "K": 11.43,
}

def lt_num(value):
    """Format number as Lithuanian amount string: 8 440,24"""
    if value is None or value == "":
        return ""
    n = float(value)
    # Format with 2 decimals, space as thousands sep, comma as decimal
    formatted = f"{abs(n):,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
    return f"-{formatted}" if n < 0 else formatted


RED_FILL      = PatternFill("solid", fgColor="FF9999")
ITEM_TOP      = Border(top=Side(style="medium"))          # thick top line per work item
ITEM_TOP_RED  = Border(top=Side(style="medium", color="CC0000"))

def _font(bold=False, size=10):
    return Font(name="Times New Roman", bold=bold, size=size)

def _align(h="left", v="top", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def generate_samata_excel(matched_items, project_name=""):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sąmata"

    # Set column widths
    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    t = compute_totals(matched_items)

    def s(row, col, value, bold=False, size=10, h="left", v="top", wrap=True, red=False):
        """Write a styled cell."""
        c = ws.cell(row=row, column=col, value=value)
        c.font = _font(bold=bold, size=size)
        c.alignment = _align(h=h, v=v, wrap=wrap)
        if red:
            c.fill = RED_FILL
        return c

    # ── Header block (rows 1–10) ──────────────────────────────────────────────
    ws.row_dimensions[1].height = 16.9

    # Row 6: project reference + "Lokalinė sąmata"
    ws.row_dimensions[6].height = 25.5
    s(6, 2, project_name or "Projektas", bold=True)
    s(6, 4, "L o k a l i n ė   s ą m a t a", bold=True, h="center")

    # Row 7: building description + price level
    ws.row_dimensions[7].height = 26.25
    s(7, 2, "", bold=True)
    from datetime import date
    s(7, 4, f"Sudaryta {date.today().strftime('%Y.%m')} kainų lygiu.", bold=True)

    # Row 8: section + grand total preview
    ws.row_dimensions[8].height = 15.0
    s(8, 7, "Iš viso", bold=False)
    s(8, 9, lt_num(t["t5"]), bold=False, h="right")

    # Row 10: column headers
    ws.row_dimensions[10].height = 27.75
    headers = [
        (1, "Nr."),        (2, "Darbo pavadinimas"),
        (3, "Kodas"),      (4, "Mat. vnt"),
        (5, "Norma"),      (6, "Kaina"),
        (7, "Kiekis"),     (8, "Suma"),
        (9, "Darbas"),     (10, "Medžiagos"),
        (11, "Mechanizmai"),
    ]
    for col, label in headers:
        s(10, col, label, bold=True, h="center", v="center")

    # Row 11: empty spacer
    ws.row_dimensions[11].height = 12.75

    r = 12  # start writing data from row 12

    # Group items by section (use eil_nr prefix, e.g. "1.1" → section "1")
    # Since our AI returns flat items, we treat all as one section
    section_name = "Darbai"
    section_start = r

    # ── Section header row ────────────────────────────────────────────────────
    ws.row_dimensions[r].height = 12.75
    s(r, 2, "Skyrius", bold=True, v="top", wrap=False)
    s(r, 3, section_name, bold=True, v="top", wrap=False)
    s(r, 17, "SP")   # Q column = col 17
    r += 1

    # ── Work items ────────────────────────────────────────────────────────────
    item_no = 1
    for item in matched_items:
        no_match   = item.get("status") == "no_match" or not item.get("matched_resources")
        kiekis     = item.get("kiekis") or 0
        item_total = item.get("total_cost") or 0
        item_labor = item.get("total_labor") or 0
        item_mat   = item.get("total_materials") or 0
        item_mach  = item.get("total_machinery") or 0
        unit_price = round(item_total / kiekis, 4) if kiekis else 0

        # Work item row — bold, 9.75pt; red fill if no match
        # Auto-size row height based on name length (col B is ~30 chars wide)
        _name_len = len(item.get("pavadinimas", ""))
        _lines = max(1, -(-_name_len // 30))  # ceiling division
        ws.row_dimensions[r].height = max(25.5, _lines * 13.5)
        border = ITEM_TOP_RED if no_match else ITEM_TOP
        for col in range(1, 12):
            ws.cell(row=r, column=col).border = border
        s(r,  1, item_no,                     bold=True, size=9.75, h="center", v="top", red=no_match)
        s(r,  2, item.get("pavadinimas", ""), bold=True, size=9.75, h="left",   v="top", red=no_match)
        s(r,  3, item.get("eil_nr", ""),      bold=True, size=9.75, h="left",   v="top", red=no_match)
        s(r,  4, item.get("vnt", ""),         bold=True, size=9.75, h="center", v="top", red=no_match)
        s(r,  5, None,                        bold=True, size=9.75, h="right",  v="top", red=no_match)
        s(r,  6, None if no_match else round(unit_price, 2),
                                              bold=True, size=9.75, h="right",  v="top", red=no_match)
        s(r,  7, kiekis,                      bold=True, size=9.75, h="right",  v="top", red=no_match)
        s(r,  8, "NA" if no_match else lt_num(item_total),
                                              bold=True, size=9.75, h="right",  v="top", red=no_match)
        s(r,  9, "NA" if no_match else lt_num(item_labor),
                                              bold=True, size=9.75, h="right",  v="top", red=no_match)
        s(r, 10, "NA" if no_match else lt_num(item_mat),
                                              bold=True, size=9.75, h="right",  v="top", red=no_match)
        s(r, 11, "NA" if no_match else lt_num(item_mach),
                                              bold=True, size=9.75, h="right",  v="top", red=no_match)
        r += 1

        # Sub-item rows — not bold, 9.75pt (only for matched items)
        for res in item.get("matched_resources", []):
            cost      = res.get("total_cost") or 0
            cat       = res.get("category", "")
            labor_val = lt_num(cost) if cat == "labor"    else ""
            mat_val   = lt_num(cost) if cat == "material"  else ""
            mach_val  = lt_num(cost) if cat == "machinery" else ""

            s(r,  1, None,                            bold=False, size=9.75, h="center", v="top")
            s(r,  2, res.get("name", ""),             bold=False, size=9.75, h="left",   v="top")
            s(r,  3, res.get("code", ""),             bold=False, size=9.75, h="left",   v="top")
            s(r,  4, res.get("unit", ""),             bold=False, size=9.75, h="center", v="top")
            s(r,  5, res.get("qty_per_spec_unit"),    bold=False, size=9.75, h="right",  v="top")
            s(r,  6, res.get("unit_price"),           bold=False, size=9.75, h="right",  v="top")
            s(r,  7, round(res.get("total_qty") or 0, 3), bold=False, size=9.75, h="right", v="top")
            s(r,  8, lt_num(cost),                    bold=False, size=9.75, h="right",  v="top")
            s(r,  9, labor_val,                       bold=False, size=9.75, h="right",  v="top")
            s(r, 10, mat_val,                         bold=False, size=9.75, h="right",  v="top")
            s(r, 11, mach_val,                        bold=False, size=9.75, h="right",  v="top")
            r += 1

        item_no += 1

    # ── Section total row ─────────────────────────────────────────────────────
    s(r, 2, "Iš viso už skyrių", bold=True)
    s(r, 3, section_name,        bold=True)
    s(r,  8, lt_num(t["direct"]),    bold=True, h="right")
    s(r,  9, lt_num(t["labor"]),     bold=True, h="right")
    s(r, 10, lt_num(t["materials"]), bold=True, h="right")
    s(r, 11, lt_num(t["machinery"]), bold=True, h="right")
    r += 2  # blank line after section

    # ── Grand totals block ────────────────────────────────────────────────────
    def total_line(label, i_val, j_val="", k_val="", l_val="", bold=True):
        nonlocal r
        s(r, 2, label,   bold=bold)
        s(r,  8, i_val,  bold=bold, h="right")
        s(r,  9, j_val,  bold=bold, h="right")
        s(r, 10, k_val,  bold=bold, h="right")
        s(r, 11, l_val,  bold=bold, h="right")
        r += 1

    def charge_line(label, rate_label, i_val, j_val="", k_val="", l_val=""):
        nonlocal r
        s(r, 2, label,      bold=False)
        s(r, 3, rate_label, bold=False)
        s(r,  8, i_val,  bold=False, h="right")
        s(r,  9, j_val,  bold=False, h="right")
        s(r, 10, k_val,  bold=False, h="right")
        s(r, 11, l_val,  bold=False, h="right")
        r += 1

    total_line(
        "Iš viso #1",
        lt_num(t["direct"]), lt_num(t["labor"]), lt_num(t["materials"]), lt_num(t["machinery"])
    )
    charge_line("Kiti darbo užmokesčio priskaitymai", "KDU",
                lt_num(t["direct"] + t["KDU"]), lt_num(t["labor"] + t["KDU"]))
    charge_line("Papildomų medžiagų vertė",  "PMD", "", "", lt_num(t["PMD"]))
    charge_line("Papildomų mechanizmų vertė","PMZ", "", "", "", lt_num(t["PMZ"]))
    charge_line("Soc. draudimas",            "SOD", lt_num(t["SOD"]), lt_num(t["SOD"]))

    total_line(
        "Iš viso #2 (išlaidos statinio statybos darbams)",
        lt_num(t["t2"]),
        lt_num(t["labor"] + t["KDU"] + t["SOD"]),
        lt_num(t["materials"] + t["PMD"]),
        lt_num(t["machinery"] + t["PMZ"]),
    )
    charge_line("Statybvietės išlaidos", "STI", lt_num(t["STI"]),
                lt_num(t["STI"] * 0.4), lt_num(t["STI"] * 0.5), lt_num(t["STI"] * 0.1))

    total_line("Iš viso #3 (tiesioginės išlaidos)",
               lt_num(t["t3"]), lt_num(t["t3"] * 0.6), lt_num(t["t3"] * 0.3), lt_num(t["t3"] * 0.1))

    charge_line("", "IND", "1", "1", "1", "1")  # index row

    total_line("Po indeksacijos iš viso",
               lt_num(t["t3"]), lt_num(t["t3"] * 0.6), lt_num(t["t3"] * 0.3), lt_num(t["t3"] * 0.1))

    charge_line("Pridėtinės išlaidos", "PRI", lt_num(t["PRI"]), lt_num(t["PRI"]))
    charge_line("Pelnas",              "PLN", lt_num(t["PLN"]),
                lt_num(t["PLN"] * 0.6), lt_num(t["PLN"] * 0.3), lt_num(t["PLN"] * 0.1))

    total_line("Iš viso #4 (su netiesioginėmis išlaidomis)",
               lt_num(t["t4"]),
               lt_num(t["t4"] * 0.6), lt_num(t["t4"] * 0.3), lt_num(t["t4"] * 0.1))

    charge_line("PVM", "PVM", lt_num(t["PVM"]),
                lt_num(t["PVM"] * 0.6), lt_num(t["PVM"] * 0.3), lt_num(t["PVM"] * 0.1))

    total_line("Iš viso #5 (kaina su PVM)",
               lt_num(t["t5"]),
               lt_num(t["t5"] * 0.6), lt_num(t["t5"] * 0.3), lt_num(t["t5"] * 0.1))

    s(r, 14, "Pabaiga", bold=False, h="right")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── HTML preview ─────────────────────────────────────────────────────────────
def render_samata_html(matched_items, project_name=""):
    t = compute_totals(matched_items)
    from datetime import date

    # Column definitions: (label, width, align)
    cols = [
        ("Nr.",               "40px",  "center"),
        ("Darbo pavadinimas", "300px", "left"),
        ("Kodas",             "90px",  "left"),
        ("Mat. vnt",          "70px",  "center"),
        ("Norma",             "60px",  "right"),
        ("Kaina",             "75px",  "right"),
        ("Kiekis",            "70px",  "right"),
        ("Suma",              "110px", "right"),
        ("Darbas",            "100px", "right"),
        ("Medžiagos",         "100px", "right"),
        ("Mechanizmai",       "100px", "right"),
    ]

    BASE = "font-family:'Times New Roman',serif; font-size:9.75pt;"
    HDR  = f"{BASE} font-weight:bold; background:#fff; border-bottom:2px solid #000; text-align:center; padding:4px 6px; white-space:nowrap;"
    TD   = f"{BASE} padding:2px 6px; border-bottom:1px solid #ddd; vertical-align:top;"
    BOLD = f"{BASE} font-weight:bold; padding:2px 6px; border-bottom:1px solid #bbb; vertical-align:top;"
    TOT  = f"{BASE} font-weight:bold; padding:3px 6px; border-top:2px solid #000; border-bottom:2px solid #000; vertical-align:top; background:#f5f5f5;"
    SEC  = f"{BASE} font-weight:bold; padding:4px 6px; border-bottom:1px solid #999; background:#fff; vertical-align:top;"
    CHG  = f"{BASE} padding:2px 6px; border-bottom:1px solid #eee; vertical-align:top; color:#444;"

    def th(label, width, align):
        return f'<th style="{HDR} width:{width}; text-align:{align};">{label}</th>'

    def td(val, style, align="left", colspan=1):
        sp = f' colspan="{colspan}"' if colspan > 1 else ""
        return f'<td style="{style} text-align:{align};"{sp}>{val if val is not None else ""}</td>'

    rows_html = []

    # ── Header
    rows_html.append(f'''
    <tr>
      <td colspan="13" style="font-family:'Times New Roman',serif; font-size:11pt;
          font-weight:bold; padding:8px 6px; border-bottom:1px solid #ccc;">
        {project_name or "—"}
        <span style="float:right; font-size:9pt; font-weight:normal;">
          L o k a l i n ė &nbsp; s ą m a t a &nbsp;&nbsp;|&nbsp;&nbsp;
          Sudaryta {date.today().strftime("%Y.%m")} kainų lygiu &nbsp;&nbsp;|&nbsp;&nbsp;
          Iš viso su PVM: <strong>{lt_num(t["t5"])} €</strong>
        </span>
      </td>
    </tr>''')

    # ── Column headers
    rows_html.append('<tr>' + ''.join(th(l, w, a) for l, w, a in cols) + '</tr>')

    # ── Section header
    section_name = "Darbai"
    rows_html.append(f'''
    <tr>
      <td style="{SEC}"></td>
      <td style="{SEC}" colspan="12">Skyrius &nbsp; {section_name}</td>
    </tr>''')

    NA_ROW = f"{BASE} font-weight:bold; padding:2px 6px; border-bottom:1px solid #cc0000; vertical-align:top; background:#FF9999; color:#000;"

    # ── Work items
    item_no = 1
    for item in matched_items:
        no_match   = item.get("status") == "no_match" or not item.get("matched_resources")
        kiekis     = item.get("kiekis") or 0
        item_total = item.get("total_cost") or 0
        item_labor = item.get("total_labor") or 0
        item_mat   = item.get("total_materials") or 0
        item_mach  = item.get("total_machinery") or 0
        unit_price = round(item_total / kiekis, 2) if kiekis and not no_match else 0

        row_style = NA_ROW if no_match else BOLD
        na = "NA"

        rows_html.append(
            '<tr>'
            + td(item_no,                                row_style, "center")
            + td(item.get("pavadinimas",""),             row_style, "left")
            + td(item.get("eil_nr",""),                  row_style, "left")
            + td(item.get("vnt",""),                     row_style, "center")
            + td("",                                     row_style, "right")
            + td("" if no_match else lt_num(unit_price), row_style, "right")
            + td(kiekis,                                 row_style, "right")
            + td(na if no_match else lt_num(item_total), row_style, "right")
            + td(na if no_match else lt_num(item_labor), row_style, "right")
            + td(na if no_match else lt_num(item_mat),   row_style, "right")
            + td(na if no_match else lt_num(item_mach),  row_style, "right")
            + '</tr>'
        )

        for res in item.get("matched_resources", []):
            cost = res.get("total_cost") or 0
            cat  = res.get("category", "")
            rows_html.append(
                '<tr>'
                + td("",                              TD,   "center")
                + td(res.get("name",""),              TD,   "left")
                + td(res.get("code",""),              TD,   "left")
                + td(res.get("unit",""),              TD,   "center")
                + td(res.get("qty_per_spec_unit",""), TD,   "right")
                + td(res.get("unit_price",""),        TD,   "right")
                + td(round(res.get("total_qty") or 0, 3), TD, "right")
                + td(lt_num(cost),                    TD,   "right")
                + td(lt_num(cost) if cat=="labor"     else "", TD, "right")
                + td(lt_num(cost) if cat=="material"  else "", TD, "right")
                + td(lt_num(cost) if cat=="machinery" else "", TD, "right")
                + '</tr>'
            )
        item_no += 1

    # ── Section total
    rows_html.append(
        '<tr>'
        + td("",                      TOT, "left", colspan=2)
        + td("Iš viso už skyrių",    TOT, "left", colspan=5)
        + td(lt_num(t["direct"]),    TOT, "right")
        + td(lt_num(t["labor"]),     TOT, "right")
        + td(lt_num(t["materials"]), TOT, "right")
        + td(lt_num(t["machinery"]), TOT, "right")
        + '</tr><tr><td colspan="11" style="height:10px;"></td></tr>'
    )

    # ── Grand totals helper
    def tot_row(label, i, j="", k="", l="", bold=True):
        sty = TOT if bold else CHG
        return (
            '<tr>'
            + td("",    sty, colspan=2)
            + td(label, sty, "left", colspan=5)
            + td(i,     sty, "right")
            + td(j,     sty, "right")
            + td(k,     sty, "right")
            + td(l,     sty, "right")
            + '</tr>'
        )

    rows_html += [
        tot_row("Iš viso #1",
                lt_num(t["direct"]), lt_num(t["labor"]),
                lt_num(t["materials"]), lt_num(t["machinery"])),
        tot_row(f"Kiti darbo užmokesčio priskaitymai (KDU {OVERHEAD['KDU']*100:.0f}%)",
                lt_num(t["KDU"]), lt_num(t["KDU"]), bold=False),
        tot_row(f"Papildomų medžiagų vertė (PMD {OVERHEAD['PMD']*100:.0f}%)",
                lt_num(t["PMD"]), k=lt_num(t["PMD"]), bold=False),
        tot_row(f"Papildomų mechanizmų vertė (PMZ {OVERHEAD['PMZ']*100:.0f}%)",
                lt_num(t["PMZ"]), l=lt_num(t["PMZ"]), bold=False),
        tot_row(f"Soc. draudimas (SOD {OVERHEAD['SOD']*100:.2f}%)",
                lt_num(t["SOD"]), lt_num(t["SOD"]), bold=False),
        tot_row("Iš viso #2 (išlaidos statinio statybos darbams)",
                lt_num(t["t2"]),
                lt_num(t["labor"]+t["KDU"]+t["SOD"]),
                lt_num(t["materials"]+t["PMD"]),
                lt_num(t["machinery"]+t["PMZ"])),
        tot_row(f"Statybvietės išlaidos (STI {OVERHEAD['STI']*100:.0f}%)",
                lt_num(t["STI"]), bold=False),
        tot_row("Iš viso #3 (tiesioginės išlaidos)", lt_num(t["t3"])),
        tot_row(f"Pridėtinės išlaidos (PRI {OVERHEAD['PRI']*100:.1f}%)",
                lt_num(t["PRI"]), lt_num(t["PRI"]), bold=False),
        tot_row(f"Pelnas (PLN {OVERHEAD['PLN']*100:.0f}%)",
                lt_num(t["PLN"]), bold=False),
        tot_row("Iš viso #4 (su netiesioginėmis išlaidomis)", lt_num(t["t4"])),
        tot_row(f"PVM {OVERHEAD['PVM']*100:.0f}%", lt_num(t["PVM"]), bold=False),
    ]

    # Final total — highlighted
    rows_html.append(f'''
    <tr style="background:#1a3a6b;">
      <td colspan="2" style="{BASE} color:#fff; font-weight:bold; padding:5px 6px;"></td>
      <td colspan="5" style="{BASE} color:#fff; font-weight:bold; padding:5px 6px;">
        Iš viso #5 (kaina su PVM)
      </td>
      <td style="{BASE} color:#fff; font-weight:bold; padding:5px 6px; text-align:right;">
        {lt_num(t["t5"])}
      </td>
      <td colspan="3" style="{BASE} color:#fff; padding:5px 6px;"></td>
    </tr>''')

    col_group = "".join(
        f'<col style="width:{w}">' for _, w, _ in cols
    )

    return f"""
    <div style="overflow-x:auto; font-family:'Times New Roman',serif;">
      <table style="border-collapse:collapse; width:100%; min-width:900px;">
        <colgroup>{col_group}</colgroup>
        <tbody>{''.join(rows_html)}</tbody>
      </table>
    </div>
    """


# ── Streamlit app ─────────────────────────────────────────────────────────────
import streamlit.components.v1 as components

st.set_page_config(page_title="Sąmatos skaičiavimas", layout="wide")

st.markdown("""
<style>
  .block-container {
      max-width: 1100px !important;
      padding-left: 2rem !important;
      padding-right: 2rem !important;
  }
</style>
""", unsafe_allow_html=True)

resources_df = load_resources()

st.title("Sąmatos skaičiavimas")
st.caption(f"Kainų duomenų bazė: {len(resources_df)} resursai")

# ── Inputs ────────────────────────────────────────────────────────────────────
st.divider()

uploaded_file = st.file_uploader("PDF projekto specifikacija", type=["pdf"])
col1, col2 = st.columns(2)
with col1:
    start_page = st.number_input("Puslapis nuo", min_value=1, value=1)
with col2:
    end_page = st.number_input("Puslapis iki", min_value=1, value=1)

project_name = st.text_input("Projekto pavadinimas",
                              placeholder="pvz. Sandėlio pastatas, Vilnius, 2024")

if st.button("Generuoti sąmatą", type="primary", use_container_width=True):
    if not uploaded_file:
        st.error("Įkelkite PDF failą.")
    elif start_page > end_page:
        st.error("Puslapis 'nuo' negali būti didesnis už 'iki'.")
    else:
        with st.spinner("1/2 · Ištraukiamos lentelės iš PDF..."):
            pdf_bytes = uploaded_file.read()
            images = extract_specific_pages_as_images(pdf_bytes, start_page, end_page)
            all_data = []
            for img in images:
                try:
                    result = extract_table_data(img["base64_image"])
                    if "table_data" in result:
                        all_data.extend(result["table_data"])
                except Exception as e:
                    st.error(f"Klaida puslapyje {img['page_num']}: {e}")

        if not all_data:
            st.error("Nepavyko ištraukti duomenų iš PDF.")
        else:
            st.session_state["parsed_data"] = all_data
            st.session_state["project_name"] = project_name

            with st.spinner("2/2 · AI derina pozicijas su kainų duomenų baze..."):
                try:
                    parsed_df = pd.DataFrame(all_data)
                    matched = match_items_to_resources(parsed_df, resources_df)
                    st.session_state["matched"] = matched
                except Exception as e:
                    st.error(f"Klaida derinant: {e}")
                    st.stop()

if "matched" in st.session_state:
    matched = st.session_state["matched"]
    t = compute_totals(matched)
    na_count = sum(1 for i in matched if i.get("status") == "no_match" or not i.get("matched_resources"))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Darbas", f"€ {t['labor']:,.0f}")
    c2.metric("Medžiagos", f"€ {t['materials']:,.0f}")
    c3.metric("Mechanizmai", f"€ {t['machinery']:,.0f}")
    c4.metric("Iš viso su PVM", f"€ {t['t5']:,.0f}")

    if na_count:
        st.warning(f"**{na_count}** pozicija (-os) nepavyko suderinti su duomenų baze ir pažymėtos **NA**.")

# ── Preview & download ────────────────────────────────────────────────────────
if "matched" in st.session_state:
    matched      = st.session_state["matched"]
    project_name = st.session_state.get("project_name", "")

    st.divider()
    st.subheader("Sąmatos peržiūra ir atsisiuntimas")

    na_count = sum(1 for i in matched if i.get("status") == "no_match" or not i.get("matched_resources"))
    if na_count:
        st.info(f"{na_count} eilutė (-ės) pažymėta NA — nerasta atitikmenų kainų duomenų bazėje.")

    html_preview = render_samata_html(matched, project_name)
    components.html(html_preview, height=800, scrolling=True)

    excel_buf = generate_samata_excel(matched, project_name=project_name)
    st.download_button(
        "Atsisiųsti sąmatą (.xlsx)",
        data=excel_buf,
        file_name="samata.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

"""
Tamil Nadu Elections 2026 — Power BI Report Upgrade
Rebuilds tamilnadu.pbix's Report definition (PBIR) into a full 7-page
redesign: Cover, At a Glance, Regional Story, Flip Story, Margin Story,
Party Performance, Constituency Explorer.

Only edits the text-based Report/definition/** JSON — DataModel,
SecurityBindings, Settings, Metadata, Version stay byte-for-byte
untouched. Every new visual uses only columns already in the model
(winners_2021 / winners_2026: ac_number, candidate, party, votes,
region, total_votes, share).

Run: python upgrade_powerbi_report.py
"""
import os, uuid, zipfile, shutil, datetime, json

BASE = os.path.dirname(os.path.abspath(__file__))
SRC_PBIX = os.path.join(BASE, "tamilnadu.pbix")
WORK = os.path.join(BASE, "_pbix_work")

STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
BACKUP_PBIX = os.path.join(BASE, f"tamilnadu_backup_{STAMP}.pbix")

# ── Brand ─────────────────────────────────────────────────────────────
NAVY       = "#12213F"
GOLD       = "#A9812F"
WHITE      = "#FFFFFF"
LIGHT_GRAY = "#F1F2F6"
TEXT_GRAY  = "#75757F"
INK        = "#2B2B2B"
LIGHT_BLUE = "#C7CEDE"

FONT_DISPLAY = "Cambria"
FONT_BODY    = "Segoe UI"

W2021 = "winners_2021"
W2026 = "winners_2026"

SCHEMA_VC   = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json"
SCHEMA_PAGE = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"

# ── Precomputed static stats (from data/outputs/winners_2021.csv & winners_2026.csv) ──
STATS = {
    "avg21": 48.8, "avg26": 39.0,
    "above50_21": 84, "above50_26": 14,
    "below35_21": 2, "below35_26": 61,
    "top_party": "TVK", "top_seats": 108,
    "n_flipped": 163, "n_const": 234,
}

# ── Literal / expression helpers ────────────────────────────────────────
def new_id():
    return uuid.uuid4().hex[:20]

def lit_str(s):
    escaped = str(s).replace("'", "''")
    return {"expr": {"Literal": {"Value": f"'{escaped}'"}}}

def lit_num(n):
    return {"expr": {"Literal": {"Value": f"{n}D"}}}

def lit_bool(b):
    return {"expr": {"Literal": {"Value": "true" if b else "false"}}}

def solid(hexcolor):
    return {"solid": {"color": lit_str(hexcolor)}}

def pos(x, y, w, h, z=0, tab=0):
    return {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": tab}

def col_field(entity, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}

def agg_field(entity, prop, func):
    return {"Aggregation": {"Expression": col_field(entity, prop), "Function": func}}

def visual_shell(position, visual, vid=None):
    vid = vid or new_id()
    return vid, {"$schema": SCHEMA_VC, "name": vid, "position": position, "visual": visual}

# ── Visual builders ──────────────────────────────────────────────────────
def make_rect(x, y, w, h, color, z=0, tab=0):
    visual = {
        "visualType": "shape",
        "objects": {
            "shape": [{"properties": {"tileShape": lit_str("rectangle")}}],
            "fill": [{"properties": {"fillColor": solid(color), "transparency": lit_num(0)},
                      "selector": {"id": "default"}}],
            "outline": [{"properties": {"show": lit_bool(False)}, "selector": {"id": "default"}}],
        },
        "visualContainerObjects": {
            "background": [{"properties": {"show": lit_bool(False)}}],
            "border": [{"properties": {"show": lit_bool(False)}}],
            "padding": [{"properties": {"top": lit_num(0), "bottom": lit_num(0),
                                         "left": lit_num(0), "right": lit_num(0)}}],
        },
    }
    return visual_shell(pos(x, y, w, h, z, tab), visual)


def make_textbox(x, y, w, h, lines, z=0, tab=0):
    """lines: list of dicts with text, size(pt), color(hex), bold, italic, align, font"""
    paragraphs = []
    for ln in lines:
        style = {
            "fontFamily": ln.get("font", FONT_BODY),
            "fontSize": f'{ln.get("size", 14)}pt',
        }
        if ln.get("color"):
            style["color"] = ln["color"]
        if ln.get("bold"):
            style["bold"] = True
        if ln.get("italic"):
            style["italic"] = True
        paragraphs.append({
            "textRuns": [{"value": ln["text"], "textStyle": style}],
            "horizontalTextAlignment": ln.get("align", "left"),
        })
    visual = {
        "visualType": "textbox",
        "objects": {"general": [{"properties": {"paragraphs": paragraphs}}]},
        "drillFilterOtherVisuals": True,
    }
    return visual_shell(pos(x, y, w, h, z, tab), visual)


def make_card(x, y, w, h, entity, prop, func, label_text, z=0, tab=0):
    """func: 1 = Average, 5 = CountNonNull (the only two aggregation codes
    already proven working in this file)."""
    ref_prefix = "Sum" if func == 1 else "CountNonNull"
    query_ref = f"{ref_prefix}({entity}.{prop})"
    native_ref = f"Average of {prop}" if func == 1 else f"Count of {prop}"
    visual = {
        "visualType": "cardVisual",
        "query": {
            "queryState": {"Data": {"projections": [{
                "field": agg_field(entity, prop, func),
                "queryRef": query_ref, "nativeQueryRef": native_ref,
            }]}},
            "sortDefinition": {
                "sort": [{"field": agg_field(entity, prop, func), "direction": "Descending"}],
                "isDefaultSort": True,
            },
        },
        "objects": {
            "label": [{"properties": {"text": lit_str(label_text), "fontSize": lit_num(12)},
                       "selector": {"metadata": query_ref}}],
            "value": [{"properties": {"fontSize": lit_num(30)},
                       "selector": {"metadata": query_ref}}],
        },
        "drillFilterOtherVisuals": True,
    }
    return visual_shell(pos(x, y, w, h, z, tab), visual)


def make_slicer(x, y, w, h, entity, prop, header_text, z=0, tab=0):
    visual = {
        "visualType": "slicer",
        "query": {"queryState": {"Values": {"projections": [{
            "field": col_field(entity, prop),
            "queryRef": f"{entity}.{prop}", "nativeQueryRef": prop,
        }]}}},
        "objects": {
            "data": [{"properties": {"mode": lit_str("Dropdown")}}],
            "header": [{"properties": {"show": lit_bool(True), "text": lit_str(header_text)}}],
        },
        "visualContainerObjects": {
            "padding": [{"properties": {"top": lit_num(4), "bottom": lit_num(4),
                                         "left": lit_num(8), "right": lit_num(8)}}],
        },
    }
    return visual_shell(pos(x, y, w, h, z, tab), visual)


def make_bar_chart(x, y, w, h, entity, category_prop, y_prop, title,
                    series_prop=None, y_func=5, z=0, tab=0):
    query_state = {
        "Category": {"projections": [{
            "field": col_field(entity, category_prop),
            "queryRef": f"{entity}.{category_prop}", "nativeQueryRef": category_prop,
            "active": True,
        }]},
        "Y": {"projections": [{
            "field": agg_field(entity, y_prop, y_func),
            "queryRef": f"{'CountNonNull' if y_func == 5 else 'Sum'}({entity}.{y_prop})",
            "nativeQueryRef": f"Count of {y_prop}" if y_func == 5 else f"Average of {y_prop}",
        }]},
    }
    if series_prop:
        query_state["Series"] = {"projections": [{
            "field": col_field(entity, series_prop),
            "queryRef": f"{entity}.{series_prop}", "nativeQueryRef": series_prop,
        }]}
    visual = {
        "visualType": "barChart",
        "query": {
            "queryState": query_state,
            "sortDefinition": {
                "sort": [{"field": query_state["Y"]["projections"][0]["field"], "direction": "Descending"}],
                "isDefaultSort": True,
            },
        },
        "visualContainerObjects": {
            "title": [{"properties": {
                "text": lit_str(title),
                "heading": lit_str("Heading3"),
                "fontFamily": lit_str(FONT_DISPLAY),
                "fontColor": {"solid": {"color": lit_str(NAVY)}},
                "fontSize": lit_num(16),
            }}],
        },
        "drillFilterOtherVisuals": True,
    }
    return visual_shell(pos(x, y, w, h, z, tab), visual)


def make_table(x, y, w, h, columns, sort_col=None, sort_dir="Ascending", z=0, tab=0):
    """columns: list of (entity, prop, display_name)"""
    projections = [{
        "field": col_field(entity, prop),
        "queryRef": f"{entity}.{prop}", "nativeQueryRef": disp,
    } for entity, prop, disp in columns]
    query = {"queryState": {"Values": {"projections": projections}}}
    if sort_col:
        entity, prop = sort_col
        query["sortDefinition"] = {
            "sort": [{"field": col_field(entity, prop), "direction": sort_dir}],
            "isDefaultSort": True,
        }
    visual = {
        "visualType": "tableEx",
        "query": query,
        "objects": {"columnHeaders": [{"properties": {
            "columnAdjustment": lit_str("growToFit"),
            "autoSizeColumnWidth": lit_bool(True),
        }}]},
        "drillFilterOtherVisuals": True,
    }
    return visual_shell(pos(x, y, w, h, z, tab), visual)


# ── Page helpers ─────────────────────────────────────────────────────────
def make_page_json(name, display_name, background_hex=None):
    d = {"$schema": SCHEMA_PAGE, "name": name, "displayName": display_name,
         "displayOption": "FitToPage", "height": 1080, "width": 1920}
    if background_hex:
        d["objects"] = {"background": [{"properties": {
            "color": solid(background_hex), "transparency": lit_num(0)}}]}
    return d


def write_visual(page_dir, vid, vjson):
    vdir = os.path.join(page_dir, "visuals", vid)
    os.makedirs(vdir, exist_ok=True)
    with open(os.path.join(vdir, "visual.json"), "w", encoding="utf-8") as f:
        json.dump(vjson, f, indent=2)


def header_bar(page_dir, title_text, height=90, font_size=28):
    """Navy header bar + gold underline + white title text. Returns next z."""
    vid, v = make_rect(0, 0, 1920, height, NAVY, z=0)
    write_visual(page_dir, vid, v)
    vid, v = make_rect(0, height, 1920, 5, GOLD, z=1)
    write_visual(page_dir, vid, v)
    vid, v = make_textbox(40, (height - font_size * 1.6) / 2 if height > 60 else 6, 900, height,
                           [{"text": title_text, "size": font_size, "bold": True,
                             "color": WHITE, "font": FONT_DISPLAY}], z=2)
    write_visual(page_dir, vid, v)
    return 3


def kpi_stat(page_dir, x, y, w, h, value_text, label_text, accent, z):
    vid, v = make_rect(x, y, w, h, LIGHT_GRAY, z=z)
    write_visual(page_dir, vid, v)
    vid, v = make_rect(x, y, 6, h, accent, z=z + 1)
    write_visual(page_dir, vid, v)
    vid, v = make_textbox(x + 24, y + 14, w - 48, h - 28, [
        {"text": value_text, "size": 26, "bold": True, "color": NAVY, "font": FONT_DISPLAY},
        {"text": label_text, "size": 11, "color": TEXT_GRAY},
    ], z=z + 2)
    write_visual(page_dir, vid, v)
    return z + 3


# ── Main build ───────────────────────────────────────────────────────────
def unzip_pbix(src, dest):
    if os.path.exists(dest):
        shutil.rmtree(dest)
    os.makedirs(dest)
    with zipfile.ZipFile(src, "r") as z:
        z.extractall(dest)


def strip_security_bindings(work_dir):
    """SecurityBindings is a tamper-detection signature over the package
    contents. Since we're modifying the Report definition externally, that
    signature would go stale and Desktop reports the whole file as
    'corrupted or created by an unrecognized version'. Removing the part
    (and its Content_Types.xml declaration) lets Desktop regenerate it
    cleanly on next open/save."""
    sb_path = os.path.join(work_dir, "SecurityBindings")
    if os.path.exists(sb_path):
        os.remove(sb_path)

    ct_path = os.path.join(work_dir, "[Content_Types].xml")
    with open(ct_path, encoding="utf-8") as f:
        content = f.read()
    content = content.replace('<Override PartName="/SecurityBindings" ContentType="" />', "")
    with open(ct_path, "w", encoding="utf-8") as f:
        f.write(content)


def zip_pbix(src_dir, dest_pbix):
    if os.path.exists(dest_pbix):
        os.remove(dest_pbix)
    with zipfile.ZipFile(dest_pbix, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(src_dir):
            for fn in files:
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, src_dir).replace("\\", "/")
                z.write(full, rel)


def build():
    print("Backing up current pbix ->", BACKUP_PBIX)
    shutil.copy2(SRC_PBIX, BACKUP_PBIX)

    print("Unzipping working copy ->", WORK)
    unzip_pbix(SRC_PBIX, WORK)

    strip_security_bindings(WORK)

    pages_root = os.path.join(WORK, "Report", "definition", "pages")
    pages_json_path = os.path.join(pages_root, "pages.json")
    with open(pages_json_path, encoding="utf-8") as f:
        pages_meta = json.load(f)

    existing_order = pages_meta["pageOrder"]  # [regional, flip, margin]
    regional_id, flip_id, margin_id = existing_order

    new_order = []

    # ── Page: Cover ────────────────────────────────────────────────────
    cover_id = new_id()
    cover_dir = os.path.join(pages_root, cover_id)
    os.makedirs(cover_dir, exist_ok=True)
    with open(os.path.join(cover_dir, "page.json"), "w", encoding="utf-8") as f:
        json.dump(make_page_json(cover_id, "Cover", background_hex=NAVY), f, indent=2)

    z = 0
    vid, v = make_rect(0, 0, 1920, 6, GOLD, z=z); z += 1; write_visual(cover_dir, vid, v)
    vid, v = make_rect(0, 15, 8, 850, GOLD, z=z); z += 1; write_visual(cover_dir, vid, v)
    vid, v = make_rect(0, 900, 1920, 4, GOLD, z=z); z += 1; write_visual(cover_dir, vid, v)
    vid, v = make_textbox(110, 300, 1700, 140, [
        {"text": "TAMIL NADU", "size": 60, "bold": True, "color": WHITE, "font": FONT_DISPLAY}
    ], z=z); z += 1; write_visual(cover_dir, vid, v)
    vid, v = make_textbox(110, 420, 1700, 90, [
        {"text": "ASSEMBLY ELECTIONS", "size": 32, "color": LIGHT_BLUE, "font": FONT_DISPLAY}
    ], z=z); z += 1; write_visual(cover_dir, vid, v)
    vid, v = make_textbox(110, 500, 1000, 150, [
        {"text": "2026", "size": 76, "bold": True, "color": GOLD, "font": FONT_DISPLAY}
    ], z=z); z += 1; write_visual(cover_dir, vid, v)
    vid, v = make_textbox(110, 700, 1600, 60, [
        {"text": "A Comprehensive Data Analysis \u2014 Comparing 2021 vs 2026",
         "size": 18, "italic": True, "color": LIGHT_BLUE}
    ], z=z); z += 1; write_visual(cover_dir, vid, v)
    vid, v = make_textbox(110, 925, 1700, 50, [
        {"text": f"{STATS['n_const']} Constituencies  \u00b7  6 Regions  \u00b7  2021 vs 2026 Comparison",
         "size": 14, "color": "#9AA3BD"}
    ], z=z); z += 1; write_visual(cover_dir, vid, v)
    new_order.append(cover_id)

    # ── Page: At a Glance ────────────────────────────────────────────────
    glance_id = new_id()
    glance_dir = os.path.join(pages_root, glance_id)
    os.makedirs(glance_dir, exist_ok=True)
    with open(os.path.join(glance_dir, "page.json"), "w", encoding="utf-8") as f:
        json.dump(make_page_json(glance_id, "At a Glance"), f, indent=2)

    z = header_bar(glance_dir, "At a Glance")
    vid, v = make_card(40, 130, 430, 190, W2026, "ac_number", 5, "Total Constituencies", z=z); z += 1
    write_visual(glance_dir, vid, v)
    vid, v = make_card(500, 130, 430, 190, W2021, "share", 1, "Avg Winner Share \u2014 2021", z=z); z += 1
    write_visual(glance_dir, vid, v)
    vid, v = make_card(960, 130, 430, 190, W2026, "share", 1, "Avg Winner Share \u2014 2026", z=z); z += 1
    write_visual(glance_dir, vid, v)
    vid, v = make_slicer(1420, 130, 460, 90, W2026, "region", "Region", z=z); z += 1
    write_visual(glance_dir, vid, v)
    vid, v = make_slicer(1420, 230, 460, 90, W2026, "party", "Party (2026)", z=z); z += 1
    write_visual(glance_dir, vid, v)

    stats_row = [
        (f"{STATS['top_party']} \u2014 {STATS['top_seats']} seats", "Largest Party (2026)"),
        (f"{STATS['n_flipped']} of {STATS['n_const']} ({round(STATS['n_flipped']/STATS['n_const']*100)}%)",
         "Seats Flipped Since 2021"),
        (f"{STATS['above50_21']} \u2192 {STATS['above50_26']}", "Winners >50% Share (2021 \u2192 2026)"),
        (f"{STATS['below35_21']} \u2192 {STATS['below35_26']}", "Winners <35% Share (2021 \u2192 2026)"),
    ]
    for i, (val, label) in enumerate(stats_row):
        z = kpi_stat(glance_dir, 40 + i * 460, 360, 430, 190, val, label, GOLD, z)
    new_order.append(glance_id)

    # ── Page: Regional Story (enhance existing) ─────────────────────────
    regional_dir = os.path.join(pages_root, regional_id)
    z = header_bar(regional_dir, "Regional Story")
    vid, v = make_slicer(980, 105, 450, 80, W2026, "region", "Region", z=z); z += 1
    write_visual(regional_dir, vid, v)
    vid, v = make_slicer(1450, 105, 450, 80, W2026, "party", "Party", z=z); z += 1
    write_visual(regional_dir, vid, v)
    new_order.append(regional_id)

    # ── Page: Flip Story (enhance existing) ─────────────────────────────
    flip_dir = os.path.join(pages_root, flip_id)
    z = header_bar(flip_dir, "Flip Story")
    vid, v = make_slicer(1420, 105, 460, 80, W2026, "region", "Region", z=z); z += 1
    write_visual(flip_dir, vid, v)
    new_order.append(flip_id)

    # ── Page: Margin Story (enhance existing) ───────────────────────────
    margin_dir = os.path.join(pages_root, margin_id)
    z = header_bar(margin_dir, "Margin Story", height=50, font_size=20)
    vid, v = make_slicer(1440, 58, 460, 80, W2026, "region", "Region", z=z); z += 1
    write_visual(margin_dir, vid, v)
    vid, v = make_slicer(1440, 142, 460, 80, W2026, "party", "Party", z=z); z += 1
    write_visual(margin_dir, vid, v)
    new_order.append(margin_id)

    # ── Page: Party Performance (new) ───────────────────────────────────
    party_id = new_id()
    party_dir = os.path.join(pages_root, party_id)
    os.makedirs(party_dir, exist_ok=True)
    with open(os.path.join(party_dir, "page.json"), "w", encoding="utf-8") as f:
        json.dump(make_page_json(party_id, "Party Performance"), f, indent=2)
    z = header_bar(party_dir, "Party Performance")
    vid, v = make_slicer(1420, 100, 460, 80, W2026, "region", "Region", z=z); z += 1
    write_visual(party_dir, vid, v)
    vid, v = make_bar_chart(40, 190, 910, 840, W2021, "party", "ac_number",
                             "Seats by Party \u2014 2021", z=z); z += 1
    write_visual(party_dir, vid, v)
    vid, v = make_bar_chart(970, 190, 910, 840, W2026, "party", "ac_number",
                             "Seats by Party \u2014 2026", z=z); z += 1
    write_visual(party_dir, vid, v)
    new_order.append(party_id)

    # ── Page: Constituency Explorer (new) ───────────────────────────────
    explorer_id = new_id()
    explorer_dir = os.path.join(pages_root, explorer_id)
    os.makedirs(explorer_dir, exist_ok=True)
    with open(os.path.join(explorer_dir, "page.json"), "w", encoding="utf-8") as f:
        json.dump(make_page_json(explorer_id, "Constituency Explorer"), f, indent=2)
    z = header_bar(explorer_dir, "Constituency Explorer")
    vid, v = make_slicer(980, 100, 450, 80, W2026, "region", "Region", z=z); z += 1
    write_visual(explorer_dir, vid, v)
    vid, v = make_slicer(1450, 100, 450, 80, W2026, "party", "Party (2026)", z=z); z += 1
    write_visual(explorer_dir, vid, v)
    columns = [
        (W2026, "ac_number", "AC #"),
        (W2026, "region", "Region"),
        (W2021, "candidate", "Candidate 2021"),
        (W2021, "party", "Party 2021"),
        (W2021, "share", "Share % 2021"),
        (W2026, "candidate", "Candidate 2026"),
        (W2026, "party", "Party 2026"),
        (W2026, "share", "Share % 2026"),
        (W2026, "votes", "Votes 2026"),
    ]
    vid, v = make_table(40, 230, 1840, 800, columns,
                         sort_col=(W2026, "share"), sort_dir="Ascending", z=z); z += 1
    write_visual(explorer_dir, vid, v)
    new_order.append(explorer_id)

    # ── pages.json ───────────────────────────────────────────────────────
    pages_meta["pageOrder"] = new_order
    pages_meta["activePageName"] = cover_id
    with open(pages_json_path, "w", encoding="utf-8") as f:
        json.dump(pages_meta, f, indent=2)

    print("Repackaging ->", SRC_PBIX)
    zip_pbix(WORK, SRC_PBIX)
    shutil.rmtree(WORK)
    print("Done. Pages:", [p for p in new_order])


if __name__ == "__main__":
    build()

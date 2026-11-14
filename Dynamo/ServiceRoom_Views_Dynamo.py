# =============================================================================
# SERVICE ROOM CALLOUTS + SECTIONS  --  Dynamo Python Script node
# Project: AZHA0286 TIG Str Model (Revit 2024)
#
# HOW TO USE
#   1. In Revit: Manage tab > Dynamo
#   2. File > New > search "Python" > add a "Python Script" node
#      (choose the CPython3 engine if the node offers an engine choice)
#   3. Open the node's editor, DELETE everything, paste this whole file
#   4. Keep MODE = "SCAN" and press Run -> read the report in OUT
#      (SCAN only finds rooms in the ARC links and lists what WILL be created)
#   5. Set MODE = "CREATE" and Run -> creates views / sheets / viewports
#   6. Ctrl+Z once undoes everything (single transaction)
#
# WHAT IT CREATES (per room)
#   - 1 callout on the group's parent plan  -> view named exactly as the ROOM
#   - 2 building sections cut through the room centre:
#       "<room> - 1" looks NORTH, "<room> - 2" looks SOUTH (180 deg apart)
#   - viewports auto-packed on the group's sheet (column grid, bottom-up)
#   - ALL Revit links switched ON in every created view, MEP link halftone
#
# GROUPS / SHEETS
#   15th floor rooms  -> existing sheet S120-A
#   Typical (2ND)     -> new sheet S120-B
#   Ground floor      -> new sheet S120-C
#
# SAFETY
#   - Non-destructive: only adds views / sheets / viewports + per-view graphics
#   - RERUN_GUARD skips rooms whose callout view name existed BEFORE this run;
#     cross-floor duplicates created inside the same run get " (2)" suffix
#   - One transaction = one undo step
# =============================================================================

import clr
import traceback

clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import Room
from RevitServices.Persistence import DocumentManager
from RevitServices.Threading import TransactionManager
from System.Collections.Generic import List

doc = DocumentManager.Instance.CurrentDBDocument

# ----------------------------- CONFIG ----------------------------------------
MODE = "SCAN"          # "SCAN" or "CREATE"  <-- switch here after checking scan

RERUN_GUARD = True     # skip rooms whose callout view name already existed

CALLOUT_PAD_MM = 750   # margin around room bbox for the callout rectangle
SEC_HALF_PAD_MM = 1000 # margin left/right of the room on the section cut line
SEC_BELOW_MM    = 600  # section extends this far below floor level
SEC_DEPTH_MM    = 8000 # section far clip depth (view direction)
SEC_ABOVE_MM    = 1800 # section extends this far above the room's top
VIEW_GAP_X_MM   = 12   # horizontal gap between viewports
VIEW_GAP_Y_MM   = 22   # vertical gap (includes viewport label)
COL_GAP_MM      = 20   # gap between layout columns

SHEET_X0_MM     = 15.0 # usable drawing area on A1 (clear of title block)
SHEET_X1_MM     = 695.0
SHEET_Y0_MM     = 15.0
SHEET_Y1_MM     = 555.0

LEGEND_VIEW_ID  = 2170624   # "LAYOUT" legend copied onto NEW sheets
LEGEND_U_MM     = 784.0     # same spot as on S120-A
LEGEND_V_MM     = 430.0

# Parent plan type used by the sample sheet: "TIG_Structural Plan - Column Layout"
PREFERRED_PLAN_TYPE_ID = 1746036
# Section type used by the sample sheet: "Building Section"
PREFERRED_SECTION_TYPE_ID = 31116

ROOM_GROUPS = [
    {
        "name":    "15TH FLOOR",
        "parent":  "15TH FLOOR TECHNICAL Copy 1 Copy 1",  # plan hosting callouts
        "level":   "15TH FLOOR TECHNICAL",                # level for room search
        "sheet":   "S120-A",                              # existing sheet
        "sheet_name": None,           # None = do not rename existing sheet
        "scale":   100,
        "region_top_mm": None,        # None = auto: just below existing content
        "rooms": [
            "GSM Room",
            "Water Meter Room (WM Room)",
            "Electrical Room",
            "FHC Room",
            "Garbage Room",
        ],
    },
    {
        "name":    "TYPICAL FLOORS (2ND)",
        "parent":  "2ND FLOOR",
        "level":   "2ND FLOOR",
        "sheet":   "S120-B",          # new sheet
        "sheet_name": "TYPICAL FLOORS SERVICE ROOMS",
        "scale":   100,
        "region_top_mm": SHEET_Y1_MM,
        "rooms": [
            "Water Meter Room (WM Room)",
            "Electrical Room",
            "BTU Meter Room",
            "Garbage Room",
            "FHC Room",
        ],
    },
    {
        "name":    "GROUND FLOOR",
        "parent":  "GROUND FLOOR",     # the TIG_Structural Plan variant
        "level":   "GROUND FLOOR",
        "sheet":   "S120-C",          # new sheet
        "sheet_name": "GROUND FLOOR SERVICE ROOMS",
        "scale":   100,
        "region_top_mm": SHEET_Y1_MM,
        "rooms": [
            "LV Room",
            "Main Telephone Room",
            "Generator Room",
            "GSM Room",
            "FCC Room",
            "Main Garbage Room",
            "Water Meter Room",
            "BTU Meter Room",
            "Irrigation Pump Room",
        ],
    },
]

# ----------------------------- HELPERS ---------------------------------------
FT = 304.8  # mm per foot


def mm(v):
    return v / FT


def name_keys(s):
    """Comparison keys for a room name: full string, de-parenthesised string,
    and each parenthetical part (so 'Water Meter Room (WM Room)' can match
    a model room named either 'Water Meter Room' or 'WM Room')."""
    if not s:
        return set()
    full = " ".join(s.upper().split())
    no_par, parts, depth = [], [], 0
    for ch in full:
        if ch == "(":
            depth += 1
            no_par.append(" ")
            continue
        if ch == ")":
            depth -= 1
            no_par.append(" ")
            continue
        if depth == 0:
            no_par.append(ch)
        else:
            parts.append(ch if ch != ")" else " ")
    base = " ".join("".join(no_par).split())
    inner = " ".join("".join(parts).split())
    keys = set([full, base])
    if inner:
        # parenthetical may hold several words; also keep it whole
        keys.add(inner)
    keys.discard("")
    return keys


def bbox_to_host(bb, tr):
    """Transform a link-doc BoundingBoxXYZ to host min/max corners (feet)."""
    xs, ys, zs = [], [], []
    for x in (bb.Min.X, bb.Max.X):
        for y in (bb.Min.Y, bb.Max.Y):
            for z in (bb.Min.Z, bb.Max.Z):
                p = tr.OfPoint(XYZ(x, y, z))
                xs.append(p.X)
                ys.append(p.Y)
                zs.append(p.Z)
    return (min(xs), min(ys), min(zs)), (max(xs), max(ys), max(zs))


report = []
warnings = []


def log(msg):
    report.append(msg)


def warn(msg):
    warnings.append(msg)
    report.append("  !! " + msg)


# ----------------------------- NAME BOOKKEEPING ------------------------------
used_names = set()
for _v in FilteredElementCollector(doc).OfClass(View):
    if _v.Name:
        used_names.add(_v.Name)
for _s in FilteredElementCollector(doc).OfClass(ViewSheet):
    if _s.Name:
        used_names.add(_s.Name)
INITIAL_NAMES = set(used_names)   # snapshot: existed BEFORE this run


def unique_name(base):
    """Return base, or 'base (2)', 'base (3)' ... to keep names unique."""
    if base not in used_names:
        used_names.add(base)
        return base
    i = 2
    while "{0} ({1})".format(base, i) in used_names:
        i += 1
    r = "{0} ({1})".format(base, i)
    used_names.add(r)
    return r


# ----------------------------- LOOKUPS ---------------------------------------
def find_level(name):
    n = " ".join(name.upper().split())
    for lv in FilteredElementCollector(doc).OfClass(Level):
        if " ".join(lv.Name.upper().split()) == n:
            return lv
    return None


PLAN_VIEWS = []
for _v in FilteredElementCollector(doc).OfClass(View):
    try:
        if _v.ViewType in (ViewType.FloorPlan, ViewType.CeilingPlan,
                           ViewType.EngineeringPlan):
            PLAN_VIEWS.append(_v)
    except Exception:
        continue


def find_plan_view(name):
    """Find a plan view by name, preferring the TIG_Structural Plan type."""
    n = " ".join(name.upper().split())
    named = [v for v in PLAN_VIEWS if " ".join(v.Name.upper().split()) == n]
    for v in named:
        if v.GetTypeId().IntegerValue == PREFERRED_PLAN_TYPE_ID:
            return v
    return named[0] if named else None


def find_section_vft():
    cand = doc.GetElement(ElementId(PREFERRED_SECTION_TYPE_ID))
    if cand is not None and isinstance(cand, ViewFamilyType) \
            and cand.ViewFamily == ViewFamily.Section:
        return cand
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if vft.ViewFamily == ViewFamily.Section:
            return vft
    return None


def find_title_block():
    tbs = []
    for fs in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            if fs.Category and fs.Category.Id.IntegerValue == \
                    int(BuiltInCategory.OST_TitleBlocks):
                tbs.append(fs)
        except Exception:
            continue
    for fs in tbs:
        if " ".join(fs.Name.upper().split()) == "STRUCTURE-01":
            return fs
    return tbs[0] if tbs else None


def find_sheet(number):
    for s in FilteredElementCollector(doc).OfClass(ViewSheet):
        if s.SheetNumber == number:
            return s
    return None


def collect_links():
    """(all link instances, MEP link instance ids, [(arch instance, doc, name)])"""
    all_inst, mep_ids, arch = [], [], []
    for li in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        all_inst.append(li)
        nm = li.Name if li.Name else ""
        if "MEP" in nm.upper():
            mep_ids.append(li.Id)
        else:
            arch.append((li, li.GetLinkDocument(), nm))
    return all_inst, mep_ids, arch


def collect_rooms(arch_links):
    """Rooms from the loaded ARC links, in HOST coordinates (feet)."""
    rooms = []
    for li, ldoc, lname in arch_links:
        if ldoc is None:
            warn("link not loaded, skipped: '{0}'".format(lname))
            continue
        tr = li.GetTotalTransform()
        cnt = 0
        for r in FilteredElementCollector(ldoc) \
                .OfCategory(BuiltInCategory.OST_Rooms) \
                .WhereElementIsNotElementType():
            if not isinstance(r, Room):
                continue
            loc = r.Location
            if loc is None or not isinstance(loc, LocationPoint):
                continue
            bb = r.get_BoundingBox(None)
            if bb is None:
                continue
            (x0, y0, z0), (x1, y1, z1) = bbox_to_host(bb, tr)
            rooms.append({
                "name": r.Name,
                "number": r.Number,
                "link": lname,
                "level_name": r.Level.Name if r.Level else "",
                "min": (x0, y0, z0),
                "max": (x1, y1, z1),
                "cx": (x0 + x1) / 2.0,
                "cy": (y0 + y1) / 2.0,
            })
            cnt += 1
        log("  link '{0}': {1} placed rooms collected".format(lname, cnt))
    return rooms


def match_room(target, rooms, level_name, level_elev_ft):
    """Best room match on a level. Returns (room_dict, note) or (None, note).
    Level test: room's link level NAME equals target, OR the room's HOST-space
    floor Z is within 1.5 m of the host level elevation."""
    tkeys = name_keys(target)
    best, best_score, ties, best_note = None, 0, [], ""
    for rm in rooms:
        lvl_ok, lvl_note = False, ""
        if level_name and " ".join(rm["level_name"].upper().split()) == \
                " ".join(level_name.upper().split()):
            lvl_ok, lvl_note = True, "level name"
        elif level_elev_ft is not None and \
                abs(rm["min"][2] - level_elev_ft) < 4.92:   # 1.5 m
            lvl_ok, lvl_note = True, "level by elevation"
        if not lvl_ok:
            continue
        rkeys = name_keys(rm["name"])
        score = 0
        for tk in tkeys:
            for rk in rkeys:
                if tk == rk:
                    score = max(score, 100)
                elif len(tk) > 2 and len(rk) > 2 and (tk in rk or rk in tk):
                    score = max(score, 50)
        if score > best_score:
            best, best_score, ties = rm, score, [rm]
            best_note = lvl_note
        elif score == best_score and score > 0:
            ties.append(rm)
    if best is None:
        return None, "no match on level '{0}'".format(level_name)
    if len(ties) > 1:
        desc = ", ".join(["{0} (no.{1})".format(t["name"], t["number"])
                          for t in ties])
        return best, "{0} pts, AMBIGUOUS: {1} -> used first".format(
            best_score, desc)
    return best, "{0} pts via {1}".format(best_score, best_note)


# ----------------------------- GRAPHICS PER VIEW -----------------------------
def ensure_links_on(v, link_instances, mep_ids):
    """Show all links in the view; halftone the MEP link."""
    try:
        v.SetCategoryHidden(ElementId(BuiltInCategory.OST_RvtLinks), False)
    except Exception as e:
        warn("view '{0}': Revit Links category: {1}".format(
            v.Name, str(e)[:120]))
    hidden = [li.Id for li in link_instances if li.IsHidden(v)]
    if hidden:
        try:
            v.UnhideElements(List[ElementId](hidden))
        except Exception as e:
            warn("view '{0}': unhide links: {1}".format(v.Name, str(e)[:120]))
    for li in link_instances:
        try:
            v.SetWorksetVisibility(li.WorksetId, WorksetVisibility.Visible)
        except Exception:
            pass
    for mid in mep_ids:
        try:
            ogs = OverrideGraphicSettings()
            ogs = ogs.SetHalftone(True)
            v.SetElementOverrides(mid, ogs)
        except Exception as e:
            warn("view '{0}': MEP halftone: {1}".format(v.Name, str(e)[:120]))


# ----------------------------- SECTION BOX -----------------------------------
def make_section_box(cx, cy, floor_z, half_w_ft, top_ft, depth_ft, face_north):
    """Section through (cx, cy) at floor level; looks north or south.
    Cut plane through the room centre; up = +Z; far clip = depth."""
    t = Transform.Identity
    t.Origin = XYZ(cx, cy, floor_z)
    if face_north:
        t.BasisZ = XYZ(0.0, 1.0, 0.0)     # view direction: north
    else:
        t.BasisZ = XYZ(0.0, -1.0, 0.0)    # view direction: south
    t.BasisY = XYZ(0.0, 0.0, 1.0)         # up
    t.BasisX = t.BasisZ.CrossProduct(t.BasisY)   # completes left-handed set
    bb = BoundingBoxXYZ()
    bb.Transform = t
    bb.Min = XYZ(-half_w_ft, mm(-SEC_BELOW_MM), 0.0)
    bb.Max = XYZ(half_w_ft, top_ft, depth_ft)
    return bb


# ----------------------------- VIEWPORT LAYOUT -------------------------------
def pack_viewports(items, region):
    """Column-major, bottom-up packing. items: [(key, w_mm, h_mm)].
    Returns [(key, center_u_mm, center_v_mm)]."""
    x0, y0, x1, y1 = region
    x, y, col_w = x0, y1, 0.0
    placed = []
    for key, w, h in items:
        cell_h = h + VIEW_GAP_Y_MM
        if y - cell_h < y0:
            x += col_w + COL_GAP_MM     # start a new column
            y = y1
            col_w = 0.0
        if x + w > x1 + 0.5:
            warn("layout: '{0}' may extend past the drawing area "
                 "(u={1:.0f} > {2:.0f}) - drag it into place".format(
                     str(key), x + w, x1))
        y -= cell_h
        placed.append((key, x + w / 2.0, y + h / 2.0 + VIEW_GAP_Y_MM / 2.0))
        if w > col_w:
            col_w = w
        y -= VIEW_GAP_Y_MM
    return placed


# ----------------------------- SCAN ------------------------------------------
def run_scan():
    log("SCAN MODE - nothing is modified")
    all_inst, mep_ids, arch = collect_links()
    log("Revit links: {0} total | MEP (halftone): {1} | arch: {2}".format(
        len(all_inst), len(mep_ids), len(arch)))
    rooms = collect_rooms(arch)
    log("Rooms available in arch links: {0}".format(len(rooms)))
    sv = find_section_vft()
    tb = find_title_block()
    log("Section view family type : {0}".format(sv.Name if sv else "NOT FOUND!"))
    log("Title block for new sheets: {0}".format(tb.Name if tb else "NOT FOUND!"))
    log("-" * 70)
    for g in ROOM_GROUPS:
        log("GROUP: {0}".format(g["name"]))
        pv = find_plan_view(g["parent"])
        if pv:
            vt = doc.GetElement(pv.GetTypeId())
            log("  parent plan : '{0}' -> id {1} ({2})".format(
                g["parent"], pv.Id.IntegerValue, vt.Name if vt else "?"))
        else:
            log("  parent plan : '{0}' -> NOT FOUND - fix config!".format(
                g["parent"]))
        lv = find_level(g["level"])
        if lv:
            log("  level       : '{0}' -> id {1}, elev {2:.2f} m".format(
                g["level"], lv.Id.IntegerValue, lv.Elevation * FT / 1000.0))
        else:
            log("  level       : '{0}' -> NOT FOUND - fix config!".format(
                g["level"]))
        sh = find_sheet(g["sheet"])
        if sh:
            log("  sheet       : '{0}' -> exists (id {1}, '{2}')".format(
                g["sheet"], sh.Id.IntegerValue, sh.Name))
        else:
            log("  sheet       : '{0}' -> WILL BE CREATED as '{1}'".format(
                g["sheet"], g["sheet_name"]))
        for target in g["rooms"]:
            rm, note = match_room(target, rooms, g["level"],
                                  lv.Elevation if lv else None)
            if rm:
                log("  [OK]   '{0}' -> room '{1}' no.{2} @ ({3:.0f}, {4:.0f}) mm "
                    "[{5}] ({6})".format(
                        target, rm["name"], rm["number"],
                        rm["cx"] * FT, rm["cy"] * FT, rm["link"], note))
            else:
                log("  [MISS] '{0}' -> {1}".format(target, note))
        log("")
    log("-" * 70)
    log("If everything looks right: set MODE = \"CREATE\" and Run again.")


# ----------------------------- CREATE ----------------------------------------
def run_create():
    all_inst, mep_ids, arch = collect_links()
    rooms = collect_rooms(arch)
    section_vft = find_section_vft()
    title_block = find_title_block()
    if section_vft is None:
        raise Exception("No Section ViewFamilyType found - nothing created.")
    log("CREATE MODE - adding views/sheets/viewports (single undo step)")
    log("links: {0} | MEP halftone ids: {1}".format(
        len(all_inst), [i.IntegerValue for i in mep_ids]))

    for g in ROOM_GROUPS:
        log("GROUP: {0}".format(g["name"]))
        pv = find_plan_view(g["parent"])
        if pv is None:
            warn("group {0}: parent plan '{1}' NOT FOUND - group skipped".format(
                g["name"], g["parent"]))
            continue
        lv = find_level(g["level"])
        lvl_elev = lv.Elevation if lv else None
        scale = g["scale"]

        # ---- sheet ----
        sheet = find_sheet(g["sheet"])
        if sheet is None:
            if title_block is None:
                warn("group {0}: no title block type - sheet {1} NOT created".format(
                    g["name"], g["sheet"]))
                continue
            sheet = ViewSheet.Create(doc, title_block.Id)
            sheet.SheetNumber = g["sheet"]
            if g["sheet_name"]:
                sheet.Name = g["sheet_name"]
            log("  created sheet {0} '{1}' (id {2})".format(
                g["sheet"], sheet.Name, sheet.Id.IntegerValue))
            leg = doc.GetElement(ElementId(LEGEND_VIEW_ID))
            if leg is not None:
                try:
                    lvp = Viewport.Create(doc, sheet.Id, leg.Id,
                                          XYZ(mm(LEGEND_U_MM), mm(LEGEND_V_MM), 0))
                    lvp.SetBoxCenter(XYZ(mm(LEGEND_U_MM), mm(LEGEND_V_MM), 0))
                    log("  legend 'LAYOUT' placed at ({0:.0f}, {1:.0f}) mm".format(
                        LEGEND_U_MM, LEGEND_V_MM))
                except Exception as e:
                    warn("legend on {0}: {1}".format(g["sheet"], str(e)[:120]))
        else:
            log("  using existing sheet {0} '{1}' (id {2})".format(
                g["sheet"], sheet.Name, sheet.Id.IntegerValue))

        # ---- sheet region (mm) ----
        ry1 = g["region_top_mm"]
        if ry1 is None:
            mins = []
            for vp in FilteredElementCollector(sheet).OfClass(Viewport):
                try:
                    mins.append(vp.GetBoxOutline().Min.Y * FT)
                except Exception:
                    pass
            ry1 = max(SHEET_Y0_MM, min(mins) - 15.0) if mins else SHEET_Y1_MM
            log("  region: below existing content, top = {0:.0f} mm".format(ry1))
        region = (SHEET_X0_MM, SHEET_Y0_MM, SHEET_X1_MM, ry1)

        # ---- per-room views ----
        callout_vft_id = pv.GetTypeId()
        layout_items = []
        n_new = 0
        for target in g["rooms"]:
            rm, note = match_room(target, rooms, g["level"], lvl_elev)
            if rm is None:
                warn("group {0}: '{1}' not found ({2}) - skipped".format(
                    g["name"], target, note))
                continue
            base = rm["name"].strip() if rm["name"].strip() else target
            if RERUN_GUARD and base in INITIAL_NAMES:
                log("  [SKIP] '{0}': view '{1}' already exists".format(target, base))
                continue
            x0, y0f, z0 = rm["min"]
            x1, y1f, z1 = rm["max"]
            try:
                # ---- callout on the parent plan ----
                cview = ViewSection.CreateCallout(
                    doc, pv.Id, callout_vft_id,
                    XYZ(x0 - mm(CALLOUT_PAD_MM), y0f - mm(CALLOUT_PAD_MM), 0),
                    XYZ(x1 + mm(CALLOUT_PAD_MM), y1f + mm(CALLOUT_PAD_MM), 0))
                cview.Name = unique_name(base)
                try:
                    cview.Scale = scale
                except Exception as e:
                    warn("callout '{0}': scale not set (view template may "
                         "control it): {1}".format(cview.Name, str(e)[:100]))
                try:
                    cview.DetailLevel = ViewDetailLevel.Fine
                except Exception:
                    pass
                ensure_links_on(cview, all_inst, mep_ids)

                # ---- 2 sections: - 1 looks north, - 2 looks south ----
                half_w = (x1 - x0) / 2.0 + mm(SEC_HALF_PAD_MM)
                top_ft = (z1 - z0) + mm(SEC_ABOVE_MM)
                depth_ft = mm(SEC_DEPTH_MM)
                sec_views = []
                for i, north in enumerate([True, False]):
                    box = make_section_box(rm["cx"], rm["cy"], z0, half_w,
                                           top_ft, depth_ft, north)
                    sv = ViewSection.CreateSection(doc, section_vft.Id, box)
                    sv.Name = unique_name("{0} - {1}".format(base, i + 1))
                    try:
                        sv.Scale = scale
                    except Exception as e:
                        warn("section '{0}': scale not set: {1}".format(
                            sv.Name, str(e)[:100]))
                    try:
                        sv.DetailLevel = ViewDetailLevel.Fine
                    except Exception:
                        pass
                    ensure_links_on(sv, all_inst, mep_ids)
                    sec_views.append(sv)

                # ---- layout sizes (estimate in mm) ----
                cw = ((x1 - x0) * FT + 2 * CALLOUT_PAD_MM) / scale
                ch = ((y1f - y0f) * FT + 2 * CALLOUT_PAD_MM) / scale
                sw = (2 * half_w * FT) / scale
                sh = ((top_ft + mm(SEC_BELOW_MM)) * FT) / scale
                layout_items.append((("callout", cview), cw, ch))
                layout_items.append((("sec N", sec_views[0]), sw, sh))
                layout_items.append((("sec S", sec_views[1]), sw, sh))
                n_new += 1
                log("  [NEW] '{0}': callout + 2 sections (room no.{1})".format(
                    base, rm["number"]))
            except Exception as e:
                warn("group {0} room '{1}': {2}".format(
                    g["name"], target, str(e)[:200]))

        # ---- place viewports on the sheet ----
        placed = pack_viewports(layout_items, region)
        n_placed = 0
        for key, cu, cv in placed:
            kind, view = key
            try:
                vp = Viewport.Create(doc, sheet.Id, view.Id,
                                     XYZ(mm(cu), mm(cv), 0))
                vp.SetBoxCenter(XYZ(mm(cu), mm(cv), 0))
                n_placed += 1
            except Exception as e:
                warn("viewport '{0}' ({1}): {2}".format(
                    view.Name, kind, str(e)[:160]))
        log("  rooms processed: {0} | viewports placed: {1}".format(
            n_new, n_placed))
        log("")


# ----------------------------- MAIN ------------------------------------------
try:
    if MODE.strip().upper() == "CREATE":
        TransactionManager.Instance.EnsureInTransaction(doc)
        try:
            run_create()
        finally:
            TransactionManager.Instance.TransactionTaskDone()
    else:
        run_scan()
    if warnings:
        report.append("")
        report.append("== {0} WARNING(S) - see '!!' lines above ==".format(
            len(warnings)))
except Exception:
    report.append("FAILED:")
    report.append(traceback.format_exc())
    try:
        TransactionManager.Instance.TransactionTaskDone()
    except Exception:
        pass

OUT = report

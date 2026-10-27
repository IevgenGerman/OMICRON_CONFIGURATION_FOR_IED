# -*- coding: utf-8 -*-
"""
omicron_xrio.py - OMICRON Test Universe test objects (.xrio) from the General_PB project
(author: German Ievgen).

  python omicron_xrio.py                      -> asks for the project folder
  python omicron_xrio.py "D:\\Projects\\Bucha"  -> project folder given
  options: --device W1G_D60 (repeatable) | --all

Project folder:
  General_PB.xlsm               Settings_IED, Device_Type, GEN, SV, project name / address
  Library\\test_object_default.xrio   master test object (saved once in Test Universe)
  DOC\\<IED>.xlsm               sheets Omicron AI <test>: column F = SV ID (subscriptions)
Result (one test object per Omicron AI sheet with subscriptions):
  Omicron\\<IED>\\test_object_<IED>_AI_<test>.xrio
  a device without any SV subscription: ONE file, nominal values left as in the master
  Omicron\\<IED>\\test_object_<IED>.xrio
  an existing file is never overwritten: a new one is written next to it with the date and
  time in the name (test_object_<IED>_AI_<test>_<YYYY-MM-DD_HHMMSS>.xrio)
  Omicron\\_xrio_report.txt

What is written into the master (only the <Value> of these parameters, the rest stays as it is):
  NAMEPLATE  DEVICE_MODEL        Name/description   <- IEDName
             MANUFACTURER        Manufacturer       <- Device_Type Manufacturer of the device type
             DEVICE_TYPE         Device type        <- GEN Order code (else Device_Type order_code)
  LOCATION   DEVICE_ADDRESS      Device address     <- GEN IP SB (else IP PB)
             SUBSTATION          Substation name    <- Project name   (workbook name PROJECT_NAME)
             SUBSTATION_ADDRESS  Substation address <- Project address (workbook name PROJECT_ADDRESS)
             BAY                 Bay name           <- Settings_IED BAY
  NOMINALVALUES   FNOM, VNOM, VPRIM_LL, INOM, IPRIM        <- SV columns N..W of the subscribed
  RESIDUALFACTORS VLNVN, ININOM, VNSEC, INSEC,                streams (all streams of one test
                  RESIDUAL_INSTRUMENT_TRANSFORMER            must have the same values)
     VPRIM_LL = Vn primary f-f kV * 1000
     VLNVN    = (Vneutral primary kV * 1000 / Vneutral secondary V) / (VPRIM_LL / VNOM)
     ININOM   = (IPRIM / INOM) / (Ineutral primary A / Ineutral secondary A)
A test whose streams have different or empty nominal values gets NO file (message + report).
Needs: Python 3, openpyxl (pip install openpyxl).
"""
import os, re, sys, glob, datetime, argparse, warnings

warnings.filterwarnings("ignore")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import omicron_configuration as oc            # noqa: E402

UserError = oc.UserError
OUT_DIR = oc.OUT_DIR
REPORT = "_xrio_report.txt"
MASTER = "test_object_default.xrio"

# SV columns N..W (in this order) = SVCB<n> nominal values of Device_Type
NOM = [("F", "F, Hz"), ("VN", "Vn primary f-f, kV"), ("VS", "Vn secondary f-f, V"), ("IP", "In primary, A"),
       ("IS", "In secondary, A"), ("RIT", "Instrumental transformer"), ("NVP", "Vneutral primary, kV"),
       ("NVS", "Vneutral secondary, V"), ("NIP", "Ineutral primary, A"), ("NIS", "Ineutral secondary, A")]
NOM_FIRST = 13                                   # column N (0-based) if the header is not found


def T(v):
    return oc.T(v)


# ---------------------------------------------------------------- values
def num(v):
    """number of a cell (number or text with . or ,); None if empty / not a number"""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def boolean(v):
    if isinstance(v, bool):
        return v
    s = T(v).upper()
    if s in ("TRUE", "ИСТИНА", "ІСТИНА", "1", "YES", "ТАК", "X"):
        return True
    if s in ("FALSE", "ЛОЖЬ", "ХИБНІСТЬ", "0", "NO", "НІ"):
        return False
    return None


def fmt(x):
    """number as written into the xrio: 110000 / 1.7914887887639932 / true"""
    if isinstance(x, bool):
        return "true" if x else "false"
    if abs(x - round(x)) < 1e-9 * max(1.0, abs(x)):
        return str(int(round(x)))
    # v0.0.4: like Test Universe (.NET "R"): 15 significant digits when that reads back
    # to the same number, else 17 (5.1955610465195559, 33.333)
    x = float(x)
    s = format(x, ".15g")
    if float(s) != x:
        s = format(x, ".17g")
    return s


def same(a, b):
    if isinstance(a, bool) or isinstance(b, bool):
        return a is b
    return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))


def show(key, x):
    return ("TRUE" if x else "FALSE") if key == "RIT" else fmt(x)


# ---------------------------------------------------------------- General_PB
def _header(rows, row_index, names):
    head = [T(x).upper() for x in rows[row_index]]
    for n in names:
        if n.upper() in head:
            return head.index(n.upper())
    return -1


def _defined(wb, name):
    try:
        dn = wb.defined_names[name]
    except Exception:
        dn = None
        try:
            dn = wb.defined_names.get(name)
        except Exception:
            pass
    if dn is None:
        return ""
    v = T(getattr(dn, "attr_text", "") or getattr(dn, "value", ""))
    if v.startswith("="):
        v = v[1:]
    if len(v) >= 2 and v[0] == '"' and v[-1] == '"':
        v = v[1:-1].replace('""', '"')
    return v.strip()


def read_gpb(gpb):
    """dict: project, address, ieds {IED: {bay, type}}, types {TYPE: {man, oc}}, gen {IED: {sb, pb, oc}},
    sv {SV ID upper: {svid, ied, nom: {key: value or None}}}"""
    wb, tmp = oc.open_wb(gpb)
    try:
        d = dict(project=_defined(wb, "PROJECT_NAME"), address=_defined(wb, "PROJECT_ADDRESS"),
                 ieds={}, types={}, gen={}, sv={})
        # Settings_IED: header row with IEDName
        rows = list(wb["Settings_IED"].iter_rows(values_only=True))
        hr = next((i for i, r in enumerate(rows[:20]) if "IEDNAME" in [T(x).upper() for x in r]), None)
        if hr is None:
            raise UserError("General_PB Settings_IED: column IEDName not found")
        cI, cB, cT = _header(rows, hr, ["IEDName"]), _header(rows, hr, ["BAY"]), _header(rows, hr, ["DeviceType"])
        for r in rows[hr + 1:]:
            r = list(r) + [None] * 10
            if T(r[cI]):
                d["ieds"][T(r[cI]).upper()] = dict(ied=T(r[cI]), bay=T(r[cB]) if cB >= 0 else "",
                                                   type=T(r[cT]) if cT >= 0 else "")
        # Device_Type: sub-headers on row 3, types from row 4 (column A)
        rows = list(wb["Device_Type"].iter_rows(values_only=True))
        cM, cO = _header(rows, 2, ["Manufacturer"]), _header(rows, 2, ["order_code"])
        for r in rows[3:]:
            r = list(r) + [None] * 5
            if T(r[0]):
                d["types"][T(r[0]).upper()] = dict(man=T(r[cM]) if cM >= 0 else "", oc=T(r[cO]) if cO >= 0 else "")
        # GEN: the IP / order code of a device are on its first row
        rows = list(wb["GEN"].iter_rows(values_only=True))
        cI = _header(rows, 0, ["IEDName"])
        cSB, cPB = _header(rows, 0, ["IP  SB", "IP SB"]), _header(rows, 0, ["IP  PB", "IP PB"])
        cOC = _header(rows, 0, ["Order code"])
        for r in rows[1:]:
            r = list(r) + [None] * 5
            k = T(r[cI]).upper() if cI >= 0 else ""
            if k and k not in d["gen"]:
                d["gen"][k] = dict(sb=T(r[cSB]) if cSB >= 0 else "", pb=T(r[cPB]) if cPB >= 0 else "",
                                   oc=T(r[cOC]) if cOC >= 0 else "")
        # SV: nominal values N..W
        rows = list(wb["SV"].iter_rows(values_only=True))
        cI, cS = _header(rows, 0, ["IEDName"]), _header(rows, 0, ["SV ID"])
        if cS < 0:
            raise UserError("General_PB SV: column SV ID not found")
        cF = _header(rows, 0, ["F, Hz"])
        if cF < 0:
            cF = NOM_FIRST
        ied = ""
        for r in rows[1:]:
            r = list(r) + [None] * 30
            if cI >= 0 and T(r[cI]):
                ied = T(r[cI])
            sid = T(r[cS])
            if not sid:
                continue
            nom = {}
            for j, (key, _) in enumerate(NOM):
                v = r[cF + j]
                nom[key] = boolean(v) if key == "RIT" else num(v)
            d["sv"][sid.upper()] = dict(svid=sid, ied=ied, nom=nom)
        return d
    finally:
        wb.close()
        if tmp:
            os.remove(tmp)


# ---------------------------------------------------------------- xrio (text patch)
def _span(text, tag, a=0, z=None):
    """(start, end) of the element <tag ...>...</tag> inside text[a:z]"""
    z = len(text) if z is None else z
    m = re.compile(r"<%s(\s[^>]*)?>" % re.escape(tag)).search(text, a, z)
    if not m:
        return None
    e = text.find("</%s>" % tag, m.end(), z)
    if e < 0:
        return None
    return m.start(), e + len(tag) + 3


def _xml(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def set_value(text, path, value):
    """replaces the <Value> of the parameter path (e.g. RIO/DEVICE/NAMEPLATE/DEVICE_MODEL)"""
    a, z = 0, len(text)
    for tag in path.split("/"):
        sp = _span(text, tag, a, z)
        if not sp:
            raise ValueError("%s not found in the master xrio" % path)
        a, z = sp
    # the own <Value> of the parameter (not of a child element)
    m = re.compile(r"<Value\s*/>|<Value>(.*?)</Value>", re.S).search(text, a, z)
    if not m:
        raise ValueError("%s: no <Value> in the master xrio" % path)
    new = "<Value>%s</Value>" % _xml(value) if value != "" else "<Value />"
    return text[:m.start()] + new + text[m.end():]


DEV = "RIO/DEVICE/"
FIELDS = [("DEVICE_MODEL", DEV + "NAMEPLATE/DEVICE_MODEL"), ("MANUFACTURER", DEV + "NAMEPLATE/MANUFACTURER"),
          ("DEVICE_TYPE", DEV + "NAMEPLATE/DEVICE_TYPE"), ("DEVICE_ADDRESS", DEV + "LOCATION/DEVICE_ADDRESS"),
          ("SUBSTATION", DEV + "LOCATION/SUBSTATION"), ("SUBSTATION_ADDRESS", DEV + "LOCATION/SUBSTATION_ADDRESS"),
          ("BAY", DEV + "LOCATION/BAY")]
NOMV = DEV + "NOMINALVALUES/"
RESF = DEV + "RESIDUALFACTORS/"


def nominal_params(n):
    """xrio parameters from the 10 SV values; raises ValueError for impossible values"""
    vprim = n["VN"] * 1000.0
    if n["VS"] == 0 or n["IS"] == 0 or n["NVS"] == 0 or n["NIS"] == 0 or n["NIP"] == 0 or vprim == 0:
        raise ValueError("a nominal value is 0 (division by 0)")
    # v0.0.4: the same order of operations as Test Universe (the last digit is the same)
    vlnvn = n["NVP"] * 1000.0 * n["VS"] / (n["NVS"] * vprim)
    ininom = n["IP"] * n["NIS"] / (n["IS"] * n["NIP"])
    return [(NOMV + "FNOM", n["F"]), (NOMV + "VNOM", n["VS"]), (NOMV + "VPRIM_LL", vprim),
            (NOMV + "INOM", n["IS"]), (NOMV + "IPRIM", n["IP"]),
            (RESF + "VLNVN", vlnvn), (RESF + "ININOM", ininom), (RESF + "VNSEC", n["NVS"]),
            (RESF + "INSEC", n["NIS"]), (RESF + "RESIDUAL_INSTRUMENT_TRANSFORMER", n["RIT"])]


def read_master(path):
    raw = open(path, "rb").read()
    bom = raw.startswith(b"\xef\xbb\xbf")
    text = raw[3:].decode("utf-8") if bom else raw.decode("utf-8")
    return text, bom


def write_xrio(path, text, bom):
    with open(path, "wb") as f:                        # bytes: line ends / tabs of the master kept
        f.write((b"\xef\xbb\xbf" if bom else b"") + text.encode("utf-8"))


# ---------------------------------------------------------------- one test
def check_streams(ied, tname, svids, g):
    """nominal values of a test or None + problems: all streams must be filled and equal"""
    probs = []
    noms = []
    for sid in svids:
        s = g["sv"].get(sid.upper())
        if not s:
            probs.append("%s / %s: SV ID %s not found on General_PB SV" % (ied, tname, sid))
            continue
        empty = [h for (k, h) in NOM if s["nom"][k] is None]
        if empty:
            probs.append("%s / %s: SV ID %s - nominal values empty on General_PB SV: %s" %
                         (ied, tname, s["svid"], "all (columns N..W)" if len(empty) == len(NOM)
                          else " | ".join(empty)))
            continue
        noms.append((s["svid"], s["nom"]))
    if probs:
        return None, probs
    ref_id, ref = noms[0]
    for sid, n in noms[1:]:
        diff = ["%s %s %s <> %s %s" % (h, ref_id, show(k, ref[k]), sid, show(k, n[k]))
                for (k, h) in NOM if not same(ref[k], n[k])]
        if diff:
            probs.append("%s / %s: SV streams have different nominal values - test object cannot be "
                         "created for this configuration: %s" % (ied, tname, "; ".join(diff)))
    return (None if probs else ref), probs


def free_name(folder, fn, stamp):
    p = os.path.join(folder, fn)
    if not os.path.exists(p):
        return fn, False
    stem, ext = os.path.splitext(fn)
    return "%s_%s%s" % (stem, stamp, ext), True


# ---------------------------------------------------------------- run
def list_devices(prj):
    return oc.list_devices(prj)


def master_path(prj):
    lib = os.path.join(prj, "Library")
    p = os.path.join(lib, MASTER)
    if os.path.isfile(p):
        return p
    c = sorted(glob.glob(os.path.join(lib, "*.xrio")))
    if len(c) == 1:
        return c[0]
    raise UserError("Master test object not found: %s" % p)


def run(prj, devices, log=print):
    """test objects of the given devices; returns (files written, report lines, problems)"""
    gpb, doc, lib, out = oc.project_paths(prj)
    mpath = master_path(prj)
    mtext, bom = read_master(mpath)
    try:                                                # the master must have all parameters
        for _, path in FIELDS:
            set_value(mtext, path, "")
        for path, _ in nominal_params(dict(F=50, VN=110, VS=100, IP=1, IS=1, RIT=True, NVP=1, NVS=1, NIP=1, NIS=1)):
            set_value(mtext, path, "")
    except ValueError as ex:
        raise UserError("Master %s: %s" % (os.path.basename(mpath), ex))
    log("Master: %s" % os.path.basename(mpath))
    log("Reading General_PB ...")
    g = read_gpb(gpb)
    os.makedirs(out, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    rep, notes, probs, made = [], [], [], 0
    if not g["project"]:
        notes.append("Project name is empty (General_PB Editor - Project - Project name)")
    if not g["address"]:
        notes.append("Project address is empty (General_PB Editor - Project - Project address)")
    for ied in devices:
        log("  " + ied)
        try:
            dev = oc.read_device(os.path.join(doc, ied + ".xlsm"))
        except Exception as ex:
            probs.append("%s: device file not read - %s" % (ied, ex))
            continue
        info = g["ieds"].get(ied.upper())
        if not info:
            notes.append("%s: not on General_PB Settings_IED - BAY / type / IP / order code empty" % ied)
            info = dict(ied=ied, bay="", type="")
        typ = g["types"].get(info["type"].upper(), {}) if info["type"] else {}
        if info["type"] and not typ:
            notes.append("%s: device type %s not on Device_Type - Manufacturer empty" % (ied, info["type"]))
        gen = g["gen"].get(ied.upper(), {})
        ocode = gen.get("oc") or typ.get("oc", "")
        ip = gen.get("sb") or gen.get("pb") or ""
        plate = dict(DEVICE_MODEL=ied, MANUFACTURER=typ.get("man", ""), DEVICE_TYPE=ocode, DEVICE_ADDRESS=ip,
                     SUBSTATION=g["project"], SUBSTATION_ADDRESS=g["address"], BAY=info["bay"])
        tests, nosv = [], []
        for tname, m in dev["tests"]:
            svids = list(dict.fromkeys(v[1] for v in m.values() if v[1]))
            if svids:
                tests.append((tname, svids))
            else:
                nosv.append(tname)
        folder = os.path.join(out, ied)
        if not tests:
            # no SV subscription at all: ONE test object, only nameplate / location,
            # the nominal values stay as in the master
            try:
                text = mtext
                for key, path in FIELDS:
                    text = set_value(text, path, plate[key])
                os.makedirs(folder, exist_ok=True)
                fn, dated = free_name(folder, "test_object_%s.xrio" % ied, stamp)
                write_xrio(os.path.join(folder, fn), text, bom)
                made += 1
                rep.append("%-60s nominal values from the master (no SV subscriptions)%s" %
                           (ied + "\\" + fn, "  (file existed - new file with date)" if dated else ""))
            except ValueError as ex:
                probs.append("%s: %s" % (ied, ex))
                rep.append("%-14s NOT CREATED (see PROBLEMS)" % ied)
            continue
        for tname in nosv:
            notes.append("%s / %s: no SV subscriptions - skipped (the device has tests with SV)" % (ied, tname))
        for tname, svids in tests:
            n, pr = check_streams(ied, tname, svids, g)
            if pr:
                probs += pr
                rep.append("%-14s %-10s NOT CREATED (see PROBLEMS)  SV %s" % (ied, tname, ", ".join(svids)))
                continue
            try:
                text = mtext
                for key, path in FIELDS:
                    text = set_value(text, path, plate[key])
                for path, x in nominal_params(n):
                    text = set_value(text, path, fmt(x))
                os.makedirs(folder, exist_ok=True)
                fn = "test_object_%s_AI_%s.xrio" % (ied, re.sub(r"[^\w\-]+", "_", tname).strip("_"))
                fn, dated = free_name(folder, fn, stamp)
                write_xrio(os.path.join(folder, fn), text, bom)
                made += 1
                rep.append("%-60s SV %s%s" % (ied + "\\" + fn, ", ".join(svids),
                                              "  (file existed - new file with date)" if dated else ""))
            except ValueError as ex:
                probs.append("%s / %s: %s" % (ied, tname, ex))
                rep.append("%-14s %-10s NOT CREATED (see PROBLEMS)" % (ied, tname))
    lines = ["omicron_xrio  %s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
             "Project : %s" % prj, "Master  : %s" % os.path.basename(mpath), "Output  : %s" % out,
             "Substation: %s / %s" % (g["project"] or "-", g["address"] or "-"), ""] + rep + \
            (["", "PROBLEMS (no file):"] + ["  " + p for p in probs] if probs else []) + \
            (["", "NOTES:"] + ["  " + n for n in notes] if notes else []) + ["", "%d file(s) written." % made]
    with open(os.path.join(out, REPORT), "w", encoding="utf-8-sig") as f:
        f.write("\n".join(lines) + "\n")
    return made, lines, probs


# ---------------------------------------------------------------- command line
def main():
    ap = argparse.ArgumentParser(description="OMICRON test objects (.xrio) from the General_PB project")
    ap.add_argument("project", nargs="?")
    ap.add_argument("--device", action="append", help="IED name (repeatable)")
    ap.add_argument("--all", action="store_true", help="all devices, no question")
    a = ap.parse_args()
    code = 0
    try:
        prj = a.project or oc.ask_folder()
        ieds = list_devices(prj)
        if not ieds:
            raise UserError("No device files in DOC")
        if a.device:
            sel = [i for i in ieds if i.lower() in [d.lower() for d in a.device]]
        else:
            sel = ieds if a.all else oc.choose(ieds, "Devices in DOC:")
        made, lines, probs = run(prj, sel)
        print("\n" + "\n".join(lines))
        code = 2 if probs else 0
    except UserError as ex:
        print("\n" + str(ex))
        code = 1
    if not a.project:
        input("Enter - close")
    return code


if __name__ == "__main__":
    sys.exit(main())

# Auto-generated: DEEL 1 notebook code (cleaned)
def copy_storylist(wb, src_name, dst_name):
    if src_name not in wb.sheetnames:
        return
    if dst_name in wb.sheetnames:
        wb.remove(wb[dst_name])
    wb.copy_worksheet(wb[src_name]).title = dst_name

def overwrite_row_by_name(ws, name, vals):
    # Overschrijf bestaande rij (op Naam productie), behoud kolomstructuur
    headers = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
    name_col = None
    for i,h in enumerate(headers, start=1):
        if str(h).strip().lower() == "naam productie":
            name_col = i
            break
    if name_col is None:
        return
    for r in range(2, ws.max_row+1):
        if normalize(ws.cell(r, name_col).value) == normalize(name):
            for c,h in enumerate(headers, start=1):
                if h in vals:
                    ws.cell(r,c).value = vals[h]
            return
# Deze cel voert de matching-engine uit en schrijft Krantenplanning.xlsx weg.
# Daarna wordt het bestand automatisch aangeboden voor download.


def _header_index(ws):
    hdr = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
    norm = [normalize(h) for h in hdr]
    return hdr, {n:i+1 for i,n in enumerate(norm)}

def _get_col(idx_map, *names):
    for nm in names:
        key = normalize(nm)
        if key in idx_map:
            return idx_map[key]
    return None

def normalize(x):
    return "" if x is None else str(x).strip()

def truthy(v):
    if v is None: return False
    if isinstance(v, bool): return v
    if isinstance(v, (int, float)): return v != 0
    return str(v).strip().lower() in ("true","ja","1","yes","waar")

def cell_set(v):
    # Splits cellen zoals "S_nws_0; S_nws_2;" naar set({"S_nws_0","S_nws_2"})
    if v is None: return set()
    return set([p.strip() for p in str(v).split(";") if p.strip()])

def find_col_any(df, substrings, required=True):
    for sub in substrings:
        for c in df.columns:
            if sub.lower() in str(c).lower():
                return c
    raise KeyError(f"Geen kolom gevonden voor: {substrings}")

def _pretty_class_token(t):
    t = normalize(t)
    if t == "a-keus":
        return "A-keus"
    if t == "b-keus":
        return "B-keus"
    return t


def _split_multi(val):
    if val is None:
        return set()
    s = str(val).strip()
    if s == "":
        return set()
    # accepteer ; , + / en ' en ' als scheiding
    parts = re.split(r"[;,+/]|en", s, flags=re.IGNORECASE)
    return {normalize(p) for p in parts if normalize(p)}

def class_allowed(pc, allowed):
    # allowed kan zijn "Alle" of bv. "A-keus" of meerdere waarden
    if allowed is None or str(allowed).strip()=="":
        return True
    if str(allowed).strip().lower()=="alle":
        return True
    allowed_set = _split_multi(allowed)
    pc_set = _split_multi(pc)
    # match zodra er overlap is
    return len(pc_set & allowed_set) > 0


def parse_range(s):
    if s is None: return None
    s=str(s).strip()
    m=re.match(r"^\s*(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)\s*$", s)
    if not m: return None
    return (float(m.group(1)), float(m.group(2)))

# -----------------------------
# Load workbooks (data_only=True is essentieel i.v.m. formules zoals Posities!M1)
# -----------------------------
plan_wb = openpyxl.load_workbook(VERHALENAANBOD_PATH, data_only=True)
pos_wb  = openpyxl.load_workbook(POSITIES_PATH, data_only=True)
tmpl_wb = openpyxl.load_workbook(TEMPLATES_PATH, data_only=True)

pos_ws   = pos_wb["Blad1"]
stats_ws = plan_wb["Stats"]
# Helper: always resolve the current Logfile worksheet (important after fallback reset)
def get_log_ws():
    if "Logfile" not in plan_wb.sheetnames:
        plan_wb.create_sheet("Logfile", 0)
        ws = plan_wb["Logfile"]
        ws.append(["Timestamp","Beschrijving"])
    return plan_wb["Logfile"]

# Reset logfile (laat header staan)
_log_ws = get_log_ws()
if _log_ws.max_row > 1:
    _log_ws.delete_rows(2, _log_ws.max_row)

def log(msg):
    ws = get_log_ws()
    row = ws.max_row + 1
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.cell(row=row, column=1, value=ts)
    ws.cell(row=row, column=2, value=msg)

# -----------------------------
# Planning order
# -----------------------------
po_ws = plan_wb["Planningsvolgorde"]
order = []
r = 2
while True:
    v = po_ws[f"A{r}"].value
    if v is None or str(v).strip()=="":
        break
    order.append(str(v).strip())
    r += 1

# -----------------------------
# Regime berekenen
# -----------------------------
Pgs_uniek = float(pos_ws["M1"].value)   # cached value van formule
B_kar     = float(stats_ws["C2"].value)
B_dwang   = float(stats_ws["D2"].value)

UkpUp = B_kar / Pgs_uniek
UkpUp_dwang = B_dwang / Pgs_uniek

regime = "Normaal"
if UkpUp < 4350:
    regime = "Verhalenschaarste"
elif UkpUp_dwang > 5400:
    regime = "Papierschaarste"

log(f"Regime={regime}, UkpUp={UkpUp:.2f}, UkpUp_dwang={UkpUp_dwang:.2f}, Pgs_uniek={int(Pgs_uniek)}")

# -----------------------------
# Templates inlezen
# -----------------------------
tmpl_ws = tmpl_wb["Blad1"]
tmpl_headers = [tmpl_ws.cell(1,c).value for c in range(1, tmpl_ws.max_column+1)]
def tcol(name): return tmpl_headers.index(name)+1

templates = []
for rr in range(2, tmpl_ws.max_row+1):
    tpl = tmpl_ws.cell(rr, tcol("Template")).value
    if tpl is None: 
        continue
    placeholders = []
    for i in range(1,6):
        v = tmpl_ws.cell(rr, tcol(f"Placeholder {i}")).value
        if v not in (None,""):
            placeholders.append(str(v).strip())
    templates.append({
        "Template": str(tpl).strip(),
        "Templatesoort": normalize(tmpl_ws.cell(rr, tcol("Templatesoort")).value),
        "Placeholders": placeholders,
        "Advertentiepositie": normalize(tmpl_ws.cell(rr, tcol("Advertentiepositie")).value),
        "Akpp": tmpl_ws.cell(rr, tcol("Akpp")).value,
    })

# -----------------------------
# Beslispaden inlezen (pandas)
# -----------------------------
bps = pd.read_excel(BESLISPAD_SPREAD_PATH, sheet_name="Blad1")
bpe = pd.read_excel(BESLISPAD_EP_PATH, sheet_name="Blad1")

def map_cols(df):
    # Robust tegen aangepaste kolomkoppen (zoals jouw update)
    return {
        "mode": find_col_any(df, ["Bovenste producties"]),
        "skip_pos": find_col_any(df, ["Sla deze stap over bij deze posities"], required=False),
        "class": find_col_any(df, ["Toegestane Classificatie"]),
        "tsoort": find_col_any(df, ["Toegestane 'Templatesoort'", "Templatesoort"]),
        # Spread-kop kan door updates veranderen, daarom zoeken we breed:
        "tw1_if": find_col_any(df, ["Beeld voor print=Bijplaat", "wordt voldaan", "toegestaan if"]),
        "tw1": find_col_any(df, ["Bij maximaal 1 van de placeholders op het template 'Tweede keus placeholder' toegestaan "]),
        "tw2": find_col_any(df, ["Bij maximaal 2 van de placeholders"]),
        "open_xs": find_col_any(df, ["XS_0 open"]),
        "open_s": find_col_any(df, ["S_nws_0 of S_lk_0 open", "S_nws_0"]),
        "open_custom": find_col_any(df, ["Toegestaan om maximaal 1 placeholder op te laten van onderstaande soort(en)"]),
        "derde": find_col_any(df, ["Derde keus placeholder"]),
        "derde_max1": find_col_any(df, ["Bij maximaal 1 van de placeholders op het template \'Derde keus placeholder\' toegestaan"]),
        "vierde": find_col_any(df, ["Bij maximaal 1 van de placeholders op het template 'VIERDE keus placeholder' toegestaan"]),
        "admatch": find_col_any(df, ["Advertentiepositie' matchen", "Advertentiepositie"]),
        "norm": find_col_any(df, ["Regime=Normaal"]),
        "pap": find_col_any(df, ["Regime=Papierschaarste"]),
        "ver": find_col_any(df, ["Regime=Verhalenschaarste"]),
        "conct": find_col_any(df, ["Concessies_beschreven"]),
    }

cols_bps = map_cols(bps)
cols_bpe = map_cols(bpe)

# -----------------------------
# Stats ranges (Ukpp_range_*) lookup per tab
# -----------------------------
# -----------------------------
# Stats ranges (Ukpp_range_* / Akpp_range_*) lookup per tab
# -----------------------------
stats_headers = [stats_ws.cell(1,c).value for c in range(1, stats_ws.max_column+1)]

def _norm_header(x):
    return "" if x is None else re.sub(r"\s+", "", str(x)).lower()

def stats_col_ci(wanted_name):
    # case/space-insensitive kolomzoeker
    wn = _norm_header(wanted_name)
    for idx, h in enumerate(stats_headers, start=1):
        if _norm_header(h) == wn:
            return idx
    raise ValueError(wanted_name)

def stats_row_for(tabname):
    # 1) exact match op positiesheetnaam in kolom A
    tn = normalize(tabname)
    for rr in range(2, stats_ws.max_row+1):
        if normalize(stats_ws.cell(rr,1).value) == tn:
            return rr

    # 2) fallback: een rij waarvan kolom A 'verhalenlijst' bevat (ongeacht exact label)
    for rr in range(2, stats_ws.max_row+1):
        label = normalize(stats_ws.cell(rr,1).value).lower()
        if "verhalenlijst" in label:
            return rr

    # 3) fallback: 'totale verhalenlijst' (oude naam)
    for rr in range(2, stats_ws.max_row+1):
        if normalize(stats_ws.cell(rr,1).value).lower() == "totale verhalenlijst":
            return rr

    # 4) laatste redmiddel: eerste datarij
    return 2

def resolve_akpp_range(code_or_range, tabname):
    # accepteer direct 'min:max' of codes zoals Akpp_range_normaal_extra
    if code_or_range is None:
        return None
    s = str(code_or_range).strip()
    direct = parse_range(s)
    if direct:
        return direct

    # code -> kolomnaam in Stats
    if s.lower().startswith("akpp_range_"):
        # voorkeur: Ukpp_range_*
        suffix = "Ukpp_range_" + s[len("Akpp_range_"):]
        candidates = [suffix, s]  # probeer ook Akpp_range_* zelf
        rr = stats_row_for(tabname)
        for cand in candidates:
            try:
                cc = stats_col_ci(cand)
            except ValueError:
                continue
            val = stats_ws.cell(rr, cc).value
            rng = parse_range(val)
            if rng:
                return rng

        # als we hier komen: code bestaat maar kolom/waarde niet gevonden/parsable
        log(f"WAARSCHUWING: Akpp-range code '{s}' niet kunnen oplossen in Stats voor '{tabname}'. Akpp-filter wordt overgeslagen.")
        return None

    return None

# -----------------------------
# Posities kolommen
# -----------------------------
pos_header = [pos_ws.cell(1,c).value for c in range(1, pos_ws.max_column+1)]
def pos_col(sub):
    for i,h in enumerate(pos_header, start=1):
        if h and sub.lower() in str(h).lower():
            return i
    raise KeyError(sub)

POS_VORM_COL = pos_col("Verschijningsvorm")
POS_POS_COL  = pos_col("Positie")
POS_AD1_COL  = pos_col("Advertentieaanbod")
POS_AD2_COL  = pos_col("tweede keus")
POS_AD3_COL  = pos_col("derde keus")
POS_AD4_COL  = pos_col("vierde keus")

def get_pos_row(posname):
    rr=2
    while True:
        v = pos_ws.cell(rr, POS_POS_COL).value
        if v is None or str(v).strip()=="":
            return None
        if str(v).strip()==posname:
            return rr
        rr += 1

# -----------------------------
# Sheet <-> DataFrame
# -----------------------------
def sheet_to_df(ws):
    headers = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
    data=[]
    for rr in range(2, ws.max_row+1):
        row = [ws.cell(rr,c).value for c in range(1, ws.max_column+1)]
        if all(v is None or str(v).strip()=="" for v in row):
            continue
        data.append(row)
    return pd.DataFrame(data, columns=headers)

def df_to_sheet(ws, df):
    ws.delete_rows(2, ws.max_row)
    for i,row in enumerate(df.itertuples(index=False), start=2):
        for j,val in enumerate(row, start=1):
            ws.cell(i,j,value=val)

def get_copy_targets(ws):
    # Tabbladen om NAAR TE KOPIËREN (stap 7) – uit cel AA1
    v = ws["AA1"].value
    return [s.strip() for s in re.split(r"[;,]", str(v)) if s.strip()] if v else []

def get_keep_targets(ws):
    # Tabbladen die NIET gestript mogen worden (stap 8) – uit cel AB1
    v = ws["AB1"].value
    return [s.strip() for s in re.split(r"[;,]", str(v)) if s.strip()] if v else []

    for rr in range(2, ws.max_row+1):
        if normalize(ws.cell(rr,1).value)==name:
            for c,val in enumerate(vals, start=1):
                ws.cell(rr,c,value=val)
            return
    rr = ws.max_row + 1
    for c,val in enumerate(vals, start=1):
        ws.cell(rr,c,value=val)

def remove_from_sheet(ws, name):
    for rr in range(2, ws.max_row+1):
        if normalize(ws.cell(rr,1).value)==name:
            ws.delete_rows(rr,1)
            return

# -----------------------------
# Matching per stap
# -----------------------------
def best_match_for_step(tabname, df_run, step_row, is_spread):
    cols = cols_bps if is_spread else cols_bpe

    mode          = step_row[cols["mode"]]
    is_bovenste   = str(mode).lower().startswith("bovenste")
    allowed_class = step_row[cols["class"]]
    allowed_tsoort= step_row[cols["tsoort"]]
    tw1_if        = truthy(step_row[cols["tw1_if"]])
    tw1           = truthy(step_row[cols["tw1"]])
    tw2           = truthy(step_row[cols["tw2"]])
    open_xs       = truthy(step_row[cols["open_xs"]])
    open_s        = truthy(step_row[cols["open_s"]])
    open_custom_raw = step_row[cols["open_custom"]] if "open_custom" in cols else None
    open_custom_set = cell_set(open_custom_raw)
    derde         = truthy(step_row[cols["derde"]])
    derde_max1    = truthy(step_row[cols["derde_max1"]])
    vierde        = truthy(step_row[cols["vierde"]])
    admatch       = normalize(step_row[cols["admatch"]])
    conc_txt      = step_row[cols["conct"]]

    rng_cell = step_row[ {"Normaal":cols["norm"], "Papierschaarste":cols["pap"], "Verhalenschaarste":cols["ver"]}[regime] ]
    rng_pair = resolve_akpp_range(rng_cell, tabname)

    pr = get_pos_row(tabname)
    ad1 = normalize(pos_ws.cell(pr, POS_AD1_COL).value) if pr else ""
    ad2 = normalize(pos_ws.cell(pr, POS_AD2_COL).value) if pr else ""
    ad3 = normalize(pos_ws.cell(pr, POS_AD3_COL).value) if pr else ""
    ad4 = normalize(pos_ws.cell(pr, POS_AD4_COL).value) if pr else ""
    ad_choice = {"Advertentieaanbod":ad1, "Advertentieaanbod_tweede keus":ad2, "Advertentieaanbod_derde keus":ad3, "Advertentieaanbod_vierde keus":ad4}.get(admatch, ad1)

    # IF Advertentieaanbod != W00 THEN range-min -1000
    if rng_pair and ad_choice and ad_choice!="W00":
        rng_pair = (rng_pair[0]-1000, rng_pair[1])

    def template_allowed(t):
        if allowed_tsoort:
            # Sta ";" en "," toe als scheiding in Toegestane Templatesoort
            allowed_tsoort_set = {normalize(p) for p in re.split(r"[;,]", str(allowed_tsoort)) if str(p).strip()}
            if normalize(t["Templatesoort"]) not in allowed_tsoort_set:
                return False
        if ad_choice and normalize(t["Advertentiepositie"]) != ad_choice:
            return False
        if rng_pair:
            try:
                ak = float(t["Akpp"])
            except:
                return False
            return rng_pair[0] <= ak <= rng_pair[1]
        return True

    tpls = [t for t in templates if template_allowed(t)]
    if not tpls:
        return None

    pool = df_run.copy()
    if "Classificatie" in pool.columns:
        pool = pool[pool["Classificatie"].apply(lambda x: class_allowed(x, allowed_class))]
    pool = pool.reset_index(drop=True)
    if pool.empty:
        return None

    max_tw  = 2 if tw2 else (1 if (tw1 or tw1_if) else 0)
    max_der = 0 if not derde else (1 if derde_max1 else len(phs))
    max_v   = 1 if vierde else 0

    # Precompute placeholder sets
    for col in ["Gewenste placeholder","Tweede keus placeholder","Derde keus placeholder","Vierde keus placeholder"]:
        pool[col+"_set"] = pool[col].apply(cell_set) if col in pool.columns else [set()]*len(pool)

    def matches(prow, ph, kind):
        return ph in prow[{"g":"Gewenste placeholder_set","t":"Tweede keus placeholder_set","d":"Derde keus placeholder_set","v":"Vierde keus placeholder_set"}[kind]]

    best = None
    for t in tpls:
        phs = t["Placeholders"]
        k = len(phs)
        if is_bovenste:
            # Neem bovenste N producties (N = aantal placeholders),
            # maar bij ex aequo Prioscore: neem ook alle extra producties met dezelfde Prioscore
            pool_use = pool
            if k < len(pool_use):
                try:
                    boundary = float(pool_use.loc[k-1].get("Prioscore", None))
                    last = k-1
                    while last + 1 < len(pool_use):
                        nxt = pool_use.loc[last+1].get("Prioscore", None)
                        try:
                            nxt_f = float(nxt)
                        except Exception:
                            break
                        if nxt_f == boundary:
                            last += 1
                        else:
                            break
                    pool_use = pool_use.iloc[:last+1]
                except Exception:
                    pool_use = pool_use.iloc[:k]
        else:
            pool_use = pool

        def _bovenste_ok(selected_indices):
            # Regels voor 'Bovenste' met ex aequo:
            # - als je r producties gebruikt, dan mag je NIET een hogere prioscore overslaan.
            # - je mag alleen wisselen binnen de boundary-tie (zelfde Prioscore als de r-de productie).
            try:
                r = len(selected_indices)
                if r == 0:
                    return True
                boundary = float(pool.loc[r-1].get("Prioscore", None))
                # verplicht: alle indices met Prioscore > boundary moeten geselecteerd zijn (prefix)
                mandatory_count = 0
                for j in range(len(pool)):
                    try:
                        pj = float(pool.loc[j].get("Prioscore", None))
                    except Exception:
                        break
                    if pj > boundary:
                        mandatory_count += 1
                    else:
                        break
                for j in range(mandatory_count):
                    if j not in selected_indices:
                        return False
                # toegestaan: indices t/m last waarbij Prioscore == boundary (ties)
                last = r-1
                while last + 1 < len(pool):
                    try:
                        nxt = float(pool.loc[last+1].get("Prioscore", None))
                    except Exception:
                        break
                    if nxt == boundary:
                        last += 1
                    else:
                        break
                return all((i <= last) for i in selected_indices)
            except Exception:
                # fallback: strict prefix-regel zonder ties
                r = len(selected_indices)
                if r == 0:
                    return True
                return set(range(r)).issubset(set(selected_indices))

        # Candidate list per placeholder
        candidates=[]
        for ph in phs:
            c=[]
            for idx, prow in pool_use.iterrows():
                if matches(prow, ph, "g"):
                    c.append((idx,"g")); 
                    continue
                if max_tw>0:
                    if tw1_if:
                        # Let op: dit blijft exact "bijplaat" (zoals in de huidige output-logica)
                        if normalize(prow.get("Beeld voor print","")).lower()=="bijplaat" and matches(prow, ph, "t"):
                            c.append((idx,"t"))
                    else:
                        if matches(prow, ph, "t"):
                            c.append((idx,"t"))
                if max_der>0 and matches(prow, ph, "d"):
                    c.append((idx,"d"))
                if max_v>0 and matches(prow, ph, "v"):
                    c.append((idx,"v"))
            if ph=="XS_0" and open_xs:
                c.append((None,"o"))
            if ph in ("S_nws_0","S_lk_0") and open_s:
                c.append((None,"o"))
            if open_custom_set and ph in open_custom_set:
                c.append((None,"o"))
            candidates.append(c)

        if any(len(c)==0 for c in candidates):
            continue

        # Brute force (max 5 placeholders)
        best_for_tpl=None
        for choice in itertools.product(*candidates):
            idxs=[i for i,_k in choice if i is not None]
            if len(idxs)!=len(set(idxs)): 
                continue  # productie mag niet dubbel
            if sum(1 for _i,k in choice if k=="t")>max_tw: 
                continue
            if sum(1 for _i,k in choice if k=="d")>max_der:
                continue
            if sum(1 for _i,k in choice if k=="v")>max_v:
                continue

            # MAXIMAAL 1× S_nws_0 of S_lk_0 open laten (samen)
            open_s_count = sum(
                1 for slot,(i,k) in enumerate(choice)
                if i is None and k=="o" and phs[slot] in ("S_nws_0","S_lk_0")
            )
            if open_s_count > 1:
                continue

            # MAXIMAAL 1× open laten uit custom lijst (uit beslispad)
            if open_custom_set:
                custom_open_count = sum(
                    1 for slot,(i,k) in enumerate(choice)
                    if i is None and k=="o" and phs[slot] in open_custom_set
                )
                if custom_open_count > 1:
                    continue

            prios=[]
            for i,_k in choice:
                if i is None: 
                    continue
                try:
                    prios.append(float(pool_use.loc[i].get("Prioscore",0)))
                except:
                    prios.append(0.0)
            if not prios:
                continue

            metric=(sum(prios)/len(prios), sum(prios))  # avg, sum
            if best_for_tpl is None or metric>best_for_tpl["metric"]:
                best_for_tpl={"metric":metric, "choice":choice, "pool_use":pool_use}

        if best_for_tpl and (best is None or best_for_tpl["metric"]>best["metric"]):
            best={"metric":best_for_tpl["metric"], "template":t, "choice":best_for_tpl["choice"], "pool_use":best_for_tpl["pool_use"], "conct":conc_txt}

    return best

# -----------------------------
# Execute pipeline
# -----------------------------
success = 0

def run_runs(runs):
    global success
    for posname in runs:
        if posname not in plan_wb.sheetnames:
            log(f"Run {posname}: tabblad ontbreekt.")
            continue

        pr = get_pos_row(posname)
        vorm = normalize(pos_ws.cell(pr, POS_VORM_COL).value) if pr else ""

        if vorm.lower()=="niet":
            plan_wb.remove(plan_wb[posname])
            log(f"Run {posname}: Verschijningsvorm=Niet -> verwijderd.")
            continue

        is_spread = (vorm.lower()=="spread")
        beslis = bps if is_spread else bpe
        cols_step = cols_bps if is_spread else cols_bpe

        ws = plan_wb[posname]
        df = sheet_to_df(ws)
        copy_targets = get_copy_targets(ws)
        keep_targets = get_keep_targets(ws)

        matched=False
        for _, step in beslis.iterrows():
            # Optioneel: sla deze stap over voor specifieke posities
            if "skip_pos" in cols_step and cols_step["skip_pos"] is not None:
                raw = step[cols_step["skip_pos"]]
                skip_set = {normalize(s) for s in re.split(r"[;,]", str(raw)) if str(s).strip()} if raw not in (None, "") else set()
                if normalize(posname) in skip_set:
                    continue

            bm = best_match_for_step(posname, df, step, is_spread)
            if bm is None:
                continue

            tpl = bm["template"]
            phs = tpl["Placeholders"]
            choice = bm["choice"]
            pool_use = bm["pool_use"]

            selected=[]
            for slot,(idx,kind) in enumerate(choice):
                if idx is None:
                    continue
                prod = pool_use.loc[idx].to_dict()
                selected.append((normalize(prod.get("Naam productie","")), phs[slot]))

            names_real=[n for n,_ph in selected]

            # Placeholder(s) die open blijven (idx=None)
            open_placeholders=[phs[slot] for slot,(idx,kind) in enumerate(choice) if idx is None]

            df_new = df[df["Naam productie"].astype(str).str.strip().isin(names_real)].copy()
            for col in ["Gekozen template","Gekozen placeholder","Plaatsing"]:
                if col not in df_new.columns:
                    df_new[col]=None

            for i,row in df_new.iterrows():
                nm = normalize(row["Naam productie"])
                df_new.at[i,"Gekozen template"] = tpl["Template"]
                df_new.at[i,"Gekozen placeholder"] = next((ph for n,ph in selected if n==nm), "")
                df_new.at[i,"Plaatsing"] = posname

            # Voeg extra rij(en) toe als er placeholder(s) open blijven
            if open_placeholders:
                # Safety: voorkom dubbele kolomnamen (pandas concat vereist unieke columns)
                if df_new.columns.duplicated().any():
                    df_new = df_new.loc[:, ~df_new.columns.duplicated()].copy()
                for oph in open_placeholders:
                    empty_row={c:None for c in df_new.columns}
                    empty_row["Naam productie"]="NOG ARTIKEL VOOR DEZE PLEK ZOEKEN"
                    empty_row["Gekozen template"]=tpl["Template"]
                    empty_row["Gekozen placeholder"]=oph
                    empty_row["Plaatsing"]=posname
                    df_new = pd.concat([df_new, pd.DataFrame([empty_row])], ignore_index=True)

            df_to_sheet(ws, df_new.reset_index(drop=True))

            
            # Copy / post-processing (standaard vs. UITW-runs)
            name_to_vals={}
            headers = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
            for rr in range(2, ws.max_row+1):
                nm = normalize(ws.cell(rr,1).value)
                if nm:
                    name_to_vals[nm] = {headers[c-1]: ws.cell(rr,c).value for c in range(1, ws.max_column+1)}

            is_uitw_run = ("-U" in posname)

            if not is_uitw_run:
                # Standaard: kopieer naar targets uit AA1 en overschrijf daar de rij
                for tr in copy_targets:
                    if tr in plan_wb.sheetnames:
                        wst = plan_wb[tr]
                        for nm,_ph in selected:
                            overwrite_row_by_name(wst, nm, name_to_vals[nm])

                # Verwijder gematchte producties van alle andere tabbladen behalve keep_targets (AB1)
                exclude=set([posname]+keep_targets)
                for sh in plan_wb.sheetnames:
                    if sh in exclude:
                        continue
                    wso = plan_wb[sh]
                    for nm in names_real:
                        remove_from_sheet(wso, nm)

            else:
                # UITW-runs: append naar Totale verhalenlijst en strip alleen binnen de UITW-groep (NM-U* of ZU-U*)
                if "Totale verhalenlijst" in plan_wb.sheetnames:
                    wtot = plan_wb["Totale verhalenlijst"]
                    tot_headers = [wtot.cell(1,c).value for c in range(1, wtot.max_column+1)]
                    tot_h2c = {str(h).strip(): i+1 for i,h in enumerate(tot_headers) if h is not None and str(h).strip() != ""}

                    for nm,_ph in selected:
                        vals = name_to_vals.get(nm, {})
                        new_row_idx = wtot.max_row + 1
                        for h, col in tot_h2c.items():
                            if h in vals:
                                wtot.cell(new_row_idx, col).value = vals[h]

                prefix_u = posname[:2] + "-U"  # "NM-U" of "ZU-U"
                for sh in plan_wb.sheetnames:
                    if sh == posname or sh == "Totale verhalenlijst":
                        continue
                    if not str(sh).startswith(prefix_u):
                        continue
                    wso = plan_wb[sh]
                    for nm in names_real:
                        remove_from_sheet(wso, nm)

            log(f"Run {posname}: stap {step['Stappen']} succesvolle match. Template={tpl['Template']}. {bm['conct']}")
            matched=True
            success += 1
            break

        if not matched:
            log(f"Run {posname}: {beslis.iloc[0]['Stappen']} tot en met {beslis.iloc[-1]['Stappen']} geen match.")



# Volgens Opzet Krantenplanner DL (afgeslankt):
# 1) Eerste 5 runs: volgorde uit Planningsvolgorde A2:A6
runs_1 = order[:5]

# 2) Daarna: ND-01, ND-02, ND-03
runs_2 = ["ND-01", "ND-02", "ND-03"]

all_runs = runs_1 + runs_2

run_runs(runs_1)

# 2) Daarna: ND-01, ND-02, ND-03
run_runs(runs_2)

# 3) Fallback: check succesvolle match op ND-01; zo niet, reset workbook (behalve Logfile) en herhaal runs
def _get_last_logline_for_run(run_name):
    if "Logfile" not in plan_wb.sheetnames:
        return ""
    ws = plan_wb["Logfile"]
    # zoek kolom 'Beschrijving'
    hdr = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
    desc_col = None
    for i,h in enumerate(hdr, start=1):
        if h is None:
            continue
        if str(h).strip().lower() == "beschrijving":
            desc_col = i
            break
    if desc_col is None:
        desc_col = 2

    last = ""
    for r in range(2, ws.max_row+1):
        v = ws.cell(r, desc_col).value
        if v is None:
            continue
        s = str(v)
        if s.startswith(f"Run {run_name}:"):
            last = s
    return last

def _is_success_logline(line):
    if not line:
        return False
    return ("succesvolle match" in line.lower()) or ("succesvol bij stap" in line.lower())

def _snapshot_logfile():
    if "Logfile" not in plan_wb.sheetnames:
        return None
    ws = plan_wb["Logfile"]
    data=[]
    for r in ws.iter_rows(values_only=True):
        data.append(list(r))
    return data

def _restore_logfile(data):
    if data is None:
        return
    if "Logfile" not in plan_wb.sheetnames:
        plan_wb.create_sheet("Logfile", 0)
    ws = plan_wb["Logfile"]
    if ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)
    for row in data:
        ws.append(row)

def _reset_workbook_keep_logfile(logdata):
    global plan_wb
    plan_wb = openpyxl.load_workbook(VERHALENAANBOD_PATH, data_only=True)
    _restore_logfile(logdata)

nd01_line = _get_last_logline_for_run("ND-01")
if not _is_success_logline(nd01_line):
    log("Fallback geactiveerd: ND-01 had geen succesvolle match. Reset workbook (behalve Logfile) en herhaal runs.")
    _logdata = _snapshot_logfile()
    _reset_workbook_keep_logfile(_logdata)
    success = 0

    # nieuwe volgorde: ND-01, dan de 5 runs uit Planningsvolgorde, dan ND-02, ND-03
    run_runs(["ND-01"])
    run_runs(runs_1)
    run_runs(["ND-02", "ND-03"])


# 3) Maak 2 kopieën van tabblad 'Totale verhalenlijst' -> 'ZU-UITW' en 'NM-UITW'
def _ensure_fresh_copy(src_name: str, dest_name: str):
    if dest_name in plan_wb.sheetnames:
        plan_wb.remove(plan_wb[dest_name])
    base_ws = plan_wb[src_name]
    new_ws = plan_wb.copy_worksheet(base_ws)
    new_ws.title = dest_name
    return new_ws

zu_ws = _ensure_fresh_copy("Totale verhalenlijst", "ZU-UITW")
nm_ws = _ensure_fresh_copy("Totale verhalenlijst", "NM-UITW")

def _sheet_to_rows(ws):
    headers = [c.value for c in ws[1]]
    # Map header -> index
    h2i = {str(h).strip(): i for i, h in enumerate(headers) if h is not None and str(h).strip() != ""}
    data = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None or str(v).strip()=="" for v in r):
            continue
        row = {h: r[i] if i < len(r) else None for h, i in h2i.items()}
        data.append(row)
    return headers, h2i, data

def _rewrite_sheet(ws, headers, rows, h2i):
    # Clear all rows except header
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row-1)
    # Write rows
    for ridx, row in enumerate(rows, start=2):
        for h, col_idx0 in h2i.items():
            ws.cell(row=ridx, column=col_idx0+1, value=row.get(h))

def _contains_any(text, needles):
    if text is None:
        return False
    s = str(text)
    return any(n in s for n in needles)

def apply_mappingregels(ws, mode: str):
    headers, h2i, rows = _sheet_to_rows(ws)

    # Vereiste kolommen
    required = [
        "Plaatsing", "Heel Limburg", "Focusregio",
        "Gewenste placeholder", "Placeholder bij enigszins geschikt",
        "Prioscore", "Gekozen template", "Gekozen placeholder", "Publicatiedwang", "Top 8"
    ]
    missing = [c for c in required if c not in h2i]
    if missing:
        raise ValueError(f"Mappingregels {mode}: ontbrekende kolommen op {ws.title}: {missing}")

    if mode == "ZU":
        # 1) Schrap: Plaatsing begint met 'ND-' of 'ZU-'
        def keep_r(r):
            p = "" if r.get("Plaatsing") is None else str(r.get("Plaatsing"))
            return not (p.startswith("ND-") or p.startswith("ZU-"))
        rows = [r for r in rows if keep_r(r)]

        # 2) Schrap: Heel Limburg=ongeschikt AND Focusregio bevat géén Parkstad/Maastricht/Sittard
        def keep_r2(r):
            hl = "" if r.get("Heel Limburg") is None else str(r.get("Heel Limburg")).strip()
            fr = r.get("Focusregio")
            if hl == "ongeschikt" and (not _contains_any(fr, ["Parkstad", "Maastricht", "Sittard"])):
                return False
            return True
        rows = [r for r in rows if keep_r2(r)]

        # 3) Enigszins geschikt + Focusregio ≠ Sittard/Parkstad/Maastricht -> vervang Gewenste placeholder
        for r in rows:
            hl = "" if r.get("Heel Limburg") is None else str(r.get("Heel Limburg")).strip()
            fr = "" if r.get("Focusregio") is None else str(r.get("Focusregio")).strip()
            if hl == "enigszins geschikt" and fr not in ("Sittard", "Parkstad", "Maastricht"):
                r["Gewenste placeholder"] = r.get("Placeholder bij enigszins geschikt")

        focus_targets = ("Parkstad", "Maastricht", "Sittard")
        def focus_bonus(fr):
            frs = "" if fr is None else str(fr)
            return 10 if any(t in frs for t in focus_targets) else 0

    else:  # "NM"
        # 1) Schrap: Plaatsing begint met 'ND-' of 'NM-'
        def keep_r(r):
            p = "" if r.get("Plaatsing") is None else str(r.get("Plaatsing"))
            return not (p.startswith("ND-") or p.startswith("NM-"))
        rows = [r for r in rows if keep_r(r)]

        # 2) Schrap: Heel Limburg=ongeschikt AND Focusregio bevat géén Noord/Midden
        def keep_r2(r):
            hl = "" if r.get("Heel Limburg") is None else str(r.get("Heel Limburg")).strip()
            fr = r.get("Focusregio")
            if hl == "ongeschikt" and (not _contains_any(fr, ["Noord", "Midden"])):
                return False
            return True
        rows = [r for r in rows if keep_r2(r)]

        # 3) Enigszins geschikt + Focusregio ≠ Noord/Midden -> vervang Gewenste placeholder
        for r in rows:
            hl = "" if r.get("Heel Limburg") is None else str(r.get("Heel Limburg")).strip()
            fr = "" if r.get("Focusregio") is None else str(r.get("Focusregio")).strip()
            if hl == "enigszins geschikt" and fr not in ("Noord", "Midden"):
                r["Gewenste placeholder"] = r.get("Placeholder bij enigszins geschikt")

        focus_targets = ("Noord", "Midden")
        def focus_bonus(fr):
            frs = "" if fr is None else str(fr)
            return 10 if any(t in frs for t in focus_targets) else 0

    # 4) Wis kolommen Prioscore / Gekozen template / Gekozen placeholder / Plaatsing
    for r in rows:
        r["Prioscore"] = None
        r["Gekozen template"] = None
        r["Gekozen placeholder"] = None
        r["Plaatsing"] = None

    # 5) Bereken nieuwe Prioscore
    def hl_score(hl):
        s = "" if hl is None else str(hl).strip()
        if s == "moet mee":
            return 5
        if s == "geschikt":
            return 3
        if s == "enigszins geschikt":
            return 1
        return 0

    for r in rows:
        score = 0
        score += hl_score(r.get("Heel Limburg"))

        pub = "" if r.get("Publicatiedwang") is None else str(r.get("Publicatiedwang")).strip()
        if pub == "Nee":
            score -= 1

        top8 = "" if r.get("Top 8") is None else str(r.get("Top 8")).strip()
        if top8 == "Ja":
            score += 3

        score += focus_bonus(r.get("Focusregio"))

        r["Prioscore"] = score

    # 6) Sorteren op Prioscore (hoogste bovenaan)
    rows.sort(key=lambda r: (r.get("Prioscore") if r.get("Prioscore") is not None else -10**9), reverse=True)

    _rewrite_sheet(ws, headers, rows, h2i)

apply_mappingregels(zu_ws, "ZU")
apply_mappingregels(nm_ws, "NM")

# 6) Kopieer ZU-UITW -> ZU-U1..ZU-U5 én ZU-UNUSED. Verwijder daarna tabblad ZU-UITW
for i in range(1, 6):
    _ensure_fresh_copy("ZU-UITW", f"ZU-U{i}")
_ensure_fresh_copy("ZU-UITW", "ZU-UNUSED")
if "ZU-UITW" in plan_wb.sheetnames:
    plan_wb.remove(plan_wb["ZU-UITW"])

# 7) Kopieer NM-UITW -> NM-U1..NM-U5 én NM-UNUSED. Verwijder daarna tabblad NM-UITW
for i in range(1, 6):
    _ensure_fresh_copy("NM-UITW", f"NM-U{i}")
_ensure_fresh_copy("NM-UITW", "NM-UNUSED")
if "NM-UITW" in plan_wb.sheetnames:
    plan_wb.remove(plan_wb["NM-UITW"])

# 8) Volg [MAPPINGREGELS EXTRA]
def _contains_any(text, needles):
    if text is None:
        return False
    s = str(text)
    return any(n in s for n in needles)

def _get_laatste_run_map():
    hdr = [pos_ws.cell(1,c).value for c in range(1, pos_ws.max_column+1)]
    def _col(name):
        for i,h in enumerate(hdr, start=1):
            if h is None:
                continue
            if str(h).strip().lower() == name.lower():
                return i
        return None
    c_pos = _col("Positie")
    c_lr  = _col("Laatste Run")
    if not c_pos or not c_lr:
        return {}
    out={}
    for r in range(2, pos_ws.max_row+1):
        p = pos_ws.cell(r, c_pos).value
        lr = pos_ws.cell(r, c_lr).value
        if p is None:
            continue
        p = str(p).strip()
        if not p:
            continue
        out[p] = ("" if lr is None else str(lr).strip().lower())
    return out

def _as_number(x):
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def _find_cols(ws, *names):
    hdr = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
    def col(name):
        for i,h in enumerate(hdr, start=1):
            if h is None:
                continue
            if str(h).strip().lower() == name.lower():
                return i
        return None
    return {n: col(n) for n in names}

def _apply_last_run_rules(ws, lr_value):
    cols = _find_cols(ws, "Heel Limburg", "Prioscore", "Gewenste placeholder", "Tweede keus placeholder")
    c_hl = cols["Heel Limburg"]; c_pr = cols["Prioscore"]
    if not c_hl or not c_pr:
        return
    for r in range(2, ws.max_row+1):
        hl = ws.cell(r, c_hl).value
        if ("" if hl is None else str(hl).strip()) != "moet mee":
            continue
        cur = _as_number(ws.cell(r, c_pr).value)
        if lr_value == "voorlaatste":
            ws.cell(r, c_pr).value = cur + 20
        elif lr_value == "laatste":
            ws.cell(r, c_pr).value = cur + 30
            c_gw = cols["Gewenste placeholder"]; c_2 = cols["Tweede keus placeholder"]
            if c_gw and c_2:
                gw = ws.cell(r, c_gw).value
                sec = ws.cell(r, c_2).value
                gw_s = "" if gw is None else str(gw).strip()
                sec_s = "" if sec is None else str(sec).strip()
                if sec_s:
                    if not gw_s:
                        ws.cell(r, c_gw).value = sec_s
                    elif sec_s not in gw_s:
                        ws.cell(r, c_gw).value = gw_s.rstrip("; ") + "; " + sec_s

def _apply_classificatie_and_focus_bonus(ws, focus_must_not, focus_bonus_needles):
    cols = _find_cols(ws, "Heel Limburg", "Focusregio", "Classificatie", "Prioscore")
    c_hl=cols["Heel Limburg"]; c_fr=cols["Focusregio"]; c_cl=cols["Classificatie"]; c_pr=cols["Prioscore"]
    if not c_hl or not c_fr:
        return
    for r in range(2, ws.max_row+1):
        hl = "" if ws.cell(r, c_hl).value is None else str(ws.cell(r, c_hl).value).strip()
        fr = ws.cell(r, c_fr).value
        if hl in ("enigszins geschikt", "geschikt") and (not _contains_any(fr, focus_must_not)):
            if c_cl:
                cl = "" if ws.cell(r, c_cl).value is None else str(ws.cell(r, c_cl).value).strip()
                # multi-value: verwijder A-keus uit de string (downgrade)
                cl_set_raw = {p.strip() for p in re.split(r"[;,+/]|\ben\b", str(cl), flags=re.IGNORECASE) if p.strip()}
                # vergelijk case-insensitive
                cl_set = {normalize(p).lower() for p in cl_set_raw if normalize(p)}
                if "a-keus" in cl_set:
                    # verwijder A-keus
                    cl_set = {p for p in cl_set if p != "a-keus"}
                    if not cl_set:
                        ws.cell(r, c_cl).value = "B-keus"
                    else:
                        # schrijf terug zonder trailing ; 
                        ws.cell(r, c_cl).value = "; ".join(_pretty_class_token(t) for t in sorted(cl_set))
        if focus_bonus_needles and c_pr and _contains_any(fr, focus_bonus_needles):
            ws.cell(r, c_pr).value = _as_number(ws.cell(r, c_pr).value) + 5

def apply_mappingregels_extra():
    lr_map = _get_laatste_run_map()
    for sh, lr in lr_map.items():
        if sh in plan_wb.sheetnames and lr in ("voorlaatste", "laatste"):
            _apply_last_run_rules(plan_wb[sh], lr)

    if "ZU-U1" in plan_wb.sheetnames:
        _apply_classificatie_and_focus_bonus(plan_wb["ZU-U1"],
                                             focus_must_not=["Parkstad","Maastricht","Sittard","Limburg-breed"],
                                             focus_bonus_needles=["Parkstad","Maastricht","Sittard"])
    if "NM-U1" in plan_wb.sheetnames:
        _apply_classificatie_and_focus_bonus(plan_wb["NM-U1"],
                                             focus_must_not=["Noord","Midden","Limburg-breed"],
                                             focus_bonus_needles=["Noord","Midden"])

apply_mappingregels_extra()

# 9) Runs NM-U1..NM-U5
run_runs(["NM-U1","NM-U2","NM-U3","NM-U4","NM-U5"])

# 10) Runs ZU-U1..ZU-U5
run_runs(["ZU-U1","ZU-U2","ZU-U3","ZU-U4","ZU-U5"])

# 11) Opschonen Totale verhalenlijst
def clean_totale_verhalenlijst():
    sh_name = "Totale verhalenlijst"
    if sh_name not in plan_wb.sheetnames:
        return
    ws = plan_wb[sh_name]
    headers = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
    def _col_idx(name):
        for i,h in enumerate(headers, start=1):
            if h is None:
                continue
            if str(h).strip().lower() == name.strip().lower():
                return i
        return None
    col_pl = _col_idx("Plaatsing")
    col_nm = _col_idx("Naam productie")
    if not col_pl or not col_nm:
        return
    rows=[]
    for r in range(2, ws.max_row+1):
        vals=[ws.cell(r,c).value for c in range(1, ws.max_column+1)]
        plaatsing = vals[col_pl-1]
        if plaatsing is None or str(plaatsing).strip()=="":
            continue
        rows.append(vals)
    rows.sort(key=lambda vals: ("" if vals[col_nm-1] is None else str(vals[col_nm-1]).strip().lower()))
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row-1)
    for vals in rows:
        ws.append(vals)

clean_totale_verhalenlijst()

# 12) Kopieer laatste 'Beschrijving' uit Logfile naar AD1 op alle tabbladen waarvoor een run is uitgevoerd
def copy_logfile_beschrijving_to_runs():
    """Schrijf per uitgevoerd run-tabblad de bijbehorende (laatste) Logfile.Beschrijving naar AD1."""
    if "Logfile" not in plan_wb.sheetnames:
        return
    log_ws = plan_wb["Logfile"]
    if log_ws.max_row < 2:
        return

    # Zoek kolom 'Beschrijving' (fallback: kolom 2)
    hdr = [log_ws.cell(1,c).value for c in range(1, log_ws.max_column+1)]
    desc_col = None
    for i,h in enumerate(hdr, start=1):
        if h is None:
            continue
        if str(h).strip().lower() == "beschrijving":
            desc_col = i
            break
    if desc_col is None:
        desc_col = 2

    latest = {}  # run_name -> beschrijving
    for r in range(2, log_ws.max_row+1):
        v = log_ws.cell(r, desc_col).value
        if v is None or str(v).strip()=="":
            continue
        s = str(v).strip()
        if not s.startswith("Run "):
            continue
        # verwacht: 'Run <TAB> ...' of 'Run <TAB>: ...'
        after = s[4:]
        run_name = after.split(":")[0].split(" ")[0].strip()
        if run_name:
            latest[run_name] = s

    for run_name, desc in latest.items():
        if run_name in plan_wb.sheetnames:
            plan_wb[run_name]["AD1"].value = desc

copy_logfile_beschrijving_to_runs()

# 14) EINDBEWERKINGEN
def apply_eindbewerkingen_placeholder_concessie():
    """IF 'Gekozen placeholder' niet voorkomt in string 'Gewenste placeholder' THEN Placeholder-concessie=Ja,
    op alle tabbladen waarvoor een Run is gedaan en op tabblad 'Totale verhalenlijst'."""
    if "Logfile" not in plan_wb.sheetnames:
        return
    log_ws = plan_wb["Logfile"]
    if log_ws.max_row < 2:
        return

    # Vind kolom 'Beschrijving' in Logfile
    hdr = [log_ws.cell(1,c).value for c in range(1, log_ws.max_column+1)]
    desc_col = None
    for i,h in enumerate(hdr, start=1):
        if h is None:
            continue
        if str(h).strip().lower() == "beschrijving":
            desc_col = i
            break
    if desc_col is None:
        desc_col = 2

    # Verzamel run-tabbladen (alle regels die met 'Run ' beginnen)
    run_tabs = []
    seen = set()
    for r in range(2, log_ws.max_row+1):
        v = log_ws.cell(r, desc_col).value
        if v is None:
            continue
        s = str(v).strip()
        if not s.startswith("Run "):
            continue
        after = s[4:]
        run_name = after.split(":")[0].split(" ")[0].strip()
        if run_name and run_name in plan_wb.sheetnames and run_name not in seen:
            seen.add(run_name)
            run_tabs.append(run_name)

    # Voeg altijd ook "Totale verhalenlijst" toe
    if "Totale verhalenlijst" in plan_wb.sheetnames and "Totale verhalenlijst" not in seen:
        run_tabs.append("Totale verhalenlijst")
        seen.add("Totale verhalenlijst")

    # Pas de regel toe per run-tabblad
    for sh in run_tabs:
        ws = plan_wb[sh]
        if ws.max_row < 2:
            continue

        headers = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
        def find_col(name):
            for i,h in enumerate(headers, start=1):
                if h is None:
                    continue
                if str(h).strip().lower() == name:
                    return i
            return None

        col_chosen = find_col("gekozen placeholder")
        col_desired = find_col("gewenste placeholder")
        col_conc = find_col("placeholder-concessie")

        if col_chosen is None or col_desired is None:
            continue

        if col_conc is None:
            # Voeg kolom toe als deze ontbreekt
            col_conc = ws.max_column + 1
            ws.cell(1, col_conc).value = "Placeholder-concessie"
            headers.append("Placeholder-concessie")

        for r in range(2, ws.max_row+1):
            chosen = ws.cell(r, col_chosen).value
            desired = ws.cell(r, col_desired).value
            if chosen is None or str(chosen).strip()=="":
                continue
            if desired is None:
                continue
            chosen_s = str(chosen).strip()
            desired_s = str(desired)
            if chosen_s not in desired_s:
                # Alleen als nog niet 'Ja' is gezet (niet overschrijven met iets anders)
                cur = ws.cell(r, col_conc).value
                if cur is None or str(cur).strip()=="":
                    ws.cell(r, col_conc).value = "Ja"
                else:
                    # laat bestaande waarde staan, maar zorg dat 'Ja' niet verloren gaat
                    if str(cur).strip().lower() != "ja":
                        ws.cell(r, col_conc).value = "Ja"

apply_eindbewerkingen_placeholder_concessie()

# -----------------------------
# EINDBEWERKINGEN v58: Naam van positie (uit [POSITIELIJST]) naar kolom AE op run-tabbladen
# -----------------------------
def apply_eindbewerkingen_naam_van_positie():
    """Vul op alle tabbladen waar een Run is gedaan in kolom AE de waarde 'Naam van positie' in,
    af te lezen uit [POSITIELIJST] (pos_ws)."""
    if "Logfile" not in plan_wb.sheetnames:
        return
    log_ws = plan_wb["Logfile"]
    if log_ws.max_row < 2:
        return

    # Vind kolom 'Beschrijving' in Logfile
    hdr = [log_ws.cell(1,c).value for c in range(1, log_ws.max_column+1)]
    desc_col = None
    for i,h in enumerate(hdr, start=1):
        if h is None:
            continue
        if str(h).strip().lower() == "beschrijving":
            desc_col = i
            break
    if desc_col is None:
        desc_col = 2

    # Verzamel run-tabbladen (alle regels die met 'Run ' beginnen)
    run_tabs = []
    seen = set()
    for r in range(2, log_ws.max_row+1):
        v = log_ws.cell(r, desc_col).value
        if v is None:
            continue
        s = str(v).strip()
        if not s.startswith("Run "):
            continue
        after = s[4:]
        run_name = after.split(":")[0].split(" ")[0].strip()
        if run_name and run_name in plan_wb.sheetnames and run_name not in seen:
            seen.add(run_name)
            run_tabs.append(run_name)

    # Bouw mapping Positie(tabblad) -> Naam van positie uit [POSITIELIJST]
    try:
        pos_headers = [pos_ws.cell(1,c).value for c in range(1, pos_ws.max_column+1)]
    except Exception:
        return

    pos_col = None
    naam_col = None
    for i,h in enumerate(pos_headers, start=1):
        if h is None:
            continue
        hn = str(h).strip().lower()
        if hn == "positie":
            pos_col = i
        if hn == "naam van positie":
            naam_col = i

    if pos_col is None or naam_col is None:
        return

    pos_map = {}
    for r in range(2, pos_ws.max_row+1):
        p = pos_ws.cell(r, pos_col).value
        n = pos_ws.cell(r, naam_col).value
        if p is None:
            continue
        p_s = str(p).strip()
        if p_s == "":
            continue
        pos_map[p_s] = "" if n is None else str(n)

    # Schrijf naar kolom AE (31)
    target_col_idx = 31  # AE
    for sh in run_tabs:
        ws = plan_wb[sh]
        if ws.max_row < 1:
            continue
        ws.cell(1, target_col_idx).value = "Naam van positie"
        naam = pos_map.get(sh, "")
        for r in range(2, ws.max_row+1):
            ws.cell(r, target_col_idx).value = naam

apply_eindbewerkingen_naam_van_positie()

# -----------------------------
# EINDBEWERKINGEN v55: Planning print tabblad
# -----------------------------
def create_planning_print():
    src_name = "Totale verhalenlijst"
    dst_name = "Planning print"
    if src_name not in plan_wb.sheetnames:
        return
    # Verwijder bestaande Planning print indien aanwezig
    if dst_name in plan_wb.sheetnames:
        plan_wb.remove(plan_wb[dst_name])
    # Kopieer sheet
    ws_new = plan_wb.copy_worksheet(plan_wb[src_name])
    ws_new.title = dst_name

    # Positioneer als eerste tabblad (meest links)
    try:
        sheets = plan_wb._sheets
        # verplaats laatste (net gekopieerde) naar index 0
        sheets.insert(0, sheets.pop(sheets.index(ws_new)))
    except Exception:
        pass

    ws = plan_wb[dst_name]
    if ws.max_row < 1:
        return

    # Verwijder kolommen
    cols_to_remove = [
        "Note", "Voorkeurspositie", "Printbeeld", "Print beeld", "Graphic", "Karakters", "Artikelsoort",
        "Gewenste placeholder", "Tweede keus placeholder", "Derde keus placeholder", "Vierde keus placeholder",
        "Placeholder bij enigszins geschikt", "Prioscore"
    ]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    # bepaal indices (1-based) die verwijderd moeten worden
    remove_idxs = []
    for i, h in enumerate(headers, start=1):
        if h is None:
            continue
        if str(h).strip() in cols_to_remove:
            remove_idxs.append(i)
    for col_idx in sorted(remove_idxs, reverse=True):
        ws.delete_cols(col_idx, 1)

    # Sorteer rijen op Plaatsing volgens vaste volgorde
    sort_order = [
        "NM-NO", "NM-MI", "ZU-MH", "ZU-SG", "ZU-PS",
        "ND-01", "ND-02", "ND-03",
        "NM-U1", "NM-U2", "NM-U3", "NM-U4", "NM-U5",
        "ZU-U1", "ZU-U2", "ZU-U3", "ZU-U4", "ZU-U5"
    ]
    order_map = {v: i for i, v in enumerate(sort_order)}
    # herlees headers na kolomverwijdering
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    plaatsing_col = None
    for i, h in enumerate(headers, start=1):
        if h is None:
            continue
        if str(h).strip().lower() == "plaatsing":
            plaatsing_col = i
            break
    if plaatsing_col is None or ws.max_row < 3:
        return

    data_rows = []
    for r in range(2, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        pl = ws.cell(r, plaatsing_col).value
        pl_s = "" if pl is None else str(pl).strip()
        key = order_map.get(pl_s, 10**9)
        data_rows.append((key, row_vals))

    data_rows.sort(key=lambda x: x[0])

    # Wis bestaande data (behoud header)
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # Schrijf gesorteerde rijen terug
    for _, row_vals in data_rows:
        ws.append(row_vals)

create_planning_print()

print(f"Klaar. Succesvolle matches: {success}")

# -----------------------------
# Save output + download
# -----------------------------
OUT_PATH = "Krantenplanning.xlsx"
plan_wb.save(OUT_PATH)

# -*- coding: utf-8 -*-
# Adapted from legacy build_zy_patch.py for the 2026-8-4 中运通达 quote.
# Maps 11 package sheets -> 33 b2b-air channels. Preserves sur block, only
# updates tiers. Dry-run by default; pass --write to apply into rates.json.
import openpyxl, json, re, os, argparse

HERE = os.path.dirname(os.path.abspath(__file__))
NEWXLSX = os.path.join(HERE, "中运通达", "latest.xlsx")
RATES = os.path.join(HERE, "..", "..", "rates.json")
OUT = os.path.join(HERE, "patch_zy.json")

def num(x):
    try:
        if x is None: return None
        s = str(x).replace(",", "").strip()
        if s in ("", "***", "/", "-"): return None
        return float(s)
    except: return None

def wb_open(path):
    return openpyxl.load_workbook(path, read_only=True, data_only=True)

def detect_groups(header_cells):
    groups = []; run = []
    for i, c in enumerate(header_cells):
        if i < 2:
            if run: groups.append(run); run=[]
            continue
        s = "" if c is None else str(c)
        m = re.search(r'(\d+)\s*KG', s, re.IGNORECASE)
        if m: run.append((i, int(m.group(1))))
        else:
            if run: groups.append(run); run=[]
    if run: groups.append(run)
    return groups

def is_note(s):
    if not s: return False
    if any(k in s for k in ("说明","：","返回目录")): return True
    if re.match(r'^[一二三四五六七八九十]+、', s): return True
    if re.match(r'^\d+[、.、]', s): return True
    return False

def expand_keys(label, live_keys):
    s = label.strip()
    if s in live_keys: return [s]
    for k in live_keys:
        if s.startswith(k+'（') or s.startswith(k+'('):
            return [k]
    parts = []
    if '(' in s:
        parts = [p for p in re.split(r'[\s\n]+', s) if p.strip()]
    else:
        parts = [p for p in re.split(r'[、，,\n]', s) if p.strip()]
    inlive = [p for p in parts if p in live_keys]
    return inlive if inlive else parts

def best_match(live_tiers, choices):
    live_dict = {t["min"]: t["rate"] for t in live_tiers}
    best = None; bestscore = None
    for ch in choices:
        sc = 0.0
        for mn, rt in ch:
            mn = float(mn); rt = float(rt)
            sc += abs(live_dict.get(mn, rt) - rt) if mn in live_dict else abs(rt)*0.5
        sc += 5.0 * abs(len(ch) - len(live_tiers))
        if bestscore is None or sc < bestscore:
            bestscore = sc; best = ch
    return best

def tiers_from_choice(choice):
    return [{"min": float(mn), "rate": float(rt)} for mn, rt in choice]

def rows_of(wb, sheet):
    return [list(r) for r in wb[sheet].iter_rows(values_only=True)]

def parse_vertical(wb, sheet, blocks):
    rows = rows_of(wb, sheet)
    out = {}
    title_rows = []
    for i, r in enumerate(rows):
        c1 = r[1] if len(r) > 1 else ""
        c1 = "" if c1 is None else str(c1)
        exact = [(tstr, cid, collapsed, lcol) for (tstr, cid, collapsed, lcol) in blocks if c1.strip() == tstr]
        sub = [(tstr, cid, collapsed, lcol) for (tstr, cid, collapsed, lcol) in blocks if tstr and tstr in c1]
        cand = exact or sub
        if cand:
            title_rows.append((i, cand[0][1], cand[0][2], cand[0][3]))
    for idx, (ti, cid, collapsed, lcol) in enumerate(title_rows):
        hdr = None
        for j in list(range(ti, min(ti+4, len(rows)))) + list(range(ti-1, max(-1, ti-4), -1)):
            if any(re.search(r'\d+\s*KG', str(c), re.IGNORECASE) for c in rows[j]):
                hdr = j; break
        if hdr is None:
            continue
        groups = detect_groups(rows[hdr])
        if not groups:
            continue
        grp = groups[0]
        price_cols = [x[0] for x in grp]
        start = min(ti, hdr)
        end = title_rows[idx+1][0] if idx+1 < len(title_rows) else len(rows)
        choices = []
        for j in range(start, end):
            rr = [("" if c is None else str(c)) for c in rows[j]]
            label = rr[lcol].strip() if lcol < len(rr) else ""
            if not label or "返回目录" in label or is_note(label):
                continue
            vals = []
            for pc in price_cols:
                n = num(rr[pc] if pc < len(rr) else "")
                if n is not None:
                    vals.append((grp[price_cols.index(pc)][1], n))
            if vals:
                choices.append((label, vals))
        out[cid] = choices
    return out

def parse_ca_haipai(wb):
    rows = rows_of(wb, "加拿大包税海派")
    header_rows = [2, 8, 14]
    chan_order = ["zy_ca_hp_ts_ups","zy_ca_hp_ts_ka","zy_ca_hp_dt_ups","zy_ca_hp_dt_ka","zy_ca_hp_hy_ups","zy_ca_hp_hy_ka"]
    out = {}; ci = 0
    for hr in header_rows:
        groups = detect_groups(rows[hr])
        nxt = header_rows[header_rows.index(hr)+1] if header_rows.index(hr)+1 < len(header_rows) else len(rows)
        for g in groups:
            cid = chan_order[ci]; ci += 1
            price_cols = [x[0] for x in g]
            lcol = g[0][0] - 1
            choices = []
            for j in range(hr+1, nxt):
                rr = [("" if c is None else str(c)) for c in rows[j]]
                label = rr[lcol].strip() if lcol < len(rr) else ""
                if not label or "返回目录" in label or is_note(label): continue
                vals = []
                for pc in price_cols:
                    n = num(rr[pc] if pc < len(rr) else "")
                    if n is not None:
                        vals.append((g[price_cols.index(pc)][1], n))
                if vals: choices.append((label, vals))
            out[cid] = choices
    return out

def parse_au(wb):
    rows = rows_of(wb, "澳洲包税海派")
    groups = detect_groups(rows[3])
    cids = ["zy_au_sea","zy_au_air","zy_au_air_m"]
    out = {}
    for gi, g in enumerate(groups):
        cid = cids[gi]
        price_cols = [x[0] for x in g]
        lcol = 1
        choices = []
        for j in range(4, 8):
            rr = [("" if c is None else str(c)) for c in rows[j]]
            label = rr[lcol].strip() if lcol < len(rr) else ""
            if not label or "返回目录" in label or is_note(label): continue
            vals = []
            for pc in price_cols:
                n = num(rr[pc] if pc < len(rr) else "")
                if n is not None:
                    vals.append((g[price_cols.index(pc)][1], n))
            if vals: choices.append((label, vals))
        out[cid] = choices
    return out

def parse_us_haika(wb):
    rows = rows_of(wb, "美国包税海卡")
    hdr = None
    for i, r in enumerate(rows):
        rc = [("" if c is None else str(c)) for c in r]
        if any("仓库代码" in x for x in rc) and any(re.search(r'\d+\s*KG', x, re.I) for x in rc):
            hdr = i; break
    if hdr is None:
        return {}
    groups = detect_groups(rows[hdr])
    cids = ["zy_us_hk_clx", "zy_us_hk_max", "zy_us_hk_cosco"]
    out = {}
    for gi, g in enumerate(groups):
        if gi >= len(cids):
            break
        cid = cids[gi]
        price_cols = [x[0] for x in g]
        lcol = g[0][0] - 1
        choices = []
        for j in range(hdr + 1, len(rows)):
            rr = [("" if c is None else str(c)) for c in rows[j]]
            label = rr[lcol].strip() if lcol < len(rr) else ""
            if not label or "返回目录" in label or is_note(label): continue
            if "单询" in label or "价格表" in label:
                continue
            vals = []
            for k, pc in enumerate(price_cols):
                n = num(rr[pc] if pc < len(rr) else "")
                if n is not None:
                    vals.append((g[k][1], n))
            if vals: choices.append((label, vals))
        out[cid] = choices
    return out

def parse_uk(wb):
    rows = rows_of(wb, "英国VAT递延")
    groups = detect_groups(rows[2])
    grp = groups[0]; price_cols = [x[0] for x in grp]
    out = {}
    for j in range(3, 9):
        rr = [("" if c is None else str(c)) for c in rows[j]]
        label = rr[1].strip() if len(rr) > 1 else ""
        if not label: continue
        vals = []
        for pc in price_cols:
            n = num(rr[pc] if pc < len(rr) else "")
            if n is not None:
                vals.append((grp[price_cols.index(pc)][1], n))
        if vals: out[label] = vals
    return out

def parse_jp(wb):
    rows = rows_of(wb, "日本自税空海派")
    groups = detect_groups(rows[2])
    grp = groups[0]; price_cols = [x[0] for x in grp]
    out = {}
    for j in range(3, 6):
        rr = [("" if c is None else str(c)) for c in rows[j]]
        label = rr[1].strip() if len(rr) > 1 else ""
        sub = rr[2].strip() if len(rr) > 2 else ""
        if not label: continue
        vals = []
        for pc in price_cols:
            n = num(rr[pc] if pc < len(rr) else "")
            if n is not None:
                vals.append((grp[price_cols.index(pc)][1], n))
        if vals: out[sub] = vals
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true", help="apply patch into rates.json")
    args = ap.parse_args()

    wb = wb_open(NEWXLSX)
    rates = json.load(open(RATES, encoding="utf-8"))
    live = {c["id"]: c for c in rates["channels"] if c.get("carrier") == "中运通达"}

    vertical_config = {
        "欧洲包税": [("欧洲包税-空派快线(普货)","zy_eu_baoshui_0",False,1),
                     ("欧洲包税-空派慢线(普货)","zy_eu_baoshui_1",False,1),
                     ("欧洲包税-卡航","zy_eu_baoshui_2",False,1),
                     ("欧洲包税-卡航卡派","zy_eu_baoshui_3",False,1),
                     ("欧洲包税-海运","zy_eu_baoshui_4",False,1),
                     ("欧洲包税-海运卡派","zy_eu_baoshui_5",False,1)],
        "美国包税空派": [("美国包税-空派(普货)","zy_us_air",False,1)],
        "美国包税海派": [("美森正班","zy_us_hp_clx",False,1),
                         ("美森加班","zy_us_hp_max",False,1),
                         ("盐田普船","zy_us_hp_cosco",False,1)],
        "加拿大包税空派": [("加拿大包税-空派(普货)","zy_ca_baoshui_0",True,1)],
        "欧洲VAT递延": [("欧洲VAT递延-空运","zy_eu_vat_air",False,1),
                        ("欧洲VAT递延-卡航","zy_eu_vat_kahang",False,1),
                        ("欧洲VAT递延-海运","zy_eu_vat_sea",False,1)],
    }

    raw_choices = {}
    for sheet, blocks in vertical_config.items():
        raw_choices.update(parse_vertical(wb, sheet, blocks))
    raw_choices.update(parse_ca_haipai(wb))
    raw_choices.update(parse_au(wb))
    raw_choices.update(parse_us_haika(wb))

    uk_choices = parse_uk(wb)
    def norm(s): return re.sub(r'[\n\s（）()（）]','',s)
    uk_map = {
        "英国VAT递延-空运快线":"zy_uk_air_kx",
        "英国VAT递延-空运慢线":"zy_uk_air_mx",
        "英国VAT递延-卡航":"zy_uk_kahang",
        "英国VAT递延-海运":"zy_uk_sea",
        "英国VAT递延-海卡\nBHX4/BHX8/LBA4/LBA8":"zy_uk_hk_bhx",
        "英国VAT递延-海卡\nLPL2/EMA3/MAN4/BHX7/LBA2/CWL1/DSA7/LTN7/XBH7/XDS1/BRS2/EMA2/LTN4/XBH9":"zy_uk_hk_lpl",
    }
    uk_cid = {norm(k): v for k, v in uk_map.items()}
    for title, vals in uk_choices.items():
        cid = uk_cid.get(norm(title))
        if cid: raw_choices[cid] = [("英国", vals)]

    jp_choices = parse_jp(wb)
    for sub, vals in jp_choices.items():
        if "海运" in sub:
            raw_choices["zy_jp_sea"] = [("日本", vals)]

    wb.close()

    patch = {}
    gaps = []
    for cid, ch in live.items():
        if ch.get("cat") != "b2b-air": continue
        if cid not in raw_choices or not raw_choices[cid]:
            gaps.append((cid, "no source sheet/block")); continue
        choices = raw_choices[cid]
        live_keys = list(ch["countries"].keys())
        parsed_map = {}
        for label, vals in choices:
            for k in expand_keys(label, live_keys):
                if k in live_keys and k not in parsed_map:
                    parsed_map[k] = vals
        new_countries = {}
        for k in live_keys:
            extra = {kk: vv for kk, vv in ch["countries"][k].items() if kk != "tiers"}
            if k in parsed_map:
                new_countries[k] = {**extra, "tiers": tiers_from_choice(parsed_map[k])}
            else:
                bm = best_match(ch["countries"][k]["tiers"], [v for (_, v) in choices])
                if bm is None:
                    gaps.append((cid, "best_match empty")); continue
                new_countries[k] = {**extra, "tiers": tiers_from_choice(bm)}
        patch[cid] = {"countries": new_countries}

    print("=== SELF-CHECK ===")
    for cid in sorted(patch):
        p = patch[cid]
        sk = list(p["countries"].keys())[:1]
        sample = {k: p["countries"][k]["tiers"][:2] for k in sk}
        print(f"{cid:20s} #c={len(p['countries']):2d} sample={sample}")
    print("\n=== GAPS ===")
    for cid in sorted(live):
        if live[cid].get("cat")!="b2b-air": continue
        if cid not in patch:
            print("  GAP:", cid)
    for g in gaps:
        print("  gap-detail:", g)

    json.dump(patch, open(OUT,"w",encoding="utf-8"), ensure_ascii=False, indent=1)
    print("\nWROTE", OUT, "channels:", len(patch))

    if not args.write:
        print("\n[DRY-RUN] not applied. Re-run with --write to update rates.json")
        return

    applied = 0
    for cid, ch in live.items():
        if cid not in patch: continue
        ch["countries"] = patch[cid]["countries"]
        ch["dest"] = list(patch[cid]["countries"].keys())
        applied += 1
    json.dump(rates, open(RATES,"w",encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"\n[WRITE] applied {applied} channels into rates.json")

if __name__ == "__main__":
    main()

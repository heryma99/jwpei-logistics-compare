# -*- coding: utf-8 -*-
# Bake 亚丰(璞景/德翼) 18 b2b-air channels from 亚丰/latest.xlsx (2026-8-4 璞景VIP).
# Preserves existing channel structure (countries set + sur); only updates tiers.
# Dry-run by default; --write to apply into rates.json.
import openpyxl, json, re, os, argparse

HERE = os.path.dirname(os.path.abspath(__file__))
NEWXLSX = os.path.join(HERE, "亚丰", "latest.xlsx")
RATES = os.path.join(HERE, "..", "..", "rates.json")
OUT = os.path.join(HERE, "patch_yf.json")

def num(x):
    try:
        if x is None: return None
        s = str(x).replace("，", "").replace(",", "").strip()
        if s in ("", "***", "/", "-", "—"): return None
        return float(s)
    except: return None

def wb_open(path):
    return openpyxl.load_workbook(path, read_only=True, data_only=True)

def expand_keys(label, live_keys):
    s = label.strip()
    if s in live_keys: return [s]
    for k in live_keys:
        if s.startswith(k + "（") or s.startswith(k + "("):
            return [k]
    if "（" in s or "(" in s:
        head = re.split(r"[（(]", s)[0].strip()
        if head in live_keys: return [head]
    parts = [p.strip() for p in re.split(r"[、，,\n]", s) if p.strip()]
    inlive = [p for p in parts if p in live_keys]
    return inlive if inlive else parts

def tiers_from_choice(choice):
    return [{"min": float(mn), "rate": float(rt)} for mn, rt in choice]

def rows_of(wb, sheet):
    return [list(r) for r in wb[sheet].iter_rows(values_only=True)]

def find_hdr(rows, sheet, markers):
    for i, r in enumerate(rows):
        j = "".join("" if c is None else str(c) for c in r)
        if all(m in j for m in markers):
            return i
    return None

# ---- US air (美国空派快线/经济线): 3 region rows -> 美西/美中/美东 ----
def parse_yf_us_regions(sheet, tier_mins, wb):
    rows = rows_of(wb, sheet)
    hdr = find_hdr(rows, sheet, ["KG"])
    for i, r in enumerate(rows):
        j = "".join("" if c is None else str(c) for c in r)
        if re.search(r"\d+\s*[-－]\s*\d+\s*kg", j, re.I) or "10-20kg" in j or "12≤W" in j:
            hdr = i; break
    out = {}
    for i in range(hdr + 1, len(rows)):
        rc = ["" if c is None else str(c) for c in rows[i]]
        label = rc[0]
        if not label or "附加费" in label or "备注" in label:
            if num(rc[1]) is None and num(rc[2]) is None:
                continue
        if "美" not in label and "USW" not in label.upper() and "USE" not in label.upper() and "USM" not in label.upper():
            continue
        vals = [num(x) for x in rc[1:] if num(x) is not None]
        if len(vals) < len(tier_mins):
            continue
        rates = vals[:len(tier_mins)]
        region = None
        if "西" in label or "USW" in label.upper(): region = "美西"
        elif "中" in label or "USM" in label.upper(): region = "美中"
        elif "东" in label or "USE" in label.upper(): region = "美东"
        if region:
            out[region] = list(zip(tier_mins, rates))
        if len(out) >= 3:
            break
    return out

# ---- EU zone sheets (欧洲空运包税/海运/快铁): 分区+国家 -> 22 countries ----
def parse_yf_eu_zones(sheet, tier_mins, wb, stop):
    rows = rows_of(wb, sheet)
    hdr = None
    for i, r in enumerate(rows):
        rc = ["" if c is None else str(c) for c in r]
        if any("21KG" in c or "21KG＋" in c or "21" in c and "KG" in c.upper() for c in rc) and any("分区" in c for c in rc):
            hdr = i; break
    if hdr is None:
        return {}
    rc = ["" if c is None else str(c) for c in rows[hdr]]
    part_idx = next((i for i, c in enumerate(rc) if c == "分区"), None)
    if part_idx is None:
        return {}
    tier_cols = [i for i in range(part_idx + 1, len(rc)) if re.search(r"\d+\s*KG", rc[i], re.I)]
    country_col = part_idx + 1
    out = {}
    for i in range(hdr + 1, len(rows)):
        rc = ["" if c is None else str(c) for c in rows[i]]
        if rc[0].strip() and rc[0].strip() not in ("清关方式", "包税", "自税"):
            break
        if not rc[country_col].strip():
            continue
        if not re.search(r"[\u4e00-\u9fa5]", rc[country_col]):
            continue
        vals = [num(rc[c]) if c < len(rc) else None for c in tier_cols]
        vals = [v for v in vals if v is not None]
        if len(vals) < len(tier_mins):
            continue
        out[rc[country_col].strip()] = list(zip(tier_mins, vals[:len(tier_mins)]))
    return out

# ---- 欧洲自税递延 空运: group by 渠道代码 (EU-NCK0/EU-NQ0) ----
def parse_yf_self_air(wb):
    rows = rows_of(wb, "欧洲自税递延渠道")
    tier_mins = [21, 45, 100, 500, 1000]
    hdr = None
    for i, r in enumerate(rows):
        rc = ["" if c is None else str(c) for c in r]
        if "21KG" in "".join(rc) and "分区" in rc and "渠道代码" in rc:
            hdr = i; break
    if hdr is None:
        return {}
    rc = ["" if c is None else str(c) for c in rows[hdr]]
    code_idx = rc.index("渠道代码")
    part_idx = rc.index("分区")
    country_col = part_idx + 1
    tier_cols = [i for i in range(part_idx + 1, len(rc)) if re.search(r"\d+\s*KG", rc[i], re.I)]
    groups = {}
    cur_code = None
    for i in range(hdr + 1, len(rows)):
        rc = ["" if c is None else str(c) for c in rows[i]]
        j = "".join(rc)
        if "海运自税" in j:
            break
        code = rc[code_idx].strip()
        if code:
            cur_code = code
        if not cur_code or not rc[country_col].strip():
            continue
        if not re.search(r"[\u4e00-\u9fa5]", rc[country_col]):
            continue
        vals = [num(rc[c]) if c < len(rc) else None for c in tier_cols]
        vals = [v for v in vals if v is not None]
        if len(vals) < len(tier_mins):
            continue
        groups.setdefault(cur_code, {})[rc[country_col].strip()] = list(zip(tier_mins, vals[:len(tier_mins)]))
    return groups

# ---- 欧洲自税递延 海运卡派 (EU-NWD1/NWD2 = 德国 warehouses) ----
def parse_yf_self_sea(wb):
    rows = rows_of(wb, "欧洲自税递延渠道")
    tier_mins = [21, 45, 100, 500, 1000]
    hdr = None
    for i, r in enumerate(rows):
        rc = ["" if c is None else str(c) for c in r]
        if "海运自税" in "".join(rc):
            hdr = i; break
    if hdr is None:
        return None
    hi = None
    for j in range(max(0, hdr - 2), min(hdr + 3, len(rows))):
        rc = ["" if c is None else str(c) for c in rows[j]]
        if "分区" in rc and any(re.search(r"\d+\s*KG", c, re.I) for c in rc):
            hi = j; break
    if hi is None:
        return None
    rc = ["" if c is None else str(c) for c in rows[hi]]
    part_idx = rc.index("分区")
    country_col = part_idx + 1
    tier_cols = [i for i in range(country_col + 1, len(rc)) if re.search(r"\d+\s*KG", rc[i], re.I)]
    for i in range(hi + 1, len(rows)):
        rc = ["" if c is None else str(c) for c in rows[i]]
        pcell = rc[part_idx] if part_idx < len(rc) else ""
        if "德国" in pcell:
            vals = [num(rc[c]) if c < len(rc) else None for c in tier_cols]
            vals = [v for v in vals if v is not None]
            if len(vals) >= len(tier_mins):
                return {"德国": list(zip(tier_mins, vals[:len(tier_mins)]))}
    return None

# ---- 英国 空派/海运 自税/包税 (single 英国 country) ----
def parse_yf_uk(sheet, wb):
    rows = rows_of(wb, sheet)
    tier_mins = [21, 51, 101, 301]
    out = {}
    for i, r in enumerate(rows):
        rc = ["" if c is None else str(c) for c in r]
        if rc and rc[0] in ("自税", "包税") and num(rc[1]) is not None:
            kind = "self" if rc[0] == "自税" else "bao"
            vals = [num(x) for x in rc[1:1 + len(tier_mins)] if num(x) is not None]
            if len(vals) >= len(tier_mins):
                out["air_" + kind] = list(zip(tier_mins, vals[:len(tier_mins)]))
    for i, r in enumerate(rows):
        rc = ["" if c is None else str(c) for c in r]
        if "英国海运" in "".join(rc):
            for j in range(i + 1, min(i + 4, len(rows))):
                rj = ["" if c is None else str(c) for c in rows[j]]
                if rj and rj[0] in ("自税", "包税") and num(rj[1]) is not None:
                    kind = "self" if rj[0] == "自税" else "bao"
                    vals = [num(x) for x in rj[1:1 + len(tier_mins)] if num(x) is not None]
                    if len(vals) >= len(tier_mins):
                        out["sea_" + kind] = list(zip(tier_mins, vals[:len(tier_mins)]))
            break
    for i, r in enumerate(rows):
        rc = ["" if c is None else str(c) for c in r]
        if "英国卡航" in "".join(rc):
            for j in range(i + 1, min(i + 4, len(rows))):
                rj = ["" if c is None else str(c) for c in rows[j]]
                if rj and rj[0] in ("自税", "包税") and num(rj[1]) is not None:
                    kind = "self" if rj[0] == "自税" else "bao"
                    vals = [num(x) for x in rj[1:1 + 5] if num(x) is not None]
                    if len(vals) >= 4:
                        out["kahang_" + kind] = list(zip([21, 45, 100, 500], vals[:4]))
            break
    return out

# ---- 美国包税海卡 普船: 美西(left col group) / 美中+美东(right) ----
def parse_yf_us_hk(wb):
    rows = rows_of(wb, "美国包税海卡(普船-整柜直送)")
    tier_mins = [22, 100, 350]
    hdr = None
    for i, r in enumerate(rows):
        rc = ["" if c is None else str(c) for c in r]
        if "22KG" in "".join(rc) and "100KG" in "".join(rc):
            hdr = i; break
    if hdr is None:
        return {}
    left_rep = None; right_rep = None
    for i in range(hdr + 1, len(rows)):
        rc = ["" if c is None else str(c) for c in rows[i]]
        if "时效" in "".join(rc) and num(rc[2]) is None:
            continue
        lvals = [num(rc[2]), num(rc[3]), num(rc[4])]
        if left_rep is None and all(v is not None for v in lvals):
            left_rep = list(zip(tier_mins, lvals))
        rvals = [num(rc[9]) if len(rc) > 9 else None, num(rc[10]) if len(rc) > 10 else None, num(rc[11]) if len(rc) > 11 else None]
        if right_rep is None:
            if all(v is not None for v in rvals):
                right_rep = list(zip(tier_mins, rvals))
            else:
                r350 = num(rc[11]) if len(rc) > 11 else None
                if r350 is not None:
                    right_rep = list(zip(tier_mins, [r350, r350, r350]))
        if left_rep and right_rep:
            break
    out = {}
    if left_rep: out["美西"] = left_rep
    if right_rep:
        out["美中"] = right_rep
        out["美东"] = right_rep
    return out

# ---- 欧盟卡航包税 UPS派送 section (country-based) ----
def parse_yf_kahang_ups(wb):
    rows = rows_of(wb, "欧盟卡航包税")
    tier_mins = [21, 45, 100, 500, 1000]
    sec = None
    for i, r in enumerate(rows):
        rc = ["" if c is None else str(c) for c in r]
        if rc and rc[0] == "UPS派送":
            sec = i; break
    if sec is None:
        return {}
    hdr = None
    for j in range(sec, min(sec + 3, len(rows))):
        rc = ["" if c is None else str(c) for c in rows[j]]
        if "派送国家" in rc and any(re.search(r"\d+\s*KG", c, re.I) for c in rc):
            hdr = j; break
    if hdr is None:
        return {}
    rc = ["" if c is None else str(c) for c in rows[hdr]]
    part_idx = rc.index("分区")
    country_col = part_idx + 1
    tier_cols = [i for i in range(country_col + 1, len(rc)) if re.search(r"\d+\s*KG", rc[i], re.I)]
    out = {}
    for i in range(hdr + 1, len(rows)):
        rc = ["" if c is None else str(c) for c in rows[i]]
        if rc and rc[0] in ("DHL派送", "卡派直送"):
            break
        if not rc[country_col].strip() or not re.search(r"[\u4e00-\u9fa5]", rc[country_col]):
            continue
        vals = [num(rc[c]) if c < len(rc) else None for c in tier_cols]
        vals = [v for v in vals if v is not None]
        if len(vals) < len(tier_mins):
            continue
        out[rc[country_col].strip()] = list(zip(tier_mins, vals[:len(tier_mins)]))
    return out

# ---- 欧盟卡航包税 卡派直送 (限时达) section: 分区 -> countries ----
def parse_yf_kahang_tx(wb):
    rows = rows_of(wb, "欧盟卡航包税")
    tier_mins = [21, 45, 100, 500, 1000]
    sec = None
    for i, r in enumerate(rows):
        rc = ["" if c is None else str(c) for c in r]
        if rc and rc[0] == "卡派直送":
            sec = i; break
    if sec is None:
        return {}
    hdr = None
    for j in range(sec, min(sec + 3, len(rows))):
        rc = ["" if c is None else str(c) for c in rows[j]]
        if "派送仓库" in rc and any(re.search(r"\d+\s*KG", c, re.I) for c in rc):
            hdr = j; break
    if hdr is None:
        return {}
    rc = ["" if c is None else str(c) for c in rows[hdr]]
    part_idx = rc.index("分区")
    tier_cols = [i for i in range(part_idx + 1, len(rc)) if re.search(r"\d+\s*KG", rc[i], re.I)]
    zone_rate = {}
    for i in range(hdr + 1, len(rows)):
        rc = ["" if c is None else str(c) for c in rows[i]]
        if rc and rc[0] in ("DHL派送", "UPS派送"):
            break
        pcell = rc[part_idx] if part_idx < len(rc) else ""
        m = re.match(r"(\d)区", pcell)
        if not m:
            continue
        zone = m.group(1)
        vals = [num(rc[c]) if c < len(rc) else None for c in tier_cols]
        vals = [v for v in vals if v is not None]
        if len(vals) >= len(tier_mins):
            zone_rate[zone] = list(zip(tier_mins, vals[:len(tier_mins)]))
    return zone_rate

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--write", action="store_true"); args = ap.parse_args()
    wb = wb_open(NEWXLSX)
    rates = json.load(open(RATES, encoding="utf-8"))
    live = {c["id"]: c for c in rates["channels"] if c.get("carrier") == "亚丰(璞景/德翼)"}

    raw = {}
    raw["yf_us_kx"] = parse_yf_us_regions("美国空派快线", [10, 21, 71, 101], wb)
    raw["yf_us_jj"] = parse_yf_us_regions("美国空派经济线", [10, 21, 71, 101], wb)
    raw["yf_us_ms"] = parse_yf_us_regions("美森限时达", [12, 51, 100], wb)
    raw["yf_eu_kx"] = parse_yf_eu_zones("欧洲空运包税渠道", [21, 45, 100, 500, 1000], wb, ["单票", "注意", "限时达", "卡派", "DHL", "UPS派送"])
    raw["yf_eu_sea"] = parse_yf_eu_zones("欧盟海运包税", [21, 45, 100, 500, 1000], wb, ["单票", "注意", "限时达", "卡派", "DHL", "UPS派送"])
    raw["yf_eu_rail"] = parse_yf_eu_zones("欧盟快铁包税", [21, 45, 100, 500, 1000], wb, ["单票", "注意", "限时达", "卡派", "DHL", "UPS派送"])
    self_air = parse_yf_self_air(wb)
    raw["yf_eu_self_ck0"] = self_air.get("EU-NCK0", {})
    raw["yf_eu_self_nq0"] = self_air.get("EU-NQ0", {})
    raw["yf_eu_self_wd"] = parse_yf_self_sea(wb)
    uk = parse_yf_uk("英国空运-铁路-海运-卡航自税包税", wb)
    raw["yf_uk_air_self"] = uk.get("air_self")
    raw["yf_uk_air_bao"] = uk.get("air_bao")
    raw["yf_uk_sea_self"] = uk.get("sea_self")
    raw["yf_uk_sea_bao"] = uk.get("sea_bao")
    raw["yf_uk_kahang_self"] = uk.get("kahang_self")
    raw["yf_uk_kahang_bao"] = uk.get("kahang_bao")
    raw["yf_us_hk"] = parse_yf_us_hk(wb)
    raw["yf_eu_kahang"] = parse_yf_kahang_ups(wb)
    zone_tx = parse_yf_kahang_tx(wb)
    raw["yf_eu_kahang_tx"] = zone_tx
    wb.close()

    patch = {}
    gaps = []
    for cid, ch in live.items():
        if ch.get("cat") != "b2b-air":
            continue
        live_keys = list(ch["countries"].keys())
        if cid not in raw or not raw[cid]:
            gaps.append((cid, "no source")); continue
        src = raw[cid]
        if isinstance(src, list):
            ckey = "英国" if cid.startswith("yf_uk_") else (live_keys[0] if live_keys else None)
            src = {ckey: src} if ckey else {}
        new_countries = {}
        if cid == "yf_eu_kahang_tx":
            zr = src
            r1 = zr.get("1"); r2 = zr.get("2"); r3 = zr.get("3")
            for k in live_keys:
                rate = r1
                if k == "法国": rate = r2 or r1
                elif k == "西班牙": rate = r3 or r1
                if rate is None:
                    gaps.append((cid, "zone rate missing")); continue
                extra = {kk: vv for kk, vv in ch["countries"][k].items() if kk != "tiers"}
                new_countries[k] = {**extra, "tiers": tiers_from_choice(rate)}
        elif cid == "yf_eu_self_wd":
            rate = next(iter(src.values())) if src else None
            if rate is None:
                gaps.append((cid, "no 德国 rate")); continue
            for k in live_keys:
                extra = {kk: vv for kk, vv in ch["countries"][k].items() if kk != "tiers"}
                new_countries[k] = {**extra, "tiers": tiers_from_choice(rate)}
        else:
            parsed_map = {}
            for label, vals in src.items():
                for k in expand_keys(label, live_keys):
                    if k in live_keys and k not in parsed_map:
                        parsed_map[k] = vals
            for k in live_keys:
                extra = {kk: vv for kk, vv in ch["countries"][k].items() if kk != "tiers"}
                if k in parsed_map:
                    new_countries[k] = {**extra, "tiers": tiers_from_choice(parsed_map[k])}
                else:
                    gaps.append((cid, "country not matched: " + k))
                    new_countries[k] = ch["countries"][k]
        patch[cid] = {"countries": new_countries}

    print("=== SELF-CHECK ===")
    for cid in sorted(patch):
        c = patch[cid]["countries"]
        sk = list(c.keys())[:1]
        sample = {k: c[k]["tiers"][:2] for k in sk}
        print(f"{cid:20s} #c={len(c):2d} sample={sample}")
    print("\n=== GAPS ===")
    for g in gaps:
        print("  gap:", g)

    json.dump(patch, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("\nWROTE", OUT, "channels:", len(patch))
    if not args.write:
        print("[DRY-RUN] re-run with --write to apply")
        return
    applied = 0
    for cid, ch in live.items():
        if cid not in patch: continue
        ch["countries"] = patch[cid]["countries"]
        ch["dest"] = list(patch[cid]["countries"].keys())
        applied += 1
    json.dump(rates, open(RATES, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"\n[WRITE] applied {applied} channels")

if __name__ == "__main__":
    main()

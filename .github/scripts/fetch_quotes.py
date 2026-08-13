# -*- coding: utf-8 -*-
"""
fetch_quotes.py — 云端定时抓报价单并自动更新基础价格（GitHub Actions 内运行）。
流程: 用 wl02@ 的 user_access_token 搜邮箱报价单 -> 下附件 -> 解压选主表
      -> 调各 bake_*.py 解析(只改价格数值) -> 白名单 diff 守铁律
      -> 若有变化: 重新生成 rates.js + 飞书群推「变了什么」卡片。
依赖: openpyxl（bake 需要）；user token 来自 secret WL02_MAIL_TOKEN（JSON）。
本地回归: 设 FETCH_LOCAL=1 则从本地 quote_pull/{carrier}/latest.xlsx 复制，不联网。
"""
import json, os, sys, urllib.request, urllib.parse, zipfile, shutil, subprocess, time, datetime, re

HERE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(os.path.dirname(HERE))          # .github/scripts -> repo root
RATES = os.path.join(REPO_ROOT, "rates.json")
RATES_JS = os.path.join(REPO_ROOT, "rates.js")
VERSION = os.path.join(REPO_ROOT, "version.json")

APP_ID = "cli_aab6341b78f95be9"
APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")   # 云端经 workflow secrets 注入，禁止硬编码
MBOX = "wl02@leading-trade.com"
CHAT_ID = "oc_cd9a6f072cdd348a08c29d09e8c9143a"
SITE_URL = "https://heryma99.github.io/jwpei-logistics-compare/"

CARRIERS = {
    "云途":     {"kw": ["云途", "客户报价单"], "main": "云途全球专线挂号（特惠普货）", "prefer_ext": "zip"},
    "中运通达": {"kw": ["中运通达", "中运"],    "main": None, "prefer_ext": "xlsx"},
    "亚丰":     {"kw": ["璞景", "德翼供应链", "德翼价格表", "德翼", "亚丰"], "main": None, "prefer_ext": "xlsx"},
}
BAKE = {
    "云途":     ["bake_yuntu.py"],
    "中运通达": ["bake_zy_package.py", "bake_commercial.py"],
    "亚丰":     ["bake_yf.py"],
}
MAIL_API = "https://open.feishu.cn/open-apis/mail/v1/user_mailboxes/" + MBOX
REFRESH_URL = "https://open.feishu.cn/open-apis/authen/v1/oidc/refresh_access_token"

FETCH_LOCAL = os.environ.get("FETCH_LOCAL", "") == "1"
LOCAL_XLSX_DIR = os.environ.get("LOCAL_XLSX_DIR", os.path.join(REPO_ROOT, "quote_pull"))

# 每趟运行对各承运商的处理结果汇总（用于诊断卡片 / 日志）
REPORT = []


def log(m):
    print(f"[fetch] {m}", flush=True)


def req_json(url, method="GET", token=None, body=None):
    h = {"Content-Type": "application/json", "User-Agent": "fetch-quotes"}
    if token:
        h["Authorization"] = "Bearer " + token
    r = urllib.request.Request(url, data=json.dumps(body).encode() if body is not None else None,
                               method=method, headers=h)
    with urllib.request.urlopen(r, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))


# ---------- token ----------
def get_app_access_token():
    """用 App Secret 换取 app_access_token，供 authen 系列接口鉴权。"""
    if not APP_SECRET:
        raise RuntimeError("FEISHU_APP_SECRET 未配置")
    url = "https://open.feishu.cn/open-apis/auth/v3/app_access_token/internal"
    r = req_json(url, "POST", body={"app_id": APP_ID, "app_secret": APP_SECRET})
    if r.get("code") != 0:
        raise RuntimeError(f"获取 app_access_token 失败: {r}")
    return r["app_access_token"]


def load_token():
    raw = os.environ.get("WL02_MAIL_TOKEN", "")
    if not raw:
        log("[skip] 无 WL02_MAIL_TOKEN secret，跳过本次抓取")
        return None
    try:
        tok = json.loads(raw)
    except Exception:
        log("[skip] WL02_MAIL_TOKEN 解析失败")
        return None
    now = int(time.time())
    expires_in = int(tok.get("expires_in") or 7200)
    if tok.get("access_token") and now < tok.get("got_at", 0) + expires_in - 300:
        return tok["access_token"]
    # refresh: authen v1 OIDC 要求在 Authorization 头传 app_access_token
    try:
        app_token = get_app_access_token()
        r = req_json(REFRESH_URL, "POST", token=app_token, body={
            "grant_type": "refresh_token",
            "refresh_token": tok["refresh_token"],
        })
        if r.get("code") == 0:
            data = r.get("data", r)
            tok["access_token"] = data["access_token"]
            tok["refresh_token"] = data.get("refresh_token", tok["refresh_token"])
            tok["expires_in"] = data.get("expires_in", 7200)
            tok["got_at"] = now
            log("refresh 成功")
            save_token_to_secret(tok)
            return tok["access_token"]
        else:
            log(f"refresh 失败: {r}")
    except Exception as e:
        log(f"refresh 失败: {e}")
    return None


def save_token_to_secret(tok):
    """刷新成功后把新 token 写回 GitHub secret，实现零维护自动续期。
    需 workflow 提供 GITHUB_TOKEN + actions:write 权限；失败仅告警不影响本次抓取。"""
    gt = os.environ.get("GITHUB_TOKEN")
    repo = os.environ.get("GITHUB_REPO")
    if not gt or not repo:
        log("[skip] 未配置 GITHUB_TOKEN/GITHUB_REPO，跳过写回 secret（后续过期需手动重授权）")
        return
    try:
        import base64 as _b64
        from nacl.public import PublicKey, SealedBox
        import nacl.encoding
        H = {"Authorization": f"Bearer {gt}", "Accept": "application/vnd.github+json",
             "Content-Type": "application/json", "X-GitHub-Api-Version": "2022-11-28"}
        def _api(method, path, body=None):
            data = json.dumps(body).encode() if body is not None else None
            req = urllib.request.Request("https://api.github.com" + path, data=data, headers=H, method=method)
            with urllib.request.urlopen(req, timeout=30) as r:
                b = r.read().decode()
                return r.status, (json.loads(b) if b.strip() else {})
        st, keyinfo = _api("GET", f"/repos/{repo}/actions/secrets/public-key")
        if st != 200:
            log(f"[warn] 获取 public-key 失败 {st}"); return
        pk = PublicKey(keyinfo["key"].encode(), encoder=nacl.encoding.Base64Encoder)
        box = SealedBox(pk)
        enc = nacl.encoding.Base64Encoder.encode(box.encrypt(json.dumps(tok).encode())).decode()
        st2, _ = _api("PUT", f"/repos/{repo}/actions/secrets/WL02_MAIL_TOKEN",
                      {"encrypted_value": enc, "key_id": keyinfo["key_id"]})
        if st2 in (201, 204):
            log("✅ 新 token 已自动写回 WL02_MAIL_TOKEN secret（零维护自动续期生效）")
        else:
            log(f"[warn] 写回 secret 返回 {st2}")
    except Exception as e:
        log(f"[warn] 写回 secret 失败（不影响本次抓取）: {e}")


# ---------- 邮件 ----------
def search_mail(token, kw):
    # 飞书邮件搜索正确端点：POST /open-apis/mail/v1/user_mailboxes/{MBOX}/search
    # （注意：不是 /messages/search，那个路径返回 404）
    try:
        r = req_json(MAIL_API + "/search", "POST", token, {"query": kw, "page_size": 15})
    except Exception as e:
        log(f"search '{kw}' 异常: {e}")
        REPORT.append(f"❌ 搜索「{kw}」异常: {e}")
        return []
    if r.get("code") != 0:
        log(f"search '{kw}' 错误: {r.get('code')} {r.get('msg')}")
        REPORT.append(f"❌ 搜索「{kw}」错误码 {r.get('code')}: {r.get('msg')}")
        return []
    msgs = r.get("data", {}).get("items", []) or []
    REPORT.append(f"· 搜索「{kw}」: 命中 {len(msgs)} 封")
    return msgs


def get_attachments(token, mid):
    r = req_json(MAIL_API + f"/messages/{mid}", "GET", token)
    if r.get("code") != 0:
        return []
    return r.get("data", {}).get("message", {}).get("attachments", []) or []


def download_url(token, mid, aid):
    # 飞书下载接口：单个 id 直接拼在 query 上（不要传 JSON 数组）；
    # 响应字段是 data.download_urls[]（不是 url_list）
    r = req_json(MAIL_API + f"/messages/{mid}/attachments/download_url?attachment_ids={aid}", "GET", token)
    if r.get("code") != 0:
        return None
    for d in r.get("data", {}).get("download_urls", []):
        if d.get("attachment_id") == aid and d.get("download_url"):
            return d["download_url"]
    return None


def pick_target(atts, prefer_ext):
    # 真实附件 attachment_type 为真（如 1）；内联图片 attachment_type 为 None
    real = [a for a in atts if a.get("attachment_type") and not str(a.get("filename", "")).lower().startswith("image")]
    if prefer_ext:
        t = next((a for a in real if str(a.get("filename", "")).lower().endswith(prefer_ext)), None)
        if t:
            return t
    for ext in (".zip", ".xlsx", ".docx"):
        t = next((a for a in real if str(a.get("filename", "")).lower().endswith(ext)), None)
        if t:
            return t
    return real[0] if real else None


def pick_main_xlsx(zippath, main_sheet):
    import openpyxl, io
    # 直接内存读取 zip 内 xlsx 成员，避免 Windows 解压中文文件名乱码导致路径找不到
    with zipfile.ZipFile(zippath) as z:
        names = [n for n in z.namelist() if n.lower().endswith(".xlsx")]
        chosen = None
        if main_sheet:
            best = -1
            for n in names:
                try:
                    wb = openpyxl.load_workbook(io.BytesIO(z.read(n)), read_only=True)
                    if main_sheet in wb.sheetnames and len(wb.sheetnames) > best:
                        chosen, best = n, len(wb.sheetnames)
                    wb.close()
                except Exception:
                    pass
        if not chosen and names:
            chosen = max(names, key=lambda n: z.getinfo(n).date_time)
        if not chosen:
            return None
        tmp = zippath + ".x"
        open(tmp, "wb").write(z.read(chosen))
        return tmp


def process_carrier(token, name, cfg):
    outdir = os.path.join(HERE, name)
    os.makedirs(outdir, exist_ok=True)
    final = os.path.join(outdir, "latest.xlsx")
    if FETCH_LOCAL:
        src = os.path.join(LOCAL_XLSX_DIR, name, "latest.xlsx")
        if not os.path.exists(src):
            log(f"[{name}] 本地回归缺 {src}")
            return False
        shutil.copy(src, final)
        log(f"[{name}] 本地回归用 {src}")
        REPORT.append(f"· [{name}] 本地回归用 {src}")
        return True
    # 联网抓取
    for kw in cfg["kw"]:
        items = search_mail(token, kw)
        if items:
            log(f"[{name}] 关键词「{kw}」命中 {len(items)} 封")
            break
    else:
        log(f"[{name}] 未搜到相关邮件")
        REPORT.append(f"· [{name}] 所有关键词均未搜到报价邮件")
        return False
    for em in items:
        mid = em.get("message_id") or em.get("id")
        atts = get_attachments(token, mid)
        tgt = pick_target(atts, cfg["prefer_ext"])
        if not tgt:
            continue
        url = download_url(token, mid, tgt["id"])
        if not url:
            log(f"[{name}] 拿不到下载链接")
            REPORT.append(f"❌ [{name}] 邮件有附件但拿不到下载链接")
            return False
        raw = os.path.join(outdir, "raw_" + re.sub(r"[^\w.]", "_", str(tgt.get("filename", "x"))))
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        data = urllib.request.urlopen(req, timeout=180).read()
        open(raw, "wb").write(data)
        if raw.lower().endswith(".zip"):
            cx = pick_main_xlsx(raw, cfg["main"])
            if cx:
                shutil.copy(cx, final)
                if cx != raw:
                    os.remove(cx)
            os.remove(raw)
        else:
            shutil.copy(raw, final)
            os.remove(raw)
        log(f"[{name}] 已落盘 {final} ({len(data)} bytes)")
        REPORT.append(f"✅ [{name}] 已下载报价单 {tgt.get('filename')} ({len(data)} bytes) 并落盘")
        return True
    log(f"[{name}] 候选邮件均无报价附件")
    REPORT.append(f"· [{name}] 候选邮件均无报价附件")
    return False


# ---------- 白名单 diff（守铁律：只允许数值变化） ----------
def is_num_keys(d):
    if not d:
        return False
    for k in d:
        try:
            float(k)
        except (TypeError, ValueError):
            return False
    return True


def norm_keys(d):
    return {float(k): k for k in d} if (d and is_num_keys(d)) else None


def struct_check(path, a, b, rep):
    if isinstance(a, dict) and isinstance(b, dict):
        na, nb = norm_keys(a), norm_keys(b)
        for k in a:
            present = (k in b) if na is None else (float(k) in nb)
            if not present:
                rep.append("FAIL 删除字段 %s.%s" % (path, k)); return False
        for k in b:
            present = (k in a) if na is None else (float(k) in na)
            if not present:
                if is_num_keys(a) or is_num_keys(b):
                    rep.append("FAIL 档位键变化 %s:+%s" % (path, k)); return False
                rep.append("WARN 新增字段 %s.%s" % (path, k))
        for k in a:
            kb = k if na is None else nb.get(float(k))
            if kb is not None and not struct_check(path + "." + k, a[k], b[kb], rep):
                return False
        return True
    if isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            rep.append("FAIL 数组长度 %s:%d->%d" % (path, len(a), len(b))); return False
        for i, (x, y) in enumerate(zip(a, b)):
            if not struct_check("%s[%d]" % (path, i), x, y, rep):
                return False
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) \
            and not isinstance(a, bool) and not isinstance(b, bool):
        return True
    if type(a) != type(b):
        rep.append("FAIL 类型变化 %s:%s->%s" % (path, type(a).__name__, type(b).__name__)); return False
    return True


def whitelist_ok(old, new):
    oc = {c["id"]: c for c in old["channels"]}
    nc = {c["id"]: c for c in new["channels"]}
    rep = []
    for cid in (set(oc) ^ set(nc)):
        rep.append("FAIL 渠道增删: " + cid)
    for cid in set(oc) & set(nc):
        struct_check("channels." + cid, oc[cid], nc[cid], rep)
    fails = [r for r in rep if r.startswith("FAIL")]
    for r in rep:
        log(r)
    return not fails


# ---------- 变更卡片（复用 notify 逻辑，内存 old/new 对比） ----------
def flatten(o, p=""):
    out = {}
    if isinstance(o, dict):
        for k, v in o.items():
            out.update(flatten(v, f"{p}{k}.")) if isinstance(v, (dict, list)) else out.update({p + k: v})
    elif isinstance(o, list):
        for i, v in enumerate(o):
            out.update(flatten(v, f"{p}[{i}].")) if isinstance(v, (dict, list)) else out.update({f"{p}[{i}]": v})
    return out


def humanize(p):
    m = re.match(r"country_overrides\.(.+?)\.volumetric_divisor$", p)
    if m: return f"{m.group(1)} 体积重除数"
    if p.endswith("volumetric_divisor"): return "体积重除数"
    m = re.match(r"countries\.(.+?)\.service_fee_pct$", p)
    if m: return f"{m.group(1)} 服务费比例"
    m = re.match(r"countries\.(.+?)\.brackets\.\[(\d+)\]\.rate$", p)
    if m: return f"{m.group(1)} 第{int(m.group(2))+1}段费率(元/kg)"
    m = re.match(r"countries\.(.+?)\.brackets\.\[(\d+)\]\.(.+)$", p)
    if m: return f"{m.group(1)} 第{int(m.group(2))+1}段 {m.group(3)}"
    m = re.match(r"countries\.(.+?)\.tiers\.\[(\d+)\]\.rate$", p)
    if m: return f"{m.group(1)} 第{int(m.group(2))+1}段费率(元/kg)"
    m = re.match(r"countries\.(.+?)\.tiers\.\[(\d+)\]\.(.+)$", p)
    if m: return f"{m.group(1)} 第{int(m.group(2))+1}段 {m.group(3)}"
    if "matrix" in p: return "分区矩阵 " + p.split(".")[-1]
    return p.split(".")[-1]


def fmt_val(p, v):
    if isinstance(v, float) and "pct" in p:
        return f"{v*100:.2f}%"
    if isinstance(v, float):
        return f"{v:g}"
    return str(v)


def compute_diff(old_d, new_d):
    oc = {c["id"]: c for c in old_d["channels"]}
    nc = {c["id"]: c for c in new_d["channels"]}
    added = [c for c in nc if c not in oc]
    removed = [c for c in oc if c not in nc]
    changed = []
    for cid in sorted(set(oc) & set(nc)):
        a, b = flatten(oc[cid]), flatten(nc[cid])
        dc = [(humanize(k), fmt_val(k, a.get(k)) if a.get(k) is not None else "—",
               fmt_val(k, b.get(k)) if b.get(k) is not None else "—")
              for k in sorted(set(a) | set(b)) if a.get(k) != b.get(k)]
        if dc:
            changed.append((cid, nc[cid].get("name", cid), dc))
    return added, removed, changed


def build_card(committer, ctime, added, removed, changed):
    lines = [f"**更新人**：{committer}", f"**更新时间**：{ctime}",
             f"**变动规模**：{len(changed)} 个渠道价格变动（新增 {len(added)} / 移除 {len(removed)}）", ""]
    if changed:
        lines.append("📝 **价格变动明细**（自动抓单解析）：")
        for cid, name, dc in changed[:18]:
            lines.append(f"\n**{name}**  `{cid}`")
            for label, ov, nv in dc[:8]:
                lines.append(f"  · {label}：{ov} → {nv}")
            if len(dc) > 8:
                lines.append(f"  · …另 {len(dc)-8} 处")
        if len(changed) > 18:
            lines.append(f"\n…共 {len(changed)} 个渠道变动")
    text = "\n".join(lines)
    return {"config": {"wide_screen_mode": True},
            "header": {"title": {"tag": "plain_text", "content": "比价易 · 报价单自动更新"},
                       "template": "blue"},
            "elements": [{"tag": "div", "text": {"tag": "lark_md", "content": text}},
                         {"tag": "hr"},
                         {"tag": "action", "actions": [{"tag": "button", "text": {"tag": "plain_text", "content": "查看比价易"},
                                                       "type": "primary", "url": SITE_URL}]}]}


def send_card(card):
    if not APP_SECRET:
        log("未配置 FEISHU_APP_SECRET，跳过发卡片")
        return
    try:
        t = req_json("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal", "POST",
                     body={"app_id": APP_ID, "app_secret": APP_SECRET})
        tok = t["tenant_access_token"]
        body = {"receive_id": CHAT_ID, "msg_type": "interactive", "content": json.dumps(card)}
        r = req_json("https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id", "POST", tok, body)
        if r.get("code") == 0:
            log("飞书卡片发送成功 -> " + CHAT_ID)
        else:
            log("飞书卡片错误: " + str(r)[:200])
    except Exception as e:
        log(f"发卡片失败: {e}")


# ---------- bake rates.js ----------
def bake_ratesjs():
    import time as _t
    d = json.load(open(RATES, encoding="utf-8"))
    eff = (d.get("meta") or {}).get("effective_date") or _t.strftime("%Y-%m-%d")
    d["generated"] = eff
    BANNER = "\n;(function(){function render(){var R=window.RATES||{};var g=R.generated||(R.meta&&R.meta.effective_date);var note=(R.meta&&R.meta.banner)||(R.meta&&R.meta.note)||'';var h=document.querySelector('header');if(!h||!g)return;var el=document.getElementById('effBanner');if(!el){el=document.createElement('div');el.id='effBanner';el.style.cssText='margin-top:8px;padding:6px 12px;border-radius:8px;font-size:13px;font-weight:600;color:#0e1117;background:linear-gradient(90deg,#ffd479,#ffb347);box-shadow:0 1px 4px rgba(0,0,0,.25);display:inline-block';h.appendChild(el);}el.textContent='\\U0001F4C5 '+note;}if(document.readyState!=='loading')render();else document.addEventListener('DOMContentLoaded',render);})();\n"
    with open(RATES_JS, "w", encoding="utf-8") as f:
        f.write("window.RATES = ")
        json.dump(d, f, ensure_ascii=False)
        f.write(";")
        f.write(BANNER)
    json.dump({"build": str(int(_t.time()))}, open(VERSION, "w", encoding="utf-8"))
    log(f"rates.js 重新烘焙 | effective_date={eff}")


def send_error_card(title, msg):
    send_card({
        "config": {"wide_screen_mode": True},
        "header": {"title": {"tag": "plain_text", "content": title}, "template": "red"},
        "elements": [{"tag": "div", "text": {"tag": "lark_md",
                  "content": "**自动抓单运行异常**（需人工排查）\n```\n" + msg[-1800:] + "\n```"}}],
    })


def main():
    try:
        _main_impl()
    except SystemExit:
        raise
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        log("❌ 运行异常: " + str(e))
        send_error_card("比价易 · 抓单异常", tb)
        raise


def _main_impl():
    if not os.path.exists(RATES):
        log("[skip] 无 rates.json")
        return
    old = json.load(open(RATES, encoding="utf-8"))
    token = load_token()
    if token is None and not FETCH_LOCAL:
        log("❌ 无法获取 wl02 邮箱 token，本次运行直接失败（不会静默跳过）")
        send_error_card("比价易 · 抓单失败", "无法获取 wl02 邮箱 user token：access_token 已过期且 refresh_token 续期失败。\n请在本机重跑 quote_pull/oauth_authorize.py 让用户在浏览器点一次「允许」，重新写入 WL02_MAIL_TOKEN secret。")
        sys.exit(1)
    # 抓取各承运商报价单
    got_any = False
    for name, cfg in CARRIERS.items():
        try:
            if process_carrier(token, name, cfg):
                got_any = True
        except Exception as e:
            log(f"[{name}] 处理异常: {e}")
    # 调 bake 解析（只改价格数值）
    for name, scripts in BAKE.items():
        for sc in scripts:
            try:
                subprocess.run([sys.executable, os.path.join(HERE, sc), "--write"],
                               check=True, capture_output=True, text=True, timeout=240)
                log(f"[{name}] {sc} 完成")
            except subprocess.CalledProcessError as e:
                log(f"[{name}] {sc} 失败: {e.stderr[-300:] if e.stderr else e}")
    new = json.load(open(RATES, encoding="utf-8"))
    # 白名单守铁律
    if not whitelist_ok(old, new):
        log("❌ 白名单 FAIL：结构/渠道被改动，禁止提交！已恢复原 rates.json")
        json.dump(old, open(RATES, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        return
    # 计算变更
    added, removed, changed = compute_diff(old, new)
    if not (added or removed or changed):
        log("✅ 无价格数值变化，跳过提交与通知")
        if os.environ.get("FETCH_DIAG") == "1":
            dlog = ["**【诊断模式】本次自动抓单明细**", ""]
            dlog += REPORT or ["（无各承运商处理记录）"]
            dlog += ["", "结论：邮箱抓取/解析均执行完毕，当前 rates.json 与最新报价单相比无价格数值变化。"]
            send_card({
                "config": {"wide_screen_mode": True},
                "header": {"title": {"tag": "plain_text", "content": "比价易 · 抓单诊断"}, "template": "grey"},
                "elements": [{"tag": "div",
                              "text": {"tag": "lark_md", "content": "\n".join(dlog)}}],
            })
        return
    # 重新烘焙 + 发卡片
    bake_ratesjs()
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    card = build_card("比价易·自动抓单", now, added, removed, changed)
    send_card(card)
    log(f"✅ 价格变动 {len(changed)} 渠道，已更新 rates.json/rates.js，待 yml 提交")


if __name__ == "__main__":
    main()

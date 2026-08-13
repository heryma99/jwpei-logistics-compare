#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
飞书群通知（应用 API 版，纯云端运行于 GitHub Actions）
- 团队在配置中心改 rates.json -> GitHub Action 触发 -> 本脚本：
  1) 拉取【本次版本】与【上一个版本】的 rates.json
  2) 计算差异（哪些渠道、哪些价格字段变了）
  3) 向「物流报价监控」群推送一张带【变更明细】的卡片
- 未配置 FEISHU_APP_SECRET / FEISHU_WEBHOOK 时安全跳过（不影响部署）
- 仅在云端运行；本地调试可设 DRY_RUN=1 只打印卡片 JSON 不发。
"""
import os, json, urllib.request, datetime, sys

REPO = os.environ.get("GITHUB_REPOSITORY", "heryma99/jwpei-logistics-compare")
GH_TOKEN = os.environ.get("GITHUB_TOKEN", "")          # 用于拉取历史 rates.json（contents:read）
BEFORE_SHA = os.environ.get("BEFORE_SHA", "")          # push 事件的上一版 commit sha
AFTER_SHA = os.environ.get("AFTER_SHA", os.environ.get("GITHUB_SHA", ""))
COMMITTER = os.environ.get("COMMITTER", "配置中心")
COMMIT_TIME = os.environ.get("COMMIT_TIME", "")
SITE_URL = "https://heryma99.github.io/jwpei-logistics-compare/"

APP_ID = "cli_aab6341b78f95be9"
APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")
CHAT_ID = "oc_cd9a6f072cdd348a08c29d09e8c9143a"
WEBHOOK = os.environ.get("FEISHU_WEBHOOK", "")

DRY_RUN = os.environ.get("DRY_RUN", "") == "1"

log = lambda m: print(f"[notify] {m}", flush=True)


# ---------- 拉取某个 ref 的 rates.json ----------
def fetch_rates(ref):
    if not ref:
        return None
    url = f"https://api.github.com/repos/{REPO}/contents/rates.json?ref={ref}"
    req = urllib.request.Request(url, method="GET")
    req.add_header("Authorization", "Bearer " + GH_TOKEN)
    req.add_header("Accept", "application/vnd.github.raw")
    req.add_header("User-Agent", "rates-notify")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception as e:
        log(f"拉取 rates.json@{ref[:8]} 失败: {e}")
        return None


# ---------- 扁平化 ----------
def flatten(o, prefix=""):
    out = {}
    if isinstance(o, dict):
        for k, v in o.items():
            if isinstance(v, (dict, list)):
                out.update(flatten(v, f"{prefix}{k}."))
            else:
                out[prefix + k] = v
    elif isinstance(o, list):
        for i, v in enumerate(o):
            if isinstance(v, (dict, list)):
                out.update(flatten(v, f"{prefix}[{i}]."))
            else:
                out[f"{prefix}[{i}]"] = v
    return out


# ---------- 人类可读标签 ----------
def humanize(path):
    p = path
    # country_overrides.{国家}.volumetric_divisor
    m = __import__("re").match(r"country_overrides\.(.+?)\.volumetric_divisor$", p)
    if m:
        return f"{m.group(1)} 体积重除数"
    if p.endswith("volumetric_divisor"):
        return "体积重除数"
    m = __import__("re").match(r"countries\.(.+?)\.service_fee_pct$", p)
    if m:
        return f"{m.group(1)} 服务费比例"
    m = __import__("re").match(r"countries\.(.+?)\.brackets\.\[(\d+)\]\.rate$", p)
    if m:
        return f"{m.group(1)} 第{int(m.group(2))+1}段费率(元/kg)"
    m = __import__("re").match(r"countries\.(.+?)\.brackets\.\[(\d+)\]\.(.+)$", p)
    if m:
        return f"{m.group(1)} 第{int(m.group(2))+1}段 {m.group(3)}"
    m = __import__("re").match(r"countries\.(.+?)\.min_box$", p)
    if m:
        return f"{m.group(1)} 最低计费重(kg)"
    if "product_fees" in p:
        return "产品附加费"
    if "surcharges" in p:
        seg = p.split(".")[-1]
        return f"附加费.{seg}"
    return p.split(".")[-1]


def fmt_val(path, v):
    if isinstance(v, float) and "pct" in path:
        return f"{v*100:.2f}%"
    if isinstance(v, float):
        return f"{v:g}"
    return str(v)


# ---------- 单渠道 diff ----------
def diff_channel(old_c, new_c):
    a = flatten(old_c)
    b = flatten(new_c)
    res = []
    for k in sorted(set(a) | set(b)):
        ov, nv = a.get(k), b.get(k)
        if ov != nv:
            res.append((humanize(k), fmt_val(k, ov) if ov is not None else "—",
                        fmt_val(k, nv) if nv is not None else "—"))
    return res


# ---------- 全量 diff ----------
def compute_diff(old_d, new_d):
    old_ch = {c["id"]: c for c in old_d["channels"]} if old_d else {}
    new_ch = {c["id"]: c for c in new_d["channels"]}
    added = [cid for cid in new_ch if cid not in old_ch]
    removed = [cid for cid in old_ch if cid not in new_ch]
    changed = []
    for cid in sorted(set(old_ch) & set(new_ch)):
        dc = diff_channel(old_ch[cid], new_ch[cid])
        if dc:
            changed.append((cid, new_ch[cid].get("name", cid), dc))
    return added, removed, changed


# ---------- 构建卡片 ----------
def build_card(committer, ctime, added, removed, changed):
    n_ch = len(changed) + len(added) + len(removed)
    lines = []
    lines.append(f"**更新人**：{committer}")
    if ctime:
        try:
            dt = datetime.datetime.fromisoformat(ctime.replace("Z", "+00:00"))
            lines.append(f"**更新时间**：{dt.strftime('%Y-%m-%d %H:%M')} (UTC+8 近似)")
        except Exception:
            lines.append(f"**更新时间**：{ctime}")
    lines.append(f"**变动规模**：{n_ch} 个渠道有变化（新增 {len(added)} / 移除 {len(removed)} / 价格变动 {len(changed)}）")
    lines.append("")

    if added:
        lines.append("➕ **新增渠道**：" + "、".join(added[:10]))
        lines.append("")
    if removed:
        lines.append("➖ **移除渠道**：" + "、".join(removed[:10]))
        lines.append("")

    if changed:
        lines.append("📝 **价格变动明细**：")
        for cid, name, dc in changed[:18]:
            lines.append(f"\n**{name}**  `{cid}`")
            for label, ov, nv in dc[:8]:
                lines.append(f"  · {label}：{ov} → {nv}")
            if len(dc) > 8:
                lines.append(f"  · …另 {len(dc)-8} 处变动")
        if len(changed) > 18:
            lines.append(f"\n…共 {len(changed)} 个渠道变动，完整见 rates.json")
    else:
        if not added and not removed:
            lines.append("（仅结构/说明微调，无价格数值变化）")

    # 飞书 markdown 文本
    text = "\n".join(lines)
    card = {
        "config": {"wide_screen_mode": True},
        "header": {
            "title": {"tag": "plain_text", "content": "比价易 · 报价已更新"},
            "template": "blue",
        },
        "elements": [
            {"tag": "div", "text": {"tag": "lark_md", "content": text}},
            {"tag": "hr"},
            {"tag": "action", "actions": [
                {"tag": "button", "text": {"tag": "plain_text", "content": "查看比价易"},
                 "type": "primary", "url": SITE_URL},
            ]},
        ],
    }
    return card


# ---------- 发送：应用 API ----------
def send_app_api(card):
    if not APP_SECRET:
        log("未配置 FEISHU_APP_SECRET，跳过应用API发送")
        return False
    # 取 tenant_access_token
    req = urllib.request.Request("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
                                 method="POST",
                                 data=json.dumps({"app_id": APP_ID, "app_secret": APP_SECRET}).encode())
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            t = json.loads(r.read().decode())
    except Exception as e:
        log(f"获取 tenant_access_token 失败: {e}")
        return False
    if t.get("code") != 0:
        log(f"tenant_access_token 错误: {t}")
        return False
    tok = t["tenant_access_token"]
    body = {"receive_id": CHAT_ID, "msg_type": "interactive", "content": json.dumps(card)}
    req2 = urllib.request.Request("https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
                                  method="POST", data=json.dumps(body).encode())
    req2.add_header("Content-Type", "application/json")
    req2.add_header("Authorization", "Bearer " + tok)
    try:
        with urllib.request.urlopen(req2, timeout=30) as r:
            resp = json.loads(r.read().decode())
    except Exception as e:
        log(f"发送消息失败: {e}")
        return False
    if resp.get("code") != 0:
        log(f"发送消息错误: {resp}")
        return False
    log(f"飞书通知(应用API)发送成功 -> chat {CHAT_ID}")
    return True


# ---------- 发送：webhook 兜底 ----------
def send_webhook(card):
    if not WEBHOOK:
        return False
    body = {"msg_type": "interactive", "card": card}
    req = urllib.request.Request(WEBHOOK, method="POST", data=json.dumps(body).encode())
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            resp = json.loads(r.read().decode())
    except Exception as e:
        log(f"webhook 发送失败: {e}")
        return False
    if resp.get("code") != 0:
        log(f"webhook 错误: {resp}")
        return False
    log("飞书通知(webhook)发送成功")
    return True


def main():
    log(f"FEISHU_APP_SECRET: {'***已配置***' if APP_SECRET else '未配置'} | BEFORE_SHA: {BEFORE_SHA[:8] if BEFORE_SHA else '空'} | AFTER_SHA: {AFTER_SHA[:8] if AFTER_SHA else '空'}")
    new_d = fetch_rates(AFTER_SHA) or json.load(open("rates.json", encoding="utf-8"))
    old_d = fetch_rates(BEFORE_SHA)
    if old_d is None:
        log("无前序版本可对比（首次/手动触发），卡片仅含摘要")
    added, removed, changed = compute_diff(old_d, new_d)
    card = build_card(COMMITTER, COMMIT_TIME, added, removed, changed)
    log(f"变更统计: 新增{len(added)} 移除{len(removed)} 价格变动{len(changed)}")
    if DRY_RUN:
        print("=== DRY_RUN 卡片 JSON ===")
        print(json.dumps(card, ensure_ascii=False, indent=2))
        return
    if not (APP_SECRET or WEBHOOK):
        log("未配置任何飞书凭证，跳过发送（不影响部署）")
        return
    ok = send_app_api(card)
    if not ok and WEBHOOK:
        send_webhook(card)


if __name__ == "__main__":
    main()

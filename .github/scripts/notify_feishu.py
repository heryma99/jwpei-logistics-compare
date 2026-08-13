#!/usr/bin/env python3
# 飞书群通知（GitHub Action 用，纯标准库）
# 优先用「应用 API」(tenant_access_token) 发到指定群；未配 APP_SECRET 时回退 webhook。
import os, json, sys, urllib.request, urllib.error

APP_ID = os.environ.get("FEISHU_APP_ID", "").strip() or "cli_aab6341b78f95be9"
APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "").strip()
CHAT_ID = os.environ.get("FEISHU_CHAT_ID", "").strip() or "oc_cd9a6f072cdd348a08c29d09e8c9143a"
WEBHOOK = os.environ.get("FEISHU_WEBHOOK", "").strip()
committer = os.environ.get("COMMITTER", "团队") or "团队"
site = "https://heryma99.github.io/jwpei-logistics-compare/"
sha = os.environ.get("RATES_SHA", "")[:8]

def http_post(url, payload, headers=None, timeout=15):
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    if headers:
        for k, v in headers.items():
            req.add_header(k, v)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))

CARD = {
    "header": {"template": "blue", "title": {"tag": "plain_text", "content": "比价易 · 报价已更新"}},
    "elements": [
        {"tag": "div", "text": {"tag": "lark_md",
            "content": f"**更新人**：{committer}\n**状态**：rates.js 已自动烘焙并部署上线\n**提交**：`{sha}`"}},
        {"tag": "action", "actions": [
            {"tag": "button", "text": {"tag": "plain_text", "content": "查看比价易"}, "type": "primary", "url": site}
        ]}
    ]
}

# 路径一：应用 API（已在群里的「马金火的飞书 CLI」机器人）
if APP_ID and APP_SECRET and CHAT_ID:
    try:
        tok = http_post("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
                        {"app_id": APP_ID, "app_secret": APP_SECRET})
        if tok.get("code", 0) != 0:
            print("[notify] 获取 tenant_access_token 失败:", tok); sys.exit(1)
        t = tok["tenant_access_token"]
        payload = {"receive_id": CHAT_ID, "msg_type": "interactive", "content": json.dumps(CARD)}
        resp = http_post(f"https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
                         payload, headers={"Authorization": "Bearer " + t})
        if resp.get("code", 0) != 0:
            print("[notify] 飞书返回错误:", resp); sys.exit(1)
        print("[notify] 飞书通知(应用API)发送成功 -> chat", CHAT_ID)
        sys.exit(0)
    except Exception as e:
        print("[notify] 应用API路径异常:", e)

# 路径二：兜底 webhook
if WEBHOOK:
    try:
        data = json.dumps({"msg_type": "interactive", "card": CARD}).encode("utf-8")
        req = urllib.request.Request(WEBHOOK, data=data, method="POST")
        req.add_header("Content-Type", "application/json")
        with urllib.request.urlopen(req, timeout=15) as r:
            resp = json.loads(r.read().decode("utf-8"))
        if resp.get("code", 0) != 0:
            print("[notify] 飞书 webhook 返回错误:", resp); sys.exit(1)
        print("[notify] 飞书通知(webhook)发送成功")
        sys.exit(0)
    except Exception as e:
        print("[notify] webhook 异常:", e); sys.exit(1)

print("[notify] 未配置 FEISHU_APP_SECRET / FEISHU_WEBHOOK，跳过飞书通知")
sys.exit(0)

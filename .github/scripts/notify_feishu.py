#!/usr/bin/env python3
# 飞书群自定义机器人 webhook 通知（GitHub Action 用，纯标准库）
import os, json, sys, urllib.request, urllib.error

URL = os.environ.get("FEISHU_WEBHOOK", "").strip()
if not URL:
    print("[notify] FEISHU_WEBHOOK 未配置，跳过飞书通知")
    sys.exit(0)

committer = os.environ.get("COMMITTER", "团队")
site = "https://heryma99.github.io/jwpei-logistics-compare/"
sha = os.environ.get("RATES_SHA", "")[:8]

card = {
    "msg_type": "interactive",
    "card": {
        "header": {"template": "blue", "title": {"tag": "plain_text", "content": "比价易 · 报价已更新"}},
        "elements": [
            {"tag": "div", "text": {"tag": "lark_md",
                "content": f"**更新人**：{committer}\n**状态**：rates.js 已自动烘焙并部署上线\n**提交**：`{sha}`"}},
            {"tag": "action", "actions": [
                {"tag": "button", "text": {"tag": "plain_text", "content": "查看比价易"}, "type": "primary", "url": site}
            ]}
        ]
    }
}

try:
    data = json.dumps(card).encode("utf-8")
    req = urllib.request.Request(URL, data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, timeout=15) as r:
        resp = json.loads(r.read().decode("utf-8"))
    if resp.get("code", 0) != 0:
        print(f"[notify] 飞书返回错误: {resp}")
        sys.exit(1)
    print("[notify] 飞书通知发送成功")
except urllib.error.HTTPError as e:
    print(f"[notify] HTTP 错误: {e.code} {e.read().decode('utf-8','replace')[:200]}")
    sys.exit(1)
except Exception as e:
    print(f"[notify] 异常: {e}")
    sys.exit(1)

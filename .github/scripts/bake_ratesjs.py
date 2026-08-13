# -*- coding: utf-8 -*-
"""配置中心改 rates.json 后，由 GitHub Action 自动烘焙 rates.js + bump version.json。
保留 rates.json 原样（含 meta.banner），只重新生成 rates.js。
"""
import json, time, os
HERE = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(os.path.dirname(HERE))   # .github/scripts -> repo root
RATES = os.path.join(REPO_ROOT, "rates.json")
RATES_JS = os.path.join(REPO_ROOT, "rates.js")
VERSION = os.path.join(REPO_ROOT, "version.json")

with open(RATES, encoding="utf-8") as f:
    d = json.load(f)

BANNER = """
;(function(){
  function render(){
    var R = window.RATES || {};
    var g = R.generated || (R.meta && R.meta.effective_date);
    var note = (R.meta && R.meta.banner) || (R.meta && R.meta.note) || "";
    var h = document.querySelector("header");
    if(!h || !g) return;
    var el = document.getElementById("effBanner");
    if(!el){
      el = document.createElement("div");
      el.id = "effBanner";
      el.style.cssText = "margin-top:8px;padding:6px 12px;border-radius:8px;font-size:13px;font-weight:600;color:#0e1117;background:linear-gradient(90deg,#ffd479,#ffb347);box-shadow:0 1px 4px rgba(0,0,0,.25);display:inline-block";
      h.appendChild(el);
    }
    el.textContent = "\U0001F4C5 " + note;
  }
  if(document.readyState !== "loading") render();
  else document.addEventListener("DOMContentLoaded", render);
})();
"""

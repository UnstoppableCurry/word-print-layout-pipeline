# -*- coding: utf-8 -*-
"""管线 A+B test 服务：上传 Word ->
  管线A: LibreOffice headless 转 PDF -> pdfplumber 解析文字+坐标
  管线B: Win 真实 Word 打印到 XPS 虚拟打印机 -> 文字+坐标+页面图
  前端逐页并排渲染 + 逐行文字/坐标对照。
端口 8899。转换引擎版本固化在 systemd 环境注释里，禁止随意升级。
"""
import json
import os
import re
import shutil
import subprocess
import uuid
from pathlib import Path

import pdfplumber
import requests
from flask import Flask, jsonify, request, send_file

BASE = Path("/opt/word2pdf-test/data")
BASE.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

ALLOWED = {".doc", ".docx"}

# 行聚类容差: 来自 win 双管线实测结论(填空文字基线漂移 2.1pt), 必须 >=3pt
Y_TOL = 3.5

LO_PY = "/opt/libreoffice7.6/program/python"
FIXER = "/opt/word2pdf-test/uno_linewrap_fix.py"

# 管线 B：Win 打印解析节点地址（用环境变量覆盖，默认仅示例）
XPS_API = os.environ.get("XPS_API", "http://127.0.0.1:8090")
XPS_SCALE = 0.75  # XPS 坐标 96dpi -> PDF pt 72dpi


def convert_to_pdf(src: Path, outdir: Path, job: str) -> Path:
    """先走 UNO 修复管线(刀锋文档长数字串显式断行), 失败回退纯转换"""
    pdf = outdir / (src.stem + ".pdf")
    fix_pdf = outdir / (src.stem + "_fix.pdf")
    try:
        subprocess.run(
            [LO_PY, FIXER, src.resolve().as_uri(), fix_pdf.resolve().as_uri()],
            check=True, timeout=240, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if fix_pdf.exists() and fix_pdf.stat().st_size > 1000:
            fix_pdf.rename(pdf)
            return pdf
        raise RuntimeError("修复管线未产出 PDF")
    except Exception as e:
        app.logger.warning("UNO 修复管线失败, 回退纯转换: %s", e)
    profile = Path(f"/tmp/lo_profile_{job}")
    cmd = [
        os.environ.get("SOFFICE_BIN", "soffice"), "--headless", "--norestore", "--nolockcheck",
        f"-env:UserInstallation=file://{profile}",
        "--convert-to", "pdf:writer_pdf_Export",
        "--outdir", str(outdir), str(src),
    ]
    subprocess.run(cmd, check=True, timeout=180,
                   stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if not pdf.exists():
        raise RuntimeError("LibreOffice 未产出 PDF")
    shutil.rmtree(profile, ignore_errors=True)
    return pdf


def parse_pdf(pdf_path: Path):
    pages = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for pi, page in enumerate(pdf.pages, 1):
            words = page.extract_words(x_tolerance=1.5, y_tolerance=Y_TOL)
            clusters = []
            for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
                for c in clusters:
                    if abs(c["top"] - w["top"]) <= Y_TOL:
                        c["words"].append(w)
                        c["top"] = min(c["top"], w["top"])
                        break
                else:
                    clusters.append({"top": w["top"], "words": [w]})
            lines = []
            for c in sorted(clusters, key=lambda c: c["top"]):
                ws = sorted(c["words"], key=lambda w: w["x0"])
                lines.append({
                    "text": "".join(w["text"] for w in ws),
                    "x0": round(min(w["x0"] for w in ws), 1),
                    "top": round(c["top"], 1),
                    "x1": round(max(w["x1"] for w in ws), 1),
                    "bottom": round(max(w["bottom"] for w in ws), 1),
                })
            pages.append({"page": pi, "width": round(page.width, 1),
                          "height": round(page.height, 1), "lines": lines})
    return pages


def call_xps_service(src: Path):
    """把 Word 文件发给 Win 打印解析服务，返回其 JSON。
    附带把每行坐标乘 0.75 换算到 PDF pt 空间，并过滤纯空白行（另保留 all_count）。"""
    with open(src, "rb") as fh:
        r = requests.post(XPS_API + "/api/print-parse",
                          files={"file": (src.name, fh, "application/octet-stream")},
                          timeout=300)
    r.raise_for_status()
    d = r.json()
    if "pages" not in d:
        raise RuntimeError("XPS 节点返回错误: " + json.dumps(d, ensure_ascii=False)[:300])
    pages = []
    for p in d["pages"]:
        lines = [{
            "text": l["text"],
            "x0": round(l["x0"] * XPS_SCALE, 1),
            "top": round(l["top"] * XPS_SCALE, 1),
            "x1": round(l["x1"] * XPS_SCALE, 1),
            "bottom": round(l["bottom"] * XPS_SCALE, 1),
        } for l in p["lines"]]
        nonblank = [l for l in lines if l["text"].strip()]
        img = p.get("image_url") or ""
        m = re.match(r"^/images/([A-Za-z0-9]+)/([A-Za-z0-9_.]+)$", img)
        pages.append({
            "page": p["page"],
            "width": round(p["width"] * XPS_SCALE, 1),
            "height": round(p["height"] * XPS_SCALE, 1),
            "lines": nonblank,
            "raw_line_count": len(lines),
            "image_url": ("/xpsimg/%s/%s" % (m.group(1), m.group(2))) if m else "",
        })
    return {"engine": d.get("engine"), "job": d.get("job"),
            "page_count": d.get("page_count"), "pages": pages}


PAGE = """<!doctype html>
<html lang="zh"><head><meta charset="utf-8"><title>word2pdf 双管线对比 test 服务</title>
<style>
 body{font-family:'Microsoft YaHei',sans-serif;margin:24px;background:#f5f6f8;color:#222}
 h1{font-size:20px} .card{background:#fff;border:1px solid #ddd;border-radius:8px;padding:16px;margin-bottom:20px}
 button{padding:8px 20px;font-size:14px;cursor:pointer}
 .pg{position:relative;background:#fff;border:1px solid #999;margin:8px 0;box-shadow:0 1px 4px #0002;flex:0 0 auto}
 .ln{position:absolute;border-bottom:1px dashed #e33;white-space:nowrap;color:#123}
 .ln:hover{background:#ffe9a8}
 .xps .ln{border-bottom-color:#2680c2;color:#0a3d62}
 table{border-collapse:collapse;font-size:12px} td,th{border:1px solid #ccc;padding:3px 8px}
 .meta{color:#666;font-size:13px}
 .ok{color:#080;font-weight:bold} .bad{color:#c00;font-weight:bold}
 .row{display:flex;gap:20px;align-items:flex-start;flex-wrap:wrap}
 .col h3{font-size:14px;margin:6px 0}
 .verdict{font-size:16px;margin:8px 0}
 img.pgimg{border:1px solid #999;box-shadow:0 1px 4px #0002;max-width:100%}
</style></head><body>
<h1>word2pdf 双管线对比 — A: LibreOffice 转 PDF &nbsp;|&nbsp; B: Win 真实 Word 打印 XPS 解析</h1>
<div class="card">
  <input type="file" id="f" accept=".doc,.docx">
  <button onclick="go()">上传并解析</button>
  <span id="st" class="meta"></span>
  <div id="pdfLink"></div>
</div>
<div id="sum"></div>
<div id="out"></div>
<script>
const S = 0.7;
function esc(t){return t.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}
function norm(t){return t.replace(/\s+/g,'')}
function renderPanel(p, cls){
  let h='<div class="pg '+cls+'" style="width:'+(p.width*S)+'px;height:'+(p.height*S)+'px">';
  for(const l of p.lines){
    const fs=Math.max(6,(l.bottom-l.top)*S*0.8);
    h+='<div class="ln" title="x0='+l.x0+' top='+l.top+' x1='+l.x1+' bottom='+l.bottom+'" style="left:'+(l.x0*S)+'px;top:'+(l.top*S)+'px;font-size:'+fs+'px">'+esc(l.text)+'</div>';
  }
  return h+'</div>';
}
async function go(){
  const f = document.getElementById('f').files[0];
  if(!f){alert('先选文件');return}
  document.getElementById('st').textContent=' 双管线转换解析中（B 管线走 Windows 打印，约 10-60 秒）…';
  document.getElementById('sum').innerHTML=''; document.getElementById('out').innerHTML=''; document.getElementById('pdfLink').innerHTML='';
  const fd = new FormData(); fd.append('file', f);
  let j;
  try{
    const r = await fetch('/api/convert', {method:'POST', body:fd});
    j = await r.json();
    if(!r.ok){document.getElementById('st').textContent=' 失败: '+(j.error||r.status);return}
  }catch(e){document.getElementById('st').textContent=' 请求失败: '+e;return}
  document.getElementById('st').textContent=' 完成';
  document.getElementById('pdfLink').innerHTML='<p><a href="'+j.pdf_url+'" target="_blank">⬇ 下载 A 管线(LO)产出的 PDF</a></p>';

  const lo = j.pages, xps = j.xps ? j.xps.pages : null;
  // ---- 汇总 ----
  let allEq = !!xps, sumRows='';
  const npg = Math.max(lo.length, xps?xps.length:0);
  for(let i=0;i<npg;i++){
    const a=lo[i], b=xps?xps[i]:null;
    const eq = a&&b && a.lines.length===b.lines.length;
    if(!eq) allEq=false;
    sumRows+='<tr><td>'+(i+1)+'</td><td>'+(a?a.lines.length:'-')+'</td><td>'+(b?b.lines.length:'-')+'</td><td class="'+(eq?'ok':'bad')+'">'+(eq?'✓':'✗')+'</td></tr>';
  }
  let sh='<div class="card"><div class="verdict">总体判定: ';
  if(!xps){ sh+='<span class="bad">B 管线(XPS)失败</span> <span class="meta">'+esc(j.xps_error||'')+'</span>'; }
  else { sh+= allEq?'<span class="ok">✓ 两管线逐页行数一致</span>':'<span class="bad">✗ 存在行数不一致的页（窜行/窜页风险）</span>';
    sh+=' <span class="meta">B 引擎: '+esc(j.xps.engine||'')+', A '+lo.length+' 页 / B '+j.xps.page_count+' 页</span>'; }
  sh+='</div><table><tr><th>页</th><th>A-LO 行数</th><th>B-XPS 行数(非空)</th><th>行数一致</th></tr>'+sumRows+'</table></div>';
  document.getElementById('sum').innerHTML=sh;

  // ---- 逐页并排 + 逐行对照 ----
  let html='';
  for(let i=0;i<npg;i++){
    const a=lo[i], b=xps?xps[i]:null;
    html+='<div class="card"><b>第 '+(i+1)+' 页</b> <span class="meta">A '+(a?a.lines.length:'-')+' 行 | B '+(b?b.lines.length+' 行(原始 '+b.raw_line_count+' 含空行)':'-')+'</span>';
    html+='<div class="row">';
    if(a) html+='<div class="col"><h3>A — LibreOffice PDF ('+a.width+'×'+a.height+' pt)</h3>'+renderPanel(a,'lo')+'</div>';
    if(b) html+='<div class="col"><h3>B — Word 打印 XPS ('+b.width+'×'+b.height+' pt)</h3>'+renderPanel(b,'xps')+(b.image_url?'<details><summary>页面渲染图</summary><img class="pgimg" style="width:'+(b.width*S)+'px" src="'+b.image_url+'"></details>':'')+'</div>';
    html+='</div>';
    // 逐行对照表
    if(a&&b){
      const n=Math.max(a.lines.length,b.lines.length);
      html+='<details open><summary>逐行对照（文字归一化比较，坐标单位 pt）</summary><table><tr><th>#</th><th>A-LO 文字</th><th>B-XPS 文字</th><th>文字</th><th>Δtop</th><th>Δx0</th></tr>';
      for(let k=0;k<n;k++){
        const la=a.lines[k], lb=b.lines[k];
        let cls='ok', txt='✓', dt='', dx='';
        if(!la||!lb){cls='bad';txt='✗ 缺行';}
        else{
          const teq = norm(la.text)===norm(lb.text);
          const d1=Math.abs(la.top-lb.top).toFixed(1), d2=Math.abs(la.x0-lb.x0).toFixed(1);
          dt=d1; dx=d2;
          if(!teq){cls='bad';txt='✗';}
          else if(parseFloat(d1)>2.5||parseFloat(d2)>3.5){cls='bad';txt='≈ 坐标漂移';}
        }
        html+='<tr class="'+cls+'"><td>'+(k+1)+'</td><td>'+(la?esc(la.text):'')+'</td><td>'+(lb?esc(lb.text):'')+'</td><td>'+txt+'</td><td>'+dt+'</td><td>'+dx+'</td></tr>';
      }
      html+='</table></details>';
    }
    html+='</div>';
  }
  document.getElementById('out').innerHTML=html;
}
</script></body></html>"""


@app.get("/")
def index():
    return PAGE


@app.get("/health")
def health():
    return jsonify(ok=True)


@app.post("/api/convert")
def api_convert():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(error="缺少文件"), 400
    ext = Path(f.filename).suffix.lower()
    if ext not in ALLOWED:
        return jsonify(error=f"只支持 {sorted(ALLOWED)}"), 400
    job = uuid.uuid4().hex[:12]
    jobdir = BASE / job
    jobdir.mkdir(parents=True, exist_ok=True)
    src = jobdir / ("src" + ext)
    f.save(src)
    try:
        pdf = convert_to_pdf(src, jobdir, job)
        pages = parse_pdf(pdf)
    except Exception as e:  # noqa: BLE001
        return jsonify(error=str(e)), 500
    resp = {"job": job, "pdf_url": f"/files/{job}/{pdf.name}", "pages": pages}
    try:
        resp["xps"] = call_xps_service(src)
    except Exception as e:  # noqa: BLE001
        app.logger.warning("XPS 管线失败: %s", e)
        resp["xps"] = None
        resp["xps_error"] = str(e)
    return jsonify(resp)


@app.get("/xpsimg/<job>/<name>")
def xpsimg(job, name):
    if not re.fullmatch(r"[A-Za-z0-9]+", job) or not re.fullmatch(r"[A-Za-z0-9_.]+", name):
        return "bad path", 400
    try:
        r = requests.get(f"{XPS_API}/images/{job}/{name}", timeout=30)
        if r.status_code != 200:
            return "not found", 404
        return app.response_class(r.content, mimetype="image/png")
    except Exception as e:  # noqa: BLE001
        return str(e), 502


@app.get("/files/<job>/<name>")
def files(job, name):
    p = (BASE / job / name).resolve()
    if not str(p).startswith(str(BASE.resolve())) or not p.exists():
        return "not found", 404
    return send_file(p)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8899, threaded=True)

from __future__ import annotations

import cgi
import base64
import html
import json
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile
from collections import defaultdict
from datetime import datetime
from difflib import SequenceMatcher
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from features.auth.pages import admin_page, login_page
from features.auth.service import AuthManager

BASE_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = BASE_DIR / "scripts"
RUNS_DIR = BASE_DIR / "web_runs"
MASTER_DIR = BASE_DIR / "input" / "master"
DATA_DIR = BASE_DIR / "data"
MONTHLY_SALES_PATH = DATA_DIR / "월매출_납품상품.xlsx"
SALES_LEDGER_PATH = DATA_DIR / "_월매출_내부자료.xlsx"
YEAR_MANUAL_PATH = DATA_DIR / "연도총매출_수기입력.json"


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


load_env_file(BASE_DIR / ".env")
AUTH = AuthManager(DATA_DIR)


def clean_supabase_url(value: str) -> str:
    url = value.strip().rstrip("/")
    if url and not url.startswith(("http://", "https://")):
        url = f"https://{url}"
    return url


SUPABASE_URL = clean_supabase_url(os.environ.get("SUPABASE_URL", ""))
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()


def supabase_request(method: str, path: str, body: object | None = None) -> object | None:
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        return None
    payload = json.dumps(body).encode("utf-8") if body is not None else None
    request = urllib.request.Request(f"{SUPABASE_URL}{path}", data=payload, method=method)
    request.add_header("apikey", SUPABASE_SERVICE_ROLE_KEY)
    request.add_header("Authorization", f"Bearer {SUPABASE_SERVICE_ROLE_KEY}")
    request.add_header("Content-Type", "application/json")
    request.add_header("Prefer", "resolution=merge-duplicates,return=representation")
    try:
        with urllib.request.urlopen(request, timeout=15) as response:
            text = response.read().decode("utf-8")
            return json.loads(text) if text else None
    except Exception as exc:
        print(f"[supabase-files] {method} {path} failed: {exc}", flush=True)
        return None


def restore_file_from_supabase(file_key: str, target_path: Path) -> None:
    if target_path.exists():
        return
    rows = supabase_request(
        "GET",
        f"/rest/v1/app_files?select=content_base64&file_key=eq.{urllib.parse.quote(file_key)}",
    )
    if not rows:
        return
    content = rows[0].get("content_base64") if isinstance(rows, list) else None
    if not content:
        return
    target_path.parent.mkdir(parents=True, exist_ok=True)
    target_path.write_bytes(base64.b64decode(content))


def backup_file_to_supabase(file_key: str, source_path: Path) -> None:
    if not source_path.exists():
        return
    body = {
        "file_key": file_key,
        "content_base64": base64.b64encode(source_path.read_bytes()).decode("ascii"),
    }
    supabase_request("POST", "/rest/v1/app_files?on_conflict=file_key", body)


def restore_sales_files_from_supabase() -> None:
    restore_file_from_supabase("sales_ledger", SALES_LEDGER_PATH)
    restore_file_from_supabase("monthly_sales", MONTHLY_SALES_PATH)


def backup_sales_files_to_supabase() -> None:
    backup_file_to_supabase("sales_ledger", SALES_LEDGER_PATH)
    backup_file_to_supabase("monthly_sales", MONTHLY_SALES_PATH)


def restore_master_file_from_supabase() -> None:
    if get_local_master_path() is not None:
        return
    restore_file_from_supabase("master_file", MASTER_DIR / "기초자료.xlsx")


def backup_master_file_to_supabase(source_path: Path) -> None:
    backup_file_to_supabase("master_file", source_path)

sys.path.insert(0, str(SCRIPTS_DIR))
from po_automation import create_processed_po, read_master, read_po_lines, write_summary_workbook  # noqa: E402


HTML_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>보니애가구 쿠팡 PO 변환</title>
  <style>
    :root {
      --bg: #f3f6fb;
      --panel: #ffffff;
      --ink: #172033;
      --muted: #667085;
      --line: #d9e1ee;
      --brand: #1f4e79;
      --brand2: #2b7a78;
      --danger: #b42318;
      --ok: #027a48;
      --shadow: 0 18px 45px rgba(16, 24, 40, .10);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Malgun Gothic", "맑은 고딕", Arial, sans-serif;
      color: var(--ink);
      background: var(--bg);
    }
    .app {
      min-height: 100vh;
      display: grid;
      grid-template-columns: 250px 1fr;
    }
    .side {
      background: #102a43;
      color: white;
      padding: 24px 18px;
    }
    .logo {
      display: flex;
      align-items: center;
      gap: 12px;
      font-weight: 800;
      font-size: 20px;
      margin-bottom: 28px;
    }
    .logo-mark {
      width: 34px;
      height: 34px;
      border-radius: 8px;
      background: #2b7a78;
      display: grid;
      place-items: center;
      font-weight: 900;
    }
    .nav-item {
      padding: 12px 13px;
      border-radius: 8px;
      color: #d9e7f3;
      margin-bottom: 6px;
      font-size: 14px;
      cursor: pointer;
      user-select: none;
    }
    .nav-item.active {
      background: rgba(255,255,255,.13);
      color: #fff;
      font-weight: 700;
    }
    .main {
      padding: 28px;
    }
    .topbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 22px;
    }
    h1 {
      font-size: 24px;
      margin: 0;
      letter-spacing: 0;
    }
    .sub {
      margin-top: 7px;
      color: var(--muted);
      font-size: 14px;
    }
    .grid {
      display: grid;
      grid-template-columns: minmax(420px, 1.1fr) minmax(320px, .9fr);
      gap: 18px;
      align-items: start;
    }
    .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }
    .panel-head {
      padding: 18px 20px;
      border-bottom: 1px solid var(--line);
      font-weight: 800;
    }
    .panel-body { padding: 20px; }
    label {
      display: block;
      font-weight: 700;
      margin-bottom: 8px;
      font-size: 14px;
    }
    input[type=file] {
      width: 100%;
      border: 1px dashed #9fb2c8;
      background: #f8fbff;
      border-radius: 8px;
      padding: 14px;
      margin-bottom: 18px;
    }
    .btn {
      border: 0;
      border-radius: 8px;
      padding: 13px 18px;
      background: var(--brand);
      color: white;
      font-weight: 800;
      cursor: pointer;
      width: 100%;
      font-size: 15px;
    }
    .btn:hover { background: #173c5e; }
    .note {
      background: #f7f9fc;
      border: 1px solid var(--line);
      padding: 12px 14px;
      border-radius: 8px;
      color: #344054;
      font-size: 13px;
      line-height: 1.55;
      margin-top: 12px;
    }
    .status {
      padding: 14px 16px;
      border-radius: 8px;
      font-size: 14px;
      line-height: 1.6;
      margin-bottom: 14px;
      border: 1px solid var(--line);
      background: #fff;
    }
    .status.ok { border-color: #abefc6; background: #ecfdf3; color: var(--ok); }
    .status.err { border-color: #fecdca; background: #fef3f2; color: var(--danger); }
    .download {
      display: inline-block;
      text-decoration: none;
      color: white;
      background: var(--brand2);
      padding: 12px 16px;
      border-radius: 8px;
      font-weight: 800;
      margin-top: 8px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    th, td {
      border-bottom: 1px solid var(--line);
      padding: 9px 8px;
      text-align: left;
    }
    th { color: #344054; background: #f8fafc; }
    @media (max-width: 880px) {
      .app { grid-template-columns: 1fr; }
      .side { display: none; }
      .main { padding: 18px; }
      .grid { grid-template-columns: 1fr; }
    }
  </style>
  <script>
    function showComingSoon(name) {
      alert(name + " 메뉴는 아직 준비 중입니다. 지금은 쿠팡 PO 변환 메뉴만 사용할 수 있습니다.");
    }
  </script>
</head>
<body>
  <div class="app">
    <aside class="side">
      <div class="logo"><div class="logo-mark">B</div><div>보니애가구<br><span style="font-size:13px;font-weight:500;color:#b8c8d9;">업무 시스템</span></div></div>
      <div class="nav-item" onclick="location.href='/master'">쿠팡 기초자료 관리</div>
      <div class="nav-item active" onclick="location.href='/'">PO변환</div>
      <div class="nav-item" onclick="location.href='/sales/folders'">월별납품관리</div>
      <div class="nav-item" onclick="location.href='/check'">수량검수</div>
      <div class="nav-item" onclick="location.href='/sales'">매출확인용</div>
      <div class="nav-item" onclick="location.href='/pallet'">파렛트/쉽먼트</div>
      <div class="nav-item" style="margin-top:28px;" onclick="showComingSoon('라벨 출력')">라벨 출력</div>
      <div class="nav-item" onclick="showComingSoon('배송완료 관리')">배송완료 관리</div>
    </aside>
    <main class="main">
      <div class="topbar">
        <div>
          <h1>쿠팡 PO 변환</h1>
          <div class="sub">기초자료는 저장해두고, 평소에는 PO만 올려 심플웍스 등록용 파일을 자동 생성합니다.</div>
        </div>
      </div>
      <div class="grid">
        <section class="panel {year_section_class}">
          <div class="panel-head">파일 업로드</div>
          <div class="panel-body">
            <form method="post" action="/master/upload" enctype="multipart/form-data">
              <label>기초자료 엑셀 <span style="color:#667085;font-weight:500;">(처음 또는 변경 시에만 선택)</span></label>
              <input type="file" name="master" accept=".xlsx" required />
              <button class="btn" type="submit">기초자료 저장</button>
            </form>
            <div style="height:18px"></div>
            <form method="post" action="/convert" enctype="multipart/form-data">
              <label>쿠팡 PO 엑셀 파일</label>
              <input type="file" name="po_files" accept=".xlsx" multiple required />
              <button class="btn" type="submit">심플웍스 등록용 만들기</button>
            </form>
            <div class="note">
              {master_status}
              <br>PO는 한 개도 되고 여러 개도 됩니다. 결과는 ZIP 파일로 내려받습니다.
              ZIP은 여러 파일을 하나로 묶은 압축 파일입니다.
            </div>
          </div>
        </section>
        <section class="panel {upload_section_class}">
          <div class="panel-head">처리 기준</div>
          <div class="panel-body">
            <table>
              <tr><th>항목</th><th>기준</th></tr>
              <tr><td>스큐 아이디</td><td>PO의 상품코드와 기초자료의 SKU ID 연결</td></tr>
              <tr><td>납품불가</td><td>행 안에 납품불가 표시가 있으면 제외</td></tr>
              <tr><td>요청</td><td>요청 수량을 넘지 않게 PO 수량 조정</td></tr>
              <tr><td>결과</td><td>심플웍스 업로드용 PO 복사본</td></tr>
            </table>
          </div>
        </section>
      </div>
      {message}
    </main>
  </div>
</body>
</html>"""


PALLET_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>보니애가구 파렛트/쉽먼트</title>
  <style>
    :root {
      --bg: #f3f6fb;
      --panel: #ffffff;
      --ink: #172033;
      --muted: #667085;
      --line: #d9e1ee;
      --brand: #1f4e79;
      --brand2: #2b7a78;
      --danger: #b42318;
      --ok: #027a48;
      --shadow: 0 18px 45px rgba(16, 24, 40, .10);
    }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Malgun Gothic", "맑은 고딕", Arial, sans-serif; color: var(--ink); background: var(--bg); }
    .app { min-height: 100vh; display: grid; grid-template-columns: 250px minmax(0, 1fr); }
    .side { background: #102a43; color: white; padding: 24px 18px; }
    .logo { display: flex; align-items: center; gap: 12px; font-weight: 800; font-size: 20px; margin-bottom: 28px; }
    .logo-mark { width: 34px; height: 34px; border-radius: 8px; background: #2b7a78; display: grid; place-items: center; font-weight: 900; }
    .nav-item { padding: 12px 13px; border-radius: 8px; color: #d9e7f3; margin-bottom: 6px; font-size: 14px; cursor: pointer; user-select: none; }
    .nav-item.active { background: rgba(255,255,255,.13); color: #fff; font-weight: 700; }
    .main { padding: 28px; min-width: 0; }
    h1 { font-size: 24px; margin: 0; letter-spacing: 0; }
    .sub { margin-top: 7px; color: var(--muted); font-size: 14px; }
    .grid { display: grid; grid-template-columns: minmax(420px, 1fr) minmax(300px, .7fr); gap: 18px; align-items: start; margin-top: 22px; }
    .panel { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; box-shadow: var(--shadow); overflow: hidden; }
    .panel-head { padding: 18px 20px; border-bottom: 1px solid var(--line); font-weight: 800; }
    .panel-body { padding: 20px; }
    label { display: block; font-weight: 700; margin-bottom: 8px; font-size: 14px; }
    input[type=file] { width: 100%; border: 1px dashed #9fb2c8; background: #f8fbff; border-radius: 8px; padding: 14px; margin-bottom: 18px; }
    .btn { border: 0; border-radius: 8px; padding: 13px 18px; background: var(--brand); color: white; font-weight: 800; cursor: pointer; width: 100%; font-size: 15px; }
    .btn:hover { background: #173c5e; }
    .note { background: #f7f9fc; border: 1px solid var(--line); padding: 12px 14px; border-radius: 8px; color: #344054; font-size: 13px; line-height: 1.55; margin-top: 12px; }
    .status { padding: 14px 16px; border-radius: 8px; font-size: 14px; line-height: 1.6; margin: 18px 0 0; border: 1px solid var(--line); background: #fff; }
    .status.ok { border-color: #abefc6; background: #ecfdf3; color: var(--ok); }
    .status.err { border-color: #fecdca; background: #fef3f2; color: var(--danger); }
    .download { display: inline-block; text-decoration: none; color: white; background: var(--brand2); padding: 12px 16px; border-radius: 8px; font-weight: 800; margin-top: 8px; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { border-bottom: 1px solid var(--line); padding: 9px 8px; text-align: left; vertical-align: top; }
    th { color: #344054; background: #f8fafc; width: 120px; }
    @media (max-width: 880px) {
      .app { grid-template-columns: 1fr; }
      .side { display: none; }
      .main { padding: 18px; }
      .grid { grid-template-columns: 1fr; }
    }
  </style>
  <script>
    function showComingSoon(name) {
      alert(name + " 메뉴는 아직 준비 중입니다.");
    }
  </script>
</head>
<body>
  <div class="app">
    <aside class="side">
      <div class="logo"><div class="logo-mark">B</div><div>보니애가구<br><span style="font-size:13px;font-weight:500;color:#b8c8d9;">업무 시스템</span></div></div>
      <div class="nav-item" onclick="location.href='/master'">쿠팡 기초자료 관리</div>
      <div class="nav-item" onclick="location.href='/'">PO변환</div>
      <div class="nav-item" onclick="location.href='/sales/folders'">월별납품관리</div>
      <div class="nav-item" onclick="location.href='/check'">수량검수</div>
      <div class="nav-item" onclick="location.href='/sales'">매출확인용</div>
      <div class="nav-item active" onclick="location.href='/pallet'">파렛트/쉽먼트</div>
      <div class="nav-item" style="margin-top:28px;" onclick="showComingSoon('라벨 출력')">라벨 출력</div>
      <div class="nav-item" onclick="showComingSoon('배송완료 관리')">배송완료 관리</div>
    </aside>
    <main class="main">
      <h1>파렛트/쉽먼트</h1>
      <div class="sub">PO 여러 개를 올리면 상품 크기와 파렛트 기준으로 납품 묶음 초안을 만듭니다.</div>
      {message}
      <div class="grid">
        <section class="panel">
          <div class="panel-head">PO 업로드</div>
          <div class="panel-body">
            <form method="post" action="/pallet/create" enctype="multipart/form-data">
              <label>파렛트/쉽먼트 초안을 만들 PO 엑셀</label>
              <input type="file" name="po_files" multiple accept=".xlsx" />
              <button class="btn" type="submit">파렛트/쉽먼트 초안 만들기</button>
            </form>
            <div class="note">{master_status}</div>
          </div>
        </section>
        <section class="panel">
          <div class="panel-head">결과 엑셀에 들어가는 내용</div>
          <div class="panel-body">
            <table>
              <tr><th>스큐 합산</th><td>같은 스큐 아이디는 수량을 합쳐 봅니다.</td></tr>
              <tr><th>PO별 상세</th><td>어떤 PO에서 온 상품인지 확인합니다.</td></tr>
              <tr><th>쉽먼트 초안</th><td>상품 크기, 박스 수량, 파렛트 수량으로 예상 파렛트 수를 봅니다.</td></tr>
            </table>
            <div class="note">
              `파렛트`는 상품을 쌓아 운반하는 받침 단위이고, `쉽먼트`는 쿠팡에 보낼 납품 묶음입니다.
              초안은 자동 계산 자료라서 최종 확정 전 사람이 한 번 확인하는 용도로 쓰면 됩니다.
            </div>
          </div>
        </section>
      </div>
    </main>
  </div>
</body>
</html>"""


CHECK_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>보니애가구 수량검수</title>
  <style>
    :root { --bg:#f3f6fb; --panel:#fff; --ink:#172033; --muted:#667085; --line:#d9e1ee; --brand:#1f4e79; --danger:#b42318; --ok:#027a48; --warn:#b54708; --shadow:0 18px 45px rgba(16,24,40,.10); }
    * { box-sizing:border-box; }
    body { margin:0; font-family:"Malgun Gothic","맑은 고딕",Arial,sans-serif; color:var(--ink); background:var(--bg); }
    .app { min-height:100vh; display:grid; grid-template-columns:250px minmax(0,1fr); }
    .side { background:#102a43; color:white; padding:24px 18px; }
    .logo { display:flex; align-items:center; gap:12px; font-weight:800; font-size:20px; margin-bottom:28px; }
    .logo-mark { width:34px; height:34px; border-radius:8px; background:#2b7a78; display:grid; place-items:center; font-weight:900; }
    .nav-item { padding:12px 13px; border-radius:8px; color:#d9e7f3; margin-bottom:6px; font-size:14px; cursor:pointer; user-select:none; }
    .nav-item.active { background:rgba(255,255,255,.13); color:#fff; font-weight:700; }
    .main { padding:28px; min-width:0; }
    h1 { font-size:24px; margin:0; letter-spacing:0; }
    .sub { margin-top:7px; color:var(--muted); font-size:14px; }
    .cards { display:grid; grid-template-columns:minmax(0,1fr); gap:18px; margin-top:22px; }
    .panel { background:var(--panel); border:1px solid var(--line); border-radius:8px; box-shadow:var(--shadow); overflow:hidden; }
    .panel-head { padding:16px 18px; border-bottom:1px solid var(--line); font-weight:800; }
    .panel-body { padding:18px; }
    .form-grid { display:grid; grid-template-columns:180px 180px minmax(260px,1fr) minmax(260px,1fr); gap:12px; align-items:end; }
    label { display:flex; flex-direction:column; gap:6px; font-size:13px; color:#344054; font-weight:800; }
    input[type=date], input[type=file] { width:100%; border:1px solid #b9c6d8; border-radius:8px; padding:10px 11px; font:inherit; background:white; }
    .btn { border:0; border-radius:8px; padding:12px 16px; background:var(--brand); color:white; font-weight:800; cursor:pointer; font-size:14px; }
    .status { padding:14px 16px; border-radius:8px; font-size:14px; line-height:1.6; border:1px solid var(--line); background:#fff; }
    .status.ok { border-color:#abefc6; background:#ecfdf3; color:var(--ok); }
    .status.err { border-color:#fecdca; background:#fef3f2; color:var(--danger); }
    .note { margin-top:12px; color:#667085; font-size:13px; line-height:1.6; }
    .summary { display:grid; grid-template-columns:repeat(4,minmax(120px,1fr)); gap:10px; }
    .summary div { border:1px solid var(--line); border-radius:8px; padding:12px; background:#fbfcfe; }
    .summary b { display:block; font-size:20px; margin-top:4px; }
    .scroll { overflow:auto; max-height:620px; }
    table { width:100%; min-width:980px; border-collapse:collapse; table-layout:fixed; font-size:13px; }
    th, td { border-right:1px solid var(--line); border-bottom:1px solid var(--line); padding:9px 10px; text-align:left; vertical-align:middle; overflow-wrap:anywhere; word-break:keep-all; }
    th:last-child, td:last-child { border-right:0; }
    th { background:#f8fafc; color:#344054; position:sticky; top:0; z-index:1; }
    .ok-text { color:var(--ok); font-weight:800; }
    .bad-text { color:var(--danger); font-weight:800; }
    .warn-text { color:var(--warn); font-weight:800; }
    @media (max-width:880px) { .app{grid-template-columns:1fr;} .side{display:none;} .main{padding:18px;} .form-grid,.summary{grid-template-columns:1fr;} }
  </style>
  <script>
    function showComingSoon(name) { alert(name + " 메뉴는 아직 준비 중입니다."); }
  </script>
</head>
<body>
  <div class="app">
    <aside class="side">
      <div class="logo"><div class="logo-mark">B</div><div>보니애가구<br><span style="font-size:13px;font-weight:500;color:#b8c8d9;">업무 시스템</span></div></div>
      <div class="nav-item" onclick="location.href='/master'">쿠팡 기초자료 관리</div>
      <div class="nav-item" onclick="location.href='/'">PO변환</div>
      <div class="nav-item" onclick="location.href='/sales/folders'">월별납품관리</div>
      <div class="nav-item active" onclick="location.href='/check'">수량검수</div>
      <div class="nav-item" onclick="location.href='/sales'">매출확인용</div>
      <div class="nav-item" onclick="location.href='/pallet'">파렛트/쉽먼트</div>
      <div class="nav-item" style="margin-top:28px;" onclick="showComingSoon('라벨 출력')">라벨 출력</div>
      <div class="nav-item" onclick="showComingSoon('배송완료 관리')">배송완료 관리</div>
    </aside>
    <main class="main">
      <h1>수량검수</h1>
      <div class="sub">쿠팡에 등록된 스큐 납품수량과 심플웍스 No 기준 수량이 같은지 확인합니다.</div>
      {message}
      <div class="cards">
        <section class="panel">
          <div class="panel-head">심플웍스 엑셀 업로드</div>
          <div class="panel-body">
            <form method="post" action="/check/run" enctype="multipart/form-data">
              <div class="form-grid">
                <label>시작일
                  <input type="date" name="date_from" value="{date_from}">
                </label>
                <label>종료일
                  <input type="date" name="date_to" value="{date_to}">
                </label>
                <label>심플웍스 엑셀
                  <input type="file" name="simpleworks_file" accept=".xlsx" multiple>
                </label>
                <label>심플웍스 캡쳐 이미지
                  <input type="file" name="simpleworks_image" accept=".png,.jpg,.jpeg,.bmp,.webp" multiple>
                </label>
              </div>
              <div style="margin-top:14px;"><button class="btn" type="submit">수량 검수하기</button></div>
            </form>
            <div class="note">심플웍스 엑셀 또는 캡쳐 이미지를 여러 개 올릴 수 있습니다. 각 파일의 상품명 앞 노란 숫자를 `심플웍스 No`로 읽고, 같은 번호는 수량을 합산합니다. 캡쳐 이미지는 OCR로 읽기 때문에 엑셀보다 정확도가 낮을 수 있습니다.</div>
          </div>
        </section>
        {result}
      </div>
    </main>
  </div>
</body>
</html>"""


SALES_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>보니애가구 매출확인용</title>
  <style>
    :root { --bg:#f3f6fb; --panel:#fff; --ink:#172033; --muted:#667085; --line:#d9e1ee; --brand:#1f4e79; --brand2:#2b7a78; --shadow:0 18px 45px rgba(16,24,40,.10); }
    * { box-sizing: border-box; }
    body { margin:0; font-family:"Malgun Gothic","맑은 고딕",Arial,sans-serif; color:var(--ink); background:var(--bg); }
    .app { min-height:100vh; display:grid; grid-template-columns:250px minmax(0, 1fr); }
    .side { background:#102a43; color:white; padding:24px 18px; }
    .logo { display:flex; align-items:center; gap:12px; font-weight:800; font-size:20px; margin-bottom:28px; }
    .logo-mark { width:34px; height:34px; border-radius:8px; background:#2b7a78; display:grid; place-items:center; font-weight:900; }
    .nav-item { padding:12px 13px; border-radius:8px; color:#d9e7f3; margin-bottom:6px; font-size:14px; cursor:pointer; user-select:none; }
    .nav-item.active { background:rgba(255,255,255,.13); color:#fff; font-weight:700; }
    .main { padding:28px; min-width:0; }
    h1 { font-size:24px; margin:0; letter-spacing:0; }
    .sub { margin-top:7px; color:var(--muted); font-size:14px; }
    .cards { display:grid; grid-template-columns:minmax(0, 1fr); gap:18px; margin-top:22px; min-width:0; }
    .upload-panel { order:1; }
    .summary-panel { order:2; }
    .detail-panel { order:3; }
    .year-panel { order:4; }
    .panel { background:var(--panel); border:1px solid var(--line); border-radius:8px; box-shadow:var(--shadow); overflow:hidden; min-width:0; }
    .panel-head { padding:16px 18px; border-bottom:1px solid var(--line); font-weight:800; display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap; }
    .panel-body { padding:18px; }
    input[type=file] { width:100%; border:1px dashed #9fb2c8; background:#f8fbff; border-radius:8px; padding:14px; margin-bottom:12px; }
    .btn { text-decoration:none; border:0; border-radius:8px; padding:10px 14px; background:var(--brand); color:white; font-weight:800; cursor:pointer; font-size:14px; }
    table { width:100%; border-collapse:collapse; table-layout:fixed; font-size:13px; }
    th, td { border-right:1px solid var(--line); border-bottom:1px solid var(--line); padding:9px 10px; text-align:left; vertical-align:middle; overflow-wrap:anywhere; word-break:keep-all; }
    th:last-child, td:last-child { border-right:0; }
    th { background:#f8fafc; color:#344054; }
    .summary-table th, .summary-table td { text-align:center; white-space:normal; }
    .summary-table th:first-child, .summary-table td:first-child { text-align:left; }
    .detail-table th:nth-child(1), .detail-table td:nth-child(1),
    .detail-table th:nth-child(3), .detail-table td:nth-child(3),
    .detail-table th:nth-child(4), .detail-table td:nth-child(4),
    .detail-table th:nth-child(5), .detail-table td:nth-child(5),
    .detail-table th:nth-child(6), .detail-table td:nth-child(6) { white-space:nowrap; }
    .detail-table td:nth-child(3), .detail-table td:nth-child(4), .detail-table td:nth-child(5) { text-align:right; }
    .qty-input, .memo-input { width:100%; border:1px solid #b9c6d8; border-radius:6px; padding:7px 8px; font:inherit; background:#fff; }
    .money-input { width:100%; border:1px solid #b9c6d8; border-radius:6px; padding:7px 8px; font:inherit; background:#fff; text-align:right; }
    .qty-input { text-align:right; min-width:64px; }
    .memo-input { min-width:120px; }
    .changed, .changed input { color:#c1121f; font-weight:800; }
    .save-row { display:flex; justify-content:flex-end; padding:12px 16px; border-bottom:1px solid var(--line); background:#fbfcfe; }
    .year-table th { text-align:center; }
    .year-table td { text-align:right; }
    .year-table td:first-child { text-align:left; font-weight:700; }
    .auto-cell { background:#f7fbff; font-weight:700; }
    .over-cell.negative { color:#c1121f; font-weight:800; }
    .lookup-row { display:grid; grid-template-columns:repeat(5, minmax(120px, 1fr)); gap:10px; padding:12px 16px; border-bottom:1px solid var(--line); background:#fbfcfe; }
    .lookup-row label { display:flex; flex-direction:column; gap:5px; color:#475467; font-size:12px; font-weight:700; }
    .lookup-row input, .lookup-row select { width:100%; border:1px solid #b9c6d8; border-radius:6px; padding:8px 9px; font:inherit; background:#fff; }
    .lookup-row .lookup-reset { align-self:end; border:1px solid #b9c6d8; background:#fff; color:#1f4e79; border-radius:6px; padding:8px 10px; font-weight:800; cursor:pointer; }
    .inline-delete-form { align-self:end; margin:0; }
    .inline-delete-form .delete-btn { width:100%; }
    .panel-actions { display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
    .toggle-btn { border:1px solid #b9c6d8; background:#fff; color:#1f4e79; border-radius:6px; padding:8px 10px; font-weight:800; cursor:pointer; font-size:13px; }
    .delete-btn { border:1px solid #fecdca; background:#fff5f5; color:#b42318; border-radius:6px; padding:7px 9px; font-weight:800; cursor:pointer; font-size:12px; }
    .collapsible-content.is-hidden { display:none; }
    .not-shown { display:none; }
    .month-folder th { background:#d7e8f7; color:#12385a; cursor:pointer; font-weight:900; }
    .month-hidden { display:none !important; }
    .resizable-table th { position:relative; }
    .col-resizer { position:absolute; top:0; right:-4px; width:8px; height:100%; cursor:col-resize; user-select:none; touch-action:none; z-index:2; }
    .col-resizer:hover, .col-resizer.active { background:rgba(31,78,121,.22); }
    .scroll { max-height:520px; overflow:auto; }
    @media (max-width:880px) { .app{grid-template-columns:1fr;} .side{display:none;} .main{padding:18px;} table{font-size:12px;} th,td{padding:8px 7px;} .lookup-row{grid-template-columns:1fr 1fr;} }
  </style>
  <script>
    function showComingSoon(name) { alert(name + " 메뉴는 아직 준비 중입니다."); }
    function initResizableTables() {
      document.querySelectorAll(".resizable-table").forEach(function(table) {
        var cols = table.querySelectorAll("col");
        var headers = table.querySelectorAll("thead th");
        headers.forEach(function(header, index) {
          if (header.querySelector(".col-resizer")) return;
          var handle = document.createElement("span");
          handle.className = "col-resizer";
          header.appendChild(handle);
          handle.addEventListener("mousedown", function(event) {
            event.preventDefault();
            handle.classList.add("active");
            var startX = event.clientX;
            var tableWidth = table.getBoundingClientRect().width;
            var targetIndex = index;
            var partnerIndex = index + 1;
            var direction = 1;
            if (index === headers.length - 1) {
              targetIndex = index;
              partnerIndex = index - 1;
              direction = -1;
            }
            var targetWidth = headers[targetIndex].getBoundingClientRect().width;
            var partnerWidth = headers[partnerIndex].getBoundingClientRect().width;
            function move(moveEvent) {
              var delta = moveEvent.clientX - startX;
              var adjustedDelta = delta * direction;
              var newTarget = Math.max(58, targetWidth + adjustedDelta);
              var newPartner = Math.max(58, partnerWidth - adjustedDelta);
              var changedTotal = newTarget + newPartner;
              var originalTotal = targetWidth + partnerWidth;
              if (changedTotal !== originalTotal) {
                if (newTarget === 58) newPartner = originalTotal - 58;
                if (newPartner === 58) newTarget = originalTotal - 58;
              }
              cols[targetIndex].style.width = (newTarget / tableWidth * 100) + "%";
              cols[partnerIndex].style.width = (newPartner / tableWidth * 100) + "%";
            }
            function stop() {
              handle.classList.remove("active");
              document.removeEventListener("mousemove", move);
              document.removeEventListener("mouseup", stop);
            }
            document.addEventListener("mousemove", move);
            document.addEventListener("mouseup", stop);
          });
        });
      });
    }
    function parseMoney(text) {
      var cleaned = String(text || "").replaceAll(",", "").replaceAll("원", "").trim();
      var value = Number(cleaned);
      return Number.isFinite(value) ? value : 0;
    }
    function money(value) {
      return Math.round(value).toLocaleString("ko-KR") + "원";
    }
    function recalcSalesScreen() {
      var dayTotals = {};
      var monthTotals = {};
      document.querySelectorAll(".detail-row").forEach(function(row) {
        var day = row.dataset.day;
        var month = row.dataset.month;
        var originalQty = Number(row.dataset.originalQty || "0");
        var qtyInput = row.querySelector(".qty-input");
        var qty = Number(qtyInput.value || "0");
        var unitPrice = Number(row.dataset.unitPrice || "0");
        var amount = qty * unitPrice;
        row.querySelector(".row-amount").textContent = money(amount);
        row.classList.toggle("changed", qty !== originalQty || !!row.querySelector(".memo-input").value.trim());
        qtyInput.classList.toggle("changed", qty !== originalQty);
        if (!dayTotals[day]) dayTotals[day] = { qty: 0, amount: 0 };
        dayTotals[day].qty += qty;
        dayTotals[day].amount += amount;
        if (!monthTotals[month]) monthTotals[month] = 0;
        monthTotals[month] += amount;
      });
      var totalQty = 0;
      var totalAmount = 0;
      Object.keys(dayTotals).forEach(function(day) {
        document.querySelectorAll('[data-day-total="' + day + '"]').forEach(function(cell) {
          cell.textContent = day + " 합계: 수량 " + dayTotals[day].qty.toLocaleString("ko-KR") + "개 / 금액 " + money(dayTotals[day].amount);
        });
        var summaryRow = document.querySelector('.summary-row[data-day="' + day + '"]');
        if (summaryRow) {
          var vat = Math.round(dayTotals[day].amount / 1.1);
          var budget = Math.round(vat * 0.035);
          summaryRow.querySelector(".summary-qty").textContent = dayTotals[day].qty.toLocaleString("ko-KR");
          summaryRow.querySelector(".summary-amount").textContent = money(dayTotals[day].amount);
          summaryRow.querySelector(".summary-vat").textContent = money(vat);
          summaryRow.querySelector(".summary-budget").textContent = money(budget);
        }
        totalQty += dayTotals[day].qty;
        totalAmount += dayTotals[day].amount;
      });
      var totalVat = Math.round(totalAmount / 1.1);
      var totalBudget = Math.round(totalVat * 0.035);
      var totalRow = document.querySelector(".summary-total-row");
      if (totalRow) {
        totalRow.querySelector(".summary-total-qty").textContent = totalQty.toLocaleString("ko-KR");
        totalRow.querySelector(".summary-total-amount").textContent = money(totalAmount);
        totalRow.querySelector(".summary-total-vat").textContent = money(totalVat);
        totalRow.querySelector(".summary-total-budget").textContent = money(totalBudget);
      }
      document.querySelectorAll(".year-row").forEach(function(row) {
        var monthText = row.querySelector("td:first-child").textContent.replace("월", "").padStart(2, "0");
        var year = new Date().getFullYear();
        var key = year + "-" + monthText;
        var monthAmount = monthTotals[key] || 0;
        var vat = Math.round(monthAmount / 1.1);
        var budget = Math.round(vat * 0.035);
        row.dataset.sales = String(vat);
        row.dataset.budget = String(budget);
        row.querySelector(".year-sales").textContent = money(vat);
        row.querySelector(".year-budget").textContent = money(budget);
      });
      recalcYearScreen();
      applyLookups();
    }
    function recalcYearScreen() {
      var totalSales = 0;
      var totalBudget = 0;
      var totalTester = 0;
      var totalAd = 0;
      var totalOver = 0;
      var totalSupport = 0;
      var totalDiscount = 0;
      var totalPartnerDiscount = 0;
      var totalExtraAd = 0;
      document.querySelectorAll(".year-row").forEach(function(row) {
        if (row.style.display === "none") return;
        var sales = Number(row.dataset.sales || "0");
        var budget = Number(row.dataset.budget || "0");
        var tester = parseMoney(row.querySelector(".tester-input").value);
        var ad = parseMoney(row.querySelector(".ad-input").value);
        var support = parseMoney(row.querySelector(".support-input").value);
        var discount = parseMoney(row.querySelector(".discount-input").value);
        var partnerDiscount = parseMoney(row.querySelector(".partner-discount-input").value);
        var extraAd = parseMoney(row.querySelector(".extra-ad-input").value);
        var over = budget - tester - ad;
        row.querySelector(".over-cell").textContent = money(over);
        row.querySelector(".over-cell").classList.toggle("negative", over < 0);
        totalSales += sales;
        totalBudget += budget;
        totalTester += tester;
        totalAd += ad;
        totalOver += over;
        totalSupport += support;
        totalDiscount += discount;
        totalPartnerDiscount += partnerDiscount;
        totalExtraAd += extraAd;
        row.querySelector(".support-display").textContent = money(support);
        row.querySelector(".discount-display").textContent = money(discount);
        row.querySelector(".partner-discount-display").textContent = money(partnerDiscount);
        row.querySelector(".extra-ad-display").textContent = money(extraAd);
      });
      var totalRow = document.querySelector(".year-total-row");
      if (totalRow) {
        totalRow.querySelector(".year-total-sales").textContent = money(totalSales);
        totalRow.querySelector(".year-total-budget").textContent = money(totalBudget);
        totalRow.querySelector(".year-total-tester").textContent = money(totalTester);
        totalRow.querySelector(".year-total-ad").textContent = money(totalAd);
        totalRow.querySelector(".year-total-over").textContent = money(totalOver);
        totalRow.querySelector(".year-total-support").textContent = money(totalSupport);
        totalRow.querySelector(".year-total-discount").textContent = money(totalDiscount);
        totalRow.querySelector(".year-total-partner-discount").textContent = money(totalPartnerDiscount);
        totalRow.querySelector(".year-total-extra-ad").textContent = money(totalExtraAd);
      }
    }
    function inDateRange(day, from, to) {
      if (from && day < from) return false;
      if (to && day > to) return false;
      return true;
    }
    function applyLookups() {
      var yearMonth = document.getElementById("year-month-lookup")?.value || "";
      document.querySelectorAll(".year-row").forEach(function(row) {
        row.style.display = !yearMonth || row.dataset.month === yearMonth ? "" : "none";
      });

      var summaryFrom = document.getElementById("summary-from")?.value || "";
      var summaryTo = document.getElementById("summary-to")?.value || "";
      var visibleSummary = { po: 0, qty: 0, amount: 0 };
      document.querySelectorAll(".summary-row").forEach(function(row) {
        var show = inDateRange(row.dataset.day, summaryFrom, summaryTo);
        row.style.display = show ? "" : "none";
        if (show) {
          visibleSummary.po += Number(row.dataset.poCount || "0");
          visibleSummary.qty += parseMoney(row.querySelector(".summary-qty").textContent);
          visibleSummary.amount += parseMoney(row.querySelector(".summary-amount").textContent);
        }
      });
      var summaryTotal = document.querySelector(".summary-total-row");
      if (summaryTotal) {
        var vat = Math.round(visibleSummary.amount / 1.1);
        var budget = Math.round(vat * 0.035);
        summaryTotal.querySelector(".summary-total-po").textContent = visibleSummary.po.toLocaleString("ko-KR");
        summaryTotal.querySelector(".summary-total-qty").textContent = visibleSummary.qty.toLocaleString("ko-KR");
        summaryTotal.querySelector(".summary-total-amount").textContent = money(visibleSummary.amount);
        summaryTotal.querySelector(".summary-total-vat").textContent = money(vat);
        summaryTotal.querySelector(".summary-total-budget").textContent = money(budget);
      }

      var detailFrom = document.getElementById("detail-from")?.value || "";
      var detailTo = document.getElementById("detail-to")?.value || "";
      var detailKeyword = (document.getElementById("detail-keyword")?.value || "").trim().toLowerCase();
      var visibleByDay = {};
      var visibleByMonth = {};
      document.querySelectorAll(".detail-row").forEach(function(row) {
        var haystack = (row.dataset.search || "").toLowerCase();
        var show = inDateRange(row.dataset.day, detailFrom, detailTo) && (!detailKeyword || haystack.includes(detailKeyword));
        row.style.display = show ? "" : "none";
        if (show) {
          visibleByDay[row.dataset.day] = true;
          visibleByMonth[row.dataset.month] = true;
        }
      });
      document.querySelectorAll(".detail-day-row").forEach(function(row) {
        row.style.display = visibleByDay[row.dataset.day] ? "" : "none";
      });
      document.querySelectorAll(".month-folder").forEach(function(row) {
        row.style.display = visibleByMonth[row.dataset.month] ? "" : "none";
      });
      recalcYearScreen();
    }
    function resetLookup(group) {
      document.querySelectorAll('[data-lookup-group="' + group + '"] input, [data-lookup-group="' + group + '"] select').forEach(function(input) {
        input.value = "";
      });
      applyLookups();
    }
    function confirmExactDelete(label, expected) {
      var typed = prompt(label + " 삭제를 진행하려면 아래 값을 그대로 입력하세요.\\n\\n" + expected);
      return typed === expected;
    }
    function confirmDeleteDay(day) {
      return confirmExactDelete(day + " 일자", day);
    }
    function confirmDeleteMonth(month) {
      return confirmExactDelete(month + " 월 전체", month);
    }
    function setLookupDeleteDay(form) {
      var from = document.getElementById("detail-from")?.value || "";
      var to = document.getElementById("detail-to")?.value || "";
      if (!from || !to || from !== to) {
        alert("일자삭제는 시작일과 종료일을 같은 날짜로 선택해야 합니다.");
        return false;
      }
      form.querySelector('input[name="delete_day"]').value = from;
      return confirmDeleteDay(from);
    }
    function setLookupDeleteMonth(form) {
      var from = document.getElementById("detail-from")?.value || "";
      var to = document.getElementById("detail-to")?.value || "";
      var month = form.querySelector('input[name="delete_month"]').value;
      if (from && to && from.slice(0, 7) !== to.slice(0, 7)) {
        alert("월삭제는 시작일과 종료일이 같은 월 안에 있어야 합니다.");
        return false;
      }
      if (from) month = from.slice(0, 7);
      else if (to) month = to.slice(0, 7);
      form.querySelector('input[name="delete_month"]').value = month;
      return confirmDeleteMonth(month);
    }
    function toggleSection(targetId, button) {
      var target = document.getElementById(targetId);
      if (!target) return;
      var hidden = target.classList.toggle("is-hidden");
      button.textContent = hidden ? "펼치기" : "숨기기";
    }
    function toggleMonthFolder(month, button) {
      var closed = button.closest(".month-folder").classList.toggle("is-closed");
      document.querySelectorAll('[data-month="' + month + '"]').forEach(function(row) {
        if (!row.classList.contains("month-folder")) row.classList.toggle("month-hidden", closed);
      });
      button.textContent = closed ? "펼치기" : "숨기기";
      applyLookups();
    }
    document.addEventListener("DOMContentLoaded", function() {
      initResizableTables();
      document.querySelectorAll(".qty-input, .memo-input").forEach(function(input) {
        input.addEventListener("input", recalcSalesScreen);
      });
      document.querySelectorAll(".money-input").forEach(function(input) {
        input.addEventListener("input", recalcYearScreen);
      });
      document.querySelectorAll(".lookup-row input, .lookup-row select").forEach(function(input) {
        input.addEventListener("input", applyLookups);
        input.addEventListener("change", applyLookups);
      });
      recalcSalesScreen();
      recalcYearScreen();
      applyLookups();
    });
  </script>
</head>
<body>
  <div class="app">
    <aside class="side">
      <div class="logo"><div class="logo-mark">B</div><div>보니애가구<br><span style="font-size:13px;font-weight:500;color:#b8c8d9;">업무 시스템</span></div></div>
      <div class="nav-item" onclick="location.href='/master'">쿠팡 기초자료 관리</div>
      <div class="nav-item" onclick="location.href='/'">PO변환</div>
      <div class="nav-item {folders_active}" onclick="location.href='/sales/folders'">월별납품관리</div>
      <div class="nav-item" onclick="location.href='/check'">수량검수</div>
      <div class="nav-item {sales_active}" onclick="location.href='/sales'">매출확인용</div>
      <div class="nav-item" onclick="location.href='/pallet'">파렛트/쉽먼트</div>
      <div class="nav-item" style="margin-top:28px;" onclick="showComingSoon('라벨 출력')">라벨 출력</div>
      <div class="nav-item" onclick="showComingSoon('배송완료 관리')">배송완료 관리</div>
    </aside>
    <main class="main">
      <h1>{page_title}</h1>
      <div class="sub">{page_sub}</div>
      {message}
      <div class="cards">
        <section class="panel year-panel {year_section_class}">
          <div class="panel-head">
            <span>연도총매출</span>
            <div class="panel-actions">
              <span style="font-size:12px;color:#667085;">매출/VAT 별도/광고비예산은 월매출에서 자동 반영</span>
              <button class="toggle-btn" type="button" onclick="toggleSection('year-sales-content', this)">숨기기</button>
            </div>
          </div>
          <div id="year-sales-content" class="collapsible-content">
          <div class="lookup-row" data-lookup-group="year">
            <label>월 조회
              <select id="year-month-lookup">
                <option value="">전체</option>
                <option value="01">1월</option><option value="02">2월</option><option value="03">3월</option>
                <option value="04">4월</option><option value="05">5월</option><option value="06">6월</option>
                <option value="07">7월</option><option value="08">8월</option><option value="09">9월</option>
                <option value="10">10월</option><option value="11">11월</option><option value="12">12월</option>
              </select>
            </label>
            <button class="lookup-reset" type="button" onclick="resetLookup('year')">조회 초기화</button>
          </div>
          <form method="post" action="/sales/year/save">
          <div class="save-row"><button class="btn" type="submit">연도총매출 저장</button></div>
          <div class="scroll" style="max-height:360px;">
            <table class="year-table resizable-table">
              <colgroup>
                <col style="width:7%;">
                <col style="width:11%;">
                <col style="width:11%;">
                <col style="width:10%;">
                <col style="width:10%;">
                <col style="width:11%;">
                <col style="width:10%;">
                <col style="width:10%;">
                <col style="width:10%;">
                <col style="width:10%;">
              </colgroup>
              <thead><tr><th>월</th><th>매출<br>(VAT 미포함)</th><th>광고비예산<br>(VAT 미포함)</th><th>체험단<br>(VAT 미포함)</th><th>광고</th><th>광고비초과</th><th>성장장려금<br>(분기)</th><th>즉시할인<br>(VAT 미포함)</th><th>즉시할인<br>(다연채)</th><th>광고비</th></tr></thead>
              <tbody>{year_rows}</tbody>
            </table>
          </div>
          </form>
          </div>
        </section>
        <section class="panel upload-panel {upload_section_class}">
          <div class="panel-head">납품 PO 업로드</div>
          <div class="panel-body">
            <form method="post" action="/sales/upload" enctype="multipart/form-data">
              <input type="file" name="sales_po_files" accept=".xlsx" multiple required>
              <button class="btn" type="submit">납품 상품 매출에 반영</button>
            </form>
            <div style="margin-top:10px;color:#667085;font-size:13px;line-height:1.6;">
              심플웍스 등록 전후로 직접 수정한 최종 PO 복사본을 여기에 올리면, 스큐 아이디 기준으로 월별 납품 내역과 매출에 누적됩니다.
            </div>
          </div>
        </section>
        <section class="panel summary-panel {summary_section_class}">
          <div class="panel-head">
            <span>일자별 합계</span>
            <div class="panel-actions">
              <a class="btn" href="/sales/download">월매출 엑셀 다운로드</a>
              <button class="toggle-btn" type="button" onclick="toggleSection('monthly-summary-content', this)">숨기기</button>
            </div>
          </div>
          <div id="monthly-summary-content" class="collapsible-content">
          <div class="lookup-row" data-lookup-group="summary">
            <label>시작일
              <input id="summary-from" type="date">
            </label>
            <label>종료일
              <input id="summary-to" type="date">
            </label>
            <button class="lookup-reset" type="button" onclick="resetLookup('summary')">조회 초기화</button>
          </div>
          <table class="summary-table resizable-table">
            <colgroup>
              <col style="width:15%;">
              <col style="width:10%;">
              <col style="width:13%;">
              <col style="width:22%;">
              <col style="width:20%;">
              <col style="width:20%;">
            </colgroup>
            <thead><tr><th>일자</th><th>PO 수</th><th>납품수량</th><th>납품상품 합계금액</th><th>VAT 별도</th><th>광고비예산</th></tr></thead>
            <tbody>{summary_rows}</tbody>
          </table>
          </div>
        </section>
        <section class="panel detail-panel">
          <div class="panel-head">{detail_title}</div>
          <div class="lookup-row" data-lookup-group="detail">
            <label>시작일
              <input id="detail-from" type="date">
            </label>
            <label>종료일
              <input id="detail-to" type="date">
            </label>
            <label>스큐/상품명/PO/메모
              <input id="detail-keyword" type="search" placeholder="찾을 내용 입력">
            </label>
            <button class="lookup-reset" type="button" onclick="resetLookup('detail')">조회 초기화</button>
            <form class="inline-delete-form" method="post" action="/sales/delete" onsubmit="return setLookupDeleteDay(this)">
              <input type="hidden" name="delete_day" value="">
              <button class="delete-btn" type="submit">조회 일자 삭제</button>
            </form>
            <form class="inline-delete-form" method="post" action="/sales/delete" onsubmit="return setLookupDeleteMonth(this)">
              <input type="hidden" name="delete_month" value="{current_month}">
              <button class="delete-btn" type="submit">조회 월 삭제</button>
            </form>
          </div>
          <form method="post" action="/sales/save">
          <div class="save-row"><button class="btn" type="submit">수량/메모 저장</button></div>
          <div class="scroll">
            <table class="detail-table resizable-table">
              <colgroup>
                <col style="width:10%;">
                <col style="width:37%;">
                <col style="width:8%;">
                <col style="width:9%;">
                <col style="width:10%;">
                <col style="width:8%;">
                <col style="width:13%;">
                <col style="width:5%;">
              </colgroup>
              <thead><tr><th>SKU ID</th><th>상품명</th><th>납품수량</th><th>단가</th><th>금액</th><th>비고</th><th>메모</th><th>삭제</th></tr></thead>
              <tbody>{detail_rows}</tbody>
            </table>
          </div>
          </form>
        </section>
      </div>
    </main>
  </div>
</body>
</html>"""


MASTER_PAGE = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>보니애가구 쿠팡 기초자료 관리</title>
  <style>
    :root {
      --bg: #f3f6fb;
      --panel: #ffffff;
      --ink: #172033;
      --muted: #667085;
      --line: #d9e1ee;
      --brand: #1f4e79;
      --brand2: #2b7a78;
      --danger: #b42318;
      --ok: #027a48;
      --shadow: 0 18px 45px rgba(16, 24, 40, .10);
    }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Malgun Gothic", "맑은 고딕", Arial, sans-serif; color: var(--ink); background: var(--bg); }
    .app { min-height: 100vh; display: grid; grid-template-columns: 250px minmax(0, 1fr); }
    .side { background: #102a43; color: white; padding: 24px 18px; }
    .logo { display: flex; align-items: center; gap: 12px; font-weight: 800; font-size: 20px; margin-bottom: 28px; }
    .logo-mark { width: 34px; height: 34px; border-radius: 8px; background: #2b7a78; display: grid; place-items: center; font-weight: 900; }
    .nav-item { padding: 12px 13px; border-radius: 8px; color: #d9e7f3; margin-bottom: 6px; font-size: 14px; cursor: pointer; user-select: none; }
    .nav-item.active { background: rgba(255,255,255,.13); color: #fff; font-weight: 700; }
    .main { padding: 28px; min-width: 0; }
    h1 { font-size: 24px; margin: 0; letter-spacing: 0; }
    .sub { margin-top: 7px; color: var(--muted); font-size: 14px; }
    .panel { margin-top: 22px; background: var(--panel); border: 1px solid var(--line); border-radius: 8px; box-shadow: var(--shadow); overflow: hidden; }
    .add-grid { display: grid; grid-template-columns: 130px minmax(260px, 1.4fr) repeat(7, minmax(95px, 1fr)) 110px; gap: 10px; padding: 16px 18px; border-bottom: 1px solid var(--line); background: #fbfdff; align-items: end; }
    .field label { display: block; font-size: 12px; font-weight: 800; color: #344054; margin-bottom: 6px; }
    .field input[type=text], .field input[type=number] { width: 100%; border: 1px solid var(--line); border-radius: 8px; padding: 10px 11px; font-size: 14px; background: white; }
    .check-field { display: flex; gap: 7px; align-items: center; height: 39px; font-size: 13px; font-weight: 700; }
    .toolbar { display: grid; grid-template-columns: minmax(260px, 460px) auto; gap: 10px; align-items: end; padding: 16px 18px; border-bottom: 1px solid var(--line); }
    .search { width: min(460px, 100%); border: 1px solid var(--line); border-radius: 8px; padding: 11px 12px; font-size: 14px; }
    .toolbar-sort { display: flex; flex-direction: column; gap: 5px; color: #475467; font-size: 12px; font-weight: 800; }
    .toolbar-sort select { width: 100%; border: 1px solid #b9c6d8; border-radius: 8px; padding: 9px 10px; font: inherit; background: white; }
    .toolbar-actions { display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; }
    .filter-row { display: grid; grid-template-columns: repeat(4, minmax(130px, 1fr)); gap: 10px; padding: 12px 18px; border-bottom: 1px solid var(--line); background: #fbfcfe; }
    .filter-row label { display: flex; flex-direction: column; gap: 5px; color: #475467; font-size: 12px; font-weight: 800; }
    .filter-row select { width: 100%; border: 1px solid #b9c6d8; border-radius: 8px; padding: 9px 10px; font: inherit; background: white; }
    .filter-buttons { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
    .filter-chip { border: 1px solid #b9c6d8; background: #fff; color: #1f4e79; border-radius: 8px; padding: 11px 14px; font-weight: 900; cursor: pointer; font-size: 14px; }
    .filter-chip.active { background: var(--brand); color: #fff; border-color: var(--brand); }
    .filter-count { color: #475467; font-size: 13px; font-weight: 800; padding: 8px 0; }
    .btn { border: 0; border-radius: 8px; padding: 11px 16px; background: var(--brand); color: white; font-weight: 800; cursor: pointer; font-size: 14px; }
    .btn.secondary { background: #475467; }
    .status { padding: 13px 16px; border-radius: 8px; font-size: 14px; line-height: 1.6; margin-top: 16px; border: 1px solid var(--line); background: #fff; }
    .status.ok { border-color: #abefc6; background: #ecfdf3; color: var(--ok); }
    .status.err { border-color: #fecdca; background: #fef3f2; color: var(--danger); }
    .table-wrap { max-height: calc(100vh - 230px); overflow: auto; }
    table { width: 100%; min-width: 900px; border-collapse: collapse; table-layout: fixed; font-size: 13px; }
    th, td { border-right: 1px solid var(--line); border-bottom: 1px solid var(--line); padding: 8px; text-align: left; vertical-align: middle; overflow-wrap: anywhere; word-break: keep-all; }
    th:last-child, td:last-child { border-right: 0; }
    th { position: sticky; top: 0; background: #f8fafc; z-index: 1; color: #344054; }
    td.sku { font-weight: 700; color: #1f4e79; white-space: nowrap; }
    input.qty { width: 100%; min-width: 72px; border: 1px solid var(--line); border-radius: 6px; padding: 7px 8px; text-align: right; }
    .product { min-width: 420px; }
    .resizable-table th { position: sticky; }
    .col-resizer { position:absolute; top:0; right:-4px; width:8px; height:100%; cursor:col-resize; user-select:none; touch-action:none; z-index:3; }
    .col-resizer:hover, .col-resizer.active { background:rgba(31,78,121,.22); }
    .muted { color: var(--muted); font-size: 12px; }
    @media (max-width: 880px) {
      .app { grid-template-columns: 1fr; }
      .side { display: none; }
      .main { padding: 18px; }
      .toolbar { grid-template-columns: 1fr; align-items: stretch; }
      .filter-row { grid-template-columns: 1fr 1fr; }
    }
  </style>
  <script>
    function numberValue(value) {
      const cleaned = String(value || '').replace(/[^0-9.-]/g, '');
      const parsed = Number(cleaned);
      return Number.isFinite(parsed) ? parsed : 0;
    }
    function rowText(tr, selector) {
      const cell = tr.querySelector(selector);
      return cell ? cell.textContent.trim() : '';
    }
    function rowInputNumber(tr, inputNamePrefix) {
      const input = tr.querySelector('input[name^="' + inputNamePrefix + '"]');
      if (!input || !String(input.value || '').trim()) return null;
      return numberValue(input.value);
    }
    function compareNullableNumber(a, b, direction) {
      const aBlank = a === null;
      const bBlank = b === null;
      if (aBlank && bBlank) return 0;
      if (aBlank) return 1;
      if (bBlank) return -1;
      const result = a - b;
      return direction === 'desc' ? -result : result;
    }
    function sortMasterRows(rows) {
      const sort = document.getElementById('sort-master') ? document.getElementById('sort-master').value : '';
      if (!sort) return rows;
      const sorted = rows.slice();
      sorted.sort((a, b) => {
        if (sort === 'name-asc' || sort === 'name-desc') {
          const result = rowText(a, '.product').localeCompare(rowText(b, '.product'), 'ko-KR', { numeric: true });
          return sort === 'name-desc' ? -result : result;
        }
        if (sort === 'amount-asc' || sort === 'amount-desc') {
          return compareNullableNumber(rowInputNumber(a, 'amount_'), rowInputNumber(b, 'amount_'), sort === 'amount-desc' ? 'desc' : 'asc');
        }
        if (sort === 'simple-desc' || sort === 'simple-asc') {
          return compareNullableNumber(rowInputNumber(a, 'simple_no_'), rowInputNumber(b, 'simple_no_'), sort === 'simple-desc' ? 'desc' : 'asc');
        }
        return 0;
      });
      return sorted;
    }
    const masterFilterStorageKey = 'bonie-master-filters';
    function saveMasterFilterState() {
      const state = {
        q: document.getElementById('q') ? document.getElementById('q').value : '',
        unavailable: document.getElementById('filter-unavailable') ? document.getElementById('filter-unavailable').value : '',
        simpleNo: document.getElementById('filter-simple-no') ? document.getElementById('filter-simple-no').value : '',
        amount: document.getElementById('filter-amount') ? document.getElementById('filter-amount').value : '',
        sort: document.getElementById('sort-master') ? document.getElementById('sort-master').value : ''
      };
      localStorage.setItem(masterFilterStorageKey, JSON.stringify(state));
    }
    function restoreMasterFilterState() {
      let state = null;
      try {
        state = JSON.parse(localStorage.getItem(masterFilterStorageKey) || 'null');
      } catch (error) {
        state = null;
      }
      if (!state) return;
      if (document.getElementById('q')) document.getElementById('q').value = state.q || '';
      if (document.getElementById('filter-unavailable')) document.getElementById('filter-unavailable').value = state.unavailable || '';
      if (document.getElementById('filter-simple-no')) document.getElementById('filter-simple-no').value = state.simpleNo || '';
      if (document.getElementById('filter-amount')) document.getElementById('filter-amount').value = state.amount || '';
      if (document.getElementById('sort-master')) document.getElementById('sort-master').value = state.sort || '';
    }
    function filterRows() {
      const q = document.getElementById('q').value.trim().toLowerCase();
      const unavailable = document.getElementById('filter-unavailable').value;
      const simpleNo = document.getElementById('filter-simple-no').value;
      const amount = document.getElementById('filter-amount').value;
      const tbody = document.querySelector('tbody');
      const rows = Array.from(tbody.querySelectorAll('tr'));
      sortMasterRows(rows).forEach(tr => tbody.appendChild(tr));
      rows.forEach(tr => {
        const matchText = tr.innerText.toLowerCase().includes(q);
        const matchUnavailable = !unavailable || tr.dataset.unavailable === unavailable;
        const hasSimpleNo = rowInputNumber(tr, 'simple_no_') !== null ? '1' : '0';
        const hasAmount = rowInputNumber(tr, 'amount_') !== null ? '1' : '0';
        const matchSimple = !simpleNo || hasSimpleNo === simpleNo;
        const matchAmount = !amount || hasAmount === amount;
        tr.style.display = matchText && matchUnavailable && matchSimple && matchAmount ? '' : 'none';
      });
      saveMasterFilterState();
      updateFilterCount();
    }
    function resetMasterFilters() {
      document.getElementById('q').value = '';
      document.getElementById('filter-unavailable').value = '';
      document.getElementById('filter-simple-no').value = '';
      document.getElementById('filter-amount').value = '';
      const sortSelect = document.getElementById('sort-master');
      if (sortSelect) sortSelect.value = '';
      localStorage.removeItem(masterFilterStorageKey);
      updateFilterChips();
      filterRows();
    }
    function setMasterFilter(filterId, value) {
      document.getElementById(filterId).value = value;
      updateFilterChips();
      filterRows();
    }
    function updateFilterChips() {
      document.querySelectorAll('.filter-chip[data-filter-id]').forEach(btn => {
        const select = document.getElementById(btn.dataset.filterId);
        btn.classList.toggle('active', select && select.value === btn.dataset.value);
      });
    }
    function updateFilterCount() {
      const rows = Array.from(document.querySelectorAll('tbody tr'));
      const visible = rows.filter(tr => tr.style.display !== 'none').length;
      const counter = document.getElementById('filter-count');
      if (counter) counter.textContent = '현재 ' + visible.toLocaleString('ko-KR') + '개 표시 / 전체 ' + rows.length.toLocaleString('ko-KR') + '개';
    }
    function initResizableTables() {
      document.querySelectorAll(".resizable-table").forEach(function(table) {
        var cols = table.querySelectorAll("col");
        var headers = table.querySelectorAll("thead th");
        headers.forEach(function(header, index) {
          if (header.querySelector(".col-resizer")) return;
          var handle = document.createElement("span");
          handle.className = "col-resizer";
          header.appendChild(handle);
          handle.addEventListener("mousedown", function(event) {
            event.preventDefault();
            handle.classList.add("active");
            var startX = event.clientX;
            var startWidth = cols[index].getBoundingClientRect().width;
            function move(moveEvent) {
              var nextWidth = Math.max(70, startWidth + moveEvent.clientX - startX);
              cols[index].style.width = nextWidth + "px";
            }
            function stop() {
              handle.classList.remove("active");
              document.removeEventListener("mousemove", move);
              document.removeEventListener("mouseup", stop);
            }
            document.addEventListener("mousemove", move);
            document.addEventListener("mouseup", stop);
          });
        });
      });
    }
    window.addEventListener("load", initResizableTables);
    window.addEventListener("load", function() {
      restoreMasterFilterState();
      updateFilterChips();
      filterRows();
    });
    function showComingSoon(name) {
      alert(name + " 메뉴는 아직 준비 중입니다. 지금은 쿠팡 PO 변환과 기초자료 관리를 사용할 수 있습니다.");
    }
  </script>
</head>
<body>
  <div class="app">
    <aside class="side">
      <div class="logo"><div class="logo-mark">B</div><div>보니애가구<br><span style="font-size:13px;font-weight:500;color:#b8c8d9;">업무 시스템</span></div></div>
      <div class="nav-item active" onclick="location.href='/master'">쿠팡 기초자료 관리</div>
      <div class="nav-item" onclick="location.href='/'">PO변환</div>
      <div class="nav-item" onclick="location.href='/sales/folders'">월별납품관리</div>
      <div class="nav-item" onclick="location.href='/check'">수량검수</div>
      <div class="nav-item" onclick="location.href='/sales'">매출확인용</div>
      <div class="nav-item" onclick="location.href='/pallet'">파렛트/쉽먼트</div>
      <div class="nav-item" style="margin-top:28px;" onclick="showComingSoon('라벨 출력')">라벨 출력</div>
      <div class="nav-item" onclick="showComingSoon('배송완료 관리')">배송완료 관리</div>
    </aside>
    <main class="main">
      <h1>쿠팡 기초자료 관리</h1>
      <div class="sub">저장된 기초자료에서 금액, 심플웍스 No, 상품 사이즈, 무게, 납품불가 표시를 수정합니다.</div>
      {message}
      <section class="panel">
        <form method="post" action="/master/upload" enctype="multipart/form-data">
          <input type="hidden" name="next" value="master">
          <div class="toolbar">
            <label>기초자료 엑셀 다시 올리기
              <input type="file" name="master" accept=".xlsx" required>
            </label>
            <div class="toolbar-actions">
              <button class="btn secondary" type="submit">현재 기초자료 교체</button>
            </div>
          </div>
        </form>
      </section>
      <form method="post" action="/master/save">
        <section class="panel">
          <div class="add-grid">
            <div class="field">
              <label>새 SKU ID</label>
              <input type="text" name="new_sku" placeholder="예: 12345678">
            </div>
            <div class="field">
              <label>상품명</label>
              <input type="text" name="new_name" placeholder="새 상품명을 입력">
            </div>
            <div class="field">
              <label>금액</label>
              <input type="number" name="new_amount" min="0" placeholder="0">
            </div>
            <div class="field">
              <label>심플웍스 No</label>
              <input type="text" name="new_simple_no" placeholder="예: 7463">
            </div>
            <div class="field">
              <label>가로(mm)</label>
              <input type="number" name="new_width_mm" min="0" placeholder="0">
            </div>
            <div class="field">
              <label>세로(mm)</label>
              <input type="number" name="new_depth_mm" min="0" placeholder="0">
            </div>
            <div class="field">
              <label>높이(mm)</label>
              <input type="number" name="new_height_mm" min="0" placeholder="0">
            </div>
            <div class="field">
              <label>무게(kg)</label>
              <input type="number" name="new_weight_kg" min="0" step="0.1" placeholder="0">
            </div>
            <div class="field">
              <label>바코드</label>
              <input type="text" name="new_barcode" placeholder="R...">
            </div>
            <label class="check-field">
              <input type="checkbox" name="new_unavailable" value="1"> 납품불가
            </label>
          </div>
          <div class="toolbar">
            <input class="search" id="q" oninput="filterRows()" placeholder="스큐 아이디 또는 상품명 검색" />
            <div class="toolbar-actions">
              <button class="btn secondary" type="button" onclick="location.href='/'">PO 변환으로 이동</button>
              <button class="btn" type="submit">수정 내용 저장</button>
            </div>
          </div>
          <div class="filter-row">
            <label>납품상태
              <select id="filter-unavailable" onchange="filterRows()">
                <option value="">전체 상품</option>
                <option value="0">납품가능 상품</option>
                <option value="1">납품불가 상품</option>
              </select>
              <span id="filter-count" class="filter-count"></span>
            </label>
            <label>정렬
              <select id="sort-master" onchange="filterRows()">
                <option value="">기본순</option>
                <option value="name-asc">상품명 오름차순</option>
                <option value="name-desc">상품명 내림차순</option>
                <option value="amount-asc">금액 낮은순</option>
                <option value="amount-desc">금액 높은순</option>
                <option value="simple-desc">심플웍스 No 큰수</option>
                <option value="simple-asc">심플웍스 No 작은수</option>
              </select>
            </label>
            <label>심플웍스 No
              <select id="filter-simple-no" onchange="filterRows()">
                <option value="">전체</option>
                <option value="1">입력 있음</option>
                <option value="0">입력 없음</option>
              </select>
            </label>
            <label>금액
              <select id="filter-amount" onchange="filterRows()">
                <option value="">전체</option>
                <option value="1">입력 있음</option>
                <option value="0">입력 없음</option>
              </select>
            </label>
          </div>
          <div class="table-wrap">
            <table class="resizable-table">
              <colgroup>
                <col style="width:130px;">
                <col style="width:420px;">
                <col style="width:120px;">
                <col style="width:120px;">
                <col style="width:95px;">
                <col style="width:95px;">
                <col style="width:95px;">
                <col style="width:95px;">
                <col style="width:100px;">
                <col style="width:170px;">
              </colgroup>
              <thead>
                <tr>
                  <th>SKU ID<br><span class="muted">스큐 아이디</span></th>
                  <th class="product">상품명</th>
                  <th>금액</th>
                  <th>심플웍스 No</th>
                  <th>가로(mm)</th>
                  <th>세로(mm)</th>
                  <th>높이(mm)</th>
                  <th>무게(kg)</th>
                  <th>납품불가</th>
                  <th>바코드</th>
                </tr>
              </thead>
              <tbody>
                {rows}
              </tbody>
            </table>
          </div>
        </section>
      </form>
    </main>
  </div>
</body>
</html>"""


def safe_name(name: str) -> str:
    return Path(name).name.replace("/", "_").replace("\\", "_")


def get_local_master_path() -> Path | None:
    if not MASTER_DIR.exists():
        return None
    files = [
        path
        for path in MASTER_DIR.glob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    ]
    if not files:
        return None
    return max(files, key=lambda path: path.stat().st_mtime)


def get_saved_master_path() -> Path | None:
    master_path = get_local_master_path()
    if master_path is not None:
        return master_path
    restore_file_from_supabase("master_file", MASTER_DIR / "기초자료.xlsx")
    return get_local_master_path()


def norm_header(value: object) -> str:
    return "".join(str(value or "").lower().split())


def find_header_col(ws, candidates: list[str]) -> int | None:
    headers = {norm_header(ws.cell(1, col).value): col for col in range(1, ws.max_column + 1)}
    normalized = [norm_header(candidate) for candidate in candidates]
    for candidate in normalized:
        if candidate in headers:
            return headers[candidate]
    for header, col in headers.items():
        if any(candidate in header for candidate in normalized):
            return col
    return None


def find_master_columns(ws) -> dict[str, int]:
    sku_col = find_header_col(ws, ["SKU ID", "상품코드", "스큐 아이디", "스큐"])
    qty_col = find_header_col(ws, ["요청", "납품가능수량", "수량"])
    name_col = find_header_col(ws, ["상품명"])
    amount_col = find_header_col(ws, ["금액", "단가", "매입가", "공급가"])
    simple_no_col = find_header_col(ws, ["심플웍스 No", "심플웍스번호", "심플웍스NO", "심플 No", "심플번호"])
    barcode_col = find_header_col(ws, ["바코드"])
    unavailable_col = find_header_col(ws, ["납품불가", "납품불가여부", "제외"])
    width_col = find_header_col(ws, ["가로(mm)", "가로", "폭", "width", "width_mm"])
    depth_col = find_header_col(ws, ["세로(mm)", "세로", "깊이", "depth", "depth_mm"])
    height_col = find_header_col(ws, ["높이(mm)", "높이", "height", "height_mm"])
    weight_col = find_header_col(ws, ["무게(kg)", "무게", "중량", "weight", "weight_kg"])
    if unavailable_col is None:
        for col in range(1, ws.max_column + 1):
            if any("납품불가" in str(ws.cell(row, col).value or "") for row in range(2, min(ws.max_row, 80) + 1)):
                unavailable_col = col
                break
    if unavailable_col is None:
        insert_at = (name_col or 3) + 1
        ws.insert_cols(insert_at)
        ws.cell(1, insert_at).value = "납품불가"
        unavailable_col = insert_at
        if barcode_col and barcode_col >= insert_at:
            barcode_col += 1
        if amount_col and amount_col >= insert_at:
            amount_col += 1
        if simple_no_col and simple_no_col >= insert_at:
            simple_no_col += 1
    if amount_col is None:
        insert_at = (name_col or 3) + 1
        ws.insert_cols(insert_at)
        ws.cell(1, insert_at).value = "금액"
        amount_col = insert_at
        if barcode_col and barcode_col >= insert_at:
            barcode_col += 1
        if unavailable_col and unavailable_col >= insert_at:
            unavailable_col += 1
        if simple_no_col and simple_no_col >= insert_at:
            simple_no_col += 1
    if simple_no_col is None:
        insert_at = (amount_col or name_col or 3) + 1
        ws.insert_cols(insert_at)
        ws.cell(1, insert_at).value = "심플웍스 No"
        simple_no_col = insert_at
        if barcode_col and barcode_col >= insert_at:
            barcode_col += 1
        if unavailable_col and unavailable_col >= insert_at:
            unavailable_col += 1
    width_col = find_header_col(ws, ["가로(mm)", "가로", "폭", "width", "width_mm"])
    depth_col = find_header_col(ws, ["세로(mm)", "세로", "깊이", "depth", "depth_mm"])
    height_col = find_header_col(ws, ["높이(mm)", "높이", "height", "height_mm"])
    weight_col = find_header_col(ws, ["무게(kg)", "무게", "중량", "weight", "weight_kg"])
    size_columns = [
        ("width_mm", width_col, "가로(mm)"),
        ("depth_mm", depth_col, "세로(mm)"),
        ("height_mm", height_col, "높이(mm)"),
        ("weight_kg", weight_col, "무게(kg)"),
    ]
    resolved_size_columns: dict[str, int] = {}
    for key, col, header in size_columns:
        if col is None:
            col = ws.max_column + 1
            ws.cell(1, col).value = header
        resolved_size_columns[key] = col
    if sku_col is None or qty_col is None or name_col is None:
        raise ValueError("기초자료에서 SKU ID, 요청, 상품명 열을 찾지 못했습니다.")
    return {
        "sku": sku_col,
        "qty": qty_col,
        "name": name_col,
        "amount": amount_col,
        "simple_no": simple_no_col,
        **resolved_size_columns,
        "barcode": barcode_col or 0,
        "unavailable": unavailable_col,
    }


def normalize_master_file(master_path: Path) -> None:
    wb = load_workbook(master_path)
    ws = wb.active
    find_master_columns(ws)
    wb.save(master_path)


def replace_saved_master_file(name: str, master_bytes: bytes) -> Path:
    MASTER_DIR.mkdir(parents=True, exist_ok=True)
    current_master = get_local_master_path()
    if current_master is not None:
        backup_dir = MASTER_DIR / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_path = backup_dir / f"{current_master.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        shutil.copy2(current_master, backup_path)
    for path in MASTER_DIR.glob("*.xlsx"):
        if path.is_file() and not path.name.startswith("~$"):
            path.unlink(missing_ok=True)
    target_path = MASTER_DIR / name
    target_path.write_bytes(master_bytes)
    normalize_master_file(target_path)
    backup_master_file_to_supabase(target_path)
    return target_path


def render_master_rows(master_path: Path) -> str:
    wb = load_workbook(master_path, data_only=True)
    ws = wb.active
    cols = find_master_columns(ws)
    rows: list[str] = []
    for row in range(2, ws.max_row + 1):
        sku = str(ws.cell(row, cols["sku"]).value or "").strip()
        name = str(ws.cell(row, cols["name"]).value or "").strip()
        if not sku and not name:
            continue
        amount = str(ws.cell(row, cols["amount"]).value or "").strip()
        simple_no = str(ws.cell(row, cols["simple_no"]).value or "").strip()
        width_mm = str(ws.cell(row, cols["width_mm"]).value or "").strip()
        depth_mm = str(ws.cell(row, cols["depth_mm"]).value or "").strip()
        height_mm = str(ws.cell(row, cols["height_mm"]).value or "").strip()
        weight_kg = str(ws.cell(row, cols["weight_kg"]).value or "").strip()
        barcode = str(ws.cell(row, cols["barcode"]).value or "").strip() if cols["barcode"] else ""
        unavailable = "납품불가" in str(ws.cell(row, cols["unavailable"]).value or "")
        checked = " checked" if unavailable else ""
        simple_flag = "1" if simple_no else "0"
        amount_flag = "1" if parse_int(amount) > 0 else "0"
        unavailable_flag = "1" if unavailable else "0"
        rows.append(
            f'<tr data-unavailable="{unavailable_flag}" data-simple-no="{simple_flag}" data-amount="{amount_flag}">'
            f'<td class="sku">{html.escape(sku)}<input type="hidden" name="row" value="{row}"></td>'
            f'<td class="product">{html.escape(name)}</td>'
            f'<td><input class="qty" name="amount_{row}" value="{html.escape(amount)}" inputmode="numeric"></td>'
            f'<td><input class="qty" name="simple_no_{row}" value="{html.escape(simple_no)}" inputmode="numeric"></td>'
            f'<td><input class="qty" name="width_mm_{row}" value="{html.escape(width_mm)}" inputmode="numeric"></td>'
            f'<td><input class="qty" name="depth_mm_{row}" value="{html.escape(depth_mm)}" inputmode="numeric"></td>'
            f'<td><input class="qty" name="height_mm_{row}" value="{html.escape(height_mm)}" inputmode="numeric"></td>'
            f'<td><input class="qty" name="weight_kg_{row}" value="{html.escape(weight_kg)}" inputmode="decimal"></td>'
            f'<td><input type="checkbox" name="unavailable_{row}" value="1"{checked}></td>'
            f"<td>{html.escape(barcode)}</td>"
            "</tr>"
        )
    return "\n".join(rows)


def save_master_form(form: cgi.FieldStorage) -> int:
    master_path = get_saved_master_path()
    if master_path is None:
        raise ValueError("저장된 기초자료가 없습니다.")

    backup_dir = MASTER_DIR / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup_path = backup_dir / f"{master_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    shutil.copy2(master_path, backup_path)

    wb = load_workbook(master_path)
    ws = wb.active
    cols = find_master_columns(ws)
    if not ws.cell(1, cols["unavailable"]).value:
        ws.cell(1, cols["unavailable"]).value = "납품불가"

    changed = 0
    new_sku = form["new_sku"].value.strip() if "new_sku" in form else ""
    new_name = form["new_name"].value.strip() if "new_name" in form else ""
    if new_sku or new_name:
        if not new_sku:
            raise ValueError("새 상품을 추가하려면 SKU ID를 입력해야 합니다.")
        if not new_name:
            raise ValueError("새 상품을 추가하려면 상품명을 입력해야 합니다.")
        existing_skus = {
            str(ws.cell(row, cols["sku"]).value or "").strip()
            for row in range(2, ws.max_row + 1)
        }
        if new_sku in existing_skus:
            raise ValueError(f"이미 등록된 SKU ID입니다: {new_sku}")
        new_row = ws.max_row + 1
        ws.cell(new_row, cols["sku"]).value = new_sku
        ws.cell(new_row, cols["qty"]).value = ""
        ws.cell(new_row, cols["name"]).value = new_name
        ws.cell(new_row, cols["amount"]).value = form["new_amount"].value.strip() if "new_amount" in form else ""
        ws.cell(new_row, cols["simple_no"]).value = form["new_simple_no"].value.strip() if "new_simple_no" in form else ""
        ws.cell(new_row, cols["width_mm"]).value = form["new_width_mm"].value.strip() if "new_width_mm" in form else ""
        ws.cell(new_row, cols["depth_mm"]).value = form["new_depth_mm"].value.strip() if "new_depth_mm" in form else ""
        ws.cell(new_row, cols["height_mm"]).value = form["new_height_mm"].value.strip() if "new_height_mm" in form else ""
        ws.cell(new_row, cols["weight_kg"]).value = form["new_weight_kg"].value.strip() if "new_weight_kg" in form else ""
        if cols["barcode"]:
            ws.cell(new_row, cols["barcode"]).value = form["new_barcode"].value.strip() if "new_barcode" in form else ""
        ws.cell(new_row, cols["unavailable"]).value = "납품불가" if "new_unavailable" in form else ""
        changed += 1

    row_values = form["row"] if "row" in form else []
    if not isinstance(row_values, list):
        row_values = [row_values]
    for item in row_values:
        row = int(item.value)
        amount_key = f"amount_{row}"
        simple_no_key = f"simple_no_{row}"
        width_key = f"width_mm_{row}"
        depth_key = f"depth_mm_{row}"
        height_key = f"height_mm_{row}"
        weight_key = f"weight_kg_{row}"
        unavailable_key = f"unavailable_{row}"
        new_amount = form[amount_key].value.strip() if amount_key in form else ""
        new_simple_no = form[simple_no_key].value.strip() if simple_no_key in form else ""
        new_width = form[width_key].value.strip() if width_key in form else ""
        new_depth = form[depth_key].value.strip() if depth_key in form else ""
        new_height = form[height_key].value.strip() if height_key in form else ""
        new_weight = form[weight_key].value.strip() if weight_key in form else ""
        new_unavailable = "납품불가" if unavailable_key in form else ""
        old_amount = str(ws.cell(row, cols["amount"]).value or "").strip()
        old_simple_no = str(ws.cell(row, cols["simple_no"]).value or "").strip()
        old_width = str(ws.cell(row, cols["width_mm"]).value or "").strip()
        old_depth = str(ws.cell(row, cols["depth_mm"]).value or "").strip()
        old_height = str(ws.cell(row, cols["height_mm"]).value or "").strip()
        old_weight = str(ws.cell(row, cols["weight_kg"]).value or "").strip()
        old_unavailable = "납품불가" if "납품불가" in str(ws.cell(row, cols["unavailable"]).value or "") else ""
        if new_amount != old_amount:
            ws.cell(row, cols["amount"]).value = new_amount
            changed += 1
        if new_simple_no != old_simple_no:
            ws.cell(row, cols["simple_no"]).value = new_simple_no
            changed += 1
        if new_width != old_width:
            ws.cell(row, cols["width_mm"]).value = new_width
            changed += 1
        if new_depth != old_depth:
            ws.cell(row, cols["depth_mm"]).value = new_depth
            changed += 1
        if new_height != old_height:
            ws.cell(row, cols["height_mm"]).value = new_height
            changed += 1
        if new_weight != old_weight:
            ws.cell(row, cols["weight_kg"]).value = new_weight
            changed += 1
        if new_unavailable != old_unavailable:
            ws.cell(row, cols["unavailable"]).value = new_unavailable
            changed += 1
    wb.save(master_path)
    backup_master_file_to_supabase(master_path)
    return changed


def update_master_amounts_from_lines(lines) -> int:
    master_path = get_saved_master_path()
    if master_path is None:
        return 0

    price_by_sku: dict[str, int] = {}
    for line in lines:
        sku = str(line.sku_id or "").strip()
        if not sku or sku in price_by_sku:
            continue
        prefer_inbound = has_inbound_sales_data(lines)
        sales_qty = get_sales_qty(line, prefer_inbound)
        sales_amount = get_sales_amount(line, prefer_inbound)
        if sales_qty <= 0 or sales_amount <= 0:
            continue
        unit_price = round(sales_amount / sales_qty) if sales_qty else parse_int(getattr(line, "purchase_price", 0))
        if unit_price <= 0 and line.available_qty:
            unit_price = round(parse_int(getattr(line, "order_amount", 0)) / line.available_qty)
        if unit_price > 0:
            price_by_sku[sku] = unit_price

    if not price_by_sku:
        return 0

    wb = load_workbook(master_path)
    ws = wb.active
    cols = find_master_columns(ws)
    changed = 0
    for row in range(2, ws.max_row + 1):
        sku = str(ws.cell(row, cols["sku"]).value or "").strip()
        if sku not in price_by_sku:
            continue
        current_amount = parse_int(ws.cell(row, cols["amount"]).value)
        if current_amount > 0:
            continue
        ws.cell(row, cols["amount"]).value = price_by_sku[sku]
        changed += 1

    if changed:
        wb.save(master_path)
    return changed


def parse_month(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return datetime.now().strftime("%Y-%m")
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(text[:19], fmt).strftime("%Y-%m")
        except ValueError:
            continue
    if len(text) >= 7:
        return text[:7].replace("/", "-")
    return datetime.now().strftime("%Y-%m")


def parse_date(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return datetime.now().strftime("%Y-%m-%d")
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(text[:19], fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    if len(text) >= 10:
        return text[:10].replace("/", "-")
    return datetime.now().strftime("%Y-%m-%d")


def parse_int(value, default: int = 0) -> int:
    text = str(value or "").replace(",", "").replace("원", "").strip()
    if not text:
        return default
    try:
        return int(float(text))
    except ValueError:
        return default


def has_inbound_sales_data(lines) -> bool:
    return any(
        parse_int(getattr(line, "inbound_qty", 0)) > 0
        or parse_int(getattr(line, "inbound_amount", 0)) > 0
        for line in lines
    )


def get_sales_qty(line, prefer_inbound: bool = False) -> int:
    inbound_qty = parse_int(getattr(line, "inbound_qty", 0))
    if prefer_inbound:
        return inbound_qty
    return parse_int(getattr(line, "available_qty", 0))


def get_sales_amount(line, prefer_inbound: bool = False) -> int:
    inbound_amount = parse_int(getattr(line, "inbound_amount", 0))
    if prefer_inbound:
        return inbound_amount
    return parse_int(getattr(line, "order_amount", 0))


def describe_sales_date_breakdown(lines) -> str:
    prefer_inbound = has_inbound_sales_data(lines)
    by_day = defaultdict(lambda: {"qty": 0, "amount": 0, "po_numbers": set()})
    for line in lines:
        qty = get_sales_qty(line, prefer_inbound)
        amount = get_sales_amount(line, prefer_inbound)
        if qty <= 0 and amount <= 0:
            continue
        day = parse_date(line.inbound_date)
        by_day[day]["qty"] += qty
        by_day[day]["amount"] += amount
        by_day[day]["po_numbers"].add(line.po_no)
    if not by_day:
        return ""
    parts = []
    for day in sorted(by_day):
        values = by_day[day]
        parts.append(
            f"{day}: PO {len(values['po_numbers'])}개, 수량 {values['qty']:,}개, 금액 {values['amount']:,}원"
        )
    return " 입고예정일 기준으로 " + " / ".join(parts) + "으로 나누어 등록했습니다."


def ensure_monthly_sales_book():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    restore_sales_files_from_supabase()
    headers = ["일자", "월", "입고예정일", "SKU ID", "상품명", "납품수량", "금액", "바코드", "비고", "수정수량", "수정메모"]
    if SALES_LEDGER_PATH.exists():
        wb = load_workbook(SALES_LEDGER_PATH)
        ws = wb.active
        changed = False
        for col, header in enumerate(headers, start=1):
            if ws.cell(1, col).value != header:
                ws.cell(1, col).value = header
                changed = True
        if changed:
            wb.save(SALES_LEDGER_PATH)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "월매출"
        ws.append(headers)
        wb.save(SALES_LEDGER_PATH)
    return wb, ws


def update_monthly_sales(lines) -> tuple[int, int]:
    wb, ws = ensure_monthly_sales_book()
    po_numbers = {line.po_no for line in lines}
    prefer_inbound = has_inbound_sales_data(lines)

    rows_to_keep = []
    for row in range(2, ws.max_row + 1):
        remarks = str(ws.cell(row, 9).value or "")
        existing_po_numbers = {
            part.strip()
            for part in remarks.replace("PO:", "").split(",")
            if part.strip()
        }
        if not (existing_po_numbers & po_numbers):
            rows_to_keep.append([ws.cell(row, col).value for col in range(1, 12)])

    ws.delete_rows(2, max(ws.max_row - 1, 0))
    for row_values in rows_to_keep:
        ws.append(row_values)

    grouped = defaultdict(lambda: {
        "month": "",
        "inbound_date": "",
        "sku": "",
        "name": "",
        "qty": 0,
        "amount": 0,
        "barcode": "",
        "po_numbers": set(),
    })
    for line in lines:
        sales_qty = get_sales_qty(line, prefer_inbound)
        sales_amount = get_sales_amount(line, prefer_inbound)
        if sales_qty <= 0 and sales_amount <= 0:
            continue
        day = parse_date(line.inbound_date)
        key = (day, line.sku_id)
        item = grouped[key]
        item["month"] = parse_month(line.inbound_date)
        item["inbound_date"] = line.inbound_date
        item["sku"] = line.sku_id
        item["name"] = item["name"] or line.product_name
        item["qty"] += sales_qty
        item["amount"] += sales_amount
        item["barcode"] = item["barcode"] or line.barcode
        item["po_numbers"].add(line.po_no)

    saved_count = 0
    for (day, _sku), item in sorted(grouped.items()):
        ws.append([
            day,
            item["month"],
            item["inbound_date"],
            item["sku"],
            item["name"],
            item["qty"],
            item["amount"],
            item["barcode"],
            ", ".join(sorted(item["po_numbers"])),
            "",
            "",
        ])
        saved_count += 1

    for column_cells in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 10), 55)
    wb.save(SALES_LEDGER_PATH)
    write_sales_display_workbook()
    backup_sales_files_to_supabase()
    return saved_count, sum(get_sales_amount(line, prefer_inbound) for line in lines)


def load_monthly_sales_summary() -> tuple[list[tuple[str, int, int, int]], list[list[object]]]:
    restore_sales_files_from_supabase()
    if not SALES_LEDGER_PATH.exists():
        return [], []
    wb = load_workbook(SALES_LEDGER_PATH, data_only=True)
    ws = wb.active
    summary = defaultdict(lambda: {"qty": 0, "amount": 0, "pos": set()})
    rows = []
    for row in range(2, ws.max_row + 1):
        day = str(ws.cell(row, 1).value or "")
        sku = str(ws.cell(row, 4).value or "")
        name = str(ws.cell(row, 5).value or "")
        original_qty = parse_int(ws.cell(row, 6).value)
        original_amount = parse_int(ws.cell(row, 7).value)
        remarks = str(ws.cell(row, 9).value or "")
        adjusted_qty_raw = str(ws.cell(row, 10).value or "").strip()
        adjusted_memo = str(ws.cell(row, 11).value or "").strip()
        adjusted_qty = parse_int(adjusted_qty_raw, original_qty) if adjusted_qty_raw else original_qty
        unit_price = round(original_amount / original_qty) if original_qty else 0
        amount = unit_price * adjusted_qty
        clean_po_numbers = []
        if not day:
            continue
        summary[day]["qty"] += adjusted_qty
        summary[day]["amount"] += amount
        for po_no in remarks.replace("PO:", "").split(","):
            po_no = po_no.strip()
            if po_no:
                summary[day]["pos"].add(po_no)
                clean_po_numbers.append(po_no)
        rows.append([
            row,
            day,
            sku,
            name,
            original_qty,
            adjusted_qty,
            unit_price,
            amount,
            ", ".join(clean_po_numbers),
            adjusted_memo,
            adjusted_qty != original_qty or bool(adjusted_memo),
        ])
    summary_rows = [
        (month, values["qty"], values["amount"], len(values["pos"]))
        for month, values in sorted(summary.items())
    ]
    return summary_rows, rows[-300:]


def sales_extra_amounts(amount: int) -> tuple[int, int]:
    vat_excluded = round(amount / 1.1)
    ad_budget = round(vat_excluded * 0.035)
    return vat_excluded, ad_budget


def load_year_manual() -> dict:
    if not YEAR_MANUAL_PATH.exists():
        return {}
    try:
        return json.loads(YEAR_MANUAL_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def save_year_manual(form: cgi.FieldStorage) -> int:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    manual = load_year_manual()
    changed = 0
    for month in range(1, 13):
        key = f"{datetime.now().year}-{month:02d}"
        if key not in manual:
            manual[key] = {}
        for field in ("tester", "ad", "support", "discount", "partner_discount", "extra_ad"):
            form_key = f"{field}_{month}"
            new_value = form[form_key].value.strip() if form_key in form else ""
            old_value = str(manual[key].get(field, ""))
            if new_value != old_value:
                manual[key][field] = new_value
                changed += 1
    YEAR_MANUAL_PATH.write_text(json.dumps(manual, ensure_ascii=False, indent=2), encoding="utf-8")
    return changed


def render_year_rows(summary_rows: list[tuple[str, int, int, int]]) -> str:
    manual = load_year_manual()
    year = datetime.now().year
    monthly_sales = defaultdict(int)
    for day, _qty, amount, _po_count in summary_rows:
        month_key = str(day)[:7]
        monthly_sales[month_key] += amount

    rows = []
    totals = defaultdict(int)
    for month in range(1, 13):
        key = f"{year}-{month:02d}"
        amount = monthly_sales[key]
        vat_excluded, ad_budget = sales_extra_amounts(amount)
        values = manual.get(key, {})
        tester = parse_int(values.get("tester", ""))
        ad = parse_int(values.get("ad", ""))
        support = parse_int(values.get("support", ""))
        discount = parse_int(values.get("discount", ""))
        partner_discount = parse_int(values.get("partner_discount", ""))
        extra_ad = parse_int(values.get("extra_ad", ""))
        over = ad_budget - tester - ad
        totals["sales"] += vat_excluded
        totals["budget"] += ad_budget
        totals["tester"] += tester
        totals["ad"] += ad
        totals["over"] += over
        totals["support"] += support
        totals["discount"] += discount
        totals["partner_discount"] += partner_discount
        totals["extra_ad"] += extra_ad
        rows.append(
            f'<tr class="year-row" data-month="{month:02d}" data-sales="{vat_excluded}" data-budget="{ad_budget}">'
            f"<td>{month}월</td>"
            f'<td class="auto-cell year-sales">{vat_excluded:,}원</td>'
            f'<td class="auto-cell year-budget">{ad_budget:,}원</td>'
            f'<td><input class="money-input tester-input" name="tester_{month}" value="{html.escape(str(values.get("tester", "")), quote=True)}"></td>'
            f'<td><input class="money-input ad-input" name="ad_{month}" value="{html.escape(str(values.get("ad", "")), quote=True)}"></td>'
            f'<td class="over-cell{" negative" if over < 0 else ""}">{over:,}원</td>'
            f'<td><input class="money-input support-input" name="support_{month}" value="{html.escape(str(values.get("support", "")), quote=True)}"><span class="support-display" style="display:none">{support:,}원</span></td>'
            f'<td><input class="money-input discount-input" name="discount_{month}" value="{html.escape(str(values.get("discount", "")), quote=True)}"><span class="discount-display" style="display:none">{discount:,}원</span></td>'
            f'<td><input class="money-input partner-discount-input" name="partner_discount_{month}" value="{html.escape(str(values.get("partner_discount", "")), quote=True)}"><span class="partner-discount-display" style="display:none">{partner_discount:,}원</span></td>'
            f'<td><input class="money-input extra-ad-input" name="extra_ad_{month}" value="{html.escape(str(values.get("extra_ad", "")), quote=True)}"><span class="extra-ad-display" style="display:none">{extra_ad:,}원</span></td>'
            f"</tr>"
        )
    rows.append(
        '<tr class="year-total-row" style="background:#fff2cc;font-weight:800;">'
        '<td>연도합계</td>'
        f'<td class="year-total-sales">{totals["sales"]:,}원</td>'
        f'<td class="year-total-budget">{totals["budget"]:,}원</td>'
        f'<td class="year-total-tester">{totals["tester"]:,}원</td>'
        f'<td class="year-total-ad">{totals["ad"]:,}원</td>'
        f'<td class="year-total-over">{totals["over"]:,}원</td>'
        f'<td class="year-total-support">{totals["support"]:,}원</td>'
        f'<td class="year-total-discount">{totals["discount"]:,}원</td>'
        f'<td class="year-total-partner-discount">{totals["partner_discount"]:,}원</td>'
        f'<td class="year-total-extra-ad">{totals["extra_ad"]:,}원</td>'
        '</tr>'
    )
    return "\n".join(rows)


def write_sales_display_workbook() -> Path:
    summary_rows, detail_rows = load_monthly_sales_summary()
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "매출확인용"

    title_fill = PatternFill("solid", fgColor="D9EAF7")
    header_fill = PatternFill("solid", fgColor="E7E6E6")
    total_fill = PatternFill("solid", fgColor="FFF2CC")
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")

    row_no = 1
    if summary_rows:
        ws.cell(row_no, 1, "일자별 납품 상품 합계")
        ws.merge_cells(start_row=row_no, start_column=1, end_row=row_no, end_column=7)
        ws.cell(row_no, 1).font = Font(bold=True, size=14)
        ws.cell(row_no, 1).fill = title_fill
        ws.cell(row_no, 1).alignment = center
        row_no += 2

        by_day = defaultdict(list)
        for _row_no, day, sku, name, _original_qty, qty, unit_price, amount, remarks, memo, _changed in detail_rows:
            by_day[day].append((sku, name, qty, unit_price, amount, remarks, memo))

        for day in sorted(by_day.keys()):
            day_qty = sum(row[2] for row in by_day[day])
            day_amount = sum(row[4] for row in by_day[day])
            ws.cell(row_no, 1, f"{day} 합계: 수량 {day_qty:,}개 / 금액 {day_amount:,}원")
            ws.merge_cells(start_row=row_no, start_column=1, end_row=row_no, end_column=7)
            ws.cell(row_no, 1).fill = title_fill
            ws.cell(row_no, 1).font = bold
            ws.cell(row_no, 1).alignment = center
            row_no += 1

            for col, header in enumerate(["SKU ID", "상품명", "납품수량", "단가", "금액", "비고", "메모"], start=1):
                cell = ws.cell(row_no, col, header)
                cell.fill = header_fill
                cell.font = bold
                cell.alignment = center
            row_no += 1

            for sku, name, qty, unit_price, amount, remarks, memo in sorted(by_day[day], key=lambda item: str(item[0])):
                ws.cell(row_no, 1, sku)
                ws.cell(row_no, 2, name)
                ws.cell(row_no, 3, qty)
                ws.cell(row_no, 4, unit_price)
                ws.cell(row_no, 5, amount)
                ws.cell(row_no, 6, remarks)
                ws.cell(row_no, 7, memo)
                ws.cell(row_no, 3).number_format = '#,##0'
                ws.cell(row_no, 4).number_format = '#,##0'
                ws.cell(row_no, 5).number_format = '#,##0'
                row_no += 1

            ws.cell(row_no, 1, "합계")
            ws.cell(row_no, 3, day_qty)
            ws.cell(row_no, 5, day_amount)
            for col in range(1, 8):
                ws.cell(row_no, col).fill = total_fill
                ws.cell(row_no, col).font = bold
            ws.cell(row_no, 3).number_format = '#,##0'
            ws.cell(row_no, 5).number_format = '#,##0'
            row_no += 2
    else:
        ws.append(["아직 누적된 월매출 자료가 없습니다."])

    widths = {"A": 14, "B": 60, "C": 12, "D": 12, "E": 14, "F": 28, "G": 32}
    for column_letter, width in widths.items():
        ws.column_dimensions[column_letter].width = width
    for cells in ws.iter_rows():
        for cell in cells:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A3"
    wb.save(MONTHLY_SALES_PATH)
    return MONTHLY_SALES_PATH


def render_sales_page(message: str = "", folder_mode: bool = False) -> str:
    summary_rows, detail_rows = load_monthly_sales_summary()
    current_month = datetime.now().strftime("%Y-%m")
    visible_detail_rows = detail_rows if folder_mode else [
        row for row in detail_rows if str(row[1]).startswith(current_month)
    ]
    if summary_rows:
        def sales_extra_amounts(amount: int) -> tuple[int, int]:
            vat_excluded = round(amount / 1.1)
            ad_budget = round(vat_excluded * 0.035)
            return vat_excluded, ad_budget

        summary_html = "\n".join(
            f'<tr class="summary-row" data-day="{html.escape(month)}" data-po-count="{po_count}">'
            f"<td>{html.escape(month)}</td><td>{po_count:,}</td>"
            f'<td class="summary-qty">{qty:,}</td><td class="summary-amount">{amount:,}원</td>'
            f'<td class="summary-vat">{sales_extra_amounts(amount)[0]:,}원</td><td class="summary-budget">{sales_extra_amounts(amount)[1]:,}원</td></tr>'
            for month, qty, amount, po_count in summary_rows
        )
        total_qty = sum(qty for _month, qty, _amount, _po_count in summary_rows)
        total_amount = sum(amount for _month, _qty, amount, _po_count in summary_rows)
        total_po_count = sum(po_count for _month, _qty, _amount, po_count in summary_rows)
        total_vat_excluded, total_ad_budget = sales_extra_amounts(total_amount)
        summary_html += (
            f'<tr class="summary-total-row" style="background:#fff2cc;font-weight:700;">'
            f'<td>총합계</td><td class="summary-total-po">{total_po_count:,}</td><td class="summary-total-qty">{total_qty:,}</td><td class="summary-total-amount">{total_amount:,}원</td>'
            f'<td class="summary-total-vat">{total_vat_excluded:,}원</td><td class="summary-total-budget">{total_ad_budget:,}원</td></tr>'
        )
    else:
        summary_html = '<tr><td colspan="6">아직 누적된 월매출 자료가 없습니다.</td></tr>'

    if visible_detail_rows:
        by_day = defaultdict(list)
        for row_no, day, sku, name, original_qty, qty, unit_price, amount, remarks, memo, changed in visible_detail_rows:
            by_day[day].append((row_no, sku, name, original_qty, qty, unit_price, amount, remarks, memo, changed))
        parts = []
        by_month = defaultdict(list)
        for day in sorted(by_day.keys(), reverse=True):
            by_month[str(day)[:7]].append(day)
        for month in sorted(by_month.keys(), reverse=True):
            if folder_mode:
                parts.append(
                    f'<tr class="month-folder is-closed" data-month="{html.escape(month, quote=True)}">'
                    f'<th colspan="8">{html.escape(month)} '
                    f'<button class="toggle-btn" type="button" onclick="toggleMonthFolder(\'{html.escape(month, quote=True)}\', this)">펼치기</button> '
                    f'<button class="delete-btn" type="submit" formaction="/sales/delete" name="delete_month" value="{html.escape(month, quote=True)}" onclick="return confirmDeleteMonth(\'{html.escape(month, quote=True)}\')">월 삭제</button></th></tr>'
                )
            for day in by_month[month]:
                day_qty = sum(row[4] for row in by_day[day])
                day_amount = sum(row[6] for row in by_day[day])
                hidden_class = " month-hidden" if folder_mode else ""
                parts.append(
                    f'<tr class="detail-day-row{hidden_class}" data-month="{html.escape(month, quote=True)}" data-day="{html.escape(str(day), quote=True)}"><th colspan="8" data-day-total="{html.escape(str(day), quote=True)}" style="background:#e8f1fb;color:#1f4e79;">'
                    f'{html.escape(str(day))} 합계: 수량 {day_qty:,}개 / 금액 {day_amount:,}원 '
                    f'<button class="delete-btn" type="submit" formaction="/sales/delete" name="delete_day" value="{html.escape(str(day), quote=True)}" onclick="return confirmDeleteDay(\'{html.escape(str(day), quote=True)}\')">일자 삭제</button></th></tr>'
                )
                for row_no, sku, name, original_qty, qty, unit_price, amount, remarks, memo, changed in sorted(by_day[day], key=lambda row: str(row[1])):
                    changed_class = " changed" if changed else ""
                    parts.append(
                        f'<tr class="detail-row {changed_class.strip()}{hidden_class}" data-day="{html.escape(str(day), quote=True)}" data-month="{html.escape(str(day)[:7], quote=True)}" data-original-qty="{original_qty}" data-unit-price="{unit_price}" data-search="{html.escape((str(sku) + " " + str(name) + " " + str(remarks) + " " + str(memo)), quote=True)}"><td>{html.escape(str(sku))}'
                        f'<input type="hidden" name="row" value="{row_no}"></td>'
                        f"<td>{html.escape(str(name))}</td>"
                        f'<td class="{changed_class.strip()}"><input class="qty-input" type="number" min="0" step="1" '
                        f'name="qty_{row_no}" value="{qty}"></td>'
                        f'<td>{unit_price:,}원</td><td class="row-amount">{amount:,}원</td><td>{html.escape(str(remarks))}</td>'
                        f'<td><input class="memo-input" type="text" name="memo_{row_no}" value="{html.escape(str(memo), quote=True)}" '
                        f'placeholder="수정 사유"></td>'
                        f'<td><button class="delete-btn" type="submit" formaction="/sales/delete" name="delete_row" value="{row_no}" onclick="return confirm(\'이 상품 줄을 삭제할까요?\')">삭제</button></td></tr>'
                    )
        detail_html = "\n".join(parts)
    else:
        detail_html = '<tr><td colspan="8">해당 월 납품 상품 내역이 없습니다.</td></tr>'
    year_rows = render_year_rows(summary_rows)
    page_title = "월별납품관리" if folder_mode else "매출확인용"
    page_sub = (
        "누적된 납품 상품 내역을 월별 폴더로 펼쳐서 확인합니다."
        if folder_mode
        else "업로드, 연도총매출, 일자별 합계, 해당월 납품 상품을 확인합니다."
    )
    detail_title = "월별 납품 상품 폴더" if folder_mode else f"{current_month} 납품 상품 내역"
    return (
        SALES_PAGE
        .replace("{message}", message)
        .replace("{year_rows}", year_rows)
        .replace("{summary_rows}", summary_html)
        .replace("{detail_rows}", detail_html)
        .replace("{page_title}", page_title)
        .replace("{page_sub}", page_sub)
        .replace("{detail_title}", detail_title)
        .replace("{current_month}", current_month)
        .replace("{sales_active}", "active" if not folder_mode else "")
        .replace("{folders_active}", "active" if folder_mode else "")
        .replace("{year_section_class}", "not-shown" if folder_mode else "")
        .replace("{upload_section_class}", "" if folder_mode else "not-shown")
        .replace("{summary_section_class}", "not-shown" if folder_mode else "")
    )


def normalize_simple_no(value: object) -> str:
    text = str(value or "").strip()
    if text.endswith(".0"):
        text = text[:-2]
    return re.sub(r"\D", "", text)


def extract_simple_no(values: list[object]) -> str:
    for value in values:
        text = str(value or "").strip()
        match = re.search(r"(\d{3,6})\s*[:：]", text)
        if match:
            return match.group(1)
    for value in values:
        no = normalize_simple_no(value)
        if 3 <= len(no) <= 6:
            return no
    return ""


def get_master_simpleworks_maps() -> tuple[dict[str, str], dict[str, dict[str, str]]]:
    master_path = get_saved_master_path()
    if master_path is None:
        raise ValueError("저장된 기초자료가 없습니다. 먼저 기초자료에 심플웍스 No를 입력해주세요.")
    wb = load_workbook(master_path)
    ws = wb.active
    cols = find_master_columns(ws)
    wb.save(master_path)
    sku_to_simple: dict[str, str] = {}
    simple_to_info: dict[str, dict[str, str]] = {}
    for row in range(2, ws.max_row + 1):
        sku = str(ws.cell(row, cols["sku"]).value or "").strip()
        simple_no = normalize_simple_no(ws.cell(row, cols["simple_no"]).value)
        name = str(ws.cell(row, cols["name"]).value or "").strip()
        if not sku or not simple_no:
            continue
        sku_to_simple[sku] = simple_no
        simple_to_info.setdefault(simple_no, {"sku": sku, "name": name})
    return sku_to_simple, simple_to_info


def parse_simpleworks_excel(path: Path) -> dict[str, int]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    header_row = None
    product_col = None
    qty_col = None
    for row in range(1, min(ws.max_row, 30) + 1):
        for col in range(1, ws.max_column + 1):
            text = norm_header(ws.cell(row, col).value)
            if text in ["상품명", "상품"]:
                header_row = row
                product_col = col
            if text == "수량":
                header_row = row if header_row is None else header_row
                qty_col = col
        if product_col and qty_col:
            break

    quantities: dict[str, int] = defaultdict(int)
    start_row = (header_row + 1) if header_row else 1
    for row in range(start_row, ws.max_row + 1):
        values = [ws.cell(row, col).value for col in range(1, ws.max_column + 1)]
        if all(value in [None, ""] for value in values):
            continue
        simple_no = extract_simple_no(values)
        if not simple_no:
            continue
        if qty_col:
            qty = parse_int(ws.cell(row, qty_col).value)
        else:
            qty = 0
            for value in reversed(values):
                text = str(value or "").strip()
                if re.search(r"\d+\s*개?$", text) or isinstance(value, (int, float)):
                    qty = parse_int(text)
                    break
        if qty > 0:
            quantities[simple_no] += qty
    return dict(quantities)


def extract_simpleworks_quantities_from_text(text: str) -> dict[str, int]:
    quantities: dict[str, int] = defaultdict(int)
    pending_simple_no = ""
    for raw_line in text.splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        if not line:
            continue
        simple_match = re.search(r"([0-9$SIl|]{3,7})\s*[:：]", line)
        if simple_match:
            pending_simple_no = (
                simple_match.group(1)
                .replace("$", "9")
                .replace("S", "5")
                .replace("I", "1")
                .replace("l", "1")
                .replace("|", "1")
            )
            pending_simple_no = re.sub(r"\D", "", pending_simple_no)
        elif not pending_simple_no:
            start_match = re.match(r"^\D*(\d{3,6})\b", line)
            if start_match:
                pending_simple_no = start_match.group(1)
        qty_matches = re.findall(r"(\d{1,4})\s*(?:개|m|M)", line)
        if not qty_matches and pending_simple_no:
            trailing_numbers = re.findall(r"\b(\d+)\b", line)
            trailing_numbers = [number for number in trailing_numbers if number != pending_simple_no]
            if trailing_numbers:
                qty_matches = [trailing_numbers[-1]]
        if pending_simple_no and qty_matches:
            quantities[pending_simple_no] += parse_int(qty_matches[-1])
            pending_simple_no = ""
    return dict(quantities)


def get_tesseract_candidates() -> list[str]:
    return [
        os.environ.get("TESSERACT_CMD", ""),
        shutil.which("tesseract") or "",
        "/usr/bin/tesseract",
        "/usr/local/bin/tesseract",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    ]


def get_tesseract_path() -> str | None:
    tesseract_path = shutil.which("tesseract")
    if tesseract_path:
        return tesseract_path
    for candidate in get_tesseract_candidates():
        if candidate and Path(candidate).exists():
            return candidate
    return None


def get_tesseract_language(tesseract_path: str) -> str:
    completed = subprocess.run(
        [tesseract_path, "--list-langs"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
        timeout=20,
    )
    langs = set(re.findall(r"^[a-zA-Z_]+$", completed.stdout, flags=re.MULTILINE))
    if "kor" in langs and "eng" in langs:
        return "kor+eng"
    if "kor" in langs:
        return "kor"
    return "eng"


def clean_ocr_simple_no(text: str) -> str:
    cleaned = (
        str(text or "")
        .replace("$", "9")
        .replace("S", "5")
        .replace("I", "1")
        .replace("l", "1")
        .replace("|", "1")
    )
    digits = re.sub(r"\D", "", cleaned)
    if len(digits) == 5 and digits.endswith("2"):
        return digits[:4]
    if len(digits) > 4:
        return digits[:4]
    return digits


def clean_ocr_qty(text: str) -> int:
    digits = re.sub(r"\D", "", str(text or ""))
    if len(digits) >= 3 and int(digits[-2:]) >= 50:
        digits = digits[:-2]
    return parse_int(digits)


def run_tesseract_text(tesseract_path: str, image_path: Path, mode: str = "6", whitelist: str = "") -> str:
    command = [tesseract_path, str(image_path), "stdout", "-l", get_tesseract_language(tesseract_path), "--psm", mode]
    if whitelist:
        command.extend(["-c", f"tessedit_char_whitelist={whitelist}"])
    completed = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="ignore", timeout=60)
    if completed.returncode != 0:
        raise ValueError(f"캡쳐 이미지 글자읽기에 실패했습니다: {completed.stderr.strip() or 'OCR 오류'}")
    return completed.stdout


def parse_simpleworks_table_image(path: Path, tesseract_path: str) -> dict[str, int]:
    try:
        from PIL import Image, ImageEnhance
    except Exception:
        return {}

    image = Image.open(path).convert("L")
    width, height = image.size
    pixels = image.load()

    horizontal_lines: list[int] = []
    for y in range(height):
        dark_count = sum(1 for x in range(width) if pixels[x, y] < 235)
        if dark_count > width * 0.45:
            if not horizontal_lines or y - horizontal_lines[-1] > 4:
                horizontal_lines.append(y)

    if len(horizontal_lines) < 3:
        return {}

    vertical_candidates: list[tuple[int, int]] = []
    start_y = horizontal_lines[1]
    for x in range(int(width * 0.55), width):
        dark_count = sum(1 for y in range(start_y, height) if pixels[x, y] < 235)
        if dark_count > (height - start_y) * 0.25:
            vertical_candidates.append((x, dark_count))
    divider_x = min(vertical_candidates, key=lambda item: item[0])[0] if vertical_candidates else int(width * 0.82)

    quantities: dict[str, int] = defaultdict(int)
    temp_dir = RUNS_DIR / "ocr_crops"
    temp_dir.mkdir(parents=True, exist_ok=True)
    row_pairs = list(zip(horizontal_lines[1:-1], horizontal_lines[2:]))
    for index, (y1, y2) in enumerate(row_pairs, start=1):
        if y2 - y1 < 18:
            continue
        top = y1 + 4
        bottom = y2 - 4
        if bottom <= top:
            continue

        no_crop = image.crop((0, top, min(55, max(45, divider_x // 9)), bottom))
        no_crop = no_crop.resize((no_crop.width * 6, no_crop.height * 6))
        no_crop = ImageEnhance.Contrast(no_crop).enhance(2.2)
        no_path = temp_dir / f"{path.stem}_no_{index}.png"
        no_crop.save(no_path)
        no_text = run_tesseract_text(tesseract_path, no_path, "7", "0123456789$SIl|")
        simple_no = clean_ocr_simple_no(no_text)

        qty_crop = image.crop((divider_x + 4, top, min(width, divider_x + 56), bottom))
        qty_crop = qty_crop.resize((qty_crop.width * 6, qty_crop.height * 6))
        qty_crop = ImageEnhance.Contrast(qty_crop).enhance(2.2)
        qty_path = temp_dir / f"{path.stem}_qty_{index}.png"
        qty_crop.save(qty_path)
        qty_text = run_tesseract_text(tesseract_path, qty_path, "7", "0123456789")
        qty = clean_ocr_qty(qty_text)

        if simple_no and qty > 0:
            quantities[simple_no] += qty

    return dict(quantities)


def parse_simpleworks_image(path: Path) -> dict[str, int]:
    tesseract_path = get_tesseract_path()
    if not tesseract_path:
        checked = ", ".join(candidate for candidate in get_tesseract_candidates() if candidate) or "없음"
        raise ValueError(
            "캡쳐 이미지를 읽으려면 OCR 프로그램인 Tesseract 설치가 필요합니다. "
            f"지금은 심플웍스 엑셀 업로드를 사용해주세요. 확인한 경로: {checked} / PATH: {os.environ.get('PATH', '')}"
        )
    table_quantities = parse_simpleworks_table_image(path, tesseract_path)
    if table_quantities:
        return table_quantities
    return extract_simpleworks_quantities_from_text(run_tesseract_text(tesseract_path, path, "6"))


def get_coupang_quantities_by_simple_no(date_from: str, date_to: str) -> tuple[dict[str, int], list[tuple[str, str, int]]]:
    sku_to_simple, _simple_to_info = get_master_simpleworks_maps()
    expected: dict[str, int] = defaultdict(int)
    unmapped: list[tuple[str, str, int]] = []
    if not SALES_LEDGER_PATH.exists():
        return {}, []
    wb = load_workbook(SALES_LEDGER_PATH, data_only=True)
    ws = wb.active
    for row in range(2, ws.max_row + 1):
        day = str(ws.cell(row, 1).value or "").strip()
        if date_from and day < date_from:
            continue
        if date_to and day > date_to:
            continue
        sku = str(ws.cell(row, 4).value or "").strip()
        name = str(ws.cell(row, 5).value or "").strip()
        original_qty = parse_int(ws.cell(row, 6).value)
        adjusted_qty_raw = str(ws.cell(row, 10).value or "").strip()
        qty = parse_int(adjusted_qty_raw, original_qty) if adjusted_qty_raw else original_qty
        simple_no = sku_to_simple.get(sku, "")
        if simple_no:
            expected[simple_no] += qty
        elif qty:
            unmapped.append((sku, name, qty))
    return dict(expected), unmapped


def render_check_result(simpleworks_qty: dict[str, int], date_from: str, date_to: str) -> str:
    _sku_to_simple, simple_to_info = get_master_simpleworks_maps()
    coupang_qty, unmapped = get_coupang_quantities_by_simple_no(date_from, date_to)
    all_simple_nos = sorted(set(coupang_qty) | set(simpleworks_qty), key=lambda value: int(value) if value.isdigit() else value)
    rows = []
    ok_count = diff_count = coupang_only = simple_only = 0
    for simple_no in all_simple_nos:
        expected = coupang_qty.get(simple_no, 0)
        actual = simpleworks_qty.get(simple_no, 0)
        diff = actual - expected
        info = simple_to_info.get(simple_no, {"sku": "", "name": ""})
        if expected == actual:
            status = '<span class="ok-text">일치</span>'
            ok_count += 1
        elif expected and actual:
            status = '<span class="bad-text">수량차이</span>'
            diff_count += 1
        elif expected:
            status = '<span class="warn-text">쿠팡만 있음</span>'
            coupang_only += 1
        else:
            status = '<span class="warn-text">심플웍스만 있음</span>'
            simple_only += 1
        rows.append(
            "<tr>"
            f"<td>{status}</td><td>{html.escape(simple_no)}</td><td>{html.escape(info.get('sku', ''))}</td>"
            f"<td>{html.escape(info.get('name', ''))}</td><td>{expected:,}</td><td>{actual:,}</td><td>{diff:,}</td>"
            "</tr>"
        )
    for sku, name, qty in unmapped:
        rows.append(
            '<tr><td><span class="bad-text">기초자료 매칭 없음</span></td>'
            f"<td></td><td>{html.escape(sku)}</td><td>{html.escape(name)}</td><td>{qty:,}</td><td>0</td><td>{-qty:,}</td></tr>"
        )
    if not rows:
        rows.append('<tr><td colspan="7">검수할 자료가 없습니다.</td></tr>')
    return (
        '<section class="panel"><div class="panel-head">검수 결과</div><div class="panel-body">'
        f'<div class="summary"><div>일치<b>{ok_count:,}</b></div><div>수량차이<b>{diff_count:,}</b></div>'
        f'<div>쿠팡만 있음<b>{coupang_only + len(unmapped):,}</b></div><div>심플웍스만 있음<b>{simple_only:,}</b></div></div>'
        '<div class="scroll" style="margin-top:14px;"><table><colgroup><col style="width:120px;"><col style="width:110px;"><col style="width:120px;"><col style="width:380px;"><col style="width:110px;"><col style="width:120px;"><col style="width:100px;"></colgroup>'
        '<thead><tr><th>상태</th><th>심플웍스 No</th><th>SKU ID</th><th>상품명</th><th>쿠팡 수량</th><th>심플웍스 수량</th><th>차이</th></tr></thead>'
        f"<tbody>{''.join(rows)}</tbody></table></div></div></section>"
    )


def render_check_page(message: str = "", result: str = "", date_from: str = "", date_to: str = "") -> str:
    today = datetime.now()
    default_from = date_from or today.strftime("%Y-%m-01")
    default_to = date_to or today.strftime("%Y-%m-%d")
    return (
        CHECK_PAGE
        .replace("{message}", message)
        .replace("{result}", result)
        .replace("{date_from}", html.escape(default_from, quote=True))
        .replace("{date_to}", html.escape(default_to, quote=True))
    )


def handle_sales_upload(form: cgi.FieldStorage) -> tuple[str, str | None]:
    po_items = form["sales_po_files"] if "sales_po_files" in form else None
    if po_items is None:
        raise ValueError("매출확인용 PO 파일을 올려주세요.")
    if not isinstance(po_items, list):
        po_items = [po_items]

    run_id = uuid.uuid4().hex[:12]
    work_dir = RUNS_DIR / f"sales_{run_id}"
    input_dir = work_dir / "uploaded_po"
    input_dir.mkdir(parents=True, exist_ok=True)

    all_lines = []
    file_count = 0
    for item in po_items:
        if not getattr(item, "filename", ""):
            continue
        name = safe_name(item.filename)
        if name.startswith("~$") or not name.lower().endswith(".xlsx"):
            continue
        po_path = input_dir / name
        po_path.write_bytes(item.file.read())
        all_lines.extend(read_po_lines(po_path))
        file_count += 1

    if file_count == 0:
        raise ValueError("처리할 PO 엑셀 파일이 없습니다.")
    if not all_lines:
        raise ValueError("PO 안에서 상품 내역을 찾지 못했습니다.")

    filled_amounts = update_master_amounts_from_lines(all_lines)
    line_count, total_amount = update_monthly_sales(all_lines)
    prefer_inbound = has_inbound_sales_data(all_lines)
    total_qty = sum(get_sales_qty(line, prefer_inbound) for line in all_lines)
    amount_message = f" 기초자료 금액 {filled_amounts}건도 자동으로 채웠습니다." if filled_amounts else ""
    date_message = describe_sales_date_breakdown(all_lines)
    return (
        f"매출확인용으로 저장했습니다. PO {file_count}개, 상품 줄 {line_count}개, "
        f"총 수량 {total_qty:,}개, 납품상품 합계금액 {total_amount:,}원입니다.{amount_message}{date_message}"
    ), None


def save_sales_detail_form(form: cgi.FieldStorage) -> int:
    wb, ws = ensure_monthly_sales_book()
    row_values = form["row"] if "row" in form else []
    if not isinstance(row_values, list):
        row_values = [row_values]

    changed = 0
    for item in row_values:
        row = parse_int(item.value)
        if row < 2 or row > ws.max_row:
            continue
        qty_key = f"qty_{row}"
        memo_key = f"memo_{row}"
        original_qty = parse_int(ws.cell(row, 6).value)
        new_qty_text = form[qty_key].value.strip() if qty_key in form else str(original_qty)
        new_memo = form[memo_key].value.strip() if memo_key in form else ""
        new_qty = parse_int(new_qty_text, original_qty)
        adjusted_value = "" if new_qty == original_qty else new_qty
        old_adjusted = str(ws.cell(row, 10).value or "").strip()
        old_memo = str(ws.cell(row, 11).value or "").strip()
        if str(adjusted_value) != old_adjusted:
            ws.cell(row, 10).value = adjusted_value
            changed += 1
        if new_memo != old_memo:
            ws.cell(row, 11).value = new_memo
            changed += 1
    wb.save(SALES_LEDGER_PATH)
    write_sales_display_workbook()
    backup_sales_files_to_supabase()
    return changed


def delete_sales_detail_form(form: cgi.FieldStorage) -> tuple[int, str]:
    wb, ws = ensure_monthly_sales_book()
    target_day = form["delete_day"].value.strip() if "delete_day" in form else ""
    target_month = form["delete_month"].value.strip() if "delete_month" in form else ""
    target_row = parse_int(form["delete_row"].value) if "delete_row" in form else 0

    rows_to_delete = []
    if target_month:
        for row in range(2, ws.max_row + 1):
            if str(ws.cell(row, 1).value or "").strip().startswith(target_month):
                rows_to_delete.append(row)
        label = f"{target_month} 월 전체"
    elif target_day:
        for row in range(2, ws.max_row + 1):
            if str(ws.cell(row, 1).value or "").strip() == target_day:
                rows_to_delete.append(row)
        label = f"{target_day} 일자"
    elif target_row:
        if 2 <= target_row <= ws.max_row:
            rows_to_delete.append(target_row)
        sku = str(ws.cell(target_row, 4).value or "").strip() if rows_to_delete else ""
        day = str(ws.cell(target_row, 1).value or "").strip() if rows_to_delete else ""
        label = f"{day} / {sku}"
    else:
        raise ValueError("삭제할 일자 또는 상품 줄을 찾지 못했습니다.")

    for row in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row, 1)
    wb.save(SALES_LEDGER_PATH)
    write_sales_display_workbook()
    backup_sales_files_to_supabase()
    return len(rows_to_delete), label


def zip_dir(source_dir: Path, zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in source_dir.rglob("*"):
            if file_path.is_file():
                zf.write(file_path, file_path.relative_to(source_dir))


def build_message(kind: str, text: str, link: str | None = None) -> str:
    cls = "ok" if kind == "ok" else "err"
    link_html = f'<br><a class="download" href="{html.escape(link)}">결과 ZIP 다운로드</a>' if link else ""
    return f'<div style="height:18px"></div><div class="status {cls}">{html.escape(text)}{link_html}</div>'


class BonnieHandler(BaseHTTPRequestHandler):
    SESSION_COOKIE = "coupang_session"
    COOKIE_SECURE = os.environ.get("COOKIE_SECURE", "").strip() == "1"

    def get_cookie(self, name: str) -> str | None:
        for part in self.headers.get("Cookie", "").split(";"):
            key, _, value = part.strip().partition("=")
            if key == name:
                return value
        return None

    def current_user(self) -> dict[str, object] | None:
        return AUTH.get_user(self.get_cookie(self.SESSION_COOKIE))

    def read_urlencoded_form(self) -> dict[str, str]:
        length = int(self.headers.get("Content-Length", "0") or "0")
        body = self.rfile.read(length).decode("utf-8")
        parsed = parse_qs(body, keep_blank_values=True)
        return {key: values[-1] if values else "" for key, values in parsed.items()}

    def send_redirect(self, location: str, cookies: list[str] | None = None) -> None:
        self.send_response(302)
        self.send_header("Location", location)
        for cookie in cookies or []:
            self.send_header("Set-Cookie", cookie)
        self.end_headers()

    def session_cookie(self, token: str) -> str:
        secure = "; Secure" if self.COOKIE_SECURE else ""
        return f"{self.SESSION_COOKIE}={token}; Path=/; HttpOnly; SameSite=Lax{secure}; Max-Age=28800"

    def clear_session_cookie(self) -> str:
        secure = "; Secure" if self.COOKIE_SECURE else ""
        return f"{self.SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Lax{secure}; Max-Age=0"

    def require_user(self) -> dict[str, object] | None:
        user = self.current_user()
        if user is None:
            self.send_redirect("/login")
            return None
        return user

    def require_permission(self, permission: str) -> dict[str, object] | None:
        user = self.require_user()
        if user is None:
            return None
        if not AUTH.has_permission(user, permission):
            self.send_html(login_page("접근 권한이 없습니다."), status=403)
            return None
        return user

    def decorate_page(self, text: str) -> str:
        user = self.current_user()
        if user is None:
            return text
        admin_link = ""
        if AUTH.has_permission(user, "admin"):
            admin_link = '<div class="nav-item" onclick="location.href=\'/admin\'">관리자모드</div>'
        user_html = (
            '<div style="margin-top:24px;padding:12px;border-top:1px solid rgba(255,255,255,.18);'
            'color:#d9e7f3;font-size:13px;line-height:1.6;">'
            f'{html.escape(str(user["username"]))}<br>'
            '<a href="/logout" style="color:#fff;font-weight:800;text-decoration:none;">로그아웃</a>'
            '</div>'
        )
        return text.replace("</aside>", f"{admin_link}{user_html}</aside>", 1)

    def page(self, message: str = "") -> str:
        saved_master = get_saved_master_path()
        if saved_master:
            master_status = f"현재 저장된 기초자료: {html.escape(saved_master.name)}"
        else:
            master_status = "저장된 기초자료가 없습니다. 처음 1회는 기초자료 엑셀을 선택해주세요."
        return self.decorate_page(HTML_PAGE.replace("{message}", message).replace("{master_status}", master_status))

    def master_page(self, message: str = "") -> str:
        master_path = get_saved_master_path()
        if master_path is None:
            rows = '<tr><td colspan="10">저장된 기초자료가 없습니다. 먼저 PO 변환 화면에서 기초자료를 올려주세요.</td></tr>'
        else:
            rows = render_master_rows(master_path)
        return self.decorate_page(MASTER_PAGE.replace("{message}", message).replace("{rows}", rows))

    def pallet_page(self, message: str = "") -> str:
        saved_master = get_saved_master_path()
        if saved_master:
            master_status = f"현재 저장된 기초자료: {html.escape(saved_master.name)}"
        else:
            master_status = "저장된 기초자료가 없습니다. 먼저 쿠팡 PO 변환 화면에서 기초자료를 1회 올려주세요."
        return self.decorate_page(PALLET_PAGE.replace("{message}", message).replace("{master_status}", master_status))

    def check_page(self, message: str = "", result: str = "", date_from: str = "", date_to: str = "") -> str:
        return self.decorate_page(render_check_page(message, result, date_from, date_to))

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/login":
            self.send_html(login_page())
            return

        if parsed.path == "/logout":
            AUTH.logout(self.get_cookie(self.SESSION_COOKIE))
            self.send_redirect("/login", [self.clear_session_cookie()])
            return

        if parsed.path == "/admin":
            if self.require_permission("admin") is None:
                return
            self.send_html(admin_page(AUTH.list_users(), AUTH.admin_email))
            return

        if parsed.path == "/":
            if self.require_permission("po_convert") is None:
                return
            self.send_html(self.page())
            return

        if parsed.path == "/master":
            if self.require_permission("master") is None:
                return
            self.send_html(self.master_page())
            return

        if parsed.path == "/pallet":
            if self.require_permission("pallet") is None:
                return
            self.send_html(self.pallet_page())
            return

        if parsed.path == "/check":
            if self.require_permission("check") is None:
                return
            self.send_html(self.check_page())
            return

        if parsed.path == "/sales":
            if self.require_permission("sales") is None:
                return
            self.send_html(self.decorate_page(render_sales_page()))
            return

        if parsed.path == "/sales/folders":
            if self.require_permission("sales") is None:
                return
            self.send_html(self.decorate_page(render_sales_page(folder_mode=True)))
            return

        if parsed.path == "/sales/download":
            if self.require_permission("sales") is None:
                return
            if not MONTHLY_SALES_PATH.exists() and SALES_LEDGER_PATH.exists():
                write_sales_display_workbook()
            if not MONTHLY_SALES_PATH.exists():
                self.send_html(render_sales_page(build_message("err", "아직 월매출 엑셀 파일이 없습니다.")), status=404)
                return
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote('월매출_납품상품.xlsx')}")
            self.send_header("Content-Length", str(MONTHLY_SALES_PATH.stat().st_size))
            self.end_headers()
            with MONTHLY_SALES_PATH.open("rb") as f:
                shutil.copyfileobj(f, self.wfile)
            return

        if parsed.path.startswith("/download/"):
            if self.require_user() is None:
                return
            run_id = unquote(parsed.path.removeprefix("/download/"))
            zip_path = RUNS_DIR / run_id / "result.zip"
            if not zip_path.exists():
                self.send_html(self.page(build_message("err", "결과 파일을 찾지 못했습니다.")), status=404)
                return
            self.send_response(200)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote('쿠팡_PO_변환결과.zip')}")
            self.send_header("Content-Length", str(zip_path.stat().st_size))
            self.end_headers()
            with zip_path.open("rb") as f:
                shutil.copyfileobj(f, self.wfile)
            return

        self.send_html(self.page(build_message("err", "없는 화면입니다.")), status=404)

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path == "/login":
            form = self.read_urlencoded_form()
            token = AUTH.login(form.get("username", "").strip(), form.get("password", ""))
            if token is None:
                self.send_html(login_page("아이디 또는 비밀번호가 맞지 않습니다."), status=401)
                return
            self.send_redirect("/", [self.session_cookie(token)])
            return

        if path != "/convert":
            if path == "/admin/users/save":
                if self.require_permission("admin") is None:
                    return
                form = parse_qs(self.rfile.read(int(self.headers.get("Content-Length", "0") or "0")).decode("utf-8"))
                username = form.get("username", [""])[-1].strip()
                password = form.get("password", [""])[-1]
                permissions = form.get("permissions", [])
                try:
                    AUTH.upsert_user(username, password, permissions)
                    self.send_html(admin_page(AUTH.list_users(), AUTH.admin_email, "저장되었습니다."))
                except Exception as exc:
                    self.send_html(admin_page(AUTH.list_users(), AUTH.admin_email, f"저장 중 오류가 났습니다: {exc}"), status=400)
                return
            if path == "/admin/users/delete":
                if self.require_permission("admin") is None:
                    return
                form = self.read_urlencoded_form()
                try:
                    AUTH.delete_user(form.get("username", "").strip())
                    self.send_html(admin_page(AUTH.list_users(), AUTH.admin_email, "삭제되었습니다."))
                except Exception as exc:
                    self.send_html(admin_page(AUTH.list_users(), AUTH.admin_email, f"삭제 중 오류가 났습니다: {exc}"), status=400)
                return
            if path == "/master/save":
                if self.require_permission("master") is None:
                    return
                self.handle_master_save()
                return
            if path == "/master/upload":
                if self.require_permission("po_convert") is None:
                    return
                self.handle_master_upload_request()
                return
            if path == "/sales/upload":
                if self.require_permission("sales") is None:
                    return
                self.handle_sales_upload_request()
                return
            if path == "/sales/save":
                if self.require_permission("sales") is None:
                    return
                self.handle_sales_save_request()
                return
            if path == "/sales/delete":
                if self.require_permission("sales") is None:
                    return
                self.handle_sales_delete_request()
                return
            if path == "/sales/year/save":
                if self.require_permission("sales") is None:
                    return
                self.handle_sales_year_save_request()
                return
            if path == "/pallet/create":
                if self.require_permission("pallet") is None:
                    return
                self.handle_pallet_create_request()
                return
            if path == "/check/run":
                if self.require_permission("check") is None:
                    return
                self.handle_check_request()
                return
            self.send_html(self.page(build_message("err", "없는 요청입니다.")), status=404)
            return

        if self.require_permission("po_convert") is None:
            return
        try:
            message, link = self.handle_convert()
            self.send_html(self.page(build_message("ok", message, link)))
        except Exception as exc:
            self.send_html(self.page(build_message("err", f"처리 중 오류가 났습니다: {exc}")), status=500)

    def handle_sales_upload_request(self) -> None:
        content_type = self.headers.get("Content-Type", "")
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type},
        )
        try:
            message, _ = handle_sales_upload(form)
            self.send_html(render_sales_page(build_message("ok", message)))
        except Exception as exc:
            self.send_html(render_sales_page(build_message("err", f"매출확인용 저장 중 오류가 났습니다: {exc}")), status=500)

    def handle_check_request(self) -> None:
        content_type = self.headers.get("Content-Type", "")
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type},
        )
        date_from = form["date_from"].value.strip() if "date_from" in form else ""
        date_to = form["date_to"].value.strip() if "date_to" in form else ""
        try:
            excel_items = form["simpleworks_file"] if "simpleworks_file" in form else []
            image_items = form["simpleworks_image"] if "simpleworks_image" in form else []
            if not isinstance(excel_items, list):
                excel_items = [excel_items]
            if not isinstance(image_items, list):
                image_items = [image_items]
            excel_items = [item for item in excel_items if getattr(item, "filename", "")]
            image_items = [item for item in image_items if getattr(item, "filename", "")]
            if not excel_items and not image_items:
                raise ValueError("심플웍스 엑셀 또는 캡쳐 이미지 파일을 올려주세요.")
            run_id = uuid.uuid4().hex[:12]
            work_dir = RUNS_DIR / f"check_{run_id}"
            work_dir.mkdir(parents=True, exist_ok=True)
            simpleworks_qty: dict[str, int] = defaultdict(int)
            excel_count = 0
            image_count = 0
            for item in excel_items:
                name = safe_name(item.filename)
                if not name.lower().endswith(".xlsx"):
                    raise ValueError("심플웍스 엑셀은 .xlsx 파일로 올려주세요.")
                file_path = work_dir / name
                file_path.write_bytes(item.file.read())
                parsed_qty = parse_simpleworks_excel(file_path)
                for simple_no, qty in parsed_qty.items():
                    simpleworks_qty[simple_no] += qty
                excel_count += 1
            for item in image_items:
                name = safe_name(item.filename)
                if not name.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".webp")):
                    raise ValueError("캡쳐 이미지는 png, jpg, jpeg, bmp, webp 파일로 올려주세요.")
                file_path = work_dir / name
                file_path.write_bytes(item.file.read())
                parsed_qty = parse_simpleworks_image(file_path)
                for simple_no, qty in parsed_qty.items():
                    simpleworks_qty[simple_no] += qty
                image_count += 1
            simpleworks_qty = dict(simpleworks_qty)
            if not simpleworks_qty:
                raise ValueError("올린 파일에서 심플웍스 No와 수량을 찾지 못했습니다.")
            result = render_check_result(simpleworks_qty, date_from, date_to)
            msg = build_message("ok", f"검수 완료: 심플웍스 엑셀 {excel_count:,}개, 캡쳐 이미지 {image_count:,}개에서 심플웍스 No {len(simpleworks_qty):,}개를 합산했습니다.")
            self.send_html(self.check_page(msg, result, date_from, date_to))
        except Exception as exc:
            self.send_html(self.check_page(build_message("err", f"수량검수 중 오류가 났습니다: {exc}"), "", date_from, date_to), status=500)

    def handle_sales_save_request(self) -> None:
        content_type = self.headers.get("Content-Type", "")
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type},
        )
        try:
            changed = save_sales_detail_form(form)
            self.send_html(render_sales_page(build_message("ok", f"수량/메모 수정사항 {changed}건을 저장했습니다.")))
        except Exception as exc:
            self.send_html(render_sales_page(build_message("err", f"수량/메모 저장 중 오류가 났습니다: {exc}")), status=500)

    def handle_sales_delete_request(self) -> None:
        content_type = self.headers.get("Content-Type", "")
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type},
        )
        try:
            deleted, label = delete_sales_detail_form(form)
            self.send_html(render_sales_page(build_message("ok", f"{label} 삭제 완료: {deleted}줄을 삭제했습니다.")))
        except Exception as exc:
            self.send_html(render_sales_page(build_message("err", f"삭제 중 오류가 났습니다: {exc}")), status=500)

    def handle_sales_year_save_request(self) -> None:
        content_type = self.headers.get("Content-Type", "")
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type},
        )
        try:
            changed = save_year_manual(form)
            self.send_html(render_sales_page(build_message("ok", f"연도총매출 수기 입력값 {changed}건을 저장했습니다.")))
        except Exception as exc:
            self.send_html(render_sales_page(build_message("err", f"연도총매출 저장 중 오류가 났습니다: {exc}")), status=500)

    def handle_master_save(self) -> None:
        content_type = self.headers.get("Content-Type", "")
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type},
        )
        try:
            changed = save_master_form(form)
            self.send_html(self.master_page(build_message("ok", f"저장되었습니다. 변경 {changed}건을 반영했습니다.")))
        except Exception as exc:
            self.send_html(self.master_page(build_message("err", f"저장 중 오류가 났습니다: {exc}")), status=500)

    def handle_master_upload_request(self) -> None:
        content_type = self.headers.get("Content-Type", "")
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type},
        )
        next_page = form["next"].value.strip() if "next" in form else ""
        try:
            item = form["master"] if "master" in form else None
            if item is None or not getattr(item, "filename", ""):
                raise ValueError("기초자료 엑셀 파일을 선택해주세요.")
            name = safe_name(item.filename)
            if name.startswith("~$") or not name.lower().endswith(".xlsx"):
                raise ValueError("기초자료는 .xlsx 파일로 올려주세요.")

            master_bytes = item.file.read()
            fd, check_name = tempfile.mkstemp(suffix=".xlsx")
            os.close(fd)
            check_path = Path(check_name)
            try:
                check_path.write_bytes(master_bytes)
                read_master(check_path)
            finally:
                check_path.unlink(missing_ok=True)

            replace_saved_master_file(name, master_bytes)
            message = build_message("ok", f"기초자료를 교체 저장했습니다: {name}")
            if next_page == "master":
                self.send_html(self.master_page(message))
            else:
                self.send_html(self.page(message))
        except Exception as exc:
            message = build_message("err", f"기초자료 저장 중 오류가 났습니다: {exc}")
            if next_page == "master":
                self.send_html(self.master_page(message), status=500)
            else:
                self.send_html(self.page(message), status=500)

    def handle_pallet_create_request(self) -> None:
        try:
            message, link = self.handle_pallet_create()
            self.send_html(self.pallet_page(build_message("ok", message, link)))
        except Exception as exc:
            self.send_html(self.pallet_page(build_message("err", f"파렛트/쉽먼트 초안 생성 중 오류가 났습니다: {exc}")), status=500)

    def handle_pallet_create(self) -> tuple[str, str]:
        content_type = self.headers.get("Content-Type", "")
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type},
        )

        saved_master = get_saved_master_path()
        if saved_master is None:
            raise ValueError("저장된 기초자료가 없습니다. 먼저 쿠팡 PO 변환 화면에서 기초자료를 1회 올려주세요.")

        po_items = form["po_files"] if "po_files" in form else None
        if po_items is None:
            raise ValueError("PO 엑셀을 올려주세요.")
        if not isinstance(po_items, list):
            po_items = [po_items]

        run_id = uuid.uuid4().hex[:12]
        work_dir = RUNS_DIR / run_id
        input_dir = work_dir / "input_po"
        output_dir = work_dir / "result"
        processed_dir = output_dir / "심플웍스_업로드용_PO복사본"
        input_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        po_paths: list[Path] = []
        for item in po_items:
            if not getattr(item, "filename", ""):
                continue
            name = safe_name(item.filename)
            if name.startswith("~$") or not name.lower().endswith(".xlsx"):
                continue
            po_path = input_dir / name
            po_path.write_bytes(item.file.read())
            po_paths.append(po_path)
        if not po_paths:
            raise ValueError("처리할 PO 엑셀 파일이 없습니다.")

        master = read_master(saved_master)
        all_lines = []
        for po_path in sorted(po_paths):
            output_file = processed_dir / po_path.name.replace(".xlsx", " - 복사본.xlsx")
            all_lines.extend(create_processed_po(po_path, output_file, master))

        filled_amounts = update_master_amounts_from_lines(all_lines)
        write_summary_workbook(all_lines, output_dir / "파렛트_쉽먼트_초안.xlsx")
        zip_path = work_dir / "result.zip"
        zip_dir(output_dir, zip_path)

        total_qty = sum(line.available_qty for line in all_lines)
        pallet_ready = sum(1 for line in all_lines if line.pallet_qty)
        amount_message = f" 기초자료 금액 {filled_amounts}건도 자동으로 채웠습니다." if filled_amounts else ""
        return (
            f"파렛트/쉽먼트 초안을 만들었습니다. PO {len(po_paths)}개, 상품 줄 {len(all_lines)}개, "
            f"총 수량 {total_qty:,}개입니다. 파렛트 기준이 있는 줄은 {pallet_ready}개입니다.{amount_message}"
        ), f"/download/{run_id}"

    def handle_convert(self) -> tuple[str, str]:
        content_type = self.headers.get("Content-Type", "")
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type},
        )

        master_item = form["master"] if "master" in form else None
        po_items = form["po_files"] if "po_files" in form else None
        if po_items is None:
            raise ValueError("PO 엑셀을 올려주세요.")
        if not isinstance(po_items, list):
            po_items = [po_items]

        run_id = uuid.uuid4().hex[:12]
        work_dir = RUNS_DIR / run_id
        input_dir = work_dir / "input_po"
        output_dir = work_dir / "result"
        processed_dir = output_dir / "심플웍스_업로드용_PO복사본"
        input_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        master_path: Path | None = None
        if master_item is not None and getattr(master_item, "filename", ""):
            master_name = safe_name(master_item.filename)
            master_path = work_dir / master_name
            master_bytes = master_item.file.read()
            master_path.write_bytes(master_bytes)
            replace_saved_master_file(master_name, master_bytes)
        else:
            saved_master = get_saved_master_path()
            if saved_master is None:
                raise ValueError("저장된 기초자료가 없습니다. 처음 1회는 기초자료 엑셀을 올려주세요.")
            master_path = saved_master

        po_paths: list[Path] = []
        for item in po_items:
            if not getattr(item, "filename", ""):
                continue
            name = safe_name(item.filename)
            if name.startswith("~$") or not name.lower().endswith(".xlsx"):
                continue
            po_path = input_dir / name
            po_path.write_bytes(item.file.read())
            po_paths.append(po_path)
        if not po_paths:
            raise ValueError("처리할 PO 엑셀 파일이 없습니다.")

        master = read_master(master_path)

        all_lines = []
        for po_path in sorted(po_paths):
            output_file = processed_dir / po_path.name.replace(".xlsx", " - 복사본.xlsx")
            all_lines.extend(create_processed_po(po_path, output_file, master))

        filled_amounts = update_master_amounts_from_lines(all_lines)

        zip_path = work_dir / "result.zip"
        zip_dir(output_dir, zip_path)
        total_qty = sum(line.available_qty for line in all_lines)
        total_amount = sum(line.order_amount for line in all_lines)
        amount_message = f" 기초자료 금액 {filled_amounts}건도 자동으로 채웠습니다." if filled_amounts else ""
        return (
            f"완료되었습니다. PO {len(po_paths)}개, 상품 줄 {len(all_lines)}개를 처리했습니다. "
            f"총 수량 {total_qty:,}개, 합계 금액 {total_amount:,}원입니다.{amount_message}"
        ), f"/download/{run_id}"

    def send_html(self, text: str, status: int = 200) -> None:
        data = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, format: str, *args) -> None:
        return


def main() -> None:
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    port = int(os.environ.get("PORT", "8765"))
    server = ThreadingHTTPServer(("0.0.0.0", port), BonnieHandler)
    print("Bonnie Coupang PO web app")
    print(f"Open: http://127.0.0.1:{port}")
    print(f"LAN:  http://<this-computer-ip>:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()

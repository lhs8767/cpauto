from __future__ import annotations

import html

from .service import PERMISSIONS


def login_page(message: str = "") -> str:
    notice = f'<div class="msg">{html.escape(message)}</div>' if message else ""
    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>쿠팡 업무 시스템 로그인</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ margin:0; min-height:100vh; display:grid; place-items:center; background:#f3f6fb; font-family:"Malgun Gothic", Arial, sans-serif; color:#172033; }}
    .box {{ width:min(420px, calc(100vw - 32px)); background:#fff; border:1px solid #d9e1ee; border-radius:8px; padding:28px; box-shadow:0 18px 45px rgba(16,24,40,.10); }}
    h1 {{ margin:0 0 8px; font-size:24px; }}
    p {{ margin:0 0 22px; color:#667085; font-size:14px; }}
    label {{ display:block; font-weight:700; margin:14px 0 7px; font-size:14px; }}
    input {{ width:100%; border:1px solid #d9e1ee; border-radius:8px; padding:13px; font-size:15px; }}
    button {{ width:100%; margin-top:20px; border:0; border-radius:8px; background:#1f4e79; color:#fff; padding:13px; font-weight:800; font-size:15px; cursor:pointer; }}
    .msg {{ margin-bottom:14px; padding:12px; border:1px solid #fecdca; background:#fef3f2; color:#b42318; border-radius:8px; font-size:14px; }}
  </style>
</head>
<body>
  <form class="box" method="post" action="/login">
    <h1>쿠팡 업무 시스템</h1>
    <p>아이디와 비밀번호로 로그인하세요.</p>
    {notice}
    <label>아이디</label>
    <input name="username" autocomplete="username" required />
    <label>비밀번호</label>
    <input name="password" type="password" autocomplete="current-password" required />
    <button type="submit">로그인</button>
  </form>
</body>
</html>"""


def admin_page(users: dict[str, dict[str, object]], admin_email: str, message: str = "") -> str:
    notice = f'<div class="msg">{html.escape(message)}</div>' if message else ""
    rows = "\n".join(_user_row(username, user, admin_email) for username, user in sorted(users.items()))
    checks = "\n".join(
        f'<label><input type="checkbox" name="permissions" value="{key}"> {html.escape(label)}</label>'
        for key, label in PERMISSIONS.items()
    )
    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>관리자모드</title>
  <style>
    * {{ box-sizing:border-box; }}
    body {{ margin:0; background:#f3f6fb; font-family:"Malgun Gothic", Arial, sans-serif; color:#172033; }}
    main {{ max-width:1100px; margin:0 auto; padding:28px; }}
    .top {{ display:flex; justify-content:space-between; align-items:center; gap:12px; margin-bottom:18px; }}
    h1 {{ margin:0; font-size:24px; }}
    a {{ color:#1f4e79; font-weight:800; text-decoration:none; }}
    section {{ background:#fff; border:1px solid #d9e1ee; border-radius:8px; padding:20px; margin-bottom:18px; box-shadow:0 12px 32px rgba(16,24,40,.08); }}
    table {{ width:100%; border-collapse:collapse; font-size:14px; }}
    th, td {{ border-bottom:1px solid #d9e1ee; padding:10px; text-align:left; vertical-align:top; }}
    th {{ background:#f8fafc; }}
    label {{ display:block; margin:8px 0; }}
    input[type=text], input[type=password] {{ width:100%; border:1px solid #d9e1ee; border-radius:8px; padding:11px; }}
    .grid {{ display:grid; grid-template-columns:1fr 1fr; gap:14px; }}
    .checks {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(180px, 1fr)); gap:2px 14px; margin-top:8px; }}
    button {{ border:0; border-radius:8px; background:#1f4e79; color:#fff; padding:11px 15px; font-weight:800; cursor:pointer; }}
    .danger {{ background:#b42318; }}
    .msg {{ margin-bottom:14px; padding:12px; border:1px solid #abefc6; background:#ecfdf3; color:#027a48; border-radius:8px; }}
  </style>
</head>
<body>
  <main>
    <div class="top">
      <h1>관리자모드</h1>
      <div><a href="/">업무 화면</a> · <a href="/logout">로그아웃</a></div>
    </div>
    {notice}
    <section>
      <h2>사용자 권한</h2>
      <table>
        <thead><tr><th>아이디</th><th>권한</th><th>관리</th></tr></thead>
        <tbody>{rows}</tbody>
      </table>
    </section>
    <section>
      <h2>사용자 추가/수정</h2>
      <form method="post" action="/admin/users/save">
        <div class="grid">
          <label>아이디<input name="username" required></label>
          <label>새 비밀번호<input name="password" type="password"></label>
        </div>
        <div class="checks">{checks}</div>
        <button type="submit">저장</button>
      </form>
    </section>
  </main>
</body>
</html>"""


def _user_row(username: str, user: dict[str, object], admin_email: str) -> str:
    permissions = user.get("permissions", [])
    labels = [PERMISSIONS[p] for p in permissions if p in PERMISSIONS]
    delete = "" if username == admin_email else (
        f'<form method="post" action="/admin/users/delete">'
        f'<input type="hidden" name="username" value="{html.escape(username)}">'
        f'<button class="danger" type="submit">삭제</button></form>'
    )
    return (
        f"<tr><td>{html.escape(username)}</td>"
        f"<td>{html.escape(', '.join(labels) or '권한 없음')}</td>"
        f"<td>{delete}</td></tr>"
    )

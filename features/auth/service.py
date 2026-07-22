from __future__ import annotations

import json
import os
import secrets
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path


PERMISSIONS = {
    "po_convert": "PO 변환",
    "master": "기초자료 관리",
    "sales": "매출/월별납품 관리",
    "check": "수량검수",
    "pallet": "파렛트/쉽먼트",
    "admin": "관리자모드",
}


class AuthManager:
    def __init__(self, data_dir: Path) -> None:
        self.data_dir = data_dir
        self.admin_email = os.environ.get("ADMIN_EMAIL", "hslee@bonie.co.kr").strip().lower()
        self.supabase_url = self._clean_supabase_url(os.environ.get("SUPABASE_URL", ""))
        self.anon_key = os.environ.get("SUPABASE_ANON_KEY", "").strip()
        self.service_key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()
        self.local_mode = not (self.supabase_url and self.anon_key and self.service_key)
        self.local_email = os.environ.get("LOCAL_PREVIEW_EMAIL", "").strip().lower()
        self.local_password = os.environ.get("LOCAL_PREVIEW_PASSWORD", "")
        self.local_sessions: set[str] = set()

    def _clean_supabase_url(self, value: str) -> str:
        url = value.strip().rstrip("/")
        if url and not url.startswith(("http://", "https://")):
            url = f"https://{url}"
        return url

    def login(self, username: str, password: str) -> str | None:
        email = username.strip().lower()
        if self.local_mode:
            if self.local_email and email == self.local_email and secrets.compare_digest(password, self.local_password):
                token = secrets.token_urlsafe(32)
                self.local_sessions.add(token)
                return token
            print("[auth] Local preview login failed", flush=True)
            return None
        try:
            data = self._request(
                "POST",
                "/auth/v1/token?grant_type=password",
                {"email": email, "password": password},
                api_key=self.anon_key,
            )
            token = str(data.get("access_token", "")).strip()
            if not token:
                print("[auth] Supabase login did not return access_token", flush=True)
                return None
            return token
        except Exception as exc:
            print(f"[auth] Supabase login failed for {email}: {exc}", flush=True)
            return None

    def logout(self, token: str | None) -> None:
        if token:
            self.local_sessions.discard(token)
        return

    def get_user(self, token: str | None) -> dict[str, object] | None:
        if not token:
            return None
        if self.local_mode:
            if token not in self.local_sessions:
                return None
            return {
                "username": self.local_email,
                "user_id": "local-preview",
                "permissions": list(PERMISSIONS),
                "access_token": token,
            }
        try:
            data = self._request(
                "GET",
                "/auth/v1/user",
                None,
                api_key=self.anon_key,
                bearer_token=token,
            )
            user_id = str(data.get("id", ""))
            email = str(data.get("email", "")).lower()
            return {
                "username": email,
                "user_id": user_id,
                "permissions": self._permissions_for_user(user_id, email),
                "access_token": token,
            }
        except Exception as exc:
            print(f"[auth] Supabase get_user failed: {exc}", flush=True)
            return None

    def has_permission(self, user: dict[str, object] | None, permission: str) -> bool:
        return bool(user and permission in user.get("permissions", []))

    def list_users(self) -> dict[str, dict[str, object]]:
        users = {self.admin_email: {"permissions": list(PERMISSIONS)}}
        try:
            for auth_user in self._list_supabase_users():
                email = str(auth_user.get("email", "")).lower()
                user_id = str(auth_user.get("id", ""))
                users[email] = {"permissions": self._permissions_for_user(user_id, email)}
        except Exception as exc:
            print(f"[auth] Supabase list_users failed: {exc}", flush=True)
        return users

    def upsert_user(self, username: str, password: str, permissions: list[str]) -> None:
        email = username.strip().lower()
        try:
            user_id = self._ensure_supabase_user(email, password)
            self._replace_permissions(user_id, permissions)
        except Exception as exc:
            print(f"[auth] Supabase upsert_user failed for {email}: {exc}", flush=True)
            raise

    def delete_user(self, username: str) -> None:
        email = username.strip().lower()
        if email == self.admin_email:
            raise ValueError("관리자 계정은 삭제할 수 없습니다.")
        try:
            user_id = self._find_user_id(email)
            if user_id:
                self._request(
                    "DELETE",
                    f"/auth/v1/admin/users/{user_id}",
                    None,
                    api_key=self.service_key,
                    bearer_token=self.service_key,
                    admin=True,
                )
        except Exception as exc:
            print(f"[auth] Supabase delete_user failed for {email}: {exc}", flush=True)
            raise

    def _permissions_for_user(self, user_id: str, email: str) -> list[str]:
        if email == self.admin_email:
            return list(PERMISSIONS)
        rows = self._request(
            "GET",
            f"/rest/v1/user_permissions?select=permission&user_id=eq.{urllib.parse.quote(user_id)}",
            None,
            api_key=self.service_key,
            bearer_token=self.service_key,
            admin=True,
        )
        return [row["permission"] for row in rows if row.get("permission") in PERMISSIONS]

    def _replace_permissions(self, user_id: str, permissions: list[str]) -> None:
        valid = [p for p in permissions if p in PERMISSIONS]
        self._request(
            "DELETE",
            f"/rest/v1/user_permissions?user_id=eq.{urllib.parse.quote(user_id)}",
            None,
            api_key=self.service_key,
            bearer_token=self.service_key,
            admin=True,
        )
        if valid:
            rows = [{"user_id": user_id, "permission": p} for p in valid]
            self._request(
                "POST",
                "/rest/v1/user_permissions",
                rows,
                api_key=self.service_key,
                bearer_token=self.service_key,
                admin=True,
            )

    def _ensure_supabase_user(self, email: str, password: str) -> str:
        user_id = self._find_user_id(email)
        if user_id:
            if password:
                self._request(
                    "PUT",
                    f"/auth/v1/admin/users/{user_id}",
                    {"password": password},
                    api_key=self.service_key,
                    bearer_token=self.service_key,
                    admin=True,
                )
            return user_id
        if not password:
            raise ValueError("새 사용자는 비밀번호가 필요합니다.")
        data = self._request(
            "POST",
            "/auth/v1/admin/users",
            {"email": email, "password": password, "email_confirm": True},
            api_key=self.service_key,
            bearer_token=self.service_key,
            admin=True,
        )
        return str(data.get("id", ""))

    def _find_user_id(self, email: str) -> str:
        for user in self._list_supabase_users():
            if str(user.get("email", "")).lower() == email:
                return str(user.get("id", ""))
        return ""

    def _list_supabase_users(self) -> list[dict[str, object]]:
        data = self._request(
            "GET",
            "/auth/v1/admin/users",
            None,
            api_key=self.service_key,
            bearer_token=self.service_key,
            admin=True,
        )
        if isinstance(data, dict):
            return data.get("users", [])
        return data

    def _request(
        self,
        method: str,
        path: str,
        body: object,
        api_key: str,
        bearer_token: str | None = None,
        admin: bool = False,
    ) -> object:
        url = f"{self.supabase_url}{path}"
        payload = json.dumps(body).encode("utf-8") if body is not None else None
        request = urllib.request.Request(url, data=payload, method=method)
        request.add_header("apikey", api_key)
        request.add_header("Authorization", f"Bearer {bearer_token or api_key}")
        request.add_header("Content-Type", "application/json")
        if admin:
            request.add_header("Prefer", "return=representation")
        try:
            with urllib.request.urlopen(request, timeout=15) as response:
                text = response.read().decode("utf-8")
                return json.loads(text) if text else {}
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            raise RuntimeError(detail or str(exc)) from exc

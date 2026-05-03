import argparse
import base64
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
import hashlib
import http.server
import json
from pathlib import Path
import re
import secrets
import sys
import time
import urllib.parse
import urllib.request
import webbrowser


HOST = "127.0.0.1"
PORT = 8765
REDIRECT_PATH = "/callback"
REDIRECT_URI = f"http://{HOST}:{PORT}{REDIRECT_PATH}"
AUTHORITY = "https://login.microsoftonline.com/common/oauth2/v2.0"
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_CLIENT_ID = "c571b0af-327b-46bd-89cc-1ddc38cf9453"
DEFAULT_PROFILE = "default"
SCOPES = ["openid", "profile", "offline_access", "User.Read", "Mail.Read", "Mail.ReadWrite", "Mail.Send"]


def app_dir() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def safe_profile_name(name: str | None) -> str:
    raw = (name or DEFAULT_PROFILE).strip() or DEFAULT_PROFILE
    return re.sub(r"[^A-Za-z0-9_.-]", "_", raw)


class HttpClient:
    @staticmethod
    def post_form(url: str, form: dict) -> dict:
        data = urllib.parse.urlencode(form).encode("utf-8")
        request = urllib.request.Request(
            url,
            data=data,
            headers={"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"},
            method="POST",
        )
        return HttpClient._open_json(request)

    @staticmethod
    def json(url: str, method: str = "GET", token: str | None = None, body: dict | None = None) -> dict:
        headers = {"Accept": "application/json"}
        if token:
            headers["Authorization"] = f"Bearer {token}"
        data = None
        if body is not None:
            headers["Content-Type"] = "application/json"
            data = json.dumps(body).encode("utf-8")
        request = urllib.request.Request(url, data=data, headers=headers, method=method)
        return HttpClient._open_json(request)

    @staticmethod
    def _open_json(request: urllib.request.Request) -> dict:
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                raw = response.read()
                return json.loads(raw.decode("utf-8")) if raw else {}
        except urllib.error.HTTPError as exc:
            details = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"HTTP {exc.code}: {details}") from exc


class TokenStore:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.tokens_dir = root / "tokens"

    def ensure(self) -> None:
        self.tokens_dir.mkdir(exist_ok=True)

    def path(self, profile: str | None) -> Path:
        return self.tokens_dir / f"{safe_profile_name(profile)}.json"

    def load(self, profile: str | None) -> dict:
        path = self.path(profile)
        if not path.exists():
            return {}
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}

    def save(self, profile: str | None, token: dict) -> None:
        self.ensure()
        self.path(profile).write_text(json.dumps(token, indent=2), encoding="utf-8")

    def profiles(self) -> list[str]:
        self.ensure()
        return sorted(path.stem for path in self.tokens_dir.glob("*.json"))


class OAuthCallbackServer(http.server.HTTPServer):
    auth_code: str | None = None
    auth_error: str | None = None
    returned_state: str | None = None


class OAuthCallbackHandler(http.server.BaseHTTPRequestHandler):
    server: OAuthCallbackServer

    def log_message(self, format: str, *args) -> None:
        return

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        if parsed.path != REDIRECT_PATH:
            self.send_response(404)
            self.end_headers()
            return
        self.server.auth_code = params.get("code", [""])[0]
        self.server.auth_error = params.get("error_description", params.get("error", [""]))[0]
        self.server.returned_state = params.get("state", [""])[0]
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write("<h2>登入完成</h2><p>可以回到終端機。</p>".encode("utf-8"))


class AuthService:
    def __init__(self, token_store: TokenStore, client_id: str = DEFAULT_CLIENT_ID) -> None:
        self.token_store = token_store
        self.client_id = client_id

    def login(self, profile: str | None) -> None:
        verifier, challenge = self._make_pkce_pair()
        oauth_state = secrets.token_urlsafe(24)
        query = urllib.parse.urlencode(
            {
                "client_id": self.client_id,
                "response_type": "code",
                "redirect_uri": REDIRECT_URI,
                "response_mode": "query",
                "scope": " ".join(SCOPES),
                "state": oauth_state,
                "code_challenge": challenge,
                "code_challenge_method": "S256",
                "prompt": "select_account",
            }
        )
        server = OAuthCallbackServer((HOST, PORT), OAuthCallbackHandler)
        print(f"正在開啟 Microsoft 登入頁，profile：{safe_profile_name(profile)}")
        webbrowser.open(f"{AUTHORITY}/authorize?{query}")
        server.handle_request()
        if server.auth_error:
            raise RuntimeError(server.auth_error)
        if not server.auth_code or server.returned_state != oauth_state:
            raise RuntimeError("授權回傳資料不完整。")
        raw_token = HttpClient.post_form(
            f"{AUTHORITY}/token",
            {
                "client_id": self.client_id,
                "grant_type": "authorization_code",
                "code": server.auth_code,
                "redirect_uri": REDIRECT_URI,
                "scope": " ".join(SCOPES),
                "code_verifier": verifier,
            },
        )
        self.token_store.save(profile, self._normalize_token(raw_token, self.client_id))

    def refresh_if_needed(self, profile: str | None) -> dict:
        token = self.token_store.load(profile)
        if not token:
            raise RuntimeError(f"Profile「{safe_profile_name(profile)}」尚未登入。")
        if time.time() < float(token.get("expires_at", 0)) - 90:
            return token
        raw_token = HttpClient.post_form(
            f"{AUTHORITY}/token",
            {
                "client_id": token.get("client_id") or self.client_id,
                "grant_type": "refresh_token",
                "refresh_token": token.get("refresh_token"),
                "scope": " ".join(SCOPES),
            },
        )
        refreshed = self._normalize_token(raw_token, token.get("client_id") or self.client_id)
        if not refreshed.get("refresh_token"):
            refreshed["refresh_token"] = token.get("refresh_token", "")
        self.token_store.save(profile, refreshed)
        return refreshed

    @staticmethod
    def _make_pkce_pair() -> tuple[str, str]:
        verifier = base64.urlsafe_b64encode(secrets.token_bytes(48)).rstrip(b"=").decode("ascii")
        challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode("ascii")).digest()).rstrip(b"=").decode("ascii")
        return verifier, challenge

    @staticmethod
    def _normalize_token(raw: dict, client_id: str) -> dict:
        return {
            "client_id": client_id,
            "access_token": raw["access_token"],
            "refresh_token": raw.get("refresh_token", ""),
            "expires_at": time.time() + int(raw.get("expires_in", 3600)),
        }


class GraphClient:
    def __init__(self, auth: AuthService, profile: str | None) -> None:
        self.auth = auth
        self.profile = profile

    def request(self, path_or_url: str, method: str = "GET", body: dict | None = None) -> dict:
        token = self.auth.refresh_if_needed(self.profile)
        url = path_or_url if path_or_url.startswith("https://") else f"{GRAPH_ROOT}{path_or_url}"
        return HttpClient.json(url, method=method, token=token["access_token"], body=body)

    def me(self) -> dict:
        return self.request("/me?$select=displayName,userPrincipalName,mail")

    def today_messages(self, folder: str) -> list[dict]:
        start_utc, end_utc = DateRange.taipei_today_utc_range()
        base = "/me/messages" if folder == "all" else f"/me/mailFolders/{folder}/messages"
        query = urllib.parse.urlencode(
            {
                "$top": "100",
                "$filter": f"receivedDateTime ge {start_utc} and receivedDateTime lt {end_utc} and isDraft eq false",
                "$orderby": "receivedDateTime desc",
                "$select": "id,subject,from,receivedDateTime,isRead,isDraft,bodyPreview,parentFolderId",
            }
        )
        data = self.request(f"{base}?{query}")
        messages = data.get("value", [])
        while data.get("@odata.nextLink"):
            data = self.request(data["@odata.nextLink"])
            messages.extend(data.get("value", []))
        return messages

    def create_reply_all_draft(self, message_id: str, comment: str) -> dict:
        return self.request(f"/me/messages/{message_id}/createReplyAll", method="POST", body={"comment": comment})

    def update_subject(self, message_id: str, subject: str) -> None:
        self.request(f"/me/messages/{message_id}", method="PATCH", body={"subject": subject})


class DateRange:
    @staticmethod
    def taipei_today_utc_range() -> tuple[str, str]:
        taipei = timezone(timedelta(hours=8))
        now = datetime.now(taipei)
        start = datetime(now.year, now.month, now.day, tzinfo=taipei)
        end = start + timedelta(days=1)
        return start.astimezone(timezone.utc).isoformat().replace("+00:00", "Z"), end.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


@dataclass
class DailyDraftResult:
    first_subject: str
    second_subject: str
    extracted_time: str
    draft_subject: str
    comment: str
    draft_id: str | None = None


class DailyDraftConfig:
    def __init__(self, data: dict) -> None:
        self.data = data

    @classmethod
    def load(cls, path: str | None) -> "DailyDraftConfig":
        config_path = Path(path) if path else app_dir() / "daily_draft_config.json"
        if not config_path.exists():
            raise RuntimeError(f"找不到設定檔：{config_path}")
        return cls(json.loads(config_path.read_text(encoding="utf-8")))

    @property
    def first_contains(self) -> str:
        return self.data["firstMessage"]["subjectContains"]

    @property
    def second_contains(self) -> str:
        return self.data["secondMessage"]["subjectContains"]

    @property
    def time_pattern(self) -> str:
        return self.data["timeExtraction"]["pattern"]

    @property
    def draft_rule(self) -> dict:
        return self.data["draft"]


class DailyDraftService:
    def __init__(self, graph: GraphClient, config: DailyDraftConfig) -> None:
        self.graph = graph
        self.config = config

    def build(self, folder: str, body_override: str | None = None, dry_run: bool = False) -> DailyDraftResult:
        messages = self.graph.today_messages(folder)
        first = self._find_latest(messages, self.config.first_contains, require_reply_all=True)
        second = self._find_latest(messages, self.config.second_contains)
        extracted_time = self._extract_time(second.get("subject") or "")
        comment = self._render_body(extracted_time, body_override)
        draft_subject = self._render_subject(first.get("subject") or "")
        result = DailyDraftResult(
            first_subject=first.get("subject") or "",
            second_subject=second.get("subject") or "",
            extracted_time=extracted_time,
            draft_subject=draft_subject,
            comment=comment,
        )
        if dry_run:
            return result
        if not self._can_create_reply_all(first):
            raise RuntimeError(self._reply_all_error(first))
        try:
            draft = self.graph.create_reply_all_draft(first["id"], comment)
        except RuntimeError as exc:
            raise RuntimeError(f"{self._reply_all_error(first)} Original error: {exc}") from exc
        self.graph.update_subject(draft["id"], draft_subject)
        result.draft_id = draft.get("id")
        return result

    def _find_latest(self, messages: list[dict], contains: str, require_reply_all: bool = False) -> dict:
        matches = [item for item in messages if contains in (item.get("subject") or "")]
        if require_reply_all:
            matches = [item for item in matches if self._can_create_reply_all(item)]
        if not matches:
            raise RuntimeError(f"今天找不到主旨包含「{contains}」的信件。")
        return sorted(matches, key=lambda item: item.get("receivedDateTime") or "", reverse=True)[0]

    @staticmethod
    def _can_create_reply_all(message: dict) -> bool:
        return bool(message.get("id")) and not message.get("isDraft") and bool(message.get("receivedDateTime"))

    @staticmethod
    def _reply_all_error(message: dict) -> str:
        subject = message.get("subject") or "(no subject)"
        message_id = message.get("id") or "(no id)"
        received = message.get("receivedDateTime") or "(no receivedDateTime)"
        is_draft = message.get("isDraft")
        return (
            "Cannot create ReplyAll draft for the selected Outlook item. "
            f"subject={subject!r}, id={message_id}, receivedDateTime={received}, isDraft={is_draft}."
        )

    def _extract_time(self, subject: str) -> str:
        match = re.search(self.config.time_pattern, subject)
        if not match:
            raise RuntimeError("第二封信件主旨中找不到指定的時間格式。")
        return match.group(1)

    def _render_body(self, extracted_time: str, body_override: str | None) -> str:
        template = body_override if body_override is not None else self.config.draft_rule.get("bodyTemplate", "{time}")
        return template.replace("{time}", extracted_time)

    def _render_subject(self, original_subject: str) -> str:
        rule = self.config.draft_rule.get("subjectRewrite", {})
        subject_body = original_subject
        if rule.get("pattern") and rule.get("replacement") is not None:
            subject_body = re.sub(rule["pattern"], rule["replacement"], original_subject, count=1)
        return f"{self.config.draft_rule.get('replyPrefix', 'RE: ')}{subject_body}"


class CliApp:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.token_store = TokenStore(root)

    def run(self, argv: list[str] | None = None) -> int:
        self.token_store.ensure()
        argv = sys.argv[1:] if argv is None else argv
        if not argv:
            argv = ["run"]
        parser = self._build_parser()
        args = parser.parse_args(argv)
        code = 0
        try:
            args.func(args)
        except Exception as exc:
            print(f"錯誤：{exc}", file=sys.stderr)
            code = 1
        finally:
            if getattr(sys, "frozen", False) and not args.no_pause:
                input("按 Enter 結束...")
        return code

    def login(self, args: argparse.Namespace) -> None:
        AuthService(self.token_store, args.client_id).login(args.profile)
        graph = self._graph(args.profile, args.client_id)
        me = graph.me()
        print(f"登入完成：{me.get('displayName') or me.get('userPrincipalName') or me.get('mail')}")

    def profiles(self, _args: argparse.Namespace) -> None:
        names = self.token_store.profiles()
        if not names:
            print("目前沒有任何已登入 profile。")
            return
        for name in names:
            print(name)

    def status(self, args: argparse.Namespace) -> None:
        me = self._graph(args.profile).me()
        print(f"Profile：{safe_profile_name(args.profile)}")
        print(f"名稱：{me.get('displayName')}")
        print(f"帳號：{me.get('userPrincipalName') or me.get('mail')}")

    def daily_draft(self, args: argparse.Namespace) -> None:
        result = self._draft_service(args.profile, args.config).build(args.folder, args.body, args.dry_run)
        self._print_draft_result(args.profile, result)
        if args.dry_run:
            print("Dry run，不建立草稿。")
        else:
            print("已建立全部回覆草稿，尚未寄出。")

    def one_click(self, args: argparse.Namespace) -> None:
        selected = self._choose_profile()
        if not self.token_store.load(selected):
            print("此 profile 尚未登入，先開啟登入流程。")
            AuthService(self.token_store, args.client_id).login(selected)
        result = self._draft_service(selected, args.config).build(args.folder, args.body, args.dry_run)
        self._print_draft_result(selected, result)
        if args.dry_run:
            print("Dry run，不建立草稿。")
        else:
            print("已建立全部回覆草稿，尚未寄出。")

    def _graph(self, profile: str | None, client_id: str = DEFAULT_CLIENT_ID) -> GraphClient:
        return GraphClient(AuthService(self.token_store, client_id), profile)

    def _draft_service(self, profile: str | None, config_path: str | None) -> DailyDraftService:
        return DailyDraftService(self._graph(profile), DailyDraftConfig.load(config_path))

    def _choose_profile(self) -> str:
        names = self.token_store.profiles()
        print("請選擇帳號 profile：")
        for index, name in enumerate(names, start=1):
            print(f"{index}. {name}")
        print("N. 新增/重新登入 profile")
        choice = input("請輸入編號、profile 名稱，或直接 Enter 使用 default：").strip()
        if not choice:
            return DEFAULT_PROFILE
        if choice.lower() == "n":
            return safe_profile_name(input("請輸入 profile 名稱，直接 Enter 使用 default：").strip())
        if choice.isdigit() and 1 <= int(choice) <= len(names):
            return names[int(choice) - 1]
        return safe_profile_name(choice)

    @staticmethod
    def _print_draft_result(profile: str | None, result: DailyDraftResult) -> None:
        print(f"Profile：{safe_profile_name(profile)}")
        print(f"第一封：{result.first_subject}")
        print(f"第二封：{result.second_subject}")
        print(f"抽出時間：{result.extracted_time}")
        print(f"草稿主旨：{result.draft_subject}")
        print(f"草稿內文：{result.comment}")
        if result.draft_id:
            print(f"草稿 ID：{result.draft_id}")

    def _build_parser(self) -> argparse.ArgumentParser:
        parser = argparse.ArgumentParser(description="Outlook Graph CLI")
        parser.add_argument("--no-pause", action="store_true")
        sub = parser.add_subparsers(dest="command", required=True)

        login_parser = sub.add_parser("login")
        self._add_profile(login_parser)
        login_parser.add_argument("--client-id", default=DEFAULT_CLIENT_ID)
        login_parser.set_defaults(func=self.login)

        profiles_parser = sub.add_parser("profiles")
        profiles_parser.set_defaults(func=self.profiles)

        status_parser = sub.add_parser("status")
        self._add_profile(status_parser)
        status_parser.set_defaults(func=self.status)

        draft_parser = sub.add_parser("daily-draft")
        self._add_profile(draft_parser)
        self._add_draft_args(draft_parser)
        draft_parser.set_defaults(func=self.daily_draft)

        run_parser = sub.add_parser("run")
        run_parser.add_argument("--client-id", default=DEFAULT_CLIENT_ID)
        self._add_draft_args(run_parser)
        run_parser.set_defaults(func=self.one_click)
        return parser

    @staticmethod
    def _add_profile(parser: argparse.ArgumentParser) -> None:
        parser.add_argument("--profile", default=DEFAULT_PROFILE)

    @staticmethod
    def _add_draft_args(parser: argparse.ArgumentParser) -> None:
        parser.add_argument("--config", default=str(app_dir() / "daily_draft_config.json"))
        parser.add_argument("--body", default=None)
        parser.add_argument("--folder", default="inbox", choices=["inbox", "all"])
        parser.add_argument("--dry-run", action="store_true")


def main(argv: list[str] | None = None) -> int:
    return CliApp(app_dir()).run(argv)


if __name__ == "__main__":
    raise SystemExit(main())

import argparse
from dataclasses import dataclass
from datetime import datetime, time
import html
import json
from pathlib import Path
import re
import sys


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = app_dir()
DEFAULT_CONFIG_PATH = APP_DIR / "daily_draft_config.json"
OL_FOLDER_INBOX = 6
OL_FORMAT_HTML = 2


class OutlookComUnavailable(RuntimeError):
    pass


class Config:
    def __init__(self, data: dict) -> None:
        self.data = data

    @classmethod
    def load(cls, path: str | None) -> "Config":
        config_path = Path(path) if path else DEFAULT_CONFIG_PATH
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


class OutlookApp:
    def __init__(self) -> None:
        try:
            import win32com.client
        except ModuleNotFoundError as exc:
            raise OutlookComUnavailable("缺少 pywin32，請先執行：python -m pip install pywin32") from exc
        self._win32 = win32com.client
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.app.GetNamespace("MAPI")

    def inbox_items(self, mailbox: str | None = None):
        if mailbox:
            recipient = self.namespace.CreateRecipient(mailbox)
            recipient.Resolve()
            if not recipient.Resolved:
                raise RuntimeError(f"找不到或無法解析信箱：{mailbox}")
            folder = self.namespace.GetSharedDefaultFolder(recipient, OL_FOLDER_INBOX)
        else:
            folder = self.namespace.GetDefaultFolder(OL_FOLDER_INBOX)
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        return items


class MailFinder:
    def __init__(self, items) -> None:
        self.items = items

    def latest_today_by_subject_contains(self, subject_contains: str):
        today = datetime.now().date()
        for item in self.items:
            if getattr(item, "Class", None) != 43:
                continue
            received = getattr(item, "ReceivedTime", None)
            subject = getattr(item, "Subject", "") or ""
            if not received or received.date() != today:
                if received and received.date() < today:
                    break
                continue
            if subject_contains in subject:
                return item
        raise RuntimeError(f"今天找不到主旨包含「{subject_contains}」的信件。")


@dataclass
class DraftPreview:
    first_subject: str
    second_subject: str
    extracted_time: str
    draft_subject: str
    body: str


class DailyDraftService:
    def __init__(self, config: Config) -> None:
        self.config = config

    def prepare(self, finder: MailFinder, body_override: str | None = None) -> tuple[object, DraftPreview]:
        first = finder.latest_today_by_subject_contains(self.config.first_contains)
        second = finder.latest_today_by_subject_contains(self.config.second_contains)
        extracted_time = self._extract_time(second.Subject)
        body_template = body_override if body_override is not None else self.config.draft_rule.get("bodyTemplate", "{time}")
        body = body_template.replace("{time}", extracted_time)
        draft_subject = self._rewrite_subject(first.Subject)
        preview = DraftPreview(
            first_subject=first.Subject,
            second_subject=second.Subject,
            extracted_time=extracted_time,
            draft_subject=draft_subject,
            body=body,
        )
        return first, preview

    def create_reply_all_draft(self, source_mail, preview: DraftPreview, display: bool = True):
        draft = source_mail.ReplyAll()
        draft.Subject = preview.draft_subject
        html_body = getattr(draft, "HTMLBody", "") or ""
        if html_body:
            draft.BodyFormat = OL_FORMAT_HTML
            source_html = getattr(source_mail, "HTMLBody", "") or ""
            history_html = self._ensure_source_body_in_history(html_body, source_html)
            draft.HTMLBody = self._prepend_html(history_html, f"{self._body_as_html(preview.body)}<br><br>")
        else:
            draft.Body = f"{preview.body}\r\n\r\n{draft.Body}"
        draft.Save()
        if display:
            draft.Display()
        return draft

    @staticmethod
    def _body_as_html(body: str) -> str:
        return "<br>".join(html.escape(line) for line in body.splitlines())

    @classmethod
    def _ensure_source_body_in_history(cls, reply_html: str, source_html: str) -> str:
        source_body = cls._extract_html_body(source_html).strip()
        if not source_body:
            return reply_html
        source_text = cls._html_text(source_body)
        reply_text = cls._html_text(reply_html)
        if source_text and source_text in reply_text:
            return reply_html
        return cls._append_html(reply_html, f"<br><br>{source_body}")

    @staticmethod
    def _extract_html_body(html_text: str) -> str:
        match = re.search(r"<body\b[^>]*>(.*?)</body>", html_text, flags=re.IGNORECASE | re.DOTALL)
        return match.group(1) if match else html_text

    @staticmethod
    def _html_text(html_text: str) -> str:
        text = re.sub(r"<[^>]+>", "", html_text)
        text = html.unescape(text)
        return re.sub(r"\s+", " ", text).strip()

    @staticmethod
    def _prepend_html(html_text: str, prefix: str) -> str:
        match = re.search(r"<body\b[^>]*>", html_text, flags=re.IGNORECASE)
        if match:
            return html_text[: match.end()] + prefix + html_text[match.end() :]
        return prefix + html_text

    @staticmethod
    def _append_html(html_text: str, suffix: str) -> str:
        match = re.search(r"</body>", html_text, flags=re.IGNORECASE)
        if match:
            return html_text[: match.start()] + suffix + html_text[match.start() :]
        return html_text + suffix

    def _extract_time(self, subject: str) -> str:
        match = re.search(self.config.time_pattern, subject)
        if not match:
            raise RuntimeError("第二封信件主旨中找不到指定的時間格式。")
        return match.group(1)

    def _rewrite_subject(self, subject: str) -> str:
        rule = self.config.draft_rule.get("subjectRewrite", {})
        subject_body = subject
        if rule.get("pattern") and rule.get("replacement") is not None:
            subject_body = re.sub(rule["pattern"], rule["replacement"], subject, count=1)
        return f"{self.config.draft_rule.get('replyPrefix', 'RE: ')}{subject_body}"


class Cli:
    def build_parser(self) -> argparse.ArgumentParser:
        parser = argparse.ArgumentParser(description="Classic Outlook COM daily draft tool")
        parser.add_argument("--config", default=str(DEFAULT_CONFIG_PATH))
        parser.add_argument("--mailbox", default=None, help="指定信箱 SMTP；不填則使用目前 Outlook 預設收件匣")
        parser.add_argument("--body", default=None, help="覆蓋設定檔內文，可用 {time}")
        parser.add_argument("--dry-run", action="store_true", help="只預覽，不建立草稿")
        parser.add_argument("--no-display", action="store_true", help="建立草稿但不開啟草稿視窗")
        return parser

    def run(self, argv: list[str] | None = None) -> int:
        parser = self.build_parser()
        args = parser.parse_args(argv)
        try:
            config = Config.load(args.config)
            outlook = OutlookApp()
            finder = MailFinder(outlook.inbox_items(args.mailbox))
            service = DailyDraftService(config)
            source_mail, preview = service.prepare(finder, args.body)
            self.print_preview(preview)
            if args.dry_run:
                print("Dry run，不建立草稿。")
                return 0
            service.create_reply_all_draft(source_mail, preview, display=not args.no_display)
            print("已建立全部回覆草稿，尚未寄出。")
            return 0
        except Exception as exc:
            print(f"錯誤：{exc}", file=sys.stderr)
            return 1

    @staticmethod
    def print_preview(preview: DraftPreview) -> None:
        print(f"第一封：{preview.first_subject}")
        print(f"第二封：{preview.second_subject}")
        print(f"抽出時間：{preview.extracted_time}")
        print(f"草稿主旨：{preview.draft_subject}")
        print(f"草稿內文：{preview.body}")


if __name__ == "__main__":
    raise SystemExit(Cli().run())

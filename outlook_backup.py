import json
import logging
import re
import threading
import tkinter as tk
from datetime import datetime, timezone
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List

import pythoncom
import win32com.client  # type: ignore
from bs4 import BeautifulSoup
from extract_msg import Message

DEFAULT_CONFIG_PATH = Path("backup_config.json")
DEFAULT_STATE_PATH = Path("backup_state.json")
DEFAULT_LOG_PATH = Path("logs") / "outlook_backup.log"


INTERVAL_PRESETS = {
    "15分": 15,
    "30分": 30,
    "1時間": 60,
    "6時間": 360,
    "12時間": 720,
    "1日": 1440,
    "1週間": 10080,
    "1か月(30日)": 43200,
    "カスタム(分)": None,
}


class OutlookBackupApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Outlook メール自動バックアップ")

        self.config_path = DEFAULT_CONFIG_PATH
        self.state_path = DEFAULT_STATE_PATH
        self.state = self._load_state()

        self.timer_id = None
        self.running_job = False

        self.account_var = tk.StringVar(value="*")
        self.sender_var = tk.StringVar()
        self.save_root_var = tk.StringVar()
        self.interval_choice_var = tk.StringVar(value="1日")
        self.custom_interval_var = tk.StringVar(value="30")

        self._build_ui()
        self._toggle_custom_interval()
        self._load_or_initialize_config()

    def _build_ui(self) -> None:
        tk.Label(self.root, text="対象Outlookアカウント（*で全件）").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        tk.Entry(self.root, width=48, textvariable=self.account_var).grid(row=0, column=1, padx=6)

        tk.Label(self.root, text="送信者メールアドレス（部分一致）").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        tk.Entry(self.root, width=48, textvariable=self.sender_var).grid(row=1, column=1, padx=6)

        tk.Label(self.root, text="保存先フォルダ").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        tk.Entry(self.root, width=48, textvariable=self.save_root_var).grid(row=2, column=1, padx=6)
        tk.Button(self.root, text="参照", command=self._browse_folder).grid(row=2, column=2, padx=6)

        tk.Label(self.root, text="バックアップ間隔").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        self.interval_combo = ttk.Combobox(
            self.root,
            width=18,
            state="readonly",
            textvariable=self.interval_choice_var,
            values=list(INTERVAL_PRESETS.keys()),
        )
        self.interval_combo.grid(row=3, column=1, sticky="w", padx=6)
        self.interval_combo.bind("<<ComboboxSelected>>", lambda _e: self._toggle_custom_interval())

        tk.Label(self.root, text="カスタム分").grid(row=3, column=1, padx=(170, 6), sticky="w")
        self.custom_interval_entry = tk.Entry(self.root, width=10, textvariable=self.custom_interval_var)
        self.custom_interval_entry.grid(row=3, column=1, padx=(230, 6), sticky="w")

        tk.Button(self.root, text="保存して自動開始", command=self.save_and_start, width=18).grid(row=4, column=0, pady=14)
        tk.Button(self.root, text="今すぐ1回実行", command=self.run_once_async, width=18).grid(row=4, column=1, sticky="w", pady=14)
        tk.Button(self.root, text="自動実行を停止", command=self.stop_schedule, width=18).grid(row=4, column=2, pady=14)

        self.status_label = tk.Label(self.root, text="状態: 初期化中", anchor="w")
        self.status_label.grid(row=5, column=0, columnspan=3, sticky="w", padx=6, pady=6)

    def _load_or_initialize_config(self) -> None:
        if not self.config_path.exists():
            self.status_label.config(text="状態: 初回設定を入力して「保存して自動開始」を押してください")
            return

        config = json.loads(self.config_path.read_text(encoding="utf-8"))
        self.account_var.set(config.get("target_account", "*"))
        self.sender_var.set(config.get("sender_filter", ""))
        self.save_root_var.set(config.get("save_root", ""))
        self.interval_choice_var.set(config.get("interval_label", "1日"))
        self.custom_interval_var.set(str(config.get("custom_interval_minutes", 30)))
        self._toggle_custom_interval()

        self.status_label.config(text="状態: 設定読込済み。自動実行を開始します")
        self.start_schedule()


    def _toggle_custom_interval(self) -> None:
        if self.interval_choice_var.get() == "カスタム(分)":
            self.custom_interval_entry.configure(state="normal")
        else:
            self.custom_interval_entry.configure(state="disabled")

    def _get_interval_minutes(self) -> int:
        label = self.interval_choice_var.get().strip() or "1日"
        value = INTERVAL_PRESETS.get(label)
        if value is None:
            custom = int(self.custom_interval_var.get().strip())
            if custom <= 0:
                raise ValueError("custom interval")
            return custom
        return int(value)

    def _browse_folder(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.save_root_var.set(folder)

    def _load_state(self) -> Dict[str, str]:
        if not self.state_path.exists():
            return {}
        try:
            return json.loads(self.state_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_state(self) -> None:
        self.state_path.write_text(json.dumps(self.state, ensure_ascii=False, indent=2), encoding="utf-8")

    def _validate_inputs(self) -> bool:
        sender = self.sender_var.get().strip()
        save_root = self.save_root_var.get().strip()

        if not sender or not save_root:
            messagebox.showerror("エラー", "送信者メールアドレス・保存先を入力してください")
            return False

        try:
            mins = self._get_interval_minutes()
            if mins <= 0:
                raise ValueError("interval")
        except Exception:
            messagebox.showerror("エラー", "バックアップ間隔の設定を見直してください")
            return False

        Path(save_root).mkdir(parents=True, exist_ok=True)
        return True

    def save_and_start(self) -> None:
        if not self._validate_inputs():
            return

        config = {
            "target_account": self.account_var.get().strip() or "*",
            "sender_filter": self.sender_var.get().strip(),
            "save_root": self.save_root_var.get().strip(),
            "interval_label": self.interval_choice_var.get().strip(),
            "custom_interval_minutes": int(self.custom_interval_var.get().strip() or "30"),
            "interval_minutes": self._get_interval_minutes(),
        }
        self.config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")

        self.status_label.config(text="状態: 設定を保存しました。自動実行を開始します")
        self.start_schedule()

    def start_schedule(self) -> None:
        self.stop_schedule()
        self.run_once_async()

    def stop_schedule(self) -> None:
        if self.timer_id is not None:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None
        self.status_label.config(text="状態: 自動実行停止中")

    def _schedule_next(self) -> None:
        mins = self._get_interval_minutes()
        self.timer_id = self.root.after(mins * 60 * 1000, self.run_once_async)
        self.status_label.config(text=f"状態: 次回実行を {mins} 分後に予約しました")

    def run_once_async(self) -> None:
        if self.running_job:
            return
        if not self._validate_inputs():
            return

        self.running_job = True
        self.status_label.config(text="状態: バックアップ実行中...")

        def _job() -> None:
            try:
                saved = self.run_backup(
                    account_filter=self.account_var.get().strip() or "*",
                    sender_filter=self.sender_var.get().strip(),
                    save_root=self.save_root_var.get().strip(),
                )
                self.root.after(0, lambda: self.status_label.config(text=f"状態: 実行完了（新規 {saved} 件）"))
            except Exception as exc:
                logging.exception("バックアップ処理に失敗しました")
                self.root.after(0, lambda: messagebox.showerror("エラー", str(exc)))
            finally:
                self.running_job = False
                self.root.after(0, self._schedule_next)

        threading.Thread(target=_job, daemon=True).start()

    @staticmethod
    def get_real_sender_smtp(mail) -> str:
        try:
            if getattr(mail, "SentOnBehalfOfName", None):
                try:
                    ex_user = mail.SentOnBehalfOf
                    if ex_user:
                        return ex_user.PrimarySmtpAddress or ""
                except Exception:
                    pass

            if getattr(mail, "SenderEmailType", "") == "EX":
                try:
                    ex_user = mail.Sender.GetExchangeUser()
                    if ex_user:
                        return ex_user.PrimarySmtpAddress or ""
                except Exception:
                    pass

            return getattr(mail, "SenderEmailAddress", "") or ""
        except Exception:
            return ""

    @staticmethod
    def _safe_name(raw: str) -> str:
        cleaned = re.sub(r'[\\/:*?"<>|\r\n]+', "_", raw).strip(" .")
        return cleaned or "(no_subject)"

    @staticmethod
    def _safe_dir(raw: str) -> str:
        cleaned = re.sub(r'[\\/:*?"<>|\r\n]+', "_", raw).strip(" .")
        return cleaned or "unknown_account"

    @staticmethod
    def _get_account_name(account_folder) -> str:
        return str(getattr(account_folder, "Name", "")).strip()

    def _select_accounts(self, namespace, account_filter: str):
        all_accounts = [acct for acct in namespace.Folders]
        if account_filter == "*":
            return all_accounts

        key = account_filter.lower()
        selected = [acct for acct in all_accounts if key in self._get_account_name(acct).lower()]
        return selected

    @staticmethod
    def _find_inbox(account_folder):
        for folder in account_folder.Folders:
            name = str(getattr(folder, "Name", "")).lower()
            if name in {"受信トレイ", "inbox"}:
                return folder
        return None

    def run_backup(self, account_filter: str, sender_filter: str, save_root: str) -> int:
        base_dir = Path(save_root)
        saved_count = 0

        pythoncom.CoInitialize()
        try:
            namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            accounts = self._select_accounts(namespace, account_filter)
            if not accounts:
                raise RuntimeError(f"対象アカウントが見つかりません: {account_filter}")

            for account in accounts:
                account_name = self._get_account_name(account)
                inbox = self._find_inbox(account)
                if inbox is None:
                    logging.warning("受信トレイが見つかりません: %s", account_name)
                    continue

                account_root = base_dir / self._safe_dir(account_name)
                save_msg_dir = account_root / "msg"
                save_attach_dir = account_root / "attachments"
                save_html_dir = account_root / "html"

                save_msg_dir.mkdir(parents=True, exist_ok=True)
                save_attach_dir.mkdir(parents=True, exist_ok=True)
                save_html_dir.mkdir(parents=True, exist_ok=True)

                saved_files: List[Path] = []
                items = inbox.Items
                items.Sort("[ReceivedTime]", True)

                for mail in items:
                    try:
                        if getattr(mail, "Class", None) != 43:
                            continue

                        sender_addr = self.get_real_sender_smtp(mail)
                        if sender_filter.lower() not in sender_addr.lower():
                            continue

                        entry_id = getattr(mail, "EntryID", "")
                        if not entry_id:
                            continue

                        state_key = f"{account_name}:{entry_id}"
                        if state_key in self.state:
                            continue

                        dt = mail.ReceivedTime.strftime("%Y%m%d_%H%M%S")
                        subject = self._safe_name(getattr(mail, "Subject", ""))[:80]
                        filename = f"{dt}_{entry_id[:8]}_{subject}.msg"
                        msg_path = save_msg_dir / filename

                        mail.SaveAs(str(msg_path), 3)
                        saved_files.append(msg_path)
                        self.state[state_key] = datetime.now(timezone.utc).isoformat()
                        saved_count += 1
                    except Exception:
                        logging.exception("MSG保存に失敗: account=%s", account_name)

                self._save_state()

                for msg_path in saved_files:
                    self._extract_attachments_and_html(msg_path, save_attach_dir, save_html_dir)
        finally:
            pythoncom.CoUninitialize()

        return saved_count

    @staticmethod
    def _extract_attachments_and_html(msg_path: Path, attach_dir: Path, html_dir: Path) -> None:
        try:
            msg = Message(str(msg_path))
            body_html = msg.htmlBody or msg.body or ""

            for att in msg.attachments:
                name = att.longFilename or att.shortFilename
                if not name:
                    continue
                out_name = re.sub(r'[\\/:*?"<>|\r\n]+', "_", name)
                with open(attach_dir / out_name, "wb") as f:
                    f.write(att.data)

            soup = BeautifulSoup(body_html, "lxml")
            base = msg_path.stem
            with open(html_dir / f"{base}.html", "w", encoding="utf-8") as f:
                f.write(soup.prettify())
        except Exception:
            logging.exception("MSG解析に失敗: %s", msg_path)


def setup_logging() -> None:
    DEFAULT_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(DEFAULT_LOG_PATH, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def main() -> None:
    setup_logging()
    root = tk.Tk()
    OutlookBackupApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

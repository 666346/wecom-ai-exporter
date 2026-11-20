#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""wecom-agent Windows 本地封装层。

所有聊天读取与导出最终都调用仓库自带命令：
    decrypt/windows/wecom_win.py <key> conversations --limit 2000 --json
    decrypt/windows/wecom_win.py <key> export <cid> --out chat.md --format markdown --with-media --json

本文件只负责：多账号管理、key 有效性校验/自动重抓、会话缓存、按群名模糊匹配、
把每次导出生成的随机 assets 目录归一化为固定 attachments/（实现真正增量合并）、
统计汇总（消息数/成功附件/missing/encrypted）、以及 5 分钟增量同步的指纹判断。
"""
import argparse
import contextlib
import filecmp
import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
from pathlib import Path

try:
    import msvcrt  # Windows 文件锁；非 Windows 下退化为仅进程内互斥
except ImportError:  # pragma: no cover
    msvcrt = None

ROOT = Path(__file__).resolve().parent
CONFIG = ROOT / "config.json"
RUNTIME = ROOT / "runtime"
CACHE = RUNTIME / "cache"
STATE = RUNTIME / "sync_state.json"
WATCH = ROOT / "watch_groups.txt"
WIN_CLI = ROOT / "decrypt" / "windows" / "wecom_win.py"
KEYS_JSON = ROOT / "decrypt" / "windows" / "wxwork_keys.json"
KEY_TXT = ROOT / "decrypt" / "windows" / "key.txt"
FIND_KEYS = ROOT / "find_all_keys.ps1"

sys.path.insert(0, str(ROOT / "decrypt"))
from wxwork_crypto import quick_verify  # noqa: E402
from local_io import connect_readonly  # noqa: E402
import export_wxwork as ex  # noqa: E402

LOCK_DIR = RUNTIME / "locks"
LOCK_FILE = LOCK_DIR / "wecom-db.lock"
_DB_LOCK = threading.RLock()
_DB_DEPTH = threading.local()


@contextlib.contextmanager
def db_lock(timeout=1800):
    """解密库互斥锁（进程内可重入 + 跨进程文件锁）。

    Windows 上 `os.replace(tmp, message.db)` 在目标被别的进程以只读方式打开时会抛
    PermissionError WinError 5；而 Web 导出、内置 5 分钟同步、终端 export_group.bat /
    sync_watch.ps1 都会同时解密同一个账号的库。所有涉及解密库的操作必须排队。
    """
    depth = getattr(_DB_DEPTH, "value", 0)
    with _DB_LOCK:
        fd = None
        if depth == 0:
            LOCK_DIR.mkdir(parents=True, exist_ok=True)
            fd = os.open(str(LOCK_FILE), os.O_RDWR | os.O_CREAT, 0o644)
            if os.fstat(fd).st_size == 0:
                os.write(fd, b"0")
            os.lseek(fd, 0, os.SEEK_SET)
            if msvcrt is not None:
                deadline = time.monotonic() + timeout
                while True:
                    try:
                        msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
                        break
                    except OSError:
                        if time.monotonic() >= deadline:
                            os.close(fd)
                            raise RuntimeError("等待解密库锁超时（可能有另一个导出/同步任务卡住）")
                        time.sleep(0.4)
        _DB_DEPTH.value = depth + 1
        try:
            yield
        finally:
            _DB_DEPTH.value = depth
            if fd is not None:
                if msvcrt is not None:
                    try:
                        os.lseek(fd, 0, os.SEEK_SET)
                        msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
                    except OSError:
                        pass
                os.close(fd)


# ---------------------------------------------------------------- 配置 / key

def load_config():
    if not CONFIG.exists():
        raise RuntimeError(f"缺配置文件 {CONFIG}")
    return json.loads(CONFIG.read_text(encoding="utf-8"))


def save_config(cfg):
    CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def wxwork_root(cfg):
    return Path(cfg["wxwork_root"])


def make_account(account_id, cfg):
    base = wxwork_root(cfg) / account_id
    return {
        "id": account_id,
        "data_dir": str(base / "Data"),
        "decrypted_dir": str(ROOT / "decrypt" / "windows" / "export" / ("dec_" + account_id)),
        "cache_dir": str(base / "Cache"),
        "key": "",
    }


def sync_accounts_from_keys(cfg, keys):
    """把新抓到 key 的账号自动补进 config（例如用户切换企微账号后）。"""
    known = {a["id"] for a in cfg["accounts"]}
    added = []
    for account_id in keys:
        if account_id in known:
            continue
        if not (wxwork_root(cfg) / account_id / "Data" / "message.db").exists():
            continue
        cfg["accounts"].append(make_account(account_id, cfg))
        known.add(account_id)
        added.append(account_id)
    changed = False
    for acc in cfg["accounts"]:
        key = keys.get(acc["id"])
        if key and acc.get("key") != key:
            acc["key"] = key
            changed = True
    if added or changed:
        save_config(cfg)
        if cfg["accounts"]:
            KEY_TXT.write_text(cfg["accounts"][0]["key"], encoding="ascii")  # server.py 需要
    return added


def rescan_keys(cfg):
    """key 失效（企微重启/切换账号）时，重新跑内存扫描。

    只用稳定的 NORMAL 模式：WIDE 模式会扫描全部内存页，在本机 8 账号环境下会崩。
    """
    print("key 已失效，正在重新扫描企业微信内存…", file=sys.stderr)
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(FIND_KEYS)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=240,
    )
    if not KEYS_JSON.exists():
        raise RuntimeError("重新扫描后仍没有 wxwork_keys.json；请确认企业微信已登录并运行")
    keys = json.loads(KEYS_JSON.read_text(encoding="utf-8"))
    added = sync_accounts_from_keys(cfg, keys)
    if added:
        print(f"新增可读取账号: {', '.join(added)}", file=sys.stderr)
    return load_config()


def ensure_key(acc, cfg):
    """返回该账号当前有效的 16 字节 key（hex）；失效则自动重扫后重试一次。"""
    mdb = Path(acc["data_dir"]) / "message.db"
    if not mdb.exists():
        raise RuntimeError(f"账号 {acc['id']} 缺 message.db: {mdb}")
    with mdb.open("rb") as fh:
        page1 = fh.read(4096)
    for attempt in range(2):
        key = acc.get("key", "")
        if key:
            try:
                if quick_verify(bytes.fromhex(key), page1):
                    return key
            except ValueError:
                pass
        cfg = rescan_keys(cfg)
        acc = next((a for a in cfg["accounts"] if a["id"] == acc["id"]), acc)
    raise RuntimeError(
        f"账号 {acc['id']} 的 key 不在企业微信内存中。请打开企业微信并登录/切换到该账号，"
        f"然后重新运行（或先执行 find_all_keys.ps1）。"
    )


# ---------------------------------------------------------------- 调用项目 CLI

def run_cli(cfg, acc, key, *args):
    env = os.environ.copy()
    env["WECOM_DATA_DIR"] = acc["data_dir"]
    env["WECOM_DECRYPTED_DIR"] = acc["decrypted_dir"]
    env["WECOM_CACHE_DIR"] = acc.get("cache_dir") or str(Path(acc["data_dir"]).parent / "Cache")
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    cmd = [cfg["python"], "-X", "utf8", str(WIN_CLI), key, *args]
    with db_lock():
        proc = subprocess.run(cmd, env=env, capture_output=True, text=True,
                              encoding="utf-8", errors="replace", timeout=1800)
    if proc.returncode != 0:
        raise RuntimeError(f"wecom_win.py 执行失败 ({' '.join(args[:2])}):\n{proc.stderr.strip() or proc.stdout.strip()}")
    return proc.stdout


# ---------------------------------------------------------------- 会话缓存 / 匹配

def refresh(acc, cfg, force=False):
    CACHE.mkdir(parents=True, exist_ok=True)
    target = CACHE / f"conv_{acc['id']}.json"
    ttl = int(cfg.get("cache_ttl_seconds", 1800))
    if not force and target.exists() and time.time() - target.stat().st_mtime < ttl:
        try:
            return json.loads(target.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            pass
    key = ensure_key(acc, cfg)
    raw = run_cli(cfg, acc, key, "conversations", "--limit", "2000", "--json")
    data = json.loads(raw)
    payload = {"account": acc["id"], "updated": time.time(),
               "conversations": data.get("conversations", [])}
    target.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return payload


def all_accounts(cfg, force=False):
    for acc in cfg["accounts"]:
        try:
            yield acc, refresh(acc, cfg, force=force)
        except Exception as exc:  # 某个账号 key 不在内存时不影响其它账号
            print(f"跳过账号 {acc['id']}: {exc}", file=sys.stderr)
            yield acc, {"account": acc["id"], "conversations": []}


def match(cfg, keyword, force=False):
    needle = keyword.casefold()
    hits = []
    for acc, data in all_accounts(cfg, force=force):
        for conv in data["conversations"]:
            if needle in str(conv.get("name", "")).casefold():
                hits.append({**conv, "account": acc["id"]})
    return {"keyword": keyword, "count": len(hits), "candidates": hits}


def locate(cfg, cid, force=False):
    for acc, data in all_accounts(cfg, force=force):
        for conv in data["conversations"]:
            if str(conv.get("id")) == str(cid):
                return acc, conv.get("name", cid)
    if not force:
        return locate(cfg, cid, force=True)
    return None, None


# ---------------------------------------------------------------- 导出

def safe_name(name):
    text = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(name))
    text = re.sub(r"\s+", " ", text).strip().strip(".")
    return text[:80] or "未命名会话"


def _merge_assets(outdir, new_dirs, chat, manifest_path):
    """把随机 chat.assets-<id>/ 归一化到固定 attachments/，只复制新增/变化的文件。"""
    fixed = outdir / "attachments"
    replaced = []
    for directory in new_dirs:
        for src in sorted(directory.rglob("*")):
            if not src.is_file():
                continue
            rel = src.relative_to(directory)
            dst = fixed / rel
            if dst.exists() and dst.stat().st_size == src.stat().st_size and filecmp.cmp(dst, src, shallow=False):
                src.unlink()
                continue
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            src.unlink()
        shutil.rmtree(directory, ignore_errors=True)
        replaced.append(directory.name)
    if replaced:
        for path in (chat, manifest_path):
            if path and Path(path).exists():
                text = Path(path).read_text(encoding="utf-8")
                for old in replaced:
                    text = text.replace(old, "attachments")
                Path(path).write_text(text, encoding="utf-8")
    return fixed


def export_one(cfg, acc, cid, name=None):
    key = ensure_key(acc, cfg)
    if not name:
        found_acc, name = locate(cfg, cid)
        if found_acc:
            acc = found_acc
            key = ensure_key(acc, cfg)
    folder = Path(cfg["export_root"]) / safe_name(name or cid)
    marker = folder / ".conversation_id"
    if marker.exists() and marker.read_text(encoding="utf-8").strip() != str(cid):
        folder = Path(cfg["export_root"]) / f"{safe_name(name or cid)}_{str(cid)[-6:]}"
    folder.mkdir(parents=True, exist_ok=True)
    marker.write_text(str(cid), encoding="utf-8")

    chat = folder / "chat.md"
    before = {p.name for p in folder.glob("chat.assets-*")}
    # 仓库自带导出命令：时间正序(ASC)、UTF-8、Markdown、带媒体清单
    raw = run_cli(cfg, acc, key, "export", str(cid),
                  "--out", str(chat), "--format", "markdown", "--with-media", "--json")
    result = json.loads(raw)
    manifest_path = Path(result["manifest"]) if result.get("manifest") else None
    new_dirs = [p for p in folder.glob("chat.assets-*") if p.name not in before]
    if new_dirs:
        _merge_assets(folder, new_dirs, chat, manifest_path)

    manifest = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path and manifest_path.exists() else {"entries": []}
    issues, attachments_ok = [], 0
    for entry in manifest.get("entries", []):
        files = entry.get("files") or []
        if entry.get("status") == "resolved" and files:
            attachments_ok += len(files)
            continue
        status = str(entry.get("status", "unknown"))
        if status in ("not_found", "cache_changed_unreadable_or_unsafe", "cache_changed_or_unreadable", "cache_root_unavailable"):
            category = "missing"
        elif status in ("encrypted_unsupported",):
            category = "encrypted"
        else:
            category = "other"
        msg = entry.get("message") or {}
        issues.append({
            "category": category,
            "status": status,
            "reason": entry.get("reason", ""),
            "kind": entry.get("kind", ""),
            "local_id": msg.get("local_id", ""),
            "hint": "请在企业微信里点开该消息/附件，等待下载缓存完成后重新导出" if category == "missing" else
                    "该文件为加密或无法识别的图片，本地无明文缓存" if category == "encrypted" else
                    "资源标识存在歧义或格式不支持，需人工确认",
        })
    return {
        "account": acc["id"], "conversation_id": str(cid), "name": name,
        "out_dir": str(folder), "chat_md": str(chat),
        "manifest": str(manifest_path) if manifest_path else None,
        "messages": result.get("messages", 0),
        "media_status": result.get("media", {}),
        "attachments_ok": attachments_ok,
        "issues": issues,
        "missing": [i for i in issues if i["category"] == "missing"],
        "encrypted": [i for i in issues if i["category"] == "encrypted"],
        "other": [i for i in issues if i["category"] == "other"],
        "exported_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }


# ---------------------------------------------------------------- 增量同步

def fingerprint(acc):
    data = Path(acc["data_dir"])
    parts = []
    for filename in ("message.db", "message.db-wal", "session.db", "session.db-wal"):
        path = data / filename
        if path.exists():
            st = path.stat()
            parts.append(f"{filename}:{st.st_size}:{st.st_mtime_ns}")
    cache = Path(acc.get("cache_dir") or (data.parent / "Cache"))
    if cache.exists():
        newest, count = 0, 0
        for path in cache.rglob("*"):
            if path.is_file():
                count += 1
                try:
                    newest = max(newest, path.stat().st_mtime_ns)
                except OSError:
                    pass
        parts.append(f"cache:{count}:{newest}")
    return hashlib.sha256("|".join(parts).encode("utf-8")).hexdigest()[:16]


def export_cache(cfg, acc, cid, name=None):
    """兜底：把该账号的全部明文媒体缓存导出到群目录 cache_all/。

    仓库 Windows 端的 --with-media 依赖 macOS 版 protobuf 资源 ID 假设，本机实测
    关联率可能为 0。这个命令改用项目自带 media 子命令，保证附件不丢，
    但文件不再与消息一一对应（适合整体喂给 AI 工具）。一次性使用，体积可能较大。
    """
    key = ensure_key(acc, cfg)
    folder = Path(cfg["export_root"]) / safe_name(name or cid)
    out = folder / "cache_all"
    raw = run_cli(cfg, acc, key, "media", "--out", str(out), "--json")
    return {"account": acc["id"], "conversation_id": str(cid),
            "out": str(out), "result": json.loads(raw)}


def parse_watch():
    entries = []
    if not WATCH.exists():
        return entries
    for line in WATCH.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        cid = line.split()[0]
        entries.append(cid)
    return entries


def sync(cfg, force=False):
    with db_lock():
        return _sync_locked(cfg, force=force)


def _sync_locked(cfg, force=False):
    RUNTIME.mkdir(parents=True, exist_ok=True)
    state = json.loads(STATE.read_text(encoding="utf-8")) if STATE.exists() else {}
    old = state.get("fingerprints", {})
    fps = {}
    for acc in cfg["accounts"]:
        try:
            fps[acc["id"]] = fingerprint(acc)
        except OSError:
            fps[acc["id"]] = ""
    changed = {a["id"] for a in cfg["accounts"] if fps[a["id"]] != old.get(a["id"])}

    cids = parse_watch()
    if not cids:
        return {"skipped": True, "reason": "watch_groups.txt 为空", "changed": sorted(changed)}

    results, errors = [], []
    for cid in cids:
        acc, name = locate(cfg, cid)
        if not acc:
            errors.append(f"{cid}: 未在任何已解锁账号的会话列表中找到")
            continue
        if not force and acc["id"] not in changed:
            results.append({"conversation_id": cid, "skipped": True, "account": acc["id"], "name": name})
            continue
        try:
            results.append(export_one(cfg, acc, cid, name))
        except Exception as exc:
            errors.append(f"{cid}: {exc}")

    state["fingerprints"] = fps
    state["last_sync"] = time.strftime("%Y-%m-%d %H:%M:%S")
    STATE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"skipped": False, "changed": sorted(changed), "results": results, "errors": errors,
            "time": state["last_sync"]}


# ---------------------------------------------------------------- Web 界面支撑层

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff", ".heic"}
DOC_EXTENSIONS = IMAGE_EXTENSIONS | {
    ".doc", ".docx", ".pdf", ".xls", ".xlsx", ".xlsm", ".csv", ".txt", ".md", ".json",
    ".log", ".xml", ".html", ".htm", ".ppt", ".pptx", ".rtf",
    ".dwg", ".dxf", ".skp", ".rvt", ".rfa", ".3ds", ".max", ".obj", ".stl", ".step", ".stp",
    ".psd", ".ai", ".cdr", ".eps", ".svg", ".zip", ".rar", ".7z",
    ".mp4", ".mov", ".avi", ".mp3", ".wav", ".silk", ".amr", ".m4a",
}
MAX_CACHE_FILES = 200
MAX_CACHE_BYTES = 150 * 1024 * 1024
MAX_CACHE_SINGLE_BYTES = 25 * 1024 * 1024
CACHE_WINDOW_DAYS = 30
CACHE_TEXT_EXTENSIONS = {".txt", ".md", ".csv", ".json", ".log", ".xml", ".html", ".htm",
                         ".docx", ".xlsx", ".xlsm", ".pdf"}
MAX_ALL_FOR_AI_CACHE_DOCS = 80
MAX_ALL_FOR_AI_CACHE_CHARS = 800 * 1024


def _norm_ts(value):
    if not value:
        return 0.0
    value = float(value)
    return value / 1000.0 if value > 20000000000 else value


def _dec_message_db(acc, cfg):
    path = Path(acc["decrypted_dir"]) / "message.db"
    if not path.exists():
        refresh(acc, cfg)  # 用项目自带命令触发一次解密，生成 dec 库
    if not path.exists():
        raise RuntimeError(f"解密库不存在: {path}")
    return path


def _load_names(acc):
    users = {}
    path = Path(acc["decrypted_dir"]) / "user.db"
    if not path.exists():
        return users
    con = connect_readonly(path)
    try:
        for row in con.execute("SELECT id, name FROM user_table"):
            if row[1]:
                users[str(row[0])] = row[1]
        try:
            for row in con.execute("SELECT wxid, name FROM wechat_contactV1"):
                if row[1] and row[1] != "-":
                    users.setdefault(str(row[0]), row[1])
        except sqlite3.Error:
            pass
    finally:
        con.close()
    return users


def _rows(acc, cfg, cid, *, descending=False, limit=None):
    db = _dec_message_db(acc, cfg)
    order = "DESC" if descending else "ASC"
    sql = ("SELECT message_id, sender_id, content_type, send_time, content "
           "FROM message_table WHERE conversation_id=? "
           f"ORDER BY send_time {order}, message_id {order}")
    if limit:
        sql += f" LIMIT {int(limit)}"
    with db_lock():
        con = connect_readonly(db)
        try:
            return [tuple(r) for r in con.execute(sql, (cid,))], _load_names(acc)
        finally:
            con.close()


def _printable_ratio(text):
    if not text:
        return 0.0
    return sum(1 for ch in text if ch.isprintable() or ch in "\r\n\t") / len(text)


def _deep_strings(blob, depth=0):
    """递归提取 protobuf 里的文本：Windows 版把正文藏在嵌套 message 里。

    仓库的 ex.render 对这类消息只返回 "[文本]" 空标签；其 _pb_strings 只扫顶层，
    所以这里自己下钻，让预览与 ALL_FOR_AI.md 拿到真实正文。
    """
    if depth > 4 or not isinstance(blob, (bytes, bytearray)) or len(blob) > 2 * 1024 * 1024:
        return []
    data = bytes(blob)
    from protobuf_fields import fields as pb_fields
    try:
        parsed = pb_fields(data)
    except Exception:
        try:
            text = data.decode("utf-8")
            return [text] if text.strip() and _printable_ratio(text) > 0.85 else []
        except UnicodeDecodeError:
            return []
    out = []
    for values in parsed.values():
        for value in values:
            if not isinstance(value, bytes) or not value:
                continue
            nested = None
            try:
                nested = pb_fields(value)
            except Exception:
                pass
            if nested and any(isinstance(x, bytes) for vals in nested.values() for x in vals):
                out.extend(_deep_strings(value, depth + 1))
                continue
            try:
                text = value.decode("utf-8")
                if text.strip() and _printable_ratio(text) > 0.85:
                    out.append(text)
            except UnicodeDecodeError:
                pass
    return out


def _render(ct, raw):
    try:
        text = str(ex.render(ct, raw, "windows")).strip()
    except Exception:
        text = ""
    placeholder = (not text
                   or re.fullmatch(r"\[[^\]]{1,12}\]", text)
                   or re.fullmatch(r"\[[^\]]{1,12}\]\s*[\dA-Za-z_\-]{4,}", text))
    if placeholder:
        candidates = _deep_strings(bytes(raw) if raw else b"")
        useful = [s.strip() for s in candidates
                  if s.strip() and len(s.strip()) > 1 and not s.strip().startswith("/data/")]
        if useful:
            text = "\n".join(dict.fromkeys(useful))[:4000]
    return text or "(无正文)"


# ---------------------------------------------------------------- 合并转发卡片解析
#
# ct=4 的复合消息即合并转发卡片，标题常见两种形态：
#   「<A>和<B>的聊天记录」/「Chat History for <A> and <B>」
# 典型样本含十几到几十条子消息，文字与图片/文件混排。
# 结构：
#   root.f2 标题；root.f1[] 子消息
#   子消息 f2 时间(Unix 秒) / f5 类型(0,2=文本 14,101=图片 15=文件) / f11 发言人
#          f101 文本正文（3~5 层 protobuf 嵌套）/ f103 媒体 payload
#   媒体 payload f2 原文件名、f4 原始字节数、f8/f10/f30/f31 32 位 hash、
#                f3/f26 下载 URL、f25 token、f44 扩展名
# 关键实测：图片 f103.f10 恰好等于本机 Cache/Image 里对应 .jpg 的 MD5，
#           6/6 精确命中，因此用「原始字节数 + MD5」做最强匹配键。

_FWD_TEXT_FIELDS = (101, 102, 105, 107)
_FWD_MEDIA_FIELDS = (103, 104, 106)
_HASH32_RE = re.compile(r"[0-9a-fA-F]{32}")
_NAME_EXT_RE = re.compile(r"\.[A-Za-z0-9]{1,8}$")


def _pb_map(blob):
    from protobuf_fields import ordered_fields
    out = {}
    try:
        for number, value in ordered_fields(bytes(blob)):
            out.setdefault(number, []).append(value)
    except Exception:
        return {}
    return out


def _pb_str(value, min_len=1):
    if not isinstance(value, (bytes, bytearray)):
        return ""
    try:
        text = bytes(value).decode("utf-8")
    except UnicodeDecodeError:
        return ""
    text = text.strip()
    if len(text) < min_len or _printable_ratio(text) < 0.85:
        return ""
    return text


def _fwd_noise(text):
    """合并转发里的 URL / hex / base64 / 成员 ID 串，一律不算正文。"""
    if not text:
        return True
    if text.startswith(("http://", "https://")) or "qlogo.cn" in text or "qpic.cn" in text:
        return True
    if re.fullmatch(r"[0-9a-fA-F]{20,}", text) or re.fullmatch(r"[A-Za-z0-9+/=_\-]{20,}", text):
        return True
    if re.fullmatch(r"\d{10,25}(?:[;,]\s*\d{10,25}){2,}", text) or re.fullmatch(r"\d{10,}", text):
        return True
    if len(text) <= 5 and not re.search(r"[\u4e00-\u9fff]", text) and not _NAME_EXT_RE.search(text):
        return True
    return False


def _clean_pb_text(value):
    text = _pb_str(value)
    if not text:
        return ""
    # 递归兜底时可能带出 protobuf 的 tag / 长度前缀控制字节，统一清掉
    text = re.sub(r"[\x00-\x1f]", "", text)
    return text.strip()


def _pb_texts(blob, depth=0):
    """递归抽出子消息正文（结构化解析失败时的兜底）。"""
    if depth > 7 or not isinstance(blob, (bytes, bytearray)) or len(blob) > 500000:
        return []
    fields = _pb_map(blob)
    if not fields:
        text = _clean_pb_text(blob)
        return [text] if text and not _fwd_noise(text) else []
    out = []
    for number in sorted(fields):
        for value in fields[number]:
            if isinstance(value, int):
                continue
            text = _clean_pb_text(value)
            if text and not _fwd_noise(text):
                out.append(text)
            elif len(value) > 3:
                out.extend(_pb_texts(value, depth + 1))
    return out


def _fwd_text(payload):
    """按 f101 → f1 → f2 → f1[] 的已知层级取文本段（含多段表情/换行）。

    直接把整个 payload 当字符串会带出 `\\x12`、`b` 这类长度前缀字节，
    所以优先按 protobuf 层级下钻；实在解不开才回退到递归文本抽取。
    """
    segments = []
    level1 = _pb_map(payload)
    for outer in (level1.get(1, []) or [payload]):
        if not isinstance(outer, (bytes, bytearray)):
            continue
        level2 = _pb_map(outer)
        if not level2:
            continue
        for middle in (level2.get(2, []) or level2.get(1, [])):
            if not isinstance(middle, (bytes, bytearray)):
                continue
            level3 = _pb_map(middle)
            candidates = []
            if level3:
                for number in (1, 2):
                    candidates.extend(v for v in level3.get(number, [])
                                      if isinstance(v, (bytes, bytearray)))
            else:
                candidates = [middle]
            for candidate in candidates:
                text = _clean_pb_text(candidate)
                if text and not _fwd_noise(text):
                    segments.append(text)
    if not segments:
        segments = [_clean_pb_text(t) for t in _pb_texts(payload)]
    return "\n".join(dict.fromkeys(s for s in segments if s and not _fwd_noise(s)))[:6000]


def _fwd_media(blob, sub_ct=0):
    """解析媒体 payload：原文件名 / 原始字节数 / hash / token / url。"""
    payload = _pb_map(blob)
    if not payload:
        return None
    original = ""
    for number in (2, 44):
        for value in payload.get(number, []):
            text = _pb_str(value)
            if text and _NAME_EXT_RE.search(text) and _valid_filename(text):
                original = _clean_filename(text)
                break
        if original:
            break
    size = payload.get(4, [0])[0]
    size = int(size) if isinstance(size, int) else 0
    hashes = []
    for number in (8, 10, 30, 31):
        for value in payload.get(number, []):
            text = _pb_str(value)
            if text and _HASH32_RE.fullmatch(text):
                hashes.append(text.lower())
    hashes = list(dict.fromkeys(hashes))
    token = next((t for t in (_pb_str(v) for v in payload.get(25, [])) if t), "")
    url = ""
    for number in (3, 26):
        for value in payload.get(number, []):
            candidate = _pb_str(value)
            if candidate.startswith("http"):
                url = candidate
                break
        if url:
            break
    if not (original or size or hashes or url):
        return None
    return {
        "name": original, "named": bool(original), "kind": "file" if original else "image",
        "size": size, "hashes": hashes, "token": token[:180], "url": url[:400],
        "source": None, "match": None, "placeholder": False, "thumb": False,
    }


def parse_forward_card(ct, raw):
    """识别 ct=4 的合并转发卡片，拆出全部子消息；不是卡片返回 None。"""
    if int(ct or 0) != 4 or not raw or len(raw) < 400:
        return None
    root = _pb_map(raw)
    subs_raw = [v for v in root.get(1, []) if isinstance(v, (bytes, bytearray))]
    if len(subs_raw) < 2:
        return None
    title = next((t for t in (_pb_str(v) for v in root.get(2, [])) if t and not _fwd_noise(t)), "")
    subs = []
    for index, blob in enumerate(subs_raw):
        fields = _pb_map(blob)
        if not fields:
            continue
        value = fields.get(2, [0])[0]
        ts = float(value) if isinstance(value, int) and 1_400_000_000 < value < 2_600_000_000 else 0.0
        sender = ""
        for number in (11, 6):
            for item in fields.get(number, []):
                candidate = _pb_str(item)
                if candidate and not _fwd_noise(candidate):
                    sender = candidate
                    break
            if sender:
                break
        sub_ct = fields.get(5, [0])[0]
        sub_ct = int(sub_ct) if isinstance(sub_ct, int) else 0
        texts = []
        for number in _FWD_TEXT_FIELDS:
            for value in fields.get(number, []):
                if isinstance(value, (bytes, bytearray)):
                    for segment in _fwd_text(value).splitlines():
                        if segment:
                            texts.append(segment)
        text = "\n".join(dict.fromkeys(t for t in texts if t))[:6000]
        media = []
        for number in _FWD_MEDIA_FIELDS:
            for value in fields.get(number, []):
                item = _fwd_media(value, sub_ct)
                if item:
                    media.append(item)
        if not (text or media):
            continue
        kind = "文件" if any(m["kind"] == "file" for m in media) else ("图片" if media else "文字")
        subs.append({
            "idx": index, "ts": ts, "time": ex.fmt_time(ts) if ts else "",
            "sender": sender or "未知", "ct": sub_ct, "kind": kind,
            "text": text, "files": media,
        })
    if len(subs) < 2:
        return None
    subs.sort(key=lambda s: (s["ts"], s["idx"]))
    return {"title": title or "合并转发聊天记录", "count": len(subs), "subs": subs,
            "images": 0, "matched": 0, "missing": 0}


def _image_suffix_of(path):
    try:
        with Path(path).open("rb") as handle:
            head = handle.read(12)
    except OSError:
        return ""
    if head[:3] == b"\xff\xd8\xff":
        return ".jpg"
    if head[:4] == b"\x89PNG":
        return ".png"
    if head[:4] in (b"GIF8", b"GIF9"):
        return ".gif"
    if head[:4] == b"RIFF" and head[8:12] == b"WEBP":
        return ".webp"
    return ""


_MD5_CACHE = {}
_MD5_LOCK = threading.Lock()


def _md5_of(path):
    try:
        stat = Path(path).stat()
        key = (str(path), stat.st_size, stat.st_mtime_ns)
    except OSError:
        return ""
    with _MD5_LOCK:
        if key in _MD5_CACHE:
            return _MD5_CACHE[key]
    try:
        digest = hashlib.md5()
        with Path(path).open("rb") as handle:
            for chunk in iter(lambda: handle.read(1 << 20), b""):
                digest.update(chunk)
        value = digest.hexdigest()
    except OSError:
        value = ""
    with _MD5_LOCK:
        if len(_MD5_CACHE) > 30000:
            _MD5_CACHE.clear()
        _MD5_CACHE[key] = value
    return value


_CACHE_INDEX_CACHE = {}
_CACHE_INDEX_LOCK = threading.Lock()


def _cached_cache_index(acc, ttl=180):
    key = acc["id"]
    now = time.time()
    with _CACHE_INDEX_LOCK:
        hit = _CACHE_INDEX_CACHE.get(key)
        if hit and now - hit[0] < ttl:
            return hit[1], hit[2]
    index, count = build_cache_index(acc)
    with _CACHE_INDEX_LOCK:
        _CACHE_INDEX_CACHE[key] = (now, index, count)
    return index, count


def resolve_forward_media(acc, card, card_ts, index=None):
    """按 文件名 → (原始字节数 + MD5) → (字节数唯一 + 时间接近) 精准找本机缓存。

    只在本账号 Cache 内找；无图片 magic 的全零占位文件视为「尚未下载」。
    """
    if index is None:
        index, _ = _cached_cache_index(acc)
    image_no = 0
    for sub in card["subs"]:
        for item in sub["files"]:
            if item["kind"] == "image":
                image_no += 1
                item["name"] = item["name"] or f"参考图{image_no}.jpg"
            source, how = None, None
            if item.get("named"):
                source = _lookup_cache(index, item["name"], sub["ts"] or card_ts)
                how = "name" if source else None
            if not source and item["size"]:
                # 文件名没对上（缓存里可能存成 uuid 名）时，用原始字节数 + MD5 兜底
                candidates = list(index.get("by_size", {}).get(item["size"], []))
                if item["kind"] == "image":
                    candidates = [c for c in candidates if _image_suffix_of(c[1])]
                if item["hashes"]:
                    for candidate in candidates:
                        if _md5_of(candidate[1]) in item["hashes"]:
                            source, how = candidate[1], "md5"
                            break
                elif len(candidates) == 1 and abs(
                        candidates[0][0] - (card_ts or candidates[0][0])) < 7 * 86400:
                    source, how = candidates[0][1], "size"
            if source and item["kind"] == "image" and not _image_suffix_of(source):
                # 只有「图片」才检查 magic：Word/PDF/工程文件本来就不是图片头
                item["placeholder"] = True
                source, how = None, None
            if source and item["kind"] == "image":
                ext = _image_suffix_of(source)
                if not item.get("named"):
                    item["name"] = f"参考图{image_no}{ext or '.jpg'}"
            item["source"] = str(source) if source else None
            item["match"] = how
            item["thumb"] = bool(source and item["kind"] == "image")
            item["size_kb"] = round(Path(source).stat().st_size / 1024, 1) if source else 0
            if item["kind"] == "image":
                card["images"] += 1
            if source:
                card["matched"] += 1
            else:
                card["missing"] += 1
    return card


def forward_thumb(acc, cid, mid, sub_i, file_i):
    """给界面返回合并转发子消息里某张图的本机绝对路径（调用方负责校验与读取）。"""
    cfg = load_config()
    messages = _message_index(cfg, acc, cid)
    message = next((m for m in messages if m["id"] == str(mid)), None)
    if not message or not message.get("forward"):
        return None, ""
    subs = message["forward"]["subs"]
    if not 0 <= sub_i < len(subs) or not 0 <= file_i < len(subs[sub_i]["files"]):
        return None, ""
    item = subs[sub_i]["files"][file_i]
    if not item.get("source"):
        return None, ""
    path = Path(item["source"]).resolve()
    root = Path(acc.get("cache_dir") or (Path(acc["data_dir"]).parent / "Cache")).resolve()
    if root not in path.parents or not path.is_file():
        return None, ""
    return path, _image_suffix_of(path)


# ---------------------------------------------------------------- 正文去噪 / 消息索引

ATTACHMENT_EXTS = (
    "docx", "doc", "pdf", "xlsx", "xls", "xlsm", "csv", "pptx", "ppt", "txt", "md",
    "wps", "et", "dps", "zip", "rar", "7z", "dwg", "dxf", "rvt", "rfa", "skp", "obj",
    "stp", "step", "stl", "fbx", "3ds", "3dm", "max", "igs", "iges", "sat", "sldprt",
    "png", "jpg", "jpeg", "gif", "webp", "bmp", "tif", "tiff", "heic", "psd", "ai", "svg",
    "mp4", "mov", "avi", "mp3", "wav", "amr", "silk", "m4a",
)
_EXT_ALT = "|".join(re.escape(e) for e in ATTACHMENT_EXTS)
_NOISE_URL_RE = re.compile(r"https?://[^\s\u4e00-\u9fff]+", re.I)
_NOISE_LINE_PATTERNS = [
    re.compile(r"^\d{15,20}$"),                       # 企微内部 15~20 位数字 ID
    re.compile(r"^\d{10,25}(?:[;,]\s*\d{10,25}){2,}$"),  # 分号拼接的系统成员 ID 列表
    re.compile(r"^[A-Za-z0-9+/=_\-]{24,}$"),         # hex / Base64 / v1_ 令牌 / md5
    re.compile(r"^[0-9a-fA-F]{32,}$"),
    re.compile(r"^[0-9a-fA-F\-]{32,}\.(?:jpg|jpeg|png|gif|webp)$", re.I),  # uuid 哈希图片名
    re.compile(r"^(?:file://|/var/|/private/|.*Containers/Data/Application/).*$", re.I),  # iOS 内部路径
]
_FILE_LINE_RE = re.compile(rf"^([^\\/:*?\"<>|]{{1,120}}\.(?:{_EXT_ALT}))$", re.I)
_FILE_INLINE_RE = re.compile(
    rf"(?<![\w\u4e00-\u9fff])([^\\/:*?\"<>|\s]{{1,120}}\.(?:{_EXT_ALT}))(?![\w\u4e00-\u9fff])", re.I)
_BAD_STEM_PUNCT = re.compile(r"[，。！？；：、,!?;:]")


def _clean_filename(name):
    """去掉 protobuf 里成对出现的包裹符，保留文件名本体。

    注意 `《我一生中的重要抉择》教学设计.docx` 里的书名号是文件名的一部分，
    只有整体成对包裹（如 `[图.zip]`）才剥离，不能无脑 strip 掉开头的《。
    """
    text = str(name).strip()
    for left, right in (("[", "]"), ("【", "】"), ("<", ">"), ("《", "》"),
                        ('"', '"'), ("'", "'"), ("“", "”"), ("‘", "’")):
        if len(text) > 2 and text.startswith(left) and text.endswith(right):
            text = text[1:-1].strip()
            break
    return text.strip(" \t[]【】<>\"'“”‘’,，;；:：")


def _valid_filename(name):
    stem = name.rsplit(".", 1)[0]
    if not stem or len(stem) < 2:
        return False
    if re.fullmatch(r"[0-9a-fA-F_\-]{24,}", stem):  # 纯内部哈希名，不是用户提到的文件
        return False
    if _BAD_STEM_PUNCT.search(stem):
        return False
    return True


def _extract_filenames(line):
    """返回 (文件名列表, 去掉文件名后的正文行)。整行是文件名时优先整行判定。"""
    whole = _FILE_LINE_RE.fullmatch(line.strip())
    if whole and _valid_filename(whole.group(1)):
        return [_clean_filename(whole.group(1))], ""
    names = []

    def repl(match):
        name = _clean_filename(match.group(1))
        if name and _valid_filename(name):
            names.append(name)
            return " "
        return match.group(0)

    rest = _FILE_INLINE_RE.sub(repl, line)
    return names, rest


def clean_message(ct, raw):
    """把 protobuf 噪音洗掉，返回 (干净正文, 该消息明确提到的附件文件名列表)。

    过滤：内部 CDN 链接、独立 Base64/hex 行、15~20 位纯数字 ID、uuid 哈希图片名、
    无意义短乱码行、纯 [类型] 占位符；并去重合并重复行。
    """
    rendered = _render(ct, raw)
    files, body_lines, seen = [], [], set()
    for raw_line in rendered.splitlines():
        line = _NOISE_URL_RE.sub("", raw_line).strip()
        if not line:
            continue
        line = re.sub(r"^\[[^\]]{1,12}\]\s*", "", line).strip()  # 去掉 [文件]/[卡片]/[图片] 标签
        line = line.strip(" \t*·-—|>")
        if not line or re.fullmatch(r"\[[^\]]{1,12}\]", line):
            continue
        if any(p.fullmatch(line) for p in _NOISE_LINE_PATTERNS):
            continue
        # 正文里夹着一整串分号拼接的成员 ID 时，只删掉这段 ID，保留其余文字
        if re.search(r"\d{10,25}(?:;\d{10,25}){2,}", line):
            line = re.sub(r"\d{10,25}(?:;\d{10,25}){2,}", " ", line).strip()
            if not line:
                continue
        # 行内仍夹着超长 base64/hex 令牌且没有中文正文 → 整行丢弃
        if re.search(r"[A-Za-z0-9+/=_\-]{40,}", line) and not re.search(r"[\u4e00-\u9fff]", line):
            continue
        # 独立的英文账号/ID 行（合并转发里的 wxid），不是正文
        if re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{7,23}", line):
            continue
        if (len(line) <= 6 and not re.search(r"[\u4e00-\u9fff]", line)
                and not _FILE_LINE_RE.fullmatch(line)):
            continue  # 'Rígr' / '1' / ']' / 短乱码
        names, rest = _extract_filenames(line)
        for name in names:
            if name and name not in files:
                files.append(name)
        rest = re.sub(r"\s{2,}", " ", rest).strip(
            " \t-—·|[]【】<>《》\"'“”‘’,，;；:：")
        if (rest and rest not in seen
                and not (len(rest) <= 3 and not re.search(r"[\u4e00-\u9fff]", rest))):
            seen.add(rest)
            body_lines.append(rest)
    return "\n".join(body_lines).strip(), files


_MSG_INDEX = {}
_MSG_INDEX_LOCK = threading.Lock()


def _message_index(cfg, acc, cid):
    """全量清洗后的消息（时间正序），按解密库 mtime/size 做缓存。"""
    db = _dec_message_db(acc, cfg)
    stamp = (db.stat().st_mtime_ns, db.stat().st_size)
    key = (acc["id"], str(cid))
    with _MSG_INDEX_LOCK:
        hit = _MSG_INDEX.get(key)
        if hit and hit[0] == stamp:
            return hit[1]
    rows, users = _rows(acc, cfg, cid)
    messages = []
    for mid, sid, ct, st, raw in rows:
        card = parse_forward_card(ct, raw)
        if card:
            resolve_forward_media(acc, card, _norm_ts(st))
            body = (f"【合并转发聊天记录】{card['title']} · {card['count']} 条子消息"
                    f"（含 {card['images']} 张图片/文件，已匹配 {card['matched']} 个）")
            files = [f["name"] for sub in card["subs"] for f in sub["files"] if f["name"]]
            messages.append({
                "id": str(mid), "ts": _norm_ts(st), "time": ex.fmt_time(st),
                "sender": users.get(str(sid), str(sid)),
                "type": "合并转发", "content": body, "files": files, "forward": card,
            })
            continue
        body, files = clean_message(ct, raw)
        messages.append({
            "id": str(mid), "ts": _norm_ts(st), "time": ex.fmt_time(st),
            "sender": users.get(str(sid), str(sid)),
            "type": ex.type_name(ct, "windows"),
            "content": body, "files": files,
        })
    with _MSG_INDEX_LOCK:
        _MSG_INDEX.clear()  # 只保留最近一个会话，避免内存膨胀
        _MSG_INDEX[key] = (stamp, messages)
    return messages


RANGE_DAYS = {"today": 0, "3d": 3, "7d": 7, "30d": 30}


def _public_forward(card):
    """发给浏览器的合并转发结构：去掉本机绝对路径 / token / url，只留展示字段。

    缩略图走 /api/thumb（后端按 cid+mid+序号反查本机路径并校验目录），
    避免把 Cache 绝对路径暴露到前端，也把 /api/messages 体积压下来。
    """
    subs = []
    for sub in card["subs"]:
        subs.append({
            "idx": sub["idx"], "ts": sub["ts"], "time": sub["time"],
            "sender": sub["sender"], "kind": sub["kind"], "text": sub["text"],
            "files": [{"name": f["name"], "kind": f["kind"], "size": f["size"],
                       "thumb": f.get("thumb", False), "match": f.get("match"),
                       "size_kb": f.get("size_kb", 0)} for f in sub["files"]],
        })
    return {"title": card["title"], "count": card["count"], "images": card["images"],
            "matched": card["matched"], "missing": card["missing"], "subs": subs}


def _parse_window_dt(value, end_of_day=False):
    if not value:
        return None
    text = str(value).strip().replace(" ", "T")
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
        try:
            struct = time.strptime(text, fmt)
            if fmt == "%Y-%m-%d" and end_of_day:
                struct = time.struct_time((struct.tm_year, struct.tm_mon, struct.tm_mday,
                                          23, 59, 59, struct.tm_wday, struct.tm_yday, -1))
            return time.mktime(struct)
        except ValueError:
            continue
    return None


def messages_window(cfg, acc, cid, range_key="all", start=None, end=None):
    """按时间窗返回全量（不再只取 20 条）清洗后的消息，时间正序。"""
    all_messages = _message_index(cfg, acc, cid)
    low = high = None
    if range_key == "custom":
        low = _parse_window_dt(start)
        high = _parse_window_dt(end, end_of_day=True)
    elif range_key in RANGE_DAYS:
        days = RANGE_DAYS[range_key]
        now = time.localtime()
        if days == 0:
            low = time.mktime((now.tm_year, now.tm_mon, now.tm_mday, 0, 0, 0, 0, 0, -1))
        else:
            low = time.time() - days * 86400
    picked = [m for m in all_messages
              if (low is None or m["ts"] >= low) and (high is None or m["ts"] <= high)]
    fmt = lambda t: time.strftime("%Y-%m-%d %H:%M", time.localtime(t)) if t else ""
    public = []
    for message in picked:
        item = dict(message)
        if item.get("forward"):
            item["forward"] = _public_forward(item["forward"])
        public.append(item)
    return {"conversation_id": str(cid), "count": len(picked), "total": len(all_messages),
            "from": fmt(low), "to": fmt(high), "range": range_key,
            "last_id": public[-1]["id"] if public else "", "messages": public}


def account_status(cfg):
    """给界面顶部用：每个账号的 key/在线状态、会话数、库大小与最后写入。"""
    accounts = []
    for acc in cfg["accounts"]:
        mdb = Path(acc["data_dir"]) / "message.db"
        key_valid = False
        if mdb.exists() and acc.get("key"):
            try:
                with mdb.open("rb") as fh:
                    key_valid = quick_verify(bytes.fromhex(acc["key"]), fh.read(4096))
            except Exception:
                pass
        conv_count = 0
        cf = CACHE / f"conv_{acc['id']}.json"
        if cf.exists():
            try:
                conv_count = len(json.loads(cf.read_text(encoding="utf-8")).get("conversations", []))
            except Exception:
                pass
        accounts.append({
            "id": acc["id"], "key_valid": key_valid, "online": bool(key_valid and mdb.exists()),
            "conversations": conv_count,
            "db_mb": round(mdb.stat().st_size / 1048576, 1) if mdb.exists() else 0,
            "last_write": time.strftime("%Y-%m-%d %H:%M", time.localtime(mdb.stat().st_mtime)) if mdb.exists() else "",
            "data_dir": acc["data_dir"],
        })
    root = wxwork_root(cfg)
    all_ids = sorted(p.name for p in root.iterdir() if (p / "Data" / "message.db").exists()) if root.exists() else []
    unlocked = {a["id"] for a in accounts if a["key_valid"]}
    running = False
    try:
        proc = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WXWork.exe", "/NH"],
                              capture_output=True, text=True, encoding="gbk",
                              errors="replace", timeout=15)
        running = "WXWork" in (proc.stdout or "")
    except Exception:
        pass
    return {"accounts": accounts, "locked_accounts": [i for i in all_ids if i not in unlocked],
            "wecom_running": running, "export_root": cfg["export_root"],
            "total_accounts": len(all_ids)}


def list_conversations(cfg, force=False):
    rows = []
    for acc, data in all_accounts(cfg, force=force):
        for conv in data["conversations"]:
            rows.append({**conv, "account": acc["id"]})
    rows.sort(key=lambda r: (-(r.get("last_timestamp") or 0), str(r.get("name", ""))))
    return {"count": len(rows), "conversations": rows}


def preview_messages(cfg, acc, cid, limit=20):
    """直接读已解密的库，不重新解密，所以预览很快。返回时间正序的最近 N 条。"""
    rows, users = _rows(acc, cfg, cid, descending=True, limit=limit)
    messages = []
    for mid, sid, ct, st, raw in reversed(rows):
        messages.append({
            "local_id": str(mid), "time": ex.fmt_time(st),
            "sender": users.get(str(sid), str(sid)),
            "type": ex.type_name(ct, "windows"),
            "content": _render(ct, raw),
        })
    return {"conversation_id": cid, "count": len(messages), "messages": messages}


def extract_document_text(path, limit=20000):
    """把附件里的 docx/pdf/xlsx/txt 正文抽成纯文本，供 ALL_FOR_AI.md 就地展开。"""
    suffix = path.suffix.lower()
    try:
        if suffix in (".txt", ".md", ".csv", ".json", ".log", ".xml", ".html", ".htm"):
            return path.read_text(encoding="utf-8", errors="replace")[:limit]
        if suffix == ".docx":
            from docx import Document
            doc = Document(str(path))
            parts = [p.text for p in doc.paragraphs]
            for table in doc.tables[:20]:
                for row in table.rows[:100]:
                    parts.append(" | ".join(cell.text for cell in row.cells))
            return "\n".join(parts)[:limit]
        if suffix in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
            parts = []
            for ws in wb.worksheets[:10]:
                parts.append(f"[Sheet] {ws.title}")
                for row in ws.iter_rows(max_row=200, values_only=True):
                    parts.append(" | ".join("" if v is None else str(v) for v in row))
            wb.close()
            return "\n".join(parts)[:limit]
        if suffix == ".pdf":
            try:
                import pdfplumber
                with pdfplumber.open(str(path)) as pdf:
                    return "\n".join((page.extract_text() or "") for page in pdf.pages[:30])[:limit]
            except Exception:
                from pypdf import PdfReader
                reader = PdfReader(str(path))
                return "\n".join((page.extract_text() or "") for page in reader.pages[:30])[:limit]
    except Exception as exc:
        return f"(文档解析失败: {exc})"
    return None


def build_all_for_ai(cfg, acc, cid, name=None, include_cache=False):
    """生成单文件 ALL_FOR_AI.md：时间正序聊天 + 附件正文 + 未关联清单 + 兜底缓存正文。"""
    folder = Path(cfg["export_root"]) / safe_name(name or cid)
    manifest_file = folder / "chat.media.json"
    manifest = json.loads(manifest_file.read_text(encoding="utf-8")) if manifest_file.exists() else {"entries": []}
    entry_map = {}
    for entry in manifest.get("entries", []):
        msg = entry.get("message") or {}
        entry_map.setdefault(str(msg.get("local_id")), []).append(entry)

    rows, users = _rows(acc, cfg, cid)
    lines = [
        f"# {name or cid} · 聊天记录与附件汇总（给 AI 用）",
        "",
        f"- conversation-id: `{cid}`",
        f"- 企微账号: `{acc['id']}`",
        f"- 生成时间: {time.strftime('%Y-%m-%d %H:%M:%S')}",
        "- 消息按时间正序排列；Word / PDF / Excel / TXT 附件与兜底缓存正文已就地展开，图片用相对路径引用。",
        "- 直接把本文件拖给 Codex / 其它 AI 聊天框即可；同目录下的 attachments/ 与 cache_files/ 是被引用的原文件。",
        "",
        "---",
        "",
    ]
    stats = {"messages": 0, "documents": 0, "images": 0, "other_files": 0, "missing_notes": 0,
             "cache_documents": 0, "cache_images": 0, "cache_other_files": 0, "unresolved": 0}
    unresolved = []
    for mid, sid, ct, st, raw in rows:
        stats["messages"] += 1
        lines.append(f"## [{ex.fmt_time(st)}] {users.get(str(sid), str(sid))} · [{ex.type_name(ct, 'windows')}]")
        lines.append("")
        lines.append(_render(ct, raw))
        lines.append("")
        for entry in entry_map.get(str(mid), []):
            status = entry.get("status")
            if status != "resolved":
                stats["missing_notes"] += 1
                unresolved.append((str(mid), str(status), str(entry.get("reason", ""))))
                continue
            for rel in entry.get("files") or []:
                path = folder / rel
                if path.suffix.lower() in IMAGE_EXTENSIONS:
                    stats["images"] += 1
                    lines.append(f"![图片]({rel})")
                    lines.append("")
                else:
                    text = extract_document_text(path) if path.exists() else None
                    if text is not None:
                        stats["documents"] += 1
                        lines.append(f"**附件正文 · `{path.name}`**")
                        lines.append("")
                        lines.append("```text")
                        lines.append(text)
                        lines.append("```")
                    else:
                        stats["other_files"] += 1
                        lines.append(f"- 附件：`[{path.name}]({rel})`")
                    lines.append("")

    # 未关联媒体汇总：Windows 端 --with-media 关联率低，让 AI 也知道缺口在哪
    stats["unresolved"] = len(unresolved)
    if unresolved:
        by_status = {}
        for _, status, _ in unresolved:
            by_status[status] = by_status.get(status, 0) + 1
        lines += [
            "---", "",
            "## 未能关联到消息的媒体（Windows 端限制）", "",
            f"> 共 {len(unresolved)} 条。Windows 版资源 ID 结构与 macOS 不同，多数为 "
            "`resource_not_cached` / `unsupported`；在企业微信里点开下载后可重新导出补齐。"
            "完整清单见 `chat.media.json`。", "",
            "| 状态 | 数量 |", "| --- | --- |",
        ]
        for status, count in sorted(by_status.items(), key=lambda kv: -kv[1]):
            lines.append(f"| `{status}` | {count} |")
        lines += ["", "前 100 条明细（消息 local_id · 状态 · 原因）：", "", "```text"]
        for mid, status, reason in unresolved[:100]:
            lines.append(f"{mid}\t{status}\t{reason}")
        lines += ["```", ""]

    # 兜底缓存：这些文件与消息不一一对应，但属于同期缓存，正文同样喂给 AI
    if include_cache:
        cache_dir = folder / "cache_files"
        files = sorted(_selected_cache_files(folder),
                       key=lambda p: (0 if p.suffix.lower() in CACHE_TEXT_EXTENSIONS else 1,
                                      -p.stat().st_mtime)) if cache_dir.exists() else []
        lines += [
            "---", "",
            "## 本机缓存兜底文件（按该群最近活跃时间窗筛选）", "",
            "> 这些文件无法与单条消息一一对应，但属于同一账号同期的本地缓存；"
            "Word / PDF / Excel / TXT 正文已就地展开，图片与工程文件用相对路径引用。", "",
        ]
        used_chars = 0
        for path in files:
            rel = path.relative_to(folder).as_posix()
            suffix = path.suffix.lower()
            if suffix in IMAGE_EXTENSIONS:
                stats["cache_images"] += 1
                lines.append(f"- 图片：`[{path.name}]({rel})`")
                continue
            if (suffix in CACHE_TEXT_EXTENSIONS and stats["cache_documents"] < MAX_ALL_FOR_AI_CACHE_DOCS
                    and used_chars < MAX_ALL_FOR_AI_CACHE_CHARS):
                text = extract_document_text(path, limit=12000)
                if text is not None:
                    stats["cache_documents"] += 1
                    used_chars += len(text)
                    lines += [f"**缓存文档正文 · `{rel}`**", "", "```text", text, "```", ""]
                    continue
            stats["cache_other_files"] += 1
            lines.append(f"- 文件：`[{path.name}]({rel})`")
        lines.append("")

    out = folder / "ALL_FOR_AI.md"
    out.write_text("\n".join(lines), encoding="utf-8")
    return {"path": str(out), "size_kb": round(out.stat().st_size / 1024, 1), **stats}


def _cache_rank(item):
    """可抽正文的文档最优先（小文件在前），其次其它工程/办公文件，最后图片。"""
    mtime, size, path = item
    suffix = path.suffix.lower()
    if suffix in CACHE_TEXT_EXTENSIONS:
        return (0, size, -mtime)
    if suffix not in IMAGE_EXTENSIONS:
        return (1, size, -mtime)
    return (2, 0, -mtime)


def _selected_cache_files(folder):
    """读本轮 manifest；旧版本导出的目录没有 manifest 时回退到目录扫描。"""
    mf = folder / "cache_manifest.json"
    if mf.exists():
        try:
            rels = json.loads(mf.read_text(encoding="utf-8")).get("selected", [])
            return [(folder / "cache_files" / r) for r in rels if (folder / "cache_files" / r).is_file()]
        except (ValueError, OSError):
            pass
    d = folder / "cache_files"
    return sorted((p for p in d.rglob("*") if p.is_file()), key=lambda p: (-p.stat().st_mtime)) if d.exists() else []


def prune_cache(folder):
    """显式清理：只删 cache_files/ 里不属于本轮 manifest 的工具生成文件。

    不自动执行；必须由界面按钮或 CLI 明确触发，避免误删用户放在导出目录里的东西。
    """
    folder = Path(folder)
    target = folder / "cache_files"
    if not target.exists():
        return {"removed": 0, "kept": 0, "path": str(target)}
    keep = {str(p.relative_to(target).as_posix()) for p in _selected_cache_files(folder)}
    removed = 0
    for existing in list(target.rglob("*")):
        if not existing.is_file():
            continue
        if existing.relative_to(target).as_posix() in keep:
            continue
        try:
            existing.unlink()
            removed += 1
        except OSError:
            pass
    for directory in sorted((p for p in target.rglob("*") if p.is_dir()),
                            key=lambda p: -len(p.parts)):
        try:
            directory.rmdir()
        except OSError:
            pass
    return {"removed": removed, "kept": len(keep), "path": str(target)}


def cache_filtered(cfg, acc, cid, name=None):
    """兜底：按该群最近活跃时间窗，把本机缓存里的 PDF/Word/图片/工程文件拷到 cache_files/。"""
    folder = Path(cfg["export_root"]) / safe_name(name or cid)
    target = folder / "cache_files"
    target.mkdir(parents=True, exist_ok=True)
    with db_lock():
        db = _dec_message_db(acc, cfg)
        con = connect_readonly(db)
        try:
            row = con.execute("SELECT MIN(send_time), MAX(send_time) FROM message_table WHERE conversation_id=?", (cid,)).fetchone()
        finally:
            con.close()
    # 群的历史跨度可能长达数月，直接用 min~max 会命中几乎全部缓存。
    # 取「最近 CACHE_WINDOW_DAYS 天活跃窗口」，并让文档/工程文件优先于图片。
    first = _norm_ts(row[0]) if row and row[0] else 0
    last = _norm_ts(row[1]) if row and row[1] else time.time()
    span_days = round((last - first) / 86400) if first else 0
    start = max(first - 86400, last - CACHE_WINDOW_DAYS * 86400) if first else last - CACHE_WINDOW_DAYS * 86400
    end = last + 2 * 86400
    cache = Path(acc.get("cache_dir") or (Path(acc["data_dir"]).parent / "Cache"))
    picked = []
    if cache.exists():
        for path in cache.rglob("*"):
            if not path.is_file() or path.suffix.lower() not in DOC_EXTENSIONS:
                continue
            try:
                st = path.stat()
                if start <= st.st_mtime <= end:
                    picked.append((st.st_mtime, st.st_size, path))
            except OSError:
                pass
    picked.sort(key=_cache_rank)
    selected, total_bytes, skipped = [], 0, 0
    for _, size, path in picked:
        if size > MAX_CACHE_SINGLE_BYTES:
            skipped += 1
            continue
        if len(selected) >= MAX_CACHE_FILES or total_bytes + size > MAX_CACHE_BYTES:
            skipped += 1
            continue
        selected.append(path)
        total_bytes += size
    selected_rel = {str(p.relative_to(cache).as_posix()) for p in selected}
    copied = 0
    for path in selected:
        dst = target / path.relative_to(cache)
        if dst.exists() and dst.stat().st_size == path.stat().st_size:
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, dst)
        copied += 1
    manifest_file = folder / "cache_manifest.json"
    fmt = lambda t: time.strftime("%Y-%m-%d %H:%M", time.localtime(t)) if t else "-"
    manifest_file.write_text(json.dumps({
        "window": [fmt(start), fmt(end)], "selected": sorted(selected_rel),
        "generated": time.strftime("%Y-%m-%d %H:%M:%S"),
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    stale = sum(1 for p in target.rglob("*") if p.is_file()
                and p.relative_to(target).as_posix() not in selected_rel)
    return {"path": str(target), "window": [fmt(start), fmt(end)],
            "window_note": f"群历史跨度 {span_days} 天，按最近 {CACHE_WINDOW_DAYS} 天活跃窗口筛选；文档优先、单文件上限 25MB",
            "matched": len(picked), "selected": len(selected), "copied": copied,
            "skipped": skipped, "stale": stale,
            "mb": round(total_bytes / 1048576, 1)}


def web_export(cfg, acc, cid, name=None, cache_fallback=False):
    """界面「一键导出给 AI」：聊天 + 媒体清单 + 附件 + 单文件汇总（可选兜底缓存）。

    整段放在互斥锁里，保证 5 分钟同步线程不会在「导出聊天」与「生成 ALL_FOR_AI.md」
    之间插进来覆盖 chat.media.json。
    """
    with db_lock():
        summary = export_one(cfg, acc, cid, name)
        # 先拷兜底缓存，再生成单文件，这样缓存文档正文也能写进 ALL_FOR_AI.md
        fallback = cache_filtered(cfg, acc, cid, name) if cache_fallback else None
        ai = build_all_for_ai(cfg, acc, cid, name, include_cache=bool(cache_fallback))
        return {**summary, "all_for_ai": ai, "cache_fallback": fallback}


CACHE_SUBDIRS = ("File", "Image", "Voice", "Video", "Emotion", "Files", "Images")


def _norm_filename(name):
    """归一化缓存文件名：小写、去空白、去 (1)/(2) 副本标记，让重名文件可比对。"""
    text = str(name).strip().lower()
    text = re.sub(r"\s*\(\d+\)(?=(?:\.[^.]+)?$)", "", text)
    text = re.sub(r"\s+", "", text)
    return text


def _safe_attachment_name(name):
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(name))
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned[:120] or "attachment"


def build_cache_index(acc):
    """索引当前账号 Cache 下的文件（完整名 / 主名 / 原始字节数三级），供精准匹配。

    by_size 供合并转发里的无文件名图片按 f103.f4 原始大小 + MD5 命中。
    """
    root = Path(acc.get("cache_dir") or (Path(acc["data_dir"]).parent / "Cache"))
    index = {"name": {}, "stem": {}, "by_size": {}, "all": []}
    count = 0
    if not root.exists():
        return index, count
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        try:
            rel = path.relative_to(root)
            if rel.parts and rel.parts[0] not in CACHE_SUBDIRS:
                continue
            stat = path.stat()
        except (OSError, ValueError):
            continue
        item = (stat.st_mtime, path, stat.st_size)
        index["name"].setdefault(_norm_filename(path.name), []).append(item)
        index["stem"].setdefault(_norm_filename(path.stem), []).append(item)
        index["by_size"].setdefault(stat.st_size, []).append(item)
        index["all"].append(item)
        count += 1
    return index, count


def _lookup_cache(index, filename, msg_ts):
    """精确名 → 主名（扩展名可不同）；重名取 mtime 最接近该条消息时间的那个。"""
    def nearest(items):
        if not items:
            return None
        return min(items, key=lambda it: abs(it[0] - (msg_ts or it[0])))[1]

    hit = nearest(index["name"].get(_norm_filename(filename)))
    if hit:
        return hit
    return nearest(index["stem"].get(_norm_filename(Path(filename).stem)))


def _reset_ai_dir(folder, export_root):
    """导出前清空该会话的 发给AI/ 目录（只动工具自有目录，严格校验路径）。"""
    root = Path(export_root).resolve()
    folder = Path(folder).resolve()
    if folder.name != "发给AI" or root not in folder.parents:
        raise RuntimeError(f"拒绝清理非工具导出目录: {folder}")
    if folder.exists():
        for child in folder.iterdir():
            if child.is_dir():
                shutil.rmtree(child, ignore_errors=True)
            else:
                try:
                    child.unlink()
                except OSError:
                    pass
    folder.mkdir(parents=True, exist_ok=True)


def precise_export(cfg, acc, cid, name=None, ids=None):
    """只导出已勾选消息 + 精准匹配到的本群附件，严格按聊天顺序编号。

    输出目录：<导出根>/<群名>/发给AI/
        000_chat.md          时间正序聊天；合并转发卡片在原位置完整展开子消息
        001_MMDD-HHMM_xxx    按聊天出现顺序重命名的附件（含卡片内图片/文件）
        ALL_FOR_AI.md        单文件汇总，按 001/002 顺序嵌入文档正文
    不再盲捞全账号缓存，也不会混入其它群或旧单子的文件。
    """
    wanted = {str(x) for x in (ids or [])}
    if not wanted:
        raise RuntimeError("没有勾选任何消息，请先在右侧列表里选择")
    all_messages = _message_index(cfg, acc, cid)
    picked = [m for m in all_messages if m["id"] in wanted]
    if not picked:
        raise RuntimeError("勾选的消息不在该会话中，请刷新右侧列表后重试")
    picked.sort(key=lambda m: (m["ts"], int(m["id"]) if m["id"].isdigit() else 0))

    folder = Path(cfg["export_root"]) / safe_name(name or cid) / "发给AI"
    _reset_ai_dir(folder, cfg["export_root"])
    index, cache_count = build_cache_index(acc)

    title = name or str(cid)
    chat_lines = [
        f"# {title} · 已选消息（时间正序）", "",
        f"- conversation-id: `{cid}`", f"- 企微账号: `{acc['id']}`",
        f"- 生成时间: {time.strftime('%Y-%m-%d %H:%M:%S')}",
        f"- 已选消息: {len(picked)} 条（仅当前勾选，不含未勾选消息与其它群文件）",
        "- 合并转发卡片已在该消息原位置按子消息原始时间正序完整展开，附件统一编号。", "",
        "---", "",
    ]
    ai_lines = [
        f"# {title} · 单文件汇总（给 AI）", "",
        f"- conversation-id: `{cid}` / 账号 `{acc['id']}`",
        f"- 已选消息 {len(picked)} 条，附件按聊天出现顺序编号 001、002…",
        "- 只含当前勾选消息与这些消息精准匹配到的附件；Word/PDF/Excel/TXT 正文就地展开。", "",
        "---", "",
    ]
    file_map, attachments, missing = {}, [], []
    seq = 0
    stats = {"messages": 0, "documents": 0, "images": 0, "other_files": 0,
             "forward_cards": 0, "forward_subs": 0}
    forward_hint = "需在企业微信里点开该合并转发卡片并点击查看原图/下载文件后重新导出"

    def emit_attachment(filename, source, ts, time_str, key, local_id,
                        forward=False, require_image=False, origin=""):
        """统一编号：主群消息与合并转发子消息共用同一个 001、002… 序列。"""
        nonlocal seq, chat_lines, ai_lines
        if key in file_map:  # 同一文件多次提及时复用同一个编号
            note = file_map[key]
            chat_lines += [note, ""]
            ai_lines += [note, ""]
            return
        if source is not None:
            source = Path(source)
            if not source.is_file() or (require_image and not _image_suffix_of(source)):
                source = None
        if source is None:
            # 未下载的文件不占编号，保证 001/002/003 在目录里连续
            note = f"[附件未下载: {filename}]"
            missing.append({
                "category": "missing", "status": "not_downloaded", "reason": filename,
                "kind": Path(filename).suffix.lower().lstrip(".") or "file",
                "local_id": local_id,
                "hint": forward_hint if forward else "请在企业微信里点开该消息/附件，等待下载缓存完成后重新导出",
            })
            chat_lines += [note, ""]
            ai_lines += [note, ""]
            file_map[key] = note
            return
        seq += 1
        number = f"{seq:03d}"
        stamp = time.strftime("%m%d-%H%M", time.localtime(ts)) if ts else "0000-0000"
        dest_name = f"{number}_{stamp}_{_safe_attachment_name(filename)}"
        destination = folder / dest_name
        shutil.copy2(source, destination)
        note = f"[附件 {number} -> {dest_name}]"
        suffix = destination.suffix.lower()
        attachments.append({"number": number, "name": filename, "renamed": dest_name,
                            "source": str(source), "time": time_str, "local_id": local_id,
                            "origin": origin or ("合并转发" if forward else "主聊天"),
                            "size_kb": round(destination.stat().st_size / 1024, 1)})
        chat_lines += [note, ""]
        ai_lines += [note, ""]
        text = (extract_document_text(destination, limit=20000)
                if suffix in (".txt", ".md", ".csv", ".json", ".log", ".xml", ".html", ".htm",
                              ".docx", ".xlsx", ".xlsm", ".pdf") else None)
        if text is not None:
            stats["documents"] += 1
            ai_lines += [f"**附件 {number} 正文 · `{dest_name}`**", "", "```text", text, "```", ""]
        elif suffix in IMAGE_EXTENSIONS:
            stats["images"] += 1
            ai_lines += [f"![附件 {number}]({dest_name})", ""]
        else:
            stats["other_files"] += 1
            ai_lines += [f"- 附件 {number}：`[{dest_name}]({dest_name})`", ""]
        file_map[key] = note

    for message in picked:
        stats["messages"] += 1
        card = message.get("forward")
        if card:
            # ---- 合并转发：在卡片原位置按子消息时间正序完整展开 ----
            stats["forward_cards"] += 1
            header = f"## [{message['time']}] {message['sender']} · [合并转发]"
            intro = (f"**{card['title']}** · {card['count']} 条子消息"
                     f"（以下按子消息原始时间正序完整展开，图片/文件统一编号）")
            chat_lines += [header, "", intro, "", "---", ""]
            ai_lines += [header, "", intro, "", "---", ""]
            for si, sub in enumerate(card["subs"]):
                stats["forward_subs"] += 1
                sub_header = f"### [{sub['time']}] {sub['sender']} · [{sub['kind']}]"
                body = sub["text"] or "（此条为图片/文件，无文字正文）"
                chat_lines += [sub_header, "", body, ""]
                ai_lines += [sub_header, "", body, ""]
                for fi, item in enumerate(sub["files"]):
                    filename = item["name"]
                    key = (_norm_filename(filename) if item.get("named")
                           else f"{message['id']}|{si}|{fi}|{_norm_filename(filename)}")
                    emit_attachment(
                        filename, item.get("source"), sub["ts"], sub["time"], key,
                        f"{message['id']}#{si + 1}", forward=True,
                        require_image=(item["kind"] == "image"),
                        origin=f"合并转发 · {card['title']}")
            chat_lines += ["---", ""]
            ai_lines += ["---", ""]
            continue

        # ---- 普通主群消息 ----
        header = f"## [{message['time']}] {message['sender']} · [{message['type']}]"
        chat_lines += [header, "", message["content"], ""]
        ai_lines += [f"## [{message['time']}] {message['sender']}", "", message["content"], ""]
        for idx, filename in enumerate(message["files"]):
            emit_attachment(filename, _lookup_cache(index, filename, message["ts"]),
                            message["ts"], message["time"], _norm_filename(filename),
                            message["id"])

    if attachments:
        ai_lines += ["---", "", "## 附件清单（按聊天出现顺序）", "",
                     "| 编号 | 消息时间 | 来源 | 文件 |", "| --- | --- | --- | --- |"]
        for item in attachments:
            ai_lines.append(f"| {item['number']} | {item['time']} | {item['origin']} | `{item['renamed']}` |")
        ai_lines.append("")
    if missing:
        ai_lines += ["---", "", "## 尚未下载的附件", "",
                     "> 合并转发里的图片/文件：需在企业微信里点开该卡片并点击查看原图/下载文件，等缓存落地后重新导出。", ""]
        for item in missing:
            ai_lines.append(f"- 消息 {item['local_id']}：`{item['reason']}`")
        ai_lines.append("")

    chat_path = folder / "000_chat.md"
    ai_path = folder / "ALL_FOR_AI.md"
    chat_path.write_text("\n".join(chat_lines), encoding="utf-8")
    ai_path.write_text("\n".join(ai_lines), encoding="utf-8")
    return {
        "account": acc["id"], "conversation_id": str(cid), "name": title,
        "out_dir": str(folder), "chat_md": str(chat_path),
        "messages": len(picked), "attachments_ok": len(attachments),
        "files": attachments, "missing": missing,
        "cache_files_scanned": cache_count,
        "all_for_ai": {"path": str(ai_path), "size_kb": round(ai_path.stat().st_size / 1024, 1), **stats},
        "exported_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }


# ---------------------------------------------------------------- 实时增量同步
#
# 前端每 8 秒静默轮询 /api/changes。直接读加密 message.db + WAL 的 stat 做水印；
# 有写入时只把 WAL 里「已提交」的页解密后补进解密库（本机 WAL 约 7 MB，
# 补丁约 0.2 秒），而不是每来一条新消息都整库重解 297 MB（约 8 秒）。
# WAL 发生 checkpoint / generation 变化时自动退回整库重解，并对整库重解做节流。

LIVE_STATE = RUNTIME / "live"
PAGE_SZ = 4096
_WAL_MAGICS = (0x377F0682, 0x377F0683)
_LIVE_FULL_LOCK = threading.Lock()
_LAST_FULL = {}


def _stat_sig(path):
    try:
        info = Path(path).stat()
        return (info.st_size, info.st_mtime_ns)
    except OSError:
        return (0, 0)


def live_signature(acc):
    main = Path(acc["data_dir"]) / "message.db"
    return {"main": _stat_sig(main), "wal": _stat_sig(str(main) + "-wal")}


def _wal_checksum(data, byteorder, state=(0, 0)):
    import struct
    first, second = state
    for left, right in struct.iter_unpack(byteorder + "II", data):
        first = (first + left + second) & 0xFFFFFFFF
        second = (second + right + first) & 0xFFFFFFFF
    return first, second


def _wal_committed(acc):
    """只读 WAL：返回 (generation_hex, commit_size, {page_no: 加密页})。

    遇到正在写入/校验失败的帧就停在最后一个完整提交，绝不把半帧当成已提交数据。
    """
    import struct
    wal = Path(acc["data_dir"]) / "message.db-wal"
    try:
        data = wal.read_bytes()
    except OSError:
        return None, 0, {}
    if len(data) < 32:
        return None, 0, {}
    magic, version, page_size = struct.unpack(">III", data[:12])
    if magic not in _WAL_MAGICS or version != 3007000 or page_size != PAGE_SZ:
        return None, 0, {}
    byteorder = "<" if magic == _WAL_MAGICS[0] else ">"
    state = _wal_checksum(data[:24], byteorder)
    if state != struct.unpack(">II", data[24:32]):
        return None, 0, {}
    generation = data[16:24]
    pending, committed, commit_size = {}, {}, None
    frame_size = PAGE_SZ + 24
    for offset in range(32, len(data) - frame_size + 1, frame_size):
        header = data[offset:offset + 24]
        if header[8:16] != generation:
            break
        page_number, size = struct.unpack(">II", header[:8])
        if not 1 <= page_number <= 0xFFFFFFFE:
            break
        page = data[offset + 24:offset + frame_size]
        state = _wal_checksum(header[:8] + page, byteorder, state)
        if state != struct.unpack(">II", header[16:24]):
            break
        pending[page_number] = page
        if size:
            committed.update(pending)
            pending.clear()
            committed = {number: payload for number, payload in committed.items()
                         if number <= size}
            commit_size = size
    return generation.hex(), commit_size or 0, committed


def sync_live_db(acc, cfg):
    """有写入则把 WAL 已提交页增量补进解密库；失败/结构变化则整库重解。

    返回 (changed, method)，method ∈ unchanged / wal / full / pending。
    """
    LIVE_STATE.mkdir(parents=True, exist_ok=True)
    state_path = LIVE_STATE / f"{acc['id']}.json"
    sig = live_signature(acc)
    try:
        previous = json.loads(state_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        previous = {}
    if (previous.get("main") == list(sig["main"])
            and previous.get("wal") == list(sig["wal"])):
        return False, "unchanged"

    dec = Path(acc["decrypted_dir"]) / "message.db"
    generation, commit_size, pages = _wal_committed(acc)
    can_patch = bool(
        dec.exists() and generation and pages and commit_size
        and previous.get("generation") == generation
        and previous.get("main_size") == sig["main"][0]
        and acc.get("key")
    )
    method = "wal"
    with db_lock():
        if can_patch:
            from wxwork_crypto import decrypt_page, has_plain_header_fragment
            size_before = dec.stat().st_size
            old_pages = {}
            try:
                key = bytes.fromhex(acc["key"])
                if 1 in pages and not has_plain_header_fragment(pages[1]):
                    raise ValueError("WAL 页1 缺明文头片段，退回整库重解")
                grown = commit_size - size_before // PAGE_SZ
                if grown > 0 and any(n not in pages for n in
                                     range(size_before // PAGE_SZ + 1, commit_size + 1)):
                    raise ValueError("WAL 缺少新增页，退回整库重解")
                with dec.open("r+b") as handle:
                    for number, page in sorted(pages.items()):
                        offset = (number - 1) * PAGE_SZ
                        if offset >= size_before:
                            old_pages[number] = b""
                        else:
                            handle.seek(offset)
                            old_pages[number] = handle.read(PAGE_SZ)
                        handle.seek(offset)
                        handle.write(decrypt_page(key, page, number))
                    handle.truncate(max(size_before, commit_size * PAGE_SZ))
                    handle.flush()
                    os.fsync(handle.fileno())
                    handle.seek(28)
                    handle.write(commit_size.to_bytes(4, "big"))
                connection = connect_readonly(dec, immutable=True)
                try:
                    connection.execute("SELECT count(*) FROM message_table").fetchone()
                finally:
                    connection.close()
            except Exception:
                # 逐页回滚，绝不让一次失败的补丁破坏已解密的库
                try:
                    with dec.open("r+b") as handle:
                        for number, buffer in old_pages.items():
                            handle.seek((number - 1) * PAGE_SZ)
                            handle.write(buffer)
                        handle.truncate(size_before)
                except OSError:
                    pass
                can_patch = False
        if not can_patch:
            now = time.monotonic()
            with _LIVE_FULL_LOCK:
                last = _LAST_FULL.get(acc["id"], 0)
                if now - last < 20:
                    return True, "pending"
                _LAST_FULL[acc["id"]] = now
            refresh(acc, cfg, force=True)  # 项目自带命令整库重解（约 8 秒）
            generation, commit_size, _ = _wal_committed(acc)
            method = "full"
        with _MSG_INDEX_LOCK:
            _MSG_INDEX.clear()

    state = {"main": list(sig["main"]), "wal": list(sig["wal"]),
             "main_size": sig["main"][0], "generation": generation or "",
             "pages": commit_size, "synced_at": time.time()}
    try:
        state_path.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass
    return True, method


def check_changes(cfg, acc, cid, last_id=""):
    """前端 5~10 秒轮询：水印检查 + WAL 增量补丁 + 只读两条 SQL，不重建消息索引。

    前端对比 last_id / total 决定是否增量刷新右侧列表，因此这里刻意不做
    全量 protobuf 解析（6228 条消息的群重建一次要几秒）。
    """
    try:
        changed, method = sync_live_db(acc, cfg)
    except Exception as exc:
        return {"ok": False, "changed": False, "error": str(exc),
                "last_updated": time.strftime("%H:%M:%S")}
    try:
        db = _dec_message_db(acc, cfg)
        with db_lock():
            con = connect_readonly(db)
            try:
                total = con.execute(
                    "SELECT count(*) FROM message_table WHERE conversation_id=?", (cid,)
                ).fetchone()[0]
                row = con.execute(
                    "SELECT message_id, send_time FROM message_table WHERE conversation_id=? "
                    "ORDER BY send_time DESC, message_id DESC LIMIT 1", (cid,)
                ).fetchone()
            finally:
                con.close()
        current_id = str(row[0]) if row else ""
        last_time = ex.fmt_time(row[1]) if row else ""
    except Exception as exc:
        return {"ok": False, "changed": False, "error": str(exc),
                "last_updated": time.strftime("%H:%M:%S")}
    return {"ok": True, "changed": bool(changed or (last_id and current_id != str(last_id))),
            "method": method, "total": total, "last_id": current_id,
            "last_time": last_time, "last_updated": time.strftime("%H:%M:%S")}


def watch_get():
    return {"ids": parse_watch(),
            "text": WATCH.read_text(encoding="utf-8-sig") if WATCH.exists() else ""}


def watch_set(cids):
    cids = [str(c).strip() for c in cids if str(c).strip()]
    header = []
    if WATCH.exists():
        for line in WATCH.read_text(encoding="utf-8-sig").splitlines():
            if line.strip().startswith("#"):
                header.append(line)
            elif line.strip():
                break
    if not header:
        header = ["# 定时同步清单 —— 每行一个 conversation-id", "# 由本地 Web 界面维护"]
    WATCH.write_text("\n".join(header + [""] + cids) + "\n", encoding="utf-8")
    return {"ids": cids}


def sync_status():
    state = json.loads(STATE.read_text(encoding="utf-8")) if STATE.exists() else {}
    return {"last_sync": state.get("last_sync"), "fingerprints": state.get("fingerprints", {}),
            "watched": parse_watch()}


def rescan():
    """界面「重新扫描企微账号与密钥」：重扫内存 → 补账号 → 清缓存 → 强制刷新会话。"""
    cfg = load_config()
    before = {a["id"] for a in cfg["accounts"]}
    scan_out, scan_err = "", ""
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(FIND_KEYS)],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=200,
        )
        scan_out = proc.stdout or ""
        scan_err = (proc.stderr or "").strip()
    except subprocess.TimeoutExpired:
        scan_err = "扫描超时（200s），沿用已有 key"
    keys = {}
    if KEYS_JSON.exists():
        try:
            keys = json.loads(KEYS_JSON.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            pass
    cfg = load_config()
    sync_accounts_from_keys(cfg, keys)
    cfg = load_config()
    added = [a["id"] for a in cfg["accounts"] if a["id"] not in before]
    for old in CACHE.glob("conv_*.json"):
        try:
            old.unlink()
        except OSError:
            pass
    refreshed, errors = [], []
    for acc in cfg["accounts"]:
        try:
            refreshed.append({"account": acc["id"],
                              "conversations": len(refresh(acc, cfg, force=True)["conversations"])})
        except Exception as exc:
            errors.append(f"{acc['id']}: {exc}")
    scan_lines = [line for line in scan_out.splitlines()
                  if ("KEY=" in line or "MATCHED" in line or "SCAN_LIMITED" in line
                      or "扫描内存" in line or "扫描模式" in line or "NO_WXWORK" in line)]
    if scan_err:
        scan_lines.append("扫描提示: " + scan_err[:200])
    return {"added": added, "keys": sorted(keys.keys()), "scan_lines": scan_lines[:20],
            "refreshed": refreshed, "errors": errors,
            "scan_limited": "SCAN_LIMITED" in scan_out, "scan_stderr": scan_err}


def open_folder(path):
    target = Path(path)
    if not target.exists():
        raise RuntimeError(f"目录不存在: {target}")
    os.startfile(str(target))
    return {"opened": str(target)}


# ---------------------------------------------------------------- CLI

def main():
    parser = argparse.ArgumentParser(description="wecom-agent Windows 导出/同步工具")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("accounts")
    p_refresh = sub.add_parser("refresh"); p_refresh.add_argument("--force", action="store_true")
    p_match = sub.add_parser("match"); p_match.add_argument("keyword"); p_match.add_argument("--force", action="store_true")
    p_export = sub.add_parser("export"); p_export.add_argument("--cid", required=True)
    p_export.add_argument("--account"); p_export.add_argument("--name")
    p_sync = sub.add_parser("sync"); p_sync.add_argument("--force", action="store_true")
    p_cache = sub.add_parser("cache"); p_cache.add_argument("--cid", required=True)
    p_cache.add_argument("--account"); p_cache.add_argument("--name")
    sub.add_parser("fingerprint")
    sub.add_parser("status")
    sub.add_parser("rescan")
    sub.add_parser("watch-get")
    sub.add_parser("sync-status")
    p_conv = sub.add_parser("conversations"); p_conv.add_argument("--force", action="store_true")
    p_prev = sub.add_parser("preview"); p_prev.add_argument("--cid", required=True)
    p_prev.add_argument("--account"); p_prev.add_argument("--name"); p_prev.add_argument("--limit", type=int, default=20)
    for name in ("all-for-ai", "cache-filtered", "web-export"):
        p = sub.add_parser(name); p.add_argument("--cid", required=True)
        p.add_argument("--account"); p.add_argument("--name")
        if name == "web-export":
            p.add_argument("--cache-fallback", action="store_true")
    p_ws = sub.add_parser("watch-set"); p_ws.add_argument("--ids", default="")
    p_of = sub.add_parser("open-folder"); p_of.add_argument("--path", required=True)
    p_prune = sub.add_parser("prune-cache"); p_prune.add_argument("--path", required=True)
    p_msgs = sub.add_parser("messages"); p_msgs.add_argument("--cid", required=True)
    p_msgs.add_argument("--account"); p_msgs.add_argument("--range", default="all")
    p_msgs.add_argument("--start"); p_msgs.add_argument("--end")
    p_pe = sub.add_parser("precise-export"); p_pe.add_argument("--cid", required=True)
    p_pe.add_argument("--account"); p_pe.add_argument("--name"); p_pe.add_argument("--ids", required=True)
    args = parser.parse_args()
    cfg = load_config()

    if args.command == "accounts":
        out = {"accounts": [{"id": a["id"], "data_dir": a["data_dir"], "key_ready": bool(a.get("key"))}
                            for a in cfg["accounts"]],
               "export_root": cfg["export_root"], "python": cfg["python"]}
    elif args.command == "refresh":
        accs = [a for a in cfg["accounts"] if not getattr(args, "account", None) or a["id"] == args.account]
        out = {"refreshed": [{"account": a["id"], "conversations": len(refresh(a, cfg, force=args.force)["conversations"])}
                             for a in accs]}
    elif args.command == "match":
        out = match(cfg, args.keyword, force=args.force)
    elif args.command == "export":
        acc = next((a for a in cfg["accounts"] if a["id"] == args.account), None) if args.account else None
        if acc is None:
            acc, name = locate(cfg, args.cid)
            if acc is None:
                raise RuntimeError(f"找不到会话 {args.cid}")
        else:
            name = args.name
        out = export_one(cfg, acc, args.cid, name or args.name)
    elif args.command == "sync":
        out = sync(cfg, force=args.force)
    elif args.command == "cache":
        acc = next((a for a in cfg["accounts"] if a["id"] == args.account), None) if args.account else None
        found_name = None
        if acc is None:
            acc, found_name = locate(cfg, args.cid)
            if acc is None:
                raise RuntimeError(f"找不到会话 {args.cid}")
        out = export_cache(cfg, acc, args.cid, args.name or found_name)
    elif args.command == "status":
        out = account_status(cfg)
    elif args.command == "conversations":
        out = list_conversations(cfg, force=args.force)
    elif args.command in ("preview", "all-for-ai", "cache-filtered", "web-export"):
        acc = next((a for a in cfg["accounts"] if a["id"] == args.account), None) if args.account else None
        nm = args.name
        if acc is None:
            acc, found = locate(cfg, args.cid)
            if acc is None:
                raise RuntimeError(f"找不到会话 {args.cid}")
            nm = nm or found
        if args.command == "preview":
            out = preview_messages(cfg, acc, args.cid, args.limit)
        elif args.command == "all-for-ai":
            has_cache = (Path(cfg["export_root"]) / safe_name(nm or args.cid) / "cache_files").exists()
            out = build_all_for_ai(cfg, acc, args.cid, nm, include_cache=has_cache)
        elif args.command == "cache-filtered":
            out = cache_filtered(cfg, acc, args.cid, nm)
        else:
            out = web_export(cfg, acc, args.cid, nm, cache_fallback=args.cache_fallback)
    elif args.command == "rescan":
        out = rescan()
    elif args.command == "watch-get":
        out = watch_get()
    elif args.command == "watch-set":
        out = watch_set([c for c in args.ids.split(",") if c.strip()])
    elif args.command == "sync-status":
        out = sync_status()
    elif args.command == "open-folder":
        out = open_folder(args.path)
    elif args.command == "prune-cache":
        out = prune_cache(args.path)
    elif args.command in ("messages", "precise-export"):
        acc = next((a for a in cfg["accounts"] if a["id"] == args.account), None) if args.account else None
        if acc is None:
            acc, found = locate(cfg, args.cid)
            if acc is None:
                raise RuntimeError(f"找不到会话 {args.cid}")
        if args.command == "messages":
            out = messages_window(cfg, acc, args.cid, args.range, args.start, args.end)
        else:
            out = precise_export(cfg, acc, args.cid, args.name,
                                 [x.strip() for x in args.ids.split(",") if x.strip()])
    else:
        out = {"fingerprints": {a["id"]: fingerprint(a) for a in cfg["accounts"]}}
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False), file=sys.stdout)
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

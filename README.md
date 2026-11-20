# 企微 AI 导出工具（Windows）

> **开源声明**：本项目基于 MIT 协议的 [freanocank/wecom-agent](https://github.com/freanocank/wecom-agent) 二次开发，保留原项目 `LICENSE` 中的完整版权声明（Copyright (c) 2026 tzwkb）与 MIT 授权条款。在此致谢原作者的工作。

纯本地运行的 Windows 企业微信聊天导出与可视化工具：把本机聊天记录、合并转发卡片和附件整理成适合 Codex / 其他 AI 直接读取的目录，并提供浏览器界面完成筛选、勾选与导出。

---

## 核心功能

- **Windows 本地解密**：从 `WXWork.exe` 内存提取 wxSQLite3 密钥，校验并解密本机 `message.db`，不依赖任何服务端。
- **多账号自动发现**：一次内存扫描匹配全部账号的库指纹；切换企微账号后可点「重新扫描企微账号与密钥」重新抓取。
- **Web 可视化界面**：`http://127.0.0.1:8765`，只监听本机回环地址；账号状态、会话搜索、消息预览、导出与同步全在浏览器完成。
- **搜索即联动**：搜索框 300 ms 防抖，自动高亮并选中第一条匹配会话，右侧立刻切换。
- **日期筛选 + 逐条勾选**：全部 / 今天 / 近 3 天 / 近 7 天 / 近 30 天 / 自定义时间段（精确到分钟）；每条消息独立勾选，支持全选、反选、清空和「从此条往后全选」。
- **合并转发深度展开**：识别 `ct=4` 的嵌套转发卡片，按子消息原始时间正序展开文字、图片与文件；图片按「原始字节数 + MD5」在本账号缓存中精准匹配并生成缩略图。
- **WAL 实时增量同步**：每 8 秒静默检查数据库水印，只把 WAL 中已提交的页解密后补进本地解密库；有新消息时增量追加到右侧列表，保留勾选状态与滚动位置。
- **按聊天顺序统一编号打包**：只导出勾选的消息；附件按出现顺序重命名为 `001_MMDD-HHMM_原文件名`，写入 `000_chat.md` 原位标记；生成单文件 `ALL_FOR_AI.md`，就地展开 Word / PDF / Excel / TXT 正文。
- **不混单**：导出前自动清空该会话的 `发给AI/` 目录，未下载附件原位标注且不占编号。

---

## Windows 一键安装与使用

### 1. 准备

- Windows 10/11，Python 3.10+（需加入 PATH），Git。
- 已安装并登录企业微信桌面版，且至少打开过要导出的会话。

### 2. 安装

```bat
git clone <你的仓库地址> wecom-ai-exporter
cd wecom-ai-exporter
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt -r requirements-documents.txt
```

> Windows 解密还需要 `pycryptodome`，已写入 `requirements.txt`。

### 3. 配置

```bat
copy config.example.json config.json
```

用记事本编辑 `config.json`，把 `REPLACE_WITH_...` 占位符换成你的实际路径：

| 字段 | 含义 |
| --- | --- |
| `python` | 项目内 `.venv\Scripts\python.exe` 的完整路径 |
| `root` | 项目根目录完整路径 |
| `wxwork_root` | 企业微信账号数据根目录（其下是各账号 ID 目录，每个含 `Data` 与 `Cache`） |
| `export_root` | 导出根目录，例如桌面下的一个文件夹 |
| `cache_ttl_seconds` | 会话列表缓存秒数 |
| `accounts` | 留空数组即可；抓取密钥后会自动填充 |

`config.json`、`watch_groups.txt`、密钥文件与导出内容都已被 `.gitignore` 忽略，不会进入版本库。

### 4. 启动

双击项目根目录的 **`start_web.bat`**，浏览器会自动打开 `http://127.0.0.1:8765`。

使用顺序：

1. 首页「账号与状态」确认已解锁账号；如需重扫，点「重新扫描企微账号与密钥」。
2. 左侧搜索群名或会话 ID，右侧自动切换到该会话。
3. 选时间范围，确认勾选消息（默认全选当前筛选结果）；可用「从此条往后全选」跳过旧单。
4. 点「一键导出给 AI」，完成后点「打开导出文件夹」，将 `发给AI/` 内文件全选拖给 AI 工具。

### 5. 可选：后台定时同步

- 在界面「每 5 分钟自动同步」区开启，或在左侧列表勾选要监视的会话。
- 传统方式：双击 `start_sync.bat` 启动隐藏窗口，`stop_sync.bat` 停止；`sync_once.bat` 立即同步一次。
- `startup_sync.vbs` 是可选的开机自启入口，默认未启用。

---

## 安全与免责声明

- **纯本地离线运行**：服务只绑定 `127.0.0.1`，不上传聊天记录、密钥、账号信息或附件到任何第三方服务器；不内置任何遥测或联网上报。
- **仅限本人账号本机数据**：请只用于读取你自己有权访问的企业微信本机数据，并遵守所在单位的数据合规制度与相关法律法规。
- **密钥只在内存与本机文件中流转**：解密密钥来自本机进程内存扫描，仅保存在被 `.gitignore` 忽略的本地文件中；请勿把 `config.json`、`*key*.txt`、`wxwork_keys.json`、`watch_groups.txt`、`runtime/` 或导出目录提交、分享或上传。
- **风险自担**：本项目按「原样」提供，不保证对企业微信未来版本的兼容性；使用解密功能前请确认你已获得相应授权。项目作者对任何数据丢失、隐私泄露或违规使用不承担责任。
- 发现敏感信息被误提交时，请立即撤销对应提交并轮换相关凭据。

---

## 致谢

- [freanocank/wecom-agent](https://github.com/freanocank/wecom-agent) —— 本项目的上游基础，MIT 协议，原版权声明完整保留在 `LICENSE` 中。
- 原项目的 Windows 解密适配、导出原语与 protobuf 解析实现。

## 协议

MIT License。详见 [LICENSE](LICENSE)。

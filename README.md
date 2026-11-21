# 企微 AI 导出工具（Windows）

> **开源声明**：本项目基于 MIT 协议的 [freanocank/wecom-agent](https://github.com/freanocank/wecom-agent) 二次开发，保留原项目 `LICENSE` 中的完整版权声明（Copyright (c) 2026 tzwkb）与 MIT 授权条款。在此致谢原作者的工作。

纯本地运行的 Windows 企业微信聊天导出与可视化工具：把本机聊天记录、**任意层嵌套的合并转发卡片**和**任意格式的附件**整理成**单个 `.zip` 压缩包**，供 Codex / 其他 AI 直接读取；另提供浏览器界面完成筛选、逐条勾选与一键导出。

---

## v1.1.0 更新

- **任意层套娃递归解析**：合并转发不再只展开一层。解析器会遍历子消息的全部 protobuf 字段，发现嵌套卡片就地递归展开（支持 2 层、3 层乃至更深，带防环保护）。`000_chat.md` 用 `### [第1层转发]` → `#### [第2层转发]` 层级标题展示，网页界面同步按层缩进。
- **取消文件后缀白名单**：不再靠写死的扩展名列表判断「是不是文件」。凡企业微信文件 / 图片 / 视频 / 语音 payload，一律从 protobuf 提取真实文件名、原始字节数与 32 位 hash，再到本机 `Cache/File`、`Cache/Image`、`Cache/Video`、`Cache/Voice` 精准匹配（优先 `字节数 + MD5`），按原格式打包。PDF / Word / Excel / PPT / WPS / 压缩包 / 工程文件（OBJ、STL、STP、NC、RVT、SKP、C4D、PSD、AI…）全部支持；可转文本的格式额外把正文就地展开进 `ALL_FOR_AI.md`，其余保留原文件并原位标注。
- **图文混排卡片**：`ct=123` 卡片里内嵌的图片与文字一起解析，图片同样进入全局编号序列。
- **短数字不再被误删**：只过滤 15~20 位的企微内部 ID，正常聊天里的 `1`、`120`、`500` 等接单 / 报价数字完整保留。
- **统计口径统一**：卡片顶部统计、子项状态与实际打包文件数 100% 一致（`附件总数 = 已匹配 + 未缓存`），不再出现「含 0 个附件，已匹配 1」这类矛盾。
- **单 ZIP 导出 + 置顶高亮**：导出先在 `runtime/tmp_export/` 临时目录生成文件，打包成 `YYYYMMDD-HHMM_会话名.zip` 后自动删除临时目录，导出根目录**只多一个 ZIP**，不再残留散落的群文件夹。点击「定位 ZIP（自动选中）」会用 `explorer.exe /select` 把资源管理器强制拉到屏幕最前并高亮该 ZIP；另有「浏览器下载 ZIP」按钮。
- **带专属图标的静默启动**：`launch_web.vbs` 无黑框启动、自动检测服务、等待端口就绪，优先用 Edge / Chrome 独立应用窗口打开界面；`assets/app.ico` 是多尺寸应用图标，可一键创建桌面快捷方式。

---

## 核心功能

- **Windows 本地解密**：从 `WXWork.exe` 内存提取 wxSQLite3 密钥，校验并解密本机 `message.db`，不依赖任何服务端。
- **多账号自动发现**：一次内存扫描匹配全部账号的库指纹；切换企微账号后可点「重新扫描企微账号与密钥」重新抓取。
- **Web 可视化界面**：`http://127.0.0.1:8765`，只监听本机回环地址；账号状态、会话搜索、消息预览、导出与同步全在浏览器完成。
- **搜索即联动**：搜索框 300 ms 防抖，自动高亮并选中第一条匹配会话，右侧立刻切换。
- **日期筛选 + 逐条勾选**：全部 / 今天 / 近 3 天 / 近 7 天 / 近 30 天 / 自定义时间段；每条消息独立勾选，支持全选、反选、清空和「从此条往后全选」。
- **合并转发递归展开**：按各层子消息原始时间正序展开文字、图片与文件；图片按「原始字节数 + MD5」在本账号缓存中精准匹配并生成缩略图。
- **WAL 实时增量同步**：每 8 秒静默检查数据库水印，只把 WAL 中已提交的页解密后补进本地解密库；有新消息时增量追加到右侧列表。
- **单 ZIP 按聊天顺序打包**：ZIP 内固定为 `000_chat.md` → `001_…` → `002_…` → … → `ALL_FOR_AI.md`，中文文件名 UTF-8 不乱码。支持「带时间戳」与「覆盖同名」两种命名模式。
- **不混单**：只导出勾选消息精准匹配到的附件，未下载附件原位标注且不占编号。

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

> `requirements.txt` 含 `pycryptodome`（解密）、`pywin32`（资源管理器置顶 / 快捷方式 / 回收站操作）与 `Pillow`（可选，重新生成图标）。`requirements-documents.txt` 含 Word / PDF / Excel 正文提取依赖；若还想展开 PPTX 正文，可额外 `pip install python-pptx`。

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

`config.json`、`watch_groups.txt`、密钥文件、导出 ZIP 与 `runtime/` 都已被 `.gitignore` 忽略，不会进入版本库。

### 4. 启动（推荐：带图标、无黑框）

`launch_web.vbs` 会以脚本所在目录为项目根目录，自动：

1. 检测 `127.0.0.1:8765` 是否已在运行；没运行则用 `pythonw.exe` 后台静默拉起 `web_app.py`；
2. 等待 `/api/health` 就绪；
3. 优先用 Edge / Chrome 的 `--app=` 独立窗口打开界面，否则用默认浏览器。

创建桌面快捷方式（可选，一次即可）：

```bat
.venv\Scripts\python.exe assets\create_shortcut.py
```

之后双击桌面上的 **「企微AI导出工具」** 即可无黑框打开。图标文件是 `assets/app.ico`；想换图案可改 `assets/make_icon.py` 后运行它重新生成。

> 传统入口 `start_web.bat` / `start_sync.bat` 仍然可用，只是会显示控制台窗口。

### 5. 使用顺序

1. 首页「账号与状态」确认已解锁账号；如需重扫，点「重新扫描企微账号与密钥」。
2. 左侧搜索群名或会话 ID，右侧自动切换到该会话。
3. 选时间范围，确认勾选消息；可用「从此条往后全选」跳过旧单。
4. 点「一键导出给 AI」，得到**一个 ZIP**。导出完成会自动在资源管理器里置顶高亮它，也可点「浏览器下载 ZIP」直接下载。
5. 把 ZIP 整个拖给 AI 工具，或解压后按 `000_chat.md` → 编号附件 → `ALL_FOR_AI.md` 顺序使用。

### 6. 可选：后台定时同步

- 在界面「每 5 分钟自动同步」区开启，或在左侧列表勾选要监视的会话。
- 传统方式：双击 `start_sync.bat` 启动隐藏窗口，`stop_sync.bat` 停止；`sync_once.bat` 立即同步一次。
- `startup_sync.vbs` 是可选的开机自启入口，默认未启用。

---

## 导出产物说明

```
YYYYMMDD-HHMM_会话名.zip
├── 000_chat.md          时间正序聊天正文；转发按 ### / #### 层级展开
├── 001_MMDD-HHMM_原文件名
├── 002_MMDD-HHMM_原文件名
├── ...
└── ALL_FOR_AI.md        单文件汇总 + 可读取文档正文就地展开 + 附件清单
```

`000_chat.md` 与 `ALL_FOR_AI.md` 在每个附件原位都标注 `[附件 00X -> 00X_原文件名]`；无法转纯文本的格式保留原格式文件在 ZIP 中。

---

## 安全与免责声明

- **纯本地离线运行**：服务只绑定 `127.0.0.1`，不上传聊天记录、密钥、账号信息或附件到任何第三方服务器；不内置任何遥测或联网上报。启动器只访问本机回环地址做健康检查。
- **仅限本人账号本机数据**：请只用于读取你自己有权访问的企业微信本机数据，并遵守所在单位的数据合规制度与相关法律法规。
- **密钥只在内存与本机文件中流转**：解密密钥来自本机进程内存扫描，仅保存在被 `.gitignore` 忽略的本地文件中；请勿把 `config.json`、`*key*.txt`、`wxwork_keys.json`、`watch_groups.txt`、`runtime/`、桌面快捷方式或导出 ZIP 提交、分享或上传。
- **风险自担**：本项目按「原样」提供，不保证对企业微信未来版本的兼容性；使用解密功能前请确认你已获得相应授权。项目作者对任何数据丢失、隐私泄露或违规使用不承担责任。
- 发现敏感信息被误提交时，请立即撤销对应提交并轮换相关凭据。

---

## 致谢

- [freanocank/wecom-agent](https://github.com/freanocank/wecom-agent) —— 本项目的上游基础，MIT 协议，原版权声明完整保留在 `LICENSE` 中。
- 原项目的 Windows 解密适配、导出原语与 protobuf 解析实现。

## 协议

MIT License。详见 [LICENSE](LICENSE)。

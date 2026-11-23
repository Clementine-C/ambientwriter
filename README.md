# Word 伴写（AmbientWriter）

在 Microsoft Word 里做到 **WPS AI 伴写**那种体验：写作时停顿一下，光标后出现灰色续写，
**Tab 接受 / Esc 放弃**，继续打字自动丢弃。

Word 的 web 加载项平台做不到行内灰字（拿不到逐字符输入、拦不住 Tab），所以这里分成两层：

| 层 | 是什么 | 在哪 |
|---|---|---|
| **伴写引擎** | 独立进程，通过 Word COM 驱动文档，键盘钩子接管 Tab/Esc | `AmbientWriter/` |
| **设置面板** | 改造过的 [word-GPT-Plus](https://github.com/Kuingsmile/word-GPT-Plus) 前端，多一个「伴写」标签页 | `patches/` 里的补丁 |

两层通过本地 HTTPS 服务的 `/api/bridge` 共享配置，改完即时生效。

## 快速开始

```powershell
# 1) 打开伴写面板（服务 + watcher 一起拉起，浏览器直达设置页）
open-banshe-settings.cmd

# 2) 在侧栏「伴写」页填 API 密钥（DeepSeek）
#    Base URL: https://api.deepseek.com/v1    模型: deepseek-chat

# 3) 打开 Word 开始写：停顿约 1.1 秒 → 灰字出现 → Tab / Esc
start-word-gptplus.cmd
```

没有密钥时走模拟续写，只演示交互，不联网。

## 目录

```
AmbientWriter/          伴写引擎（C# 源码 + 构建脚本）
  src/Watcher.cs          主逻辑：轮询光标、插入灰字、Tab/Esc 钩子
  build.ps1               编译（csc，免管理员）
  make-certs.ps1          生成本地 CA + 服务器证书
  config.template.ini     配置模板
patches/                  对 word-gpt-plus 的全部改动（补丁 + 两个新文件）
tools/ui-check/           无头 Edge 驱动的面板 UI 自检脚本
*.cmd                     启动器 / 修复脚本
STATUS.md                 完整的踩坑与验证记录
使用说明.md               日常使用与排障
```

## 已实现的细节

- 触发：上下文 ≥ 10 字符**或**以句末标点结尾；`Tab` 接受、`Esc` 放弃、其它键丢弃
- 接受后 1.5 秒冷却，避免连续弹建议
- 灰色文字用 Word 自己的 `Range.End` 定位，删除前比对文本，避免误删正文
- 切换文档自动丢弃建议，不会在别的文档里删除
- 模型失败/返回空 → 写入 `status.json`，面板显示最近错误
- 位置：`%LOCALAPPDATA%\AmbientWriter\`（`config.ini`、`bridge.json`、`status.json`、`watcher.log`）
- 诊断：`AmbientWriterWatcher.exe --selftest` 会真调一次模型并写日志

## 已知限制

- **Office 2024 零售版**不会持久化 web 加载项：每次 Word 会话都要在「开始 → 加载项 → GPT Plus」
  手点一次。这是 Office 侧的机制，脚本无法绕过（详见 `STATUS.md` 第 25–28 轮）。
- 伴写引擎只对 Windows 桌面版 Word 有效；Word 网页版 / Mac 不可用。
- 面板依赖本地 HTTPS 服务，服务没起时面板读不到配置（伴写本身仍会用 `config.ini` 的旧值）。

## 许可证

伴写引擎为本仓库原创；`patches/` 中的改动基于 MIT 许可的
[word-GPT-Plus](https://github.com/Kuingsmile/word-GPT-Plus)，遵循其 MIT 许可。

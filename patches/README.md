# word-gpt-plus 本地改动（补丁）

本目录记录我们对上游 **word-GPT-Plus** 的全部修改，仓库本身不收录那份 300+ MB 的检出。

- 上游：`https://github.com/Kuingsmile/word-GPT-Plus`（MIT）
- 基线提交：`3e48047cea17be6c210e1950f62b5b890d3008cb`

## 改了什么

| 文件 | 改动 |
|---|---|
| `src/main.ts` | 支持 `https://<host>:3000/settings` 深链接直接进设置页 |
| `src/pages/SettingsPage.vue` | 新增「伴写」标签页（开关 / API 密钥 / Base URL / 模型 / 停顿毫秒 / 后台状态 / 退出进程），通过 `/api/bridge` 与伴写进程通信 |
| `serve-https.cjs`（新增） | 本地 HTTPS 静态服务 + `/api/bridge`、`/api/status` 两个端点 |
| `vite.selfhost.config.js`（新增） | 用本地证书起 `vite preview` 的配置 |

## 怎么还原

```powershell
git clone https://github.com/Kuingsmile/word-GPT-Plus.git word-gpt-plus
cd word-gpt-plus
git checkout 3e48047cea17be6c210e1950f62b5b890d3008cb

# 1) 打补丁（两个已跟踪文件的改动）
git apply --whitespace=nowarn ..\patches\word-gpt-plus.patch

# 2) 拷入两个新文件
Copy-Item ..\patches\new-files\*.cjs, ..\patches\new-files\*.js . -Force

# 3) 构建
pnpm install --config.engine-strict=false --ignore-scripts
pnpm --config.engine-strict=false build
```

`dist/` 由构建产生，不要提交。

## 注意

- 补丁是在 LF 工作区上生成的；如果检出后 git 把行尾转成 CRLF，`git apply` 可能报
  whitespace 冲突，加 `--ignore-space-change` 或先设置 `git config core.autocrlf false`。
- 上游更新后补丁可能失效，`SettingsPage.vue` 是冲突高发区（我们加了 155 行）。

<p align="center">
  <img src="assets/cover.png" width="720" alt="dsh-basic-office">
</p>

# dsh-basic-office

DSH 办公文件插件：让 DeepSeek Harness 的 Agent **阅读**常见二进制文档（PDF / DOCX / XLSX / PPTX / CSV），并**生成**常见办公文件（DOCX / XLSX / PPTX / PDF / CSV）。

DSH 原生的 `read` / `write` / `edit` 只处理 UTF-8 文本（`read_image` 只处理图片），二进制办公文档是空白。本插件补齐这一块，且**不与任何原生工具重名/冲突**（全部使用 `office_*` 前缀）。

## 提供的工具

| 工具 | 说明 |
|---|---|
| `office_read_pdf` | 提取 PDF 文本，支持 `page_start` / `page_end` 页范围与 `max_chars` 截断 |
| `office_read_docx` | 提取 Word .docx 文本（标题、段落、表格文字） |
| `office_read_xlsx` | 读取 .xlsx / .xls / .csv，按工作表返回制表符分隔的行；支持 `sheet` / `range` / `max_rows` |
| `office_read_pptx` | 提取 PowerPoint .pptx 每页幻灯片的文字（含表格文字） |
| `office_write_docx` | 生成 Word 文档：居中标题、`heading1-3`/`quote`/普通段落（粗体/斜体/对齐/颜色）、带表头底色的表格 |
| `office_write_xlsx` | 生成 Excel 工作簿：多工作表、字符串/数字/布尔/空单元格、自动列宽 |
| `office_write_pptx` | 生成 PPT：16:9 宽屏、标题/副标题/分级项目符号/表格 |
| `office_write_pdf` | 生成 PDF：多页、标题/段落/项目符号/表格；中文等非 ASCII 文本自动使用 Windows 系统字体（微软雅黑等） |
| `office_write_csv` | 生成 CSV：可选表头、`comma`/`tab`/`semicolon` 分隔符、默认带 UTF-8 BOM（Excel 打开中文不乱码） |

## 设计要点（与原生能力对齐）

- **沙箱一致**：生成文件与原生 `write` 一样受会话沙箱策略约束——`workspace-write` 下只能写入工作区与系统临时目录，越界会得到与原生相同的 `[sandbox: ...]` 拒绝提示，并支持 `sandbox_permissions` + `justification` 的一次性升级（需用户批准）。
- **读前保护**：生成文件遵循原生 fs 观察策略（先读后覆写），未读过的已有文件不会被直接覆盖。
- **路径解析**：相对路径按会话工作区解析，与原生 `read` / `write` 一致。
- **观察事件**：读取/生成都会发出 `fs/observed`，与原生工具链互通。
- **规范化**：工具定义遵循官方文档 `docs/user/develop/basic/tool.md` 与 `docs/cookbook/adding-a-tool.md`（`defineTool` + `ctx.tools.register`、参数 DSL、规范返回值、`output.render`、纯函数 `presentCall`）。

## 安装（用户）

```sh
# 无需 git，直接从 GitHub 下载安装（已实测可用）：
dsh plugin --profile web add https://codeload.github.com/pqkisvery666/dsh-basic-office/tar.gz/refs/heads/main

# 或机器上有 git 时：
dsh plugin --profile web add git+https://github.com/pqkisvery666/dsh-basic-office.git
```

安装后 **重启 dsh web 服务**（`Ctrl+C` 后重新 `dsh web`），新工具即对所有会话生效。

```sh
dsh plugin --profile web up      # 更新
dsh plugin --profile web remove dsh-basic-office   # 卸载
```

## 安装（作者：放入插件目录的真实拷贝方式）

插件以 npm tarball 形式安装：`dsh plugin add` 会把 tarball **解包成真实目录**放进 profile 的 `node_modules`（与 dsh-better-sidebar 等同级），**不使用链接**，profile 与开发目录完全解耦。

```sh
# 1. 在插件开发目录里打 tarball：
cd dsh-basic-office   # 插件源码目录
pnpm pack             # 生成 dsh-basic-office-<version>.tgz

# 2. 把 tarball 放进 profile 的 plugins 目录（可复用，后续重装/升级都从它还原）：
mkdir -p ~/.dsh/profiles/web/plugins
cp dsh-basic-office-0.1.0.tgz ~/.dsh/profiles/web/plugins/

# 3. 从 tarball 安装（--config.minimumReleaseAge=0 仅用于绕过 pnpm 11 供应链策略对
#    近期发布的旧包（dsh-better-sidebar 等）的误报，属一次性临时参数）：
dsh plugin --profile web add "C:\Users\19065\.dsh\profiles\web\plugins\dsh-basic-office-0.1.0.tgz" --config.minimumReleaseAge=0
```

安装后 **重启 dsh web 服务**（`Ctrl+C` 后重新 `dsh web`），新工具即对所有会话生效。

更新插件：改完源码 → 重新 `pnpm pack` → 覆盖 `~/.dsh/profiles/web/plugins/` 里的 tarball → `dsh plugin --profile web up`（或 remove 后重新 add）。

### 卸载

```sh
dsh plugin --profile web remove dsh-basic-office
```

## 在会话中使用

直接对模型说即可，例如：

- 「读一下 `C:\path\报告.pdf` 的前 10 页」
- 「把这个表格生成 `data.xlsx` 和 `data.csv`」
- 「把以下内容生成一份 `周报.docx`（带标题和表格）和一个 5 页的 `汇报.pptx`」
- 「把这段总结输出成 `总结.pdf`」

## 开发

```sh
pnpm install        # 安装依赖
node test/smoke.mjs # 冒烟测试：5 种格式生成↔解析回环 + 全工具面执行
```

## 目录结构

```
dsh-basic-office/
  cordis.patch.yml   # dsh.bundle.patch 层：插入插件行（office-tools）
  lib/index.js       # Cordis 插件入口：9 个工具的注册与执行
  lib/parsers.js     # 解析侧纯函数（PDF/DOCX/XLSX/PPTX）
  lib/generators.js  # 生成侧纯函数（DOCX/XLSX/PPTX/PDF/CSV）
  test/smoke.mjs     # 冒烟测试
```

## 已知限制

- 扫描版 PDF（纯图片无文字层）无法提取文本（无 OCR）。
- 老式二进制 `.doc` / `.xls` 仅 `.xls` 受支持；`.doc` 请先另存为 `.docx`。
- 非 Windows 平台生成 PDF 时若无 CJK 系统字体，中文无法渲染（西文正常）。
- PPTX 读取不含图片/绘图中的文字。

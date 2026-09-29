# DBX Query Report：查询前先看 SQL，确认后生成报表

用已经接入 DBX MCP 的 AI Agent 查询数据库：先核对表结构和统计口径，展示带注释的完整 SQL；你确认后，Agent 才执行查询，并交付 Excel、离线 HTML 图表和 SQL 文件。

这是一份供 Agent 执行的技能，不是 DBX 安装程序。技能包不含数据库连接、库表字段、账号密码或真实业务数据。

## 安装前准备

- 按 [DBX 官方 MCP 文档](https://dbxio.com/en/docs/mcp)完成 DBX 安装与 MCP 接入，并让目标 Agent 能看到 DBX 的连接、元数据和只读查询工具。请在 DBX 中限制可访问的连接与数据库，优先使用只读账号。
- 目标 Agent 支持读取 Agent Skills，能写入本地文件，并能运行 Python 3.10 或更高版本。随包报表脚本只用 Python 标准库，不需要安装 Excel 或第三方 Python 包。

技能文件遵循 [Agent Skills 格式](https://agentskills.io/specification)。DBX MCP 的安装与连接步骤请看官方文档；本仓库只负责接入完成后的查询和交付流程。

## 安装

### 方法一：通过 skills CLI 安装

```bash
npx skills add https://github.com/limuyun88/dbx-query-report
```

命令会让你选择目标 Agent 和安装范围。也可以明确指定，例如安装到当前项目的 Codex：

```bash
npx skills add https://github.com/limuyun88/dbx-query-report --skill dbx-query-report -a codex
```

`npx skills add` 是可选的第三方安装工具，需要 Node.js 和 npx；其安装参数以 [skills CLI 文档](https://github.com/vercel-labs/skills)为准。

### 方法二：下载 ZIP 手动安装

下载仓库 ZIP，解压后把**整个** `dbx-query-report` 文件夹放到目标 Agent 的技能目录。不要只复制 `SKILL.md`，因为运行时还会用到 `references/` 和 `scripts/`。

| Agent | 项目级目录示例 |
| --- | --- |
| Codex | 项目根目录下的 `.agents/skills/dbx-query-report/` |
| Claude Code | 项目根目录下的 `.claude/skills/dbx-query-report/` |
| 其他支持 Agent Skills 的 Agent | 按该 Agent 官方文档指定的技能目录 |

放好后，确认路径中有 `dbx-query-report/SKILL.md`，并重新开启 Agent 会话。Codex 和 Claude Code 的路径分别见其[技能文档](https://developers.openai.com/codex/skills)与[技能文档](https://code.claude.com/docs/en/skills)。

### 验证安装

在新会话中输入：

```text
使用 dbx-query-report。先检查你是否能读取该技能、发现 DBX MCP 的元数据和只读查询工具；只报告检查结果，不查询业务数据。
```

成功时，Agent 应能说明它找到了技能和可用的 DBX 工具。找不到工具时，先检查目标 Agent 的 MCP 连接；仅看到技能名称还不足以执行数据库查询。

## 使用

用自然语言提出业务问题，并说明连接、期间与已确定的口径。技能按以下顺序工作：

1. 核对目标连接、表结构及字段注释；对影响结果但尚未确定的口径提问。
2. 展示**完整、带业务注释的只读 SQL**，说明日期范围、指标算法和仍待确认的假设。
3. 你确认后再执行；结果完整性检查通过后，生成三种文件。

如果只需要 SQL，明确说“先展示 SQL，不执行”。如果结果已在当前会话查出，明确说“只导出已有结果，不重新查询”。

### 示例一：先看 SQL，再执行

```text
使用 dbx-query-report。请用我已授权的 DBX 连接，
统计 2026 年 1 月和 2 月的月度销售额，并分别列出每月销售额前 10 的商品。
先核对日期字段、有效订单状态、商品标识和销售额算法。
先给我完整 SQL，不要执行；等我确认后再查询并导出报表。
```

Agent 应先展示可审阅的完整 SQL。你确认后，可以回复：

```text
确认执行刚才展示的 SQL，并生成 Excel、离线 HTML 和 SQL 文件。
```

商品榜单按月分别排名，是否跨区域合计、退款和优惠如何处理，均以你确认的口径为准。示例没有指定真实连接或字段名，Agent 不能自行编造。

### 示例二：只核对结构

```text
使用 dbx-query-report，只查看我指定连接下的目标表结构和字段注释。
帮我指出写销售额 SQL 还需要确认哪些口径，不查询订单明细。
```

### 示例三：只导出当前结果

```text
使用 dbx-query-report，把本会话已经查出的完整结果导出为 Excel、离线 HTML 和 SQL。
不要再次查询数据库。
```

交付目录中的 `report.xlsx` 保存全部返回分组，`report.html` 可离线打开，`report.sql` 保存实际执行的 SQL。结果 JSON 和交付文件存放在你的工作区，不属于公开技能包。若 DBX 返回结果被截断，Agent 不应把它标成完整榜单或完整合计。

## 不连接 DBX 的本地演示

仓库提供一份**完全虚构**的测试数据，可单独验证三件套导出。从仓库根目录运行：

```bash
python scripts/export_report.py --input tests/fixtures/report.json --out-dir demo-output
```

运行后会新建 `demo-output/`，其中包含 `report.xlsx`、`report.html` 和 `report.sql`。目标目录必须尚不存在，以免覆盖文件。这个演示不会调用 DBX，也不能代替真实 MCP 连通性测试。

## 常见问题

**Agent 找得到技能，却找不到 DBX 工具？** 检查 DBX MCP 是否接入了当前 Agent、连接是否允许访问，以及新会话是否加载了 MCP。

**SQL 已执行，但没有完整报表？** 检查查询是否失败或被 DBX 工具截断；部分结果不能冒充完整合计或前 10 榜单。

**只想要 SQL？** 在请求里写“先展示 SQL，不执行”。技能会停在预览阶段。

**导出脚本提示输出目录已存在？** 换一个新目录名；脚本不会覆盖旧交付。

## 验证范围与许可

本地已使用虚构数据验证导出脚本、XLSX 读取、金额和空值、两个月分别展示商品榜、HTML 桌面与窄屏布局。不同 Agent 对技能发现、MCP 工具和本地运行权限的支持各有差异；目前尚未在所有 Agent 上完成真实 DBX 端到端测试。

开发者可运行 `python -m unittest discover -s tests -v`。详见 [SKILL.md](SKILL.md)、[导出数据格式](references/export-format.md)和 [MIT 许可证](LICENSE)。

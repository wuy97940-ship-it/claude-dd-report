# 安装指南

## 方式 A：装进 Claude Code（推荐）

### A1. 放置仓库

把本仓库 clone（或复制）到一个固定位置，例如：

```powershell
git clone https://github.com/<你的用户名>/claude-dd-report.git D:\ai-tools\claude-dd-report
```

> 路径基准：SKILL.md 与 SOP 内的全部相对路径（`tools\`、`workflows\`、`templates\`、`docs\`）
> 均以**仓库根目录**为基准，放置后不要单独移动子目录。

### A2. 注册 Skill

把 skill 挂进 Claude Code 的技能目录（Windows 用 junction，不复制文件，仓库更新即时生效）：

```powershell
# 用户级（所有项目可用）
New-Item -ItemType Junction -Path "$env:USERPROFILE\.claude\skills\dd-report" -Target "D:\ai-tools\claude-dd-report\skills\dd-report"

# 或项目级（仅当前项目）
New-Item -ItemType Junction -Path "<项目>\.claude\skills\dd-report" -Target "D:\ai-tools\claude-dd-report\skills\dd-report"
```

Linux/macOS 用软链接：

```bash
ln -s /path/to/claude-dd-report/skills/dd-report ~/.claude/skills/dd-report
```

### A3. 注册 Agent（可选但推荐）

```powershell
Copy-Item "D:\ai-tools\claude-dd-report\agents\dd-expert.md" "$env:USERPROFILE\.claude\agents\dd-expert.md"
```

> dd-expert 的 frontmatter 里引用了企查查 MCP（`mcp__qcc-*`）。未配置企查查时删掉这些工具名即可，
> SKILL 内置了公开渠道降级路径。

### A4. 验证

新开 Claude Code 会话，输入：

```
按 examples/content.example.json 生成一份示例尽调报告，然后跑交付门禁
```

它应当：调 dd_report_builder 构建 → 跑 dd_report_selfcheck 报告 FAIL=0 → 告知产物路径。

## 方式 B：装进其他 Agent（通用）

任何支持 "SKILL.md 约定"（Anthropic Agent Skills 规范）的 Agent：

1. 把 `skills/dd-report/SKILL.md` 及其所在目录整体放入该 Agent 的技能加载路径
2. `tools/` 需要 Python 3.10+ 与 `python-docx`、`lxml`；把 SKILL.md 中 `tools\*.py` 的调用按你的路径解释
3. 把 SKILL.md 的触发描述（frontmatter `description`）注册进该 Agent 的技能路由

无技能机制的 Agent（如通用助手）：直接把 SKILL.md 当系统提示词的一部分喂给它，
并告诉它工具脚本的可执行路径，同样可以驱动全流程。

## Python 依赖清单

| 包 | 用途 | 必需性 |
|---|---|---|
| python-docx | builder/门禁/手术全链 | 必需 |
| lxml | docx_edit_lib XML 手术 | 必需 |
| pywin32 | dd_word_finalize（Word COM） | 可选（Windows+Office） |
| PyMuPDF (fitz) | 渲染目检截图 | 可选 |
| openpyxl | dd_extract_xlsx 底稿提取 | 可选 |

## 贵司格式定制（一次性 10 分钟）

1. 打开 `templates/build_master.py`，按贵司规范改字体/字号/行距/页边距
2. 运行 `python templates/build_master.py` 重新生成 `templates/dd_report_master.docx`
3. 更彻底的做法：拿一份**贵司已定稿的尽调报告**直接当母版——builder 的 capture_exemplars
   会自动从定稿报告识别 8 类格式样本，无需任何配置（母版里不要有批注和修订残留，
   有就先跑 `python tools/dd_comment_closure.py --close 定稿报告.docx`）

# WPS_Skills 项目 Bug 修复报告

> 执行人：Claude少校 | 时间：2026-03-21 14:04:17 +08:00
> 时间校验：本机14:04:17 / Cloudflare 14:04:19 / Apple 14:04:20，偏差≤3秒，通过

## 一、修复概览

| 编号 | 类型 | 问题描述 | 优先级 | 状态 |
|------|------|---------|--------|------|
| B1 | 代码Bug | wps_word_generate_toc 工具重复定义 | P0 | ✅ 已修复 |
| B2 | 代码Bug | wps_word_set_font 工具重复定义 | P0 | ✅ 已修复 |
| B3 | 代码Bug | Excel/PPT index.ts 缺失个体工具导出 | P1 | ✅ 已修复 |
| B4 | 代码Bug | tools/index.ts 注释工具数不准确(64→145) | P1 | ✅ 已修复 |
| B5 | 文档Bug | README.md 描述与实际能力不符 | P0 | ✅ 已修复 |
| B6 | Issue #6 | MCP Server 连接失败 | P0 | ✅ 已修复 |
| B7 | Issue #5 | Linux 平台加载项路径错误 | P1 | ✅ 已修复 |
| B8 | Issue #4 | WPS 加载项 arguments error | P0 | ✅ 已修复 |
| B9 | 文档 | README 重构为人类+AI双模式 | P1 | ✅ 已修复 |
| B10 | 文档 | SKILL.md 工具列表同步 | P1 | ✅ 已修复 |

## 二、详细修复记录

### B1: wps_word_generate_toc 重复定义
- **根因**: format.ts 和 document.ts 都定义了工具名 `wps_word_generate_toc`
- **修复**: document.ts 中重命名为 `wps_word_generate_doc_toc`
- **文件**: `wps-office-mcp/src/tools/word/document.ts`

### B2: wps_word_set_font 重复定义
- **根因**: format.ts 和 content.ts 都定义了工具名 `wps_word_set_font`
- **修复**: content.ts 中重命名为 `wps_word_set_font_style`，变量名同步更新
- **文件**: `wps-office-mcp/src/tools/word/content.ts`, `word/index.ts`

### B3: index.ts 导出缺失
- **根因**: workbook.ts/data-advanced.ts/textbox.ts 的个体导出未添加到各 index.ts
- **修复**: 在 excel/index.ts 添加 workbook(10对) 和 data-advanced(7对) 导出；ppt/index.ts 添加 textbox(7对) 导出
- **文件**: `excel/index.ts`, `ppt/index.ts`

### B4: tools/index.ts 注释不准确
- **根因**: 注释声称"共64个"工具，实际145个
- **修复**: 更新注释为 Excel(65)+Word(24)+PPT(42)+Common(2)+内置(12)=145
- **文件**: `wps-office-mcp/src/tools/index.ts`

### B5: README.md 重构
- **修复**: 重构为人类+AI双模式说明，准确描述MCP工具数133个（+12内置=145），添加 Claude Code/Cursor/Augment 支持说明
- **文件**: `README.md`

### B6: Issue #6 - MCP连接失败
- **根因**: dist/ 过期编译产物含重复工具名，ToolRegistry.register() 抛出异常导致 Server 崩溃
- **修复**:
  1. tool-registry.ts 重复注册改为 warn+skip 而非 throw
  2. install.sh/auto-install-mac.sh 编译前先 rm -rf dist
- **文件**: `wps-office-mcp/src/server/tool-registry.ts`, `scripts/install.sh`, `scripts/auto-install-mac.sh`

### B7: Issue #5 - Linux平台加载项问题
- **根因**: Linux 加载项目录名错误，缺少 WPS 要求的尾部后缀；publish.xml 不支持 Linux
- **修复**:
  1. install.sh 目标目录改为正确名称
  2. update_publish_xml() 扩展支持 Linux
  3. INSTALL.md 新增 Linux 安装指南
- **文件**: `scripts/install.sh`, `INSTALL.md`

### B8: Issue #4 - arguments error
- **根因**: wps-claude-assistant/manifest.xml 缺少 ribbon/scripts/permissions 声明
- **修复**: 补齐 manifest.xml 完整声明，对齐 Windows 版
- **文件**: `wps-claude-assistant/manifest.xml`

### B9-B10: SKILL.md 同步更新
- **修复**: 4个 SKILL.md 文件的工具列表与代码实际注册一致
  - wps-excel: 46→65个
  - wps-word: 22→24个
  - wps-ppt: 35→42个
  - wps-office: 综合133个

## 三、Git 提交记录

| Commit | 描述 |
|--------|------|
| 56063ca | fix: 修复 GitHub Issues #4/#5/#6 三个安装与启动问题 |
| df7deca | fix: 修复代码层7个bug（重名工具/导出缺失/注释不符/README重构） |
| ded040e | docs: T17 SKILL.md同步更新至最新工具列表 |

## 四、验证状态

- ✅ Word 工具名唯一性 - 无重复
- ✅ Excel/PPT index.ts 导出完整性 - 全部导出
- ✅ README.md 数据准确性 - MCP工具133+12内置=145
- ✅ SKILL.md 工具列表一致性 - 与代码同步
- ✅ tool-registry.ts 容错性 - 重复注册不再崩溃
- ✅ manifest.xml 完整性 - ribbon/scripts 声明齐全
- ✅ install.sh Linux 路径修正 - 加载项目录名正确

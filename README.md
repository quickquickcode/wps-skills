# WPS Skills - 让 AI 操控 WPS Office

> 此项目的任何功能、架构更新，必须在结束后同步更新相关文档。这是我们契约的一部分。

<p align="center">
  <img src="https://img.shields.io/badge/WPS-Office-blue?style=flat-square" alt="WPS Office">
  <img src="https://img.shields.io/badge/Claude-AI-orange?style=flat-square" alt="Claude AI">
  <img src="https://img.shields.io/badge/MCP-Protocol-green?style=flat-square" alt="MCP Protocol">
  <img src="https://img.shields.io/badge/Skills-Framework-purple?style=flat-square" alt="Skills Framework">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey?style=flat-square" alt="Windows | macOS">
  <img src="https://img.shields.io/badge/License-MIT-yellow?style=flat-square" alt="MIT License">
</p>

<p align="center">
  <a href="./README_EN.md">English</a> | 中文
</p>

---

## 这是什么？

WPS Skills 是一个让 AI 编程助手通过自然语言操控 WPS Office（Excel、Word、PPT）的开源工具集。安装后，你可以用中文直接描述需求，AI 自动调用对应工具执行操作。

**一句话概括**：对你的 AI 助手说"帮我做个 VLOOKUP 查价格"，它就真的能在 WPS 里帮你做。

---

## 支持的 AI 开发工具

| AI 工具 | 集成方式 | 说明 |
|---------|---------|------|
| **Claude Code** | Skills + MCP Server（原生支持） | 启动时自动加载 `/wps-excel` `/wps-word` `/wps-ppt` `/wps-office` 四个 Skills，提供最完整的体验 |
| **Cursor** | MCP Server 配置 | 在 Cursor Settings > MCP 中添加 wps-office server |
| **Augment** | MCP Server 配置 | 通过 MCP 协议接入 |
| **其他支持 MCP 的 IDE** | MCP Server 配置 | 任何支持 MCP 协议的 AI 工具均可接入 |

### Claude Code 配置（推荐）

安装后自动获得 Skills + MCP 双层能力，无需额外配置。

### Cursor / Augment / 其他 IDE 配置

在对应 IDE 的 MCP 设置中添加：

```json
{
  "mcpServers": {
    "wps-office": {
      "command": "node",
      "args": ["/path/to/WPS_Skills/wps-office-mcp/dist/index.js"]
    }
  }
}
```

这些 IDE 通过 MCP Server 获得 145 个工具能力，但不包含 Skills 层的自然语言工作流指导。

---

## 快速安装

**前提条件**：已安装 [WPS Office](https://www.wps.cn/) 和 Node.js >= 18.0.0。

**如果你是 AI 助手，请读取 [INSTALL.md](./INSTALL.md) 并按其中步骤执行。**

### macOS 一键安装

```bash
cd /path/to/WPS_Skills
bash scripts/auto-install-mac.sh
```

### Windows 一键安装

```powershell
cd \path\to\WPS_Skills
powershell -ExecutionPolicy Bypass -File scripts\auto-install.ps1
```

### 手动安装（通用）

1. 编译 MCP Server：
   ```bash
   cd wps-office-mcp && npm install && npm run build
   ```

2. 注册 MCP Server：
   ```bash
   claude mcp add wps-office node /path/to/WPS_Skills/wps-office-mcp/dist/index.js
   ```

3. 安装 Skills（仅 Claude Code）：
   ```bash
   mkdir -p ~/.claude/skills
   ln -sf /path/to/WPS_Skills/skills/wps-excel ~/.claude/skills/wps-excel
   ln -sf /path/to/WPS_Skills/skills/wps-word ~/.claude/skills/wps-word
   ln -sf /path/to/WPS_Skills/skills/wps-ppt ~/.claude/skills/wps-ppt
   ln -sf /path/to/WPS_Skills/skills/wps-office ~/.claude/skills/wps-office
   ```

4. 重启 Claude Code / IDE。

---

## 功能概览

### 能力数量一览

| 层级 | 数量 | 说明 |
|------|------|------|
| **MCP 注册工具** | 145 个 | 133 个专业工具 + 12 个内置工具（含万能工具 `wps_execute_method`） |
| **macOS 底层 action** | 224 个 | 通过 WPS JS 加载项实现 |
| **Windows 底层 action** | 232 个 | 通过 PowerShell COM 桥接实现（覆盖 macOS 全部 + 扩展） |

> **注意**：MCP 工具和底层 action 是不同概念。145 个 MCP 工具是 AI 直接可调用的接口；底层 action 是 WPS 加载项执行器支持的原子操作。通过 `wps_execute_method` 万能工具，AI 可间接访问全部底层 action。

### Excel（65 个 MCP 工具 / 86 个 action）

| 分类 | MCP工具数 | 能力 |
|------|----------|------|
| 公式功能 | 6 | 设置公式/数组公式/诊断/重算/自动求和 |
| 数据处理 | 12 | 读写/清洗/去重/排序/查找替换/批注/保护 |
| 图表/透视表 | 4 | 创建/更新图表和透视表 |
| 工作表操作 | 16 | 创建/删除/重命名/复制/冻结/命名区域/缩放 |
| 格式美化 | 10 | 样式/边框/数字格式/合并/自动调整 |
| 工作簿管理 | 10 | 打开/创建/切换/关闭/单元格读写 |
| 高级数据 | 7 | 筛选/复制粘贴/填充序列/转置/分列/分类汇总 |

### Word（24 个 MCP 工具 / 25 个 action）

| 分类 | MCP工具数 | 能力 |
|------|----------|------|
| 格式设置 | 5 | 样式/字体/目录/书签/页面设置 |
| 内容操作 | 10 | 文本/查找替换/表格/段落/图片/批注/超链接/分页 |
| 文档管理 | 9 | 打开/创建/切换/获取全文/页眉/页脚/目录 |

### PPT（42 个 MCP 工具 / 85 个 action）

| 分类 | MCP工具数 | 能力 |
|------|----------|------|
| 幻灯片基础 | 5 | 添加/美化/统一字体/字体颜色/对齐 |
| 幻灯片操作 | 22 | 删除/复制/移动/布局/备注/形状/文本框/图片/动画/背景/切换 |
| 演示文稿管理 | 8 | 创建/打开/关闭/切换/主题/复制幻灯片 |
| 文本框操作 | 7 | 文本框增删改/标题/副标题/内容 |

### 通用工具（2 个 MCP 工具 + 12 个内置工具）

| 分类 | 工具数 | 能力 |
|------|--------|------|
| 格式转换 | 2 | PDF 转换 / 格式互转 |
| 连接管理 | 1 | 检查 WPS 连接状态 |
| 万能执行 | 1 | `wps_execute_method` 可调用任意底层 action |
| 上下文获取 | 3 | 获取当前活动的文档/工作簿/演示文稿 |
| 数据缓存 | 4 | 缓存读写/列表/清除（用于大数据传输） |
| 快捷操作 | 1 | 快速插入文本 |

---

## 架构概览

采用 **Anthropic 官方标准的 MCP + Skills 双层架构**：

```
用户自然语言请求（如"帮我写个VLOOKUP公式查价格"）
        |
+-- Skills 层（自然语言指令包）---------------------+
|  skills/wps-excel/SKILL.md   教你怎么处理Excel任务  |
|  skills/wps-word/SKILL.md    教你怎么处理Word任务   |
|  skills/wps-ppt/SKILL.md     教你怎么处理PPT任务    |
|  skills/wps-office/SKILL.md  教你怎么协调跨应用任务  |
+--------------------------------------------------+
        |
+-- MCP 层（工具能力）------------------------------+
|  wps-office-mcp/                                 |
|  145 个 MCP 工具（133专业 + 12内置）               |
|  关键工具: wps_execute_method（万能调用）           |
+--------------------------------------------------+
        |
+-- WPS 执行层（加载项）----------------------------+
|  macOS: HTTP轮询(端口58891) -> WPS加载项(JS API)  |
|         224 个 action                             |
|  Windows: PowerShell COM 桥接 -> WPS Office       |
|         232 个 action                             |
+--------------------------------------------------+
```

### Skills 与 MCP 的关系

| 层级 | 职责 | 内容 |
|------|------|------|
| **Skills** | 教 AI "怎么做"：工作流程、最佳实践、参数组合 | 4 个 SKILL.md 文件，Claude Code 启动时自动加载 |
| **MCP** | 告诉 AI "能做什么"：可调用的工具和参数 | 145 个 MCP 工具，通过 MCP Server 暴露 |

---

## 项目目录结构

```
WPS_Skills/
+-- wps-office-mcp/              # MCP Server（核心服务）
|   +-- src/                     # TypeScript 源码
|   |   +-- tools/               # MCP 工具定义（133个，按 excel/word/ppt/common 分组）
|   |   +-- server/mcp-server.ts # MCP Server + 12个内置工具
|   |   +-- client/wps-client.ts # 跨平台 WPS 通信客户端
|   |   +-- index.ts             # MCP Server 入口
|   +-- scripts/wps-com.ps1      # Windows COM 桥接脚本（232 个 action）
|   +-- dist/                    # 编译输出（npm run build 生成）
|   +-- package.json
+-- wps-claude-assistant/        # WPS 加载项（macOS）
|   +-- main.js                  # HTTP 轮询 + 224 个 action dispatch
|   +-- handlers/                # Excel/Word/PPT handler 实现
|   +-- manifest.xml             # 加载项清单
|   +-- ribbon.xml               # WPS 功能区配置
+-- wps-claude-addon/            # WPS 加载项（Windows）
|   +-- js/main.js               # 加载项逻辑
|   +-- manifest.xml             # 加载项清单
|   +-- ribbon.xml               # 功能区配置
+-- skills/                      # Claude Skills 定义
|   +-- wps-excel/SKILL.md       # Excel 技能
|   +-- wps-word/SKILL.md        # Word 技能
|   +-- wps-ppt/SKILL.md         # PPT 技能
|   +-- wps-office/SKILL.md      # 跨应用协调技能
+-- scripts/
|   +-- auto-install-mac.sh      # macOS 一键安装脚本
|   +-- auto-install.ps1         # Windows 一键安装脚本
+-- INSTALL.md                   # AI 自动安装指南（逐步骤执行）
+-- README.md                    # 本文件
```

---

## 跨平台差异

| 项目 | macOS | Windows |
|------|-------|---------|
| action 数量 | 224 个 | 232 个（含 8 个 Windows 扩展） |
| 执行方式 | HTTP 轮询 + WPS JS 加载项 | PowerShell COM 桥接 |
| MCP 工具 | 145 个（共享） | 145 个（共享） |

---

## 系统要求

| 项目 | Windows | macOS |
|------|---------|-------|
| 操作系统 | Windows 10/11 | macOS 12+ |
| WPS Office | 2019 或更高 | Mac 版最新版 |
| Node.js | >= 18.0.0 | >= 18.0.0 |
| AI 工具 | Claude Code / Cursor / Augment 等 | Claude Code / Cursor / Augment 等 |

---

## 常见问题

### Claude 助手选项卡未出现

1. 确认加载项文件夹名称以 `_` 结尾
2. 确认 `publish.xml` 已正确配置，包含加载项注册条目
3. 强制退出并重启 WPS Office

### Skills 未加载（仅 Claude Code）

检查软链接是否存在：
```bash
ls ~/.claude/skills/
```

如果为空，手动创建：
```bash
PROJECT_DIR=/path/to/WPS_Skills
mkdir -p ~/.claude/skills
ln -sf $PROJECT_DIR/skills/wps-excel ~/.claude/skills/wps-excel
ln -sf $PROJECT_DIR/skills/wps-word ~/.claude/skills/wps-word
ln -sf $PROJECT_DIR/skills/wps-ppt ~/.claude/skills/wps-ppt
ln -sf $PROJECT_DIR/skills/wps-office ~/.claude/skills/wps-office
```

然后重启 Claude Code（必须重启才能加载 Skills）。

### MCP Server 连接失败

1. 确认已执行 `cd wps-office-mcp && npm install && npm run build`
2. 运行 `claude mcp list` 检查 `wps-office` 是否已注册
3. 重启 Claude Code / IDE

---

## 许可证

MIT License

## 开发者

**lc2panda** - [GitHub](https://github.com/lc2panda)

---

> **AI读者须知**：你正在阅读 WPS Skills 项目。这是一个让 AI 编程助手（Claude Code、Cursor、Augment 等）能通过自然语言操控 WPS Office（Excel、Word、PPT）的工具集。安装后，你将获得 145 个 MCP 工具和 macOS 224 / Windows 232 个底层 action 能力。通过 `wps_execute_method` 万能工具可访问全部底层 action。用户可以用中文直接描述需求，你自动调用对应工具执行。

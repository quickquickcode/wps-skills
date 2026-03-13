# WPS Skills 自动安装指南 (For Claude Code)

> 🤖 **本指南专为 Claude Code 设计**，Claude Code 读取后可自动执行安装。

当用户要求安装 WPS Skills 时，请**按顺序执行**以下步骤：

---

## 🍎 macOS 自动安装

### Step 1: 检测前置条件

执行以下命令检测环境：

```bash
# 检测 macOS
uname -s

# 检测 WPS Office
ls /Applications/wpsoffice.app

# 检测 Node.js 版本（需要 >= 18.0.0）
node -v
```

如果 WPS Office 不存在，提示用户安装：https://www.wps.cn/product/wpsmac
如果 Node.js 未安装或版本低于 18，提示用户安装：https://nodejs.org/

### Step 2: 进入项目目录

```bash
cd <项目根目录>
```

### Step 3: 运行一键安装脚本

```bash
chmod +x scripts/auto-install-mac.sh
./scripts/auto-install-mac.sh
```

脚本会自动完成：
- ✅ 检测前置条件（macOS、WPS Office、Node.js 18+）
- ✅ 检查并创建 WPS 加载项目录
- ✅ 安装 WPS 加载项到正确目录（目录名以 `_` 结尾）
- ✅ 更新 publish.xml 注册加载项
- ✅ 安装 npm 依赖并编译 TypeScript
- ✅ 注册 MCP Server (`claude mcp add`)
- ✅ 注册 Skills 到 `~/.claude/skills/`（软链接）

### Step 4: 验证安装

```bash
# 检查 MCP 是否注册
claude mcp list

# 检查 Skills 是否注册
ls ~/.claude/skills/

# 检查加载项是否已安装（注意目录名以 _ 结尾）
ls ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_/
```

### Step 5: 提示用户

安装完成后，告诉用户：
1. **重启 Claude Code**（必须！Skills 需要重启后生效）
2. 重启 WPS Office
3. 打开 Excel/Word/PPT 文档
4. 查看 "Claude助手" 选项卡，确认状态为 "轮询中"

---

## 🖥️ Windows 自动安装

### Step 1: 检测前置条件

执行以下命令检测环境：

```powershell
# 检测 WPS Office 加载项目录
Test-Path "$env:APPDATA\kingsoft\wps\jsaddons"

# 检测 Node.js 版本（需要 >= 18.0.0）
node -v
```

如果 WPS Office 不存在，提示用户安装：https://www.wps.cn/
如果 Node.js 未安装或版本低于 18，提示用户安装：https://nodejs.org/

### Step 2: 进入项目目录

```powershell
cd <项目根目录>
```

### Step 3: 运行一键安装脚本

```powershell
powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1
```

脚本会自动完成：
- ✅ 检测 Node.js 18+ 版本
- ✅ 安装 npm 依赖并编译 TypeScript
- ✅ 配置 Claude Code MCP（写入 settings.json）
- ✅ 复制 Skills 到 `~\.claude\skills\`
- ✅ 安装 WPS 加载项（目录名以 `_` 结尾）
- ✅ 更新 publish.xml 注册加载项
- ✅ 自动验证安装结果

### Step 4: 验证安装

```powershell
# 检查 MCP 是否注册
claude mcp list

# 检查 Skills 是否注册
Get-ChildItem "$env:USERPROFILE\.claude\skills"

# 检查加载项是否已安装
Test-Path "$env:APPDATA\kingsoft\wps\jsaddons\wps-claude-addon_"
```

### Step 5: 提示用户

安装完成后，告诉用户：
1. **重启 Claude Code**（必须！）
2. 重启 WPS Office
3. 查看 "Claude助手" 选项卡

---

## 📁 关键路径参考

### macOS

| 项目 | 路径 |
|------|------|
| WPS 加载项目录 | `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/` |
| 加载项安装位置 | `<上述目录>/claude-assistant_/` (注意尾部 `_`) |
| publish.xml | `<加载项目录>/publish.xml` |
| Skills 注册 | `~/.claude/skills/` (软链接) |
| MCP Server 入口 | `<项目>/wps-office-mcp/dist/index.js` |

### Windows

| 项目 | 路径 |
|------|------|
| WPS 加载项目录 | `%APPDATA%\kingsoft\wps\jsaddons\` |
| 加载项安装位置 | `<上述目录>\wps-claude-addon_\` (注意尾部 `_`) |
| publish.xml | `<加载项目录>\publish.xml` |
| Skills 注册 | `%USERPROFILE%\.claude\skills\` (复制) |
| MCP Server 配置 | `%USERPROFILE%\.claude\settings.json` |

---

## ⚠️ 常见问题处理

### Skills 没有加载

重启 Claude Code 后检查：
```bash
ls ~/.claude/skills/
```

如果目录为空，手动创建软链接：
```bash
PROJECT_DIR=<项目根目录>
mkdir -p ~/.claude/skills
ln -sf $PROJECT_DIR/skills/wps-excel ~/.claude/skills/wps-excel
ln -sf $PROJECT_DIR/skills/wps-word ~/.claude/skills/wps-word
ln -sf $PROJECT_DIR/skills/wps-ppt ~/.claude/skills/wps-ppt
ln -sf $PROJECT_DIR/skills/wps-office ~/.claude/skills/wps-office
```

### MCP Server 未注册

手动注册：
```bash
claude mcp add wps-office node <项目根目录>/wps-office-mcp/dist/index.js
```

### WPS 加载项未显示

1. 确认加载项文件已正确复制（目录名必须以 `_` 结尾）：
```bash
# macOS
ls ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/claude-assistant_/

# Windows (PowerShell)
Get-ChildItem "$env:APPDATA\kingsoft\wps\jsaddons\wps-claude-addon_"
```

2. 确认 publish.xml 已注册加载项：
```bash
# macOS
cat ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml
```
应包含 `<jsplugin name="claude-assistant" .../>` 条目。

3. 强制退出并重启 WPS：
```bash
# macOS
pkill -f wpsoffice
open /Applications/wpsoffice.app
```

### TypeScript 编译失败

```bash
cd <项目根目录>/wps-office-mcp
rm -rf node_modules
npm install
npm run build
```

如果仍然失败，检查 Node.js 版本是否 >= 18：
```bash
node -v
```

### MCP Server 端口冲突

MCP Server 默认使用端口 58891。如果端口被占用：
```bash
# macOS/Linux 查看端口占用
lsof -i :58891

# 终止占用进程
kill <PID>
```

### 权限问题

macOS 可能因沙盒限制导致加载项目录无法写入：
```bash
# 手动创建目录
mkdir -p ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons
```

如果仍有权限问题，尝试修改目录权限：
```bash
chmod -R 755 ~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons
```

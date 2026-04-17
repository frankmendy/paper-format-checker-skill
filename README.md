<h1 align="center">📄 Paper Format Checker Skill</h1>

<p align="center">
  <strong>基于"网络空间安全学院论文格式模板"的自动化排版修正工具</strong>
</p>

<p align="center">
  <a href="CHANGELOG.md"><img src="https://img.shields.io/badge/version-2.0-blue.svg" alt="Version"></a>
  <a href="#"><img src="https://img.shields.io/badge/python-3.8+-green.svg" alt="Python"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-yellow.svg" alt="License"></a>
</p>

---

## 📝 简介 (Introduction)

**Paper Format Checker** 是一款专为学术论文排版打造的自动化修正技能。它能够一键接管并修复 Word (`.docx`) 文档中繁杂的格式问题，让你从无尽的排版调整中解放出来，将更多精力投入到学术创作本身。

本技能通过智能识别论文结构，自动修正字体、字号、对齐方式、缩进、行距、页码以及图表题注，确保最终输出的文档 100% 符合官方格式模板的要求。

## 📦 安装指南 (Installation)

### 环境要求

- **操作系统**: Windows 10/11 (需要 Microsoft Word COM 接口支持)
- **Python**: 3.8 或更高版本
- **Microsoft Word**: 2016 或更高版本（推荐）

### 依赖安装

在使用本技能前，请确保已安装必要的 Python 依赖包：

```bash
pip install python-docx pywin32
```

**依赖说明：**
- `python-docx`: 用于读取和写入 Word 文档
- `pywin32`: 用于 Word COM 自动化（更新目录、刷新域等）

### AI Agent 集成

本技能设计为 AI Agent 的扩展能力，推荐在 [Trae](https://www.trae.ai/) 或 [OpenClaw](https://github.com/openclaw/openclaw) 等支持 Tool/Skill 机制的 AI 编程助手或代理环境中使用。

### 方式一：在 Trae 中安装（推荐）
1. **下载 Trae**：前往 [Trae 官网 (trae.ai)](https://www.trae.ai/) 下载并安装适用于 Windows 的最新版 IDE。
2. **导入技能**：
   - 打开 Trae 的 AI 对话侧边栏。
   - 点击聊天框上方的 **Skills**（技能）按钮，或者直接在聊天框输入 `/skill`。
   - 在弹出的技能面板中，选择通过 GitHub/Gitee 仓库 URL 导入。
   - 填入本仓库的地址（任选其一）：
     - GitHub: `https://github.com/frankmendy/paper-format-checker-skill`
     - Gitee: `https://gitee.com/frankmendy/paper-format-checker-skill`
3. 安装完成后，即可在任意项目中随时唤起它为你排版论文。

### 方式二：在 OpenClaw 等其他 AI Agent 中安装
我们也强烈推荐在强大的开源 AI Agent 框架（如 **OpenClaw** 等）中挂载并使用本技能！
- **安装步骤**：将本仓库克隆到本地，或根据你使用的 Agent 平台的扩展规范，将本仓库的 `SKILL.md` 和对应的 Python 脚本注册为一项本地 Tool/Skill。
- **⚠️ 环境注意（跨平台必看）**：本技能依赖于底层的 **Windows 操作系统** 和 **Microsoft Word COM 接口** 来实现诸如“自动刷新目录页码”、“修复前导点”等高级排版功能。
  - **如果你的 OpenClaw/Agent 运行在 Linux、macOS 或 WSL 无头环境中**：请在调用执行脚本时，加上 `--no_word_update` 参数以跳过强制调用 Word 的阶段。这样工具依然能完成 90% 的排版（正文、标题、摘要等），你只需将最终的 `_fixed.docx` 下载到 Windows 本地，打开并手动按 `F9` 键更新一次目录即可！

## 📖 使用教程 (Tutorial)

使用本技能非常简单，只需通过自然语言与 Trae 交互即可完成所有排版工作。

### 第一步：准备文档
确保你的论文是 `.docx` 格式，且已经包含了基本的章节标题、正文、图片和表格。不需要你手动去调行距、字体等细节，只要内容完整即可。

### 第二步：唤起技能并执行
在 Trae 的对话框中，直接告诉 AI 你需要排版论文，并提供文件的**绝对路径**。
例如，你可以这样发送消息：
> "帮我用 paper_format_checker 修复一下这篇论文的格式：`D:\论文\我的毕业论文.docx`"

或者直接使用指令唤起：
> "`/skill paper_format_checker` 请处理文件 `D:\论文\我的毕业论文.docx`"

### 第三步：等待处理与获取结果
AI 接收到指令后，会自动在后台启动 Word 进程进行处理。整个过程包括：
1. **摘要与标题处理**：修正中英文标题、摘要和关键词的字体与缩进。
2. **正文排版**：对各级标题（第一章、1.1、1.1.1 等）进行字体加粗、字号调整和缩进处理，同时将正文段落统一为首行缩进 2 字符、固定行距 20 磅。
3. **图表与题注**：自动将图片居中，将图/表题注调整为五号楷体/Times New Roman 并居中对齐。
4. **目录与页码重建**：强制刷新目录域，确保所有目录项拥有正确的右对齐前导点（`......`）和加粗层级。

处理完成后，AI 会在原文件所在目录下生成一个名为 `原文件名_fixed.docx` 的新文件（原文件不会被覆盖）。你可以直接打开它查看完美排版后的最终效果！

## 🎯 论文应具备的内容 (Required Paper Structure)

本技能的自动化修正依赖于论文的标准化结构。为了让工具能够准确识别并应用相应的格式规则，你的论文**必须**包含以下三个核心部分：

1. **摘要与关键词部分 (Abstract & Keywords)**
   - 包含中文论文题目、中文摘要标签与正文、中文关键词。
   - 包含英文论文题目、英文 Abstract 标签与正文、英文 Key words。
   
2. **目录部分 (Table of Contents)**
   - 包含“目录”标题。
   - 包含各级章节标题及对应页码（推荐使用 Word 自动生成的目录域，即使格式混乱也会被自动重置）。
   
3. **正文与结尾部分 (Body & Appendices)**
   - **正文**：以“第一章”或相应一级标题为起点的核心内容区域，包含标准层级（如 1.1, 1.1.1）的小节。
   - **图表**：插入在正文中的图片、表格及其对应的图名、表名。
   - **特殊章节**：文档末尾需包含“致谢”、“参考文献”等标志性收尾章节（如有“附录”也需包含在内）。

---

## 💡 最佳写作要求 (Best Practices)

为了让格式修复达到最完美的“一键出图”效果，建议在撰写初稿时遵循以下规范：

- **🔖 结构清晰，合理分页**
  摘要、目录、正文各主要部分之间，请使用 **分页符 (Page Break)** 或 **分节符 (Section Break)** 进行分隔，绝对不要使用连续敲击回车键来挤出新页。
- **🔢 使用标准的标题编号**
  章节标题请采用标准的 `第一章`、`1.1`、`1.1.1` 格式。标题编号与标题文字之间建议**保留一个空格**（例如：`1.1 研究背景`）。
- **🖼️ 图表和题注位置规范**
  图片请设置为“嵌入型 (In Line with Text)”。**图名必须写在图片附近，表名必须写在表格附近**，中间尽量不要插入多余的空段落。
- **🚫 避免过度复杂的自定义样式**
  尽量使用 Word 默认的 `Normal` (正文) 样式进行书写，不要手动在段落属性中设置复杂的“悬挂缩进”、“多级列表”或“制表位”，修复工具会自动接管一切。
- **📑 生成自动目录**
  提交修复前，请使用 Word 内置功能插入一次“自动目录”（引用 -> 目录 -> 自动目录）。

---

## 🚀 如何使用 (Usage)

在对话中直接唤起本技能，并告诉它你需要修复的论文路径即可。例如：

> **User:** “帮我修复一下这篇论文的格式：`D:\论文\我的毕业论文.docx`”
> 
> **Skill:** 工具将自动加载并处理文档，修复完成后，会在同一目录下生成一个带有 `_fixed` 后缀的全新文档。

---

## 📋 版本历史 (Changelog)

查看 [CHANGELOG.md](CHANGELOG.md) 了解详细的版本更新历史。

### v2.0 (2025-04-17)
- ✨ **重大更新**: 修复二、三级标题段前段后"自动"配置问题
- 🔧 **新功能**: 添加依赖检查机制，启动时自动检查必要依赖
- 📚 **文档**: 更新格式规则文档，添加新的排坑经验

### v1.0 (初始版本)
- ✨ **基础功能**: 自动修复论文格式，匹配官方模板
- 🎯 **三阶段修复**: 摘要、目录、正文与题注
- 🔄 **循环复检**: 确保修复质量

## 🤝 参与贡献 (Contributing)

我们欢迎任何形式的贡献！如果你在使用过程中发现了可以优化的逻辑，或是想增加对更多特殊格式的支持，可以通过以下方式参与贡献：

1. **Fork** 本项目
2. 创建你的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交你的更改 (`git commit -m 'Add some AmazingFeature'`)
4. 将你的修改推送到分支 (`git push origin feature/AmazingFeature`)
5. 发起一个 **Pull Request**

*在提交代码前，请确保你已经阅读并理解了项目的代码规范，并通过了本地的格式校验测试。*

---

## 🐛 问题反馈 (Issues)

由于不同版本的 Word 差异以及每篇论文的原始排版千奇百怪，工具偶尔可能会遇到无法完美处理的边缘情况（Edge Cases）。

如果你在使用中遇到了问题，例如：
- 某段正文被错误识别为标题
- 图表题注的缩进未能完全清除
- 运行脚本时出现 COM 接口死锁或权限报错

请前往 [Issues 页面](../../../../issues) 提交问题报告。在提交时，请尽量附上：
1. 你的操作系统和 Word 版本
2. 出现报错的完整控制台日志
3. 如果方便，请附上引发问题的 `.docx` 样例文件（注意脱敏，删除个人隐私内容）

---

## 📄 开源协议 (License)

本项目基于 **MIT License** 开源。
这意味着你可以自由地使用、修改、分发本工具，甚至用于商业用途。详情请参阅项目根目录下的 `LICENSE` 文件。

---

<p align="center">
  <i>Let the tool handle the formatting, you focus on the research.</i>
</p>
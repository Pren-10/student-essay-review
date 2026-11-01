# 学生作文点评与修改 Skill

把学生作文照片、PDF 或文字整理成“作文点评 + 修改后参考文章”的 Word 文件。先确认需求，再保留原文主线完成细节、心理与语言修改，核验字数后套用固定模板。

本 skill 适用于新东方（XDF）教育场景的作文修改规范，面向日常教学和教师批改工作，将识别、点评、修改、字数核验与 Word 排版整合为高效流程，生成结构清晰、排版直观、便于教师携带和分享的作文修改文件。这里的“适用于”指工作场景与流程设计，不代表新东方官方授权或官方发布的制度文件。

## 使用

将本仓库作为 `student-essay-review` 文件夹放入支持 SKILL.md 的助手技能目录，或在聊天中提供本文件夹的 SKILL.md 路径。

触发示例：

- “按作文工作流处理这篇。”——先询问要求。
- “按默认要求处理这份 PDF。”——直接按默认规则执行。
- “五年级，保留原标题，正文 650—750 字，重点补心理变化。”——按本次要求执行。

默认正文 650—750 字，点评约 200 字另算；保留原题。可覆盖字数、标题、年级和改写程度。文件夹批处理约定见 [SKILL.md](SKILL.md)。

## 依赖和范围

- 作文阅读、点评与改写由能读取图片/文件的 AI 助手完成，脚本只负责格式与计数。
- Word 生成：Windows PowerShell 5.1+、.NET，无第三方包下载。
- PDF 转图片：Windows PowerShell 5.1 和系统 Windows PDF 组件。
- Word 预览：本机安装 Microsoft Word；后台只读打开成品并导出 PDF。
- 模板中文字体为华文宋体（STSong）。缺少该字体时可能发生替代，需检查实际显示；本仓库不分发字体。
- 不含学生原件、学生成品、账号配置、密码、API key 或完整聊天记录。

## 手动生成与验证

准备 JSON（字段见 SKILL.md），然后运行：

```powershell
& './scripts/build-docx.ps1' -ContentPath './work/content.json' -OutputPath './outputs/修改稿.docx'
& './scripts/render-word.ps1' -InputPaths './outputs/修改稿.docx' -OutputDir './work/preview'
& './scripts/render-pdf.ps1' -InputPdf './work/preview/修改稿.pdf' -OutputPrefix './work/preview/page'
```

`render-pdf.ps1` 请在 Windows PowerShell 5.1 进程中运行。按设备的脚本执行策略授权运行，不需要永久修改系统策略。默认计数同时验证含标点/不含标点，可用 `-CountMode WithPunctuation` 或 `WithoutPunctuation` 明确选择。输出文件存在时拒绝覆盖。

Word 模板保留黄色栏目高亮、A4 页面、16 磅正文、20 磅居中加粗标题和首行两字缩进。模板已去除原学生内容和个人元数据。文字和长度不同，页数也可能不同。

## 验证情况

公开包已实测：有效 JSON 生成 Word、默认两种字数限制、拒绝覆盖已有成品、拒绝超出字数要求的内容、两页 PDF 转 PNG，以及模板 XML 和个人信息清理检查。

当前受限会话启动 Word COM 返回 `80070520`（登录会话不可用），因此本次公开包未完成新的 Word 视觉渲染验证。Word 导出脚本需要在可访问桌面 Word 的 Windows 用户会话中运行。此前工作流试做通过不等于此包在所有设备上已验证。

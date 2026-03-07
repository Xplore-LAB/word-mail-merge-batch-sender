# Word Mail Merge Batch Sender

> 让 Word 邮件合并支持“按行自动添加附件”的 VBA 批量邮件工具。  
> A simple VBA tool that adds attachment support to Word Mail Merge batch emails.

<p align="center">
  <a href="https://github.com/Xplore0114/word-mail-merge-batch-sender/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/Xplore0114/word-mail-merge-batch-sender?style=flat-square"></a>
  <a href="https://github.com/Xplore0114/word-mail-merge-batch-sender/blob/main/LICENSE"><img alt="License" src="https://img.shields.io/github/license/Xplore0114/word-mail-merge-batch-sender?style=flat-square"></a>
  <a href="https://github.com/Xplore0114/word-mail-merge-batch-sender"><img alt="Platform" src="https://img.shields.io/badge/platform-Windows%20%2B%20Office-blue?style=flat-square"></a>
</p>

这是一个面向国内用户的 Office VBA 小工具，主要解决一个非常实际的问题：

**Word 自带的 Mail Merge 可以批量发邮件，但不能为每一封邮件自动添加对应附件。**

这个项目基于 `Word + Outlook + Excel + VBA`，在保留 Word 模板排版和邮件合并能力的基础上，为每封邮件自动匹配并附加对应文件。

如果这个项目对你有帮助，欢迎点一个 Star。

---

## 这个项目解决什么问题

很多老师、培训机构、HR、行政、财务都会遇到同一种需求：

- 每封邮件正文不一样
- 每封邮件附件也不一样
- 还想继续用熟悉的 Word 模板和 Excel 名单

例如：

- 给家长批量发送作业或成绩单
- 给客户批量发送报价单、合同、对账单
- 给员工批量发送 offer、通知、资料包
- 给学员批量发送课程安排和附件材料

Word 原生邮件合并擅长“批量正文”，但不擅长“批量附件”。

这个项目就是专门补上这一块。

---

## 为什么值得用

- 继续使用熟悉的 `Word Mail Merge`
- 继续使用 Excel 维护名单
- 继续通过 Outlook 发邮件
- 不需要额外安装复杂系统
- 每封邮件可以自动带不同附件
- 适合普通办公用户，不是只给程序员用的

---

## 核心功能 Features

- 使用 Word Mail Merge 生成个性化邮件正文
- 使用 Outlook 批量发送邮件
- 支持按 Excel 每一行自动附加对应附件
- 支持多 Outlook 账号选择
- 支持两种模式
  - 直接发送 `Send`
  - 仅生成草稿 `Draft`
- 发送前检查常见问题
- 发送失败时给出错误提示
- 预留错误报告扩展空间

---

## 最大亮点：附件自动匹配

这是 Word 原生邮件合并最常见、也最让人头疼的短板。

假设你的 Excel 数据如下：

| Email_Addresses | Attachment |
|---|---|
| parent1@email.com | C:\Homework\math.pdf |
| parent2@email.com | C:\Homework\science.pdf |

发送后：

- 发给 `parent1@email.com` 的邮件自动附带 `math.pdf`
- 发给 `parent2@email.com` 的邮件自动附带 `science.pdf`

也就是说：

**正文跟着 Mail Merge 走，附件跟着 Excel 每一行走。**

这就是这个项目的核心价值。

---

## 和 Word 原生邮件合并有什么区别

| 功能 | Word Mail Merge | 本项目 |
|---|---|---|
| 批量邮件 | Yes | Yes |
| 个性化正文 | Yes | Yes |
| 保留 Word 模板排版 | Yes | Yes |
| 自动添加附件 | No | Yes |
| Outlook 账号选择 | No | Yes |
| 发送前检查 | No | Yes |
| 错误提示 | No | Yes |
| 草稿模式 | Limited | Yes |

---

## 3 分钟看懂使用流程

### 第一步：准备 Word 模板

你照常用 Word 写邮件模板，并插入邮件合并字段。

例如：

```text
尊敬的「Parent_Name」家长：
您好！
这是关于学生「Student_Name」的一条学习提醒。
本周作业内容：
「Homework_Task」
提交截止时间：
「Deadline」
请您提醒孩子按时完成作业。
感谢您的配合！
班主任
```

模板文件建议放在：

```text
template/mail_template.docm
```

> 模板必须是 `.docm`，因为需要运行宏。

### 第二步：准备 Excel 名单

Excel 至少包含以下字段：

| 字段 | 是否必须 | 说明 |
|---|---|---|
| Email_Addresses | 必须 | 收件人邮箱 |
| Attachment | 可选 | 附件完整路径 |

其他列都可以作为 Word 邮件模板变量使用。

示例：

| Email_Addresses | Parent_Name | Student_Name | Homework_Task | Deadline | Attachment |
|---|---|---|---|---|---|
| parent1@email.com | 张先生 | 张小明 | 数学练习册第12-15页 | 2026-03-20 | C:\Homework\math.pdf |
| parent2@email.com | 李女士 | 李小红 | 科学实验报告 | 2026-03-22 | C:\Homework\science.pdf |

示例文件建议放在：

```text
example/sample_data.xlsx
```

### 第三步：在 Word 中连接数据源

```text
邮件 -> 选择收件人 -> 使用现有列表
```

然后选择你的 Excel 文件。

### 第四步：插入合并域

```text
邮件 -> 插入合并域
```

例如：

```text
«Parent_Name»
«Student_Name»
«Homework_Task»
```

### 第五步：运行宏

```text
Alt + F8
```

运行：

```text
WDSE_Final_AutoSend
```

### 第六步：按提示发送

程序通常会提示你：

1. 选择 Outlook 发件账号
2. 输入邮件标题
3. 选择发送模式
4. 查看发送前扫描结果
5. 确认发送

---

## 适合哪些人用

- 学校老师 / 教务老师
- 培训机构运营人员
- HR / 行政 / 招聘人员
- 财务 / 客服 / 销售支持
- 需要批量发邮件但不想折腾复杂系统的人

如果你平时已经在用 Word、Excel、Outlook，这个工具会比较顺手。

---

## 适用环境 Requirements

- Windows
- Microsoft Word
- Microsoft Outlook
- Microsoft Excel
- 启用 VBA 宏支持

> 这是一个基于 Office VBA 的桌面工具，不适用于网页端 Office。

---

## 仓库结构 Repository Structure

```text
word-mail-merge-batch-sender
│
├─ README.md
├─ LICENSE
├─ CHANGELOG.md
│
├─ src
│  └─ WordBatchMailSender.bas
│
├─ template
│  └─ mail_template.docm
│
├─ example
│  └─ sample_data.xlsx
│
└─ docs
   └─ screenshot.png
```

---

## 当前状态 Current Status

当前仓库先发布的是占位版，用来先公开项目结构和说明文档。

后续会逐步补充：

- `src/WordBatchMailSender.bas`
- `template/mail_template.docm`
- `example/sample_data.xlsx`
- `docs/screenshot.png`

如果你现在看到的是占位文件，不是你看错了，是我故意先把“门面”和结构搭起来了，方便项目先公开、先被看到、后续再补完整内容。

---

## 常见问题 Troubleshooting

### 1. 邮件无法发送

可能原因：

- Outlook 没有打开
- Outlook 没有登录可用账号
- 本机安全策略阻止宏调用邮件发送

建议处理：

- 先打开 Outlook 并确认可以正常收发邮件
- 重新运行宏
- 检查 Office 宏安全设置

### 2. 附件发送失败

可能原因：

- Excel 中的附件路径错误
- 附件文件不存在
- 使用了相对路径，而不是完整路径

建议处理：

- 使用完整路径，例如 `C:\Homework\math.pdf`
- 确认文件真实存在
- 尽量不要改动附件目录结构

### 3. 邮件字段没有替换

可能原因：

- Excel 列名与 Word 合并字段名称不一致

建议处理：

- 删除原字段
- 重新通过“插入合并域”插入一次
- 确认 Excel 表头没有多余空格

### 4. 宏无法运行

可能原因：

- Word 禁止宏执行
- 文件不是 `.docm`
- VBA 模块未正确导入

建议处理：

- 启用宏
- 确认模板文件为 `.docm`
- 检查 `src/WordBatchMailSender.bas` 是否已正确导入

### 5. 为什么不用第三方群发系统

因为很多实际办公场景里，用户已经在使用：

- Word 写模板
- Excel 管数据
- Outlook 发邮件

这个项目的思路不是替换你的办公习惯，而是在你原本的流程上补上“自动附件”这个关键能力。

---

## 截图 Screenshot

当前截图仍为占位状态，后续建议补充以下内容：

- 发送前扫描窗口
- Outlook 账号选择窗口
- 直接发送 / 生成草稿效果图
- Excel 示例数据截图

占位路径：

```text
docs/screenshot.png
```

---

## Roadmap

- [x] 公开仓库与中文 README
- [x] 发布占位版开源说明
- [ ] 补充真实 VBA 源码
- [ ] 补充 Word 模板
- [ ] 补充 Excel 示例数据
- [ ] 补充操作截图
- [ ] 增加更完整的错误报告说明
- [ ] 增加更详细的使用演示

---

## 对国外用户的简短说明

This repository is mainly written for Chinese users.

But the project itself is simple:

- Use Word Mail Merge for personalized email body
- Use Outlook to send emails
- Use Excel to map recipients and attachment paths
- Add one attachment per row automatically

If needed, an English README can be added later.

---

## License

MIT License

---

## GitHub Topics

`vba`, `word-mail-merge`, `outlook`, `batch-email`, `office-automation`, `email-automation`

---

## One-line Pitch

一个让 Word 邮件合并支持“按行自动添加附件”的 VBA 批量邮件工具。

A simple VBA tool that adds attachment support to Word Mail Merge batch emails.

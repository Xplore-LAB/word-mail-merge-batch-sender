# Word Mail Merge Batch Sender

一个用于解决 `Word 邮件合并（Mail Merge）无法自动添加附件` 问题的 VBA 工具。

A simple VBA tool for sending batch emails with attachments using Word Mail Merge and Outlook.

这个项目面向国内用户，文档以中文为主，同时保留必要英文，方便 GitHub 搜索、海外用户理解，以及后续协作。

如果这个项目对你有帮助，欢迎点一个 Star。

---

## 为什么做这个项目

Word 自带的邮件合并可以批量生成个性化邮件正文，但有一个很常见的痛点：

- 不能为每一封邮件自动附加对应附件

而现实里很多场景都需要“正文个性化 + 每封邮件带不同附件”，例如：

- 学校给家长发送作业或成绩单
- 培训机构批量发送通知和资料
- HR 批量发送 offer、合同或入职材料
- 财务批量发送账单、回执、对账单

这个项目就是为了解决这个问题。

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
- 可扩展错误报告能力

---

## 最大亮点 Highlight

**支持批量邮件自动附带附件。**

这是 Word 原生邮件合并最不方便的一点，也是本项目最核心的价值。

例如 Excel 中这样配置：

| Email_Addresses | Attachment |
|---|---|
| parent1@email.com | C:\Homework\math.pdf |
| parent2@email.com | C:\Homework\science.pdf |

发送后：

- 第一封邮件自动带上 `math.pdf`
- 第二封邮件自动带上 `science.pdf`

---

## 与 Word 原生邮件合并的区别

| 功能 | Word Mail Merge | 本项目 |
|---|---|---|
| 批量邮件 | Yes | Yes |
| 个性化正文 | Yes | Yes |
| 保留 Word 模板排版 | Yes | Yes |
| 自动添加附件 | No | Yes |
| Outlook 账号选择 | No | Yes |
| 发送前检查 | No | Yes |
| 错误提示 | No | Yes |

---

## 适用环境 Requirements

- Microsoft Word
- Microsoft Outlook
- Microsoft Excel
- Windows
- 启用 VBA 宏支持

> 说明：这是一个基于 Office VBA 的桌面工具，不适用于网页端 Office。

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

## Excel 数据结构 Data Format

程序最低依赖以下字段：

| 字段 | 是否必须 | 说明 |
|---|---|---|
| Email_Addresses | 必须 | 收件人邮箱 |
| Attachment | 可选 | 附件完整路径 |

其他字段都可以作为 Word 模板中的邮件合并字段使用。

### 示例字段 Example

| Email_Addresses | Parent_Name | Student_Name | Homework_Task | Deadline | Attachment |
|---|---|---|---|---|---|
| parent1@email.com | 张先生 | 张小明 | 数学练习册第12-15页 | 2026-03-20 | C:\Homework\math.pdf |
| parent2@email.com | 李女士 | 李小红 | 科学实验报告 | 2026-03-22 | C:\Homework\science.pdf |

---

## Word 模板示例 Template Example

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

---

## 使用方法 Quick Start

### 1. 打开 Word 模板

打开：

```text
template/mail_template.docm
```

> 模板必须是 `.docm` 格式。

### 2. 连接 Excel 数据源

在 Word 中：

```text
邮件 -> 选择收件人 -> 使用现有列表
```

选择：

```text
example/sample_data.xlsx
```

### 3. 插入邮件合并字段

```text
邮件 -> 插入合并域
```

例如：

```text
«Parent_Name»
«Student_Name»
«Homework_Task»
```

### 4. 运行宏

按：

```text
Alt + F8
```

运行：

```text
WDSE_Final_AutoSend
```

### 5. 按提示完成发送

程序通常会提示你：

1. 选择 Outlook 发件账号
2. 输入邮件标题
3. 选择发送模式
4. 查看发送前扫描结果
5. 确认发送

---

## 占位文件 Placeholder Files

当前仓库已先公开占位版文档，方便先上线项目。

后续你需要补充这些真实文件：

- `src/WordBatchMailSender.bas`
- `template/mail_template.docm`
- `example/sample_data.xlsx`
- `docs/screenshot.png`

为了避免用户困惑，当前仓库中的这些文件会先使用占位说明。

---

## 常见问题 Troubleshooting

### 1. 邮件无法发送

可能原因：

- Outlook 没有打开
- Outlook 没有登录可用账号
- 本机安全策略阻止宏调用邮件发送

建议处理：

- 先打开 Outlook 并确认可正常收发邮件
- 重新运行宏
- 检查 Office 宏安全设置

### 2. 附件发送失败

可能原因：

- Excel 中的附件路径错误
- 附件文件不存在
- 使用了相对路径而不是完整路径

建议处理：

- 使用完整路径，例如 `C:\Homework\math.pdf`
- 确认文件实际存在

### 3. 邮件字段没有替换

可能原因：

- Excel 列名和 Word 合并字段名称不一致

建议处理：

- 删除原字段
- 重新通过“插入合并域”插入一次

### 4. 宏无法运行

可能原因：

- Word 禁止宏执行
- 文件不是 `.docm`
- VBA 模块未正确导入

建议处理：

- 启用宏
- 确认模板文件为 `.docm`
- 检查 `src/WordBatchMailSender.bas` 是否已正确导入

---

## 发布计划 Roadmap

- [x] 公开仓库与中文 README
- [ ] 补充 VBA 源码
- [ ] 补充 Word 模板
- [ ] 补充 Excel 示例数据
- [ ] 补充操作截图
- [ ] 增加更完整的错误报告说明

---

## 截图 Screenshot

当前为占位状态，后续建议补充：

- 发送前扫描窗口
- 账号选择窗口
- 邮件发送 / 草稿生成效果图

占位文件路径：

```text
docs/screenshot.png
```

---

## License

MIT License

---

## 适合 GitHub 搜索的关键词 Topics

`vba`, `word-mail-merge`, `outlook`, `batch-email`, `office-automation`, `email-automation`

---

## 一句话介绍 One-line Pitch

一个让 Word 邮件合并支持“按行自动添加附件”的 VBA 批量邮件工具。

A simple VBA tool that adds attachment support to Word Mail Merge batch emails.

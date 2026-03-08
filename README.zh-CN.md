# Word 邮件合并批量发信工具

<p align="center">
  <img src="assets/banner.svg" alt="project banner" width="100%" />
</p>

<p align="center">
  <img src="assets/icon.svg" alt="Word Mail Merge Batch Sender 图标" width="112" />
</p>

<p align="center"><strong>让 Word 邮件合并支持“按行自动添加附件”。</strong></p>

<p align="center">
  <a href="README.md">English Docs</a> · <a href="docs/architecture.md">架构说明</a> · <a href="docs/quickstart.md">快速开始</a>
</p>

![Windows](https://img.shields.io/badge/platform-Windows%20%2B%20Office-blue)
![VBA](https://img.shields.io/badge/language-VBA-1F6FEB)
![Outlook](https://img.shields.io/badge/mail-Outlook-0A66C2)
![Mail Merge](https://img.shields.io/badge/workflow-Word%20Mail%20Merge-22C55E)

这是一个很实用的 Office VBA 小工具，专门解决 Word 邮件合并里最常见的痛点之一：

**正文可以批量个性化，但附件不能按 Excel 每一行自动匹配。**

这个项目保留了大家熟悉的办公流程：

- Word 负责模板排版
- Excel 负责名单和字段
- Outlook 负责实际发信
- VBA 负责把“逐行附件”这一步补上

## 适合谁用

- 学校老师、教务、培训机构运营
- HR、行政、招聘、销售支持
- 财务、客服、通知类岗位
- 已经在用 Word + Excel + Outlook 的办公用户

## 核心功能

- 用 Word Mail Merge 生成个性化邮件正文
- 从 Excel 读取收件人和附件路径
- 每一行数据可对应不同附件
- 支持选择 Outlook 发件账号
- 支持“直接发送”或“仅生成草稿”
- 发送前检查常见问题
- 给出更清晰的失败提示

## 快速开始

### 你需要的文件

- `template/mail_template.docx`
- `example/sample_data.xlsx`
- `src/WordBatchMailSender.bas`

### 基本流程

1. 打开 Word 模板。
2. 将模板连接到 Excel 数据源。
3. 在 Word 中插入邮件合并域。
4. 在 VBA 编辑器中导入 `src/WordBatchMailSender.bas`。
5. 运行 `BatchMailMergeSend`。
6. 选择发件账号、标题和发送模式。

## 项目结构

```text
word-mail-merge-batch-sender
├── assets/
│   ├── banner.svg
│   └── icon.svg
├── docs/
│   ├── architecture.md
│   ├── quickstart.md
│   ├── screenshot.png.placeholder.txt
│   └── troubleshooting.md
├── example/
│   └── sample_data.xlsx
├── src/
│   └── WordBatchMailSender.bas
├── template/
│   └── mail_template.docx
├── CHANGELOG.md
├── LICENSE
├── README.md
└── README.zh-CN.md
```

## 文档

- `README.md`
- `docs/architecture.md`
- `docs/quickstart.md`
- `docs/troubleshooting.md`

## 限制说明

- 面向 Windows + Office 桌面场景
- 依赖本机 Outlook 可用
- 需要 Office 宏权限

## 许可证

MIT

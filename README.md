# 批量邮件发送工具 - 企业微信邮箱版

通过企业微信邮箱 SMTP 批量发送个性化邮件，支持 Excel 名单导入、模板变量替换，以及通过 IMAP 按日期导出收件箱邮件到 Excel。

## 快速启动

```bash
cd email-sender
pip install -r requirements.txt
python app.py
```

浏览器打开 http://localhost:5050 即可使用。

## 功能

- 📤 **批量发送** — Excel 名单导入，模板变量（`{{列名}}`）替换，附件按行匹配
- 🎨 **HTML 模板** — 内置参赛提醒、报名确认、结果通知等模板
- 📥 **收件箱导出** — 通过 IMAP 按日期范围拉取收件箱邮件，导出到 Excel（日期、邮件地址、回复内容）
- ⚡ **实时进度** — 后台异步发送/拉取，前端实时查看进度

## 使用步骤

### 发送邮件

1. **配置 SMTP** — 填写企业微信邮箱地址和密码，点击「测试连接」确认
2. **上传名单** — 上传包含邮箱列的 Excel 文件（.xls / .xlsx / .csv），选择邮箱所在列
3. **上传附件**（可选）— 按 Excel 中「附件文件名列」匹配发送个性化附件
4. **编辑邮件** — 选择模板或自定义 HTML 正文，用 `{{列名}}` 插入个性化变量
5. **批量发送** — 确认无误后点击发送，实时查看进度和结果

### 导出收件箱邮件

1. 在 Step 6「收件箱邮件导出」卡片中：
   - 确认 IMAP 服务器（企业微信邮箱默认 `imap.exmail.qq.com:993`）
   - 选择开始与结束日期（默认近 7 天）
   - 点击「测试连接」确认，然后点击「导出为 Excel」
2. 等待拉取完成后，点击「下载 Excel」保存文件
3. 导出字段：日期、邮件地址、发件人、主题、回复内容

> 邮箱账号和密码复用 Step 1 中的 SMTP 配置，无需重复填写。

## 企业微信邮箱配置

| 协议 | 服务器 | 端口 |
|------|--------|------|
| SMTP（发信） | smtp.exmail.qq.com | 465 (SSL) |
| IMAP（收信） | imap.exmail.qq.com | 993 (SSL) |

密码使用邮箱登录密码或客户端专用密码。客户端专用密码在企业微信邮箱「设置 → 邮箱绑定 → 安全登录」中生成。

## 环境变量

复制 `.env.example` 为 `.env`，填入真实值可自动填充前端表单。

```bash
SMTP_HOST=smtp.exmail.qq.com
SMTP_PORT=465
SENDER_EMAIL=your_email@example.com
SENDER_PASSWORD=your_password
SENDER_NAME=CSDN

# IMAP（收件箱导出）
IMAP_HOST=imap.exmail.qq.com
IMAP_PORT=993
```

## 模板变量

邮件主题和正文中可使用 `{{列名}}` 格式的变量，发送时会自动替换为 Excel 中对应行的值。

例如 Excel 中有「姓名」「邮箱」列，正文中写 `Hi {{姓名}}` 会自动替换为每个收件人的姓名。

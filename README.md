# outlook-fetcher-cli

命令行工具，用于访问 Outlook 邮箱、日历和联系人。

## 安装

```bash
# 从源码构建
cargo build --release -p outlook-fetcher-cli

# 或使用构建脚本
./build.sh  # Linux/macOS
build.bat   # Windows
```

## 用法

```bash
outlook-fetcher-cli [OPTIONS] --token <TOKEN> <COMMAND>
```

### 全局选项

| 选项 | 环境变量 | 说明 |
|------|----------|------|
| `-t, --token` | `OUTLOOK_REFRESH_TOKEN` | Microsoft OAuth 刷新令牌 |
| `-c, --client-id` | `OUTLOOK_CLIENT_ID` | Azure AD 应用程序 ID（可选） |

### 邮件命令

| 命令 | 说明 |
|------|------|
| `inbox` | 列出收件箱邮件 |
| `unread` | 列出未读邮件 |
| `unread-count` | 获取未读邮件数量 |
| `search <query>` | 按主题搜索邮件 |
| `get <id>` | 获取邮件详情 |
| `read <id>` | 标记邮件为已读 |
| `delete <id>` | 删除邮件 |
| `send <to> -s <subject> -b <body>` | 发送邮件 |
| `reply <id> -m <message>` | 回复邮件 |
| `forward <id> -t <to>` | 转发邮件 |
| `poll <since>` | 轮询新邮件 |
| `attachments <id>` | 列出附件 |
| `download -e <email_id> -a <attachment_id> -o <output>` | 下载附件 |
| `folders` | 列出邮件文件夹 |
| `me` | 显示当前用户信息 |

### 草稿命令

| 命令 | 说明 |
|------|------|
| `drafts` | 列出草稿 |
| `create-draft` | 创建草稿 |
| `send-draft <id>` | 发送草稿 |

### 日历命令

| 命令 | 说明 |
|------|------|
| `events` | 列出日历事件 |
| `event <id>` | 获取事件详情 |
| `create-event -s <subject> --start <datetime> --end <datetime>` | 创建事件 |
| `delete-event <id>` | 删除事件 |
| `accept-event <id>` | 接受邀请 |
| `decline-event <id>` | 拒绝邀请 |

### 联系人命令

| 命令 | 说明 |
|------|------|
| `contacts` | 列出联系人 |
| `contact <id>` | 获取联系人详情 |
| `create-contact --first-name <name> --last-name <name>` | 创建联系人 |
| `delete-contact <id>` | 删除联系人 |

## 示例

### 邮件操作

```bash
# 列出收件箱
outlook-fetcher-cli inbox

# 列出未读邮件
outlook-fetcher-cli unread

# 获取未读数量
outlook-fetcher-cli unread-count

# 搜索邮件
outlook-fetcher-cli search "发票"

# 获取邮件详情
outlook-fetcher-cli get "MESSAGE_ID"

# 发送邮件
outlook-fetcher-cli send "recipient@example.com" -s "主题" -b "邮件内容"

# 发送 HTML 邮件
outlook-fetcher-cli send "recipient@example.com" -s "主题" -b "<h1>Hello</h1>" --html

# 回复邮件
outlook-fetcher-cli reply "MESSAGE_ID" -m "感谢您的来信"

# 回复全部
outlook-fetcher-cli reply "MESSAGE_ID" -m "感谢" --all

# 转发邮件
outlook-fetcher-cli forward "MESSAGE_ID" -t "other@example.com" -c "请查看"

# 轮询新邮件
outlook-fetcher-cli poll "2024-01-10T00:00:00Z"

# 列出附件
outlook-fetcher-cli attachments "MESSAGE_ID"

# 下载附件
outlook-fetcher-cli download -e "MESSAGE_ID" -a "ATTACHMENT_ID" -o "./file.pdf"
```

### 草稿操作

```bash
# 列出草稿
outlook-fetcher-cli drafts

# 创建草稿
outlook-fetcher-cli create-draft -s "主题" -b "内容" -t "to@example.com"

# 发送草稿
outlook-fetcher-cli send-draft "DRAFT_ID"
```

### 日历操作

```bash
# 列出事件
outlook-fetcher-cli events

# 列出指定日期范围的事件
outlook-fetcher-cli events --start "2024-01-01" --end "2024-01-31"

# 获取事件详情
outlook-fetcher-cli event "EVENT_ID"

# 创建事件
outlook-fetcher-cli create-event -s "会议" --start "2024-01-15T10:00:00" --end "2024-01-15T11:00:00"

# 创建带地点和参与者的事件
outlook-fetcher-cli create-event -s "会议" --start "2024-01-15T10:00:00" --end "2024-01-15T11:00:00" \
  -l "会议室A" -a "user1@example.com,user2@example.com"

# 创建在线会议
outlook-fetcher-cli create-event -s "在线会议" --start "2024-01-15T10:00:00" --end "2024-01-15T11:00:00" --online

# 接受邀请
outlook-fetcher-cli accept-event "EVENT_ID"

# 拒绝邀请
outlook-fetcher-cli decline-event "EVENT_ID" -c "时间冲突"
```

### 联系人操作

```bash
# 列出联系人
outlook-fetcher-cli contacts

# 搜索联系人
outlook-fetcher-cli contacts -s "张三"

# 获取联系人详情
outlook-fetcher-cli contact "CONTACT_ID"

# 创建联系人
outlook-fetcher-cli create-contact --first-name "张" --last-name "三" -e "zhangsan@example.com"

# 创建完整联系人
outlook-fetcher-cli create-contact --first-name "张" --last-name "三" \
  -e "zhangsan@example.com" -m "13800138000" -c "公司名" -j "工程师"

# 删除联系人
outlook-fetcher-cli delete-contact "CONTACT_ID"
```

## License

MIT

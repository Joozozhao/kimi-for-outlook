# Kimi for Outlook

AI 邮件助手 Outlook Web Add-in，集成 Kimi AI 能力，直接在 Outlook 中处理邮件。

## 功能

- 📥 **邮件分类** — 自动判断优先级、类别、建议操作
- ✍️ **起草回复** — 根据邮件内容生成得体回复，支持多种语气
- 📅 **安排会议** — 提取会议要点、建议时间、生成邀请草稿
- 🎭 **调整语气** — 正式/友好/简洁/委婉/紧急，一键改写

## 技术栈

- Office.js (Outlook Add-in API)
- Kimi API (Moonshot AI)
- 纯前端静态页面，部署于 Vercel

## 部署步骤

### 1. 准备代码

本目录包含完整代码，直接推送到 GitHub：

```bash
cd outlook-addin
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/kimi-for-outlook.git
git push -u origin main
```

### 2. 部署到 Vercel

1. 访问 [vercel.com](https://vercel.com)，用 GitHub 账号登录
2. 点击 "Add New Project"
3. 导入 `kimi-for-outlook` 仓库
4. 框架预设选 **Other**，根目录留空
5. 点击 **Deploy**

部署完成后，Vercel 会给你一个域名，例如 `https://kimi-for-outlook.vercel.app`

### 3. 更新 manifest.xml

把 `manifest.xml` 中所有的 `https://PLACEHOLDER.vercel.app` 替换为你的真实域名：

```bash
# macOS
sed -i '' 's|https://PLACEHOLDER.vercel.app|https://你的域名.vercel.app|g' manifest.xml
```

需要替换的位置：
- `<IconUrl>`
- `<HighResolutionIconUrl>`
- `<AppDomain>`
- `<SourceLocation>`（2处）
- `<bt:Image>`（3处）
- `<bt:Url>`（2处）

### 4. 获取 Kimi API Key

1. 访问 [platform.moonshot.cn](https://platform.moonshot.cn)
2. 注册/登录账号
3. 进入「API Key 管理」创建新 Key
4. 首次使用插件时输入此 Key（仅保存在本地浏览器）

### 5. 安装到 Outlook

#### 方式 A：Outlook 网页版（最简单）

1. 打开 [Outlook Web](https://outlook.live.com 或 https://outlook.office.com)
2. 点击右上角齿轮 ⚙️ → **查看所有 Outlook 设置**
3. 选择 **邮件** → **自定义操作** → **添加加载项**
4. 选择 **我的加载项** → **+ 添加自定义加载项** → **从文件添加**
5. 上传更新后的 `manifest.xml`

#### 方式 B：Outlook 桌面版

1. 打开 Outlook
2. 点击 **文件** → **管理加载项**
3. 点击 **+ 我的加载项** → **添加自定义加载项** → **从文件添加**
4. 选择 `manifest.xml`

#### 方式 C：管理员部署（企业用户）

将 manifest.xml 上传到 Microsoft 365 管理员中心的「集成应用」中，分发给组织内用户。

## 目录结构

```
outlook-addin/
├── manifest.xml          # Office Add-in 清单（需替换域名）
├── src/
│   ├── index.html        # 任务面板主页面
│   ├── app.js            # 业务逻辑 + Kimi API 调用
│   ├── styles.css        # 样式
│   ├── commands.html     # Office 功能入口
│   └── assets/           # 图标文件
│       ├── icon-16.png
│       ├── icon-32.png
│       ├── icon-64.png
│       ├── icon-80.png
│       └── icon-128.png
├── package.json
└── vercel.json           # Vercel 部署配置
```

## 安全说明

- API Key 仅存储在用户本地浏览器的 `localStorage` 中
- 所有 AI 请求直接从浏览器发送到 Kimi 官方 API，不经过中间服务器
- 邮件内容仅在用户主动点击功能按钮时才会发送给 AI

## 常见问题

**Q: 插件显示空白？**
A: 检查 Vercel 域名是否正确配置到 manifest.xml 中，且所有 URL 一致。

**Q: API Key 错误？**
A: 点击插件右上角 ⚙️ 重新输入 Key。确保 Key 来自 platform.moonshot.cn 且有足够余额。

**Q: 无法获取邮件内容？**
A: 确保在 Outlook 阅读或撰写邮件时打开插件，不要在新邮件窗口外使用。

## License

MIT

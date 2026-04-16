# 飞书文档提取工具

通过粘贴飞书文档 URL，自动提取完整文章内容（Markdown 格式，图片本地化）。

## 功能

- 粘贴飞书文档 URL，一键提取
- 支持 Markdown 输出
- 图片自动下载到本地
- 支持 Token 认证访问需要权限的文档
- 跨平台桌面应用

## 安装

```bash
pip install -r requirements.txt
python main.py
```

## 权限说明

首次使用需要配置飞书应用权限：

1. 打开 [飞书开放平台](https://open.feishu.cn/developer)
2. 创建自建应用，获取 App ID 和 App Secret
3. 开通以下权限：
   - `docx:document:readonly`
   - `drive:drive:readonly`
4. 用 App ID/Secret 获取 Access Token

## 许可证

MIT

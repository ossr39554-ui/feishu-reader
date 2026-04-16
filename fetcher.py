"""飞书文档获取器 - 支持公开访问和 Token 认证"""
import re
import requests
import os
import time

BASE_URL = "https://open.feishu.cn/open-apis"


class FeishuFetcher:
    def __init__(self, token: str = None):
        self.token = token
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": "FeishuReader/1.0"})

    def extract_doc_token(self, url: str) -> str:
        """从 URL 中提取 doc_token"""
        patterns = [
            r'/docx/([a-zA-Z0-9]+)',
            r'/doc/([a-zA-Z0-9]+)',
            r'feishu\.cn/doc[xy]/([a-zA-Z0-9]+)',
            r'docs\.feishu\.cn/doc[xy]/([a-zA-Z0-9]+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, url)
            if match:
                return match.group(1)
        raise ValueError(f"无法从 URL 中提取 doc_token: {url}")

    def get_blocks(self, doc_token: str, page_token: str = None) -> dict:
        """获取文档 blocks (新版 docx API)"""
        url = f"{BASE_URL}/docx/v1/documents/{doc_token}/blocks"
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token

        headers = {}
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"

        resp = self.session.get(url, params=params, headers=headers, timeout=30)
        if resp.status_code == 401 or resp.status_code == 403:
            return None  # 需要认证
        resp.raise_for_status()
        return resp.json()

    def get_all_blocks(self, doc_token: str) -> list:
        """递归获取所有 blocks"""
        all_blocks = []
        page_token = None

        while True:
            data = self.get_blocks(doc_token, page_token)
            if not data:
                return None  # 需要认证

            items = data.get("data", {}).get("items", [])
            all_blocks.extend(items)

            # 检查是否有下一页
            page_token = data.get("data", {}).get("page_token")
            has_more = data.get("data", {}).get("has_more", False)

            if not has_more or not page_token:
                break

        return all_blocks

    def download_image(self, image_key: str, save_dir: str) -> str:
        """下载图片并返回本地路径"""
        url = f"{BASE_URL}/im/v1/images/{image_key}"
        headers = {}
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"

        resp = self.session.get(url, headers=headers, timeout=30)
        resp.raise_for_status()

        # 从 Content-Type 获取扩展名
        content_type = resp.headers.get("Content-Type", "image/png")
        ext = content_type.split("/")[-1]
        if ext == "jpeg":
            ext = "jpg"

        filename = f"image_{image_key}.{ext}"
        filepath = os.path.join(save_dir, filename)

        with open(filepath, "wb") as f:
            f.write(resp.content)

        return filename

    def fetch_document(self, url: str, output_dir: str, log_callback=None) -> dict:
        """获取文档并下载所有图片，返回结果"""
        def log(msg):
            if log_callback:
                log_callback(msg)

        doc_token = self.extract_doc_token(url)
        log(f"提取到 doc_token: {doc_token}")

        # 尝试获取 blocks
        log("正在获取文档内容...")
        blocks = self.get_all_blocks(doc_token)

        if blocks is None:
            # 公开获取失败，需要 Token
            if not self.token:
                raise Exception("文档需要认证才能访问，请输入 Access Token")
            raise Exception("Token 无效或已过期")

        log(f"获取到 {len(blocks)} 个 blocks")

        return {
            "doc_token": doc_token,
            "blocks": blocks,
            "image_map": {},
        }

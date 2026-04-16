"""飞书 Block → Markdown 转换器 (新版 docx API)"""
import os


class BlockParser:
    # Block 类型常量（新版 docx API）
    BLOCK_TYPE_PAGE = 1
    BLOCK_TYPE_TEXT = 2
    BLOCK_TYPE_HEADING1 = 3
    BLOCK_TYPE_HEADING2 = 4
    BLOCK_TYPE_HEADING3 = 5
    BLOCK_TYPE_BULLET = 12
    BLOCK_TYPE_ORDERED = 13
    BLOCK_TYPE_CODE = 14
    BLOCK_TYPE_QUOTE = 15
    BLOCK_TYPE_TABLE = 17
    BLOCK_TYPE_IMAGE = 27
    BLOCK_TYPE_SHEET = 30

    def __init__(self, image_map: dict):
        self.image_map = image_map

    def extract_text_from_elements(self, elements: list) -> str:
        """从 elements 数组中提取文本"""
        if not elements:
            return ""
        result = ""
        for elem in elements:
            if isinstance(elem, dict):
                text_run = elem.get("text_run", {})
                content = text_run.get("content", "")
                result += content
            else:
                result += str(elem)
        return result

    def parse_text(self, block: dict) -> str:
        """解析段落/文本"""
        text_obj = block.get("text", {})
        elements = text_obj.get("elements", [])
        return self.extract_text_from_elements(elements)

    def parse_heading(self, block: dict, level: int) -> str:
        """解析标题"""
        key = f"heading{level}"
        heading_obj = block.get(key, {})
        elements = heading_obj.get("elements", [])
        text = self.extract_text_from_elements(elements)
        return "#" * level + " " + text

    def parse_bullet(self, block: dict) -> str:
        """解析无序列表"""
        bullet_obj = block.get("bullet", {})
        elements = bullet_obj.get("elements", [])
        text = self.extract_text_from_elements(elements)
        return "- " + text

    def parse_ordered(self, block: dict) -> str:
        """解析有序列表"""
        ordered_obj = block.get("ordered", {})
        elements = ordered_obj.get("elements", [])
        text = self.extract_text_from_elements(elements)
        return "1. " + text

    def parse_code(self, block: dict) -> str:
        """解析代码块"""
        code_obj = block.get("code", {})
        elements = code_obj.get("elements", [])
        text = self.extract_text_from_elements(elements)
        language = code_obj.get("language", "")
        return f"```{language}\n{text}\n```"

    def parse_quote(self, block: dict) -> str:
        """解析引用"""
        quote_obj = block.get("quote", {})
        elements = quote_obj.get("elements", [])
        text = self.extract_text_from_elements(elements)
        return "> " + text

    def parse_table(self, block: dict) -> str:
        """解析表格"""
        cells = block.get("cells", [])
        if not cells:
            return ""
        result = []
        for row in cells:
            row_text = []
            for cell in row:
                cell_text = self.extract_text_from_elements(cell.get("elements", []))
                row_text.append(cell_text)
            result.append("| " + " | ".join(row_text) + " |")
        if result:
            col_count = len(result[0].split("|")[:-1])
            separator = "|" + "|".join(["---"] * col_count) + "|"
            result.insert(1, separator)
        return "\n".join(result)

    def parse_image(self, block: dict) -> str:
        """解析图片 - 新版 API 使用 image.token"""
        image_info = block.get("image", {})
        token = image_info.get("token", "")

        if token in self.image_map and self.image_map[token]:
            filename = self.image_map[token]
            return f"![image](assets/{filename})"
        elif token:
            return f"![image](assets/image_{token}.png)"
        return ""

    def parse_sheet(self, block: dict) -> str:
        """解析表格（新版）"""
        sheet_obj = block.get("sheet", {})
        token = sheet_obj.get("token", "")
        return f"[表格: {token}]"

    def parse_block(self, block: dict) -> str:
        """解析单个 block"""
        block_type = block.get("block_type", 0)

        parsers = {
            self.BLOCK_TYPE_TEXT: lambda: self.parse_text(block),
            self.BLOCK_TYPE_HEADING1: lambda: self.parse_heading(block, 1),
            self.BLOCK_TYPE_HEADING2: lambda: self.parse_heading(block, 2),
            self.BLOCK_TYPE_HEADING3: lambda: self.parse_heading(block, 3),
            self.BLOCK_TYPE_BULLET: lambda: self.parse_bullet(block),
            self.BLOCK_TYPE_ORDERED: lambda: self.parse_ordered(block),
            self.BLOCK_TYPE_CODE: lambda: self.parse_code(block),
            self.BLOCK_TYPE_QUOTE: lambda: self.parse_quote(block),
            self.BLOCK_TYPE_TABLE: lambda: self.parse_table(block),
            self.BLOCK_TYPE_IMAGE: lambda: self.parse_image(block),
            self.BLOCK_TYPE_SHEET: lambda: self.parse_sheet(block),
        }

        parser = parsers.get(block_type)
        if parser:
            return parser()
        return ""

    def parse(self, blocks: list) -> str:
        """解析所有 blocks 并生成 Markdown"""
        result = []

        for block in blocks:
            content = self.parse_block(block)
            if content:
                result.append(content)

        return "\n\n".join(result)

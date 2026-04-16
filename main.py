"""飞书云文档提取工具 - PyQt6 桌面应用"""
import sys
import os
import datetime
import re
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QMessageBox,
    QProgressBar, QFileDialog
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt

from fetcher import FeishuFetcher
from parser import BlockParser


class FetchWorker(QThread):
    """后台获取线程"""
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str, str)

    def __init__(self, url, token, output_dir):
        super().__init__()
        self.url = url
        self.token = token
        self.output_dir = output_dir

    def run(self):
        try:
            self.log_signal.emit("开始获取文档...")

            fetcher = FeishuFetcher(self.token if self.token else None)

            # 获取文档
            result = fetcher.fetch_document(
                self.url,
                self.output_dir,
                log_callback=self.log_signal.emit
            )

            self.log_signal.emit("正在解析文档内容...")

            # 解析为 Markdown
            parser = BlockParser(result["image_map"])
            markdown = parser.parse(result["blocks"])

            # 生成文件名
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            doc_name = re.sub(r'[^\w\u4e00-\u9fff-]', '_', result["doc_token"])
            md_filename = f"{doc_name}_{timestamp}.md"
            md_filepath = os.path.join(self.output_dir, md_filename)

            # 保存 Markdown
            with open(md_filepath, "w", encoding="utf-8") as f:
                f.write(markdown)

            self.log_signal.emit(f"完成！文件已保存至:\n{md_filepath}")
            self.finished_signal.emit(True, md_filepath, "")

        except Exception as e:
            error_msg = str(e)
            self.log_signal.emit(f"错误: {error_msg}")
            self.finished_signal.emit(False, "", error_msg)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("飞书文档提取工具")
        self.setMinimumWidth(600)
        self.setup_ui()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # URL 输入
        url_layout = QHBoxLayout()
        url_layout.addWidget(QLabel("文档 URL:"))
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("粘贴飞书文档链接，如 https://xxx.feishu.cn/doc/xxx")
        url_layout.addWidget(self.url_input)
        layout.addLayout(url_layout)

        # Token 输入（可选）
        token_layout = QHBoxLayout()
        token_layout.addWidget(QLabel("Access Token:"))
        self.token_input = QLineEdit()
        self.token_input.setPlaceholderText("可选：用于访问需要权限的文档")
        self.token_input.setEchoMode(QLineEdit.EchoMode.Password)
        token_layout.addWidget(self.token_input)
        layout.addLayout(token_layout)

        # 提取按钮
        self.extract_btn = QPushButton("开始提取")
        self.extract_btn.clicked.connect(self.start_fetch)
        layout.addWidget(self.extract_btn)

        # 进度条
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        # 日志区域
        log_label = QLabel("处理日志:")
        layout.addWidget(log_label)
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(200)
        layout.addWidget(self.log_area)

        # 状态栏
        self.statusBar().showMessage("就绪")

    def start_fetch(self):
        url = self.url_input.text().strip()
        if not url:
            QMessageBox.warning(self, "提示", "请输入文档 URL")
            return

        token = self.token_input.text().strip()

        # 选择输出目录
        output_dir = QFileDialog.getExistingDirectory(
            self, "选择保存位置", os.path.expanduser("~/Desktop")
        )
        if not output_dir:
            return

        # 禁用按钮
        self.extract_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)  # 不确定进度
        self.log_area.clear()
        self.statusBar().showMessage("处理中...")

        # 启动后台线程
        self.worker = FetchWorker(url, token, output_dir)
        self.worker.log_signal.connect(self.append_log)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def append_log(self, msg):
        self.log_area.append(msg)
        # 滚动到底部
        self.log_area.verticalScrollBar().setValue(
            self.log_area.verticalScrollBar().maximum()
        )

    def on_finished(self, success, filepath, error):
        self.extract_btn.setEnabled(True)
        self.progress.setVisible(False)
        self.statusBar().showMessage("完成" if success else "失败")

        if success:
            QMessageBox.information(
                self, "完成",
                f"文档已成功提取！\n\n保存位置: {filepath}"
            )
        else:
            QMessageBox.critical(self, "错误", f"提取失败: {error}")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

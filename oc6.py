import sys
import os
import cv2
import numpy as np
import pytesseract
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt, QTimer, QPoint
from PyQt5.QtGui import QPainter, QPen, QColor, QBrush
from PIL import ImageGrab

# ====================== 自动获取当前目录下的Tesseract路径 ======================
# 获取当前脚本所在文件夹
current_dir = os.path.dirname(os.path.abspath(__file__))
# 拼接当前目录下的 tesseract.exe 路径
tesseract_path = os.path.join(current_dir, "Tesseract-OCR", "tesseract.exe")
pytesseract.pytesseract.tesseract_cmd = tesseract_path
OCR_LANG = 'chi_sim+eng'
# =========================================================================


class OverlayWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setGeometry(200, 200, 400, 220)

        self.dragging = False
        self.resizing = False
        self.drag_start_pos = QPoint()
        self.handle_size = 18

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setPen(QPen(QColor(255, 0, 0), 2))
        painter.drawRect(1, 1, self.width() - 2, self.height() - 2)

        painter.setPen(QPen(QColor(0, 100, 255), 1))
        painter.setBrush(QColor(0, 150, 255))
        painter.drawRect(0, self.height() - self.handle_size,
                         self.handle_size, self.handle_size)

        painter.setBrush(QColor(255, 140, 0))
        painter.drawRect(self.width() - self.handle_size,
                         self.height() - self.handle_size,
                         self.handle_size, self.handle_size)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            pos = event.pos()
            if (0 <= pos.x() <= self.handle_size and
                    self.height() - self.handle_size <= pos.y() <= self.height()):
                self.dragging = True
                self.drag_start_pos = event.globalPos() - self.frameGeometry().topLeft()

            elif (self.width() - self.handle_size <= pos.x() <= self.width() and
                  self.height() - self.handle_size <= pos.y() <= self.height()):
                self.resizing = True
                self.drag_start_pos = event.globalPos()

    def mouseMoveEvent(self, event):
        if self.dragging:
            self.move(event.globalPos() - self.drag_start_pos)

        if self.resizing:
            delta = event.globalPos() - self.drag_start_pos
            new_w = self.width() + delta.x()
            new_h = self.height() + delta.y()
            if new_w > 80 and new_h > 60:
                self.setFixedSize(new_w, new_h)
                self.drag_start_pos = event.globalPos()

    def mouseReleaseEvent(self, event):
        self.dragging = False
        self.resizing = False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("xx识别工具")
        self.setGeometry(100, 100, 900, 700)

        self.overlay = OverlayWindow()
        self.overlay.hide()

        self.excel_df = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.do_ocr)

        self.color_map = {
            "红色": "red",
            "蓝色": "blue",
            "绿色": "green",
            "橙色": "orange",
            "紫色": "purple",
            "黑色": "black"
        }

        self.content_key = "题目内容"
        self.answer_key = "答案"
        self.content_color = "red"
        self.answer_color = "blue"

        self.init_ui()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # ====================== 快捷键提示框 ======================
        self.label_hotkey = QLabel()
        self.label_hotkey.setStyleSheet("color: #225599; font-weight: bold; padding: 4px;")
        self.label_hotkey.setText(" 快捷键：F7 = 窗口最小化 + 停止识别")
        layout.addWidget(self.label_hotkey)
        # ==========================================================

        # ====================== 关键字 + 颜色 + 应用按钮 ======================
        key_layout = QHBoxLayout()

        self.edit_content = QLineEdit()
        self.edit_content.setPlaceholderText("列名关键字")
        self.edit_content.setText("题目内容")

        self.edit_answer = QLineEdit()
        self.edit_answer.setPlaceholderText("列名关键字")
        self.edit_answer.setText("答案")

        self.cbo_content_color = QComboBox()
        self.cbo_content_color.addItems(list(self.color_map.keys()))
        self.cbo_content_color.setCurrentText("红色")

        self.cbo_answer_color = QComboBox()
        self.cbo_answer_color.addItems(list(self.color_map.keys()))
        self.cbo_answer_color.setCurrentText("蓝色")

        self.btn_apply = QPushButton("应用设置")
        self.btn_apply.setStyleSheet("background-color:#f0a030; color:white; padding:4px 10px;")

        key_layout.addWidget(QLabel("列名关键字:"))
        key_layout.addWidget(self.edit_content)
        key_layout.addWidget(QLabel("颜色:"))
        key_layout.addWidget(self.cbo_content_color)

        key_layout.addWidget(QLabel("列名关键字:"))
        key_layout.addWidget(self.edit_answer)
        key_layout.addWidget(QLabel("颜色:"))
        key_layout.addWidget(self.cbo_answer_color)
        key_layout.addWidget(self.btn_apply)

        layout.addLayout(key_layout)
        # ====================================================================

        btn_layout = QHBoxLayout()
        self.btn_start = QPushButton("开始识别")
        self.btn_stop = QPushButton("停止识别")
        self.btn_load = QPushButton("导入Excel")
        btn_layout.addWidget(self.btn_start)
        btn_layout.addWidget(self.btn_stop)
        btn_layout.addWidget(self.btn_load)
        layout.addLayout(btn_layout)

        layout.addWidget(QLabel("OCR识别文字："))
        self.txt_ocr = QTextEdit()
        self.txt_ocr.setReadOnly(True)
        layout.addWidget(self.txt_ocr)

        layout.addWidget(QLabel("Excel最佳匹配（按匹配度排序）："))
        self.txt_match = QTextEdit()
        self.txt_match.setReadOnly(True)
        layout.addWidget(self.txt_match)

        # 按钮绑定
        self.btn_apply.clicked.connect(self.apply_settings)
        self.btn_start.clicked.connect(self.start_ocr)
        self.btn_stop.clicked.connect(self.stop_ocr)
        self.btn_load.clicked.connect(self.load_excel)

    # ====================== F7 快捷键功能 ======================
    def keyPressEvent(self, event):
        if event.key() == Qt.Key_F7:
            self.stop_ocr()        # 停止识别
            self.showMinimized()   # 窗口最小化
            self.txt_ocr.append("\n[F7] 已停止识别并最小化窗口")
        super().keyPressEvent(event)
    # ===========================================================

    def apply_settings(self):
        self.content_key = self.edit_content.text().strip()
        self.answer_key = self.edit_answer.text().strip()
        self.content_color = self.color_map[self.cbo_content_color.currentText()]
        self.answer_color = self.color_map[self.cbo_answer_color.currentText()]
        self.txt_match.append("<br> ===== 已应用高亮设置 =====")

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(filter="Excel Files (*.xlsx *.xls)")
        if not path:
            return
        try:
            self.excel_df = pd.read_excel(path).fillna("").astype(str)
            tip = f"已导入 Excel：共 {len(self.excel_df)} 行"
            self.txt_match.setHtml(tip)
        except Exception as e:
            self.txt_match.setHtml(f"导入失败：{str(e)}")

    def start_ocr(self):
        self.overlay.show()
        self.timer.start(300)

    def stop_ocr(self):
        self.timer.stop()
        self.overlay.hide()

    def do_ocr(self):
        try:
            x, y = self.overlay.x(), self.overlay.y()
            w, h = self.overlay.width(), self.overlay.height()

            img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
            gray = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)
            text = pytesseract.image_to_string(gray, lang=OCR_LANG).strip()
            self.txt_ocr.setText(text)

            if not text or self.excel_df is None:
                return

            self.match_and_sort(text)

        except Exception as e:
            self.txt_ocr.setText(f"识别异常：{str(e)}")

    def match_and_sort(self, ocr_text):
        words = [w.strip() for w in ocr_text.split() if len(w.strip()) >= 2]
        if not words:
            self.txt_match.setHtml("未提取到有效关键词")
            return

        matches = []
        for idx, row in self.excel_df.iterrows():
            line = " ".join(row.values).lower()
            score = sum(1 for w in words if w.lower() in line)
            if score > 0:
                matches.append((-score, idx + 2, row))

        matches.sort()
        if matches:
            html_list = []
            for score, row_num, row_data in matches[:15]:
                row_parts = []
                for col_name, value in row_data.items():
                    if self.content_key and self.content_key in col_name:
                        row_parts.append(f'<span style="color:{self.content_color}; font-weight:bold;">{value}</span>')
                    elif self.answer_key and self.answer_key in col_name:
                        row_parts.append(f'<span style="color:{self.answer_color}; font-weight:bold;">{value}</span>')
                    else:
                        row_parts.append(value)

                row_str = " | ".join(row_parts)
                html_list.append(f"第{row_num}行 [{-score}分]：{row_str}")

            self.txt_match.setHtml("<br><br>".join(html_list))
        else:
            self.txt_match.setHtml("未匹配到相关内容")

    def closeEvent(self, event):
        self.overlay.close()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())

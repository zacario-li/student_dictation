import sys
import threading
import time
import re
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextEdit, QLineEdit, QComboBox, QFileDialog
)
from PySide6.QtGui import QIcon, QTextCursor, QTextCharFormat, QColor, QFont, QPainter, QPen, QBrush
from PySide6.QtCore import Qt, QTimer, Signal
import pygame
import os
import uuid
import requests
from io import BytesIO
from PIL import Image, ImageDraw
import hashlib
import base64
import json
import websocket
import ssl
import pandas as pd
import qdarkstyle
import math

class CircularCountdownWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.total_seconds = 1
        self.remaining_seconds = 1
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.on_tick)
        self.setMinimumSize(60, 60)
        self.setMaximumSize(60, 60)
        self.running = False
        self.finished_callback = None

    def start(self, seconds, finished_callback=None):
        self.total_seconds = seconds
        self.remaining_seconds = seconds
        self.running = True
        self.finished_callback = finished_callback
        self.timer.start(50)  # 20fps for smoothness
        self.update()

    def stop(self):
        self.running = False
        self.timer.stop()
        self.update()

    def on_tick(self):
        if not self.running:
            return
        self.remaining_seconds -= 0.05
        if self.remaining_seconds <= 0:
            self.remaining_seconds = 0
            self.running = False
            self.timer.stop()
            self.update()
            if self.finished_callback:
                self.finished_callback()
            return
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        rect = self.rect().adjusted(5, 5, -5, -5)
        # 背景圆
        painter.setPen(QPen(QColor('#cccccc'), 6))
        painter.drawEllipse(rect)
        # 进度扇形
        if self.total_seconds > 0:
            angle = 360 * (self.remaining_seconds / self.total_seconds)
            painter.setPen(QPen(QColor('#4f8cff'), 6))
            painter.drawArc(rect, 90 * 16, -angle * 16)
        # 中心文字
        painter.setPen(QPen(QColor('#222222'), 1))
        painter.setFont(QFont("微软雅黑", 10, QFont.Bold))
        text = f"{int(self.remaining_seconds + 0.05)}s" if self.running and self.remaining_seconds > 0 else ""
        painter.drawText(rect, Qt.AlignCenter, text)

class WordAnnouncer(QWidget):
    countdown_start = Signal(float)
    countdown_hide = Signal()
    def __init__(self):
        super().__init__()
        self.setWindowTitle("小学生词语默写播报器")
        self.setWindowIcon(QIcon("app_icon.ico"))
        self.resize(400, 500)
        self.is_playing = False
        self.is_paused = False
        self.play_thread = None
        self.tts_thread = None
        self.current_word_index = -1
        self.total_words = 0
        self.tts_engine = 'edge'  # 可选 'edge' 或 'pyttsx3'
        pygame.mixer.init()  # 初始化pygame音频
        pygame.mixer.music.set_volume(1.0)  # 设置最大音量
        # Excel相关
        self.lesson_words = {}  # 课名 -> 词语列表
        self.excel_loaded = False
        self.load_excel_words()
        self.init_ui()

        # 信号连接
        self.countdown_start.connect(self._start_countdown_mainthread)

    def load_excel_words(self):
        excel_path = 'words.xlsx'  # 你可以修改为实际Excel文件名
        if not os.path.exists(excel_path):
            return
        try:
            df = pd.read_excel(excel_path, header=None, engine='openpyxl')
            # 第一行是课名，后面每列是该课的词语
            for col in df:
                lesson = str(df[col][0]).strip()
                words = [str(x).strip() for x in df[col][1:] if pd.notna(x) and str(x).strip()]
                if lesson and words:
                    self.lesson_words[lesson] = words
            self.excel_loaded = True
        except Exception as e:
            print(f"读取Excel失败: {e}")
            self.lesson_words = {}
            self.excel_loaded = False

    def init_ui(self):
        # 倒计时控件先初始化
        self.countdown = CircularCountdownWidget()
        self.countdown.show()
        layout = QVBoxLayout()

        # 选择TTS引擎
        tts_layout = QHBoxLayout()
        tts_label = QLabel("选择语音播报引擎:")
        self.tts_combo = QComboBox()
        self.tts_combo.addItem("Edge-TTS(免费,联网)", 'edge')
        self.tts_combo.addItem("pyttsx3(免费,离线)", 'pyttsx3')
        self.tts_combo.setCurrentIndex(0)
        self.tts_combo.currentIndexChanged.connect(self.on_tts_selected)
        tts_layout.addWidget(tts_label)
        tts_layout.addWidget(self.tts_combo)
        tts_layout.addStretch(1)
        layout.addLayout(tts_layout)

        # 选择Excel文件按钮
        file_layout = QHBoxLayout()
        self.file_label = QLabel("当前Excel: words.xlsx")
        self.file_button = QPushButton("选择Excel文件")
        self.file_button.clicked.connect(self.on_choose_excel)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.file_button)
        layout.addLayout(file_layout)

        # 课选择
        self.lesson_combo = QComboBox()
        self.lesson_combo.addItem("选择课文（可选）")
        if self.excel_loaded:
            for lesson in self.lesson_words:
                self.lesson_combo.addItem(lesson)
        self.lesson_combo.currentIndexChanged.connect(self.on_lesson_selected)
        layout.addWidget(self.lesson_combo)

        # 间隔设置和倒计时进度条同一行
        interval_row = QHBoxLayout()
        interval_label = QLabel("播报间隔(秒):")
        self.interval_input = QLineEdit("3")
        interval_row.addWidget(interval_label)
        interval_row.addWidget(self.interval_input)
        interval_row.addWidget(self.countdown)
        interval_row.addStretch(1)
        layout.addLayout(interval_row)

        # 词语输入
        words_label_layout = QHBoxLayout()
        self.words_label = QLabel("请输入要播报的词组(每行一个，或用空格分隔):")
        self.progress_label = QLabel("当前词语：0/0")
        words_label_layout.addWidget(self.words_label)
        words_label_layout.addWidget(self.progress_label)
        words_label_layout.addStretch(1)
        layout.addLayout(words_label_layout)

        self.text_area = QTextEdit()
        layout.addWidget(self.text_area)

        # 按钮
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("开始播报")
        self.pause_button = QPushButton("暂停")
        self.stop_button = QPushButton("停止播报")
        self.clear_button = QPushButton("清除内容")
        self.pause_button.setEnabled(False)
        self.stop_button.setEnabled(False)

        self.start_button.clicked.connect(self.on_start)
        self.pause_button.clicked.connect(self.on_pause)
        self.stop_button.clicked.connect(self.on_stop)
        self.clear_button.clicked.connect(self.on_clear)

        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.pause_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addWidget(self.clear_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

        self.setStyleSheet("""
            QWidget {
                background: #f5f6fa;
            }
            QPushButton {
                background-color: #4f8cff;
                color: white;
                border-radius: 8px;
                padding: 6px 16px;
                font-size: 14px;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
            QLineEdit, QTextEdit {
                border: 1px solid #cccccc;
                border-radius: 6px;
                padding: 4px;
                background: #ffffff;
                font-size: 14px;
                color: #222222;
            }
            QComboBox {
                border-radius: 6px;
                padding: 4px;
                font-size: 14px;
                color: #222222;
                background: #ffffff;
            }
            QComboBox QAbstractItemView {
                background: #ffffff;
                color: #222222;
                selection-background-color: #4f8cff;
                selection-color: #ffffff;
                font-size: 14px;
            }
            QLabel {
                font-size: 14px;
                color: #222222;
            }
            QComboBox:disabled {
                background: #eeeeee;
                color: #aaaaaa;
                border: 1px solid #cccccc;
            }
        """)

    def on_lesson_selected(self, idx):
        if idx <= 0:
            return
        lesson = self.lesson_combo.currentText()
        words = self.lesson_words.get(lesson, [])
        if words:
            self.text_area.setPlainText('\n'.join(words))
            self.total_words = len(words)
            self.update_progress_label()

    def on_start(self):
        if not self.is_playing:
            # 格式化输入
            raw_text = self.text_area.toPlainText().strip()
            words = re.split(r'[ \t\u3000\r\n]+', raw_text)
            words = [word for word in words if word.strip()]
            formatted_text = '\n'.join(words)
            self.text_area.setPlainText(formatted_text)
            self.total_words = len(words)
            self.update_progress_label()

            self.is_playing = True
            self.is_paused = False
            self.start_button.setEnabled(False)
            self.pause_button.setEnabled(True)
            self.stop_button.setEnabled(True)
            self.clear_button.setEnabled(False)
            self.lesson_combo.setEnabled(False)
            self.play_thread = threading.Thread(target=self.play_words)
            self.play_thread.start()

    def on_pause(self):
        if self.is_playing:
            self.is_paused = not self.is_paused
            self.pause_button.setText('继续' if self.is_paused else '暂停')

    def on_stop(self):
        self.is_playing = False
        self.is_paused = False
        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        self.clear_button.setEnabled(True)
        self.current_word_index = -1
        self.highlight_current_word(-1)
        self.lesson_combo.setEnabled(True)
        self.update_progress_label()

    def on_clear(self):
        if not self.is_playing:
            self.text_area.clear()
            self.current_word_index = -1
            self.highlight_current_word(-1)
            self.total_words = 0
            self.update_progress_label()

    def highlight_current_word(self, index):
        cursor = self.text_area.textCursor()
        fmt = QTextCharFormat()
        # 清除所有高亮
        cursor.select(QTextCursor.Document)
        fmt.setBackground(Qt.white)
        cursor.setCharFormat(fmt)
        # 高亮当前词
        if index >= 0:
            lines = self.text_area.toPlainText().split('\n')
            start = sum(len(l) + 1 for l in lines[:index])
            length = len(lines[index])
            cursor.setPosition(start)
            cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, length)
            fmt.setBackground(QColor(255, 255, 0))  # 黄色高亮
            cursor.setCharFormat(fmt)
            self.text_area.setTextCursor(cursor)
        self.update_progress_label()

    def say_text(self, text):
        """根据 tts_engine 选择 TTS 服务"""
        if self.tts_engine == 'edge':
            self._say_text_edge(text)
        elif self.tts_engine == 'pyttsx3':
            self._say_text_pyttsx3(text)
        else:
            print("未知TTS引擎")

    def _say_text_edge(self, text):
        """使用 edge-tts 合成并播放语音"""
        def tts_and_play():
            try:
                import asyncio
                import edge_tts
                mp3_path = f"_edge_tts_{uuid.uuid4().hex}.mp3"
                rate = "-30%"  # 语速调慢，越负越慢，可根据需要调整
                async def run():
                    communicate = edge_tts.Communicate(text, "zh-CN-XiaoxiaoNeural", rate=rate)
                    await communicate.save(mp3_path)
                asyncio.run(run())
                pygame.mixer.music.load(mp3_path)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    pygame.time.Clock().tick(10)
                pygame.mixer.music.unload()
                try:
                    os.remove(mp3_path)
                except:
                    pass
            except Exception as e:
                print(f"edge-tts异常: {str(e)}")
        self.tts_thread = threading.Thread(target=tts_and_play)
        self.tts_thread.start()

    def _say_text_pyttsx3(self, text):
        """使用 pyttsx3 (离线) 合成并播放语音"""
        def tts_and_play():
            try:
                import pyttsx3
                engine = pyttsx3.init()
                # 自动选择中文语音（Windows 下通常有 Microsoft Huihui/Microsoft Xiaoxiao）
                voices = engine.getProperty('voices')
                for v in voices:
                    # 兼容不同 pyttsx3 版本和平台
                    lang = ''
                    if hasattr(v, 'languages') and v.languages:
                        # 有些 pyttsx3 版本是 bytes，有些是 str
                        try:
                            lang = v.languages[0]
                            if isinstance(lang, bytes):
                                lang = lang.decode('utf-8', errors='ignore')
                        except Exception:
                            lang = ''
                    if ('zh' in lang.lower()) or ('chinese' in v.name.lower()):
                        engine.setProperty('voice', v.id)
                        break
                # 设置语速，数值越小越慢，100~150较为自然
                engine.setProperty('rate', 130)  # 语速可调，推荐130左右
                engine.setProperty('volume', 1.0)
                engine.say(text)
                engine.runAndWait()
                engine.stop()
            except ImportError:
                print("未安装 pyttsx3，请先运行: pip install pyttsx3")
            except Exception as e:
                print(f"pyttsx3异常: {str(e)}")
        self.tts_thread = threading.Thread(target=tts_and_play)
        self.tts_thread.start()

    def _start_countdown_mainthread(self, interval):
        finished = getattr(self, '_countdown_finished_event', None)
        def on_countdown_finished():
            if finished:
                finished.set()
        self.countdown.show()
        self.countdown.start(interval, finished_callback=on_countdown_finished)

    def play_words(self):
        try:
            interval = float(self.interval_input.text())
        except ValueError:
            interval = 10

        words = self.text_area.toPlainText().strip().split('\n')
        words = [word for word in words if word.strip()]

        if not words:
            self.on_stop()
            return

        # 播放开始提示
        self.say_text("准备开始")
        time.sleep(3)

        for i, word in enumerate(words):
            if not self.is_playing:
                break
            self.current_word_index = i
            self.highlight_current_word(i)
            for _ in range(2):
                if not self.is_playing:
                    break
                while self.is_paused:
                    time.sleep(0.1)
                    if not self.is_playing:
                        break
                if not self.is_playing:
                    break
                self.say_text(word)
                time.sleep(3)
            if self.is_playing:
                # 启动倒计时控件（主线程）
                finished = threading.Event()
                self._countdown_finished_event = finished
                self.countdown_start.emit(interval)
                while not finished.is_set() and self.is_playing and not self.is_paused:
                    time.sleep(0.05)
        if self.is_playing:
            self.current_word_index = len(words) - 1
            self.highlight_current_word(self.current_word_index)
            self.say_text("默写结束")
            while pygame.mixer.music.get_busy():
                time.sleep(0.1)
        self.on_stop()

    def on_choose_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.file_label.setText(f"当前Excel: {os.path.basename(file_path)}")
            self.lesson_words = {}
            self.excel_loaded = False
            try:
                df = pd.read_excel(file_path, header=None, engine='openpyxl')
                for col in df:
                    lesson = str(df[col][0]).strip()
                    words = [str(x).strip() for x in df[col][1:] if pd.notna(x) and str(x).strip()]
                    if lesson and words:
                        self.lesson_words[lesson] = words
                self.excel_loaded = True
            except Exception as e:
                print(f"读取Excel失败: {e}")
                self.lesson_words = {}
                self.excel_loaded = False
            # 刷新下拉框
            self.lesson_combo.clear()
            self.lesson_combo.addItem("选择课文（可选）")
            if self.excel_loaded:
                for lesson in self.lesson_words:
                    self.lesson_combo.addItem(lesson)

    def update_progress_label(self):
        current = self.current_word_index + 1 if self.current_word_index >= 0 else 0
        self.progress_label.setText(f"当前词语：{current}/{self.total_words}")

    def on_tts_selected(self, idx):
        self.tts_engine = self.tts_combo.currentData()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    font = QFont("微软雅黑", 12)
    app.setFont(font)
    app.setStyleSheet(qdarkstyle.load_stylesheet())
    window = WordAnnouncer()
    window.show()
    sys.exit(app.exec())
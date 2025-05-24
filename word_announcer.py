import wx
import time
import threading
import asyncio
import edge_tts
import pygame
import os
import uuid
import requests
from io import BytesIO
from PIL import Image, ImageDraw
import re

class WordAnnouncerFrame(wx.Frame):
    def __init__(self):
        super().__init__(parent=None, title='小学生词语默写播报器', size=(400, 800))
        
        # 设置应用图标
        try:
            # 创建图标文件
            icon_path = "app_icon.ico"
            if not os.path.exists(icon_path):
                # 使用PIL创建图标
                img = Image.new('RGBA', (32, 32), (0, 120, 212, 255))  # 蓝色背景
                draw = ImageDraw.Draw(img)
                # 绘制书本
                draw.rectangle([8, 4, 24, 28], fill=(255, 255, 255, 255))  # 白色书本
                draw.line([8, 12, 24, 12], fill=(0, 120, 212, 255), width=2)  # 书页线
                draw.line([8, 18, 24, 18], fill=(0, 120, 212, 255), width=2)  # 书页线
                # 保存为ICO文件
                img.save(icon_path, format='ICO')
            
            # 加载图标
            icon = wx.Icon(icon_path)
            self.SetIcon(icon)
        except Exception as e:
            print(f"无法加载图标: {str(e)}")
        
        # 控制变量
        self.is_playing = False
        self.is_paused = False
        self.play_thread = None
        self.tts_thread = None
        self.current_word_index = -1
        pygame.mixer.init()  # 初始化pygame音频
        pygame.mixer.music.set_volume(1.0)  # 设置最大音量
        
        self.init_ui()
        self.Centre()
        
    def init_ui(self):
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        
        # 间隔时间设置
        interval_box = wx.BoxSizer(wx.HORIZONTAL)
        interval_label = wx.StaticText(panel, label='播报间隔(秒):')
        self.interval_input = wx.TextCtrl(panel, value='3')
        interval_box.Add(interval_label, 0, wx.ALL | wx.CENTER, 5)
        interval_box.Add(self.interval_input, 0, wx.ALL, 5)
        vbox.Add(interval_box, 0, wx.ALL | wx.EXPAND, 5)
        
        # 词组输入区域
        words_label = wx.StaticText(panel, label='请输入要播报的词组(每行一个，或用空格分隔):')
        vbox.Add(words_label, 0, wx.ALL, 5)
        
        # 文本区域
        self.text_area = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.TE_RICH2 | wx.HSCROLL)
        self.text_area.Bind(wx.EVT_TEXT, self.on_text_change)  # 添加文本变化事件处理
        vbox.Add(self.text_area, 1, wx.ALL | wx.EXPAND, 5)
        
        # 控制按钮
        button_box = wx.BoxSizer(wx.HORIZONTAL)
        self.start_button = wx.Button(panel, label='开始播报')
        self.pause_button = wx.Button(panel, label='暂停')
        self.stop_button = wx.Button(panel, label='停止播报')
        self.clear_button = wx.Button(panel, label='清除内容')
        
        self.pause_button.Disable()
        self.stop_button.Disable()
        
        self.start_button.Bind(wx.EVT_BUTTON, self.on_start)
        self.pause_button.Bind(wx.EVT_BUTTON, self.on_pause)
        self.stop_button.Bind(wx.EVT_BUTTON, self.on_stop)
        self.clear_button.Bind(wx.EVT_BUTTON, self.on_clear)
        
        button_box.Add(self.start_button, 0, wx.ALL, 5)
        button_box.Add(self.pause_button, 0, wx.ALL, 5)
        button_box.Add(self.stop_button, 0, wx.ALL, 5)
        button_box.Add(self.clear_button, 0, wx.ALL, 5)
        vbox.Add(button_box, 0, wx.ALL | wx.CENTER, 5)
        
        panel.SetSizer(vbox)
        
    def on_start(self, event):
        if not self.is_playing:
            # 新增：自动格式化文本区域内容为每行一个词语
            raw_text = self.text_area.GetValue().strip()
            words = re.split(r'[ \t\u3000\r\n]+', raw_text)
            words = [word for word in words if word.strip()]
            formatted_text = '\n'.join(words)
            self.text_area.SetValue(formatted_text)

            self.is_playing = True
            self.is_paused = False
            self.start_button.Disable()
            self.pause_button.Enable()
            self.stop_button.Enable()
            self.clear_button.Disable()
            self.play_thread = threading.Thread(target=self.play_words)
            self.play_thread.start()
    
    def on_pause(self, event):
        if self.is_playing:
            self.is_paused = not self.is_paused
            self.pause_button.SetLabel('继续' if self.is_paused else '暂停')
            self.update_arrow()
    
    def on_stop(self, event):
        self.is_playing = False
        self.is_paused = False
        self.start_button.Enable()
        self.pause_button.Disable()
        self.stop_button.Disable()
        self.clear_button.Enable()
        self.update_arrow()
    
    def on_clear(self, event):
        if not self.is_playing:
            self.text_area.SetValue('')
            self.current_word_index = -1  # 清除内容时才重置索引
            self.update_arrow()
    
    def on_text_change(self, event):
        """处理文本变化事件"""
        text = self.text_area.GetValue()
        if not text:
            return
        
        # 支持空格、制表符、全角空格等分隔符
        if ((' ' in text or '\t' in text or '\u3000' in text) and '\n' not in text):
            # 用正则分割所有空白字符（空格、制表符、全角空格等）
            words = re.split(r'[ \t\u3000]+', text.strip())
            formatted_text = '\n'.join([w for w in words if w])
            self.text_area.SetValue(formatted_text)
            self.text_area.SetInsertionPointEnd()
    
    def update_arrow(self):
        """更新箭头位置"""
        if self.current_word_index < 0:
            return
        
        # 获取所有文本
        all_text = self.text_area.GetValue()
        lines = all_text.split('\n')
        
        # 计算当前词组的位置
        start_pos = 0
        for i in range(self.current_word_index):
            start_pos += len(lines[i]) + 1  # +1 for newline
        
        # 先移除已有箭头
        word = lines[self.current_word_index]
        if word.endswith(" ←"):
            word = word[:-2]
        lines[self.current_word_index] = word
        
        if self.is_playing and not self.is_paused:
            # 红色箭头
            lines[self.current_word_index] = word + " ←"
            new_text = '\n'.join(lines)
            self.text_area.SetValue(new_text)
            arrow_pos = start_pos + len(word)
            self.text_area.SetStyle(arrow_pos, arrow_pos + 2, wx.TextAttr(wx.RED))
        else:
            # 灰色箭头
            lines[self.current_word_index] = word + " ←"
            new_text = '\n'.join(lines)
            self.text_area.SetValue(new_text)
            arrow_pos = start_pos + len(word)
            self.text_area.SetStyle(arrow_pos, arrow_pos + 2, wx.TextAttr(wx.Colour(128,128,128)))  # 灰色
        
        self.text_area.ShowPosition(start_pos)
    
    def say_text(self, text):
        """使用Edge-TTS合成并播放语音，声音大且自然"""
        def tts_and_play():
            try:
                mp3_path = f"_edge_tts_temp_{uuid.uuid4().hex}.mp3"
                async def generate_speech():
                    # 使用高质量中文女声
                    communicate = edge_tts.Communicate(
                        text,
                        "zh-CN-XiaoxiaoNeural",
                        rate="+0%",  # 正常语速
                        volume="+50%"  # 增加50%音量
                    )
                    await communicate.save(mp3_path)
                
                # 在单独的线程中运行异步函数
                asyncio.run(generate_speech())
                
                # 播放音频
                pygame.mixer.music.load(mp3_path)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    pygame.time.Clock().tick(10)
                pygame.mixer.music.unload()
                
                try:
                    os.remove(mp3_path)
                except:
                    pass  # 忽略文件删除错误
            except Exception as e:
                print(f"TTS错误: {str(e)}")
                # 如果TTS失败，使用备用方案
                try:
                    import win32com.client
                    speaker = win32com.client.Dispatch("SAPI.SpVoice")
                    speaker.Volume = 100
                    speaker.Speak(text)
                except:
                    pass
        
        # 在单独的线程中运行TTS和播放
        self.tts_thread = threading.Thread(target=tts_and_play)
        self.tts_thread.start()
    
    def play_words(self):
        try:
            interval = float(self.interval_input.GetValue())
        except ValueError:
            interval = 10

        # 新增：对输入内容做分割处理，支持空格、Tab、全角空格、换行
        raw_text = self.text_area.GetValue().strip()
        # 先把所有分隔符都换成换行
        words = re.split(r'[ \t\u3000\r\n]+', raw_text)
        words = [word for word in words if word.strip()]

        if not words:
            wx.CallAfter(self.on_stop, None)
            return
            
        # 播放开始提示
        wx.CallAfter(self.say_text, "准备开始")
        time.sleep(3)
        
        # 播报词组
        for i, word in enumerate(words):
            if not self.is_playing:
                break
                
            # 更新当前词组索引和箭头
            self.current_word_index = i
            wx.CallAfter(self.update_arrow)
            
            # 播报两次
            for _ in range(2):
                if not self.is_playing:
                    break
                while self.is_paused:
                    time.sleep(0.1)
                    if not self.is_playing:
                        break
                if not self.is_playing:
                    break
                wx.CallAfter(self.say_text, word)
                time.sleep(3)
                
            if self.is_playing:
                time.sleep(interval)
        
        # 播放结束提示
        if self.is_playing:
            wx.CallAfter(setattr, self, "current_word_index", len(words) - 1)  # 保证箭头在最后一个词
            wx.CallAfter(self.update_arrow)
            wx.CallAfter(self.say_text, "默写结束")
            # 等待"默写结束"播报完毕
            while pygame.mixer.music.get_busy():
                time.sleep(0.1)

        wx.CallAfter(self.on_stop, None)

def main():
    app = wx.App()
    frame = WordAnnouncerFrame()
    frame.Show()
    app.MainLoop()

if __name__ == '__main__':
    main() 
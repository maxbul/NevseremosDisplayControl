import sys
import subprocess
import re
import cv2
import numpy as np
import ctypes
import time
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QVBoxLayout, QHBoxLayout,
                             QWidget, QSystemTrayIcon, QMenu, QAction, QPushButton, 
                             QComboBox, QCheckBox, QSpinBox, QDoubleSpinBox, QGroupBox)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QImage, QPixmap, QIcon, QColor
from pynput import keyboard, mouse

class MonitorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Smart Monitor Controller v2.10")
        self.resize(460, 880)

        # Стани
        self.is_active = True # Головний вимикач
        self.video_active = False
        self.motion_detected = False
        self.hid_active = False
        self.presence_detected = False
        self.monitor_is_off = False
        
        self.seconds_without_motion = 0
        self.off_start_time = 0      
        self.continuous_motion_sec = 0 

        self.cap = None
        self.avg_frame = None
        self.reference_frame = None

        self.VK_SCROLL = 0x91
        self.KEYEVENTF_KEYUP = 0x0002

        self.init_ui()
        self.create_tray_icon()

        # Слухачі HID
        self.kb_listener = keyboard.Listener(on_press=self.on_hid_event)
        self.mouse_listener = mouse.Listener(on_move=self.on_hid_event, on_click=self.on_hid_event, on_scroll=self.on_hid_event)
        self.kb_listener.start()
        self.mouse_listener.start()

        self.logic_timer = QTimer()
        self.logic_timer.timeout.connect(self.process_logic)
        self.logic_timer.start(1000)

        self.cam_timer = QTimer()
        self.cam_timer.timeout.connect(self.update_camera)

        if self.cam_select.count() >= 1:
            self.toggle_camera()

    def init_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        group_style = "QGroupBox { border: 2px solid black; border-radius: 5px; margin-top: 10px; font-weight: bold; } " \
                      "QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px 0 3px; }"

        # --- ГОЛОВНА КНОПКА КЕРУВАННЯ ---
        self.btn_main_toggle = QPushButton("ПРАЦЮЄ")
        self.btn_main_toggle.setFixedHeight(50)
        self.btn_main_toggle.setStyleSheet("font-weight: bold; font-size: 14pt; background-color: #d4edda; color: #155724; border: 2px solid #155724;")
        self.btn_main_toggle.clicked.connect(self.toggle_main_system)
        main_layout.addWidget(self.btn_main_toggle)

        # --- СТАТУСИ ТА КАМЕРА ---
        top_box = QWidget()
        top_layout = QVBoxLayout(top_box)
        self.status_label = QLabel("Статус: Очікування...")
        self.timer_label = QLabel("Без активності: 0 сек")
        self.video_msg_label = QLabel("")
        self.video_msg_label.setStyleSheet("color: #ff0000; font-weight: bold;")
        top_layout.addWidget(self.status_label)
        top_layout.addWidget(self.timer_label)
        top_layout.addWidget(self.video_msg_label)
        
        cam_sel_layout = QHBoxLayout()
        self.cam_select = QComboBox()
        self.refresh_cameras()
        self.btn_cam = QPushButton("Пуск")
        self.btn_cam.clicked.connect(self.toggle_camera)
        cam_sel_layout.addWidget(self.cam_select)
        cam_sel_layout.addWidget(self.btn_cam)
        top_layout.addLayout(cam_sel_layout)
        
        self.video_frame = QLabel("Камера вимкнена")
        self.video_frame.setFixedSize(320, 240)
        self.video_frame.setStyleSheet("background: black; border: 1px solid gray;")
        self.video_frame.setAlignment(Qt.AlignCenter)
        top_layout.addWidget(self.video_frame, alignment=Qt.AlignCenter)
        main_layout.addWidget(top_box)

        # --- ГРУПА: ДЕТЕКЦІЯ ПРИСУТНОСТІ ---
        presence_group = QGroupBox("Детекція присутності")
        presence_group.setStyleSheet(group_style)
        presence_layout = QVBoxLayout()
        h_presence = QHBoxLayout()
        self.ref_frame_label = QLabel("Немає фото")
        self.ref_frame_label.setFixedSize(120, 90)
        self.ref_frame_label.setStyleSheet("background: #eee; border: 1px solid gray;")
        self.ref_frame_label.setAlignment(Qt.AlignCenter)
        self.btn_ref = QPushButton("Запам'ятати\nпорожню кімнату")
        self.btn_ref.setFixedHeight(90)
        self.btn_ref.clicked.connect(self.capture_reference)
        h_presence.addWidget(self.ref_frame_label)
        h_presence.addWidget(self.btn_ref)
        presence_layout.addLayout(h_presence)
        self.check_ref_mode = QCheckBox("Вкл. (блокувати, якщо я в кадрі)")
        presence_layout.addWidget(self.check_ref_mode)
        presence_group.setLayout(presence_layout)
        main_layout.addWidget(presence_group)

        # --- ГРУПА: КЕРУВАННЯ ЖИВЛЕННЯМ ---
        power_group = QGroupBox("Керування живленням")
        power_group.setStyleSheet(group_style)
        power_layout = QVBoxLayout()
        h_off = QHBoxLayout()
        self.check_auto_off = QCheckBox("Гасити при відсутності більше")
        self.check_auto_off.setChecked(True)
        self.spin_timeout = QSpinBox()
        self.spin_timeout.setRange(5, 3600)
        self.spin_timeout.setValue(10)
        h_off.addWidget(self.check_auto_off)
        h_off.addWidget(self.spin_timeout)
        h_off.addWidget(QLabel("сек."))
        power_layout.addLayout(h_off)
        self.check_auto_on = QCheckBox("Вмикати при виявленні руху/вводу")
        self.check_auto_on.setChecked(True)
        power_layout.addWidget(self.check_auto_on)
        power_group.setLayout(power_layout)
        main_layout.addWidget(power_group)

        # --- ГРУПА: ЗАТРИМКА ПРОБУДЖЕННЯ ---
        wake_group = QGroupBox("Затримка пробудження")
        wake_group.setStyleSheet(group_style)
        wake_layout = QVBoxLayout()
        self.check_smart_wake = QCheckBox("Вкл.")
        wake_layout.addWidget(self.check_smart_wake)
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Якщо вимкнено більше ніж (хв):"))
        self.spin_smart_off_min = QSpinBox()
        self.spin_smart_off_min.setValue(3)
        row1.addWidget(self.spin_smart_off_min)
        wake_layout.addLayout(row1)
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Потрібно впевненого руху (сек):"))
        self.spin_smart_motion_sec = QSpinBox()
        self.spin_smart_motion_sec.setValue(10)
        row2.addWidget(self.spin_smart_motion_sec)
        wake_layout.addLayout(row2)
        wake_group.setLayout(wake_layout)
        main_layout.addWidget(wake_group)

        # --- ГРУПА: ОПТИМІЗАЦІЯ ---
        opt_group = QGroupBox("Оптимізація навантаження на ПК")
        opt_group.setStyleSheet(group_style)
        opt_layout = QVBoxLayout()
        self.check_optimize = QCheckBox("Активувати")
        self.check_optimize.setChecked(True)
        opt_layout.addWidget(self.check_optimize)
        f_layout = QHBoxLayout()
        f_layout.addWidget(QLabel("Кадрів в сек (звичайний):"))
        self.spin_fps = QSpinBox()
        self.spin_fps.setRange(1, 30)
        self.spin_fps.setValue(3)
        f_layout.addWidget(self.spin_fps)
        opt_layout.addLayout(f_layout)
        f_layout2 = QHBoxLayout()
        f_layout2.addWidget(QLabel("Інтервал при присутності (сек):"))
        self.spin_eco_interval = QDoubleSpinBox()
        self.spin_eco_interval.setValue(2.0)
        f_layout2.addWidget(self.spin_eco_interval)
        opt_layout.addLayout(f_layout2)
        opt_group.setLayout(opt_layout)
        main_layout.addWidget(opt_group)

        self.check_scroll_led = QCheckBox("Блимати Scroll Lock при русі")
        self.check_scroll_led.setChecked(True)
        main_layout.addWidget(self.check_scroll_led)

        main_layout.addStretch()
        self.setCentralWidget(central_widget)

    def toggle_main_system(self):
        self.is_active = not self.is_active
        if self.is_active:
            self.btn_main_toggle.setText("ПРАЦЮЄ")
            self.btn_main_toggle.setStyleSheet("font-weight: bold; font-size: 14pt; background-color: #d4edda; color: #155724; border: 2px solid #155724;")
            if self.cap is None: self.toggle_camera()
        else:
            self.btn_main_toggle.setText("ЗУПИНЕНО")
            self.btn_main_toggle.setStyleSheet("font-weight: bold; font-size: 14pt; background-color: #f8d7da; color: #721c24; border: 2px solid #721c24;")
            if self.cap is not None: self.toggle_camera()
            self.seconds_without_motion = 0
            self.video_msg_label.setText("")

    def on_hid_event(self, *args):
        if self.is_active: self.hid_active = True

    def create_tray_icon(self):
        pix = QPixmap(16, 16)
        pix.fill(QColor("blue"))
        self.tray_icon = QSystemTrayIcon(QIcon(pix), self)
        menu = QMenu()
        menu.addAction("Розгорнути", self.showNormal)
        menu.addAction("Вихід", self.close_app)
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.show()

    def refresh_cameras(self):
        for i in range(2):
            cap = cv2.VideoCapture(i, cv2.CAP_DSHOW)
            if cap.isOpened():
                self.cam_select.addItem(f"Камера {i}", i)
                cap.release()

    def toggle_camera(self):
        if self.cap is None:
            idx = self.cam_select.currentData()
            self.cap = cv2.VideoCapture(idx, cv2.CAP_DSHOW)
            self.cam_timer.start(33)
            self.btn_cam.setText("Відключити")
        else:
            self.cam_timer.stop()
            self.cap.release()
            self.cap = None
            self.btn_cam.setText("Пуск")
            self.video_frame.setText("Камера вимкнена")

    def capture_reference(self):
        self.btn_ref.setText("ЧЕРЕЗ 5 СЕКУНД\nСФОТОГРАФУЄМО ПУСТУ КІМНАТУ\n(вийди з кадру, розбійнику!)")
        QTimer.singleShot(5000, self._do_capture_ref)

    def _do_capture_ref(self):
        if self.cap:
            ret, frame = self.cap.read()
            if ret:
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                self.reference_frame = cv2.GaussianBlur(gray, (21, 21), 0)
                rgb_ref = cv2.cvtColor(self.reference_frame, cv2.COLOR_GRAY2RGB)
                h, w, ch = rgb_ref.shape
                img_ref = QImage(rgb_ref.data, w, h, ch * w, QImage.Format_RGB888)
                self.ref_frame_label.setPixmap(QPixmap.fromImage(img_ref).scaled(120, 90, Qt.KeepAspectRatio))
                self.check_ref_mode.setChecked(True)
                self.btn_ref.setText("Запам'ятати\nпорожню кімнату")

    def update_camera(self):
        if self.cap is None or not self.is_active: return
        ret, frame = self.cap.read()
        if not ret: return

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (21, 21), 0)
        if self.avg_frame is None: self.avg_frame = gray.copy().astype("float")
        cv2.accumulateWeighted(gray, self.avg_frame, 0.5)
        delta = cv2.absdiff(gray, cv2.convertScaleAbs(self.avg_frame))
        thresh = cv2.threshold(delta, 25, 255, cv2.THRESH_BINARY)[1]
        self.motion_detected = np.sum(thresh) > 12000

        self.presence_detected = False
        if self.reference_frame is not None and self.check_ref_mode.isChecked():
            diff = cv2.absdiff(self.reference_frame, gray)
            ref_thresh = cv2.threshold(diff, 30, 255, cv2.THRESH_BINARY)[1]
            if np.sum(ref_thresh) > 50000: self.presence_detected = True

        # Динамічний інтервал
        interval = 33
        if self.check_optimize.isChecked():
            interval = int(self.spin_eco_interval.value() * 1000) if self.presence_detected else int(1000 / self.spin_fps.value())
        if self.cam_timer.interval() != interval: self.cam_timer.start(interval)

        if self.motion_detected:
            cv2.putText(frame, "MOTION!", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)
            if self.check_scroll_led.isChecked(): self.toggle_scroll_lock()
        
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb.shape
        img = QImage(rgb.data, w, h, ch * w, QImage.Format_RGB888)
        self.video_frame.setPixmap(QPixmap.fromImage(img).scaled(320, 240, Qt.KeepAspectRatio))

    def toggle_scroll_lock(self):
        ctypes.windll.user32.keybd_event(self.VK_SCROLL, 0x46, 0, 0)
        ctypes.windll.user32.keybd_event(self.VK_SCROLL, 0x46, self.KEYEVENTF_KEYUP, 0)

    def process_logic(self):
        if not self.is_active: return
        try:
            # Для виявлення ПАУЗИ в йотубі шукаємо прапор EXECUTION (виконання)
            res = subprocess.run(['powercfg', '/requests'], capture_output=True, text=True, encoding='cp866')
            # Перевіряємо DISPLAY (для плеєрів) або EXECUTION/ВЫПОЛНЕНИЕ (для браузерів)
            m_disp = re.search(r'DISPLAY:\s*(.*?)(?=\n[A-ZА-Я]{3,}:|\Z)', res.stdout, re.DOTALL)
            m_exec = re.search(r'(?:ВЫПОЛНЕНИЕ|EXECUTION):\s*(.*?)(?=\n[A-ZА-Я]{3,}:|\Z)', res.stdout, re.DOTALL)
            
            disp_active = m_disp and m_disp.group(1).strip() not in ["Нет.", "None", ""]
            exec_active = m_exec and m_exec.group(1).strip() not in ["Нет.", "None", ""]
            
            self.video_active = disp_active or exec_active
        except: self.video_active = False

        self.video_msg_label.setText("⚠️ ВИЯВЛЕНО ВІДЕО" if self.video_active else "")

        if self.motion_detected: self.continuous_motion_sec += 1
        else: self.continuous_motion_sec = 0

        can_wake_up = False
        if self.hid_active: can_wake_up = True
        elif self.motion_detected:
            if self.check_smart_wake.isChecked() and self.monitor_is_off:
                if (time.time() - self.off_start_time) > (self.spin_smart_off_min.value() * 60):
                    if self.continuous_motion_sec >= self.spin_smart_motion_sec.value(): can_wake_up = True
                else: can_wake_up = True
            else: can_wake_up = True

        is_active = self.motion_detected or self.hid_active or self.video_active or (self.presence_detected if self.check_ref_mode.isChecked() else False)

        if is_active:
            self.seconds_without_motion = 0
            if self.monitor_is_off and can_wake_up and self.check_auto_on.isChecked():
                self.wake_up_monitor()
            self.hid_active = False
        else:
            self.seconds_without_motion += 1

        self.status_label.setText(f"Рух: {'Так' if self.motion_detected else 'Ні'} | Ввід: {'Так' if self.hid_active else 'Ні'} | Presence: {'Так' if self.presence_detected else 'Ні'}")
        self.timer_label.setText(f"Без активності: {self.seconds_without_motion} сек")

        if self.check_auto_off.isChecked() and self.seconds_without_motion >= self.spin_timeout.value():
            if not self.video_active and not self.monitor_is_off:
                self.turn_off_monitor()

    def turn_off_monitor(self):
        ctypes.windll.user32.SendMessageW(0xFFFF, 0x0112, 0xF170, 2)
        self.monitor_is_off = True
        self.off_start_time = time.time()

    def wake_up_monitor(self):
        ctypes.windll.user32.keybd_event(0x10, 0, 0, 0)
        ctypes.windll.user32.keybd_event(0x10, 0, 2, 0)
        self.monitor_is_off = False

    def close_app(self):
        if self.cap: self.cap.release()
        self.kb_listener.stop()
        self.mouse_listener.stop()
        QApplication.quit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    win = MonitorApp()
    win.show()
    sys.exit(app.exec_())
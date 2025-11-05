# diploma_generator.py
# Восстановлено из .exe, очищено, работает на Python 3.13
# Требует: wxPython, python-pptx, openpyxl, comtypes, psutil

import os
import json
import threading
import queue
import time
import psutil
import wx
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import comtypes.client
from datetime import datetime
import re


# === КОНВЕРТАЦИЯ PPTX → PDF ===
def pptx_to_pdf(input_pptx, output_pdf, stop_event):
    powerpoint = None
    try:
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 0
        deck = powerpoint.Presentations.Open(input_pptx, WithWindow=False)

        if stop_event.is_set():
            deck.Close()
            powerpoint.Quit()
            return False

        deck.SaveAs(output_pdf, 17)  # 17 = PDF
        deck.Close()
        powerpoint.Quit()
        return True

    except Exception as e:
        # Убиваем зависшие PowerPoint
        for proc in psutil.process_iter(['name']):
            if proc.info['name'].lower() == 'powerpnt.exe':
                proc.terminate()
                try:
                    proc.wait(timeout=3)
                except psutil.TimeoutExpired:
                    proc.kill()
        if powerpoint:
            try:
                powerpoint.Quit()
            except:
                pass
        raise Exception(f"Ошибка конвертации: {e}")


# === ЗАМЕНА ТЕКСТА ===
def replace_text(shape, placeholder, value, font_settings=None):
    if not shape.has_text_frame:
        return

    for paragraph in shape.text_frame.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        if placeholder in full_text:
            new_text = full_text.replace(placeholder, str(value))

            paragraph.clear()
            run = paragraph.add_run()
            run.text = new_text

            if font_settings and font_settings.get("use_custom"):
                run.font.name = font_settings["name"]
                run.font.size = Pt(font_settings["size"])
                run.font.bold = font_settings.get("bold", False)
            else:
                run.font.color.rgb = RGBColor(127, 127, 127)

            paragraph.alignment = PP_ALIGN.CENTER


# === ОСНОВНАЯ ГЕНЕРАЦИЯ ===
def generate_diplomas(
    excel_path, ppt_template, output_dir, column_mapping,
    error_handling, default_values, font_settings,
    sort_column, enable_sorting,
    log_queue, progress_queue, eta_queue, stop_event
):
    try:
        wb = load_workbook(excel_path, read_only=True)
        ws = wb.active

        # Заголовки
        headers = []
        for i, cell in enumerate(ws[1]):
            headers.append(cell.value or f"Столбец {chr(65 + i)}")

        participants = []
        skipped = []

        # Читаем строки
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if stop_event.is_set():
                log_queue.put("Генерация прервана")
                return False

            row_dict = dict(zip(headers, row))
            skip = False

            for ph, col in column_mapping.items():
                val = row_dict.get(col)
                if val in (None, ""):
                    if error_handling == "stop":
                        skipped.append(f"Строка {row_idx}: пустое поле '{ph}'")
                        skip = True
                        break
                    elif error_handling == "default":
                        row_dict[col] = default_values.get(ph, "—")

            if not skip:
                participants.append((row_idx, row_dict))

        if skipped:
            log_queue.put(f"Пропущено строк: {len(skipped)}")

        os.makedirs(output_dir, exist_ok=True)
        total = len(participants)
        times = []

        for idx, (row_idx, data) in enumerate(participants, 1):
            if stop_event.is_set():
                log_queue.put("Прервано")
                return False

            start = time.time()

            # Копируем шаблон
            temp_pptx = f"temp_{idx}.pptx"
            prs = Presentation(ppt_template)

            # Замена
            for shape in prs.slides[0].shapes:
                for ph, col in column_mapping.items():
                    val = data.get(col, "")
                    replace_text(shape, f"{{{ph}}}", val, font_settings)

            prs.save(temp_pptx)

            # Папка
            folder = output_dir
            if enable_sorting and sort_column and sort_column in data:
                sub = str(data[sort_column]).strip()
                if sub:
                    folder = os.path.join(output_dir, sub)
                    os.makedirs(folder, exist_ok=True)

            # Имя файла
            name = data.get(column_mapping.get("NAME", ""), "diploma")
            name = re.sub(r'[<>:"/\\|?*]', "_", str(name))
            pdf_path = os.path.join(folder, f"{name}.pdf")

            # Конвертация
            if not pptx_to_pdf(temp_pptx, pdf_path, stop_event):
                log_queue.put(f"Ошибка: {name}.pdf")
            else:
                log_queue.put(f"Сохранено: {name}.pdf")

            os.remove(temp_pptx)

            # Прогресс
            elapsed = time.time() - start
            times.append(elapsed)
            avg = sum(times) / len(times)
            remain = int((total - idx) * avg)
            m, s = divmod(remain, 60)
            eta_queue.put(f"{m:02d}:{s:02d}")
            progress_queue.put(int(idx / total * 100))

        log_queue.put(f"Готово! Сохранено: {total}")
        eta_queue.put("00:00")
        progress_queue.put(100)
        return True

    except Exception as e:
        log_queue.put(f"Ошибка: {e}")
        return False


# === GUI ===
class DiplomaGeneratorApp(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Генератор дипломов", size=(1000, 400))
        self.SetMinSize((1000, 400))
        self.SetMaxSize((1000, 400))

        self.excel_path = self.pptx_path = self.output_dir = ""
        self.column_mapping = {}
        self.placeholders = []
        self.error_handling = "stop"
        self.default_values = {}
        self.font_settings = {"use_custom": False}
        self.sort_column = ""
        self.enable_sorting = True

        self.stop_event = threading.Event()
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.eta_queue = queue.Queue()
        self.thread = None

        self.setup_ui()
        self.load_config()
        self.Bind(wx.EVT_CLOSE, self.on_close)

        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.update_ui, self.timer)
        self.timer.Start(100)

    def setup_ui(self):
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Левая панель
        left = wx.Panel(panel)
        grid = wx.GridBagSizer(5, 5)

        # Excel
        grid.Add(wx.StaticText(left, label="Excel:"), (0,0))
        self.excel_name = wx.StaticText(left, label="Не выбран")
        grid.Add(self.excel_name, (0,1))
        self.excel_ctrl = wx.TextCtrl(left, style=wx.TE_READONLY)
        grid.Add(self.excel_ctrl, (0,2), flag=wx.EXPAND)
        btn = wx.Button(left, label="Выбрать", size=(100,30))
        btn.Bind(wx.EVT_BUTTON, self.choose_excel)
        grid.Add(btn, (0,3))

        # PPTX
        grid.Add(wx.StaticText(left, label="PPTX:"), (1,0))
        self.pptx_name = wx.StaticText(left, label="Не выбран")
        grid.Add(self.pptx_name, (1,1))
        self.pptx_ctrl = wx.TextCtrl(left, style=wx.TE_READONLY)
        grid.Add(self.pptx_ctrl, (1,2), flag=wx.EXPAND)
        btn = wx.Button(left, label="Выбрать", size=(100,30))
        btn.Bind(wx.EVT_BUTTON, self.choose_pptx)
        grid.Add(btn, (1,3))

        # Папка
        grid.Add(wx.StaticText(left, label="Папка:"), (2,0))
        self.out_name = wx.StaticText(left, label="Не выбрана")
        grid.Add(self.out_name, (2,1))
        self.out_ctrl = wx.TextCtrl(left, style=wx.TE_READONLY)
        grid.Add(self.out_ctrl, (2,2), flag=wx.EXPAND)
        btn = wx.Button(left, label="Выбрать", size=(100,30))
        btn.Bind(wx.EVT_BUTTON, self.choose_output)
        grid.Add(btn, (2,3))

        # Ошибки
        grid.Add(wx.StaticText(left, label="Ошибки:"), (3,0))
        self.err_choice = wx.Choice(left, choices=["Остановить", "Пропустить", "По умолчанию"])
        self.err_choice.SetSelection(0)
        self.err_choice.Bind(wx.EVT_CHOICE, self.on_error_change)
        grid.Add(self.err_choice, (3,1), span=(1,2), flag=wx.EXPAND)

        # Сортировка
        grid.Add(wx.StaticText(left, label="Сортировка:"), (4,0))
        self.sort_check = wx.CheckBox(left, label="В папки")
        self.sort_check.Bind(wx.EVT_CHECKBOX, self.on_sort_change)
        grid.Add(self.sort_check, (4,1))
        self.sort_choice = wx.Choice(left)
        self.sort_choice.Bind(wx.EVT_CHOICE, self.on_sort_change)
        grid.Add(self.sort_choice, (4,2))

        left.SetSizer(grid)

        # Правая панель
        right = wx.Panel(panel)
        vsizer = wx.BoxSizer(wx.VERTICAL)
        vsizer.Add(wx.StaticText(right, label="Прогресс:"), flag=wx.LEFT, border=5)
        self.gauge = wx.Gauge(right, range=100)
        vsizer.Add(self.gauge, flag=wx.EXPAND)
        self.eta = wx.StaticText(right, label="Осталось: 00:00")
        vsizer.Add(self.eta, flag=wx.TOP, border=5)
        vsizer.Add(wx.StaticText(right, label="Лог:"), flag=wx.LEFT|wx.TOP, border=5)
        self.log = wx.TextCtrl(right, style=wx.TE_MULTILINE|wx.TE_READONLY, size=(-1,150))
        vsizer.Add(self.log, proportion=1, flag=wx.EXPAND)
        right.SetSizer(vsizer)

        # Кнопки
        btns = wx.BoxSizer(wx.HORIZONTAL)
        self.map_btn = wx.Button(panel, label="Сопоставление")
        self.map_btn.Bind(wx.EVT_BUTTON, self.open_mapping)
        self.map_btn.Enable(False)
        self.gen_btn = wx.Button(panel, label="Запустить")
        self.gen_btn.Bind(wx.EVT_BUTTON, self.start_gen)
        self.gen_btn.Enable(False)
        self.stop_btn = wx.Button(panel, label="Прервать")
        self.stop_btn.Bind(wx.EVT_BUTTON, self.stop_gen)
        self.stop_btn.Enable(False)
        btns.Add(self.map_btn, flag=wx.RIGHT, border=10)
        btns.Add(self.gen_btn, flag=wx.RIGHT, border=10)
        btns.Add(self.stop_btn)

        # Финал
        sizer.Add(left, proportion=1, flag=wx.EXPAND|wx.ALL, border=10)
        sizer.Add(right, proportion=1, flag=wx.EXPAND|wx.ALL, border=10)
        main = wx.BoxSizer(wx.VERTICAL)
        main.Add(sizer, proportion=1, flag=wx.EXPAND)
        main.Add(btns, flag=wx.ALIGN_CENTER|wx.ALL, border=10)
        panel.SetSizer(main)

    def update_ui(self, evt):
        while True:
            try: msg = self.log_queue.get_nowait()
            except: break
            else: self.log.AppendText(f"{datetime.now():%H:%M:%S} | {msg}\n")
        while True:
            try: p = self.progress_queue.get_nowait()
            except: break
            else: self.gauge.SetValue(p)
        while True:
            try: e = self.eta_queue.get_nowait()
            except: break
            else: self.eta.SetLabel(f"Осталось: {e}")

    def choose_excel(self, e): self._choose_file("excel", "*.xlsx", self.choose_excel_cb)
    def choose_pptx(self, e): self._choose_file("pptx", "*.pptx", self.choose_pptx_cb)
    def choose_output(self, e): self._choose_dir(self.choose_output_cb)

    def _choose_file(self, attr, wc, cb):
        with wx.FileDialog(self, "Выбрать", wildcard=wc, style=wx.FD_OPEN) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                path = dlg.GetPath()
                setattr(self, attr + "_path", path)
                getattr(self, attr + "_ctrl").SetValue(path)
                getattr(self, attr + "_name").SetLabel(os.path.basename(path))
                self.log_queue.put(f"{attr.upper()}: {os.path.basename(path)}")
                cb()

    def _choose_dir(self, cb):
        with wx.DirDialog(self, "Папка") as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self.output_dir = dlg.GetPath()
                self.out_ctrl.SetValue(self.output_dir)
                self.out_name.SetLabel(os.path.basename(self.output_dir) or "Папка")
                self.log_queue.put(f"Папка: {self.output_dir}")
                cb()

    def choose_excel_cb(self): self.update_buttons()
    def choose_pptx_cb(self): self.scan_placeholders(); self.update_buttons()
    def choose_output_cb(self): self.update_buttons()

    def scan_placeholders(self):
        try:
            prs = Presentation(self.pptx_path)
            self.placeholders = []
            for shape in prs.slides[0].shapes:
                if shape.has_text_frame:
                    text = shape.text_frame.text
                    self.placeholders.extend(re.findall(r"\{([^}]+)\}", text))
            self.placeholders = list(set(self.placeholders))
            self.log_queue.put(f"Плейсхолдеры: {', '.join(self.placeholders)}")
            self.sort_choice.Set(self.placeholders)
            self.sort_check.Enable(True)
            self.sort_choice.Enable(True)
        except Exception as e:
            self.log_queue.put(f"Ошибка шаблона: {e}")

    def update_buttons(self):
        ready = bool(self.excel_path and self.pptx_path and self.output_dir and self.placeholders)
        self.map_btn.Enable(ready)
        self.gen_btn.Enable(ready and self.column_mapping)

    def on_error_change(self, e):
        self.error_handling = ["stop", "skip", "default"][self.err_choice.GetSelection()]
        if self.error_handling == "default":
            self.open_defaults()

    def on_sort_change(self, e):
        self.enable_sorting = self.sort_check.GetValue()
        self.sort_column = self.sort_choice.GetStringSelection()

    def open_defaults(self):
        dlg = wx.Dialog(self, title="Значения по умолчанию")
        sizer = wx.BoxSizer(wx.VERTICAL)
        entries = {}
        for ph in self.placeholders:
            hs = wx.BoxSizer(wx.HORIZONTAL)
            hs.Add(wx.StaticText(dlg, label=f"{ph}:"), flag=wx.RIGHT, border=5)
            txt = wx.TextCtrl(dlg, value=self.default_values.get(ph, ""))
            hs.Add(txt, proportion=1, flag=wx.EXPAND)
            entries[ph] = txt
            sizer.Add(hs, flag=wx.EXPAND|wx.ALL, border=5)
        btn = wx.Button(dlg, label="Сохранить")
        btn.Bind(wx.EVT_BUTTON, lambda e: [self.default_values.update({p: t.GetValue() for p,t in entries.items()}), dlg.Destroy()])
        sizer.Add(btn, flag=wx.ALIGN_CENTER|wx.ALL, border=10)
        dlg.SetSizer(sizer)
        dlg.Fit()
        dlg.ShowModal()

    def open_mapping(self, e):
        dlg = wx.Dialog(self, title="Сопоставление", size=(600,400))
        sizer = wx.BoxSizer(wx.VERTICAL)
        grid = wx.GridBagSizer(5,5)
        self.mapping_choices = {}
        wb = load_workbook(self.excel_path, read_only=True)
        headers = [c.value or f"Колонка {i+1}" for i, c in enumerate(wb.active[1])]

        for i, ph in enumerate(self.placeholders):
            grid.Add(wx.StaticText(dlg, label=f"{ph}:"), (i,0), flag=wx.ALIGN_CENTER_VERTICAL)
            ch = wx.Choice(dlg, choices=["Игнорировать"] + headers)
            ch.SetSelection(0)
            self.mapping_choices[ph] = ch
            grid.Add(ch, (i,1), flag=wx.EXPAND)

        sizer.Add(grid, proportion=1, flag=wx.EXPAND|wx.ALL, border=10)
        btns = wx.BoxSizer(wx.HORIZONTAL)
        auto = wx.Button(dlg, label="Авто")
        auto.Bind(wx.EVT_BUTTON, lambda e: self.auto_map(headers))
        save = wx.Button(dlg, label="Сохранить")
        save.Bind(wx.EVT_BUTTON, lambda e: [self.save_mapping(), dlg.Destroy()])
        btns.Add(auto, flag=wx.RIGHT, border=5)
        btns.Add(save)
        sizer.Add(btns, flag=wx.ALIGN_CENTER|wx.ALL, border=10)
        dlg.SetSizer(sizer)
        dlg.Fit()
        dlg.ShowModal()

    def auto_map(self, headers):
        synonyms = {
            "NAME": ["ФИО", "Имя", "Name"],
            "DATE": ["Дата", "Date"],
            "TIME": ["Часы", "Hours"]
        }
        for ph, ch in self.mapping_choices.items():
            for h in headers:
                if any(syn.lower() in h.lower() for syn in synonyms.get(ph, [ph])):
                    ch.SetStringSelection(h)
                    break

    def save_mapping(self):
        self.column_mapping = {ph: ch.GetStringSelection() for ph, ch in self.mapping_choices.items() if ch.GetStringSelection() != "Игнорировать"}
        self.update_buttons()

    def start_gen(self, e):
        if not self.column_mapping:
            wx.MessageBox("Сопоставьте поля!", "Ошибка", wx.OK|wx.ICON_ERROR)
            return
        self.stop_event.clear()
        self.thread = threading.Thread(target=self.run_gen, daemon=True)
        self.thread.start()
        self.gen_btn.Enable(False)
        self.stop_btn.Enable(True)

    def run_gen(self):
        success = generate_diplomas(
            self.excel_path, self.pptx_path, self.output_dir,
            self.column_mapping, self.error_handling, self.default_values,
            self.font_settings, self.sort_column, self.enable_sorting,
            self.log_queue, self.progress_queue, self.eta_queue, self.stop_event
        )
        wx.CallAfter(self.finish_gen, success)

    def finish_gen(self, success):
        self.gen_btn.Enable(bool(self.column_mapping))
        self.stop_btn.Enable(False)
        if success:
            wx.MessageBox("Готово!", "Успех", wx.OK|wx.ICON_INFORMATION)

    def stop_gen(self, e):
        self.stop_event.set()
        self.log_queue.put("Прерывание...")
        self.stop_btn.Enable(False)

    def on_close(self, e):
        if self.thread and self.thread.is_alive():
            self.stop_event.set()
            self.thread.join(timeout=3)
        for f in os.listdir():
            if f.startswith("temp_") and f.endswith((".pptx", ".pdf")):
                try: os.remove(f)
                except: pass
        self.timer.Stop()
        self.Destroy()

    def load_config(self):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                cfg = json.load(f)
                for k in ["excel_path", "pptx_path", "output_dir", "column_mapping", "error_handling", "default_values", "sort_column", "enable_sorting"]:
                    if k in cfg:
                        setattr(self, k, cfg[k])
                # Восстановление UI
                if self.excel_path: self.excel_ctrl.SetValue(self.excel_path); self.excel_name.SetLabel(os.path.basename(self.excel_path))
                if self.pptx_path: self.pptx_ctrl.SetValue(self.pptx_path); self.pptx_name.SetLabel(os.path.basename(self.pptx_path)); self.scan_placeholders()
                if self.output_dir: self.out_ctrl.SetValue(self.output_dir); self.out_name.SetLabel(os.path.basename(self.output_dir) or "Папка")
                self.err_choice.SetSelection(["stop", "skip", "default"].index(self.error_handling))
                self.sort_check.SetValue(self.enable_sorting)
                if self.sort_column and self.placeholders: self.sort_choice.SetSelection(self.placeholders.index(self.sort_column))
                self.update_buttons()
        except: pass


if __name__ == "__main__":
    app = wx.App()
    frame = DiplomaGeneratorApp()
    frame.Show()
    app.MainLoop()

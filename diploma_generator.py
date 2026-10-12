"""Desktop interface for both diploma generation modes."""
import multiprocessing
if __name__ == '__main__':
    # Frozen Office workers must skip loading the GUI and its libraries.
    multiprocessing.freeze_support()

import copy
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import queue
import threading
import sys
from datetime import datetime
from app_version import VERSION

import wx
import wx.lib.scrolledpanel

from diploma_core import (
    Cancelled, DiplomaError, MODES, POLICIES, Settings, app_data_dir,
    auto_mapping, generate, load_settings, prepare_plan, read_excel,
    read_template, save_settings,
)

BG = '#F2F5FA'
INK = '#19283F'
MUTED = '#65758B'
ACCENT = '#375BDD'


class DiplomaGeneratorApp(wx.Frame):
    def __init__(self):
        super().__init__(None, title='Генератор дипломов', size=(1180, 850))
        self.SetMinSize((1000, 760))
        icon_path = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent)) / 'assets' / 'app.ico'
        if icon_path.is_file():
            self.SetIcon(wx.Icon(str(icon_path), wx.BITMAP_TYPE_ICO))
        self.settings = Settings()
        self.events = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None
        self.busy = False
        self.closing = False
        self.locked_controls = []
        self.last_output = ''
        self.setup_ui()
        try:
            self.settings = load_settings()
        except DiplomaError as exc:
            self.log(str(exc))
        self.apply_settings()
        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.on_timer, self.timer)
        self.timer.Start(100)
        self.Centre()

    def label(self, parent, text, size=10, colour=INK, bold=False):
        item = wx.StaticText(parent, label=text)
        item.SetForegroundColour(colour)
        item.SetFont(wx.Font(size, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD if bold else wx.FONTWEIGHT_NORMAL,
                             faceName='Segoe UI'))
        return item

    def button(self, parent, text, handler, primary=False):
        item = wx.Button(parent, label=text, size=(-1, 36))
        item.Bind(wx.EVT_BUTTON, handler)
        if primary:
            item.SetBackgroundColour(ACCENT)
            item.SetForegroundColour('white')
            item.SetFont(wx.Font(10, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL,
                                 wx.FONTWEIGHT_BOLD, faceName='Segoe UI'))
        return item

    def card(self, parent, sizer, title, subtitle=None, proportion=0):
        panel = wx.Panel(parent)
        panel.SetBackgroundColour('white')
        outer = wx.BoxSizer(wx.VERTICAL)
        outer.Add(self.label(panel, title, 13, bold=True), 0, wx.LEFT | wx.RIGHT | wx.TOP, 14)
        if subtitle:
            label = self.label(panel, subtitle, 9, MUTED)
            label.Wrap(480)
            outer.Add(label, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
        content = wx.BoxSizer(wx.VERTICAL)
        outer.Add(content, 1, wx.EXPAND | wx.ALL, 14)
        panel.SetSizer(outer)
        sizer.Add(panel, proportion, wx.EXPAND | wx.BOTTOM, 12)
        return panel, content

    def setup_ui(self):
        root = wx.Panel(self)
        root.SetBackgroundColour(BG)
        root.SetFont(wx.Font(10, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL,
                            wx.FONTWEIGHT_NORMAL, faceName='Segoe UI'))
        outer = wx.BoxSizer(wx.VERTICAL)
        header = wx.BoxSizer(wx.HORIZONTAL)
        heading = wx.BoxSizer(wx.VERTICAL)
        heading.Add(self.label(root, 'Дипломы для каждого', 25, bold=True))
        heading.Add(self.label(root, 'Excel → шаблон PowerPoint → готовые PDF', 11, MUTED), 0, wx.TOP, 6)
        header.Add(heading, 1)
        self.badge = self.label(root, 'ГОТОВ К РАБОТЕ', 9, ACCENT, True)
        header.Add(self.badge, 0, wx.ALIGN_CENTER_VERTICAL)
        outer.Add(header, 0, wx.EXPAND | wx.ALL, 24)
        body = wx.BoxSizer(wx.HORIZONTAL)
        left = wx.lib.scrolledpanel.ScrolledPanel(root)
        left.SetBackgroundColour(BG)
        left_sizer = wx.BoxSizer(wx.VERTICAL)
        panel, content = self.card(left, left_sizer, '01  Исходные файлы')
        self.paths = {}
        for key, title, hint in (
            ('excel_path', 'Таблица участников', 'Excel · .xlsx'),
            ('pptx_path', 'Шаблон диплома', 'PowerPoint · .pptx · первый слайд'),
            ('output_dir', 'Папка результата', 'Отдельный PDF для каждого участника'),
        ):
            content.Add(self.label(panel, title, 10, bold=True), 0, wx.BOTTOM, 4)
            line = wx.BoxSizer(wx.HORIZONTAL)
            field = wx.TextCtrl(panel)
            field.SetHint(hint)
            field.SetMinSize((-1, 32))
            field.Bind(wx.EVT_TEXT, self.on_path_changed)
            browse = self.button(panel, 'Выбрать…', lambda event, k=key: self.choose_path(k))
            line.Add(field, 1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 8)
            line.Add(browse, 0)
            content.Add(line, 0, wx.EXPAND | wx.BOTTOM, 12)
            self.paths[key] = field
            self.locked_controls.extend([field, browse])
        panel, content = self.card(left, left_sizer, '02  Режим генерации')
        self.mode = wx.Choice(panel, choices=list(MODES.values()))
        self.mode.Bind(wx.EVT_CHOICE, self.on_mode_changed)
        content.Add(self.mode, 0, wx.EXPAND)
        self.mode_hint = self.label(panel, '', 9, MUTED)
        content.Add(self.mode_hint, 0, wx.EXPAND | wx.TOP | wx.BOTTOM, 10)
        self.mapping_btn = self.button(panel, 'Настроить поля и значения по умолчанию', self.open_mapping)
        content.Add(self.mapping_btn, 0, wx.EXPAND)
        self.locked_controls.extend([self.mode, self.mapping_btn])
        panel, content = self.card(left, left_sizer, '03  Правила и оформление')
        content.Add(self.label(panel, 'Если обязательное поле пустое', 9, MUTED), 0, wx.BOTTOM, 5)
        self.policy = wx.Choice(panel, choices=list(POLICIES.values()))
        content.Add(self.policy, 0, wx.EXPAND | wx.BOTTOM, 12)
        self.fit = wx.CheckBox(panel, label='Уменьшать длинный текст без переноса строк')
        self.fit.SetToolTip('Сохраняет шрифт и оформление шаблона. Уменьшает только размер, если текст не помещается. Явные переносы в шаблоне остаются.')
        content.Add(self.fit, 0, wx.BOTTOM, 12)
        self.sort = wx.CheckBox(panel, label='Распределять PDF по папкам')
        self.sort.Bind(wx.EVT_CHECKBOX, self.on_sort_changed)
        content.Add(self.sort, 0, wx.BOTTOM, 5)
        self.sort_field = wx.ComboBox(panel, style=wx.CB_READONLY)
        content.Add(self.sort_field, 0, wx.EXPAND)
        self.locked_controls.extend([self.policy, self.fit, self.sort, self.sort_field])
        left.SetSizer(left_sizer)
        left.SetupScrolling(scroll_x=False)
        right = wx.lib.scrolledpanel.ScrolledPanel(root)
        right.SetBackgroundColour(BG)
        right_sizer = wx.BoxSizer(wx.VERTICAL)
        panel, content = self.card(right, right_sizer, 'Проверка перед запуском',
                                   'Проверьте поля и количество дипломов до экспорта.')
        self.summary = self.label(panel, 'Выберите таблицу и шаблон.', 11, MUTED)
        self.summary.Wrap(420)
        content.Add(self.summary, 0, wx.EXPAND | wx.BOTTOM, 16)
        self.check_btn = self.button(panel, 'Проверить данные', self.check_data)
        content.Add(self.check_btn, 0, wx.EXPAND)
        self.locked_controls.append(self.check_btn)
        panel, content = self.card(right, right_sizer, 'Генерация')
        self.status = self.label(panel, 'Ожидание запуска', 12, bold=True)
        content.Add(self.status, 0, wx.BOTTOM, 12)
        self.gauge = wx.Gauge(panel, range=100, size=(-1, 12))
        content.Add(self.gauge, 0, wx.EXPAND | wx.BOTTOM, 10)
        self.progress_text = self.label(panel, '0 из 0  ·  осталось —', 9, MUTED)
        content.Add(self.progress_text, 0, wx.BOTTOM, 14)
        actions = wx.BoxSizer(wx.HORIZONTAL)
        self.start_btn = self.button(panel, 'Создать дипломы', self.start_generation, True)
        self.stop_btn = self.button(panel, 'Остановить', self.stop_generation)
        self.stop_btn.Enable(False)
        actions.Add(self.start_btn, 1, wx.RIGHT, 8)
        actions.Add(self.stop_btn, 0)
        content.Add(actions, 0, wx.EXPAND)
        self.open_btn = self.button(panel, 'Открыть папку результата', self.open_output)
        content.Add(self.open_btn, 0, wx.EXPAND | wx.TOP, 8)
        panel, content = self.card(right, right_sizer, 'Журнал операций', proportion=1)
        self.log_ctrl = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_DONTWRAP)
        self.log_ctrl.SetMinSize((-1, 100))
        self.log_ctrl.SetBackgroundColour('#F8FAFD')
        self.log_ctrl.SetFont(wx.Font(9, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL,
                                     wx.FONTWEIGHT_NORMAL, faceName='Consolas'))
        content.Add(self.log_ctrl, 1, wx.EXPAND)
        log_button = self.button(panel, 'Открыть подробный журнал', self.open_log)
        content.Add(log_button, 0, wx.EXPAND | wx.TOP, 8)
        right.SetSizer(right_sizer)
        right.SetupScrolling(scroll_x=False)
        body.Add(left, 6, wx.EXPAND | wx.RIGHT, 20)
        body.Add(right, 5, wx.EXPAND)
        outer.Add(body, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 24)
        outer.Add(self.label(root, f'Версия {VERSION} · Шаблон сохраняет своё оформление. Для экспорта нужен Microsoft PowerPoint.', 9, MUTED),
                  0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 24)
        root.SetSizer(outer)
        self.SetBackgroundColour(BG)

    def apply_settings(self):
        for key, field in self.paths.items():
            field.ChangeValue(getattr(self.settings, key))
        self.mode.SetSelection(list(MODES).index(self.settings.mode))
        self.policy.SetSelection(list(POLICIES).index(self.settings.error_handling))
        self.fit.SetValue(self.settings.auto_fit)
        self.sort.SetValue(self.settings.enable_sorting)
        self.refresh_mode_hint()
        self.refresh_sort_fields()

    def collect_settings(self):
        settings = copy.deepcopy(self.settings)
        for key, field in self.paths.items():
            setattr(settings, key, field.GetValue().strip().strip('"'))
        settings.mode = list(MODES)[self.mode.GetSelection()]
        settings.error_handling = list(POLICIES)[self.policy.GetSelection()]
        settings.auto_fit = self.fit.GetValue()
        settings.enable_sorting = self.sort.GetValue()
        settings.sort_column = self.sort_field.GetValue()
        return settings

    def persist(self):
        self.settings = self.collect_settings()
        try:
            save_settings(self.settings)
        except Exception as exc:
            logging.exception('Could not save settings')
            self.log(f'Настройки не сохранены: {exc}')

    def refresh_mode_hint(self):
        team = self.mode.GetSelection() == 1
        self.mode_hint.SetLabel(
            'Каждая строка — команда. FullN1, FullN2… → отдельные дипломы.\n{NAME} — участник; {TRAINER} — первый тренер FullTr….' if team else
            'Каждая строка — один диплом, как в orig.\nСвяжите поля шаблона с любыми колонками Excel.')
        self.mode_hint.Wrap(490)
        self.mode_hint.GetParent().Layout()

    def refresh_sort_fields(self):
        selected = self.settings.sort_column or self.sort_field.GetValue()
        choices = set()
        try:
            if self.paths['excel_path'].GetValue():
                headers, _ = read_excel(self.paths['excel_path'].GetValue(), headers_only=True)
                choices.update(headers)
            if self.paths['pptx_path'].GetValue():
                _, placeholders = read_template(self.paths['pptx_path'].GetValue())
                choices.update(placeholders)
        except DiplomaError:
            pass
        if self.mode.GetSelection() == 1:
            choices.update(['NAME', 'TRAINER'])
        self.sort_field.SetItems(sorted(choices))
        if selected in choices:
            self.sort_field.SetStringSelection(selected)
        elif choices:
            self.sort_field.SetStringSelection('TEAMNAME' if 'TEAMNAME' in choices else sorted(choices)[0])
        self.sort_field.Enable(self.sort.GetValue() and not self.busy)

    def on_path_changed(self, event):
        if event.GetEventObject() in [self.paths.get('excel_path'), self.paths.get('pptx_path')]:
            self.settings.column_mapping = {}
        if hasattr(self, 'summary'):
            self.summary.SetLabel('Файлы изменены. Проверьте данные перед запуском.')
        event.Skip()

    def on_mode_changed(self, event):
        self.settings.column_mapping = {}
        self.refresh_mode_hint()
        self.refresh_sort_fields()
        self.summary.SetLabel('Режим изменён. Проверьте сопоставление полей.')

    def on_sort_changed(self, event):
        self.refresh_sort_fields()

    def choose_path(self, key):
        if key == 'output_dir':
            dialog = wx.DirDialog(self, 'Папка для PDF')
        else:
            extension = 'xlsx' if key == 'excel_path' else 'pptx'
            dialog = wx.FileDialog(self, 'Выберите файл', wildcard=f'*.{extension}|*.{extension}',
                                   style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        with dialog:
            if dialog.ShowModal() == wx.ID_OK:
                self.paths[key].SetValue(dialog.GetPath())
                if key in ('excel_path', 'pptx_path'):
                    self.settings.column_mapping = {}
                self.refresh_sort_fields()

    def report_error(self, exc):
        logging.exception('Operation failed')
        message = str(exc) if isinstance(exc, DiplomaError) else f'Не удалось выполнить операцию: {exc}'
        self.log(message)
        wx.MessageBox(message, 'Не удалось продолжить', wx.OK | wx.ICON_ERROR, self)

    def open_mapping(self, event):
        try:
            settings = self.collect_settings()
            headers, _ = read_excel(settings.excel_path, headers_only=True)
            _, placeholders = read_template(settings.pptx_path)
        except Exception as exc:
            self.report_error(exc)
            return
        mapping = auto_mapping(placeholders, headers, settings.mode)
        mapping.update({ph: h for ph, h in settings.column_mapping.items() if h in headers or not h})
        dialog = wx.Dialog(self, title='Поля диплома', size=(820, 600), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        dialog.SetMinSize((680, 400))
        panel = wx.Panel(dialog)
        panel.SetBackgroundColour(BG)
        outer = wx.BoxSizer(wx.VERTICAL)
        outer.Add(self.label(panel, 'Свяжите шаблон с таблицей', 18, bold=True), 0, wx.ALL, 18)
        note = self.label(panel, 'Каждое поле шаблона получает значение из выбранной колонки.\nЗначения справа используются при выборе «Заполнить по умолчанию».', 10, MUTED)
        outer.Add(note, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 18)
        scroll = wx.lib.scrolledpanel.ScrolledPanel(panel)
        grid = wx.FlexGridSizer(cols=3, hgap=12, vgap=12)
        grid.AddGrowableCol(1, 1)
        grid.AddGrowableCol(2, 1)
        for title in ['Поле шаблона', 'Колонка Excel', 'По умолчанию']:
            grid.Add(self.label(scroll, title, 10, bold=True))
        choices, defaults = {}, {}
        for ph in placeholders:
            grid.Add(self.label(scroll, '{' + ph + '}'), 0, wx.ALIGN_CENTER_VERTICAL)
            automatic = settings.mode == 'team' and ph in ('NAME', 'TRAINER')
            choice = wx.Choice(scroll, choices=['Не сопоставлено'] + headers)
            if automatic:
                choice.SetItems(['FullN… (участник)' if ph == 'NAME' else 'FullTr… (первый тренер)'])
                choice.SetSelection(0)
                choice.Enable(False)
            else:
                choice.SetStringSelection(mapping.get(ph, '') or 'Не сопоставлено')
            default = wx.TextCtrl(scroll, value=settings.default_values.get(ph, ''))
            grid.Add(choice, 0, wx.EXPAND)
            grid.Add(default, 0, wx.EXPAND)
            choices[ph], defaults[ph] = choice, default
        scroll.SetSizer(grid)
        scroll.SetupScrolling(scroll_x=False)
        outer.Add(scroll, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 18)
        actions = wx.BoxSizer(wx.HORIZONTAL)
        def autofill(event):
            guessed = auto_mapping(placeholders, headers, settings.mode)
            for ph, choice in choices.items():
                if choice.IsEnabled():
                    choice.SetStringSelection(guessed.get(ph, '') or 'Не сопоставлено')
        actions.Add(self.button(panel, 'Автосопоставление', autofill), 0)
        actions.AddStretchSpacer()
        cancel = wx.Button(panel, wx.ID_CANCEL, 'Отмена')
        save = wx.Button(panel, wx.ID_OK, 'Сохранить')
        save.SetDefault()
        actions.Add(cancel, 0, wx.RIGHT, 8)
        actions.Add(save, 0)
        outer.Add(actions, 0, wx.EXPAND | wx.ALL, 18)
        panel.SetSizer(outer)
        wrapper = wx.BoxSizer(wx.VERTICAL)
        wrapper.Add(panel, 1, wx.EXPAND)
        dialog.SetSizer(wrapper)
        dialog.CentreOnParent()
        with dialog:
            if dialog.ShowModal() == wx.ID_OK:
                self.settings.column_mapping = {ph: (c.GetStringSelection() if c.GetSelection() > 0 else '')
                                                for ph, c in choices.items() if c.IsEnabled()}
                self.settings.default_values = {ph: entry.GetValue() for ph, entry in defaults.items()}
                self.persist()
                self.log('Сопоставление полей сохранено.')
                self.refresh_sort_fields()

    def set_busy(self, busy, checking=False):
        self.busy = busy
        for item in self.locked_controls:
            item.Enable(not busy)
        self.sort_field.Enable(not busy and self.sort.GetValue())
        self.start_btn.Enable(not busy)
        self.stop_btn.Enable(busy)
        self.badge.SetLabel('ПРОВЕРКА' if checking else 'ГЕНЕРАЦИЯ' if busy else 'ГОТОВ К РАБОТЕ')
        self.status.SetLabel('Проверка данных…' if checking else 'Создаём дипломы…' if busy else self.status.GetLabel())

    def launch(self, checking):
        if self.busy:
            return
        self.persist()
        settings = copy.deepcopy(self.settings)
        self.stop_event.clear()
        self.set_busy(True, checking)
        self.gauge.SetValue(0)
        self.progress_text.SetLabel('Подготовка данных…')
        def work():
            try:
                plan = prepare_plan(settings, self.stop_event)
                self.events.put(('plan', plan))
                if checking:
                    self.events.put(('checked', None))
                else:
                    count = generate(settings, self.log,
                                     lambda done, total, eta: self.events.put(('progress', (done, total, eta))),
                                     self.stop_event, plan=plan)
                    self.events.put(('complete', (count, settings.output_dir)))
            except Cancelled:
                self.events.put(('cancelled', None))
            except Exception as exc:
                logging.exception('Background operation failed')
                self.events.put(('error', str(exc)))
            finally:
                self.events.put(('idle', None))
        self.worker = threading.Thread(target=work, name='diploma-generation', daemon=True)
        self.worker.start()

    def check_data(self, event):
        self.launch(True)

    def start_generation(self, event):
        self.launch(False)

    def stop_generation(self, event):
        self.stop_event.set()
        self.stop_btn.Enable(False)
        self.status.SetLabel('Останавливаем…')
        self.log('Запрошена остановка. Уже созданные PDF сохраняются.')

    def log(self, message):
        self.events.put(('log', str(message)))
        logging.info('%s', message)

    def on_timer(self, event):
        for _ in range(100):
            try:
                kind, payload = self.events.get_nowait()
            except queue.Empty:
                break
            if kind == 'log':
                self.log_ctrl.AppendText(f'{datetime.now():%H:%M:%S}  {payload}\n')
            elif kind == 'plan':
                self.checked_warnings = payload.warnings[:30]
                selected = self.sort_field.GetValue()
                choices = sorted(set(payload.headers + payload.placeholders +
                                     (['NAME', 'TRAINER'] if self.settings.mode == 'team' else [])))
                self.sort_field.SetItems(choices)
                if selected in choices:
                    self.sort_field.SetStringSelection(selected)
                self.summary.SetLabel(f'Дипломов: {len(payload.records)}\nПропущено: {payload.skipped} · предупреждений: {len(payload.warnings)}\nПоля: ' + ', '.join(payload.placeholders))
                self.summary.Wrap(420)
                self.summary.GetParent().Layout()
                preview = [f'Строка {r.row}: {r.values.get("NAME", "")}'
                           for r in payload.records[:5]]
                self.log('Первые дипломы:\n' + '\n'.join(preview))
            elif kind == 'progress':
                done, total, eta = payload
                self.gauge.SetValue(round(done / total * 100))
                self.progress_text.SetLabel(f'{done} из {total}  ·  осталось {eta // 60:02d}:{eta % 60:02d}')
            elif kind == 'checked':
                # A check has no export stage to emit the warnings.
                for warning in self.checked_warnings:
                    self.log(warning)
                self.status.SetLabel('Данные проверены')
                self.progress_text.SetLabel('Можно запускать генерацию')
                self.log('Проверка завершена. Экспорт PDF при проверке не выполнялся.')
            elif kind == 'complete':
                count, self.last_output = payload
                self.status.SetLabel(f'Готово! Дипломов: {count}')
                self.gauge.SetValue(100)
                self.log(f'Генерация завершена. Папка: {self.last_output}')
            elif kind == 'cancelled':
                self.status.SetLabel('Остановлено')
                self.log('Операция остановлена. Можно начать снова.')
            elif kind == 'error':
                self.status.SetLabel('Не удалось завершить')
                self.log(payload)
                if not self.closing:
                    wx.MessageBox(payload + '\n\nПодробности доступны в журнале.', 'Ошибка', wx.OK | wx.ICON_ERROR, self)
            elif kind == 'idle':
                self.set_busy(False)
                if self.closing:
                    self.timer.Stop()
                    self.Destroy()
                    return

    def open_output(self, event):
        path = self.last_output or self.paths['output_dir'].GetValue()
        try:
            if not path or not Path(path).is_dir():
                raise DiplomaError('Папка результата ещё не создана или недоступна.')
            os.startfile(str(Path(path).resolve()))
        except Exception as exc:
            self.report_error(exc)

    def open_log(self, event):
        try:
            os.startfile(str(app_data_dir() / 'diploma.log'))
        except Exception as exc:
            self.report_error(exc)

    def on_close(self, event):
        self.persist()
        if self.busy:
            self.closing = True
            self.stop_generation(None)
            if event.CanVeto():
                event.Veto()
                return
        self.timer.Stop()
        self.Destroy()


def main():
    multiprocessing.freeze_support()
    app = wx.App(False)
    try:
        handler = RotatingFileHandler(app_data_dir() / 'diploma.log', maxBytes=2_000_000,
                                      backupCount=3, encoding='utf-8')
        logging.basicConfig(level=logging.INFO, handlers=[handler],
                            format='%(asctime)s %(levelname)s %(message)s')
        frame = DiplomaGeneratorApp()
        frame.Show()
        app.MainLoop()
    except Exception as exc:
        logging.exception('Application startup failed')
        wx.MessageBox(f'Не удалось запустить приложение: {exc}', 'Ошибка запуска', wx.OK | wx.ICON_ERROR)


if __name__ == '__main__':
    main()

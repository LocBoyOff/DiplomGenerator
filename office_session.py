"""One isolated PowerPoint session per batch, with cancellable COM calls."""
import ctypes
import logging
import multiprocessing as mp
from pathlib import Path
import time
import traceback

from diploma_core import Cancelled, DiplomaError, PLACEHOLDER, _terminate_owned, fit_powerpoint_text, format_value

LOG = logging.getLogger(__name__)


def _frame(slide, target):
    shapes = slide.Shapes
    for index in target["path"][:-1]:
        shapes = shapes.Item(index).GroupItems
    shape = shapes.Item(target["path"][-1])
    if target["cell"]:
        row, column = target["cell"]
        shape = shape.Table.Cell(row, column).Shape
    return shape.TextFrame2


def _worker(connection, template_path, targets):
    import comtypes
    import comtypes.client
    import psutil
    app = deck = slide = None
    owned = False
    stage = "инициализация Office"
    comtypes.CoInitialize()
    try:
        existing = {p.pid for p in psutil.process_iter(["name"])
                    if (p.info["name"] or "").lower() == "powerpnt.exe"}
        app = comtypes.client.CreateObject("PowerPoint.Application")
        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(ctypes.c_void_p(int(app.HWND)), ctypes.byref(pid))
        owned = bool(pid.value and pid.value not in existing)
        if owned:
            connection.send(("owned", (pid.value, psutil.Process(pid.value).create_time())))
        stage = "открытие шаблона"
        deck = app.Presentations.Open(template_path, True, False, False)
        compiled = []
        for target in targets:
            text = _frame(deck.Slides.Item(1), target).TextRange.Text
            matches = [(len(text[:m.start()].encode("utf-16-le")) // 2 + 1,
                        len(m.group().encode("utf-16-le")) // 2, m.group(1))
                       for m in PLACEHOLDER.finditer(text)]
            if matches:
                compiled.append((target, matches))
        connection.send(("ready", None))
        while True:
            job = connection.recv()
            if job is None:
                break
            stage = "создание диплома из шаблона"
            # A fresh slide also resets fonts reduced for the previous participant.
            slide = deck.Slides.Item(1).Duplicate().Item(1)
            for target, matches in compiled:
                frame = _frame(slide, target)
                if job["auto_fit"] and not target["cell"]:
                    frame.AutoSize = 0
                    frame.WordWrap = 0
                text = frame.TextRange
                for start, length, placeholder in reversed(matches):
                    text.Characters(start, length).Text = format_value(job["values"].get(placeholder, ""))
            frame = text = None
            if job["auto_fit"]:
                stage = "уменьшение текста до размеров блоков"
                report = fit_powerpoint_text(deck, [t for t, _ in compiled], slide.SlideIndex)
                connection.send(("fitted", report))
            stage = "сохранение PDF"
            deck.PrintOptions.Ranges.ClearAll()
            page_range = deck.PrintOptions.Ranges.Add(slide.SlideIndex, slide.SlideIndex)
            deck.ExportAsFixedFormat(job["pdf"], 2, 2, PrintRange=page_range, RangeType=4)
            deck.PrintOptions.Ranges.ClearAll()
            page_range = None
            slide.Delete()
            slide = None
            connection.send(("success", None))
    except (EOFError, BrokenPipeError):
        pass
    except BaseException as exc:
        try:
            connection.send(("error", {"message": f"Этап: {stage}. {exc}",
                                        "traceback": traceback.format_exc()}))
        except (OSError, EOFError):
            pass
    finally:
        if deck is not None:
            try:
                deck.Saved = -1  # Only the private batch template is discarded.
                deck.Close()
            except Exception:
                LOG.exception("Could not close batch template")
        if app is not None and owned:
            try:
                import gc
                frame = text = page_range = slide = deck = None
                gc.collect()
                if app.Presentations.Count == 0:
                    app.Quit()
                else:
                    # A user may have opened another document during this batch.
                    connection.send(("released", None))
            except Exception:
                LOG.exception("Could not quit batch PowerPoint")
        slide = deck = app = None
        comtypes.CoUninitialize()
        connection.close()


class PowerPointSession:
    def __init__(self, template_path, targets, stop_event, timeout=90):
        self.stop_event, self.timeout = stop_event, timeout
        self.identity = None
        context = mp.get_context("spawn")
        self.connection, child = context.Pipe()
        self.process = context.Process(target=_worker,
                                       args=(child, str(Path(template_path).resolve()), targets))
        self.started = False
        try:
            if stop_event.is_set():
                raise Cancelled()
            self.process.start()
            self.started = True
            child.close()
            self._wait("ready")
        except BaseException:
            child.close()
            self.close()
            raise

    def _wait(self, expected, fit_report=None, started=None):
        started = time.monotonic() if started is None else started
        while True:
            if self.stop_event.is_set():
                raise Cancelled()
            if time.monotonic() - started > self.timeout:
                raise DiplomaError(f"PowerPoint не завершил операцию за {self.timeout} секунд. "
                                   "Проверьте диалоги Office и доступность шаблона.")
            try:
                available = self.connection.poll(.05)
            except OSError:
                raise DiplomaError("Связь с процессом PowerPoint прервана.") from None
            if available:
                try:
                    kind, payload = self.connection.recv()
                except EOFError:
                    raise DiplomaError("Процесс экспорта PowerPoint завершился неожиданно.") from None
                if kind == "owned":
                    self.identity = payload
                elif kind == "released":
                    self.identity = None
                elif kind == "fitted":
                    if fit_report:
                        fit_report(payload)
                elif kind == "error":
                    LOG.error("Office traceback:\n%s", payload["traceback"])
                    raise DiplomaError("Не удалось выполнить экспорт через PowerPoint.\n" + payload["message"])
                elif kind == expected:
                    return
            elif not self.process.is_alive():
                raise DiplomaError("Процесс экспорта PowerPoint завершился неожиданно.")

    def export(self, values, pdf, auto_fit=True, fit_report=None):
        self.start_export(values, pdf, auto_fit)
        return self.finish_export(fit_report)

    def start_export(self, values, pdf, auto_fit=True):
        if self.stop_event.is_set():
            raise Cancelled()
        self.pending_path = Path(pdf).resolve()
        self.pending_started = time.monotonic()
        try:
            self.connection.send({"values": values, "pdf": str(self.pending_path), "auto_fit": auto_fit})
        except OSError:
            raise DiplomaError("Связь с процессом PowerPoint прервана.") from None

    def finish_export(self, fit_report=None):
        self._wait("success", fit_report, self.pending_started)
        path = self.pending_path
        if not path.is_file():
            raise DiplomaError("PowerPoint не создал PDF.")
        with path.open("rb") as stream:
            if stream.read(5) != b"%PDF-":
                raise DiplomaError("PowerPoint не создал корректный PDF.")
        return True

    def close(self):
        if self.started:
            try:
                self.connection.send(None)
            except (OSError, BrokenPipeError):
                pass
            deadline = time.monotonic() + 2
            while self.process.is_alive() and time.monotonic() < deadline:
                try:
                    if self.connection.poll(.05):
                        kind, payload = self.connection.recv()
                        if kind == "owned":
                            self.identity = payload
                        elif kind == "released":
                            self.identity = None
                except (EOFError, OSError):
                    break
            self.process.join(timeout=.2)
            if self.process.is_alive():
                self.process.terminate()
                self.process.join(timeout=3)
            _terminate_owned(self.identity)
            self.process.close()
            self.started = False
        self.connection.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

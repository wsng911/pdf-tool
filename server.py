import io
import json
import math
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

import fitz  # PyMuPDF
import webview
from PIL import Image


A4_W = 595.275590551
A4_H = 841.88976378
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
def resource_base_dir():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent

class CancelledError(Exception):
    pass


def mm_to_pt(mm):
    return float(mm) * 72.0 / 25.4


def safe_int(value, default):
    try:
        return int(value)
    except Exception:
        return default


def safe_float(value, default):
    try:
        return float(value)
    except Exception:
        return default


def default_output_dir():
    home = Path.home()
    downloads = home / "Downloads"
    return downloads if downloads.exists() else home


def ensure_unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    i = 1
    while True:
        candidate = parent / f"{stem} ({i}){suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def open_with_system(path: str):
    if not path:
        return False
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        elif system == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return True
    except Exception:
        return False


def notify_system(title: str, message: str):
    try:
        system = platform.system()
        if system == "Darwin":
            subprocess.Popen([
                "osascript",
                "-e",
                f'display notification {json.dumps(message)} with title {json.dumps(title)}'
            ])
        elif system == "Linux" and shutil.which("notify-send"):
            subprocess.Popen(["notify-send", title, message])
        # Windows 这里不强依赖额外库，避免增加依赖
    except Exception:
        pass


def detect_ghostscript():
    candidates = []

    if platform.system() == "Windows":
        names = ["gswin64c.exe", "gswin32c.exe", "gs.exe"]
    else:
        names = ["gs"]

    for name in names:
        p = shutil.which(name)
        if p:
            candidates.append(Path(p))

    if platform.system() == "Windows":
        roots = [
            os.environ.get("ProgramFiles"),
            os.environ.get("ProgramFiles(x86)"),
        ]
        patterns = [
            "gs/*/bin/gswin64c.exe",
            "gs/*/bin/gswin32c.exe",
            "Ghostscript/*/bin/gswin64c.exe",
            "Ghostscript/*/bin/gswin32c.exe",
        ]
        for root in roots:
            if not root:
                continue
            root_path = Path(root)
            for pattern in patterns:
                for p in root_path.glob(pattern):
                    candidates.append(p)

    seen = set()
    uniq = []
    for p in candidates:
        sp = str(p)
        if sp not in seen and p.exists():
            seen.add(sp)
            uniq.append(p)

    for gs_path in uniq:
        try:
            result = subprocess.run(
                [str(gs_path), "-version"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=CREATE_NO_WINDOW,
            )
            version = (result.stdout or result.stderr).strip().splitlines()[0].strip()
            if result.returncode == 0 and version:
                return {
                    "ok": True,
                    "path": str(gs_path),
                    "version": version
                }
        except Exception:
            continue

    return {"ok": False, "path": "", "version": ""}


def parse_page_range(expr: str, total_pages: int):
    if total_pages <= 0:
        return []

    expr = (expr or "").strip()
    if not expr:
        return list(range(total_pages))

    result = []
    seen = set()

    parts = [p.strip() for p in expr.split(",") if p.strip()]
    if not parts:
        return list(range(total_pages))

    for part in parts:
        if "-" in part:
            a, b = part.split("-", 1)
            try:
                start = int(a.strip())
                end = int(b.strip())
            except Exception:
                raise ValueError(f"无效页码范围：{part}")

            if start > end:
                start, end = end, start

            start = max(1, start)
            end = min(total_pages, end)
            for n in range(start, end + 1):
                idx = n - 1
                if idx not in seen:
                    seen.add(idx)
                    result.append(idx)
        else:
            try:
                n = int(part)
            except Exception:
                raise ValueError(f"无效页码：{part}")

            if 1 <= n <= total_pages:
                idx = n - 1
                if idx not in seen:
                    seen.add(idx)
                    result.append(idx)

    if not result:
        raise ValueError("页码范围没有命中任何页面")

    return result


def compute_target_rect(src_w, src_h, margin_mm, auto_rotate=True):
    margin = mm_to_pt(margin_mm)
    box_w = max(10, A4_W - 2 * margin)
    box_h = max(10, A4_H - 2 * margin)

    rotate = 0
    w, h = float(src_w), float(src_h)

    if auto_rotate:
        if (w > h and box_w < box_h) or (w < h and box_w > box_h):
            rotate = 90
            w, h = h, w

    scale = min(box_w / w, box_h / h)
    new_w = w * scale
    new_h = h * scale

    x0 = (A4_W - new_w) / 2
    y0 = (A4_H - new_h) / 2
    return fitz.Rect(x0, y0, x0 + new_w, y0 + new_h), rotate


def pixmap_to_jpeg_bytes(pix, quality=80):
    mode = "L" if pix.n < 3 else "RGB"
    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
    bio = io.BytesIO()
    img.save(bio, format="JPEG", quality=quality, optimize=True)
    return bio.getvalue()


class ProgressTracker:
    def __init__(self, api, total_steps: int):
        self.api = api
        self.total_steps = max(1, total_steps)
        self.done_steps = 0
        self.lock = threading.Lock()

    def current_percent(self):
        with self.lock:
            return int(self.done_steps * 100 / self.total_steps)

    def step(self, status: str):
        with self.lock:
            self.done_steps += 1
            percent = int(self.done_steps * 100 / self.total_steps)
            percent = max(0, min(100, percent))
        self.api.emit("progress", {"percent": percent, "status": status})

    def status(self, status: str, percent=None):
        if percent is None:
            percent = self.current_percent()
        percent = max(0, min(100, int(percent)))
        self.api.emit("progress", {"percent": percent, "status": status})


class Api:
    def __init__(self):
        self.window = None
        self.worker = None
        self.cancel_event = threading.Event()
        self.worker_lock = threading.Lock()

    def set_window(self, window):
        self.window = window

    def emit(self, event, data=None):
        if not self.window:
            return
        try:
            payload = json.dumps({"event": event, "data": data}, ensure_ascii=False)
            self.window.evaluate_js(f"window.__pdfEvent({payload})")
        except Exception:
            pass

    # ── dialogs / basic ──────────────────────────────────────────────

    def get_default_output_dir(self):
        return str(default_output_dir())

    def pick_files(self):
        if not self.window:
            return []
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=("PDF Files (*.pdf)", "All files (*.*)")
        )
        return [str(Path(p)) for p in (result or [])]

    def pick_folder(self):
        if not self.window:
            return None
        result = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        if not result:
            return None
        if isinstance(result, (list, tuple)):
            return str(result[0]) if result else None
        return str(result)

    def pick_output_dir(self):
        return self.pick_folder()

    def scan_folder(self, folder):
        try:
            p = Path(folder)
            if not p.exists() or not p.is_dir():
                return []
            pdfs = [str(x) for x in p.rglob("*") if x.is_file() and x.suffix.lower() == ".pdf"]
            pdfs.sort(key=lambda s: s.lower())
            return pdfs
        except Exception:
            return []

    def get_pdf_info(self, files):
        total_pages = 0
        valid_files = 0
        for file in files or []:
            try:
                with fitz.open(str(file)) as doc:
                    total_pages += doc.page_count
                    valid_files += 1
            except Exception:
                continue
        return {"file_count": valid_files, "total_pages": total_pages}

    def check_ghostscript(self):
        return detect_ghostscript()

    def open_directory(self, path):
        return open_with_system(str(path))

    def quit(self):
        try:
            self.cancel_event.set()
            if self.window:
                self.window.destroy()
        except Exception:
            pass
        return {"ok": True}

    # ── process control ───────────────────────────────────────────────

    def cancel_process(self):
        self.cancel_event.set()
        return {"ok": True}

    def start_process(self, config):
        with self.worker_lock:
            if self.worker and self.worker.is_alive():
                return {"ok": False, "error": "已有任务正在运行"}

            files = (config or {}).get("files") or []
            if not files:
                return {"ok": False, "error": "请先添加 PDF 文件"}

            self.cancel_event.clear()
            self.worker = threading.Thread(
                target=self._worker_run,
                args=(config,),
                daemon=True
            )
            self.worker.start()
            return {"ok": True}

    def _worker_run(self, config):
        try:
            files = [str(Path(f)) for f in (config.get("files") or [])]
            page_range = (config.get("page_range") or "").strip()

            jobs = []
            for file in files:
                try:
                    src = Path(file)
                    if not src.exists():
                        raise FileNotFoundError("文件不存在")
                    with fitz.open(str(src)) as doc:
                        indices = parse_page_range(page_range, doc.page_count)
                    jobs.append({"file": str(src), "indices": indices})
                except Exception as e:
                    self.emit("file_error", {"file": file, "error": str(e)})

            if not jobs:
                raise RuntimeError("没有可处理的有效 PDF 文件")

            total_steps = sum(len(job["indices"]) for job in jobs)
            tracker = ProgressTracker(self, total_steps)
            output_paths = []

            gs_info = detect_ghostscript()
            if config.get("mode") == "gray_gs" and not gs_info.get("ok"):
                tracker.status("未找到 Ghostscript，自动改用稳定栅格灰度", percent=0)

            for i, job in enumerate(jobs, start=1):
                if self.cancel_event.is_set():
                    raise CancelledError()

                src_file = job["file"]
                indices = job["indices"]
                label = Path(src_file).name

                tracker.status(f"准备处理 {label} ({i}/{len(jobs)})")

                out_path = self._make_output_path(src_file, config)
                try:
                    self._process_one(src_file, indices, config, out_path, tracker, gs_info)
                    output_paths.append(out_path)
                except CancelledError:
                    raise
                except Exception as e:
                    self.emit("file_error", {"file": src_file, "error": str(e)})

            if self.cancel_event.is_set():
                raise CancelledError()

            if not output_paths:
                raise RuntimeError("没有成功输出任何文件")

            output_cfg = config.get("output") or {}

            if output_cfg.get("open_first_pdf") and output_paths:
                open_with_system(output_paths[0])

            if output_cfg.get("open_dir_after") and output_paths:
                open_with_system(str(Path(output_paths[0]).parent))

            if output_cfg.get("notify"):
                notify_system("PDF处理完成", f"已输出 {len(output_paths)} 个文件")

            self.emit("done", {"output_paths": output_paths})

        except CancelledError:
            self.emit("cancelled", None)
        except Exception as e:
            self.emit("error", str(e))
        finally:
            self.cancel_event.clear()
            with self.worker_lock:
                self.worker = None

    # ── output path ───────────────────────────────────────────────────

    def _make_output_path(self, src_file, config):
        src = Path(src_file)
        output_cfg = config.get("output") or {}

        use_source_dir = bool(output_cfg.get("use_source_dir"))
        out_dir = src.parent if use_source_dir else Path(output_cfg.get("dir") or default_output_dir())
        out_dir.mkdir(parents=True, exist_ok=True)

        suffix = str(output_cfg.get("suffix") or "")
        mode = config.get("mode") or "color"

        if mode in ("gray_gs", "gray_raster") and output_cfg.get("auto_suffix_gray"):
            if "gray" not in suffix.lower():
                suffix = suffix + "_gray"

        filename = f"{src.stem}{suffix}.pdf" if suffix else f"{src.stem}.pdf"
        return str(ensure_unique_path(out_dir / filename))

    # ── process implementations ───────────────────────────────────────

    def _process_one(self, src_file, page_indices, config, out_path, tracker, gs_info):
        mode = config.get("mode") or "color"
        auto_rotate = bool(config.get("auto_rotate_a4", True))
        margin_mm = safe_float(config.get("margin_mm"), 8)
        raster_dpi = max(72, safe_int(config.get("raster_dpi"), 200))

        compress_cfg = config.get("compress") or {}
        compress_enabled = bool(compress_cfg.get("enabled"))

        if compress_enabled:
            self._build_compressed_pdf(
                src_file=src_file,
                page_indices=page_indices,
                out_path=out_path,
                tracker=tracker,
                mode=mode,
                auto_rotate=auto_rotate,
                margin_mm=margin_mm,
                compress_cfg=compress_cfg,
            )
            return

        if mode == "color":
            self._build_vector_pdf(
                src_file=src_file,
                page_indices=page_indices,
                out_path=out_path,
                tracker=tracker,
                auto_rotate=auto_rotate,
                margin_mm=margin_mm,
            )
        elif mode == "gray_gs":
            if gs_info.get("ok"):
                with tempfile.TemporaryDirectory() as td:
                    temp_pdf = str(Path(td) / "vector_base.pdf")
                    self._build_vector_pdf(
                        src_file=src_file,
                        page_indices=page_indices,
                        out_path=temp_pdf,
                        tracker=tracker,
                        auto_rotate=auto_rotate,
                        margin_mm=margin_mm,
                    )
                    if self.cancel_event.is_set():
                        raise CancelledError()

                    tracker.status(
                        f"Ghostscript 转灰度：{Path(src_file).name}",
                        percent=min(tracker.current_percent(), 99)
                    )
                    self._ghostscript_gray(temp_pdf, out_path, gs_info["path"])
            else:
                self._build_raster_pdf(
                    src_file=src_file,
                    page_indices=page_indices,
                    out_path=out_path,
                    tracker=tracker,
                    grayscale=True,
                    dpi=raster_dpi,
                    jpeg_quality=85,
                    auto_rotate=auto_rotate,
                    margin_mm=margin_mm,
                )
        elif mode == "gray_raster":
            self._build_raster_pdf(
                src_file=src_file,
                page_indices=page_indices,
                out_path=out_path,
                tracker=tracker,
                grayscale=True,
                dpi=raster_dpi,
                jpeg_quality=85,
                auto_rotate=auto_rotate,
                margin_mm=margin_mm,
            )
        else:
            raise ValueError(f"未知输出模式：{mode}")

    def _build_vector_pdf(self, src_file, page_indices, out_path, tracker, auto_rotate, margin_mm):
        label = Path(src_file).name
        out_doc = fitz.open()

        try:
            with fitz.open(src_file) as src_doc:
                total = len(page_indices)

                for i, idx in enumerate(page_indices, start=1):
                    if self.cancel_event.is_set():
                        raise CancelledError()

                    src_page = src_doc.load_page(idx)
                    rect, rotate = compute_target_rect(
                        src_page.rect.width,
                        src_page.rect.height,
                        margin_mm,
                        auto_rotate
                    )

                    page = out_doc.new_page(width=A4_W, height=A4_H)
                    page.show_pdf_page(rect, src_doc, idx, rotate=rotate)

                    if tracker:
                        tracker.step(f"正在处理 {label} - 第 {i}/{total} 页")

            out_doc.save(out_path, garbage=4, deflate=True)
        finally:
            out_doc.close()

    def _build_raster_pdf(
        self,
        src_file,
        page_indices,
        out_path,
        tracker,
        grayscale,
        dpi,
        jpeg_quality,
        auto_rotate,
        margin_mm,
    ):
        label = Path(src_file).name
        out_doc = fitz.open()

        try:
            with fitz.open(src_file) as src_doc:
                total = len(page_indices)

                for i, idx in enumerate(page_indices, start=1):
                    if self.cancel_event.is_set():
                        raise CancelledError()

                    src_page = src_doc.load_page(idx)
                    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
                    colorspace = fitz.csGRAY if grayscale else fitz.csRGB
                    pix = src_page.get_pixmap(matrix=matrix, colorspace=colorspace, alpha=False)
                    img_bytes = pixmap_to_jpeg_bytes(pix, quality=jpeg_quality)

                    rect, rotate = compute_target_rect(
                        src_page.rect.width,
                        src_page.rect.height,
                        margin_mm,
                        auto_rotate
                    )

                    page = out_doc.new_page(width=A4_W, height=A4_H)
                    page.insert_image(rect, stream=img_bytes, rotate=rotate)

                    if tracker:
                        tracker.step(f"正在处理 {label} - 第 {i}/{total} 页")

            out_doc.save(out_path, garbage=4, deflate=True)
        finally:
            out_doc.close()

    def _build_compressed_pdf(self, src_file, page_indices, out_path, tracker, mode, auto_rotate, margin_mm, compress_cfg):
        grayscale = mode != "color"
        base_dpi = max(72, safe_int(compress_cfg.get("max_dpi"), 160))
        base_quality = min(95, max(25, safe_int(compress_cfg.get("jpeg_quality"), 72)))
        target_mb = max(0.1, safe_float(compress_cfg.get("target_mb"), 10))
        label = Path(src_file).name

        with tempfile.TemporaryDirectory() as td:
            pass1 = str(Path(td) / "pass1.pdf")
            self._build_raster_pdf(
                src_file=src_file,
                page_indices=page_indices,
                out_path=pass1,
                tracker=tracker,
                grayscale=grayscale,
                dpi=base_dpi,
                jpeg_quality=base_quality,
                auto_rotate=auto_rotate,
                margin_mm=margin_mm,
            )

            chosen = Path(pass1)
            size1 = chosen.stat().st_size
            target_bytes = int(target_mb * 1024 * 1024)

            if size1 > target_bytes and not self.cancel_event.is_set():
                ratio = math.sqrt(target_bytes / max(1, size1))
                ratio = max(0.35, min(0.9, ratio))

                dpi2 = max(72, int(base_dpi * ratio))
                q2 = max(35, int(base_quality * (0.75 + 0.25 * ratio)))

                pass2 = str(Path(td) / "pass2.pdf")
                tracker.status(f"压缩优化中：{label}", percent=min(tracker.current_percent(), 99))

                self._build_raster_pdf(
                    src_file=src_file,
                    page_indices=page_indices,
                    out_path=pass2,
                    tracker=None,
                    grayscale=grayscale,
                    dpi=dpi2,
                    jpeg_quality=q2,
                    auto_rotate=auto_rotate,
                    margin_mm=margin_mm,
                )

                size2 = Path(pass2).stat().st_size

                if size2 <= target_bytes and size1 > target_bytes:
                    chosen = Path(pass2)
                elif size1 <= target_bytes and size2 > target_bytes:
                    chosen = Path(pass1)
                else:
                    # 都没到目标 or 都到目标，选更接近目标的
                    if abs(size2 - target_bytes) < abs(size1 - target_bytes):
                        chosen = Path(pass2)

            shutil.copyfile(str(chosen), out_path)

    def _ghostscript_gray(self, in_pdf, out_pdf, gs_path):
        if self.cancel_event.is_set():
            raise CancelledError()

        cmd = [
            gs_path,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            "-dNOPAUSE",
            "-dBATCH",
            "-dSAFER",
            "-dAutoRotatePages=/None",
            "-sColorConversionStrategy=Gray",
            "-dProcessColorModel=/DeviceGray",
            "-dOverrideICC",
            "-dDetectDuplicateImages=true",
            "-dCompressFonts=true",
            f"-sOutputFile={out_pdf}",
            in_pdf,
        ]

        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=CREATE_NO_WINDOW,
        )

        if result.returncode != 0:
            msg = (result.stderr or result.stdout or "Ghostscript 执行失败").strip()
            raise RuntimeError(msg)

        if not Path(out_pdf).exists():
            raise RuntimeError("Ghostscript 未生成输出文件")


def main():
    base_dir = resource_base_dir()
    html_path = base_dir / "index.html"

    api = Api()
    window = webview.create_window(
        title="PDF处理工具",
        url=str(html_path),
        js_api=api,
        width=920,
        height=860,
        min_size=(860, 760),
    )
    api.set_window(window)

    webview.start(debug=False)


if __name__ == "__main__":
    main()

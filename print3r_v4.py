"""
requirements.txt

pillow
requests
pywin32
# plus your existing thermalprinter dependency
"""

import os
import io
import sys
import json
import time
import random
import textwrap
import tempfile
import threading
import datetime as dt
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any

# GUI
import tkinter as tk
from tkinter import filedialog, messagebox

# Imaging
from PIL import Image, ImageEnhance, ImageTk, ImageGrab

# Printer + Serial
from thermalprinter import ThermalPrinter
import serial.tools.list_ports

# Networking
import requests

# Outlook (optional, guarded)
try:
    import win32com.client as win32
    from win32com.client import constants, gencache
    import pythoncom
    HAVE_WIN32 = True
except Exception:
    HAVE_WIN32 = False
    win32 = None
    constants = None
    pythoncom = None

# Timezone (Python 3.9+)
try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None  # best-effort fallback

# =========================
# ===== USER CONFIG =======
# =========================
WIDTH_PIXELS = 384                   # keep from current app
LINE_WIDTH_CHARS = 32                # monospaced characters per line (tweak for your unit)
TIMEZONE_NAME = "Europe/Zurich"
CITY_NAME = "Zurich"
USE_ONLINE_SOURCES = True            # if False, rely on cached/fallbacks

# Wall-mount mode: rotate all output 180° and print bottom-up
UPSIDE_DOWN_MODE = True

# Meme sources (NSFW allowed)
MEME_SUBREDDITS = ["wholesomememes", "ProgrammerHumor"]
ALLOW_NSFW = True

# Networking
HTTP_TIMEOUT = 6  # seconds

# Printing
HEAT_TIME = 110  # keep from current app

# Pillow 10+ compatible resampling
RESAMPLE_FILTER = getattr(Image, 'Resampling', Image).LANCZOS

# Cache dir
CACHE_DIR = os.path.join(tempfile.gettempdir(), "thermal_print3r_cache")
os.makedirs(CACHE_DIR, exist_ok=True)

# Weather constants
ZURICH_LAT = 47.3769
ZURICH_LON = 8.5417

# Weather code map → (ascii label, description) for display
WEATHER_ICON_MAP = {
    0: ("SUN", "Clear"),
    1: ("SUN", "Mainly clear"),
    2: ("CLOUDS", "Partly cloudy"),
    3: ("CLOUDS", "Overcast"),
    45: ("FOG", "Fog"),
    48: ("FOG", "Depositing rime fog"),
    51: ("RAIN", "Light drizzle"),
    53: ("RAIN", "Drizzle"),
    55: ("RAIN", "Dense drizzle"),
    56: ("RAIN", "Freezing drizzle"),
    57: ("RAIN", "Freezing drizzle"),
    61: ("RAIN", "Slight rain"),
    63: ("RAIN", "Rain"),
    65: ("RAIN", "Heavy rain"),
    66: ("RAIN", "Freezing rain"),
    67: ("RAIN", "Freezing rain"),
    71: ("SNOW", "Slight snow"),
    73: ("SNOW", "Snow"),
    75: ("SNOW", "Heavy snow"),
    77: ("SNOW", "Snow grains"),
    80: ("RAIN", "Rain showers"),
    81: ("RAIN", "Rain showers"),
    82: ("RAIN", "Violent rain showers"),
    85: ("SNOW", "Snow showers"),
    86: ("SNOW", "Snow showers"),
    95: ("TSTMS", "Thunderstorm"),
    96: ("TSTMS", "Thunderstorm hail"),
    99: ("TSTMS", "Thunderstorm hail"),
}

# --- Printer-safe text helpers (remove emojis etc.) ---
CP437_SAFE_REPLACEMENTS = {
    "—": "-", "–": "-", "“": '"', "”": '"', "‘": "'", "’": "'",
    # Common emoji/labels → ASCII
    "☀️": "SUN", "🌤️": "SUN", "⛅": "CLOUDS", "☁️": "CLOUDS",
    "🌫️": "FOG", "🌦️": "SHWRS", "🌧️": "RAIN",
    "❄️": "SNOW", "⛈️": "TSTMS",
    "📅": "CAL", "🤣": "JOKE",
}

def sanitize_for_printer(s: str) -> str:
    """Replace unsupported chars with rough ASCII equivalents; ensure CP437-safe."""
    for k, v in CP437_SAFE_REPLACEMENTS.items():
        s = s.replace(k, v)
    try:
        s.encode("cp437", errors="strict")
        return s  # already safe
    except Exception:
        return s.encode("cp437", errors="replace").decode("cp437")


def get_tz() -> dt.tzinfo:
    """Return timezone object; fall back to local if zoneinfo missing."""
    if ZoneInfo:
        try:
            return ZoneInfo(TIMEZONE_NAME)
        except Exception:
            pass
    return dt.datetime.now().astimezone().tzinfo


def wrap(text: str, width: int) -> List[str]:
    """Simple hard wrap for monospaced printer."""
    lines: List[str] = []
    for para in text.splitlines():
        if not para:
            lines.append("")
            continue
        lines.extend(textwrap.wrap(para, width=width, replace_whitespace=False, drop_whitespace=False))
    return lines


# Wind bearing → cardinal
CARDINALS = ["N","NNE","NE","ENE","E","ESE","SE","SSE","S","SSW","SW","WSW","W","WNW","NW","NNW"]
def bearing_to_cardinal(deg: Optional[float]) -> str:
    try:
        d = (float(deg) % 360.0)
        idx = int((d + 11.25) // 22.5) % 16
        return CARDINALS[idx]
    except Exception:
        return "?"


# Weather asset filenames (monochrome PNGs you provide)
WEATHER_ICON_FILES = {
    "clear": "sun.png",
    "cloudy": "cloud.png",
    "rain": "rain.png",
    "snow": "snow.png",
    "thunder": "thunder.png",
}
def weather_category_from_code(code: int) -> str:
    code = int(code)
    if code in (0, 1): return "clear"
    if code in (2, 3, 45, 48): return "cloudy"
    if code in (51, 53, 55, 56, 57, 61, 63, 65, 80, 81, 82): return "rain"
    if code in (71, 73, 75, 77, 85, 86): return "snow"
    if code in (95, 96, 99): return "thunder"
    return "cloudy"


def parse_openmeteo_local(t: str, tz: dt.tzinfo) -> dt.datetime:
    """Open-Meteo returns local wall time without offset when timezone param is used."""
    try:
        d = dt.datetime.fromisoformat(t)
    except Exception:
        d = dt.datetime.strptime(t, "%Y-%m-%dT%H:%M")
    return d.replace(tzinfo=tz)


def get_asset_path(*parts) -> Optional[str]:
    """Return the first existing path among script dir / CWD / PyInstaller temp."""
    candidates: List[Path] = []
    try:
        candidates.append(Path(__file__).resolve().parent.joinpath(*parts))
    except Exception:
        pass
    candidates.append(Path.cwd().joinpath(*parts))
    if hasattr(sys, "_MEIPASS"):
        try:
            candidates.append(Path(getattr(sys, "_MEIPASS")).joinpath(*parts))  # type: ignore
        except Exception:
            pass
    for p in candidates:
        if p.exists():
            return str(p)
    return None


def ensure_filename(img: Image.Image, name: str = "generated.png") -> Image.Image:
    """Attach a filename attribute to PIL images created in-memory."""
    if not hasattr(img, "filename"):
        try:
            img.filename = name
        except Exception:
            pass
    return img


def trim_white_borders(img: Image.Image, thresh: int = 250) -> Image.Image:
    """Crop away white/near-white borders to avoid extra vertical whitespace on printed icons."""
    gray = img.convert("L")
    mask = gray.point(lambda p: 255 if p < thresh else 0)
    bbox = mask.getbbox()
    if bbox:
        return img.crop(bbox)
    return img


# =========================
# Helper Classes
# =========================

class PrinterHelper:
    """Wrapper around ThermalPrinter with simple text/image utilities and one retry."""

    def __init__(self, port_getter, heat_time: int = HEAT_TIME, line_width: int = LINE_WIDTH_CHARS):
        self.port_getter = port_getter
        self.heat_time = heat_time
        self.line_width = line_width

    # ------- Core open/retry -------
    def _open(self) -> ThermalPrinter:
        port = self.port_getter()
        if not port:
            raise RuntimeError("Printer not connected")
        return ThermalPrinter(port=port, heat_time=self.heat_time)

    def _with_retry(self, fn, *args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except Exception:
            time.sleep(0.3)
            return fn(*args, **kwargs)

    # ------- Text as image (for upside-down + stability) -------
    def _render_text_ticket(self, lines: List[str], padding: int = 6, line_gap: int = 2) -> Image.Image:
        """Render list of lines into a WIDTH_PIXELS-wide 1-bit image."""
        from PIL import ImageDraw, ImageFont

        # Choose a monospaced font if available; fallback to default bitmap font.
        font = None
        candidate_paths = [
            r"C:\Windows\Fonts\consola.ttf",    # Consolas
            r"C:\Windows\Fonts\cour.ttf",       # Courier New
            r"C:\Windows\Fonts\lucon.ttf",      # Lucida Console
        ]
        for p in candidate_paths:
            if os.path.exists(p):
                try:
                    font = ImageFont.truetype(p, 18)
                    break
                except Exception:
                    font = None
        if font is None:
            # Bitmap fallback (monospace-ish)
            try:
                from PIL import ImageFont as _IF
                font = _IF.load_default()
            except Exception:
                pass

        # Measure line height
        # Use a temporary draw to measure text bbox robustly across Pillow versions
        tmp_img = Image.new("L", (10, 10), 255)
        draw = ImageDraw.Draw(tmp_img)
        bbox = draw.textbbox((0, 0), "Ag", font=font)
        line_h = (bbox[3] - bbox[1]) if bbox else 12
        line_h = max(10, line_h)

        height = padding * 2 + len(lines) * (line_h + line_gap) - line_gap
        height = max(1, height)

        canvas = Image.new("L", (WIDTH_PIXELS, height), 255)
        d = ImageDraw.Draw(canvas)

        y = padding
        for ln in lines:
            # Lines already wrapped and often pre-centered with spaces.
            d.text((padding, y), ln, fill=0, font=font)
            y += line_h + line_gap

        # Convert to 1-bit (sharp)
        ticket = canvas.convert("1")
        return ensure_filename(ticket, "text_ticket.png")

    def print_text(self, text: str, cancel_flag_getter=None, feed_lines: int = 5):
        """Kept for API compatibility: now renders text to an image and prints upside-down."""
        # Wrap and sanitize similarly to legacy behavior
        lines = []
        for line in text.splitlines():
            if line.strip() == "<SEPARATOR>":
                lines.append("-" * self.line_width)
            else:
                lines.extend(wrap(line, self.line_width))
        self.print_lines(lines, cancel_flag_getter=cancel_flag_getter, feed_lines=feed_lines)

    def print_lines(self, lines: List[str], cancel_flag_getter=None, feed_lines: int = 5):
        """Render lines to image and print band-wise (rotated if UPSIDE_DOWN_MODE)."""
        ticket = self._render_text_ticket(lines)
        self.print_image_bandwise(ticket, feed_lines=feed_lines, cancel_flag_getter=cancel_flag_getter)

    # ------- Image printing -------
    def _prep_image_for_output(self, img: Image.Image) -> Image.Image:
        """Resize to width, convert to 1-bit, rotate if needed."""
        # Ensure width
        if img.width != WIDTH_PIXELS:
            factor = WIDTH_PIXELS / img.width
            new_size = (WIDTH_PIXELS, max(1, int(round(img.height * factor))))
            img = img.resize(new_size, RESAMPLE_FILTER)
        # Convert to 1-bit if not already
        if img.mode != "1":
            img = img.convert("1", dither=Image.FLOYDSTEINBERG)
        # Rotate 180° for wall-mount mode
        if UPSIDE_DOWN_MODE:
            img = img.rotate(180, expand=True)
        return ensure_filename(img, getattr(img, "filename", "buffer.png"))

    def print_image(self, img: Image.Image, feed_lines: int = 10):
        """Simple image print (rotated if needed)."""
        img = self._prep_image_for_output(img)
        with self._open() as printer:
            self._with_retry(printer.image, img)
            self.safe_feed(printer, feed_lines)

    def print_image_bandwise(
        self,
        img: Image.Image,
        band_height: int = 64,
        base_sleep: float = 0.025,
        dark_bonus: float = 0.20,
        feed_lines: int = 8,
        cancel_flag_getter=None,
    ):
        """
        Print a 1-bit image in vertical bands with adaptive delays to avoid overrunning
        the printer when lots of black/dark pixels need more heat time.

        - band_height: rows per chunk (multiples of 8 are safe for most printers)
        - base_sleep: minimum pause after each band (seconds)
        - dark_bonus: extra pause scaled by band 'ink' coverage (0..1)
        """
        img = self._prep_image_for_output(img)
        H = img.height
        with self._open() as printer:
            y = 0
            while y < H:
                if cancel_flag_getter and cancel_flag_getter():
                    break
                band = img.crop((0, y, img.width, min(H, y + band_height)))
                # Send this band
                self._with_retry(printer.image, ensure_filename(band, f"band_{y}.png"))
                # Darkness heuristic to slow down on heavy coverage
                g = band.convert("L")
                hist = g.histogram()  # 256 bins
                black_pixels = hist[0] if hist else 0
                total = band.width * band.height if band.width and band.height else 1
                coverage = min(1.0, max(0.0, black_pixels / float(total)))
                time.sleep(base_sleep + coverage * dark_bonus)
                y += band_height
            self.safe_feed(printer, feed_lines)

    # ------- Misc -------
    def center(self, s: str) -> str:
        s = sanitize_for_printer(s.strip())
        return s.center(self.line_width)

    def separator(self) -> str:
        return "-" * self.line_width

    @staticmethod
    def safe_feed(printer: ThermalPrinter, n: int = 5):
        for _ in range(max(0, n)):
            try:
                printer.out("")
            except Exception:
                pass

    @staticmethod
    def list_com_ports() -> List[object]:
        return list(serial.tools.list_ports.comports())

    @staticmethod
    def guess_printer_port() -> Optional[str]:
        for p in serial.tools.list_ports.comports():
            try:
                with ThermalPrinter(port=p.device, heat_time=HEAT_TIME):
                    return p.device
            except Exception:
                continue
        return None


class CalendarFetcher:
    """Fetch today's events from Outlook Desktop using COM, default profile."""

    def __init__(self, tz: dt.tzinfo):
        self.tz = tz

    def get_today_events(self) -> List[Dict[str, str]]:
        if not HAVE_WIN32:
            raise RuntimeError("pywin32 not installed; Outlook COM unavailable.")

        co_inited = False
        try:
            if pythoncom is not None:
                pythoncom.CoInitialize()
                co_inited = True

            outlook = gencache.EnsureDispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            cal = ns.GetDefaultFolder(constants.olFolderCalendar)

            today = dt.datetime.now(self.tz).date()
            start = dt.datetime.combine(today, dt.time.min).astimezone(self.tz)
            end = dt.datetime.combine(today, dt.time.max).astimezone(self.tz)

            start_str = start.strftime("%m/%d/%Y %H:%M %p")
            end_str = end.strftime("%m/%d/%Y %H:%M %p")

            items = cal.Items
            items.IncludeRecurrences = True
            items.Sort("[Start]")
            restriction = f"[Start] >= '{start_str}' AND [End] <= '{end_str}'"
            try:
                ritems = items.Restrict(restriction)
            except Exception:
                ritems = items

            events: List[Dict[str, str]] = []
            try:
                for appt in ritems:
                    try:
                        st_str = dt.datetime.strptime(str(appt.Start), "%m/%d/%y %H:%M:%S").strftime("%H:%M")
                        en_str = dt.datetime.strptime(str(appt.End), "%m/%d/%y %H:%M:%S").strftime("%H:%M")
                    except Exception:
                        st_str = str(appt.Start)[11:16]
                        en_str = str(appt.End)[11:16]
                    subject = str(getattr(appt, "Subject", "") or "").strip()
                    location = str(getattr(appt, "Location", "") or "").strip()
                    events.append({
                        "start": st_str,
                        "end": en_str,
                        "subject": subject,
                        "location": location
                    })
            except Exception:
                pass

            return events
        except Exception as e:
            raise RuntimeError(f"Unable to access Outlook: {e}")
        finally:
            try:
                if co_inited:
                    pythoncom.CoUninitialize()
            except Exception:
                pass


class WeatherFetcher:
    """Open-Meteo client and formatting helpers."""

    def __init__(self, lat: float, lon: float, tzname: str):
        self.lat = lat
        self.lon = lon
        self.tzname = tzname

    def _cache_path(self) -> str:
        return os.path.join(CACHE_DIR, "last_weather.json")

    def get_current_and_hourly(self) -> Dict:
        base = "https://api.open-meteo.com/v1/forecast"
        params = {
            "latitude": self.lat,
            "longitude": self.lon,
            "hourly": "temperature_2m,precipitation_probability,weathercode,windspeed_10m,winddirection_10m",
            "current_weather": "true",
            "timezone": self.tzname,
            "daily": "temperature_2m_max,temperature_2m_min,weathercode",
        }
        if not USE_ONLINE_SOURCES:
            raise RuntimeError("Online sources disabled by config.")
        try:
            r = requests.get(base, params=params, timeout=HTTP_TIMEOUT)
            r.raise_for_status()
            data = r.json()
            with open(self._cache_path(), "w", encoding="utf-8") as f:
                json.dump(data, f)
            return data
        except Exception as e:
            try:
                with open(self._cache_path(), "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                raise RuntimeError(f"Weather fetch failed: {e}")

    @staticmethod
    def icon_label_from_code(code: int) -> Tuple[str, str]:
        return WEATHER_ICON_MAP.get(int(code), ("CLOUDS", "Cloudy"))


class FunFetcher:
    """Memes and jokes with online sources, fallbacks, caching."""

    UA = {"User-Agent": "ThermalPrint3r/1.0 (+https://example.local)"}

    def __init__(self):
        self.meme_cache_path = os.path.join(CACHE_DIR, "last_meme.png")
        self.joke_cache_path = os.path.join(CACHE_DIR, "last_joke.txt")

    # ===== Memes =====
    @staticmethod
    def _is_image_url(url: str) -> bool:
        ext = url.lower().split("?")[0].rsplit(".", 1)[-1]
        return ext in ("png", "jpg", "jpeg", "bmp", "gif", "webp")

    def _download_image(self, url: str) -> Optional[Image.Image]:
        try:
            r = requests.get(url, headers=self.UA, timeout=HTTP_TIMEOUT)
            r.raise_for_status()
            return Image.open(io.BytesIO(r.content)).convert("L")
        except Exception:
            return None

    def get_random_meme(self) -> Optional[Image.Image]:
        if USE_ONLINE_SOURCES:
            # Meme API first
            try:
                subs = ",".join(MEME_SUBREDDITS) if MEME_SUBREDDITS else ""
                url = "https://meme-api.com/gimme" + (f"/{subs}" if subs else "")
                r = requests.get(url, headers=self.UA, timeout=HTTP_TIMEOUT)
                r.raise_for_status()
                j = r.json()
                nsfw = bool(j.get("nsfw", False))
                if ALLOW_NSFW or not nsfw:
                    img_url = j.get("url") or ""
                    if self._is_image_url(img_url):
                        img = self._download_image(img_url)
                        if img:
                            try: img.save(self.meme_cache_path)
                            except Exception: pass
                            return img
            except Exception:
                pass
            # Reddit hot fallback
            try:
                sub = random.choice(MEME_SUBREDDITS) if MEME_SUBREDDITS else "wholesomememes"
                url = f"https://www.reddit.com/r/{sub}/hot.json?limit=50&raw_json=1"
                r = requests.get(url, headers=self.UA, timeout=HTTP_TIMEOUT)
                r.raise_for_status()
                posts = r.json().get("data", {}).get("children", [])
                random.shuffle(posts)
                for ch in posts:
                    d = ch.get("data", {})
                    if d.get("is_video"): continue
                    if (not ALLOW_NSFW) and d.get("over_18"): continue
                    img_url = d.get("url_overridden_by_dest") or d.get("url") or ""
                    if not self._is_image_url(img_url): continue
                    img = self._download_image(img_url)
                    if img:
                        try: img.save(self.meme_cache_path)
                        except Exception: pass
                        return img
            except Exception:
                pass
            # XKCD fallback
            try:
                latest = requests.get("https://xkcd.com/info.0.json", headers=self.UA, timeout=HTTP_TIMEOUT).json()
                maxn = int(latest.get("num"))
                for _ in range(5):
                    n = random.randint(1, maxn)
                    j = requests.get(f"https://xkcd.com/{n}/info.0.json", headers=self.UA, timeout=HTTP_TIMEOUT).json()
                    img_url = j.get("img", "")
                    if img_url:
                        img = self._download_image(img_url)
                        if img:
                            try: img.save(self.meme_cache_path)
                            except Exception: pass
                            return img
            except Exception:
                pass
        # Offline cache/local
        try:
            if os.path.exists(self.meme_cache_path):
                return Image.open(self.meme_cache_path).convert("L")
        except Exception:
            pass
        assets_dir = get_asset_path("assets", "memes")
        if assets_dir and os.path.isdir(assets_dir):
            candidates = [f for f in os.listdir(assets_dir) if f.lower().endswith((".png", ".jpg", ".jpeg", ".bmp"))]
            if candidates:
                try:
                    p = os.path.join(assets_dir, random.choice(candidates))
                    return Image.open(p).convert("L")
                except Exception:
                    pass
        return None

    # ===== Jokes =====
    def get_random_joke(self) -> str:
        if USE_ONLINE_SOURCES:
            # icanhazdadjoke
            try:
                headers = {**self.UA, "Accept": "application/json"}
                r = requests.get("https://icanhazdadjoke.com/", headers=headers, timeout=HTTP_TIMEOUT)
                r.raise_for_status()
                j = r.json()
                joke = (j.get("joke") or "").strip()
                if joke:
                    try:
                        with open(self.joke_cache_path, "w", encoding="utf-8") as f:
                            f.write(joke)
                    except Exception:
                        pass
                    return joke
            except Exception:
                pass
            # JokeAPI v2
            try:
                flags = [] if ALLOW_NSFW else ["nsfw"]
                url = "https://v2.jokeapi.dev/joke/Any?type=single" + (f"&blacklistFlags={','.join(flags)}" if flags else "")
                j = requests.get(url, headers=self.UA, timeout=HTTP_TIMEOUT).json()
                if j.get("type") == "single":
                    joke = (j.get("joke") or "").strip()
                    if joke:
                        try:
                            with open(self.joke_cache_path, "w", encoding="utf-8") as f:
                                f.write(joke)
                        except Exception:
                            pass
                        return joke
            except Exception:
                pass
            # Official Joke API
            try:
                j = requests.get("https://official-joke-api.appspot.com/jokes/random", headers=self.UA, timeout=HTTP_TIMEOUT).json()
                setup = (j.get("setup") or "").strip()
                punch = (j.get("punchline") or "").strip()
                joke = f"{setup} — {punch}".strip(" —")
                if joke:
                    try:
                        with open(self.joke_cache_path, "w", encoding="utf-8") as f:
                            f.write(joke)
                    except Exception:
                        pass
                    return joke
            except Exception:
                pass
        # Cache / fallback
        try:
            if os.path.exists(self.joke_cache_path):
                with open(self.joke_cache_path, "r", encoding="utf-8") as f:
                    joke = f.read().strip()
                    if joke:
                        return joke
        except Exception:
            pass
        fallback = [
            "I told my computer I needed a break, and it said 'No problem - I'll go to sleep.'",
            "Why do programmers prefer dark mode? Because light attracts bugs.",
            "I would tell you a UDP joke, but you might not get it.",
            "There are 10 kinds of people: those who understand binary and those who don't.",
        ]
        return random.choice(fallback)


# =========================
# Main App
# =========================

class ThermalPrintTool:
    def __init__(self):
        # Image state for printing (no main preview now)
        self.source_image: Optional[Image.Image] = None
        self.display_image: Optional[Image.Image] = None

        self.print_cancel_flag = False
        self.printer_port: Optional[str] = None

        # Helpers
        self.printer_helper = PrinterHelper(lambda: self.printer_port)
        self.calendar_fetcher = CalendarFetcher(get_tz())
        self.weather_fetcher = WeatherFetcher(ZURICH_LAT, ZURICH_LON, TIMEZONE_NAME)
        self.fun_fetcher = FunFetcher()

        # Main window (minimal controls)
        self.window = tk.Tk()
        self.window.title("Thermal Print3r Tool — Desk Buddy")
        try:
            self.window.iconbitmap('icon.ico')
        except Exception:
            pass
        self.window.configure(background="#2C3E50")

        # Menu
        self._build_menu()

        # Top: status + refresh
        top = tk.Frame(self.window, bg="#2C3E50")
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=6)
        self.lbl_printer_status = tk.Label(top, text="Printer: Detecting...", fg="#ECF0F1", bg="#2C3E50", font=("Arial", 12))
        self.lbl_printer_status.pack(side=tk.LEFT)
        tk.Button(top, text="Refresh Printer", command=self.refresh_printer,
                  bg="#1ABC9C", fg="#FFFFFF", font=("Arial", 10)).pack(side=tk.RIGHT)

        # Left column with ONLY the requested buttons
        left = tk.Frame(self.window, bg="#34495E", relief=tk.RIDGE, borderwidth=2)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        def bigbtn(text, cmd, color="#1ABC9C"):
            b = tk.Button(left, text=text, command=cmd, bg=color, fg="#FFFFFF", font=("Arial", 12, "bold"))
            b.pack(pady=6, fill=tk.X)
            return b

        bigbtn("Print Image", self.open_image_print_dialog)
        bigbtn("Print Text", self.open_text_print_dialog)
        bigbtn("Print Today’s Calendar", self._threaded_print_calendar, "#9B59B6")
        bigbtn("Print Weather", self._threaded_print_weather, "#3498DB")
        bigbtn("Print Meme", self._threaded_print_meme, "#E67E22")
        bigbtn("Print Joke", self._threaded_print_joke, "#E67E22")
        bigbtn("Cancel Print", self.cancel_print, "#95A5A6")

        # Status line (bottom of left)
        self.lbl_status = tk.Label(left, text="Last print: —", fg="#BDC3C7", bg="#34495E", anchor="w", font=("Consolas", 9))
        self.lbl_status.pack(side=tk.BOTTOM, fill=tk.X, padx=6, pady=6)

        # Start printer thread for image prints
        self.print_event = threading.Event()
        self.printer_thread = threading.Thread(target=self.print_thread_function, args=(self.print_event,), daemon=True)
        self.printer_thread.start()

        # Initial printer detection
        self.refresh_printer()
        self.window.mainloop()

    # ---------- Menu ----------
    def _build_menu(self):
        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)
        m_tools = tk.Menu(menubar, tearoff=0)
        m_tools.add_command(label="Test Print", command=self._threaded_test_print)
        m_tools.add_command(label="Detect COM Ports", command=self.menu_detect_com_ports)
        menubar.add_cascade(label="Tools", menu=m_tools)

    # ---------- Status ----------
    def update_status(self, msg: str, ok: bool = True):
        def do():
            self.lbl_status.config(text=f"Last print: {'OK' if ok else 'ERROR'} — {msg}",
                                   fg="#2ECC71" if ok else "#E74C3C")
        try:
            self.window.after(0, do)
        except Exception:
            pass

    # ---------- Printer ----------
    def refresh_printer(self):
        com_port = self.detect_printer()
        if com_port:
            self.lbl_printer_status.config(text=f"Printer: {com_port} - Connected", fg="green")
            self.printer_port = com_port
        else:
            self.lbl_printer_status.config(text="Printer: Not Found", fg="red")
            self.printer_port = None

    def detect_printer(self) -> Optional[str]:
        ports = list(serial.tools.list_ports.comports())
        for port in ports:
            try:
                with ThermalPrinter(port=port.device, heat_time=HEAT_TIME) as printer:
                    if hasattr(printer, 'identify'):
                        try:
                            if printer.identify() == "ThermalPrinter":
                                return port.device
                        except Exception:
                            pass
                    else:
                        return port.device
            except Exception:
                continue
        return None

    # ---------- Printing threads ----------
    def print_thread_function(self, event):
        """Handles printing of self.display_image in a background thread (image jobs)."""
        while True:
            if event.wait(1):
                if self.printer_port is None:
                    print("Printer not connected.")
                    event.clear()
                    continue
                try:
                    # Use band-wise printing and rotation inside helper
                    self.printer_helper.print_image_bandwise(
                        self.display_image,
                        band_height=64,
                        base_sleep=0.025,
                        dark_bonus=0.20,
                        feed_lines=10,
                        cancel_flag_getter=lambda: self.print_cancel_flag
                    )
                    self.update_status("Image printed", ok=True)
                except Exception as e:
                    print(f"Printing error: {e}")
                    self.update_status(f"Image print error: {e}", ok=False)
                event.clear()

    def cancel_print(self):
        """Cancels ongoing image/text prints (checked between bands)."""
        self.print_cancel_flag = True

    # ---------- IMAGE PRINT DIALOG ----------
    def open_image_print_dialog(self):
        """Modal dialog with: Original + Processed previews, Brightness, Contrast, Open, Rotate, Print."""
        dlg = tk.Toplevel(self.window)
        dlg.title("Print Image")
        dlg.configure(bg="#2C3E50")
        dlg.transient(self.window)
        dlg.grab_set()

        # Layout: previews (left) and controls (right)
        previews = tk.Frame(dlg, bg="#2C3E50")
        previews.pack(side=tk.LEFT, padx=12, pady=10)
        controls = tk.Frame(dlg, bg="#2C3E50")
        controls.pack(side=tk.RIGHT, padx=12, pady=10, fill=tk.Y)

        # Headers
        tk.Label(previews, text="Original Image", fg="#ECF0F1", bg="#2C3E50").grid(row=0, column=0, padx=5, pady=4)
        tk.Label(previews, text="Processed Image", fg="#ECF0F1", bg="#2C3E50").grid(row=0, column=1, padx=5, pady=4)

        lbl_orig = tk.Label(previews, bg="#34495E")
        lbl_proc = tk.Label(previews, bg="#34495E")
        lbl_orig.grid(row=1, column=0, padx=5, pady=5)
        lbl_proc.grid(row=1, column=1, padx=5, pady=5)

        # Local dialog state
        state = {
            "src": (self.source_image.copy() if self.source_image is not None else None),
            "brightness": 0.0,
            "contrast": 0.0,
            "imgtk_orig": None,
            "imgtk_proc": None,
            "processed_1bit": None,
        }

        def load_file():
            filename = filedialog.askopenfilename(parent=dlg)
            if not filename:
                return
            try:
                state["src"] = Image.open(filename)
                refresh_previews()
            except Exception as e:
                messagebox.showerror("Open Image", f"Failed to open: {e}", parent=dlg)

        def rotate_local():
            if state["src"] is not None:
                state["src"] = state["src"].rotate(90, expand=True)
                refresh_previews()

        def refresh_previews():
            # No image yet
            if state["src"] is None:
                lbl_orig.config(image="", text="No image loaded", fg="#ECF0F1", bg="#2C3E50")
                lbl_proc.config(image="", text="No image loaded", fg="#ECF0F1", bg="#2C3E50")
                state["processed_1bit"] = None
                return

            # Resize original to printer width for processing, but show scaled previews
            factor = WIDTH_PIXELS / state["src"].width
            new_size = (WIDTH_PIXELS, max(1, round(state["src"].height * factor)))
            orig_resized = state["src"].resize(new_size, RESAMPLE_FILTER)

            # Apply adjustments to make processed preview + printable image
            proc = ImageEnhance.Brightness(orig_resized).enhance(1 + state["brightness"])
            proc = ImageEnhance.Contrast(proc).enhance(1 + state["contrast"])
            proc_1b = proc.convert("1", dither=Image.FLOYDSTEINBERG)
            state["processed_1bit"] = ensure_filename(proc_1b, "dlg_print.bmp")

            # Screen previews (scale down if large)
            def to_preview(image):
                max_w = 280
                scale = min(1.0, max_w / image.width)
                im = image.resize((int(image.width * scale), int(image.height * scale)), RESAMPLE_FILTER)
                imgtk = ImageTk.PhotoImage(im)
                return imgtk

            imgtk_o = to_preview(orig_resized.convert("L"))
            imgtk_p = to_preview(proc_1b)
            lbl_orig.config(image=imgtk_o);  lbl_orig.image = imgtk_o
            lbl_proc.config(image=imgtk_p);  lbl_proc.image = imgtk_p

        # Controls on right
        tk.Button(controls, text="Open Image...", command=load_file, bg="#1ABC9C", fg="#FFFFFF").pack(fill=tk.X, pady=3)
        tk.Button(controls, text="Rotate 90°", command=rotate_local, bg="#1ABC9C", fg="#FFFFFF").pack(fill=tk.X, pady=3)

        tk.Label(controls, text="Brightness", fg="#ECF0F1", bg="#2C3E50").pack(pady=(8, 0))
        s_b = tk.Scale(controls, from_=-1.0, to=1.0, resolution=0.1, orient=tk.HORIZONTAL,
                       length=200, bg="#34495E", fg="#ECF0F1",
                       command=lambda v: (state.update({"brightness": float(v)}), refresh_previews()))
        s_b.set(0.0); s_b.pack(pady=2)

        tk.Label(controls, text="Contrast", fg="#ECF0F1", bg="#2C3E50").pack(pady=(8, 0))
        s_c = tk.Scale(controls, from_=-1.0, to=1.0, resolution=0.1, orient=tk.HORIZONTAL,
                       length=200, bg="#34495E", fg="#ECF0F1",
                       command=lambda v: (state.update({"contrast": float(v)}), refresh_previews()))
        s_c.set(0.0); s_c.pack(pady=2)

        def do_print_and_close():
            if state["processed_1bit"] is None:
                messagebox.showerror("Print Image", "No image loaded.", parent=dlg)
                return
            # Save into app state and trigger background image print
            self.source_image = state["src"]
            self.display_image = state["processed_1bit"]
            self.print_event.set()
            dlg.destroy()

        tk.Button(controls, text="Print", command=do_print_and_close,
                  bg="#27AE60", fg="#FFFFFF", font=("Arial", 11, "bold")).pack(fill=tk.X, pady=(12, 3))
        tk.Button(controls, text="Close", command=dlg.destroy, bg="#95A5A6", fg="#FFFFFF").pack(fill=tk.X, pady=2)

        refresh_previews()

        # Center dialog over main
        self.window.update_idletasks(); dlg.update_idletasks()
        x = self.window.winfo_x() + (self.window.winfo_width() // 2) - (dlg.winfo_reqwidth() // 2)
        y = self.window.winfo_y() + (self.window.winfo_height() // 2) - (dlg.winfo_reqheight() // 2)
        dlg.geometry(f"+{x}+{y}")

    # ---------- TEXT PRINT DIALOG ----------
    def open_text_print_dialog(self):
        """Modal dialog with a multiline text box; auto-paste clipboard text if available."""
        dlg = tk.Toplevel(self.window)
        dlg.title("Print Text")
        dlg.configure(bg="#2C3E50")
        dlg.transient(self.window)
        dlg.grab_set()

        tk.Label(dlg, text="Text to print:", fg="#ECF0F1", bg="#2C3E50").pack(anchor="w", padx=10, pady=(10,4))
        txt = tk.Text(dlg, width=48, height=15, bg="#34495E", fg="#ECF0F1", insertbackground="#ECF0F1")
        txt.pack(padx=10, pady=4)

        # Auto-paste clipboard text if present
        try:
            clip = self.window.clipboard_get()
            if isinstance(clip, str) and clip.strip():
                txt.insert("1.0", clip.strip())
        except Exception:
            pass

        def do_print_text_and_close():
            content = txt.get("1.0", "end").strip("\n")
            if not content.strip():
                messagebox.showerror("Print Text", "Nothing to print.", parent=dlg)
                return
            threading.Thread(target=self._print_text_job, args=(content,), daemon=True).start()
            dlg.destroy()

        footer = tk.Frame(dlg, bg="#2C3E50"); footer.pack(fill=tk.X, padx=10, pady=8)
        tk.Button(footer, text="Print", command=do_print_text_and_close, bg="#27AE60",
                  fg="#FFFFFF", font=("Arial", 11, "bold")).pack(side=tk.RIGHT, padx=(6,0))
        tk.Button(footer, text="Close", command=dlg.destroy, bg="#95A5A6", fg="#FFFFFF").pack(side=tk.RIGHT)

        # Center
        self.window.update_idletasks(); dlg.update_idletasks()
        x = self.window.winfo_x() + (self.window.winfo_width() // 2) - (dlg.winfo_reqwidth() // 2)
        y = self.window.winfo_y() + (self.window.winfo_height() // 2) - (dlg.winfo_reqheight() // 2)
        dlg.geometry(f"+{x}+{y}")

    def _print_text_job(self, text: str):
        try:
            self.print_cancel_flag = False
            header = self.printer_helper.center("TEXT PRINT")
            lines = [header, self.printer_helper.separator()]
            lines.extend(wrap(text, LINE_WIDTH_CHARS))
            # Render + print (rotated inside helper)
            self.printer_helper.print_lines(lines, cancel_flag_getter=lambda: self.print_cancel_flag)
            self.update_status("Text printed", ok=True)
        except Exception as e:
            self.update_status(f"Text print error: {e}", ok=False)

    # ---------- Desk Buddy workers ----------
    def _threaded_print_calendar(self):
        threading.Thread(target=self.print_today_calendar, daemon=True).start()

    def _threaded_print_weather(self):
        threading.Thread(target=self.print_weather, daemon=True).start()

    def _threaded_print_meme(self):
        threading.Thread(target=self.print_meme, daemon=True).start()

    def _threaded_print_joke(self):
        threading.Thread(target=self.print_joke, daemon=True).start()

    def _threaded_test_print(self):
        threading.Thread(target=self.menu_test_print_impl, daemon=True).start()

    # ---------- Implementations ----------
    def print_today_calendar(self):
        self.print_cancel_flag = False
        try:
            events = self.calendar_fetcher.get_today_events()
            today_str = dt.datetime.now(get_tz()).strftime("%Y-%m-%d")
            lines = [self.printer_helper.center(f"CALENDAR - {today_str}"),
                     self.printer_helper.separator()]
            if not events:
                lines.append("No events today.")
            else:
                for ev in events:
                    loc = f"  [{ev['location']}]" if ev.get("location") else ""
                    line = f"{ev['start']}-{ev['end']}  {ev['subject']}{loc}"
                    lines.extend(wrap(line, LINE_WIDTH_CHARS))
            self.printer_helper.print_lines(lines, cancel_flag_getter=lambda: self.print_cancel_flag)
            self.update_status("Calendar printed", ok=True)
        except Exception as e:
            self.update_status(f"Calendar error: {e}", ok=False)
            messagebox.showerror("Calendar", f"Failed to print calendar: {e}")

    def _print_weather_icon(self, weather_code: int):
        try:
            category = weather_category_from_code(weather_code)
            base_dir = get_asset_path("assets", "weather")
            path: Optional[str] = None

            preferred = WEATHER_ICON_FILES.get(category, "cloud.png").lower()
            if base_dir:
                p = Path(base_dir)
                files = [c for c in p.iterdir() if c.is_file() and c.suffix.lower() in (".png", ".jpg", ".jpeg", ".bmp")]
                for f in files:
                    if f.name.lower() == preferred:
                        path = str(f); break
                if not path:
                    for f in files:
                        if category in f.name.lower():
                            path = str(f); break
                if not path and files:
                    path = str(files[0])

            if path and os.path.isfile(path):
                img = Image.open(path)
                if img.mode in ("RGBA", "LA"):
                    bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
                    bg.paste(img, (0, 0), img)
                    img = bg.convert("L")
                else:
                    img = img.convert("L")
            else:
                from PIL import ImageDraw, ImageFont
                img = Image.new("L", (256, 80), 255)
                d = ImageDraw.Draw(img)
                label = category.upper()
                try:
                    f = ImageFont.truetype("arial.ttf", 28)
                except Exception:
                    f = ImageFont.load_default()
                d.text((10, 25), label, fill=0, font=f)

            img = trim_white_borders(img)
            target_w = min(220, WIDTH_PIXELS)
            scale = target_w / img.width
            new_size = (int(target_w), max(1, int(img.height * scale)))
            icon = img.resize(new_size, RESAMPLE_FILTER).convert("1", dither=Image.FLOYDSTEINBERG)

            canvas = Image.new("1", (WIDTH_PIXELS, icon.height), 1)
            x = (WIDTH_PIXELS - icon.width) // 2
            canvas.paste(icon, (x, 0))
            canvas = ensure_filename(canvas, "weather_icon.png")
            # Rotate/print inside helper
            self.printer_helper.print_image(canvas, feed_lines=0)
        except Exception as e:
            self.update_status(f"Weather icon skipped: {e}", ok=False)

    def print_weather(self):
        self.print_cancel_flag = False
        try:
            data = self.weather_fetcher.get_current_and_hourly()
            tz = get_tz()

            head_lines = [self.printer_helper.center(f"WEATHER - {CITY_NAME}"),
                          self.printer_helper.separator()]
            self.printer_helper.print_lines(head_lines, cancel_flag_getter=lambda: self.print_cancel_flag, feed_lines=0)

            cw = data.get("current_weather", {}) or {}
            cur_code = int(cw.get("weathercode", 2))
            self._print_weather_icon(cur_code)
            self.printer_helper.print_lines([], feed_lines=1)

            cur_temp = cw.get("temperature")
            cur_wind = cw.get("windspeed")
            cur_wdir = cw.get("winddirection")
            cur_dir = bearing_to_cardinal(cur_wdir)
            now_line = f"Now: {int(round(cur_temp))}°C  Wind {int(round(cur_wind))} km/h {cur_dir}"
            now_block = [now_line, self.printer_helper.separator()]
            self.printer_helper.print_lines(now_block, cancel_flag_getter=lambda: self.print_cancel_flag, feed_lines=1)

            hourly = data.get("hourly", {}) or {}
            times = hourly.get("time", [])
            temps = hourly.get("temperature_2m", [])
            probs = hourly.get("precipitation_probability", [])
            wspeeds = hourly.get("windspeed_10m", [])
            wdirs = hourly.get("winddirection_10m", [])

            parsed_times = [parse_openmeteo_local(t, tz) for t in times]
            now_aw = dt.datetime.now(tz).replace(minute=0, second=0, microsecond=0)

            start_idx = next((i for i, tt in enumerate(parsed_times) if tt >= now_aw),
                             max(0, len(parsed_times) - 12))
            end_idx = min(len(parsed_times), start_idx + 12)

            lines: List[str] = [
                f"Next 12h (from {parsed_times[start_idx].strftime('%H:%M')}):",
                "Time  T  P%  Wind"
            ]

            for i in range(start_idx, end_idx):
                hour = parsed_times[i].strftime("%H:%M")
                tC = temps[i] if i < len(temps) else None
                pp = probs[i] if i < len(probs) and probs[i] is not None else 0
                ws = wspeeds[i] if i < len(wspeeds) else 0
                wd = wdirs[i] if i < len(wdirs) else None
                wdc = bearing_to_cardinal(wd)
                line = f"{hour} {int(round(tC))}° {int(pp)}%  {wdc}{int(round(ws))}"
                lines.append(line)

            daily = data.get("daily", {}) or {}
            tmin = daily.get("temperature_2m_min", [None])[0]
            tmax = daily.get("temperature_2m_max", [None])[0]
            lines.append(self.printer_helper.separator())
            lines.append(f"Today: min {int(round(tmin))}° / max {int(round(tmax))}°")

            self.printer_helper.print_lines(lines, cancel_flag_getter=lambda: self.print_cancel_flag, feed_lines=3)
            self.update_status("Weather printed", ok=True)
        except Exception as e:
            self.update_status(f"Weather error: {e}", ok=False)
            messagebox.showerror("Weather", f"Failed to print weather: {e}")

    def print_meme(self):
        self.print_cancel_flag = False
        try:
            img = self.fun_fetcher.get_random_meme()
            if img is None:
                self._print_fallback_no_meme()
                self.update_status("No meme available", ok=False)
                return
            w = WIDTH_PIXELS
            factor = w / img.width
            new_size = (w, max(1, int(round(img.height * factor))))
            img_resized = img.resize(new_size, RESAMPLE_FILTER)
            img_1bit = img_resized.convert("1", dither=Image.FLOYDSTEINBERG)
            img_1bit = ensure_filename(img_1bit, "meme.png")
            # Band-wise + rotate inside helper
            self.printer_helper.print_image_bandwise(
                img_1bit,
                band_height=64,
                base_sleep=0.025,
                dark_bonus=0.20,
                feed_lines=8,
                cancel_flag_getter=lambda: self.print_cancel_flag
            )
            self.update_status("Meme printed", ok=True)
        except Exception as e:
            self.update_status(f"Meme error: {e}", ok=False)
            self._print_fallback_no_meme()

    def print_joke(self):
        self.print_cancel_flag = False
        try:
            joke = self.fun_fetcher.get_random_joke()
            header = self.printer_helper.center("RANDOM JOKE")
            lines = [header, self.printer_helper.separator()]
            lines.extend(wrap(joke, LINE_WIDTH_CHARS))
            self.printer_helper.print_lines(lines, cancel_flag_getter=lambda: self.print_cancel_flag)
            self.update_status("Joke printed", ok=True)
        except Exception as e:
            self.update_status(f"Joke error: {e}", ok=False)
            messagebox.showerror("Joke", f"Failed to print joke: {e}")

    def _print_fallback_no_meme(self):
        try:
            lines = [self.printer_helper.center("MEME"), self.printer_helper.separator(), "No meme available."]
            self.printer_helper.print_lines(lines)
        except Exception:
            pass

    # ---------- Menu actions ----------
    def menu_detect_com_ports(self):
        ports = PrinterHelper.list_com_ports()
        guess = PrinterHelper.guess_printer_port()
        if not ports:
            messagebox.showinfo("COM Ports", "No serial ports found.")
            return
        lines = []
        for p in ports:
            dev = getattr(p, "device", "UNKNOWN")
            desc = getattr(p, "description", "")
            mark = "  [Likely printer]" if guess and dev == guess else ""
            lines.append(f"{dev}  —  {desc}{mark}")
        messagebox.showinfo("COM Ports", "\n".join(lines))

    def menu_test_print_impl(self):
        try:
            lines = [
                self.printer_helper.center("Thermal Print3r Tool"),
                self.printer_helper.center("Desk Buddy Test"),
                self.printer_helper.separator(),
                "Sample line 1: Hello, world!",
                "Sample line 2: ##########",
                "<SEPARATOR>",
                "Wrap test: " + ("The quick brown fox jumps over the lazy dog. " * 2)
            ]
            self.printer_helper.print_lines(lines)
            self.update_status("Test print OK", ok=True)
        except Exception as e:
            self.update_status(f"Test print error: {e}", ok=False)
            messagebox.showerror("Test Print", f"Failed: {e}")

    def _threaded_test_print(self):
        threading.Thread(target=self.menu_test_print_impl, daemon=True).start()


# -------------------------
# Entry point
# -------------------------

if __name__ == "__main__":
    ThermalPrintTool()

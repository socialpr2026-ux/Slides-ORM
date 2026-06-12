from __future__ import annotations

from pathlib import Path
import json
import os
import re
import shutil
import subprocess
import time

from PIL import Image


WINDOWS_BROWSER_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]

LINUX_BROWSER_NAMES = [
    "google-chrome",
    "google-chrome-stable",
    "chromium",
    "chromium-browser",
    "microsoft-edge",
]

COOKIE_BUTTON_PATTERNS = [
    r"^(принять|согласен|согласна|хорошо|понятно|ок)$",
    r"(принять.*cookie|accept.*cookie|accept all|allow all)",
    r"^(accept|agree|got it|ok)$",
]

BLOCKED_TEXT_PATTERNS = [
    r"access denied",
    r"forbidden",
    r"captcha",
    r"robot check",
    r"проверка безопасности",
    r"доступ ограничен",
]

CAPTURE_VERSION = 23


def _browser_executable() -> str:
    for env_name in ("AI_ORM_BROWSER_PATH", "PLAYWRIGHT_CHROMIUM_EXECUTABLE"):
        value = os.environ.get(env_name, "").strip()
        if value and Path(value).exists():
            return value
    for path in WINDOWS_BROWSER_PATHS:
        if Path(path).exists():
            return path
    for name in LINUX_BROWSER_NAMES:
        resolved = shutil.which(name)
        if resolved:
            return resolved
    return ""


def _text_snippets(text: str, limit: int = 6) -> list[str]:
    clean = re.sub(r"\s+", " ", str(text or "")).strip()
    if not clean:
        return []
    candidates: list[str] = []
    for sentence in re.split(r"(?<=[.!?])\s+", clean):
        sentence = sentence.strip(" .,!?:-—")
        if len(sentence) >= 18:
            candidates.append(sentence[:150])
    if len(clean) >= 32:
        candidates.extend([clean[:150], clean[max(0, len(clean) // 2 - 75): len(clean) // 2 + 75]])
    unique: list[str] = []
    for item in candidates:
        item = re.sub(r"\s+", " ", item).strip()
        if item and item not in unique:
            unique.append(item)
    return unique[:limit]


def _write_meta(path: Path, meta: dict) -> None:
    path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


def _looks_like_browser_error_screenshot(path: Path) -> bool:
    """Detect Chrome/Edge network block pages so they are not treated as source screenshots."""
    try:
        with Image.open(path) as source:
            image = source.convert("RGB").resize((320, 180))
            hsv = image.convert("HSV")
    except Exception:
        return False
    pixels = list(image.getdata())
    hsv_pixels = list(hsv.getdata())
    if not pixels:
        return False
    mean_brightness = sum(sum(pixel) / 3 for pixel in pixels) / len(pixels)
    mean_saturation = sum(pixel[1] for pixel in hsv_pixels) / len(hsv_pixels)
    near_gray = sum(1 for r, g, b in pixels if max(r, g, b) - min(r, g, b) <= 8) / len(pixels)
    very_light = sum(1 for r, g, b in pixels if (r + g + b) / 3 >= 232) / len(pixels)
    return mean_brightness >= 232 and mean_saturation <= 8 and near_gray >= 0.92 and very_light >= 0.70


def _cached_result(output_path: Path, meta_path: Path, url: str, width_px: int, height_px: int) -> dict | None:
    if not output_path.exists() or not meta_path.exists():
        return None
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if (
        meta.get("status") == "captured"
        and meta.get("url") == url
        and int(meta.get("width_px", 0) or 0) == int(width_px)
        and int(meta.get("height_px", 0) or 0) == int(height_px)
        and int(meta.get("capture_version", 0) or 0) == CAPTURE_VERSION
    ):
        if _looks_like_browser_error_screenshot(output_path):
            return None
        meta["status"] = "cached"
        meta["image_path"] = str(output_path)
        return meta
    return None


def _frame_image(path: Path, width_px: int, height_px: int) -> None:
    with Image.open(path) as source:
        image = source.convert("RGB")
    image.save(path)


def _click_cookie_buttons(page) -> None:
    for pattern in COOKIE_BUTTON_PATTERNS:
        regex = re.compile(pattern, re.IGNORECASE)
        for locator_factory in (
            lambda: page.get_by_role("button", name=regex),
            lambda: page.get_by_text(regex),
        ):
            try:
                locator = locator_factory().first
                if locator.count() > 0:
                    locator.click(timeout=900)
                    page.wait_for_timeout(250)
                    return
            except Exception:
                continue


def _hide_fixed_cookie_overlays(page) -> None:
    page.evaluate(
        """
        () => {
          const needles = ['cookie', 'cookies', 'consent', 'gdpr', 'privacy', 'сookie', 'куки', 'install', 'app', 'приложение', 'установите'];
          for (const el of Array.from(document.querySelectorAll('body *'))) {
            const visibleText = (el.innerText || '').slice(0, 500);
            const text = `${el.id || ''} ${el.className || ''} ${el.getAttribute('aria-label') || ''} ${visibleText}`.toLowerCase();
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            const zIndex = Number.parseInt(style.zIndex || '0', 10) || 0;
            const isNamedConsent = needles.some((needle) => text.includes(needle));
            const isLargeOverlay = zIndex >= 1000 && rect.width > window.innerWidth * 0.15 && rect.height > window.innerHeight * 0.10;
            const isFloating = (style.position === 'fixed' || style.position === 'sticky') && rect.width > 18 && rect.height > 18;
            const isOverlayPosition = isFloating || style.position === 'absolute';
            if (((isNamedConsent || isLargeOverlay) && isOverlayPosition) || isFloating) {
              el.style.setProperty('display', 'none', 'important');
            }
          }
        }
        """
    )


def _hide_inline_ads(page) -> None:
    page.evaluate(
        """
        () => {
          const needles = ['реклама', 'advertisement', 'sponsored', 'promoted'];
          for (const el of Array.from(document.querySelectorAll('body *'))) {
            const visibleText = (el.innerText || '').slice(0, 800);
            const text = `${el.id || ''} ${el.className || ''} ${el.getAttribute('aria-label') || ''} ${visibleText}`.toLowerCase();
            if (!needles.some((needle) => text.includes(needle))) continue;
            const rect = el.getBoundingClientRect();
            const textLength = (el.innerText || '').trim().length;
            const imageCount = el.querySelectorAll('img, iframe, picture').length;
            const looksLikeAdUnit = rect.width > 80 && rect.height > 35 && (textLength < 520 || rect.height < 270);
            const looksLikeAdGrid = rect.width > 220 && rect.height > 80 && imageCount >= 2;
            if (looksLikeAdUnit || looksLikeAdGrid) {
              el.style.setProperty('visibility', 'hidden', 'important');
              el.style.setProperty('max-height', '0px', 'important');
              el.style.setProperty('overflow', 'hidden', 'important');
            }
          }
        }
        """
    )


def _scroll_to_text(page, target_text: str, target_ratio: float, compact_crop: bool) -> dict:
    snippets = _text_snippets(target_text)
    if not snippets:
        return {"matched_text": False, "matched_snippet": ""}
    return page.evaluate(
        """
        ([snippets, targetRatio, expandMetadata, compactCrop]) => {
          const normalize = (value) => (value || '').toLowerCase().replace(/\\s+/g, ' ').trim();
          const wanted = snippets.map(normalize).filter(Boolean);
          if (!wanted.length || !document.body) return {matched_text: false, matched_snippet: ''};
          let best = null;
          let bestSnippet = '';
          let bestScore = 0;
          const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
          while (walker.nextNode()) {
            const node = walker.currentNode;
            const text = normalize(node.nodeValue);
            if (!text || text.length < 12) continue;
            for (let snippetIndex = 0; snippetIndex < wanted.length; snippetIndex += 1) {
              const snippet = wanted[snippetIndex];
              let score = 0;
              if (text.includes(snippet)) score = snippet.length;
              else if (snippet.includes(text) && text.length > 24) score = text.length;
              if (score > 0) score += Math.max(0, wanted.length - snippetIndex) * 180;
              if (score > bestScore && node.parentElement) {
                best = node.parentElement;
                bestSnippet = snippet;
                bestScore = score;
              }
            }
          }
          if (!best) return {matched_text: false, matched_snippet: ''};
          let chosen = best;
          let cursor = best;
          let chosenScore = -Infinity;
          const snippetLength = bestSnippet.length || 80;
          const maxTextLength = Math.min(1400, Math.max(520, snippetLength * 3.5));
          const metaRegex = /(\\b\\d{1,2}[.\\s](январ|феврал|март|апрел|мая|июн|июл|август|сентябр|октябр|ноябр|декабр)|\\b\\d{1,2}\\.\\d{1,2}\\.\\d{2,4}|\\b20\\d{2}\\b|нравится|ответить|автор|мама|опубликовано|дата|отзыв рекомендуют|рекомендую|отзывов|репутация)/;
          const authorRegex = /(мама|автор|опубликовано|отзывов|репутация|^[а-яa-z0-9_. -]{2,32}\\s*[·,])/;
          const adRegex = /(реклама|advertisement|ad choices|sponsored|промо|читать также|похожие материалы)/;
          for (let i = 0; i < 10 && cursor; i += 1) {
            const rect = cursor.getBoundingClientRect();
            const text = normalize(cursor.innerText || cursor.textContent || '');
            const hasMeta = metaRegex.test(text);
            const hasAuthorish = authorRegex.test(text);
            const lengthOk = text.length <= maxTextLength || (hasMeta && text.length <= 2600);
            if (rect.width > 180 && rect.height > 25 && text.includes(bestSnippet) && lengthOk) {
              let score = 1000;
              if (hasMeta) score += 350;
              if (hasAuthorish) score += 120;
              score += Math.min(rect.width, 900) / 10;
              if (text.length > snippetLength * 1.15) score += 35;
              if (/ответ|ответить|нравится|мама|автор|рекоменд/.test(text)) score += 25;
              if (rect.width > 350) score += 20;
              if (adRegex.test(text)) score -= 700;
              score -= Math.abs(text.length - snippetLength * 1.6) / 6;
              score -= Math.max(0, rect.height - 260) / 2;
              if (score > chosenScore) {
                chosen = cursor;
                chosenScore = score;
              }
            }
            cursor = cursor.parentElement;
          }
          chosen.scrollIntoView({block: 'center', inline: 'nearest'});
          document.querySelectorAll('[data-ai-orm-crop-target="1"]').forEach((el) => el.removeAttribute('data-ai-orm-crop-target'));
          chosen.setAttribute('data-ai-orm-crop-target', '1');
          const rect = chosen.getBoundingClientRect();
          const anchorRect = best.getBoundingClientRect();
          const pageWidth = Math.max(document.documentElement.scrollWidth, document.body ? document.body.scrollWidth : 0, window.innerWidth);
          const pageHeight = Math.max(document.documentElement.scrollHeight, document.body ? document.body.scrollHeight : 0, window.innerHeight);
          let x = rect.left + window.scrollX - 12;
          let y = rect.top + window.scrollY - 12;
          let width = rect.width + 24;
          let height = rect.height + 24;
          if (expandMetadata) {
            const metaPadTop = Math.min(120, Math.max(34, height * 0.28));
            const metaPadBottom = Math.min(64, Math.max(26, height * 0.14));
            y -= metaPadTop;
            height += metaPadTop + metaPadBottom;
          }
          if (compactCrop && targetRatio && height > 0) {
            const maxClipHeight = Math.max(150, Math.min(height, (width / targetRatio) * 1.32));
            if (height > maxClipHeight) {
              const matchTop = anchorRect.top + window.scrollY;
              const desiredTop = expandMetadata ? matchTop - Math.min(132, maxClipHeight * 0.42) : matchTop - 16;
              y = Math.max(0, desiredTop);
              height = maxClipHeight;
            }
          }
          width = Math.min(width, pageWidth - 2);
          height = Math.min(height, pageHeight - 2);
          x = Math.max(0, Math.min(x, pageWidth - width - 1));
          y = Math.max(0, Math.min(y, pageHeight - height - 1));
          return {matched_text: true, matched_snippet: bestSnippet, clip: {x, y, width, height}};
        }
        """,
        [
            snippets,
            float(target_ratio or 0),
            bool(re.search(r"forum\.baby|irecommend|otzyv|vseotzyvy", str(target_text or "") + " " + str(getattr(page, 'url', '') or ""), re.IGNORECASE)),
            bool(compact_crop),
        ],
    )


def _blocked_by_page_text(text: str) -> bool:
    lower = str(text or "").lower()
    return any(re.search(pattern, lower) for pattern in BLOCKED_TEXT_PATTERNS)


def _capture_with_chrome_cli(
    *,
    browser_path: str,
    url: str,
    output_path: Path,
    meta_path: Path,
    base_meta: dict,
    target_text: str,
    width_px: int,
    height_px: int,
    timeout_ms: int,
) -> dict:
    viewport_width = max(640, min(int(width_px), 1400))
    viewport_height = max(420, min(int(height_px), 1600))
    timeout_s = max(15, int(timeout_ms / 1000))
    common_args = [
        browser_path,
        "--headless=new",
        "--disable-gpu",
        "--disable-dev-shm-usage",
        "--disable-extensions",
        "--disable-infobars",
        "--no-sandbox",
        "--no-first-run",
        "--no-default-browser-check",
        "--hide-scrollbars",
        "--lang=ru-RU",
        f"--window-size={viewport_width},{viewport_height}",
    ]
    try:
        screenshot_result = subprocess.run(
            [*common_args, f"--screenshot={output_path}", "--virtual-time-budget=2500", url],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=timeout_s,
        )
        if screenshot_result.returncode != 0 or not output_path.exists() or output_path.stat().st_size <= 0:
            meta = {
                **base_meta,
                "status": "failed",
                "reason": "chrome_cli_screenshot_failed",
                "browser_path": browser_path,
                "error": f"browser exited with code {screenshot_result.returncode}",
            }
            _write_meta(meta_path, meta)
            return meta

        _frame_image(output_path, width_px, height_px)
        if _looks_like_browser_error_screenshot(output_path):
            meta = {
                **base_meta,
                "status": "failed",
                "reason": "browser_error_or_network_block_page",
                "browser_path": browser_path,
                "capture_mode": "chrome_cli",
            }
            _write_meta(meta_path, meta)
            return meta
        meta = {
            **base_meta,
            "status": "captured",
            "reason": "ok",
            "browser_path": browser_path,
            "matched_text": False,
            "matched_snippet": "",
            "crop_mode": "viewport",
            "capture_mode": "chrome_cli",
        }
        _write_meta(meta_path, meta)
        return meta
    except Exception as exc:
        meta = {
            **base_meta,
            "status": "failed",
            "reason": "chrome_cli_exception",
            "browser_path": browser_path,
            "error": str(exc),
        }
        _write_meta(meta_path, meta)
        return meta


def capture_browser_screenshot(
    *,
    url: str,
    output_path: Path,
    target_text: str = "",
    width_px: int,
    height_px: int,
    timeout_ms: int = 15_000,
    crop_to_match: bool = True,
    compact_crop: bool = False,
) -> dict:
    """Capture a real browser screenshot when Playwright and a local browser are available."""
    output_path = Path(output_path)
    meta_path = output_path.with_suffix(".json")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    url = str(url or "").strip()
    width_px = int(width_px)
    height_px = int(height_px)
    base_meta = {
        "url": url,
        "image_path": str(output_path),
        "width_px": width_px,
        "height_px": height_px,
        "capture_version": CAPTURE_VERSION,
        "captured_at_epoch": int(time.time()),
    }
    if not url.lower().startswith(("http://", "https://")):
        meta = {**base_meta, "status": "failed", "reason": "missing_url"}
        _write_meta(meta_path, meta)
        return meta
    cached = _cached_result(output_path, meta_path, url, width_px, height_px)
    if cached:
        return cached
    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:
        browser_path = _browser_executable()
        if browser_path:
            meta = _capture_with_chrome_cli(
                browser_path=browser_path,
                url=url,
                output_path=output_path,
                meta_path=meta_path,
                base_meta=base_meta,
                target_text=target_text,
                width_px=width_px,
                height_px=height_px,
                timeout_ms=timeout_ms,
            )
            meta["playwright_error"] = str(exc)
            _write_meta(meta_path, meta)
            return meta
        meta = {**base_meta, "status": "failed", "reason": "playwright_unavailable", "error": str(exc)}
        _write_meta(meta_path, meta)
        return meta
    browser_path = _browser_executable()
    if not browser_path:
        meta = {**base_meta, "status": "failed", "reason": "browser_unavailable"}
        _write_meta(meta_path, meta)
        return meta

    viewport = {
        "width": max(640, min(width_px, 900)),
        "height": max(420, min(height_px, 1500)),
    }
    browser = None
    context = None
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                executable_path=browser_path,
                headless=True,
                args=[
                    "--disable-dev-shm-usage",
                    "--disable-extensions",
                    "--disable-infobars",
                    "--disable-blink-features=AutomationControlled",
                    "--no-first-run",
                    "--lang=ru-RU",
                ],
            )
            context = browser.new_context(
                viewport=viewport,
                device_scale_factor=1,
                locale="ru-RU",
                ignore_https_errors=True,
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
                ),
            )
            page = context.new_page()
            response = page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            status = response.status if response else None
            if status and status >= 400:
                meta = {**base_meta, "status": "failed", "reason": "http_error", "http_status": status, "browser_path": browser_path}
                _write_meta(meta_path, meta)
                return meta
            try:
                page.wait_for_load_state("networkidle", timeout=5_000)
            except Exception:
                pass
            _click_cookie_buttons(page)
            _hide_fixed_cookie_overlays(page)
            _hide_inline_ads(page)
            body_text = ""
            try:
                body_text = page.locator("body").inner_text(timeout=2_000)[:3_000]
            except Exception:
                pass
            if _blocked_by_page_text(body_text):
                meta = {**base_meta, "status": "failed", "reason": "blocked_page", "http_status": status, "browser_path": browser_path}
                _write_meta(meta_path, meta)
                return meta
            match = _scroll_to_text(page, target_text, width_px / max(height_px, 1), compact_crop)
            page.wait_for_timeout(650)
            _hide_fixed_cookie_overlays(page)
            _hide_inline_ads(page)
            crop_mode = "viewport"
            if crop_to_match and match.get("matched_text"):
                try:
                    clip = match.get("clip") or None
                    if clip:
                        page.screenshot(path=str(output_path), clip=clip, full_page=True)
                        crop_mode = "matched_clip"
                    else:
                        page.locator('[data-ai-orm-crop-target="1"]').screenshot(path=str(output_path), timeout=5_000)
                        crop_mode = "matched_element"
                except Exception:
                    try:
                        page.locator('[data-ai-orm-crop-target="1"]').screenshot(path=str(output_path), timeout=5_000)
                        crop_mode = "matched_element"
                    except Exception:
                        page.screenshot(path=str(output_path), full_page=False)
            else:
                page.screenshot(path=str(output_path), full_page=False)
            _frame_image(output_path, width_px, height_px)
            meta = {
                **base_meta,
                "status": "captured",
                "reason": "ok",
                "http_status": status,
                "browser_path": browser_path,
                "matched_text": bool(match.get("matched_text")),
                "matched_snippet": match.get("matched_snippet", ""),
                "crop_mode": crop_mode,
            }
            _write_meta(meta_path, meta)
            return meta
    except Exception as exc:
        meta = {**base_meta, "status": "failed", "reason": "exception", "error": str(exc), "browser_path": browser_path}
        _write_meta(meta_path, meta)
        return meta
    finally:
        try:
            if context:
                context.close()
        except Exception:
            pass
        try:
            if browser:
                browser.close()
        except Exception:
            pass

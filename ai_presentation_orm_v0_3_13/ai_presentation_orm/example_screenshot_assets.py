from __future__ import annotations

from pathlib import Path
from urllib.parse import quote, urlparse
from urllib.request import Request, urlopen
import json
import mimetypes
import time


ASSET_DOWNLOAD_VERSION = 1
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"


def _write_meta(path: Path, meta: dict) -> None:
    path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


def _read_json_url(url: str, timeout: int = 25) -> dict:
    req = Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"})
    with urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def _download_bytes(url: str, timeout: int = 40) -> tuple[bytes, str]:
    req = Request(url, headers={"User-Agent": USER_AGENT, "Accept": "image/*,*/*;q=0.8"})
    with urlopen(req, timeout=timeout) as response:
        content_type = response.headers.get("Content-Type", "").split(";", 1)[0].strip().lower()
        return response.read(), content_type


def _direct_download_url(url: str) -> tuple[str, str]:
    low = url.lower()
    if "disk.yandex." in low or "yadi.sk/" in low:
        api = "https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key=" + quote(url, safe="")
        payload = _read_json_url(api)
        href = str(payload.get("href") or "")
        if not href:
            raise ValueError("Yandex Disk download href is missing")
        return href, "yandex_public_api"
    return url, "direct_url"


def _looks_like_image(data: bytes, content_type: str) -> bool:
    if content_type.startswith("image/"):
        return True
    signatures = [b"\x89PNG\r\n\x1a\n", b"\xff\xd8\xff", b"GIF87a", b"GIF89a", b"RIFF"]
    return any(data.startswith(signature) for signature in signatures)


def _extension(content_type: str, source_url: str) -> str:
    parsed_suffix = Path(urlparse(source_url).path).suffix.lower()
    if parsed_suffix in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}:
        return parsed_suffix
    guessed = mimetypes.guess_extension(content_type or "") or ".png"
    return ".jpg" if guessed == ".jpe" else guessed


def fetch_original_screenshot_asset(*, url: str, output_dir: Path, stem: str) -> dict:
    """Download a ready screenshot asset without image transformations."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    url = str(url or "").strip()
    meta_path = output_dir / f"{stem}.json"
    base_meta = {
        "source_url": url,
        "asset_download_version": ASSET_DOWNLOAD_VERSION,
        "downloaded_at_epoch": int(time.time()),
    }
    if not url.lower().startswith(("http://", "https://")):
        meta = {**base_meta, "status": "failed", "reason": "missing_url"}
        _write_meta(meta_path, meta)
        return meta
    if meta_path.exists():
        try:
            cached = json.loads(meta_path.read_text(encoding="utf-8"))
            image_path = Path(cached.get("image_path", ""))
            if (
                cached.get("status") == "downloaded"
                and cached.get("source_url") == url
                and int(cached.get("asset_download_version", 0) or 0) == ASSET_DOWNLOAD_VERSION
                and image_path.exists()
            ):
                cached["status"] = "download_cached"
                return cached
        except Exception:
            pass
    try:
        direct_url, method = _direct_download_url(url)
        data, content_type = _download_bytes(direct_url)
        if not _looks_like_image(data, content_type):
            meta = {**base_meta, "status": "failed", "reason": "not_image", "content_type": content_type, "download_method": method}
            _write_meta(meta_path, meta)
            return meta
        image_path = output_dir / f"{stem}{_extension(content_type, direct_url)}"
        image_path.write_bytes(data)
        meta = {
            **base_meta,
            "status": "downloaded",
            "reason": "ok",
            "image_path": str(image_path),
            "content_type": content_type,
            "byte_count": len(data),
            "download_method": method,
        }
        _write_meta(meta_path, meta)
        return meta
    except Exception as exc:
        meta = {**base_meta, "status": "failed", "reason": "exception", "error": str(exc)}
        _write_meta(meta_path, meta)
        return meta

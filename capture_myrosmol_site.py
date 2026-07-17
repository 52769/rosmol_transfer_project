from __future__ import annotations

import asyncio
import json
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

from playwright.async_api import BrowserContext, Page, Response, async_playwright

BASE_URL = "https://myrosmol.ru/projects?my=1"
ROOT = Path(__file__).resolve().parent
PROFILE_DIR = ROOT / "diagnostic_browser_profile"
OUTPUT_ROOT = ROOT / "diagnostics_myrosmol"

SENSITIVE_HEADER_PARTS = (
    "authorization",
    "cookie",
    "token",
    "secret",
    "session",
    "csrf",
    "xsrf",
)
SENSITIVE_QUERY_PARTS = (
    "token",
    "auth",
    "password",
    "passwd",
    "session",
    "cookie",
    "code",
    "secret",
)


def safe_name(value: str, limit: int = 90) -> str:
    value = re.sub(r"[\\/:*?\"<>|\r\n]+", "_", value or "")
    value = re.sub(r"\s+", " ", value).strip(" ._")
    return (value[:limit] or "page").rstrip(" .")


def redact_url(url: str) -> str:
    try:
        parts = urlsplit(url)
        query = []
        for key, value in parse_qsl(parts.query, keep_blank_values=True):
            if any(part in key.casefold() for part in SENSITIVE_QUERY_PARTS):
                value = "[REMOVED]"
            query.append((key, value))
        return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment))
    except Exception:
        return url


def sanitize_headers(headers: list[dict[str, Any]] | None) -> list[dict[str, Any]]:
    cleaned: list[dict[str, Any]] = []
    for header in headers or []:
        name = str(header.get("name", ""))
        item = dict(header)
        if any(part in name.casefold() for part in SENSITIVE_HEADER_PARTS):
            item["value"] = "[REMOVED]"
        cleaned.append(item)
    return cleaned


def sanitize_har(raw_path: Path, clean_path: Path) -> None:
    with raw_path.open("r", encoding="utf-8") as stream:
        data = json.load(stream)

    log = data.get("log", {})
    for entry in log.get("entries", []):
        request = entry.get("request", {})
        response = entry.get("response", {})

        request["url"] = redact_url(str(request.get("url", "")))
        request["headers"] = sanitize_headers(request.get("headers"))
        request["cookies"] = []
        response["headers"] = sanitize_headers(response.get("headers"))
        response["cookies"] = []

        url_lower = str(request.get("url", "")).casefold()
        if any(part in url_lower for part in ("/login", "/signin", "/auth/", "/token")):
            post_data = request.get("postData")
            if isinstance(post_data, dict):
                post_data["text"] = "[REMOVED LOGIN/AUTH BODY]"
                post_data["params"] = []

    with clean_path.open("w", encoding="utf-8") as stream:
        json.dump(data, stream, ensure_ascii=False)


def append_log(path: Path, text: str) -> None:
    stamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    with path.open("a", encoding="utf-8") as stream:
        stream.write(f"[{stamp}] {text}\n")


async def close_cookie_banner(page: Page) -> None:
    try:
        button = page.get_by_role("button", name=re.compile(r"^\s*Принять\s*$", re.I))
        if await button.count() and await button.first.is_visible(timeout=500):
            await button.first.click(timeout=1500)
    except Exception:
        pass


async def snapshot(page: Page, output_dir: Path, index: int, note: str = "") -> int:
    index += 1
    title = ""
    try:
        title = await page.title()
    except Exception:
        pass
    stem = f"{index:03d}_{safe_name(note or title or 'snapshot')}"
    html_path = output_dir / f"{stem}.html"
    png_path = output_dir / f"{stem}.png"
    meta_path = output_dir / f"{stem}.json"

    try:
        html_path.write_text(await page.content(), encoding="utf-8")
    except Exception as exc:
        html_path.write_text(f"<!-- Не удалось сохранить DOM: {exc} -->", encoding="utf-8")
    try:
        await page.screenshot(path=str(png_path), full_page=True)
    except Exception:
        try:
            await page.screenshot(path=str(png_path), full_page=False)
        except Exception:
            pass

    meta = {
        "index": index,
        "note": note,
        "title": title,
        "url": redact_url(page.url),
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Снимок сохранён: {stem}")
    return index


async def save_api_response(response: Response, output_dir: Path, counter: list[int], log_path: Path) -> None:
    try:
        request = response.request
        if request.resource_type not in {"xhr", "fetch"}:
            return
        if "/api/" not in response.url:
            return

        content_type = (response.headers.get("content-type") or "").casefold()
        if not any(token in content_type for token in ("json", "text", "javascript")):
            return

        body = await response.body()
        if len(body) > 8 * 1024 * 1024:
            append_log(log_path, f"API body skipped (>8MB): {response.status} {redact_url(response.url)}")
            return

        counter[0] += 1
        url_parts = urlsplit(response.url)
        suffix = ".json" if "json" in content_type else ".txt"
        filename = f"{counter[0]:04d}_{response.status}_{safe_name(url_parts.path.replace('/', '_'))}{suffix}"
        target = output_dir / filename
        target.write_bytes(body)

        meta = {
            "status": response.status,
            "method": request.method,
            "url": redact_url(response.url),
            "resource_type": request.resource_type,
            "content_type": content_type,
        }
        target.with_suffix(target.suffix + ".meta.json").write_text(
            json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as exc:
        append_log(log_path, f"API response capture error: {exc}")


async def wait_for_manual_login(playwright) -> None:
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    context = await playwright.chromium.launch_persistent_context(
        user_data_dir=str(PROFILE_DIR),
        headless=False,
        viewport={"width": 1440, "height": 1000},
        accept_downloads=True,
        args=["--start-maximized"],
    )
    page = context.pages[0] if context.pages else await context.new_page()
    try:
        await page.goto(BASE_URL, wait_until="domcontentloaded", timeout=60_000)
    except Exception:
        pass

    print("\nЭТАП 1. АВТОРИЗАЦИЯ")
    print("В открытом браузере войдите в аккаунт вручную.")
    print("Когда откроется страница «Мои проекты», вернитесь в консоль и нажмите Enter.")
    await asyncio.to_thread(input, "> ")
    await context.close()


async def record_scenario(playwright, output_dir: Path) -> None:
    raw_har = output_dir / "network_raw.har"
    clean_har = output_dir / "network_sanitized.har"
    trace_path = output_dir / "playwright_trace.zip"
    events_log = output_dir / "browser_events.log"
    api_dir = output_dir / "api_responses"
    snapshots_dir = output_dir / "snapshots"
    api_dir.mkdir(parents=True, exist_ok=True)
    snapshots_dir.mkdir(parents=True, exist_ok=True)

    context: BrowserContext = await playwright.chromium.launch_persistent_context(
        user_data_dir=str(PROFILE_DIR),
        headless=False,
        viewport={"width": 1440, "height": 1000},
        accept_downloads=True,
        record_har_path=str(raw_har),
        record_har_content="embed",
        record_har_mode="full",
        args=["--start-maximized"],
    )
    await context.tracing.start(screenshots=True, snapshots=True, sources=False)
    page = context.pages[0] if context.pages else await context.new_page()

    api_counter = [0]

    def on_console(message) -> None:
        append_log(events_log, f"CONSOLE {message.type}: {message.text}")

    def on_page_error(error) -> None:
        append_log(events_log, f"PAGEERROR: {error}")

    def on_request_failed(request) -> None:
        append_log(
            events_log,
            f"REQUESTFAILED {request.method} {redact_url(request.url)} :: {request.failure}",
        )

    def on_response(response: Response) -> None:
        asyncio.create_task(save_api_response(response, api_dir, api_counter, events_log))

    page.on("console", on_console)
    page.on("pageerror", on_page_error)
    page.on("requestfailed", on_request_failed)
    page.on("response", on_response)

    try:
        await page.goto(BASE_URL, wait_until="domcontentloaded", timeout=60_000)
    except Exception as exc:
        append_log(events_log, f"Initial navigation error: {exc}")
    await close_cookie_banner(page)

    print("\nЭТАП 2. ЗАПИСЬ РУЧНОГО СЦЕНАРИЯ")
    print("В браузере вручную выполните полный перенос одного тестового проекта:")
    print("1) откройте «Добавить проект» и выберите нужный шаблон;")
    print("2) заполните название и регион, нажмите «Создать черновик»;")
    print("3) после создания заполните хотя бы по одному полю на каждой нужной вкладке;")
    print("4) в календарном плане добавьте задачу и мероприятие;")
    print("5) в софинансировании добавьте партнёра;")
    print("6) в расходах добавьте одну строку товара и сохраните проект.")
    print("\nПосле каждого важного экрана вернитесь в консоль и нажмите S, чтобы сохранить DOM и скриншот.")
    print("Когда закончите весь сценарий, нажмите Q.")

    snapshot_index = 0
    snapshot_index = await snapshot(page, snapshots_dir, snapshot_index, "начало записи")

    while True:
        command = (await asyncio.to_thread(input, "\n[S] снимок, [Q] завершить, [H] инструкция: ")).strip().casefold()
        if command in {"q", "й", "quit", "готово"}:
            break
        if command in {"h", "р", "help", "?"}:
            print("S — сохранить текущий DOM и скриншот; Q — завершить и собрать архив.")
            continue
        note = ""
        if command.startswith("s "):
            note = command[2:].strip()
        elif command not in {"", "s", "ы"}:
            note = command
        else:
            note = await asyncio.to_thread(input, "Краткое название экрана (можно Enter): ")
        snapshot_index = await snapshot(page, snapshots_dir, snapshot_index, note)

    snapshot_index = await snapshot(page, snapshots_dir, snapshot_index, "конец записи")
    await context.tracing.stop(path=str(trace_path))
    await context.close()

    if raw_har.exists():
        sanitize_har(raw_har, clean_har)
        raw_har.unlink(missing_ok=True)

    manifest = output_dir / "ЧТО_ВНУТРИ.txt"
    manifest.write_text(
        """Диагностическая запись myrosmol.ru\n\n"
        "playwright_trace.zip — пошаговая запись действий, DOM и скриншотов.\n"
        "network_sanitized.har — сетевые запросы и ответы без cookie/authorization-заголовков.\n"
        "api_responses — отдельно сохранённые ответы API.\n"
        "snapshots — HTML, PNG и URL каждого снимка, сделанного командой S.\n"
        "browser_events.log — ошибки JavaScript, консоль и неудачные запросы.\n"
        "\nВажно: перед отправкой можно открыть network_sanitized.har текстовым редактором и проверить, что там нет данных, которые вы не хотите передавать.\n"
        """,
        encoding="utf-8",
    )


def make_zip(output_dir: Path) -> Path:
    zip_path = output_dir.with_suffix(".zip")
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted(output_dir.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(output_dir.parent))
    return zip_path


async def async_main() -> int:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = OUTPUT_ROOT / f"myrosmol_diagnostic_{stamp}"
    output_dir.mkdir(parents=True, exist_ok=True)

    async with async_playwright() as playwright:
        await wait_for_manual_login(playwright)
        await record_scenario(playwright, output_dir)

    zip_path = make_zip(output_dir)
    print("\nДиагностика завершена.")
    print(f"Архив для отправки: {zip_path}")
    print("После загрузки архива в чат выйдите из аккаунта myrosmol.ru, чтобы завершить сохранённую сессию.")
    return 0


def main() -> None:
    try:
        raise SystemExit(asyncio.run(async_main()))
    except KeyboardInterrupt:
        print("\nОстановлено пользователем.")
        raise SystemExit(130)
    except Exception as exc:
        print(f"\nОШИБКА ДИАГНОСТИКИ: {exc}", file=sys.stderr)
        raise SystemExit(1)


if __name__ == "__main__":
    main()

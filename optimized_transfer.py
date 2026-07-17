from __future__ import annotations

"""Ускоренный запуск исходного transfer.py без удаления его функций.

Модуль добавляет ограниченный параллелизм для независимых аккаунтов и
неинтерактивные политики ошибок/дубликатов. В интерактивном режиме автоматически
используется один поток, поэтому все исходные вопросы и выборы сохраняются.
"""
import argparse
import asyncio
import os
import sys
from datetime import datetime
from typing import Any, Sequence

import transfer as core


def _extract_fast_args(argv: list[str]) -> tuple[argparse.Namespace, list[str]]:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--workers', type=int, default=1)
    parser.add_argument('--error-policy', choices=('ask','retry','skip','stop'), default='ask')
    parser.add_argument('--duplicate-policy', choices=('ask','existing','new'), default='ask')
    parser.add_argument('--fast-mode', action='store_true')
    known, rest = parser.parse_known_args(argv[1:])
    known.workers = max(1, min(4, known.workers))
    return known, [argv[0], *rest]

FAST, CLEAN_ARGV = _extract_fast_args(sys.argv)
sys.argv[:] = CLEAN_ARGV

_original_ask_error_action = core.ask_error_action
_original_choose_existing = core.choose_existing_project_or_new
_original_run_download = core.run_download
_original_run_upload = core.run_upload


async def ask_error_action_fast(step_name: str, error: Exception, *, skip_label: str='пропустить только этот шаг') -> str:
    policy = FAST.error_policy
    if policy == 'ask':
        return await _original_ask_error_action(step_name, error, skip_label=skip_label)
    print(f'\nОШИБКА НА ЭТАПЕ: {step_name}\nПричина: {error}')
    print(f'Автоматическая политика: {policy}')
    return policy


def _created_key(item: dict[str, Any]) -> str:
    value = str(item.get('createdAt') or item.get('created_at') or '')
    return value


async def choose_existing_fast(projects: Sequence[dict[str, Any]], title: str):
    if FAST.duplicate_policy == 'ask':
        return await _original_choose_existing(projects, title)
    matches = core.find_existing_projects_by_title(projects, title)
    if not matches:
        return 'new', None
    if FAST.duplicate_policy == 'new':
        print(f'Найден дубль «{title}»: по политике создаётся новый проект.')
        return 'new', None
    editable = [p for p in matches if core._existing_project_is_editable(p)]
    if not editable:
        print(f'Найден дубль «{title}», но редактируемых черновиков нет; создаётся новый.')
        return 'new', None
    selected = sorted(editable, key=_created_key, reverse=True)[0]
    print(f'Найден дубль «{title}»: выбран существующий черновик ID {selected.get("ID")}.')
    return 'existing', selected


async def _parallel_map(rows, worker, workers: int, stop_on_error: bool):
    sem = asyncio.Semaphore(workers)
    results = [None] * len(rows)
    stop_event = asyncio.Event()

    async def one(index, row):
        if stop_event.is_set():
            return
        async with sem:
            if stop_event.is_set():
                return
            result = await worker(row)
            results[index] = result
            print(f'[готово {index+1}/{len(rows)}] {row.fio}: {result.status}')
            if stop_on_error and result.status == 'ошибка':
                stop_event.set()

    await asyncio.gather(*(one(i, row) for i, row in enumerate(rows)))
    return [r for r in results if r is not None]


async def run_download_fast(browser, rows, settings, interactive_project_selection):
    workers = FAST.workers
    if interactive_project_selection and workers > 1:
        print('Интерактивный выбор проекта требует последовательного режима; workers=1.')
        workers = 1
    if workers == 1:
        return await _original_run_download(browser, rows, settings, interactive_project_selection)
    print(f'Параллельное скачивание: {workers} аккаунта(ов).')
    return await _parallel_map(
        rows,
        lambda account: core.process_download_account(browser, account, settings, interactive_project_selection),
        workers,
        settings.stop_on_error,
    )


async def run_upload_fast(browser, rows, settings):
    workers = FAST.workers
    if (FAST.error_policy == 'ask' or FAST.duplicate_policy == 'ask') and workers > 1:
        print('Интерактивные решения требуют последовательного режима; workers=1.')
        workers = 1
    if workers == 1:
        return await _original_run_upload(browser, rows, settings)
    print(f'Параллельный перенос: {workers} аккаунта(ов).')
    return await _parallel_map(
        rows,
        lambda account: core.process_upload_account(browser, account, settings),
        workers,
        settings.stop_on_error,
    )


core.ask_error_action = ask_error_action_fast
core.choose_existing_project_or_new = choose_existing_fast
core.run_download = run_download_fast
core.run_upload = run_upload_fast

if __name__ == '__main__':
    core.main()

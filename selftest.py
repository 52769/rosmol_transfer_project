from pathlib import Path
import ast, hashlib, tempfile, shutil
import transfer, legacy_transfer, optimized_transfer, gui

base=Path(__file__).resolve().parent

def defs(path):
    tree=ast.parse(path.read_text(encoding='utf-8'))
    return {n.name for n in tree.body if isinstance(n,(ast.FunctionDef,ast.AsyncFunctionDef,ast.ClassDef))}
orig=defs(base/'legacy_transfer.py'); current=defs(base/'transfer.py')
missing=orig-current
assert not missing, f'Потеряны определения: {sorted(missing)}'
assert hashlib.sha256((base/'legacy_transfer.py').read_bytes()).digest()==hashlib.sha256((base/'transfer.py').read_bytes()).digest(), 'Основное ядро изменено относительно legacy-копии'
settings=transfer.load_settings(base/'config_transfer.yaml')
assert settings.base_url.startswith('https://')
for name,reader in [('исходник.xls',transfer.read_source_accounts),('итог.xls',transfer.read_target_accounts)]:
    p=base/name
    if p.exists():
        rows=reader(p); print(f'{name}: {len(rows)} строк')
# Проверка парсера на первом DOCX, если он есть
for doc in (base/'скачанные_проекты').glob('*.docx'):
    project=transfer.parse_project_docx(doc,'')
    print(f'Парсер DOCX: {doc.name}; проект={project.title}; мероприятий={len(project.calendar)}; расходов={len(project.expenses)}')
    break
# Проверка отчёта без персональных данных
sample=transfer.TransferResult(row_number=1,fio='Тест',operation='самопроверка',project_name='Тестовый проект',status='успешно')
csv_path,xls_path=transfer.write_reports([sample],settings)
assert csv_path.exists() and xls_path.exists()
csv_path.unlink(); xls_path.unlink()

# UTF-8 transport and legacy console decoder
env, _ = gui._playwright_environment()
assert env.get("PYTHONUTF8") == "1"
assert env.get("PYTHONIOENCODING") == "utf-8"
sample_text = "Выберите проект: Отчёт XLS: C:\\Тест\\отчёт.xls\n"
for encoding in ("utf-8", "cp866", "cp1251"):
    decoded = gui.TransferGUI._decode_worker_output(sample_text.encode(encoding))
    assert decoded == sample_text, (encoding, decoded)
print("Кодировка UTF-8/CP866/CP1251: OK")
print(f'Определений сохранено: {len(orig)}')
print('Самопроверка пройдена.')

#!/usr/bin/env python3
"""Самопроверка репозитория vk-tmc-projects.

Запуск из корня репозитория:
    python3 tools/check_repo.py            # проверка, отчёт в консоль
    python3 tools/check_repo.py --index    # проверка + пересборка INDEX_НД.md

Что проверяется:
  1. Ссылки на файлы в инструкциях (context.md), чек-листах, правилах и README
     проектов — указанный файл существует где-нибудь в репозитории.
  2. Пары «текстовая версия НД (.md) — оригинал»: у каждого текста НД в docs/
     есть файл с тем же именем в originals/ и наоборот.
  3. originals/README.md каждого проекта перечисляет все файлы папки originals/.
  4. В корневом README.md перечислены все папки проектов.

Код выхода: 0 — замечаний нет, 1 — есть замечания.
Исключения (заведомо несуществующие имена-примеры) — в tools/check_repo_ignore.txt.
"""
import os
import re
import sys
import unicodedata
import urllib.parse
from collections import defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)

ND_PREFIX = re.compile(r'^(ГОСТ|ОСТ|ТУ|ТР ТС|ТР_ТС|АТК|ФНП|РД|ПБ|СТО|СНиП|СП|ПНАЭ|РТМ)[ _]')
REF = re.compile(r'(?<![\w/])([\wА-Яа-яЁё][\wА-Яа-яЁё.\-]*\.(?:md|xlsx|docx|py|txt))(?![\wА-Яа-яЁё])')
SKIP_DIRS = {'.git', 'tools'}
SKIP_SECTION = re.compile(r'ссыл|Чего нет|Отсутству|которых нет|Замечания по составу|Об исходных файлах', re.I)


def nfc(s):
    return unicodedata.normalize('NFC', s)


def projects():
    return sorted(d for d in os.listdir('.') if os.path.isdir(d) and d not in SKIP_DIRS
                  and not d.startswith('.') and os.path.isfile(os.path.join(d, 'context.md')))


def all_files():
    out = []
    for dp, dn, fn in os.walk('.'):
        dn[:] = [d for d in dn if d not in SKIP_DIRS]
        for f in fn:
            out.append(nfc(os.path.relpath(os.path.join(dp, f), '.')))
    return out


def is_nd_name(name):
    return bool(ND_PREFIX.match(nfc(name)))


def load_ignore():
    p = os.path.join('tools', 'check_repo_ignore.txt')
    ign = set()
    if os.path.isfile(p):
        for line in open(p, encoding='utf-8'):
            line = line.split('#', 1)[0].strip()
            if line:
                ign.add(nfc(line))
    return ign


def stem(name):
    return name.rsplit('.', 1)[0]


def main():
    make_index = '--index' in sys.argv
    files = all_files()
    names = {os.path.basename(f) for f in files}
    stems = {stem(n) for n in names}
    ignore = load_ignore()
    problems = defaultdict(list)
    projs = projects()

    # 1. Ссылки на файлы
    for p in projs:
        srcs = []
        for f in files:
            if not f.startswith(p + '/') or not f.endswith('.md'):
                continue
            base = os.path.basename(f)
            if is_nd_name(base) or base.startswith(('Шаблон_', 'Образец_', 'КРК_')) or '/originals/' in f:
                continue  # тексты НД и образцы цитируют чужие имена файлов
            if base == 'CHANGELOG.md':
                continue  # история законно упоминает удалённые файлы
            srcs.append(f)
        for f in srcs:
            section = ''
            for n, line in enumerate(open(f, encoding='utf-8'), 1):
                line = nfc(urllib.parse.unquote(line))
                if line.startswith('## ') or line.startswith('# '):
                    section = line
                # разделы-отчёты, которые нарочно перечисляют отсутствующие файлы
                if SKIP_SECTION.search(section):
                    continue
                for m in REF.finditer(line):
                    r = m.group(1).rstrip('.')
                    if r in ignore or r in names or stem(r) in stems:
                        continue
                    # имя с пробелом («ГОСТ 17375-2001.md») регулярка видит с середины
                    if any(x.endswith(r) for x in names) or any(x.endswith(stem(r)) for x in stems):
                        continue
                    # «имя_docx.docx» → в репозитории есть «имя_docx.md» или «имя.docx»
                    s = stem(r)
                    if s.endswith(('_docx', '_xlsx')) and (s in stems or s[:-5] in stems):
                        continue
                    problems['Ссылка на несуществующий файл'].append(f'{f}:{n} → {r}')

    # 2. Пары текст НД ↔ оригинал
    for p in projs:
        docs = [f for f in files if f.startswith(p + '/docs/') and f.endswith('.md') and is_nd_name(os.path.basename(f))]
        origs = [f for f in files if f.startswith(p + '/originals/') and is_nd_name(os.path.basename(f))]
        ost = {stem(os.path.basename(f)) for f in origs}
        dst = {stem(os.path.basename(f)) for f in docs}
        for f in docs:
            if stem(os.path.basename(f)) not in ost:
                problems['Текст НД без оригинала в originals/'].append(f)
        for f in origs:
            if stem(os.path.basename(f)) not in dst:
                problems['Оригинал НД без текстовой версии в docs/'].append(f)

    # 3. originals/README.md перечисляет все файлы
    for p in projs:
        od = os.path.join(p, 'originals')
        rd = os.path.join(od, 'README.md')
        if not os.path.isdir(od):
            continue
        if not os.path.isfile(rd):
            problems['Нет originals/README.md'].append(p)
            continue
        txt = nfc(open(rd, encoding='utf-8').read())
        for f in sorted(os.listdir(od)):
            if f != 'README.md' and nfc(f) not in txt:
                problems['Файл не указан в originals/README.md'].append(f'{p}/originals/{nfc(f)}')

    # 4. Корневой README перечисляет все проекты
    root_readme = nfc(open('README.md', encoding='utf-8').read())
    for p in projs:
        if f'`{p}`' not in root_readme:
            problems['Проект не указан в корневом README.md'].append(p)

    if make_index:
        write_index(files, projs)

    total = sum(len(v) for v in problems.values())
    print(f'Проектов: {len(projs)} ({", ".join(projs)}); файлов: {len(files)}')
    if not total:
        print('Замечаний нет.')
        return 0
    for k, v in problems.items():
        print(f'\n## {k} ({len(v)})')
        for x in v:
            print('  -', x)
    print(f'\nИтого замечаний: {total}')
    return 1


def nd_key(name):
    s = stem(nfc(name)).replace('_', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def write_index(files, projs):
    groups = defaultdict(lambda: {'pdf': [], 'md': []})
    for f in files:
        base = os.path.basename(f)
        if not is_nd_name(base):
            continue
        proj = f.split('/', 1)[0]
        if proj not in projs:
            continue
        key = nd_key(base)
        if '/originals/' in f and not f.endswith('.md'):
            groups[key]['pdf'].append(f)
        elif '/docs/' in f and f.endswith('.md'):
            groups[key]['md'].append(f)

    def kind(k):
        for pref in ['ГОСТ', 'ОСТ', 'ТУ', 'ТР ТС', 'ФНП', 'АТК']:
            if k.startswith(pref + ' '):
                return pref
        return 'Прочие'

    def sort_key(k):
        nums = re.findall(r'\d+', k)
        return (int(nums[0]) if nums else 0, k)

    order = ['ГОСТ', 'ОСТ', 'ТУ', 'ТР ТС', 'ФНП', 'АТК', 'Прочие']
    by_kind = defaultdict(list)
    for k in groups:
        by_kind[kind(k)].append(k)

    L = ['# Указатель НД во всех проектах репозитория',
         '',
         '> Файл собирается автоматически: `python3 tools/check_repo.py --index`. Вручную не править — пересобрать после добавления или удаления НД.',
         '> Для каждого документа — где лежит оригинал (PDF) и текстовая версия (.md). Нормы брать по оригиналу; текстовая версия — для поиска.',
         '> Один и тот же НД может лежать в нескольких проектах — это копии одного файла; открывать любую.',
         '']
    n = 0
    for kd in order:
        ks = sorted(by_kind.get(kd, []), key=sort_key)
        if not ks:
            continue
        L += [f'## {kd} ({len(ks)})', '', '| НД | Оригинал (PDF) | Текстовая версия (.md) |', '|---|---|---|']
        for k in ks:
            g = groups[k]
            pdf = '<br>'.join(f'`{x}`' for x in sorted(g['pdf'])) or '—'
            md = '<br>'.join(f'`{x}`' for x in sorted(g['md'])) or '—'
            L.append(f'| {k} | {pdf} | {md} |')
            n += 1
        L.append('')
    L.insert(5, f'Всего документов: **{n}**.\n')
    open('INDEX_НД.md', 'w', encoding='utf-8').write('\n'.join(L))
    print(f'INDEX_НД.md пересобран: {n} документов')


if __name__ == '__main__':
    sys.exit(main())

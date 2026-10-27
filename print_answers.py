"""Печатные файлы «с ответами» для психологии — как у БЖД.

Правильный ответ: ✔ + жирный + рамка (под ч/б принтер), неправильный: ✗ обычным.
Оформление (поля, шрифт, стили) берётся из готового docx БЖД — шаблон по умолчанию
«Загрузки/Блок В — с ответами.docx»; из него заменяется только текст.

    python print_answers.py [--only-correct] [шаблон.docx]

Итог: Загрузки/Психология уровень N — с ответами.docx (+ .pdf, если есть Word).
С --only-correct — «… — только правильные.docx»: под вопросом остаются одни верные ответы.
"""
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parent
DOWNLOADS = Path.home() / "Downloads"
DEFAULT_TEMPLATE = DOWNLOADS / "Блок В — с ответами.docx"
LETTERS = "абвгдежзик"

LEVELS = [
    ("level1.json", "Психология — уровень 1", "Психология · уровень 1",
     "в каждом один правильный ответ", False),
    ("level2.json", "Психология — уровень 2", "Психология · уровень 2",
     "в каждом несколько правильных ответов", True),
]

SYM = ('<w:rFonts w:ascii="Segoe UI Symbol" w:cs="Segoe UI Symbol" '
       'w:eastAsia="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol"/>')
BOX = ('<w:pBdr>' + ''.join(
    f'<w:{s} w:val="single" w:color="000000" w:sz="10" w:space="2"/>'
    for s in ("top", "bottom", "left", "right")) + '</w:pBdr>')
BOLD = '<w:b/><w:bCs/>'
BLACK = '<w:color w:val="000000"/>'


def run(text, rpr=""):
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def plural(n, one, few, many):
    if n % 10 == 1 and n % 100 != 11:
        return one
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return few
    return many


def answer_par(text, ok, keep_next, after=30, border=None):
    keep = '<w:keepNext/>' if keep_next else '<w:keepNext w:val="false"/>'
    border = (BOX if ok else "") if border is None else border
    rpr = BOLD if ok else BLACK
    return (f'<w:p><w:pPr>{keep}<w:keepLines/>{border}<w:spacing w:after="{after}"/>'
            f'<w:ind w:left="680" w:hanging="340"/></w:pPr>'
            + run("✔\t" if ok else "✗\t", SYM + rpr) + run(text, rpr) + '</w:p>')


def correct_par(text, keep_next):
    keep = '<w:keepNext/>' if keep_next else '<w:keepNext w:val="false"/>'
    return (f'<w:p><w:pPr>{keep}<w:keepLines/><w:spacing w:after="30"/>'
            f'<w:ind w:left="680" w:hanging="340"/></w:pPr>'
            + run("✔	", SYM + BOLD) + run(text, BOLD) + '</w:p>')


def question_xml(num, q, multi, only_correct=False):
    text = " ".join(q["q"])
    out = [('<w:p><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="60"/>'
            '<w:ind w:left="340" w:hanging="340"/></w:pPr>'
            + run(f"{num}. ", BOLD) + run(text, BOLD) + '</w:p>')]
    if only_correct:
        ok = [a for a in q["a"] if a["ok"]]
        out += [correct_par(a["t"], i < len(ok) - 1) for i, a in enumerate(ok)]
        return "".join(out)
    if multi:
        k, n = sum(a["ok"] for a in q["a"]), len(q["a"])
        note = f'Правильных {k} {plural(k, "ответ", "ответа", "ответов")} из {n}'
        out.append('<w:p><w:pPr><w:keepNext/><w:spacing w:after="60"/><w:ind w:left="340"/></w:pPr>'
                   + run(note, '<w:i/><w:iCs/>' + BLACK + '<w:sz w:val="19"/><w:szCs w:val="19"/>')
                   + '</w:p>')
    last = len(q["a"]) - 1
    for i, a in enumerate(q["a"]):
        out.append(answer_par(f'{LETTERS[i]}) {a["t"]}', a["ok"], i < last))
    return "".join(out)


def body_xml(data, title, subtitle, only_correct=False):
    n = len(data)
    head = [
        '<w:p><w:pPr><w:spacing w:after="80"/><w:jc w:val="center"/></w:pPr>'
        + run(title, BOLD + '<w:sz w:val="32"/><w:szCs w:val="32"/>') + '</w:p>',
        '<w:p><w:pPr><w:spacing w:after="160"/><w:jc w:val="center"/></w:pPr>'
        + run(f'{n} {plural(n, "вопрос", "вопроса", "вопросов")} · {subtitle}',
              BLACK + '<w:sz w:val="22"/><w:szCs w:val="22"/>') + '</w:p>',
    ]
    if only_correct:
        head.append('<w:p><w:pPr><w:spacing w:after="120"/><w:jc w:val="center"/>'
                    '<w:pBdr><w:bottom w:val="single" w:color="999999" w:sz="6" w:space="8"/></w:pBdr></w:pPr>'
                    + run("Показаны только правильные ответы", '<w:i/><w:iCs/>' + BLACK) + '</w:p>')
        return "".join(head)
    head += [
        answer_par("Правильный ответ — галочка, жирный шрифт, рамка", True, False, after=40),
        answer_par("Неправильный ответ — крестик, обычный шрифт", False, False, after=120,
                   border='<w:pBdr><w:bottom w:val="single" w:color="999999" w:sz="6" w:space="8"/></w:pBdr>'),
    ]
    return "".join(head)


def build_docx(template, out_path, data, title, footer, subtitle, multi, only_correct=False):
    with zipfile.ZipFile(template) as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    doc = files["word/document.xml"].decode("utf-8")
    start = doc.index("<w:body>") + len("<w:body>")
    end = doc.rindex("<w:sectPr")
    content = body_xml(data, title, subtitle, only_correct) + "".join(
        question_xml(i + 1, q, multi, only_correct) for i, q in enumerate(data))
    files["word/document.xml"] = (doc[:start] + content + doc[end:]).encode("utf-8")

    ftr = files["word/footer1.xml"].decode("utf-8")
    ftr, n = re.subn(r'>[^<>]*· стр\. <', f'>{escape(footer)} · стр. <', ftr)
    if n != 1:
        raise SystemExit("В шаблоне не найден колонтитул «… · стр.»")
    files["word/footer1.xml"] = ftr.encode("utf-8")

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in files.items():
            zout.writestr(name, blob)


def to_pdf(docx_paths):
    """PDF через установленный Word (COM). Без Word — пропуск с сообщением."""
    pairs = ";".join(f"'{p}','{p.with_suffix('.pdf')}'" for p in docx_paths)
    ps = (
        "$ErrorActionPreference='Stop'; $w=New-Object -ComObject Word.Application; $w.Visible=$false; "
        f"$l=@({pairs}); try {{ for($i=0;$i -lt $l.Count;$i+=2){{ "
        "$d=$w.Documents.Open($l[$i],$false,$true); $d.Fields.Update()|Out-Null; "
        "$d.ExportAsFixedFormat($l[$i+1],17); $d.Close($false) } } finally { $w.Quit() }"
    )
    res = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                         capture_output=True, text=True)
    if res.returncode != 0:
        print("PDF не сделан (нужен Word):", res.stderr.strip()[:300])


def main():
    args = sys.argv[1:]
    only_correct = "--only-correct" in args
    args = [a for a in args if a != "--only-correct"]
    template = Path(args[0]) if args else DEFAULT_TEMPLATE
    suffix = "только правильные" if only_correct else "с ответами"
    if not template.exists():
        raise SystemExit(f"Нет шаблона: {template}")
    made = []
    for fname, title, footer, subtitle, multi in LEVELS:
        data = json.loads((ROOT / "data" / fname).read_text(encoding="utf-8"))
        out = DOWNLOADS / f"{title.replace(' — ', ' ')} — {suffix}.docx"
        build_docx(template, out, data, title, footer, subtitle, multi, only_correct)
        made.append(out)
        print("docx:", out)
    to_pdf(made)


if __name__ == "__main__":
    main()

"""Собирает index.html: вшивает вопросы из data/*.json в template.html."""
import json
from pathlib import Path

ROOT = Path(__file__).parent


def compact(path):
    questions = json.loads((ROOT / path).read_text(encoding='utf8'))
    return [{'q': q['q'], 'a': [{'t': a['t'], 'ok': bool(a['ok'])} for a in q['a']]} for q in questions]


data = {'A': compact('data/level1.json'), 'V': compact('data/level2.json')}
payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
html = (ROOT / 'template.html').read_text(encoding='utf8').replace('/*DATA*/', payload)
(ROOT / 'index.html').write_text(html, encoding='utf8')
print('index.html:', len(data['A']), '+', len(data['V']), 'вопросов,', len(html.encode()) // 1024, 'КБ')

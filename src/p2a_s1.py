"""Opt-in S1 trial: replace only the ordinary duplicate sheet and trial metadata.

No shared/array formulas, caches, account optimization or price changes.
The accepted p2a_excel generator is used unchanged and remains independently usable.
"""
from pathlib import Path
import hashlib
import json
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET

from . import p2a_excel as baseline
from .p2a import from_p1
from .lifeplan import demo

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'m': MAIN}
VARIANT = 'P2A-S1-prototype-v1'
ET.register_namespace('', MAIN)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def base_source_id():
    return baseline.identity(from_p1(demo()))[1]


def source_id():
    return hashlib.sha256((VARIANT + '\n' + base_source_id()).encode() + Path(__file__).read_text(encoding='utf-8').encode()).hexdigest()


def identity(plan):
    return baseline.identity(plan)[0], source_id()


def destination(path, source=None):
    p = Path(path).resolve()
    if p.is_relative_to(Path(__file__).resolve().parents[1] / 'references'):
        raise ValueError('参照原本へ出力しません')
    if p.exists() or (source is not None and p == Path(source).resolve()):
        raise FileExistsError('既存ファイルを上書きしません: ' + str(p))
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def sheet_members(z):
    rels = {r.attrib['Id']: r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    result = {}
    for s in ET.fromstring(z.read('xl/workbook.xml')).find('m:sheets', NS):
        target = rels[s.attrib['{' + REL + '}id']]
        result[s.attrib['name']] = target.lstrip('/') if target.startswith('/') else 'xl/' + target
    return result


def stats(path):
    result = dict(bytes=Path(path).stat().st_size, formulas=0, shared=0, arrays=0, caches=0, max_chars=0, sheets={})
    with zipfile.ZipFile(path) as z:
        for name, member in sheet_members(z).items():
            n = chars = longest = shared = arrays = caches = 0
            with z.open(member) as stream:
                for _, c in ET.iterparse(stream, events=('end',)):
                    if c.tag == '{' + MAIN + '}c':
                        f, v = c.find('m:f', NS), c.find('m:v', NS)
                        if f is not None:
                            n += 1
                            length = len(f.text or '') + 1
                            chars += length
                            longest = max(longest, length)
                            shared += int(f.attrib.get('t') == 'shared')
                            arrays += int(f.attrib.get('t') == 'array')
                            caches += int(v is not None and bool(v.text))
                        c.clear()
            info = z.getinfo(member)
            result['sheets'][name] = dict(formulas=n, formula_chars=chars, max_chars=longest,
                compressed_bytes=info.compress_size, xml_bytes=info.file_size)
            for key, value in [('formulas', n), ('shared', shared), ('arrays', arrays), ('caches', caches)]:
                result[key] += value
            result['max_chars'] = max(result['max_chars'], longest)
    return result


def numeric_cell(address, value):
    c = ET.Element('{' + MAIN + '}c', r=address, t='n')
    ET.SubElement(c, '{' + MAIN + '}v').text = str(value)
    return c


def text_cell(address, value):
    c = ET.Element('{' + MAIN + '}c', r=address, t='inlineStr')
    ET.SubElement(ET.SubElement(c, '{' + MAIN + '}is'), '{' + MAIN + '}t').text = value
    return c


def text_value(root, address):
    c = root.find(f'.//m:c[@r="{address}"]', NS)
    if c is None:
        return None
    return ''.join(t.text or '' for t in c.findall('.//m:t', NS))


def duplicate_formula(matrix_row):
    return f'=IF(SUMPRODUCT(--(FREQUENCY($B$2:$IQ$2*--(\'月次計上\'!B{matrix_row}:IQ{matrix_row}>0),$IT$3:$IT$253)>1),$IU$3:$IU$254)>0,1,0)'


def rewrite_duplicates(data, months):
    root = ET.fromstring(data)
    body = root.find('m:sheetData', NS)
    rows = {int(r.attrib['r']): r for r in body}
    if text_value(root, 'A2') != '前の同項目列':
        raise ValueError('基準の累積表ではありません')
    removed = groups = flags = 0
    for row_number, row in rows.items():
        for c in list(row):
            address = c.attrib['r']
            column = re.match(r'[A-Z]+', address)[0]
            number = 0
            for ch in column:
                number = number * 26 + ord(ch) - 64
            f = c.find('m:f', NS)
            if 2 <= number <= 251:
                if row_number == 2:
                    assert f is not None
                    f.text = f'IF({column}1="",0,MATCH({column}1,$B$1:$IQ$1,0))'
                    groups += 1
                elif row_number >= 3:
                    assert f is not None and 3 <= row_number <= months + 2
                    row.remove(c)
                    removed += 1
            elif number == 252 and row_number >= 3:
                assert f is not None and 3 <= row_number <= months + 2
                f.text = duplicate_formula(row_number + 4)[1:]
                flags += 1
    assert (groups, removed, flags) == (250, 250 * months, months)
    old = root.find('.//m:c[@r="A2"]', NS)
    rows[2].remove(old)
    rows[2].insert(0, text_cell('A2', '数値グループID（初出位置）'))
    rows[1].append(text_cell('IT1', '整数ビン上限'))
    rows[1].append(text_cell('IU1', '0ビン除外マスク'))
    for offset in range(252):
        r = offset + 3
        if r not in rows:
            rows[r] = ET.Element('{' + MAIN + '}row', r=str(r))
        if offset <= 250:
            rows[r].append(numeric_cell(f'IT{r}', offset))
        rows[r].append(numeric_cell(f'IU{r}', int(offset != 0)))
    body[:] = [rows[r] for r in sorted(rows)]
    root.find('m:dimension', NS).set('ref', f'A1:IU{max(rows)}')
    return ET.tostring(root, encoding='utf-8')


def rewrite_meta(data, baseline_sha):
    root = ET.fromstring(data)
    if text_value(root, 'B2') != base_source_id():
        raise ValueError('ffe49efと同じ通常式生成コードの基準が必要です')
    body = root.find('m:sheetData', NS)
    row2 = body.find('m:row[@r="2"]', NS)
    row2.remove(row2.find('m:c[@r="B2"]', NS))
    row2.append(text_cell('B2', source_id()))
    for row, label, value in [(3, '試作方式', VARIANT), (4, '基準生成コードSHA256', base_source_id()), (5, '基準xlsx SHA256', baseline_sha)]:
        r = ET.SubElement(body, '{' + MAIN + '}row', r=str(row))
        r.extend([text_cell('A' + str(row), label), text_cell('B' + str(row), value)])
    root.find('m:dimension', NS).set('ref', 'A1:B5')
    return ET.tostring(root, encoding='utf-8')


def transform(original, output):
    """Preserve source and all non-target XML bytes. Only ordinary formulas used."""
    output = destination(output, original)
    before_hash = sha(original)
    before = stats(original)
    if before['shared'] or before['arrays'] or before['caches']:
        raise ValueError('キャッシュ空の通常式生成物を指定してください')
    n = (before['sheets']['重複確認']['formulas'] - 500) // 252
    if not 1 <= n <= 1200 or before['sheets']['重複確認']['formulas'] != 252 * n + 500:
        raise ValueError('250ソース・1200か月以内の基準だけに対応します')
    with zipfile.ZipFile(original) as source:
        members = sheet_members(source)
        edits = {members['重複確認']: rewrite_duplicates(source.read(members['重複確認']), n),
                 members['_検証メタ']: rewrite_meta(source.read(members['_検証メタ']), before_hash)}
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix='.xlsx', delete=False) as f:
            temporary = Path(f.name)
        try:
            with zipfile.ZipFile(temporary, 'w', compression=zipfile.ZIP_DEFLATED) as target:
                for entry in source.infolist():
                    target.writestr(entry, edits[entry.filename] if entry.filename in edits else source.read(entry.filename))
            temporary.replace(output)
        finally:
            temporary.unlink(missing_ok=True)
    assert sha(original) == before_hash
    after = stats(output)
    assert after['formulas'] == before['formulas'] - 250 * n
    assert after['shared'] == after['arrays'] == after['caches'] == 0
    assert after['max_chars'] <= 8192
    return dict(months=n, original_sha256=before_hash, output_sha256=sha(output), before=before, after=after)


def generate(plan, output):
    destination(output)
    with tempfile.TemporaryDirectory(prefix='p2a-s1-baseline-') as temporary:
        original = Path(temporary) / 'normal.xlsx'
        w = baseline.generate(plan, original)
        w.close()
        del w
        return transform(original, output)


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--baseline', required=True)
    parser.add_argument('--output', required=True)
    args = parser.parse_args()
    print(json.dumps(transform(args.baseline, args.output), ensure_ascii=False, indent=2))

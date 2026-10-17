"""Independent bounded input checker. Never calls the converter or writes XLSX."""
import hashlib
import io
from pathlib import Path
import stat
import struct
import time
import zipfile
import xml.etree.ElementTree as ET

from .common import GateError

ORIGINAL_SHA = 'ea158ca2052c1c2ba4d5b7bd30e4304846770e82392bea4229506d973f131e21'
WORKBOOK_SHA = '2a22873f10491ef5f43000f6fb3bec8d1d10781214762f37585a8c3427acc5a8'
WORKBOOK = 'xl/workbook.xml'
FILE = '01_function_gate.xlsx'
MAX_ARCHIVE = 131072
MAX_EXPANDED = 45235
PARTS = ('docProps/app.xml', 'docProps/core.xml', 'xl/theme/theme1.xml',
         'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml',
         'xl/styles.xml', '_rels/.rels', WORKBOOK, 'xl/_rels/workbook.xml.rels', '[Content_Types].xml')
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def digest(data):
    return hashlib.sha256(data).hexdigest()


def read_bounded(path):
    path = Path(path)
    # Existing parent links and reparse points are also refused.
    for item in (path, *path.parents):
        s = item.lstat()
        if stat.S_ISLNK(s.st_mode) or getattr(s, 'st_file_attributes', 0) & 0x400:
            raise ValueError('リンク/reparse pointを入力に使えません')
    s = path.stat()
    if not stat.S_ISREG(s.st_mode) or s.st_nlink != 1 or s.st_size > MAX_ARCHIVE:
        raise ValueError('通常の独立ファイル/ZIPサイズ制限に不一致')
    with path.open('rb') as stream:
        raw = stream.read(MAX_ARCHIVE + 1)
    if len(raw) > MAX_ARCHIVE:
        raise ValueError('ZIP読み取りサイズ超過')
    return raw


def package(raw):
    """Closed 11-part layout; bounded reads, central/local headers, CRC/XML."""
    started = time.monotonic()
    if len(raw) > MAX_ARCHIVE:
        raise ValueError('ZIPサイズ超過')
    payloads, metadata = {}, []
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        infos = z.infolist()
        if tuple(i.filename for i in infos) != PARTS or z.comment:
            raise ValueError('ZIP part集合/順序/コメント不一致（重複/別名も不可）')
        if sum(i.file_size for i in infos) > MAX_EXPANDED:
            raise ValueError('ZIP展開サイズ超過')
        # No ZIP64, multi-disk, or appended records for this fixed fixture.
        end = raw[-22:]
        if len(end) != 22 or end[:4] != b'PK\x05\x06':
            raise ValueError('ZIP終端レコード不一致')
        _, disk, central_disk, n_disk, n, size, offset, comment = struct.unpack('<4s4H2LH', end)
        if disk or central_disk or n_disk != 11 or n != 11 or comment or offset + size != len(raw) - 22:
            raise ValueError('ZIP分割/ZIP64/追加レコードを許可しません')
        cursor = 0
        for i in infos:
            mode = i.external_attr >> 16
            if (i.flag_bits != 0 or i.compress_type != zipfile.ZIP_DEFLATED or i.extra or i.comment
                    or i.volume or i.reserved or i.orig_filename != i.filename
                    or i.create_system not in (0, 3) or stat.S_IFMT(mode) not in (0, stat.S_IFREG)
                    or i.file_size < 0 or i.file_size > MAX_EXPANDED or i.header_offset != cursor):
                raise ValueError('ZIP entry/暗号化/特殊entry/headerに想定外値')
            header = raw[cursor:cursor+30]
            if len(header) != 30:
                raise ValueError('ZIP local header欠落')
            sig, version, flags, method, mt, md, crc, compressed, expanded, nl, el = struct.unpack('<4s5H3L2H', header)
            year, month, day, hour, minute, second = i.date_time
            dos_time = (hour << 11) | (minute << 5) | (second // 2)
            dos_date = ((year - 1980) << 9) | (month << 5) | day
            if (sig != b'PK\x03\x04' or version != i.extract_version or flags != i.flag_bits
                    or method != i.compress_type or crc != i.CRC or compressed != i.compress_size
                    or expanded != i.file_size or mt != dos_time or md != dos_date
                    or el or raw[cursor+30:cursor+30+nl] != i.filename.encode('ascii')):
                raise ValueError('ZIP local/central header不整合')
            cursor += 30 + nl + el + compressed
            if cursor > offset:
                raise ValueError('ZIP payload範囲不整合')
            with z.open(i) as stream:
                content = stream.read(i.file_size + 1)
            if len(content) != i.file_size:
                raise ValueError('ZIP展開byte数不一致')
            if b'\x00' in content or b'<!DOCTYPE' in content.upper() or b'<!ENTITY' in content.upper():
                raise ValueError('DTD/entity禁止')
            ET.fromstring(content)  # every part in this fixed fixture is XML
            payloads[i.filename] = content
            metadata.append(dict(name=i.filename, bytes=i.file_size, compression=i.compress_type,
                                 timestamp=list(i.date_time), attributes=i.external_attr,
                                 creator=i.create_system, create_version=i.create_version,
                                 extract_version=i.extract_version, internal_attr=i.internal_attr))
            if time.monotonic() - started > 10:
                raise GateError('TIMEOUT', 'ZIP静的検査10秒超過')
        if cursor != offset:
            raise ValueError('ZIP part間の余分なデータ')
    return payloads, metadata


def verify(original, contrast):
    """No converter helpers: byte diff + XML tree + all part payloads."""
    before_raw, after_raw = read_bounded(original), read_bounded(contrast)
    if digest(before_raw) != ORIGINAL_SHA:
        raise ValueError('固定原本SHA不一致')
    before, bm = package(before_raw)
    after, am = package(after_raw)
    if bm != am:
        raise ValueError('ZIP metadata差分（payload以外）の不一致')
    if digest(before[WORKBOOK]) != WORKBOOK_SHA:
        raise ValueError('固定workbook.xml SHA不一致')
    left, right = before[WORKBOOK], after[WORKBOOK]
    changed = [i for i, (a, b) in enumerate(zip(left, right)) if a != b]
    if len(left) != 705 or len(left) != len(right) or len(changed) != 1:
        raise ValueError('workbook.xmlは厳密に1byte差だけを許可')
    index = changed[0]
    if left[index] != ord('1') or right[index] != ord('0'):
        raise ValueError('変更byteは1から0だけ')
    lt, rt = ET.fromstring(left), ET.fromstring(right)
    lcs, rcs = lt.findall('m:calcPr', NS), rt.findall('m:calcPr', NS)
    if len(lcs) != 1 or len(rcs) != 1:
        raise ValueError('calcPrの一意性違反')
    if lcs[0].attrib != dict(calcId='0', calcMode='auto', fullCalcOnLoad='1', forceFullCalc='1'):
        raise ValueError('原本calcPr想定外')
    if rcs[0].attrib != dict(calcId='0', calcMode='auto', fullCalcOnLoad='1', forceFullCalc='0'):
        raise ValueError('対照calcPr想定外')
    # This is a checker-only normalization, not the builder's replacement logic.
    rcs[0].attrib['forceFullCalc'] = '1'
    if ET.tostring(lt) != ET.tostring(rt):
        raise ValueError('対象属性以外のXML木変更')
    entries = []
    for name in PARTS:
        if name != WORKBOOK and before[name] != after[name]:
            raise ValueError('対象外partのbyte変更: ' + name)
        entries.append(dict(part=name, before_sha256=digest(before[name]), after_sha256=digest(after[name]),
                            bytes=len(before[name]), changed=name == WORKBOOK))
    # Check formulas/cache independently of the converter, without Excel eval.
    formulas = []
    for name in PARTS:
        if name.startswith('xl/worksheets/'):
            for c in ET.fromstring(after[name]).findall('.//m:c', NS):
                f, v = c.find('m:f', NS), c.find('m:v', NS)
                if f is not None:
                    if f.attrib or not f.text or len(f.text)+1 > 97 or v is not None and v.text:
                        raise ValueError('通常式/式長/空キャッシュ不一致')
                    formulas.append((name, c.attrib['r'], f.text))
    if len(formulas) != 24:
        raise ValueError('24式以外を許可しません')
    return dict(original_sha256=ORIGINAL_SHA, input_sha256=digest(after_raw),
                workbook_before_sha256=WORKBOOK_SHA, workbook_after_sha256=digest(right),
                change=dict(part=WORKBOOK, path='/workbook/calcPr/@forceFullCalc',
                            before='1', after='0', byte_offset=index),
                preserved_attributes=dict(calcId='0', calcMode='auto', fullCalcOnLoad='1'),
                parts=entries, formula_count=24, nonempty_caches=0,
                formula_payload_sha256=digest(repr(formulas).encode('utf-8')))

#!/usr/bin/env python3
"""Independently compare DOCX bold/italic with raw HWP or HWPX character records.

Requires Python olefile for HWP. Reports whitespace-free text alignment separately
from styling so inserted/deleted text cannot silently count as a full match.
"""
import argparse
from collections import Counter
from difflib import SequenceMatcher
import hashlib
import json
from pathlib import Path
import struct
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import zlib

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
HH = '{http://www.hancom.co.kr/hwpml/2011/head}'
HP = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'
OFF = {'0', 'false', 'off'}


def on(node):
    return node is not None and node.get(W + 'val', 'true').lower() not in OFF


def glyphs(text, bold, italic):
    return [(c, bool(bold), bool(italic)) for c in text if not c.isspace()]


def docx_glyphs(path):
    with ZipFile(path) as zipfile:
        document = ET.fromstring(zipfile.read('word/document.xml'))
        styles = ET.fromstring(zipfile.read('word/styles.xml'))
    definitions = {s.get(W + 'styleId'): s for s in styles.findall(W + 'style')}
    defaults = styles.find(f'{W}docDefaults/{W}rPrDefault/{W}rPr')
    default_paragraph = next((sid for sid, s in definitions.items()
                              if s.get(W + 'type') == 'paragraph' and s.get(W + 'default', '').lower() in {'1', 'true', 'on'}), None)

    def apply_style(sid, state):
        chain, seen = [], set()
        while sid in definitions and sid not in seen:
            seen.add(sid)
            definition = definitions[sid]
            chain.append(definition)
            parent = definition.find(W + 'basedOn')
            sid = parent.get(W + 'val') if parent is not None else None
        for definition in reversed(chain):
            for flag in ['b', 'i']:
                # Style-level true toggles; false leaves the inherited value.
                if on(definition.find(f'{W}rPr/{W}{flag}')):
                    state[flag] = not state[flag]
        return state

    result = []
    for paragraph in document.iter(W + 'p'):
        style = paragraph.find(f'{W}pPr/{W}pStyle')
        state = {flag: on(defaults.find(W + flag)) if defaults is not None else False for flag in ['b', 'i']}
        state = apply_style(style.get(W + 'val') if style is not None else default_paragraph, state)
        for run in paragraph.iter(W + 'r'):
            effective = state.copy()
            rpr = run.find(W + 'rPr')
            if rpr is not None:
                char_style = rpr.find(W + 'rStyle')
                if char_style is not None:
                    effective = apply_style(char_style.get(W + 'val'), effective)
                for flag in ['b', 'i']:
                    node = rpr.find(W + flag)
                    if node is not None:
                        effective[flag] = on(node)
            for text in run.findall(W + 't'):
                result += glyphs(text.text or '', effective['b'], effective['i'])
    return result


def hwpx_glyphs(path):
    with ZipFile(path) as zipfile:
        header = ET.fromstring(zipfile.read('Contents/header.xml'))
        shapes = {s.get('id'): (s.find(HH + 'bold') is not None, s.find(HH + 'italic') is not None)
                  for s in header.iter(HH + 'charPr')}
        names = sorted((name for name in zipfile.namelist() if name.startswith('Contents/section') and name.endswith('.xml')),
                       key=lambda name: int(Path(name).stem.replace('section', '')))
        result = []
        for name in names:
            section = ET.fromstring(zipfile.read(name))
            for run in section.iter(HP + 'run'):
                bold, italic = shapes.get(run.get('charPrIDRef'), (False, False))
                for text in run.findall(HP + 't'):
                    result += glyphs(''.join(text.itertext()), bold, italic)
    return result


def records(data):
    offset = 0
    while offset < len(data):
        header, = struct.unpack_from('<I', data, offset)
        offset += 4
        tag, level, size = header & 1023, (header >> 10) & 1023, header >> 20
        if size == 4095:
            size, = struct.unpack_from('<I', data, offset)
            offset += 4
        if offset + size > len(data):
            raise ValueError('truncated HWP record')
        yield tag, level, data[offset:offset + size]
        offset += size


def hwp_glyphs(path):
    import olefile
    with olefile.OleFileIO(path) as ole:
        compressed = struct.unpack_from('<I', ole.openstream('FileHeader').read(), 36)[0] & 1
        def stream(name):
            data = ole.openstream(name).read()
            return zlib.decompress(data, -15) if compressed else data
        shapes = []
        for tag, _, data in records(stream('DocInfo')):
            if tag == 21:
                attr, = struct.unpack_from('<I', data, 46)
                shapes.append((bool(attr & 2), bool(attr & 1)))
        result = []
        names = sorted((name for name in ole.listdir() if len(name) == 2 and name[0] == 'BodyText' and name[1].startswith('Section')),
                       key=lambda name: int(name[1][7:]))
        for name in names:
            rs = list(records(stream(name)))
            for idx, (tag, level, data) in enumerate(rs):
                if tag != 66:
                    continue
                text, pairs = b'', [(0, 0)]
                for child_tag, child_level, child_data in rs[idx + 1:]:
                    if child_tag == 66 or child_level <= level:
                        break
                    if child_level == level + 1 and child_tag == 67:
                        text = child_data
                    elif child_level == level + 1 and child_tag == 68:
                        pairs = list(struct.iter_unpack('<II', child_data))
                pos = 0
                while pos * 2 < len(text):
                    char, = struct.unpack_from('<H', text, pos * 2)
                    if char == 13:
                        break
                    if char in {2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25}:
                        pos += 8
                        continue
                    shape_id = 0
                    for start, sid in pairs:
                        if start > pos:
                            break
                        shape_id = sid
                    if char >= 32:
                        bold, italic = shapes[shape_id]
                        result += glyphs(chr(char), bold, italic)
                    pos += 1
        return result


def counts(chars):
    result = Counter(('bold' if b else 'regular') + ('+italic' if i else '') for _, b, i in chars)
    return dict(result)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('outputs', type=Path, nargs='+')
    args = parser.parse_args()
    source = docx_glyphs(args.source)
    report = {'source': str(args.source), 'source_sha256': hashlib.sha256(args.source.read_bytes()).hexdigest(),
              'source_characters': len(source), 'source_styles': counts(source), 'outputs': []}
    for path in args.outputs:
        output = hwp_glyphs(path) if path.suffix.lower() == '.hwp' else hwpx_glyphs(path)
        source_text, output_text = ''.join(c[0] for c in source), ''.join(c[0] for c in output)
        matcher = SequenceMatcher(None, source_text, output_text, autojunk=False)
        matched = mismatch_bold = mismatch_italic = added_bold = added_italic = 0
        examples = []
        for left, right, size in matcher.get_matching_blocks():
            matched += size
            for index in range(size):
                original, converted = source[left + index], output[right + index]
                mismatch_bold += original[1] != converted[1]
                mismatch_italic += original[2] != converted[2]
                added_bold += not original[1] and converted[1]
                added_italic += not original[2] and converted[2]
                if original[1:] != converted[1:] and len(examples) < 5:
                    examples.append({'source_index': left + index, 'output_index': right + index,
                                     'source': original[1:], 'output': converted[1:],
                                     'context': source_text[max(0, left + index - 12):left + index + 30]})
        report['outputs'].append({'path': str(path), 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
                                  'characters': len(output), 'styles': counts(output),
                                  'text_equal': source_text == output_text, 'matched_source_characters': matched,
                                  'unmatched_source_characters': len(source) - matched,
                                  'unmatched_output_characters': len(output) - matched,
                                  'bold_mismatches': mismatch_bold, 'italic_mismatches': mismatch_italic,
                                  'added_bold_characters': added_bold, 'added_italic_characters': added_italic,
                                  'examples': examples})
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()

#!/usr/bin/env python3
"""Dump HWP 5 CFB streams and decoded DocInfo/BodyText records."""

from __future__ import annotations

import argparse
import struct
import sys
import zlib
from pathlib import Path

try:
    import olefile
except ImportError as exc:  # pragma: no cover - diagnostic environment guard
    raise SystemExit("hwp-dump.py requires the Python 'olefile' package") from exc


BEGIN = 0x10
TAG_NAMES = {
    BEGIN + 0: "DOCUMENT_PROPERTIES",
    BEGIN + 1: "ID_MAPPINGS",
    BEGIN + 2: "BIN_DATA",
    BEGIN + 3: "FACE_NAME",
    BEGIN + 4: "BORDER_FILL",
    BEGIN + 5: "CHAR_SHAPE",
    BEGIN + 6: "TAB_DEF",
    BEGIN + 7: "NUMBERING",
    BEGIN + 8: "BULLET",
    BEGIN + 9: "PARA_SHAPE",
    BEGIN + 10: "STYLE",
    BEGIN + 11: "DOC_DATA",
    BEGIN + 14: "COMPATIBLE_DOCUMENT",
    BEGIN + 15: "LAYOUT_COMPATIBILITY",
    BEGIN + 50: "PARA_HEADER",
    BEGIN + 51: "PARA_TEXT",
    BEGIN + 52: "PARA_CHAR_SHAPE",
    BEGIN + 53: "PARA_LINE_SEG",
    BEGIN + 55: "CTRL_HEADER",
    BEGIN + 56: "LIST_HEADER",
    BEGIN + 57: "PAGE_DEF",
    BEGIN + 58: "FOOTNOTE_SHAPE",
    BEGIN + 59: "PAGE_BORDER_FILL",
    BEGIN + 60: "SHAPE_COMPONENT",
    BEGIN + 61: "TABLE",
    BEGIN + 69: "SHAPE_COMPONENT_PICTURE",
    BEGIN + 71: "CTRL_DATA",
}

ID_MAPPING_ITEMS = [
    ("BIN_DATA", BEGIN + 2),
    ("FACE_NAME_HANGUL", BEGIN + 3),
    ("FACE_NAME_LATIN", BEGIN + 3),
    ("FACE_NAME_HANJA", BEGIN + 3),
    ("FACE_NAME_JAPANESE", BEGIN + 3),
    ("FACE_NAME_OTHER", BEGIN + 3),
    ("FACE_NAME_SYMBOL", BEGIN + 3),
    ("FACE_NAME_USER", BEGIN + 3),
    ("BORDER_FILL", BEGIN + 4),
    ("CHAR_SHAPE", BEGIN + 5),
    ("TAB_DEF", BEGIN + 6),
    ("NUMBERING", BEGIN + 7),
    ("BULLET", BEGIN + 8),
    ("PARA_SHAPE", BEGIN + 9),
    ("STYLE", BEGIN + 10),
    ("MEMO_SHAPE", BEGIN + 76),
    ("TRACK_CHANGE", BEGIN + 80),
    ("TRACK_CHANGE_AUTHOR", BEGIN + 81),
]


def u16(data: bytes, offset: int) -> int:
    return struct.unpack_from("<H", data, offset)[0]


def u32(data: bytes, offset: int) -> int:
    return struct.unpack_from("<I", data, offset)[0]


def parse_records(data: bytes) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    offset = 0
    while offset < len(data):
        if offset + 4 > len(data):
            raise ValueError(f"{len(data) - offset} trailing record byte(s)")
        header = u32(data, offset)
        tag = header & 0x3FF
        level = (header >> 10) & 0x3FF
        size = (header >> 20) & 0xFFF
        offset += 4
        if size == 0xFFF:
            if offset + 4 > len(data):
                raise ValueError("missing extended record size")
            size = u32(data, offset)
            offset += 4
        end = offset + size
        if end > len(data):
            raise ValueError(f"record size {size} exceeds stream at {offset}")
        records.append({"tag": tag, "level": level, "data": data[offset:end]})
        offset = end
    return records


def read_utf16(data: bytes, offset: int) -> tuple[str, int]:
    length = u16(data, offset)
    offset += 2
    end = offset + length * 2
    if end > len(data):
        raise ValueError("truncated UTF-16 field")
    return data[offset:end].decode("utf-16le"), end


def style_detail(data: bytes) -> str:
    name, offset = read_utf16(data, 0)
    english, offset = read_utf16(data, offset)
    if offset + 8 > len(data):
        return "STYLE(truncated)"
    style_type = data[offset]
    next_id = data[offset + 1]
    language = struct.unpack_from("<h", data, offset + 2)[0]
    para_id = u16(data, offset + 4)
    char_id = u16(data, offset + 6)
    return (
        f"name={name!r} english={english!r} type={style_type} "
        f"next={next_id} lang={language} paraShape={para_id} charShape={char_id}"
    )


def numbering_detail(data: bytes) -> str:
    levels: list[str] = []
    offset = 0
    for level in range(1, 8):
        if offset + 14 > len(data):
            return "NUMBERING(truncated)"
        attr = u32(data, offset)
        adjust = u16(data, offset + 4)
        distance = u16(data, offset + 6)
        char_shape = u32(data, offset + 8)
        offset += 12
        number_format, offset = read_utf16(data, offset)
        levels.append(
            f"L{level}[attr=0x{attr:08x},adjust={adjust},distance={distance},"
            f"charShape={char_shape},format={number_format!r}]"
        )
    return " ".join(levels)


def record_detail(tag: int, data: bytes) -> str:
    if tag == BEGIN + 10:
        return style_detail(data)
    if tag == BEGIN + 7:
        return numbering_detail(data)
    if tag == BEGIN + 8 and len(data) >= 10:
        return f"bullet={chr(u16(data, 8))!r}"
    if tag == BEGIN + 9 and len(data) >= 32:
        attr = u32(data, 0)
        return (
            f"headType={(attr >> 23) & 3} level={(attr >> 25) & 7} "
            f"numberingOrBulletId={u16(data, 30)}"
        )
    if tag == BEGIN + 50 and len(data) >= 12:
        return (
            f"nchars=0x{u32(data, 0):08x} paraShape={u16(data, 8)} "
            f"style={data[10]} divideSort=0x{data[11]:02x}"
        )
    if tag == BEGIN + 61 and len(data) >= 8:
        return f"rows={u16(data, 4)} columns={u16(data, 6)}"
    return ""


def dump_id_mappings(records: list[dict[str, object]]) -> None:
    mappings = [record for record in records if record["tag"] == BEGIN + 1]
    if len(mappings) != 1:
        print(f"ERROR ID_MAPPINGS record count={len(mappings)}")
        return
    data = mappings[0]["data"]
    assert isinstance(data, bytes)
    if len(data) < 72:
        print(f"ERROR ID_MAPPINGS size={len(data)}")
        return
    values = [u32(data, index * 4) for index in range(18)]
    actual_by_tag: dict[int, int] = {}
    for record in records:
        tag = int(record["tag"])
        actual_by_tag[tag] = actual_by_tag.get(tag, 0) + 1
    face_total = sum(values[1:8])
    actual_faces = actual_by_tag.get(BEGIN + 3, 0)
    for index, (name, tag) in enumerate(ID_MAPPING_ITEMS):
        if 1 <= index <= 7:
            status = "OK" if face_total == actual_faces else "ERROR"
            actual = actual_faces if index == 1 else "shared"
        else:
            actual = actual_by_tag.get(tag, 0)
            status = "OK" if values[index] == actual else "ERROR"
        print(f"  [{index:02}] {name}: declared={values[index]} actual={actual} {status}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("hwp", type=Path)
    args = parser.parse_args()

    with olefile.OleFileIO(args.hwp) as ole:
        stream_paths = [path for path in ole.listdir(streams=True, storages=False)]
        print("# CFB streams")
        for path in sorted(stream_paths):
            data = ole.openstream(path).read()
            print(f"{'/'.join(path)}\t{len(data)}")

        header = ole.openstream(["FileHeader"]).read()
        compressed = len(header) >= 40 and (u32(header, 36) & 1) != 0
        for path in (["DocInfo"], ["BodyText", "Section0"]):
            stored = ole.openstream(path).read()
            raw = zlib.decompress(stored, -15) if compressed else stored
            records = parse_records(raw)
            name = "/".join(path)
            print(f"\n# {name}: stored={len(stored)} raw={len(raw)} records={len(records)} leftover=0")
            for record in records:
                tag = int(record["tag"])
                level = int(record["level"])
                data = record["data"]
                assert isinstance(data, bytes)
                detail = record_detail(tag, data)
                suffix = f" {detail}" if detail else ""
                print(
                    f"{'  ' * level}{TAG_NAMES.get(tag, f'TAG_{tag}')}"
                    f"({tag}) size={len(data)} level={level}{suffix}"
                )
            if path == ["DocInfo"]:
                print("\n# ID_MAPPINGS")
                dump_id_mappings(records)
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""Builds xl/vbaProject.bin from vba/*.bas without Excel.

The container follows MS-OVBA / MS-CFB. The performance cache is left out (_VBA_PROJECT carries the
'no cache' header), so Excel compiles the source itself when the workbook opens.
"""
import struct, sys

CLS = {"ThisWorkbook": "0{00020819-0000-0000-C000-000000000046}", "Sheet": "0{00020820-0000-0000-C000-000000000046}"}
OFFICE_REF = bytes.fromhex(
    "1600060000004f66666963653e000c0000004f00660066006900630065000d009e000000940000002a5c477b32444638443034432d3542"
    "46412d313031422d424445352d3030414130303434444535327d23322e30233023433a5c50726f6772616d2046696c65735c436f6d6d6f"
    "6e2046696c65735c4d6963726f736f6674205368617265645c4f464649434531365c4d534f2e444c4c234d6963726f736f667420"
    "4f66666963652031362e30204f626a656374204c696272617279000000000000")
PROJECT_ID = "{5CF9CDAB-B77B-4FE4-BD7E-D29900070E1B}"
PROTECT = 'CMG="25270E6B34DA38DA38DA38DA38"\r\nDPB="25270E6B34DF35DF35DF"\r\nGC="25270E6B34DF35DF3520"\r\n'


def compress(data: bytes) -> bytes:
    """MS-OVBA compression (LZ77 with the spec's variable offset/length split)."""
    out = bytearray(b"\x01")
    for c in range(0, len(data), 4096):
        ch = data[c:c + 4096]; body = bytearray(); pos = 0
        while pos < len(ch):
            flag_at = len(body); body.append(0); flag = 0
            for bit in range(8):
                if pos >= len(ch): break
                bc = max((pos - 1).bit_length(), 4) if pos > 1 else 4
                max_len = (0xFFFF >> bc) + 3
                best_len, best_off = 0, 0
                for start in range(max(0, pos - (1 << bc)), pos):
                    n = 0
                    while pos + n < len(ch) and n < max_len and ch[start + n] == ch[pos + n]: n += 1
                    if n > best_len: best_len, best_off = n, pos - start
                if best_len >= 3:
                    body += struct.pack("<H", ((best_off - 1) << (16 - bc)) | (best_len - 3)); flag |= 1 << bit; pos += best_len
                else:
                    body.append(ch[pos]); pos += 1
            body[flag_at] = flag
        out += struct.pack("<H", ((len(body) + 2 - 3) & 0x0FFF) | 0x3000 | 0x8000) + body
    return bytes(out)


def rec(rid, payload=b""):
    return struct.pack("<HI", rid, len(payload)) + payload


def build_dir(mods):
    d = rec(1, struct.pack("<I", 1)) + rec(2, struct.pack("<I", 0x409)) + rec(0x14, struct.pack("<I", 0x409))
    d += rec(3, struct.pack("<H", 1252)) + rec(4, b"VBAProject") + rec(5) + rec(0x40) + rec(6) + rec(0x3D)
    d += rec(7, struct.pack("<I", 0)) + rec(8, struct.pack("<I", 0))
    d += struct.pack("<HI", 9, 4) + struct.pack("<IH", 0x68FBFAC5, 4) + rec(0x0C) + rec(0x3C)
    d += OFFICE_REF + rec(0x0F, struct.pack("<H", len(mods))) + rec(0x13, struct.pack("<H", 0xFFFF))
    for name, kind in mods:
        n = name.encode(); u = name.encode("utf-16le")
        d += rec(0x19, n) + rec(0x47, u) + rec(0x1A, n) + rec(0x32, u) + rec(0x1C) + rec(0x48)
        d += rec(0x31, struct.pack("<I", 0)) + rec(0x1E, struct.pack("<I", 0)) + rec(0x2C, struct.pack("<H", 0xFFFF))
        d += rec(0x21 if kind == "std" else 0x22) + rec(0x2B)
    return d + rec(0x10)


def cfb(streams):
    """streams: {path: bytes}; paths like 'PROJECT' or 'VBA/dir'. Minimal CFB v3 writer."""
    S, MS, CUT = 512, 64, 4096
    ents = [dict(name="Root Entry", type=5, data=None)]
    ents.append(dict(name="VBA", type=1, data=None))
    for p, b in streams.items():
        ents.append(dict(name=p.split("/")[-1], type=2, data=b, parent="VBA" if "/" in p else "Root Entry"))
    for i, e in enumerate(ents): e["idx"] = i
    key = lambda e: (len(e["name"]), e["name"].upper())

    def link(parent_name, kids):
        kids = sorted(kids, key=key)
        def build(lo, hi):
            if lo >= hi: return 0xFFFFFFFF
            m = (lo + hi) // 2; e = kids[m]
            e["left"] = build(lo, m); e["right"] = build(m + 1, hi); e["color"] = 1
            return e["idx"]
        return build(0, len(kids))
    for e in ents: e.setdefault("left", 0xFFFFFFFF); e.setdefault("right", 0xFFFFFFFF); e.setdefault("child", 0xFFFFFFFF); e.setdefault("color", 1)
    vba_kids = [e for e in ents if e.get("parent") == "VBA"]
    root_kids = [ents[1]] + [e for e in ents if e.get("parent") == "Root Entry"]
    ents[1]["child"] = link("VBA", vba_kids); ents[0]["child"] = link("Root", root_kids)

    # mini stream
    mini, minifat = bytearray(), []
    for e in ents:
        if e["type"] == 2 and len(e["data"]) < CUT:
            e["start"] = len(mini) // MS; n = -(-len(e["data"]) // MS) or 1
            mini += e["data"].ljust(n * MS, b"\0")
            minifat += list(range(len(minifat) + 1, len(minifat) + n)) + [0xFFFFFFFE]
    n_dir = -(-len(ents) * 128 // S); n_mf = -(-len(minifat) * 4 // S); n_ms = -(-len(mini) // S)
    big = [e for e in ents if e["type"] == 2 and len(e["data"]) >= CUT]
    n_big = sum(-(-len(e["data"]) // S) for e in big)
    n_fat = 1
    while n_fat * 128 < n_fat + n_dir + n_mf + n_ms + n_big: n_fat += 1
    fat = [0xFFFFFFFD] * n_fat; pos = n_fat
    def chain(n):
        nonlocal pos
        s = pos; fat.extend(list(range(pos + 1, pos + n)) + [0xFFFFFFFE]); pos += n; return s
    dir_s = chain(n_dir); mf_s = chain(n_mf) if n_mf else 0xFFFFFFFE; ms_s = chain(n_ms) if n_ms else 0xFFFFFFFE
    for e in big: e["start"] = chain(-(-len(e["data"]) // S))
    fat += [0xFFFFFFFF] * (n_fat * 128 - len(fat))
    ents[0]["start"] = ms_s; ents[0]["size"] = len(mini)
    d = bytearray()
    for e in ents:
        nm = e["name"].encode("utf-16le") + b"\0\0"
        size = e.get("size", len(e["data"]) if e["data"] is not None else 0)
        d += nm.ljust(64, b"\0") + struct.pack("<HBBIII", len(nm), e["type"], e["color"], e["left"], e["right"], e["child"])
        d += b"\0" * 16 + struct.pack("<IQQ", 0, 0, 0) + struct.pack("<IIi", e.get("start", 0) if e["type"] != 1 else 0, size, 0)
    d = bytes(d).ljust(n_dir * S, b"\0")
    # empty directory slots must read as unused
    for k in range(len(ents), n_dir * 4):
        d = d[:k * 128] + b"\0" * 64 + struct.pack("<HBBIII", 0, 0, 1, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF) + d[k * 128 + 80:]
    hdr = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\0" * 16 + struct.pack("<HHHHHH", 0x3E, 3, 0xFFFE, 9, 6, 0) + b"\0" * 4
    difat = list(range(n_fat)) + [0xFFFFFFFF] * (109 - n_fat)
    hdr += struct.pack("<IIIIIIIII", 0, n_fat, dir_s, 0, CUT, mf_s, n_mf, 0xFFFFFFFE, 0) + struct.pack("<109I", *difat)
    out = bytearray(hdr) + struct.pack(f"<{len(fat)}I", *fat) + d
    out += struct.pack(f"<{len(minifat)}I", *minifat).ljust(n_mf * S, b"\xff") if n_mf else b""
    out += bytes(mini).ljust(n_ms * S, b"\0")
    for e in big: out += e["data"].ljust(-(-len(e["data"]) // S) * S, b"\0")
    return bytes(out)


def build(bas_path, sheets, out_path):
    src = open(bas_path, encoding="utf-8").read().replace("\r\n", "\n").replace("\n", "\r\n")
    std_name = [l for l in src.split("\r\n") if l.startswith("Attribute VB_Name")][0].split('"')[1]
    mods = [("ThisWorkbook", "doc")] + [(s, "doc") for s in sheets] + [(std_name, "std")]
    st = {"VBA/dir": compress(build_dir(mods)), "VBA/_VBA_PROJECT": b"\xcc\x61\xff\xff\x00\x01\x00"}
    for name, kind in mods:
        if kind == "std": text = src
        else:
            base = CLS["ThisWorkbook"] if name == "ThisWorkbook" else CLS["Sheet"]
            text = (f'Attribute VB_Name = "{name}"\r\nAttribute VB_Base = "{base}"\r\nAttribute VB_GlobalNameSpace = False\r\n'
                    'Attribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n'
                    'Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n')
        st["VBA/" + name] = compress(text.encode("cp1252"))
    docs = "".join(f"Document={n}/&H00000000\r\n" for n, k in mods if k == "doc")
    proj = (f'ID="{PROJECT_ID}"\r\n{docs}Module={std_name}\r\nName="VBAProject"\r\nHelpContextID="0"\r\nVersionCompatible32="393222000"\r\n'
            f'{PROTECT}\r\n[Host Extender Info]\r\n&H00000001={{3832D640-CF90-11CF-8E43-00A0C911005A}};VBE;&H00000000\r\n\r\n[Workspace]\r\n'
            + "".join(f"{n}=0, 0, 0, 0, C\r\n" for n, _ in mods))
    st["PROJECT"] = proj.encode("cp1252")
    st["PROJECTwm"] = b"".join(n.encode() + b"\0" + n.encode("utf-16le") + b"\0\0" for n, _ in mods) + b"\0\0"
    open(out_path, "wb").write(cfb(st))


if __name__ == "__main__":
    build("vba/modLabourPlan.bas", [f"Sheet{i}" for i in range(1, 6)], "build/vbaProject.bin")
    print("wrote build/vbaProject.bin")

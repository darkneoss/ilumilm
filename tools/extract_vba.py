"""Extrae los modulos VBA del SEAD Street Lighting Tool.

Los streams _VBA_PROJECT_CUR/VBA/* estan comprimidos con el algoritmo RLE de
MS-OVBA. olefile nos da el stream crudo; la descompresion son las ~45 lineas de
abajo, asi evitamos depender de oletools (que no esta instalado).
"""
import struct
import sys
from pathlib import Path

import olefile


def decompress(data: bytes) -> bytes:
    """Descomprime un CompressedContainer de MS-OVBA (2.4.1.3.6)."""
    if not data or data[0] != 0x01:
        raise ValueError("no es un CompressedContainer (falta el signature byte 0x01)")
    out = bytearray()
    i = 1
    while i < len(data) - 1:
        header = struct.unpack("<H", data[i:i + 2])[0]
        i += 2
        size = (header & 0x0FFF) + 3
        compressed = bool(header & 0x8000)
        end = i + size - 2
        chunk_start = len(out)
        if not compressed:
            out += data[i:i + 4096]
            i += 4096
            continue
        while i < end:
            flags = data[i]
            i += 1
            for bit in range(8):
                if i >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[i])
                    i += 1
                else:
                    token = struct.unpack("<H", data[i:i + 2])[0]
                    i += 2
                    # el reparto de bits offset/longitud depende de cuanto
                    # llevamos escrito en ESTE chunk, no en la salida total
                    written = len(out) - chunk_start
                    bits = 4
                    while (1 << bits) < written:
                        bits += 1
                    bits = max(4, min(12, bits))
                    length = (token & ((1 << (16 - bits)) - 1)) + 3
                    offset = (token >> (16 - bits)) + 1
                    for _ in range(length):
                        out.append(out[len(out) - offset])
    return bytes(out)


def find_source(stream: bytes) -> bytes:
    """El stream trae metadatos antes del container; lo localizamos probando."""
    for start in range(len(stream)):
        if stream[start] != 0x01:
            continue
        try:
            text = decompress(stream[start:])
        except Exception:
            continue
        if b"Attribute VB_Name" in text[:400]:
            return text
    raise ValueError("no se encontro el codigo fuente en el stream")


def main(xls: Path, outdir: Path) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    ole = olefile.OleFileIO(str(xls))
    for entry in ole.listdir():
        if len(entry) < 3 or entry[1] != "VBA":
            continue
        name = entry[-1]
        if name in ("_VBA_PROJECT", "dir"):
            continue
        try:
            source = find_source(ole.openstream(entry).read())
        except Exception as exc:
            print(f"  !! {name}: {exc}")
            continue
        dest = outdir / f"{name}.bas"
        dest.write_bytes(source)
        print(f"  -> {dest.name} ({len(source)} bytes)")


if __name__ == "__main__":
    main(Path(sys.argv[1]), Path(sys.argv[2]))

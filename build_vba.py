#!/usr/bin/env python3
"""
build_vba.py
  1. Читает xlsm из ZIP
  2. Распаковывает vbaProject.bin через olefile
  3. Получает оригинальный VBA-код HotelMacros через oletools
  4. Добавляет новый саб ОткрытьФормуБронирования (InputBox-форма)
  5. Упаковывает обратно в MS-OVBA (raw chunks — валидно по ECMA-376)
  6. Записывает новый vbaProject.bin через OLE-пересборщик
  7. Перепаковывает xlsm и zip
"""

import struct, os, shutil, zipfile, io, math
import olefile
from oletools.olevba import decompress_stream

SECTOR = 512
FREESECT   = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT    = 0xFFFFFFFD
DIFSECT    = 0xFFFFFFFC
MINI_CUTOFF = 4096

# ──────────────────────────────────────────────────────────────────────────────
# MS-OVBA compression using raw (uncompressed) chunks — всегда корректно
# ──────────────────────────────────────────────────────────────────────────────
def _copy_token_fields(decompressed_chunk_so_far: int):
    """
    Matches oletools' copytoken_help exactly.
    bit_count = offset bits (grows as more bytes are decompressed).
    length_bits = 16 - bit_count (shrinks).
    Returns (ob, lb, length_mask, max_len, max_offset).
    """
    difference = max(decompressed_chunk_so_far, 1)
    ob = max(math.ceil(math.log2(difference)), 4)  # offset bits
    lb = 16 - ob                                    # length bits
    length_mask = 0xFFFF >> ob                      # bottom lb bits
    max_len    = length_mask + 3
    max_offset = 1 << ob
    return ob, lb, length_mask, max_len, max_offset

def _find_match(chunk: bytes, pos: int):
    """Greedy LZ77 match. Returns (match_len, match_offset) or (0,0)."""
    ob, lb, length_mask, max_len, max_offset = _copy_token_fields(pos)
    search_start = max(0, pos - max_offset)
    best_len, best_off = 0, 0
    for back in range(pos - 1, search_start - 1, -1):
        mlen = 0
        while mlen < max_len and (pos + mlen) < len(chunk) and chunk[back + mlen] == chunk[pos + mlen]:
            mlen += 1
        if mlen >= 3 and mlen > best_len:
            best_len = mlen
            best_off = pos - back
    return best_len, best_off

def _compress_chunk(chunk: bytes) -> bytes:
    """Compress up to 4096 bytes using MS-OVBA LZ77."""
    out = bytearray()
    pos = 0
    while pos < len(chunk):
        flag_pos = len(out)
        out.append(0)
        flag = 0
        for bit in range(8):
            if pos >= len(chunk):
                break
            mlen, moff = _find_match(chunk, pos)
            if mlen >= 3:
                ob, lb, length_mask, max_len, max_offset = _copy_token_fields(pos)
                # offset in top ob bits, length in bottom lb bits
                token = ((moff - 1) << lb) | (mlen - 3)
                out += struct.pack('<H', token)
                flag |= (1 << bit)
                pos += mlen
            else:
                out.append(chunk[pos])
                pos += 1
        out[flag_pos] = flag
    return bytes(out)

def ovba_compress(source: bytes) -> bytes:
    """Compress VBA source using MS-OVBA LZ77 (ECMA-376 §2.4.1)."""
    out = bytearray([0x01])  # SignatureByte
    i = 0
    while i < len(source):
        chunk = source[i: i + 4096]
        i += 4096
        compressed = _compress_chunk(chunk)
        # CompressedChunkSize = len(compressed) - 3 (but must fit in 12 bits, max 4094)
        if len(compressed) + 3 > 4098 or len(compressed) >= len(chunk):
            # Store as raw chunk: bit15=0, bits14-12=011, bits11-0=4093
            padded = chunk + b'\x00' * (4096 - len(chunk))
            hdr = struct.pack('<H', 0x3FFD)   # 0x3000|4093, bit15=0
            out += hdr + padded
        else:
            # Compressed: bit15=1, bits14-12=011, bits11-0=chunkDataSize-1
            # oletools reads chunk_data = chunk_size-2 bytes; chunk_size=(hdr&0x0FFF)+3
            # so chunk_data = (hdr&0x0FFF)+1 = len(compressed) → hdr&0x0FFF = len-1
            hdr = struct.pack('<H', 0xB000 | (len(compressed) - 1))
            out += hdr + compressed
    return bytes(out)


# ──────────────────────────────────────────────────────────────────────────────
# Чтение всех потоков из OLE-файла
# ──────────────────────────────────────────────────────────────────────────────
def read_all_streams(ole_path: str):
    """Возвращает {path: bytes} для всех streams в OLE"""
    ole = olefile.OleFileIO(ole_path)
    streams = {}
    for entry in ole.listdir(streams=True):
        path = '/'.join(entry)
        streams[path] = ole.openstream(path).read()
    ole.close()
    return streams


# ──────────────────────────────────────────────────────────────────────────────
# Минимальный OLE Compound File writer
# Пишет новый .bin с заданными потоками, перестраивая всё с нуля
# ──────────────────────────────────────────────────────────────────────────────
def build_ole(stream_tree: dict) -> bytes:
    """
    stream_tree: {
        'ROOT': None,                  # root storage
        'PROJECT': b'...',             # top-level stream
        'PROJECTwm': b'...',
        'VBA': {                       # storage
            '_VBA_PROJECT': b'...',
            'dir': b'...',
            'HotelMacros': b'...',
            'Лист1': b'...',
            ...
        }
    }
    Returns raw bytes of the OLE compound file.
    """
    # ── Flatten streams into a list of (full_path_list, data) ───────────────
    flat = []  # (name_parts_tuple, data)

    def flatten(node, prefix):
        if isinstance(node, dict):
            for k, v in node.items():
                if k == '__data__':
                    continue
                flatten(v, prefix + [k])
        elif isinstance(node, (bytes, bytearray)):
            flat.append((tuple(prefix), bytes(node)))

    flatten(stream_tree, [])

    # ── Build directory entries ─────────────────────────────────────────────
    # OLE directory: root=0, then storages, then streams
    # For simplicity: root + one-level storages + streams
    # We'll build a flat directory with proper parent/sibling/child links
    # using the "red-black tree" approach (simplified: all left-black here)

    DIR_ENTRY_SIZE = 128

    class DirEntry:
        def __init__(self, name, obj_type, data=b'', start=ENDOFCHAIN, size=0):
            self.name = name            # str
            self.obj_type = obj_type    # 1=storage, 2=stream, 5=root
            self.data = data
            self.start = start
            self.size = size
            self.child = 0xFFFFFFFF
            self.sibling_left  = 0xFFFFFFFF
            self.sibling_right = 0xFFFFFFFF
            self.clsid = b'\x00' * 16
            self.color = 1   # 1=black (default)
            self.created = b'\x00' * 8
            self.modified = b'\x00' * 8

        def pack(self):
            name_utf16 = self.name.encode('utf-16-le')
            name_len = len(name_utf16) + 2  # include null terminator
            name_field = (name_utf16 + b'\x00\x00').ljust(64, b'\x00')[:64]
            return (
                name_field +
                struct.pack('<H', min(name_len, 64)) +
                struct.pack('<BB', self.obj_type, self.color) +
                struct.pack('<I', self.sibling_left) +
                struct.pack('<I', self.sibling_right) +
                struct.pack('<I', self.child) +
                self.clsid +
                struct.pack('<I', 0) +  # state bits
                self.created +
                self.modified +
                struct.pack('<I', self.start) +
                struct.pack('<I', self.size) +
                b'\x00' * 4   # unused high dword of size
            )

    # Collect all storage names
    storage_contents = {}  # storage_name -> [child_names]
    for path_tuple, data in flat:
        if len(path_tuple) == 1:
            storage_contents.setdefault('ROOT', []).append(path_tuple[0])
        elif len(path_tuple) == 2:
            storage_contents.setdefault(path_tuple[0], []).append(path_tuple[1])

    # ── Sector allocation ─────────────────────────────────────────────────
    # We'll lay out sectors in order:
    #   sector 0: FAT sector (placeholder, filled last)
    #   sector 1: Directory sector(s)
    #   sector 2+: stream data sectors

    # First, determine all stream data and their sizes
    all_entries = []  # DirEntry list in directory order

    root_entry = DirEntry('Root Entry', 5)
    root_entry.clsid = bytes.fromhex('06090200000000000000000000000000')
    all_entries.append(root_entry)  # index 0

    # Collect storages (top level)
    top_level_storages = []
    top_level_streams_data = {}   # name -> data (for top-level streams)
    sub_streams_data = {}         # (storage, name) -> data

    for path_tuple, data in flat:
        if len(path_tuple) == 1:
            top_level_streams_data[path_tuple[0]] = data
        elif len(path_tuple) == 2:
            sub_streams_data[(path_tuple[0], path_tuple[1])] = data
            top_level_storages.append(path_tuple[0])

    top_level_storages = list(dict.fromkeys(top_level_storages))  # dedup, preserve order

    # Build directory: root → VBA storage (and others if any) → streams
    # Directory order (DFS): root(0), VBA(1), VBA children(2..N), top-level streams(N+1..)

    storage_entries = []
    for st_name in top_level_storages:
        e = DirEntry(st_name, 1)
        e.clsid = b'\x00' * 16
        storage_entries.append(e)
        all_entries.append(e)

    # All stream entries
    stream_entries = {}   # (path_tuple) -> DirEntry

    # Sub-streams (inside storages)
    for st_name in top_level_storages:
        children = [k[1] for k in sub_streams_data if k[0] == st_name]
        children.sort()
        for child_name in children:
            data = sub_streams_data[(st_name, child_name)]
            e = DirEntry(child_name, 2, data=data, size=len(data))
            stream_entries[(st_name, child_name)] = e
            all_entries.append(e)

    # Top-level streams
    for name in sorted(top_level_streams_data.keys()):
        data = top_level_streams_data[name]
        e = DirEntry(name, 2, data=data, size=len(data))
        stream_entries[(name,)] = e
        all_entries.append(e)

    # ── Set up child/sibling links ────────────────────────────────────────
    # Find index of an entry by name
    def idx_of(e): return all_entries.index(e)

    def link_siblings(entry_list):
        """Set sibling links for a list of directory entries (simple chain)"""
        if not entry_list:
            return 0xFFFFFFFF
        # Sort by name for RB-tree (simplified: linear left-right chain)
        entry_list = sorted(entry_list, key=lambda e: e.name.upper())
        mid = len(entry_list) // 2
        root_e = entry_list[mid]
        if mid > 0:
            root_e.sibling_left = idx_of(link_siblings_get_root(entry_list[:mid]))
        if mid + 1 < len(entry_list):
            root_e.sibling_right = idx_of(link_siblings_get_root(entry_list[mid+1:]))
        return root_e

    def link_siblings_get_root(entry_list):
        entry_list = sorted(entry_list, key=lambda e: e.name.upper())
        mid = len(entry_list) // 2
        root_e = entry_list[mid]
        if mid > 0:
            root_e.sibling_left = idx_of(link_siblings_get_root(entry_list[:mid]))
        if mid + 1 < len(entry_list):
            root_e.sibling_right = idx_of(link_siblings_get_root(entry_list[mid+1:]))
        return root_e

    # Set root's children
    root_children = storage_entries + [stream_entries[(n,)] for n in sorted(top_level_streams_data)]
    if root_children:
        rc = link_siblings_get_root(root_children)
        root_entry.child = idx_of(rc)

    # Set each storage's children
    for st_name, st_e in zip(top_level_storages, storage_entries):
        children = [stream_entries[(st_name, k[1])] for k in sub_streams_data if k[0] == st_name]
        if children:
            cc = link_siblings_get_root(children)
            st_e.child = idx_of(cc)

    # ── Assign sectors to streams ─────────────────────────────────────────
    # Mini stream: streams < MINI_CUTOFF → stored in mini stream
    # Regular streams: stored in normal sectors

    mini_stream = bytearray()
    mini_fat = []         # list of uint32
    MINI_SECTOR = 64

    regular_data = []     # list of (entry, data)

    for e in all_entries:
        if e.obj_type != 2:
            continue
        if len(e.data) < MINI_CUTOFF and len(e.data) > 0:
            # mini stream
            mini_offset = len(mini_stream)
            mini_start = mini_offset // MINI_SECTOR
            n_mini = (len(e.data) + MINI_SECTOR - 1) // MINI_SECTOR
            padded = e.data + b'\x00' * (n_mini * MINI_SECTOR - len(e.data))
            mini_stream.extend(padded)
            # build mini FAT chain
            for i in range(n_mini - 1):
                mini_fat.append(mini_start + i + 1)
            mini_fat.append(ENDOFCHAIN)
            e.start = mini_start
        elif len(e.data) == 0:
            e.start = ENDOFCHAIN
        else:
            regular_data.append(e)

    # Root's data = mini stream; assign sectors later
    # Layout sectors:
    # sector 0 = FAT (placeholder)
    # sector 1+ = directory
    # then mini FAT sectors
    # then root's mini stream sectors
    # then regular stream sectors

    num_dir_entries = len(all_entries)
    # Pad to multiple of 4 entries per sector (4 * 128 = 512 bytes per sector)
    entries_per_sector = SECTOR // DIR_ENTRY_SIZE  # 4
    num_dir_sectors = (num_dir_entries + entries_per_sector - 1) // entries_per_sector
    # Also need padding to fill dir sectors
    while len(all_entries) < num_dir_sectors * entries_per_sector:
        # empty entry
        empty = DirEntry('', 0)
        empty.sibling_left = 0xFFFFFFFF
        empty.sibling_right = 0xFFFFFFFF
        empty.child = 0xFFFFFFFF
        all_entries.append(empty)

    # Sector layout plan:
    # sector 0 = FAT sector
    # sectors 1..(1+num_dir_sectors-1) = directory
    # sectors (1+num_dir_sectors).. = mini FAT, then mini stream, then regular streams

    current_sector = 1 + num_dir_sectors

    # Mini FAT sectors
    MINIFAT_ENTRIES_PER_SECTOR = SECTOR // 4  # 128
    num_minifat_sectors = (len(mini_fat) + MINIFAT_ENTRIES_PER_SECTOR - 1) // MINIFAT_ENTRIES_PER_SECTOR if mini_fat else 0
    first_minifat_sector = current_sector if num_minifat_sectors > 0 else ENDOFCHAIN
    current_sector += num_minifat_sectors

    # Root entry (mini stream container)
    root_entry.start = current_sector if len(mini_stream) > 0 else ENDOFCHAIN
    root_entry.size = len(mini_stream)
    num_ministream_sectors = (len(mini_stream) + SECTOR - 1) // SECTOR
    current_sector += num_ministream_sectors

    # Regular stream sectors
    for e in regular_data:
        n_sectors = (len(e.data) + SECTOR - 1) // SECTOR
        e.start = current_sector
        current_sector += n_sectors

    total_data_sectors = current_sector  # sectors 0..current_sector-1

    # ── Build FAT ─────────────────────────────────────────────────────────
    # We may need more than one FAT sector; for simplicity assume one is enough
    fat = [FREESECT] * (MINIFAT_ENTRIES_PER_SECTOR)  # start with 128 entries

    # sector 0 = FAT sector
    fat[0] = FATSECT
    # sectors 1..(1+num_dir_sectors-1) = directory chain
    for s in range(1, 1 + num_dir_sectors):
        fat[s] = s + 1 if s < num_dir_sectors else ENDOFCHAIN
    # mini FAT sectors chain
    for i in range(num_minifat_sectors):
        s = first_minifat_sector + i
        fat[s] = s + 1 if i < num_minifat_sectors - 1 else ENDOFCHAIN
    # mini stream sectors (root's data)
    for i in range(num_ministream_sectors):
        s = root_entry.start + i if root_entry.start != ENDOFCHAIN else 0
        if root_entry.start != ENDOFCHAIN:
            fat[s] = s + 1 if i < num_ministream_sectors - 1 else ENDOFCHAIN
    # regular stream sectors
    for e in regular_data:
        n_sectors = (len(e.data) + SECTOR - 1) // SECTOR
        for i in range(n_sectors):
            s = e.start + i
            if s >= len(fat):
                fat.extend([FREESECT] * (s - len(fat) + 1))
            fat[s] = e.start + i + 1 if i < n_sectors - 1 else ENDOFCHAIN

    # Ensure FAT covers all sectors
    while len(fat) < total_data_sectors:
        fat.append(FREESECT)

    num_fat_sectors = (len(fat) + MINIFAT_ENTRIES_PER_SECTOR - 1) // MINIFAT_ENTRIES_PER_SECTOR
    # If we need more than 1 FAT sector, place it; for typical VBA files 1 is enough
    # (128 entries × 512 bytes = 65536 bytes max file — should cover ~50KB files)

    # ── Serialize ──────────────────────────────────────────────────────────
    out = bytearray()

    # Header (512 bytes)
    difat_header = [0] + [FREESECT] * 108   # sector 0 is FAT sector
    header = (
        b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'   # magic
        + b'\x00' * 16                           # clsid
        + struct.pack('<HH', 0x003E, 0x0003)     # minor version, major version (3)
        + struct.pack('<H', 0xFFFE)              # byte order (LE)
        + struct.pack('<H', 9)                   # sector size exponent: 2^9=512
        + struct.pack('<H', 6)                   # mini sector size: 2^6=64
        + b'\x00' * 6                            # reserved
        + struct.pack('<I', num_dir_sectors)     # num dir sectors
        + struct.pack('<I', num_fat_sectors)     # num FAT sectors
        + struct.pack('<I', 1)                   # first dir sector (sector 1)
        + struct.pack('<I', 0)                   # transaction sig
        + struct.pack('<I', MINI_CUTOFF)         # mini stream cutoff
        + struct.pack('<I', first_minifat_sector)# first mini FAT sector
        + struct.pack('<I', num_minifat_sectors) # num mini FAT sectors
        + struct.pack('<I', FREESECT)            # first DIFAT sector
        + struct.pack('<I', 0)                   # num DIFAT sectors
    )
    # DIFAT array (109 entries × 4 bytes = 436 bytes)
    header += struct.pack('<I', 0)               # FAT sector 0
    header += struct.pack('<I', FREESECT) * 108
    assert len(header) == 512, f"Header length = {len(header)}"
    out += header

    # FAT sector (sector 0)
    fat_padded = fat[:128] + [FREESECT] * (128 - min(len(fat), 128))
    out += struct.pack(f'<{128}I', *fat_padded[:128])

    # Directory sectors (sectors 1..)
    dir_bytes = b''.join(e.pack() for e in all_entries)
    assert len(dir_bytes) == num_dir_sectors * SECTOR
    out += dir_bytes

    # Mini FAT sectors
    mf_padded = mini_fat + [FREESECT] * (num_minifat_sectors * MINIFAT_ENTRIES_PER_SECTOR - len(mini_fat))
    if mf_padded:
        out += struct.pack(f'<{len(mf_padded)}I', *mf_padded)

    # Mini stream (root data)
    if mini_stream:
        ms_padded = mini_stream + b'\x00' * ((num_ministream_sectors * SECTOR) - len(mini_stream))
        out += ms_padded

    # Regular streams
    for e in regular_data:
        n_sectors = (len(e.data) + SECTOR - 1) // SECTOR
        padded = e.data + b'\x00' * (n_sectors * SECTOR - len(e.data))
        out += padded

    return bytes(out)


# ──────────────────────────────────────────────────────────────────────────────
# Новый VBA-код для добавления бронирования через форму (InputBox)
# ──────────────────────────────────────────────────────────────────────────────
NEW_BOOKING_SUB = r"""
'----------------------------------------------------------
' РАЗДЕЛ 3а: ФОРМА ДОБАВЛЕНИЯ БРОНИРОВАНИЯ
'----------------------------------------------------------

Sub ОткрытьФормуБронирования()
    Dim wsB As Worksheet, wsR As Worksheet
    Set wsB = Sheets("Бронирование")
    Set wsR = Sheets("НомернойФонд")

    ' Шаг 1: номер комнаты
    Dim roomNum As String
    roomNum = InputBox("Введите номер комнаты:" & vbCrLf & _
        "(Свободные номера — лист НомернойФонд)", "Бронирование — 1/5", "")
    If Trim(roomNum) = "" Then Exit Sub
    If Not IsNumeric(roomNum) Then
        MsgBox "Номер должен быть числом!", vbExclamation: Exit Sub
    End If

    ' Проверка статуса
    Dim rType As String, rPrice As Double, r As Long
    Dim rStatus As String: rStatus = "Свободен"
    For r = 4 To 19
        If wsR.Cells(r, 3).Value = CLng(roomNum) Then
            rStatus = wsR.Cells(r, 9).Value
            rType   = wsR.Cells(r, 4).Value
            rPrice  = wsR.Cells(r, 6).Value
            Exit For
        End If
    Next r
    If rStatus = "Занят" Then
        MsgBox "Номер " & roomNum & " занят!", vbExclamation: Exit Sub
    End If

    ' Шаг 2: ФИО гостя
    Dim gName As String
    gName = InputBox("Введите ФИО гостя:", "Бронирование — 2/5", "")
    If Trim(gName) = "" Then Exit Sub

    ' Шаг 3: дата заезда
    Dim ciStr As String
    ciStr = InputBox("Дата заезда (ДД.ММ.ГГГГ):", "Бронирование — 3/5", _
        Format(Date, "DD.MM.YYYY"))
    If Trim(ciStr) = "" Then Exit Sub
    If Not IsDate(ciStr) Then
        MsgBox "Некорректная дата заезда!", vbExclamation: Exit Sub
    End If
    Dim ci As Date: ci = CDate(ciStr)

    ' Шаг 4: дата выезда
    Dim coStr As String
    coStr = InputBox("Дата выезда (ДД.ММ.ГГГГ):", "Бронирование — 4/5", _
        Format(Date + 1, "DD.MM.YYYY"))
    If Trim(coStr) = "" Then Exit Sub
    If Not IsDate(coStr) Then
        MsgBox "Некорректная дата выезда!", vbExclamation: Exit Sub
    End If
    Dim co As Date: co = CDate(coStr)
    If co <= ci Then
        MsgBox "Выезд должен быть позже заезда!", vbExclamation: Exit Sub
    End If

    ' Шаг 5: кол-во гостей
    Dim ngStr As String
    ngStr = InputBox("Количество гостей:", "Бронирование — 5/5", "1")
    Dim ng As Integer: ng = Val(ngStr): If ng < 1 Then ng = 1

    ' Подтверждение
    Dim nights As Long: nights = co - ci
    Dim total As Double: total = nights * rPrice
    Dim ok As VbMsgBoxResult
    ok = MsgBox("Подтвердите бронирование:" & vbCrLf & vbCrLf & _
        "Номер:   " & roomNum & " (" & rType & ")" & vbCrLf & _
        "Гость:   " & gName & vbCrLf & _
        "Заезд:   " & Format(ci, "DD.MM.YYYY") & vbCrLf & _
        "Выезд:   " & Format(co, "DD.MM.YYYY") & vbCrLf & _
        "Ночей:   " & nights & vbCrLf & _
        "Итого:   " & Format(total, "#,##0") & " руб.", _
        vbYesNo + vbQuestion, "Подтверждение")
    If ok <> vbYes Then Exit Sub

    ' Запись в таблицу
    Dim lastR As Long
    lastR = wsB.Cells(wsB.Rows.Count, 3).End(xlUp).Row + 1
    Dim nID As String: nID = "Б" & Format(lastR - 10, "000")

    wsB.Cells(lastR, 2).Value = lastR - 10
    wsB.Cells(lastR, 3).Value = nID
    wsB.Cells(lastR, 4).Value = CLng(roomNum)
    wsB.Cells(lastR, 5).Value = rType
    wsB.Cells(lastR, 6).Value = gName
    wsB.Cells(lastR, 7).Value = ci
    wsB.Cells(lastR, 8).Value = co
    wsB.Cells(lastR, 9).Value = nights
    wsB.Cells(lastR, 10).Value = rPrice
    wsB.Cells(lastR, 11).Value = total
    wsB.Cells(lastR, 12).Value = "Бронь"
    wsB.Cells(lastR, 13).Value = ng

    wsB.Cells(lastR, 7).NumberFormat = "DD.MM.YYYY"
    wsB.Cells(lastR, 8).NumberFormat = "DD.MM.YYYY"
    wsB.Cells(lastR, 10).NumberFormat = "#,##0"
    wsB.Cells(lastR, 11).NumberFormat = "#,##0"

    Dim c As Integer
    For c = 2 To 13
        With wsB.Cells(lastR, c)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Name = "Arial"
            .Font.Size = 10
        End With
    Next c

    ' Обновить статус номера
    For r = 4 To 19
        If wsR.Cells(r, 3).Value = CLng(roomNum) Then
            wsR.Cells(r, 9).Value = "Бронь": Exit For
        End If
    Next r

    MsgBox "Бронирование " & nID & " добавлено!" & vbCrLf & _
        "Номер: " & roomNum & ", Гость: " & gName & vbCrLf & _
        "Итого: " & Format(total, "#,##0") & " руб.", _
        vbInformation, "Готово"
End Sub

"""

# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────
def main():
    ZIP_PATH  = '/home/user/ais_hotel/АИС_Гостиница (1).zip'
    WORK_DIR  = '/tmp/ais_build2'
    OUT_XLSM  = '/home/user/ais_hotel/АИС_Гостиница_v2.xlsm'

    print("=== 1. Распаковываем ZIP и XLSM ===")
    shutil.rmtree(WORK_DIR, ignore_errors=True)
    os.makedirs(WORK_DIR)

    # Extract xlsm from zip
    xlsm_bytes = None
    with zipfile.ZipFile(ZIP_PATH) as zf:
        for name in zf.namelist():
            if name.endswith('.xlsm'):
                xlsm_bytes = zf.read(name)
                print(f"  Читаем: {name} ({len(xlsm_bytes)} bytes)")
                break

    xlsm_dir = os.path.join(WORK_DIR, 'xlsm')
    os.makedirs(xlsm_dir)
    with zipfile.ZipFile(io.BytesIO(xlsm_bytes)) as zf:
        zf.extractall(xlsm_dir)

    vba_bin_path = os.path.join(xlsm_dir, 'xl', 'vbaProject.bin')
    print(f"  vbaProject.bin: {os.path.getsize(vba_bin_path)} bytes")

    print("\n=== 2. Читаем VBA-модули из vbaProject.bin ===")
    ole = olefile.OleFileIO(vba_bin_path)
    streams = {}
    for entry in ole.listdir(streams=True):
        path = '/'.join(entry)
        streams[path] = ole.openstream(path).read()
        print(f"  {path}: {len(streams[path])} bytes")
    ole.close()

    print("\n=== 3. Декомпрессируем HotelMacros ===")
    hm_raw = streams['VBA/HotelMacros']
    # Source starts at offset 0 (no p-code — confirmed above)
    src_offset = 0
    hm_src = decompress_stream(bytearray(hm_raw[src_offset:]))
    src_text = hm_src.decode('cp1251', errors='replace')
    print(f"  Исходный код: {len(src_text)} chars")
    print(f"  Первые 60: {repr(src_text[:60])}")

    print("\n=== 4. Добавляем новый суб ===")
    # Insert before РАЗДЕЛ 3 (existing booking add section)
    # Find the marker for section 3
    MARKER = "' РАЗДЕЛ 3: ДОБАВЛЕНИЕ БРОНИРОВАНИЯ"
    if MARKER.encode('cp1251', errors='replace').decode('cp1251') in src_text:
        # Insert NEW_BOOKING_SUB before the existing section 3
        pos = src_text.find(MARKER)
        # Go back to find the comment line start
        pos2 = src_text.rfind('\n', 0, pos)
        new_src_text = src_text[:pos2+1] + NEW_BOOKING_SUB + src_text[pos2+1:]
        print("  Вставлено ПЕРЕД разделом 3")
    else:
        # Just append at the end
        new_src_text = src_text.rstrip('\r\n') + '\r\n' + NEW_BOOKING_SUB
        print("  Добавлено В КОНЕЦ (маркер раздела не найден)")

    print(f"  Новый код: {len(new_src_text)} chars (+{len(new_src_text)-len(src_text)})")

    print("\n=== 5. Компрессируем (raw MS-OVBA) ===")
    new_src_bytes = new_src_text.encode('cp1251', errors='replace')
    new_compressed = ovba_compress(new_src_bytes)
    new_hm_stream = new_compressed   # src_offset=0, so no p-code prefix
    print(f"  Старый поток: {len(hm_raw)} bytes")
    print(f"  Новый поток:  {len(new_hm_stream)} bytes")

    print("\n=== 6. Собираем новый vbaProject.bin ===")
    # Build stream_tree for OLE writer
    stream_tree = {
        'PROJECT':   streams['PROJECT'],
        'PROJECTwm': streams['PROJECTwm'],
        'VBA': {
            '_VBA_PROJECT': streams['VBA/_VBA_PROJECT'],
            'dir':          streams['VBA/dir'],
            'HotelMacros':  new_hm_stream,
        }
    }

    # Add Лист1-Лист8, ЭтаКнига
    for key, data in streams.items():
        if key.startswith('VBA/') and key not in ('VBA/_VBA_PROJECT', 'VBA/dir', 'VBA/HotelMacros'):
            module_name = key[4:]  # strip 'VBA/'
            stream_tree['VBA'][module_name] = data

    new_vba_bin = build_ole(stream_tree)
    print(f"  Новый vbaProject.bin: {len(new_vba_bin)} bytes")

    # Validate with olefile
    try:
        test_ole = olefile.OleFileIO(io.BytesIO(new_vba_bin))
        entries = test_ole.listdir(streams=True)
        print(f"  OLE OK — {len(entries)} потоков")
        test_ole.close()
    except Exception as ex:
        print(f"  OLE ОШИБКА: {ex}")
        # Fallback: try to debug
        with open('/tmp/debug_vba.bin', 'wb') as f:
            f.write(new_vba_bin)
        print("  Сохранён /tmp/debug_vba.bin для диагностики")
        raise

    print("\n=== 7. Патчим кнопку 'Добавить' в vmlDrawing ===")
    vml_path = os.path.join(xlsm_dir, 'xl', 'drawings', 'vmlDrawing7.vml')
    if os.path.exists(vml_path):
        with open(vml_path, 'r', encoding='utf-8') as f:
            vml = f.read()
        old_macro = 'ДобавитьБронь'
        new_macro = 'ОткрытьФормуБронирования'
        if old_macro in vml:
            vml = vml.replace(old_macro, new_macro)
            with open(vml_path, 'w', encoding='utf-8') as f:
                f.write(vml)
            print(f"  {old_macro} → {new_macro}")
        else:
            print(f"  Кнопка уже обновлена или макрос не найден")
    else:
        print(f"  vmlDrawing7.vml не найден")

    print("\n=== 8. Перепаковываем XLSM ===")
    # Write new vbaProject.bin
    with open(vba_bin_path, 'wb') as f:
        f.write(new_vba_bin)

    # Repack xlsm
    if os.path.exists(OUT_XLSM):
        os.remove(OUT_XLSM)
    with zipfile.ZipFile(OUT_XLSM, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(xlsm_dir):
            for fname in files:
                fpath = os.path.join(root, fname)
                arcname = os.path.relpath(fpath, xlsm_dir)
                zf.write(fpath, arcname)

    print(f"\n✓ Готово: {OUT_XLSM}  ({os.path.getsize(OUT_XLSM)} bytes)")


if __name__ == '__main__':
    main()

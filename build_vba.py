#!/usr/bin/env python3
"""
build_vba.py  — inject ОткрытьФормуБронирования into existing VBA project.

Strategy: patch the ORIGINAL vbaProject.bin from the working .xlsm file.
Instead of rebuilding OLE from scratch (which breaks Excel), we:
  1. Parse the original OLE structure (FAT, directory, mini-stream)
  2. Replace only the HotelMacros stream data (expanding if needed)
  3. Update FAT chains and directory entry size
  4. Preserve everything else byte-for-byte

This guarantees Excel compatibility since 95% of the binary is untouched.
"""

import struct, os, shutil, zipfile, io, math, copy
import olefile
from oletools.olevba import decompress_stream

SECTOR = 512
MINI_SECTOR = 64
FREESECT   = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT    = 0xFFFFFFFD
MINI_CUTOFF = 4096

# ──────────────────────────────────────────────────────────────────────────────
# MS-OVBA compression (LZ77)
# ──────────────────────────────────────────────────────────────────────────────
def _copy_token_fields(pos: int):
    difference = max(pos, 1)
    ob = max(math.ceil(math.log2(difference)), 4)
    lb = 16 - ob
    length_mask = 0xFFFF >> ob
    max_len = length_mask + 3
    max_offset = 1 << ob
    return ob, lb, length_mask, max_len, max_offset

def _find_match(chunk: bytes, pos: int):
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
    out = bytearray([0x01])
    i = 0
    while i < len(source):
        chunk = source[i: i + 4096]
        i += 4096
        compressed = _compress_chunk(chunk)
        if len(compressed) + 3 > 4098 or len(compressed) >= len(chunk):
            padded = chunk + b'\x00' * (4096 - len(chunk))
            hdr = struct.pack('<H', 0x3FFD)
            out += hdr + padded
        else:
            hdr = struct.pack('<H', 0xB000 | (len(compressed) - 1))
            out += hdr + compressed
    return bytes(out)


# ──────────────────────────────────────────────────────────────────────────────
# OLE patching: replace a stream in an existing valid OLE file
# ──────────────────────────────────────────────────────────────────────────────
def patch_ole_stream(ole_data: bytes, stream_path: str, new_data: bytes) -> bytes:
    """
    Replace the data of `stream_path` (e.g. 'VBA/HotelMacros') in the OLE
    compound file `ole_data`, returning the modified file bytes.
    Handles expansion (new data larger than old) by appending new sectors.
    """
    buf = bytearray(ole_data)

    # ── Parse header ─────────────────────────────────────────────────────
    sector_pow = struct.unpack_from('<H', buf, 30)[0]
    sector_size = 1 << sector_pow
    assert sector_size == 512, f"Only 512-byte sectors supported, got {sector_size}"

    num_fat_sectors = struct.unpack_from('<I', buf, 44)[0]
    first_dir_sector = struct.unpack_from('<I', buf, 48)[0]
    mini_cutoff = struct.unpack_from('<I', buf, 56)[0]
    first_minifat_sector = struct.unpack_from('<I', buf, 60)[0]
    num_minifat_sectors = struct.unpack_from('<I', buf, 64)[0]

    # DIFAT: first 109 entries in header at offset 76
    difat = []
    for i in range(109):
        v = struct.unpack_from('<I', buf, 76 + i * 4)[0]
        if v not in (FREESECT, ENDOFCHAIN):
            difat.append(v)

    def sec_offset(sec_id):
        return 512 + sec_id * sector_size

    # ── Read FAT ─────────────────────────────────────────────────────────
    fat = []
    for fat_sec in difat[:num_fat_sectors]:
        off = sec_offset(fat_sec)
        for i in range(sector_size // 4):
            fat.append(struct.unpack_from('<I', buf, off + i * 4)[0])

    def get_chain(start):
        chain = []
        s = start
        seen = set()
        while s not in (ENDOFCHAIN, FREESECT) and s < len(fat):
            if s in seen:
                break
            seen.add(s)
            chain.append(s)
            s = fat[s]
        return chain

    def write_fat_entry(idx, val):
        fat[idx] = val
        fat_sec_idx = idx // (sector_size // 4)
        entry_in_sec = idx % (sector_size // 4)
        if fat_sec_idx < len(difat):
            fat_sec = difat[fat_sec_idx]
            off = sec_offset(fat_sec) + entry_in_sec * 4
            if off + 4 <= len(buf):
                struct.pack_into('<I', buf, off, val)

    def alloc_sector():
        """Find a free sector in existing FAT, or extend."""
        for i in range(len(fat)):
            if fat[i] == FREESECT:
                # Make sure sector data area exists
                needed = sec_offset(i) + sector_size
                if needed > len(buf):
                    buf.extend(b'\x00' * (needed - len(buf)))
                return i
        # All 128 entries used — append (shouldn't happen for small files)
        new_id = len(fat)
        fat.append(FREESECT)
        buf.extend(b'\x00' * sector_size)
        return new_id

    # ── Read directory ───────────────────────────────────────────────────
    dir_chain = get_chain(first_dir_sector)
    dir_bytes = bytearray()
    for s in dir_chain:
        dir_bytes += buf[sec_offset(s):sec_offset(s) + sector_size]

    DIR_ENTRY_SIZE = 128
    num_entries = len(dir_bytes) // DIR_ENTRY_SIZE

    def read_dir_entry(idx):
        off = idx * DIR_ENTRY_SIZE
        name_raw = dir_bytes[off:off + 64]
        name_len = struct.unpack_from('<H', dir_bytes, off + 64)[0]
        name = name_raw[:max(0, name_len - 2)].decode('utf-16-le', errors='replace')
        obj_type = dir_bytes[off + 66]
        child = struct.unpack_from('<I', dir_bytes, off + 76)[0]
        left = struct.unpack_from('<I', dir_bytes, off + 68)[0]
        right = struct.unpack_from('<I', dir_bytes, off + 72)[0]
        start = struct.unpack_from('<I', dir_bytes, off + 116)[0]
        size = struct.unpack_from('<I', dir_bytes, off + 120)[0]
        return {
            'name': name, 'type': obj_type, 'child': child,
            'left': left, 'right': right,
            'start': start, 'size': size, 'dir_offset': off
        }

    # ── Find the target stream by walking the directory tree ─────────────
    parts = stream_path.split('/')
    target_entry = None

    def find_entry(node_idx, name):
        if node_idx == 0xFFFFFFFF or node_idx >= num_entries:
            return None
        e = read_dir_entry(node_idx)
        if e['name'].lower() == name.lower():
            return e
        result = find_entry(e['left'], name)
        if result:
            return result
        return find_entry(e['right'], name)

    # Navigate: root → storage → stream
    root = read_dir_entry(0)
    current_child = root['child']
    for i, part in enumerate(parts):
        entry = find_entry(current_child, part)
        if entry is None:
            raise ValueError(f"Stream path '{stream_path}' not found at part '{part}'")
        if i < len(parts) - 1:
            current_child = entry['child']
        else:
            target_entry = entry

    if target_entry is None:
        raise ValueError(f"Stream '{stream_path}' not found")

    old_size = target_entry['size']
    old_start = target_entry['start']
    new_size = len(new_data)

    print(f"  Патчим {stream_path}: {old_size} → {new_size} bytes")

    is_mini = old_size < mini_cutoff and old_size > 0

    if is_mini:
        # Stream is in mini-stream — need to handle mini FAT and mini stream
        # Read mini FAT
        mini_fat = []
        if first_minifat_sector != ENDOFCHAIN:
            mf_chain = get_chain(first_minifat_sector)
            for s in mf_chain:
                off = sec_offset(s)
                for i in range(sector_size // 4):
                    mini_fat.append(struct.unpack_from('<I', buf, off + i * 4)[0])

        # Read mini stream (root entry's data)
        root_start = root['start']
        root_size = root['size']
        ms_chain = get_chain(root_start)
        mini_stream = bytearray()
        for s in ms_chain:
            mini_stream += buf[sec_offset(s):sec_offset(s) + sector_size]
        mini_stream = mini_stream[:root_size]

        if new_size < mini_cutoff:
            # Still fits in mini-stream
            old_mini_chain = []
            ms = old_start
            seen = set()
            while ms not in (ENDOFCHAIN, FREESECT) and ms < len(mini_fat):
                if ms in seen:
                    break
                seen.add(ms)
                old_mini_chain.append(ms)
                ms = mini_fat[ms]

            old_mini_sectors = len(old_mini_chain)
            new_mini_sectors = (new_size + MINI_SECTOR - 1) // MINI_SECTOR

            if new_mini_sectors <= old_mini_sectors:
                # Fits in existing mini-sectors
                for i, ms_idx in enumerate(old_mini_chain[:new_mini_sectors]):
                    chunk_start = i * MINI_SECTOR
                    chunk_end = min((i + 1) * MINI_SECTOR, new_size)
                    data_chunk = new_data[chunk_start:chunk_end]
                    ms_off = ms_idx * MINI_SECTOR
                    mini_stream[ms_off:ms_off + len(data_chunk)] = data_chunk
                    if len(data_chunk) < MINI_SECTOR:
                        mini_stream[ms_off + len(data_chunk):ms_off + MINI_SECTOR] = b'\x00' * (MINI_SECTOR - len(data_chunk))
                # Free extra mini-sectors
                for i in range(new_mini_sectors, old_mini_sectors):
                    mini_fat[old_mini_chain[i]] = FREESECT
                if new_mini_sectors > 0:
                    mini_fat[old_mini_chain[new_mini_sectors - 1]] = ENDOFCHAIN
            else:
                # Need more mini-sectors — extend mini-stream
                # Use existing chain, then allocate new mini-sectors
                for i, ms_idx in enumerate(old_mini_chain):
                    chunk_start = i * MINI_SECTOR
                    chunk_end = min((i + 1) * MINI_SECTOR, new_size)
                    data_chunk = new_data[chunk_start:chunk_end]
                    ms_off = ms_idx * MINI_SECTOR
                    mini_stream[ms_off:ms_off + len(data_chunk)] = data_chunk

                # Allocate additional mini-sectors
                prev = old_mini_chain[-1] if old_mini_chain else None
                for i in range(old_mini_sectors, new_mini_sectors):
                    new_ms_idx = len(mini_stream) // MINI_SECTOR
                    chunk_start = i * MINI_SECTOR
                    chunk_end = min((i + 1) * MINI_SECTOR, new_size)
                    data_chunk = new_data[chunk_start:chunk_end]
                    padded = data_chunk + b'\x00' * (MINI_SECTOR - len(data_chunk))
                    mini_stream += padded
                    # Extend mini FAT
                    while len(mini_fat) <= new_ms_idx:
                        mini_fat.append(FREESECT)
                    mini_fat[new_ms_idx] = ENDOFCHAIN
                    if prev is not None:
                        mini_fat[prev] = new_ms_idx
                    prev = new_ms_idx

            # Write mini-stream back to root's sectors
            new_ms_size = len(mini_stream)
            new_ms_sectors_needed = (new_ms_size + sector_size - 1) // sector_size
            old_ms_sectors = len(ms_chain)

            if new_ms_sectors_needed > old_ms_sectors:
                # Need to extend root's sector chain
                last_sec = ms_chain[-1]
                for _ in range(new_ms_sectors_needed - old_ms_sectors):
                    new_sec = alloc_sector()
                    write_fat_entry(new_sec, ENDOFCHAIN)
                    write_fat_entry(last_sec, new_sec)
                    ms_chain.append(new_sec)
                    last_sec = new_sec

            # Write mini-stream data to sectors
            padded_ms = mini_stream + b'\x00' * (new_ms_sectors_needed * sector_size - len(mini_stream))
            for i, s in enumerate(ms_chain[:new_ms_sectors_needed]):
                off = sec_offset(s)
                buf[off:off + sector_size] = padded_ms[i * sector_size:(i + 1) * sector_size]

            # Update root size in directory
            root_dir_off = 0  # root is always entry 0
            struct.pack_into('<I', dir_bytes, root_dir_off + 120, new_ms_size)

            # Write mini FAT back
            if first_minifat_sector != ENDOFCHAIN:
                mf_chain = get_chain(first_minifat_sector)
                mf_padded = mini_fat + [FREESECT] * (len(mf_chain) * (sector_size // 4) - len(mini_fat))
                for i, s in enumerate(mf_chain):
                    off = sec_offset(s)
                    for j in range(sector_size // 4):
                        idx = i * (sector_size // 4) + j
                        if idx < len(mf_padded):
                            struct.pack_into('<I', buf, off + j * 4, mf_padded[idx])

        else:
            # Was mini, now needs regular sectors — complex case
            # For simplicity, raise — this shouldn't happen for our use case
            raise NotImplementedError("Mini→regular sector transition not implemented")

    else:
        # Regular sectors (>= MINI_CUTOFF or was already in regular)
        old_chain = get_chain(old_start)
        old_sectors = len(old_chain)
        new_sectors_needed = (new_size + sector_size - 1) // sector_size

        if new_sectors_needed <= old_sectors:
            # Write into existing sectors
            for i, s in enumerate(old_chain[:new_sectors_needed]):
                off = sec_offset(s)
                data_start = i * sector_size
                data_end = min((i + 1) * sector_size, new_size)
                chunk = new_data[data_start:data_end]
                padded = chunk + b'\x00' * (sector_size - len(chunk))
                buf[off:off + sector_size] = padded
            # Free unused sectors
            for i in range(new_sectors_needed, old_sectors):
                write_fat_entry(old_chain[i], FREESECT)
            if new_sectors_needed > 0:
                write_fat_entry(old_chain[new_sectors_needed - 1], ENDOFCHAIN)
        else:
            # Write into existing + append new sectors
            for i, s in enumerate(old_chain):
                off = sec_offset(s)
                data_start = i * sector_size
                data_end = min((i + 1) * sector_size, new_size)
                chunk = new_data[data_start:data_end]
                padded = chunk + b'\x00' * (sector_size - len(chunk))
                buf[off:off + sector_size] = padded

            # Allocate new sectors from free pool
            last_sec = old_chain[-1] if old_chain else None
            for i in range(old_sectors, new_sectors_needed):
                new_sec = alloc_sector()

                # Write data into the sector
                data_start = i * sector_size
                data_end = min((i + 1) * sector_size, new_size)
                chunk = new_data[data_start:data_end]
                padded = chunk + b'\x00' * (sector_size - len(chunk))
                off = sec_offset(new_sec)
                buf[off:off + sector_size] = padded

                # Update FAT: mark as end of chain
                write_fat_entry(new_sec, ENDOFCHAIN)

                # Link from previous sector
                if last_sec is not None:
                    write_fat_entry(last_sec, new_sec)

                last_sec = new_sec

    # ── Update directory entry size ──────────────────────────────────────
    target_dir_off = target_entry['dir_offset']
    struct.pack_into('<I', dir_bytes, target_dir_off + 120, new_size)

    # ── Write modified directory back to sectors ─────────────────────────
    for i, s in enumerate(dir_chain):
        off = sec_offset(s)
        buf[off:off + sector_size] = dir_bytes[i * sector_size:(i + 1) * sector_size]

    return bytes(buf)


# ──────────────────────────────────────────────────────────────────────────────
# New VBA sub for booking form
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
    WORK_DIR  = '/tmp/ais_build3'
    OUT_XLSM  = '/home/user/ais_hotel/АИС_Гостиница_v2.xlsm'

    print("=== 1. Распаковываем ZIP и XLSM ===")
    shutil.rmtree(WORK_DIR, ignore_errors=True)
    os.makedirs(WORK_DIR)

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
    orig_vba_size = os.path.getsize(vba_bin_path)
    print(f"  vbaProject.bin: {orig_vba_size} bytes")

    print("\n=== 2. Читаем VBA-код HotelMacros ===")
    ole = olefile.OleFileIO(vba_bin_path)
    hm_raw = ole.openstream('VBA/HotelMacros').read()
    ole.close()
    print(f"  HotelMacros stream: {len(hm_raw)} bytes")

    hm_src = decompress_stream(bytearray(hm_raw))
    src_text = hm_src.decode('cp1251', errors='replace')
    print(f"  Исходный код: {len(src_text)} chars")

    print("\n=== 3. Добавляем ОткрытьФормуБронирования ===")
    if 'ОткрытьФормуБронирования' in src_text:
        print("  Суб уже существует — пропускаем")
        new_src_text = src_text
    else:
        MARKER = "' РАЗДЕЛ 3: ДОБАВЛЕНИЕ БРОНИРОВАНИЯ"
        if MARKER in src_text:
            pos = src_text.find(MARKER)
            pos2 = src_text.rfind('\n', 0, pos)
            new_src_text = src_text[:pos2+1] + NEW_BOOKING_SUB + src_text[pos2+1:]
            print("  Вставлено ПЕРЕД разделом 3")
        else:
            new_src_text = src_text.rstrip('\r\n') + '\r\n' + NEW_BOOKING_SUB
            print("  Добавлено В КОНЕЦ")

    print(f"  Новый код: {len(new_src_text)} chars (+{len(new_src_text)-len(src_text)})")

    print("\n=== 4. Компрессируем (MS-OVBA LZ77) ===")
    new_src_bytes = new_src_text.encode('cp1251', errors='replace')
    new_compressed = ovba_compress(new_src_bytes)
    print(f"  Оригинал: {len(hm_raw)} bytes → Новый: {len(new_compressed)} bytes")

    # Verify roundtrip
    rt = decompress_stream(bytearray(new_compressed))
    rt_text = rt.decode('cp1251', errors='replace')
    assert 'ОткрытьФормуБронирования' in rt_text, "Roundtrip failed!"
    print(f"  Roundtrip OK ({len(rt_text)} chars)")

    print("\n=== 5. Патчим vbaProject.bin (оригинальный OLE) ===")
    with open(vba_bin_path, 'rb') as f:
        orig_vba = f.read()

    patched_vba = patch_ole_stream(orig_vba, 'VBA/HotelMacros', new_compressed)
    print(f"  Результат: {len(patched_vba)} bytes (было {len(orig_vba)})")

    # Validate with olefile
    test_ole = olefile.OleFileIO(io.BytesIO(patched_vba))
    entries = test_ole.listdir(streams=True)
    print(f"  OLE валидация: {len(entries)} потоков")
    # Verify HotelMacros is readable
    hm_test = test_ole.openstream('VBA/HotelMacros').read()
    hm_test_src = decompress_stream(bytearray(hm_test))
    hm_test_text = hm_test_src.decode('cp1251', errors='replace')
    assert 'ОткрытьФормуБронирования' in hm_test_text, "Patched stream is corrupted!"
    print(f"  HotelMacros OK: {len(hm_test_text)} chars, 'ОткрытьФормуБронирования' present")
    test_ole.close()

    # Write patched vbaProject.bin
    with open(vba_bin_path, 'wb') as f:
        f.write(patched_vba)

    print("\n=== 6. Патчим кнопку 'Добавить' в vmlDrawing ===")
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
        elif new_macro in vml:
            print(f"  Кнопка уже привязана к {new_macro}")
        else:
            print(f"  Макрос кнопки не найден")
    else:
        print(f"  vmlDrawing7.vml не найден")

    print("\n=== 7. Перепаковываем XLSM ===")
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

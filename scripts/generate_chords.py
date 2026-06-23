#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
コード表.xlsx からウクレレ(4弦)のコードダイアグラム画像を自動生成するスクリプト。

仕組み:
  - Excelの各シート(ルート音ごと)には複数のコードブロックが並んでいる。
  - 各ブロックは「コード名セル」+「フレット番号行(1,2,3,4,5)」+ グリッド枠(4弦x5フレット)で構成。
  - 実際の押弦位置(○)・セーハ(縦長の角丸長方形)は、xlsxパッケージ内の
    xl/drawings/drawingN.xml に図形(ellipse / flowChartTerminator)として
    セル座標(col/row + EMUオフセット)で記録されている。
  - openpyxlはオートシェイプを読めないため、drawing XMLを直接パースする。

出力:
  - chords/<コード名>.png を生成(ファイル名はOSで安全な文字に変換)
"""

import os
import re
import json
import zipfile
import shutil
import unicodedata
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont

XLSX_PATH = os.environ.get("CHORD_XLSX", "コード表.xlsx")
OUT_DIR = os.environ.get("CHORD_OUT_DIR", "docs/chords")
WORK_DIR = "_xlsx_extract_tmp"

NUM_STRINGS = 4   # ウクレレ
NUM_FRETS = 5     # 表示するフレット数


# ----------------------------------------------------------------------------
# 1. xlsxを展開してdrawing XMLを取得
# ----------------------------------------------------------------------------
def extract_xlsx(xlsx_path, work_dir):
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir, exist_ok=True)
    with zipfile.ZipFile(xlsx_path) as z:
        z.extractall(work_dir)
    return work_dir


def get_sheet_to_drawing_map(work_dir, sheet_names_in_order):
    """workbook.xmlのsheet順とsheetN.xmlの対応、さらにdrawing relsを辿る"""
    # workbook.xml.rels: rId -> sheetN.xml
    with open(os.path.join(work_dir, "xl", "_rels", "workbook.xml.rels"), encoding="utf-8") as f:
        wb_rels = f.read()
    rid_to_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', wb_rels))

    with open(os.path.join(work_dir, "xl", "workbook.xml"), encoding="utf-8") as f:
        wb_xml = f.read()
    # <sheet name="C" sheetId="1" r:id="rId1"/>
    sheet_entries = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb_xml)

    name_to_sheetfile = {}
    for name, rid in sheet_entries:
        target = rid_to_target.get(rid)
        if target:
            target = target.replace("worksheets/", "")
            name_to_sheetfile[name] = target  # e.g. sheet1.xml

    name_to_drawing = {}
    for name, sheetfile in name_to_sheetfile.items():
        rels_path = os.path.join(work_dir, "xl", "worksheets", "_rels", sheetfile + ".rels")
        if not os.path.exists(rels_path):
            continue
        with open(rels_path, encoding="utf-8") as f:
            rels = f.read()
        m = re.search(r'Type="[^"]*drawing"[^>]*Target="([^"]+)"', rels)
        if m:
            drawing_target = m.group(1).replace("../drawings/", "")
            name_to_drawing[name] = drawing_target  # e.g. drawing1.xml
    return name_to_drawing


# ----------------------------------------------------------------------------
# 2. drawing XMLから図形(ellipse / flowChartTerminator)を抽出
# ----------------------------------------------------------------------------
def parse_shapes(drawing_path):
    with open(drawing_path, encoding="utf-8") as f:
        content = f.read()

    shapes = []
    anchors = re.findall(r"<xdr:twoCellAnchor[^>]*>(.*?)</xdr:twoCellAnchor>", content, re.S)
    for block in anchors:
        m = re.search(
            r"<xdr:from><xdr:col>(\d+)</xdr:col><xdr:colOff>(\d+)</xdr:colOff>"
            r"<xdr:row>(\d+)</xdr:row><xdr:rowOff>(\d+)</xdr:rowOff></xdr:from>"
            r"<xdr:to><xdr:col>(\d+)</xdr:col><xdr:colOff>(\d+)</xdr:colOff>"
            r"<xdr:row>(\d+)</xdr:row><xdr:rowOff>(\d+)</xdr:rowOff></xdr:to>",
            block,
        )
        if not m:
            continue
        fc, fco, fr, fro, tc, tco, tr, tro = map(int, m.groups())
        prst_m = re.search(r'prst="(\w+)"', block)
        prst = prst_m.group(1) if prst_m else None
        if prst not in ("ellipse", "flowChartTerminator"):
            continue
        rot = bool(re.search(r'<a:xfrm rot="\d+"', block))
        shapes.append(
            {
                "type": prst,
                "from_col": fc,
                "from_row": fr,
                "to_col": tc,
                "to_row": tr,
                "rotated": rot,
            }
        )
    return shapes


# ----------------------------------------------------------------------------
# 3. ワークシートからコードブロック(コード名 + グリッド位置)を検出
# ----------------------------------------------------------------------------
def find_chord_blocks(ws):
    blocks = []
    max_row = ws.max_row
    max_col = ws.max_column

    for row in range(1, max_row + 1):
        for col in range(1, max_col - NUM_FRETS + 2):
            vals = [ws.cell(row=row, column=col + i).value for i in range(NUM_FRETS)]
            # 「1,2,3,4,5」だけでなく、ハイポジション用の「10,11,12,13,14」のような
            # 任意開始値の連続する整数列もフレット数字行として認識する。
            if (
                all(isinstance(v, int) for v in vals)
                and vals == list(range(vals[0], vals[0] + NUM_FRETS))
                and vals[0] >= 1
            ):
                start_fret = vals[0]

                # コード名を探す(フレット行の上、同じ列付近)
                name = None
                for r2 in range(row - 1, max(row - 8, 0), -1):
                    for c2 in (col, col - 1, col + 1, col - 2, col + 2):
                        if c2 < 1:
                            continue
                        v = ws.cell(row=r2, column=c2).value
                        if isinstance(v, str) and v.strip():
                            name = v.strip()
                            break
                    if name:
                        break
                if name is None:
                    continue  # 空テンプレ枠はスキップ

                # グリッド本体はフレット数字行の「上」にある。
                # 罫線は NUM_STRINGS行分(弦と弦の境界線がNUM_STRINGS本)、
                # フレット数字行の直前の行までがグリッドの下端境界。
                grid_row1_top = row - (NUM_STRINGS + 1)  # 1-indexed: グリッド最上端の行

                blocks.append(
                    {
                        "name": name,
                        "fret_row": row,
                        "start_col0": col - 1,  # 0-indexed: フレット1列目の列
                        "grid_row0_top": grid_row1_top - 1,  # 0-indexed top row
                        "start_fret": start_fret,  # このグリッドの最初の列が実際は何フレットか
                    }
                )
    return blocks


def shapes_in_block(shapes, block):
    """ブロックのグリッド範囲(0-indexed col/row)に属する図形を抜き出し、
    (string_index 1-4, fret_index 1-5, type) に変換する。"""
    col0 = block["start_col0"]  # 0-indexed col of フレット1
    row0 = block["grid_row0_top"]  # 0-indexed top row of grid (弦1の上端境界線)

    results = []
    for s in shapes:
        fret = s["from_col"] - col0  # 1,2,3,4,5
        if not (1 <= fret <= NUM_FRETS):
            continue
        if s["type"] == "flowChartTerminator":
            # セーハ: from_row~to_row が複数弦にわたる
            str_from = s["from_row"] - row0 + 1
            str_to = s["to_row"] - row0
            if str_from < 1 or str_to > NUM_STRINGS:
                continue
            results.append({"type": "bar", "fret": fret, "string_from": str_from, "string_to": str_to})
        else:
            string_idx = s["from_row"] - row0 + 1
            if not (1 <= string_idx <= NUM_STRINGS):
                continue
            results.append({"type": "dot", "fret": fret, "string": string_idx})
    return results


# ----------------------------------------------------------------------------
# 4. 画像描画
# ----------------------------------------------------------------------------
def safe_filename(name):
    # ファイル名に使えない文字を置換 (# / は特に注意)
    replacements = {
        "#": "s", "♭": "b", "/": "_on_", "\\": "_",
        "?": "", "*": "", ":": "-", '"': "", "<": "", ">": "", "|": "",
    }
    out = name
    for k, v in replacements.items():
        out = out.replace(k, v)
    return out


def draw_chord_image(name, shapes_data, start_fret=1, size=(166, 163)):
    W, H = size
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))  # 背景は完全透明
    draw = ImageDraw.Draw(img)

    WHITE = (255, 255, 255, 255)  # 不透明な白(線・文字・丸)

    # レイアウト定数(元画像の比率を再現)
    title_h = int(H * 0.58)
    grid_top = title_h
    grid_bottom = H - 1
    grid_right = W - 1

    is_open_position = start_fret == 1

    # 0フレット(ナット)を表す太い縦線の幅を左端に確保する。
    # ハイポジション(start_fret != 1)の場合はナットの太線は引かず、
    # 代わりに開始フレット番号を表示するための余白を確保する。
    if is_open_position:
        nut_w = max(3, int(W * 0.045))
    else:
        nut_w = max(1, int(W * 0.009))
    fret_label_w = int(W * 0.10) if not is_open_position else 0
    grid_left = nut_w + fret_label_w

    grid_h = grid_bottom - grid_top
    grid_w = grid_right - grid_left

    fret_w = grid_w / NUM_FRETS
    # 弦(横線)は NUM_STRINGS 本。弦と弦の間隔は (NUM_STRINGS-1) 個分。
    string_gap = grid_h / (NUM_STRINGS - 1)

    # タイトル(コード名)
    font_path = None
    for fp in (
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Black.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    ):
        if os.path.exists(fp):
            font_path = fp
            break

    text = name
    max_text_w = W - int(W * 0.06)  # 左右マージン分を確保
    font_size = int(title_h * 0.6)
    min_font_size = int(title_h * 0.28)
    while font_size > min_font_size:
        font = ImageFont.truetype(font_path, font_size, index=0) if font_path else ImageFont.load_default()
        bbox = draw.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        if tw <= max_text_w:
            break
        font_size -= 2
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    margin_left = int(W * 0.03)
    draw.text((margin_left - bbox[0], (title_h - th) / 2 - bbox[1]), text, fill=WHITE, font=font)

    # ナット(0フレット)の線(左端)。オープンポジションのみ太線にする。
    draw.line([(grid_left, grid_top), (grid_left, grid_bottom)], fill=WHITE, width=nut_w)

    # ハイポジションの場合、グリッド左に開始フレット番号を小さく表示する
    if not is_open_position:
        label_font_size = max(10, int(string_gap * 0.9))
        label_font = (
            ImageFont.truetype(font_path, label_font_size, index=0)
            if font_path
            else ImageFont.load_default()
        )
        label_text = str(start_fret)
        lbbox = draw.textbbox((0, 0), label_text, font=label_font)
        lw, lh = lbbox[2] - lbbox[0], lbbox[3] - lbbox[1]
        label_x = nut_w + (fret_label_w - lw) / 2 - lbbox[0]
        label_y = grid_top + (string_gap - lh) / 2 - lbbox[1]
        draw.text((label_x, label_y), label_text, fill=WHITE, font=label_font)

    # 縦線(フレットの区切り線): ナットの右、フレット1〜5の右端まで
    line_w = max(1, int(W * 0.009))
    for i in range(1, NUM_FRETS + 1):
        x = grid_left + i * fret_w
        draw.line([(x, grid_top), (x, grid_bottom)], fill=WHITE, width=line_w)

    # 横線(弦): NUM_STRINGS本、等間隔
    for j in range(NUM_STRINGS):
        y = grid_top + j * string_gap
        draw.line([(grid_left, y), (grid_right, y)], fill=WHITE, width=line_w)

    # セーハ・丸 (string番号 = 上から何本目の弦(線)か。1始まり)
    for sd in shapes_data:
        fret = sd["fret"]
        cx = grid_left + (fret - 0.5) * fret_w
        if sd["type"] == "dot":
            string = sd["string"]
            cy = grid_top + (string - 1) * string_gap
            r = min(fret_w, string_gap) * 0.32
            draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=WHITE)
        else:  # bar (セーハ): string_from弦からstring_to弦までを結ぶ縦長カプセル
            y_top = grid_top + (sd["string_from"] - 1) * string_gap
            y_bot = grid_top + (sd["string_to"] - 1) * string_gap
            r = fret_w * 0.30
            pad = r * 0.5
            draw.rounded_rectangle(
                [cx - r, y_top - pad, cx + r, y_bot + pad], radius=r, fill=WHITE
            )

    return img


# ----------------------------------------------------------------------------
# メイン処理
# ----------------------------------------------------------------------------
def main():
    work_dir = extract_xlsx(XLSX_PATH, WORK_DIR)
    wb = load_workbook(XLSX_PATH, data_only=False)

    name_to_drawing = get_sheet_to_drawing_map(work_dir, wb.sheetnames)

    os.makedirs(OUT_DIR, exist_ok=True)

    manifest = []
    generated = 0
    skipped_sheets = []
    removed_exact_dups = []
    kept_variant_dups = []

    # シート内の同名コードをグルーピングし、重複を判定してから描画する
    for sheet_name in wb.sheetnames:
        if sheet_name not in name_to_drawing:
            skipped_sheets.append(sheet_name)
            continue
        ws = wb[sheet_name]
        drawing_file = name_to_drawing[sheet_name]
        drawing_path = os.path.join(work_dir, "xl", "drawings", drawing_file)
        if not os.path.exists(drawing_path):
            skipped_sheets.append(sheet_name)
            continue

        shapes = parse_shapes(drawing_path)
        blocks = find_chord_blocks(ws)

        # 同名コードごとにグルーピング
        by_name = {}
        for block in blocks:
            by_name.setdefault(block["name"], []).append(block)

        used_filenames_in_sheet = {}

        for name, block_list in by_name.items():
            # 各ブロックの押弦データを計算し、内容で重複排除する
            seen_signatures = []
            unique_blocks = []  # (block, shapes_data)
            for block in block_list:
                sd = shapes_in_block(shapes, block)
                signature = (
                    block["start_fret"],
                    tuple(sorted(tuple(sorted(d.items())) for d in sd)),
                )
                if signature in seen_signatures:
                    # 完全に同じ運指の重複 → 捨てる
                    removed_exact_dups.append(f"{sheet_name}シート「{name}」(row={block['fret_row']})")
                    continue
                seen_signatures.append(signature)
                unique_blocks.append((block, sd))

            if len(unique_blocks) > 1:
                kept_variant_dups.append(
                    f"{sheet_name}シート「{name}」({len(unique_blocks)}種類の異なる運指)"
                )

            for block, sd in unique_blocks:
                img = draw_chord_image(name, sd, start_fret=block["start_fret"])
                base_fname = safe_filename(name) + ".png"
                if base_fname in used_filenames_in_sheet:
                    used_filenames_in_sheet[base_fname] += 1
                    root, ext = os.path.splitext(base_fname)
                    fname = f"{root}_{used_filenames_in_sheet[base_fname]}{ext}"
                else:
                    used_filenames_in_sheet[base_fname] = 1
                    fname = base_fname

                out_path = os.path.join(OUT_DIR, fname)
                img.save(out_path)
                manifest.append({"name": name, "file": fname, "sheet": sheet_name})
                generated += 1

    with open(os.path.join(OUT_DIR, "..", "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"生成枚数: {generated}")
    if skipped_sheets:
        print(f"スキップしたシート(図形なし/未対応): {skipped_sheets}")
    if removed_exact_dups:
        print("\n[自動排除] 完全に同じ運指の重複コードを除外しました:")
        for d in removed_exact_dups:
            print("  - " + d)
    if kept_variant_dups:
        print("\n[情報] 同名だが運指が異なるコードが見つかったため、両方を残しました(要確認):")
        for d in kept_variant_dups:
            print("  - " + d)


if __name__ == "__main__":
    main()

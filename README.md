# make_grid

A4横向きWordファイルに5列×4行（計20マス）の番号付きグリッドを生成するスクリプト。

## 使い方

```bash
python make_grid.py
```

実行すると同じフォルダに `grid_20.docx` が出力される。

## 設定値（make_grid.py 冒頭の定数）

| 定数 | デフォルト | 説明 |
|---|---|---|
| `COLS` | 5 | 列数 |
| `ROWS` | 4 | 行数 |
| `PAGE_WIDTH_CM` | 29.7 | ページ幅（cm、A4横） |
| `PAGE_HEIGHT_CM` | 21.0 | ページ高さ（cm、A4横） |
| `MARGIN_CM` | 1.0 | 四辺の余白（cm） |
| `CELL_NUMBER_FONT_SIZE` | 8 | セル番号のフォントサイズ（pt） |
| `CELL_MARGIN_DXA` | 80 | セル内余白（DXA単位） |

## 必要ライブラリ

```bash
pip install python-docx
```

## テスト

```bash
pip install pytest
python -m pytest test_make_grid.py -v
```

# docxファイルをpdf化し、結合して一つのpdfにするやつ

macOS・Windows向け（どちらもMicrosoft Wordのインストールが必要）。uvでの手順はWindows 11で動作確認済み、macOSでは未確認。

## 必要なもの

- Microsoft Word（変換はWordが行う）
- [uv](https://docs.astral.sh/uv/)（Pythonの本体と依存パッケージはuvが用意する）

## 使い方

1. `./data/{適当な名前}/`に一括ダウンロードしたフォルダ（解凍済み）を配置
2. 以下をターミナルで実行（途中でWordへのアクセス許可が求められることがある）
    ```
    uv sync
    uv run main.py -i ./data/{適当な名前}/
    ```
3. `./out.pdf`が作成される

OSごとの依存（macOSは`appscript`、Windowsは`pywin32`）は`uv.lock`に記録されており、`uv sync`が自動で選ぶ。

## 設定

```
uv run main.py -h

usage: main.py [-h] -i INPUT_DIR [-o OUTPUT_FILE] [-d]

options:
  -h, --help            show this help message and exit
  -i INPUT_DIR, --input-dir INPUT_DIR
                        一括ダウンロードしたdocxファイルを配置したディレクトリ
  -o OUTPUT_FILE, --output-file OUTPUT_FILE
                        結合したpdfを出力するファイル
  -d, --delete-tmp-files
                        中間ファイルを削除するかどうか
```

## デモ

```
uv run main.py -i ./demo/
```

## 不具合メモ

- zipファイルが解凍できない場合（macOS・Linux）
- 以下のコマンドを実行すると治るかも

```
# a.zipが壊れている場合
zip -FF a.zip --out a-fixed.zip
unzip -O CP932 a-fixed.zip
```

- 参考
  - https://qiita.com/ktateish/items/28b69c18d64d50d6d5b0

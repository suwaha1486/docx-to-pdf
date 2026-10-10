"""一括ダウンロードしたdocxファイルをpdfに変換し、1つのpdfに結合する"""

import argparse
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path

# docx2pdfはファイルごとに進捗バーを出すので止める（全体の進捗は自前で出す）
# tqdmは読み込み時に環境変数を見るため、docx2pdfのimportより前に設定する
os.environ.setdefault("TQDM_DISABLE", "1")

import docx2pdf  # noqa: E402
import pypdf  # noqa: E402


def natural_key(text: str) -> list:
    """student2がstudent10より前に来るよう、数字を数値として比較する"""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", text)]


def collect_word_files(input_dir: Path) -> list[Path]:
    files = [
        p
        # macOSではglobが大文字小文字を区別するので、.DOCXも拾えるよう拡張子は自前で見る
        for p in input_dir.glob("*/*")
        if p.is_file()
        and p.suffix.lower() == ".docx"
        # Wordで開いている間にできるロックファイル（~$で始まる）は変換できないので飛ばす
        and not p.name.startswith("~$")
    ]
    return sorted(files, key=lambda p: natural_key(f"{p.parent.name}/{p.name}"))


def clean_name(word_file: Path) -> str:
    # たまに拡張子「.docx」を2つ付けたファイルが提出されている場合がある
    # 拡張子が正しくないと、docx2pdfがエラーを吐くので、不要な拡張子を削除する
    name = word_file.name
    if name.lower().endswith(".docx.docx"):
        name = name[:-5]
    # docx2pdfは小文字の「.docx」しか受け付けないので、.DOCXなどを揃える
    return name[:-5] + ".docx"


def flat_name(word_file: Path) -> str:
    """提出者のフォルダ名を付けた、1つのフォルダに並べられるファイル名"""
    return word_file.parent.name + "-" + clean_name(word_file)


def error_message(e: BaseException) -> str:
    # WordのCOMエラーは長いタプルになるので、Wordが返した説明文の1行目だけを出す
    excepinfo = getattr(e, "excepinfo", None)
    if excepinfo and excepinfo[2]:
        return excepinfo[2].splitlines()[0]
    return str(e) or type(e).__name__


def convert(word_file: Path, work_dir: Path) -> Path:
    """1ファイルを、名前を整えて作業フォルダにコピーしてからpdfに変換する"""
    # a-b/c.docx と a/b-c.docx のように整えた名前が衝突しうるので、呼び出し側が提出物ごとに別のwork_dirを渡す
    work_dir.mkdir()
    docx = work_dir / flat_name(word_file)
    shutil.copy(word_file, docx)
    pdf = docx.with_suffix(".pdf")
    docx2pdf.convert(docx, pdf, keep_active=True)
    if not pdf.exists():
        raise RuntimeError("pdfが作成されなかった")
    return pdf


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "-i",
        "--input-dir",
        type=Path,
        required=True,
        help="一括ダウンロードしたフォルダ（{提出者}/{ファイル}.docx の形）",
    )
    parser.add_argument(
        "-o",
        "--output-file",
        type=Path,
        default=Path("out.pdf"),
        help="結合したpdfを出力するファイル（既定: ./out.pdf）",
    )
    parser.add_argument(
        "-k",
        "--keep-pdfs",
        type=Path,
        metavar="DIR",
        help="変換した個別のpdfをこのフォルダに残す（既定では残さない）",
    )
    args = parser.parse_args()

    word_files = collect_word_files(args.input_dir)
    if not word_files:
        print(f"docxファイルが見つからない: {args.input_dir / '*' / '*.docx'}", file=sys.stderr)
        return 1
    print(f"{len(word_files)}件のdocxファイルを変換する: {args.input_dir}")

    pdfs: list[tuple[Path, Path]] = []
    failed: list[tuple[Path, str]] = []
    # 変換に失敗したファイルをWordが掴んだままのことがあるので、後片付けの失敗は無視する
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        work_dir = Path(tmp)
        for n, word_file in enumerate(word_files, 1):
            label = f"{word_file.parent.name}/{word_file.name}"
            print(f"[{n}/{len(word_files)}] {label}")
            try:
                pdfs.append((word_file, convert(word_file, work_dir / f"{n:04d}")))
            # macOS版のdocx2pdfは失敗時にsys.exitするため、SystemExitも拾って次のファイルへ進む
            except (Exception, SystemExit) as e:
                failed.append((word_file, error_message(e)))
                print(f"  -> 失敗: {failed[-1][1]}", file=sys.stderr)

        if not pdfs:
            print("1件も変換できなかった", file=sys.stderr)
            return 1

        writer = pypdf.PdfWriter()
        for word_file, pdf in pdfs:
            # しおり（目次）に提出者名を入れて、結合後のpdfで目的の提出物へ飛べるようにする
            writer.append(pdf, outline_item=f"{word_file.parent.name} - {Path(clean_name(word_file)).stem}")
        args.output_file.parent.mkdir(parents=True, exist_ok=True)
        try:
            writer.write(args.output_file)
        except PermissionError:
            # Windowsでは、出力先のpdfをビューアで開いていると上書きできない
            print(f"書き込めない（pdfを開いていたら閉じて再実行）: {args.output_file}", file=sys.stderr)
            return 1

        if args.keep_pdfs:
            args.keep_pdfs.mkdir(parents=True, exist_ok=True)
            # 名前の衝突を避けつつ結合順に並ぶよう、作業フォルダの連番を頭に付ける
            for _, pdf in pdfs:
                shutil.copy(pdf, args.keep_pdfs / f"{pdf.parent.name}_{pdf.name}")

    print(f"結合したpdfを作成した（{len(pdfs)}件）: {args.output_file}")
    if failed:
        print(f"\n変換できなかったファイル（{len(failed)}件）:", file=sys.stderr)
        for word_file, reason in failed:
            print(f"  {word_file}: {reason}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

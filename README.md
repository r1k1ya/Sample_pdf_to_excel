# PDF to Excel Converter

## 概要
本プロジェクトは、PDFファイルをExcelファイルに変換するツールです。GUIを備えており、簡単にPDFを選択して変換できます。  
サンプルPDF付(sample_sales_report.pdf)

## 特徴
- PDF内の**表データ**をExcelのシートに変換
- PDF内の**テキストデータ**もExcelに保存
- **GUI付き**で直感的に操作可能
- `pdfplumber`を使用してPDFデータを解析
- `openpyxl`を使用してExcelファイルを作成

## 動作環境
- Python 3.x
- Windows 10 / 11 (他のOSでも動作する可能性あり)

## インストール
### Python環境で実行する場合
1. 必要なライブラリをインストール
   ```sh
   pip install pdfplumber pandas openpyxl tkinter
   ```
2. `pdf_to_excel.py` を実行
   ```sh
   python pdf_to_excel.py
   ```

### EXEファイルとして実行する場合
`pyinstaller`を使用してEXEファイルを作成済みのため、Pythonがインストールされていなくても実行可能です。

1. `dist/pdf_to_excel.exe` をダウンロード
2. `pdf_to_excel.exe` をダブルクリックで起動

## EXEファイルの作成方法
EXEファイルを自作したい場合は、以下の手順で作成できます。
```sh
pip install pyinstaller
pyinstaller --onefile --noconsole pdf_to_excel.py
```
このコマンドを実行すると、`dist/` フォルダ内に `pdf_to_excel.exe` が作成されます。

## 使い方
1. アプリを起動
2. `PDFを選択して変換` ボタンをクリック
3. 変換したいPDFファイルを選択
4. 変換が完了すると、元のPDFと同じフォルダにExcelファイルが出力される

## ライセンス
このプロジェクトはMITライセンスのもとで公開されています。


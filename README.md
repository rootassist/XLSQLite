# XLSQLite

## 概要

ExcelからSQLiteを使うためのアドインです。

このアドインを組み込むことで、Excel 上から SQLite データベースを作成・操作することができます。また定義されたユーザー関数を用いて直接SELECT文を実行し、その結果を配列式あるいはスピルによってセル範囲に返すこともできます。

## SQLite for Excel

内部の処理でSQLite for Excelを利用しています。
https://github.com/govert/SQLiteForExcel

## ファイルの説明

### アドイン本体

- XLSQLite64J.xlam
  アドイン本体です(64bit/UI日本語/スピル対応済み)。Excelアドインとして組み込む必要があります。

### デモプログラム

- SQLite_DemoJ.xlsx
  アドインのデモプログラムです(64bit/スピル対応済み)。
- XLSQLiteDemo.sqlite
  デモで使用するデータベースです。

### SQLite関係ファイル

- SQLite3.dll
  SQLiteの本体です。アドインと同じフォルダに置いてください。
  このリポジトリには64bit版を置いています。
  32bit版のExcelで実行するときには、SQLiteのプロジェクトのホームページから32bit版のSQLite3.dllおよびSQLite3_StdCall.dllを入手して置いてください。

### オリジナル(移植元)

- XLSQLite.xlam
  オリジナルのアドインです。
- SQLite_Demo.xlsm
  オリジナルのデモプログラムです。

## 使い方

このアドインをExcelアドインとして組み込むと以下が利用できるようになります。

### ユーザー定義関数 SQLite_Query

指定されたSELECT文を実行し、結果を配列式もしくはスピルでセル範囲に展開します。

- 引数1: データベースファイルのファイル名(パス指定可)
- 引数2: SELECT文の文字列
- 引数3: 結果の返し方(Trueならスピル、Falseなら配列式)

### ツールバー

XLSQLiteツールバーがリボンに組み込まれます。以下の機能があります。

- Create/Add SQLite table
  SQLiteのテーブルの管理ができます(作成、変更など)。
- SQL Editor
  SELECT文を入力することで、結果を新しいワークブックに出力できます。

## プロジェクトのソースコード

sourcesディレクトリの下には本プロジェクトのVBAソースコードが格納されており、以下のプロジェクトをpre-commitで利用することで、コミット時に自動的に抽出されています。
<https://github.com/rootassist/extract_vba_source>
※forkした上で仕様変更とバグ修正を行っています。

具体的には、上記のリポジトリからextract_vba_source.pyをダウンロードして.git/hooks/の下に置きます。また.git/hooks/pre-commit には以下のように記載します。

```sh:pre-commit
#!/bin/sh
python .git/hooks/extract_vba_source.py \
                      --orig-extension \
                      --dest ./sources \
                      --src-encoding 'shift_jis' \
                      --out-encoding 'utf8' \
                      --recursive \
                      .
git add -- ./sources
```

## オリジナルのライセンス

このプロジェクトのオリジナルはMark Camilleriによって作成されたもので、これを64bit対応、UIの日本語化、およびスピル対応を行いました。
オリジナルはライセンスはMITライセンスとなっています。
<https://www.gatekeeperforexcel.com/other-freebies.html>

LICENSE:

The MIT License (MIT)
Copyright (c) 2013 Mark Camilleri

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

## ライセンス

The MIT License (MIT)
Copyright (c) 2023 Ryoji Nemoto

以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。

上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な部分に記載するものとします。

ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。

# Daten（堕天）

Daten（堕天）は、複雑な帳票（神Excel）のデータを通常のExcelの表（Proxy Excel）に変換し、プログラムから読み書きできるようにするソフトウェアです。Proxy Excelで加工したデータは、元の神Excelに書き戻せます。

このリポジトリでは、Datenのインストーラー、説明書、サンプルを公開しています。Datenの概要、開発の背景、利用条件は、[Webサイト](https://info-a11y.jp/daten/)をご覧ください。

## ダウンロード

[Releases](https://github.com/Information-Accessibility-Institute/daten/releases/latest) から、インストーラー（`Daten-<版>-win-x64-Setup.exe`）をダウンロードしてください。

- 対応環境：Windows x64。.NETの導入は不要です。管理者権限も不要です。
- インストールされるもの：Converter（神ExcelからProxy Excelを作る画面のプログラム）、Bridge（神ExcelとProxy Excelの間でデータを転記するコマンド）、Pythonクライアント、サンプル。
- Pythonクライアントを使うには、Python 3.9以降が必要です（Python本体は含まれていません）。
- 現在の版（0.8.1）は評価版で、使用期限は2027年1月31日です。期限を過ぎると、新しい変換や書き戻しができなくなります。作成済みのファイルは期限後も使えます。

インストール後は、スタートメニューの「すべて」（すべてのアプリ）にある「Daten」フォルダーの「はじめに」から始めてください。既定のインストール先は `%LOCALAPPDATA%\Programs\Daten` です。`AppData` は隠しフォルダーなので、エクスプローラーのアドレス欄にこれを入れて開いてください。

## 説明書

| 文書 | 内容 |
|---|---|
| [Converter](docs/Converter.md) | 神ExcelからProxy Excelを作り、見出しの名前を確認・確定する |
| [Bridge](docs/Bridge.md) | 神ExcelとProxy Excelの間でデータを転記する。JSONで読み書きする |
| [Python](docs/Python.md) | PythonからProxy Excelを読み書きする |
| [既知の問題と制限](docs/KnownIssues.md) | 分かっている問題と制限、その回避方法 |

**うまく動かないときは、まず[既知の問題と制限](docs/KnownIssues.md)をご覧ください。** 特に、Datenは数式を再計算しません。転記した後のファイルを提出・配布・集計に使う前に、Excelで開いてCtrl+Alt+F9で再計算し、保存し直してください。

## サンプル

[examples/fund](examples/fund/README.md) は、地方公共団体の決算の帳票にある「基金の状況」を使った例です。PythonでProxy Excelに値を入れ、神Excelへ転記し、Excelで再計算した結果を読み出すまでを、手順と期待値とともに示しています。インストール先の `examples` フォルダーにも同じものがあります。

## 利用条件

- 自分の組織の業務に使う場合は、個人・企業・大学・行政機関などを問わず無料です。ほかの人や組織から請け負った仕事に使う場合の条件は、[利用条件 ― 3分でわかるご案内](https://info-a11y.jp/daten/license-quick-guide/)をご覧ください。
- 正式な条件は、[Daten（堕天）ソフトウェア使用許諾契約](LICENSE.txt)です。インストール時に同意していただきます。
- Datenはオープンソースではありません。ConverterとBridgeのソースコードは公開していません。
- Pythonクライアント（インストール先の `python` フォルダーの `daten` パッケージ）と、サンプル（[examples](examples/)）は、MITライセンスで提供しています。サンプルは自由に使ってください。サンプルの帳票は総務省「地方財政状況調査」の様式をもとにしています（詳しくは[examples](examples/README.md)をご覧ください）。
- ソフトウェアは無保証で提供します。重要な業務で使う場合は、変換結果の確認と、元のファイルのバックアップをお願いします。

## お問い合わせ

不具合の報告やお問い合わせは、[メール](mailto:contact@info-a11y.jp)でお願いします。このリポジトリでは、IssueとPull Requestを受け付けていません。帳票のファイルを添付していただく場合は、個人情報や機密情報を含まないものにしてください。

提供者：[合同会社情報アクセシビリティ機構](https://info-a11y.jp/)

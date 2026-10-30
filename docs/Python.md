# Python から Proxy を読み書きする

Python の `daten` は、Bridge を呼び出して Proxy を読み書きするクライアントです。

## 1. 導入と Bridge の指定

Python 3.9 以降を用意してください。Python 本体は Daten に含まれていません。クライアントは、インストール先の `python` フォルダーにあります。PyPI では配布していません。

既定のインストール先は `%LOCALAPPDATA%\Programs\Daten`（たとえば `C:\Users\<ユーザー名>\AppData\Local\Programs\Daten`）です。`AppData` は隠しフォルダーなので、エクスプローラーでたどるより、アドレス欄に `%LOCALAPPDATA%\Programs\Daten` と入れて開くのが簡単です。

PowerShell でインストール先に移動してから、次を実行します。PYTHONPATH を設定するだけで使え、インストールやネットワーク接続は不要です。インストール先を変えた場合は、1行目を実際の場所に置き換えてください。

```powershell
cd "$env:LOCALAPPDATA\Programs\Daten"
$env:PYTHONPATH = (Join-Path $PWD 'python')
$env:DATEN_BRIDGE = (Join-Path $PWD 'bridge/Daten.Bridge.exe')
& $env:DATEN_BRIDGE --api-version
```

インストール先で `python -m pip install ./python` を実行して、ローカルにインストールすることもできます。この場合、setuptools などがなければ、取得にネットワーク接続が必要です。

Python クライアントは MIT ライセンスで提供します。Converter と Bridge には、同梱の LICENSE.txt（使用許諾契約）が適用されます。

Bridge は、関数の `cli=` 引数、環境変数 `DATEN_BRIDGE`、PATH 上の `Daten.Bridge` の順に探します。Bridge を呼び出す前に、JSON API の版が `1` であることを確認します。

```python
import daten

from pathlib import Path
records = daten.extract("confirmed.xlsx", cli=Path("bridge/Daten.Bridge.exe").resolve())
```

openpyxl や pandas は必要ありません。

## 2. 確定した Proxy を準備する

[Converter のガイド](Converter.md)に従って神Excel から Proxy を作成し、Excel で行・列の見出し名を確認・修正して確定します。テーブルの列名と見出しセルの文字は一致させ、管理情報（Custom XML）を保ってください。

確定した後の Proxy は、プログラムがデータを更新するためのものです。人が作業する場所としては使わないでください。確定した後に見出し名を変えたくなった場合は、Proxy を作り直して、確認・確定し直します。

管理情報のない古い Proxy は、Converter で作り直して、確定した名前を付け直してください。Python の extract・embed・辞書風 API では、管理情報の検査を省略できません。出力には、入力とは異なる、まだ存在しないパスを指定してください。

## 3. レコードで読み、指定した定数だけを更新する

```python
import daten

records = daten.extract("confirmed.xlsx")
# 一つのテーブルだけなら、sheet と table を両方指定する
selected = daten.extract("confirmed.xlsx", sheet="05", table="Daten_Table_1")
```

各レコードは、sheet・table・row・column・value と、formula・cache・numberFormat・dateSystem を持ちます。名前は、実際の抽出結果で確認してください。数式セルの value は保存されている計算結果（キャッシュ値）で、最新の計算結果とは限りません。キャッシュがあれば `cache` は `unverified`、なければ `value` は None（JSON では null）、`cache` は `missing` です。

次は、数値の定数のうちゼロでないものをゼロに変える例です。真偽値を数値と区別し、書込み用には最小限のレコードを作ります。

```python
patches = []
for record in records:
    if (record["formula"] is None
            and type(record["value"]) in (int, float)
            and record["value"] != 0):
        patch = {k: record[k] for k in ("sheet", "table", "row", "column")}
        patches.append(dict(patch, value=0))

if patches:
    result = daten.embed("confirmed.xlsx", patches, "edited-proxy.xlsx")
    print(result["changes"])
```

この例はすべての数値の定数を対象にしています。実際の業務では、更新すべきキーだけを選んでください。指定しなかったセルはそのまま保たれます。None は空欄、空文字は文字列、ゼロは数値です。value がないレコードはエラーになり、数式セルへの embed は拒否されます。

値には、組込みの int・float・str・bool・None を使います。数値は Excel と同じ倍精度で扱い、NaN・無限大や、変換できない巨大な整数は拒否されます。datetime などの自動変換はなく、日付はシリアル値のまま扱います。空でないセルは、同じ型の値で更新してください。

## 4. 辞書風 API で少数のセルを扱う

次の名前・値は例です。実在する定数セルを選んでください。

```python
from daten import open_proxy

with open_proxy("confirmed.xlsx") as book:
    table = book.table("05", "Daten_Table_1")
    value = table["地方税", "決算額"]
    print(value)
```

output を省略した Book は読取り専用です。更新するには出力先を指定します。

```python
with open_proxy("confirmed.xlsx", output="edited-proxy.xlsx") as book:
    table = book.table("05", "Daten_Table_1")
    table["地方税", "決算額"] = 55_000_000
    print(table["地方税", "決算額"])  # メモリ上の変更後の値
```

行・列の名前は、抽出結果の名前と完全に一致させて指定します。大文字と小文字も区別します。添字は文字列二つのタプルだけで、A1 形式の番地では指定できません。get・update・反復など、dict の機能すべてを提供するものではありません。

最初にテーブル全体を読み込むため、少数のセルを扱う場合でも、大きなテーブルでは時間とメモリが必要です。数式セルは、キャッシュがあっても、読取り・代入とも FormulaCellError になります。数式やキャッシュの情報が必要なら、extract を使ってください。

## 5. 保存される条件

with ブロックを正常に抜けたときだけ、保留していた変更をまとめて保存します。同じセルに何度か代入した場合は、最後の値が使われます。出力ができあがるまで、元のファイルは変更されません。

- 一度も代入しなければ、出力ファイルは作られません。
- ブロックから例外が出た場合は保存しません。例外はそのまま伝わります。
- セッションは1回限りです。with の外での Book・Table の操作や、一度抜けた Book での再開は RuntimeError になります。続けて処理する場合は、open_proxy をもう一度呼び、出力には毎回新しいパスを指定してください。
- 同じセッションでは、元の空でないセルの型が保たれます。数値→None→文字列のように、空欄を経由して型を変えることはできません。

変更がないときに出力が作られないのは、辞書風 API の仕様です。新しいファイルが必ずできる前提で、後続の処理を書かないでください。

## 6. 最新の神Excel から更新して戻す

次の例では、値が実際に異なる場合だけ代入し、変更がなければ kami-to-proxy の結果をそのまま proxy-to-kami に渡します。出力のファイル名は、すべてまだ存在しないものにしてください。

```python
from daten import kami_to_proxy, proxy_to_kami, open_proxy

kami_to_proxy(kami_data="kami.xlsx", proxy_template="proxy.xlsx", output_proxy="new-proxy.xlsx")

changed = False
with open_proxy("new-proxy.xlsx", output="edited-proxy.xlsx") as book:
    table = book.table("05", "Daten_Table_1")
    desired = 55_000_000
    if table["地方税", "決算額"] != desired:
        table["地方税", "決算額"] = desired
        changed = True

source = "edited-proxy.xlsx" if changed else "new-proxy.xlsx"
proxy_to_kami(proxy_data=source, kami_template="kami.xlsx", output_kami="new-kami.xlsx")
```

例外が出たら処理を止め、proxy-to-kami に進まないでください。proxy-to-kami は、変更したセルだけでなく、対応するすべての定数を同期します。kami-to-proxy から proxy-to-kami までの間に、ほかの人が神Excel を変更しないようにしてください。Book はデータベースのトランザクションではなく、競合の検出や排他は行いません。

**数式の計算結果について**：kami-to-proxy は、神Excel の数式セルに保存されている計算結果（キャッシュ値）を Proxy に転記します。神Excel が再計算・保存済みかどうかは確認しません。実行前に神Excel を Excel で開き、Ctrl+Alt+F9 で再計算して保存してください。

**再計算に関する 0.8.1 の既知の制限**：proxy-to-kami 後の神Excel と embed 後の Proxy をプログラムで読むと、数式セルから古い計算結果が返ることがあります。openpyxl で `data_only=True` を指定しても、再計算は行われません。計算結果を使う場合は、Excel で出力を開き、**Ctrl+Alt+F9（すべて再計算）** を実行して保存し直してから読んでください。詳しくは [Bridge のガイド](Bridge.md#再計算に関する-081-の既知の制限)を参照してください。

dry-run、上書き、管理情報の照合の省略が必要な場合は、[Bridge のガイド](Bridge.md)に従って、Bridge を直接使ってください。Python の関数には、これらの引数はありません。

## 7. エラーに対処する

```python
import daten

try:
    records = daten.extract("confirmed.xlsx")
except daten.DatenError as error:
    print(error)
    print(error.returncode)
    print(error.diagnostics)
    print(error.stderr)
    raise
```

| 例外・状況 | 確認すること |
|---|---|
| DatenError | Bridge の場所、API の版、管理情報、テーブル、診断の内容、保存先。kami-to-proxy・proxy-to-kami では stderr のテキストも確認する |
| FormulaCellError | 数式セルを通常の値として読み書きしていないか |
| KeyError | 行・列の名前が完全に一致しているか、テーブルの選択が正しいか。メッセージに示された欠けている項目を確認する |
| TypeError | 添字が文字列二つのタプルか、値の型が対応しているか、空でないセルの型と一致しているか |
| ValueError | 有限でない値、sheet・table の片方だけの指定、入力と同じ出力先などがないか |
| PermissionError | 読取り専用の Book に代入していないか。output を指定したか |
| FileExistsError | 出力先がまだ存在しないパスか。既存のファイルを黙って上書きすることはない |
| RuntimeError | Book・Table を有効な with の中で使っているか。DatenError も RuntimeError の一種なので、具体的な型を確認する |

## 8. 基金の往復例

[基金の状況](../examples/fund/README.md)は、`open_proxy` で定数を設定し、神Excel へ転記し、Excel で再計算・保存し、Bridge で数式の計算結果を Proxy に転記して読み出す例です。インストール先の `examples` フォルダーにも同じ例があります。手順と期待値は、リンク先を参照してください。

両方向の転記では、データセルの書式のうち、数値の表示形式と日付体系だけを比較します。それ以外の書式の違いは検査せず、転記先の書式はそのまま保たれます。

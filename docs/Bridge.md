# Bridge でデータを連携する

Bridge は、確定した Proxy の見出しを保ったまま、神Excel との間で定数の値を同期したり、JSON のレコードとの間でデータをやり取りしたりするコマンドラインのプログラムです。

## 1. 準備する

Bridge は、インストール先の `bridge\Daten.Bridge.exe` です。既定のインストール先は `%LOCALAPPDATA%\Programs\Daten`（たとえば `C:\Users\<ユーザー名>\AppData\Local\Programs\Daten`）です。`AppData` は隠しフォルダーなので、エクスプローラーでたどるより、アドレス欄に `%LOCALAPPDATA%\Programs\Daten` と入れて開くのが簡単です。インストール時に PATH への登録を選んだ場合は、`Daten.Bridge.exe` だけで起動できます。インストール前から開いていた端末は、開き直してください。登録しなかった場合は、フルパスで起動します。既定のインストール先なら、PowerShell では次のとおりです。

```powershell
& "$env:LOCALAPPDATA\Programs\Daten\bridge\Daten.Bridge.exe" --help
```

以下の例は、PATH に登録した場合の書き方です。Bridge 自体に Python は不要です。

```powershell
Daten.Bridge.exe --version
Daten.Bridge.exe --api-version
Daten.Bridge.exe --help
```

各コマンドの終了コードを確認してください。JSON API の版は `1` です。

あらかじめ、[Converter](Converter.md) で神Excel から Proxy を作成し、Excel で行・列の見出し名を確認・修正して確定しておきます。Proxy には、Converter が作る Excel のテーブルと管理情報（Custom XML）が必要です。

神Excel と Proxy は、同じ名前・同じ順序のシートの、同じセル番地どうしが対応している必要があります。行・列の移動や並べ替え、シート名の変更には追従しません。出力には、入力とは異なる、まだ存在しないパスを指定し、出力先のフォルダーは作っておいてください。

## 2. 最新の神Excel から Proxy へ取り込む（kami-to-proxy）

`kami.xlsx` を最新の神Excel、`proxy.xlsx` を確定した Proxy とします。まず、保存せずに検査だけを行えます。

```powershell
Daten.Bridge.exe kami-to-proxy --kami-data kami.xlsx --proxy-template proxy.xlsx --output-proxy new-proxy.xlsx --dry-run
```

dry-run でも `--output-proxy` は必須です。対象のテーブル、定数の数、数式の計算結果の対象数・変更数、変更の総数とエラーを確認します。出力ファイルは作られません。問題がなければ、実際に取り込みます。

```powershell
Daten.Bridge.exe kami-to-proxy --kami-data kami.xlsx --proxy-template proxy.xlsx --output-proxy new-proxy.xlsx
```

確定した見出しはそのままに、対象のすべての定数と数式の計算結果を神Excel に合わせた `new-proxy.xlsx` が作られます。各操作の成功を確認してから、次に進んでください。

**数式の計算結果について**：kami-to-proxy は、神Excel の数式セルに保存されている計算結果（キャッシュ値）を Proxy へ転記します。神Excel が再計算・保存済みかどうかは確認しません。実行前に神Excel を Excel で開き、Ctrl+Alt+F9 で再計算して保存してください。転記した件数と、再計算を確認していない旨は、標準出力に報告されます。

## 3. データを JSON へ取り出す（extract）

```powershell
Daten.Bridge.exe extract --proxy new-proxy.xlsx --output records.json
```

一つのテーブルだけを取り出す場合は、シートとテーブルを両方指定します。次の名前は例なので、実際の名前に置き換えてください。

```powershell
Daten.Bridge.exe extract --proxy new-proxy.xlsx --sheet "05" --table Daten_Table_1 --output table-records.json
```

`--output` を省略すると、JSON を標準出力に返します。ファイルが必要なら `--output` を使ってください。PowerShell のリダイレクト（`>`）で保存したファイルは、UTF-8 になるとは限りません。

Windows PowerShell 5.1 で Bridge の出力を変数に取り込む場合は、事前に次を設定してください。

```powershell
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$result = Daten.Bridge.exe --help
```

各レコードは、sheet・table・row・column で更新の対象を特定します。数式セルの value は、保存されている計算結果（キャッシュ値）で、最新の計算結果とは限りません。formula と cache も確認してください。

## 4. 指定した定数だけを更新する（embed）

UTF-8 で `patches.json` を作ります。次のキーは例です。実際の extract の結果にある、定数セルのキーを使ってください。

```json
{
  "version": 1,
  "records": [
    {
      "sheet": "05",
      "table": "Daten_Table_1",
      "row": "地方税",
      "column": "決算額",
      "value": 55000000
    }
  ]
}
```

```powershell
Daten.Bridge.exe embed --proxy new-proxy.xlsx --data patches.json --output edited-proxy.xlsx
```

指定しなかったセルは変更されません。value に null を指定すると空欄になり、空文字列やゼロとは区別されます。空でないセルは、同じ型の値で更新してください。数式セルには書き込めません。extract の結果をすべてそのまま embed に渡さず、更新したい定数のレコードだけを用意してください。

**embed 後の Proxy も再計算されません。** 計算結果を使う前に、[再計算に関する 0.8.1 の既知の制限](#再計算に関する-081-の既知の制限)に従って、再計算・保存してください。

成功すると、version と、実際に変更したセルの数 changes を返します。すべてのレコードを検査してから出力を作り、エラーがあれば出力は作りません。embed は dry-run と `--overwrite` を受け付けません。JSON を標準入力から渡す場合は、`--data -` と指定します。

## 5. 神Excel へ戻す（proxy-to-kami）

```powershell
Daten.Bridge.exe proxy-to-kami --proxy-data edited-proxy.xlsx --kami-template kami.xlsx --output-kami new-kami.xlsx --dry-run
```

検査が成功したことを確認してから、保存します。

```powershell
Daten.Bridge.exe proxy-to-kami --proxy-data edited-proxy.xlsx --kami-template kami.xlsx --output-kami new-kami.xlsx
```

**proxy-to-kami は、embed で変更したセルだけでなく、対応するすべての定数セルを同期します。** 古い Proxy や空欄にした Proxy から proxy-to-kami すると、その値も神Excel に反映されます。最新の神Excel から kami-to-proxy、必要な定数を更新、同じ神Excel へ proxy-to-kami、という順で作業し、その間にほかの人が神Excel を更新しないようにしてください。同時に更新されても、検出はできません。

### 再計算に関する 0.8.1 の既知の制限

> **提出・配布・集計に使う前に、Excel で出力を開き、Ctrl+Alt+F9（すべて再計算）を実行してから保存し直してください。**
>
> Daten は再計算しません。proxy-to-kami 後の神Excel と embed 後の Proxy では、定数が更新されても、数式の計算結果（キャッシュ値）とグラフに保存された値は古いまま残ります。正常に終了しても、計算結果が最新であるとは限りません。
>
> Excel で開いても、自動で再計算されるとは限りません。計算方法が「手動」に設定されたブックでは、なおさら明示的な再計算が必要です。

再計算・保存した後に、値・数式・グラフ・帳票の見た目を確認してください。

## 6. オプション

### 既存の出力を置き換える

kami-to-proxy と proxy-to-kami では、`--overwrite` を付けると既存の出力を置き換えられます。

```powershell
Daten.Bridge.exe kami-to-proxy --kami-data kami.xlsx --proxy-template proxy.xlsx --output-proxy new-proxy.xlsx --overwrite
```

`--overwrite` を付けても、入力ファイルそのものを出力先にすることはできません。ジャンクションなど別名のパスを経由して入力と同じファイルを指定した場合も拒否されます。extract と embed では、常にまだ存在しない出力先を指定してください。

### 管理情報（Custom XML）の照合を省略する

通常は、元の帳票から Proxy を作り直し、確定した名前を付け直してください。ほかのツールで編集して管理情報が失われた場合などで、同じ帳票・同じ配置であることを利用者が確認できたときだけ、kami-to-proxy と proxy-to-kami で次を指定できます。

```powershell
Daten.Bridge.exe proxy-to-kami --proxy-data proxy.xlsx --kami-template kami.xlsx --output-kami new-kami.xlsx --skip-custom-xml-validation
```

管理情報の確認と、元の見出しとの照合を省略します。神Excel と Proxy を取り違えても検出できない場合があるため、その旨の警告が標準エラーに1回出ます。テーブル・数式・キー・セル範囲・日付体系・数値書式などの検査は省略されません。この指定は Proxy を修復するものではありません。管理情報のない Proxy は、extract と embed には使えません。

## 7. 結果を確認し、失敗に対処する

PowerShell では、各コマンドの直後に `$LASTEXITCODE` を確認します。0 は成功、1 は処理の失敗、2 は引数のエラーです。自動処理では、0 以外なら後続の処理を止めてください。

```powershell
if ($LASTEXITCODE -ne 0) { throw "Daten Bridge failed: $LASTEXITCODE" }
```

kami-to-proxy と proxy-to-kami の結果はテキストで、extract と embed のエラーは標準エラーの JSON で返されます。JSON の diagnostics が空でも、error を確認してください。構造の不一致は、総数と最大100件の内容が表示されます。

| 症状 | 対処 |
|---|---|
| 必須の引数がない・未知の引数 | `--help` で、操作に対応する引数を確認する。シートとテーブルは両方指定する |
| 出力が既存・入力と同じ・フォルダーがない | 出力先を修正する。dry-run でも出力先の条件を検査する |
| 管理情報がない・壊れている・対応していない | 対応する元の帳票から Converter で Proxy を作り直し、確定した名前を付け直す |
| 元の見出し・テーブルの範囲・シートなどの不一致 | 対になる帳票か、配置や名前を変えていないかを確認する |
| キー・型・数式の拒否 | 未知・重複・空白のキー、空でないセルの型、数式セルへの書込みがないかを確認する |
| 数値書式・日付体系の不一致 | 対応関係と表示形式を確認する。この不一致は許容されない |
| ファイルにアクセスできない・保存の失敗 | 他のアプリで開いていないか、権限、パスを確認して再実行する。失敗したときの出力は使わない |

両方向の転記では、データセルの書式のうち、数値の表示形式（formatCode）と日付体系（1900年／1904年）だけを比較します。フォント・塗りつぶし・罫線などの違いは検査しません。転記先の書式はそのまま保たれます。

[基金の状況の往復例](../examples/fund/README.md)では、Python で定数を設定し、神Excel へ転記し、Excel で再計算・保存し、数式の計算結果を Proxy に取り込んで読み出すまでの手順と期待値を示しています。

## 8. 年度の更新と確定した名前の保守

新年度の帳票から新しい Proxy を作った場合は、新旧の確定した行・列の見出し名を、表ごとに比べます。変わっていない項目は、外部プログラムの対応付けをそのまま使えます。追加・削除・改称・階層の変更だけを人が確認して、転記の規則を更新します。帳票の表記が変わっても業務上同じ項目なら、前年と同じ名前を付けて確定することもできます。

Bridge は、名前が同じでも、配置の違う帳票どうしを結び付けることはしません。異なる年度の帳票には、それぞれ対応する Proxy を用意してください。

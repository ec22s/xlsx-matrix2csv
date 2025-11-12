# xlsx-matrix2csv

- マトリクス型のExcelデータ → `KEY, VALUE` 型のCSVに変換出力します

- CSVの元となる配列をJSON出力することもできます

- 変換イメージ

  |Group A|Group B|Group C|Group D|
  |----|----|----|----|
  |Data A-1|Data B-1|Data C-1|Data D-1|
  |Data A-2| |Data C-2|Data D-2|
  |Data A-3|Data B-3|Data C-3| |
  |Data A-4|Data B-4| |Data D-4|

  ↓
  ```
  "Group A","Data A-1"
  "Group A","Data A-2"
  "Group A","Data A-3"
  "Group A","Data A-4"
  "Group B","Data B-1"
  "Group B","Data B-3"
  "Group B","Data B-4"
  "Group C","Data C-1"
  "Group C","Data C-2"
  "Group C","Data C-3"
  "Group D","Data D-1"
  "Group D","Data D-2"
  "Group D","Data D-4"
  ```

<br>

- 使用例

  - 準備として `npm i` を実行し、各種設定をJSONファイルに入力します。サンプルとして `config.example.json` がソースにあります

  - `npm run example` を実行すると、設定サンプルにもとづきテスト用Excelファイルのデータを変換します

  - `package.json` に設定ファイルのパスや `npm run` のコマンドがあるので、自由に変更できます。また設定ファイルを直接指定して `node . XXXX.json` と実行することもできます

<br>

- 空白セルは結果のCSVに現れません

- 結果の各行毎に共通で追加する列を設定できます

  - 設定箇所 `config.example.json` > `csv.extraData`

  - 例 共通に追加する列群が `Add-1, Add-2, Add-3` の場合、下のようになります

    ```
    "Group A","Data A-1","Add-1","Add-2","Add-3"
    "Group A","Data A-2","Add-1","Add-2","Add-3"
    "Group A","Data A-3","Add-1","Add-2","Add-3"
    "Group A","Data A-4","Add-1","Add-2","Add-3"
    "Group B","Data B-1","Add-1","Add-2","Add-3"
    "Group B","Data B-3","Add-1","Add-2","Add-3"
    ...
    (以下同様)
    ```

- 共通で追加する行列を設定し、各行を複製して追加することもできます

  - 設定箇所は前項と同じ

  - 例 共通に追加する行列が
    ```
    Add-1-1, Add-1-2
    Add-2-1, Add-2-2
    ```
    の場合、下のようになります

    ```
    "Group A","Data A-1","Add-1-1","Add-1-2"
    "Group A","Data A-1","Add-2-1","Add-2-2" ← 複製された行
    "Group A","Data A-2","Add-1-1","Add-1-2"
    "Group A","Data A-2","Add-2-1","Add-2-2" ← 複製された行
    "Group A","Data A-3","Add-1-1","Add-1-2"
    ...
    (以下同様)
    ```
  - この例が `config.example.json` で設定されています

<br>

- そんな特殊なケースが仕事で実際あったので作りました。

<br>

- 設定ファイルについて

  - 設定項目は下記のとおりです。具体例は `confing.example.json` を見て下さい。

    ```
    {
      "csv": {
        "basePath": 出力ファイルの拡張子以外のパス,
        "overWrite": 同名の出力ファイルを上書きするか否か,
        "timestamp": 出力ファイルにタイムスタンプを付与するか否か,
        "delimCol": 列区切り文字,
        "delimRow": 行区切り文字,
        "extraData": 各行に追加する行または行列を2次元配列で指定
      },
      "json": {
        "output": JSONファイルを出力するか否か,
        "basePath": 出力ファイルの拡張子以外のパス,
        "indent": 1以上なら整形してこの数の分だけスペースでインデント,
        "overWrite": 同名の出力ファイルを上書きするか否か,
        "timestamp": 出力ファイルにタイムスタンプを付与するか否か
      },
      "xlsx": {
        "file": 読み込むExcelファイル,
        "sheetIndex": シート番号 (0始まり),
        "range": 変換対象とする範囲の左上セル番地, 右下セル番地 (A1形式)
      }
    }
    ```

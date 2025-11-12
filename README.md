# xlsx-matrix2csv (WIP)

- 神Excelの縦持ちデータをノーマルなCSVに変換出力します

  - イメージ
  
    |Group A|Group B|Group C|Group D|
    |----|----|----|----|
    |Data A-1|Data B-1|Data C-1|Data D-1|
    |Data A-2| |Data C-2|Data D-2|
    |Data A-3|Data B-3|Data C-3| |
    |Data A-4|Data B-4| |Data D-4|
    
    ↓
    ```
    "Group A","Data A-1","Data A-2","Data A-3","Data A-4"
    "Group B","Data B-1","","Data B-3","Data B-4"
    "Group C","Data C-1","Data C-2","Data C-3",""
    "Group D","Data D-1","Data D-2","","Data D-4"
    ```

<br>

- 結果の各行毎に共通で追加する列を設定できます

  - 各行を複製して共通の行・列を追加することもできます

  - 参考 `config.example.json` > `csv.extraData`

<br>

- CSVの元となる配列をJSON出力することもできます (予定)
